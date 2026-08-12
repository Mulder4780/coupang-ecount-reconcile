# -*- coding: utf-8 -*-
"""
recheck_plan.py — 밴드 전량 재확인, 다음에 훑을 글 번호를 캐시에서 뽑는다.

왜: 재수집(수정글 감지)은 글당 5초+ 걸려 세션 하나로 못 끝난다. "어디까지 했나"를
사람이 기억하게 하지 않고, 캐시의 captured_at(재수집 시각) 유무로 기계가 판정한다.
  · 캐시에 없는 번호      → 신규/삭제 후보 (우선)
  · captured_at 없는 번호 → 아직 옛 수집분 그대로 (재수집 대상)

쓰는 법:
  python band/recheck_plan.py                 # 밴드별 남은 구간 요약
  python band/recheck_plan.py --band 90610953 --limit 60
      → __grabStart 에 넣을 JS 배열을 그대로 출력 (band/grab_posts.js 참조)
"""
import sys, os, json, argparse
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):   # pythonw 는 sys.stdout 이 None 이다([43])
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")
# 근거 한 장의 자리 — session_handoff·convert_dump 와 **같은 파일**이다.
# 상수로 둔 이유: 함수 안에서 매번 경로를 조립하면 시험이 진짜 파일을 건드려야 한다.
PROBE_LOG = os.path.join(os.path.dirname(HERE), "reports", "밴드_확인시각.json")
# ★ 수정글 감지 시대의 시작(2026-08-04, 상세 페이지 재수집 도입). 7월의 피드/API
#   덤프도 capturedAt 을 갖고 있어 '유무'만 보면 옛 수집이 재수집으로 오판된다 —
#   이 시각 이후의 captured_at 만 재수집 완료로 인정한다(convert_dump 는 사실만 기록).
ERA_MS = 1785769200000  # 2026-08-04 00:00 KST


def load(band):
    p = os.path.join(CACHE, f"{band}.json")
    if not os.path.exists(p):
        return None
    return json.load(open(p, encoding="utf-8")).get("posts") or {}


SCOPE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "collect_scope.json")


def scope():
    """어디까지 훑을지 — **사람이 정한 값**을 파일에서 읽는다 (2026-08-06).

    대화에 남긴 결정은 다음 세션·다른 계정이 모른다. 그래서 band/collect_scope.json 이
    기억하고, 이 도구는 그것을 기본값으로 쓴다. 범위를 넓히는 것은 사람의 결정이다.
    """
    try:
        return json.load(open(SCOPE, encoding="utf-8"))
    except Exception:
        return {}


QUIET_LIMIT_DAYS = 1          # 밴드 신선도 한도와 같다(session_handoff.FRESH_LIMIT)

# ★ 근거가 없을 때 **위쪽을 몇 개까지 찔러 볼까** (2026-08-11 실사고).
#   글 번호는 이어지므로 `hi+1` 하나만 열어 봐도 "새 글이 있나"는 답이 나온다.
#   그런데 예전에는 근거가 하루만 낡아도 `ahead`(40) 를 통째로 목록에 넣었다 —
#   없는 번호 한 개가 iframe 9초 + 본문 12초를 꽉 채우므로 40개면 **약 14분**을
#   쓰고 수확은 0 이다(2026-08-11 16:10 회차가 그렇게 낭비됐다).
#   그래서 근거가 없으면 **싸게 존재만 확인**하고, 있는 것이 확인되면
#   그다음 회차가 이어받는다(수확이 들어오면 `hi` 가 올라가 사다리가 저절로 오른다).
PROBE_AHEAD = 5

# 캐시가 '이건 업무 기록이 아니다'라고 표시해 둔 갈래들. 이름이 세 벌인 이유는
# 역사다(`contaminated` 는 clean_contaminated·convert_dump, `absent` 는 real_latest,
# 나머지는 옛 도구). **읽는 자리를 한 곳으로 모아** 도구마다 다른 낱말을 보고
# 서로 다른 것을 거르는 일을 막는다 — 한쪽만 거르면 다른 쪽은 없는 번호를 긁는다.
DEAD_FLAGS = ("deleted", "contaminated", "absent", "tainted", "ghost", "dirty")


def is_dead(v):
    """이 글은 다시 훑을 대상이 아닌가(삭제·오염·유령)."""
    return isinstance(v, dict) and any(v.get(f) for f in DEAD_FLAGS)


def _rec(band):
    """근거 한 줄 — `reports/밴드_확인시각.json` 의 그 밴드 항목."""
    try:
        return (json.load(open(PROBE_LOG, encoding="utf-8")) or {}).get(str(band)) or {}
    except Exception:
        return {}


def _age_days(seen, today=None):
    try:
        day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
        return (datetime.strptime(day, "%Y-%m-%d")
                - datetime.strptime(str(seen)[:10], "%Y-%m-%d")).days
    except ValueError:
        return None


def absent_line(band, posts=None, today=None):
    """`(cut, 이유)` — **이 번호부터 위는 없다**고 확인된 지점. 없으면 `(None, 이유)`.

    근거는 `reports/밴드_확인시각.json` 한 장이다 — convert_dump 가 덤프의
    missing·notime 지문에서 만들고 real_latest 가 사람 확인으로 적는다.
    session_handoff 의 '조용함' 판정과 **같은 파일**을 본다(두 곳이 어긋나면 한쪽은
    긁으라 하고 다른 쪽은 조용하다고 해서 사람이 무엇을 믿을지 모르게 된다).

    ★ **근거 없이 '조용함'을 단정하지 않는다.** 셋 중 하나면 근거가 없는 것으로 본다:
      · 아예 없다
      · **한도보다 오래됐다** — 그 사이 새 글이 올라왔을 수 있다(다시 밀림이다)
      · **추월됐다** — 근거가 "N 부터 없다"는데 캐시에 N 이상의 **진짜 글**(작성시각이
        있는 수확)이 이미 들어와 있다. 그때는 근거가 틀린 것이다. 실측 2026-08-11:
        84789192 근거 `없음확인 3539`(08-09) · 캐시에 3539 가 진짜 글로 있었다.
        이것을 안 걸러 내면 **실재하는 글을 유령으로 표시**하게 된다.
      근거가 없다고 40개를 긁으라는 뜻은 아니다 — 그때는 `PROBE_AHEAD` 만 찔러 본다.
    """
    return judge_absent(_rec(band), posts, today)


def judge_absent(rec, posts=None, today=None):
    """근거 **한 장**을 믿어도 되는지만 가린다 — 파일은 여기서 안 읽는다.

    읽는 자리가 둘이라(수집 계획 `absent_line` · 인계 문서 `session_handoff`)
    **판정은 여기 하나**에 둔다. 그런데 파일까지 여기서 읽으면 부르는 쪽이 제 손으로
    읽어 둔 근거를 대 볼 수 없다 — 그러면 두 화면이 서로 다른 한 장을 놓고 답하게 된다.
    그래서 근거는 **받고**, 판단만 한다.
    """
    rec = rec or {}
    try:
        n = int(rec.get("없음확인") or 0)
    except (TypeError, ValueError):
        n = 0
    seen = str(rec.get("확인시각") or "")[:10]
    if n <= 0 or not seen:
        return None, "근거 없음"
    age = _age_days(seen, today)
    if age is None or not (0 <= age <= QUIET_LIMIT_DAYS):
        return None, "근거가 낡음(%s 확인 · %s일 지남)" % (seen, "?" if age is None else age)
    if posts:
        top = trusted_hi(posts)
        if top is not None and top >= n:
            return None, "근거 추월됨(%s 부터 없다는데 %s 가 실제로 수확됐다)" % (n, top)
    return n, "%s 부터 없음 확인(%s)" % (n, seen)


def trusted_hi(posts):
    """**믿을 수 있는 최대 번호** — 작성시각이 있고 죽지 않은 글 중 가장 큰 번호.

    작성시각이 있다는 것만이 "실제로 열려서 수확됐다"는 증거다(검증 [130]).
    """
    ns = [int(k) for k, v in (posts or {}).items()
          if str(k).isdigit() and isinstance(v, dict)
          and v.get("created_at") and not is_dead(v)]
    return max(ns) if ns else None


def _confirmed_quiet(band, hi, today=None):
    """`hi` 바로 위가 '없음'으로 확인됐고, 그 확인이 아직 최근인가."""
    rec = _rec(band)
    try:
        if int(rec.get("없음확인") or 0) != int(hi) + 1:
            return False
    except (TypeError, ValueError):
        return False
    seen = str(rec.get("확인시각") or "")[:10]
    if not seen:
        return False
    age = _age_days(seen, today)
    return age is not None and 0 <= age <= QUIET_LIMIT_DAYS


def _absent_from(band, today=None, posts=None):
    """`absent_line` 의 값만 — 옛 호출자를 위해 남긴다."""
    return absent_line(band, posts, today)[0]


def plan(band, posts, floor=0, ahead=0, today=None):
    ks = sorted(int(k) for k in posts if str(k).isdigit())
    if not ks:
        return None
    lo, hi_cache = ks[0], ks[-1]
    have = set(ks)

    def _e(n):
        v = posts.get(str(n))
        return v if isinstance(v, dict) else {}

    # ★ **믿을 수 있는 최대 번호**로 위쪽을 잰다 (2026-08-07 2차 사고).
    #   예전에는 캐시가 가진 가장 큰 번호를 그대로 `hi` 로 썼다. 그런데 오염 사고로
    #   **존재하지도 않는 번호**가 캐시에 들어앉아 있었고, 그 유령이 `hi` 를 밀어 올렸다
    #   (90610953: 캐시 5464 · 실제 최신 5437). 그러면 `hi+1` 부터 만든 '새 글 후보'가
    #   처음부터 없는 번호가 되고, 없는 번호를 긁는 것이 바로 오염을 만드는 행위다.
    #   **사고가 사고를 먹여 살리는 고리**라서, 끊는 자리는 여기다.
    #   작성시각이 있는 것만이 "실제로 열려서 수확됐다"는 증거다.
    hi = trusted_hi(posts)
    if hi is None:
        hi = hi_cache
    # ★ 캐시 **위쪽**(새 글)을 먼저 본다. 예전에는 구멍만 봐서, 마지막 수집 이후 올라온
    #   글이 영원히 대상 밖이었다 — 2026-08-06 쿠팡AS 밴드가 8/4 에 멈춰 있는데도
    #   "구멍 0" 이라 아무도 몰랐고, 8/5 돌발AS 가 1건으로 보고됐다.
    # ★ 방금 '없다'고 확인한 번호를 또 훑지 않는다 (2026-08-07 지시).
    #   글 번호는 이어지므로 **없음확인이 N 이면 N 이상은 전부 없다.** 그런데도 매 회차
    #   40번을 헛짚었고(글당 5초 → 3분), 그 헛짚음이 바로 오늘 오염 사고의 입구였다.
    #   확인이 오래되면 그 사이 새 글이 올라왔을 수 있으므로 그때는 다시 훑는다 —
    #   판정 기준은 session_handoff 의 '조용함' 과 같은 근거 파일 하나다.
    # ★ 근거가 없으면 **적게 찔러 본다** (2026-08-11 실사고). 예전에는 근거가 하루만
    #   낡아도 이 자리가 40개를 그대로 쏟았고, 그 40개는 아직 없는 번호라 한 개씩
    #   21초를 채우고 no-time 으로 버려졌다 — 14분에 수확 0. 존재 여부는 `hi+1`
    #   하나로 답이 나오므로 `PROBE_AHEAD` 만 넣고, 있는 것이 확인되면 `hi` 가
    #   올라가 **다음 회차가 이어받는다.** 근거(cut)가 살아 있으면 그 아래는 실재하는
    #   글이므로 예전대로 `ahead` 까지 간다 — 실재하는 글을 긁는 것은 낭비가 아니다.
    cut, absent_why = absent_line(band, posts, today)   # 이 번호부터 위는 없다(확인됨)
    span = int(ahead or 0)
    probing = cut is None
    if probing:
        span = min(span, PROBE_AHEAD)
    new = [n for n in range(hi + 1, hi + 1 + span)
           if n not in have and (cut is None or n < cut)]
    # ★ 아래쪽 구멍(2026-08-06 발견). 예전에는 **보유한 것 중 가장 작은 번호부터**만
    #   구멍으로 봤다. 그래서 캐시가 4196~5424 여도 "구멍 0"이라 나왔고, 1~4195(2023-03
    #   부터의 4,195건)가 통째로 수집 대상 밖에 있었다 — 없는 줄도 몰랐다.
    #   floor 를 1 로 주면 밴드 개설 이후 전량이 대상이 된다.
    gaps = [n for n in range(max(1, floor), hi + 1) if n not in have]
    # 삭제된 글은 다시 훑지 않는다(2026-08-05). 밴드가 지운 글을 열면 '삭제됨'이 아니라
    # 밴드 홈을 돌려주므로 수집은 언제나 실패한다 — 목록에 남겨 두면 영원히 반복된다.
    stale = sorted(int(k) for k, v in posts.items()
                   if str(k).isdigit() and not v.get("deleted")
                   and int(v.get("captured_at") or 0) < ERA_MS)
    dead = sum(1 for v in posts.values() if isinstance(v, dict) and v.get("deleted"))
    # ★ **본문은 있는데 날짜가 없는 글**도 다시 가져와야 한다 (2026-08-07 발견).
    #   밴드는 본문(.postText)을 먼저 칠하고 작성시각(.time)을 조금 뒤에 채운다.
    #   본문을 보자마자 가져간 글은 날짜가 빈 채로 저장됐다 — 621건이었다.
    #   이 글들은 구멍도 아니고(번호가 있다) 오래된 것도 아니라(오늘 받았다)
    #   **어느 목록에도 안 잡혔다.** 캐시 숫자는 늘어나는데 대조는 안 되는,
    #   가장 알아채기 어려운 종류의 구멍이다. 날짜가 없으면 어떤 작업과도 못 맞춘다.
    #
    # ★ 그런데 그 621건의 정체는 '날짜만 빈 진짜 글'이 아니었다 (2026-08-07 2차).
    #   밴드가 `/post/<번호>` 를 iframe 으로 열면 **피드로 되돌린다.** 그래서 껍데기에
    #   남은 피드 맨 위 글이 잡혔고, 번호 수백 개가 **같은 본문**을 갖게 됐다
    #   (98건→본문 2종 · 523건→7종). 즉 남의 본문을 베껴온 가짜 기록이다.
    #   `clean_contaminated.py` 가 이것들을 `contaminated` 로 표시했다.
    #
    #   **다시 훑지 않는다** — 표본 3/3 이 리다이렉트로 확인돼 **삭제된 글**로 판정났다.
    #   목록에 넣으면 매 회차 621번(약 50분)을 헛돈다.
    #
    # ★ 다만 오염을 **한 통에 담아 두면 안 된다** (2026-08-07 3차). 두 갈래가 섞여 있다:
    #     · 실제 최신 글 번호 **이하** → 삭제된 진짜 번호. 위 판정대로 훑지 않는다.
    #     · 실제 최신 글 번호 **위**   → 처음부터 **없던 번호**(유령).
    #   갈라 두는 이유는 재수집 여부가 아니라 **`hi` 오염을 막기 위해서**다. 유령이
    #   `hi` 를 밀어 올리면 그 위로 '새 글 후보'가 생기고, 없는 번호를 긁는 그 행위가
    #   바로 오염을 만든다 — 사고가 사고를 먹여 살린다. 가르는 선은 `real_latest.py`
    #   가 적어 둔 '없음확인'이고, 없으면 믿을 수 있는 최대 번호(`hi`)다.
    contaminated = sorted(int(k) for k, v in posts.items()
                          if str(k).isdigit() and isinstance(v, dict)
                          and v.get("contaminated"))
    ceiling = (cut - 1) if cut is not None else hi
    # 유령 = 없는 것으로 확인된 번호. `absent` 표시(real_latest.py)와 천장 위 오염을 합친다.
    ghost = sorted({n for n in contaminated if n > ceiling}
                   | {int(k) for k, v in posts.items()
                      if str(k).isdigit() and isinstance(v, dict)
                      and (v.get("absent") or (cut is not None and int(k) >= cut))})
    deleted_known = [n for n in contaminated if n <= ceiling]   # 삭제 판정 — 훑지 않는다
    # `absent`(없는 번호로 확인됨)도 뺀다 — 안 빼면 유령이 '날짜없음'으로 되돌아온다.
    #
    # ★ 표시만 믿으면 안 된다 (2026-08-07 실측). Z: 옛 덤프는 매 회차 다시 병합되는데,
    #   재병합은 `contaminated` 는 지켜도 `absent` 는 지우고 지나간다. 그래서 표시해 둔
    #   유령 22건이 한 회차 만에 '날짜없음'으로 되살아나 재수집 목록 맨 앞에 섰다
    #   — 없는 번호를 다시 긁는, 바로 그 고리다.
    #   그러므로 **확인된 경계선(`cut`)을 표시보다 위에 둔다.** 경계선은 덤프가
    #   건드리지 않는 reports/밴드_확인시각.json 에 있어서 재병합을 살아남는다.
    dateless = sorted(int(k) for k, v in posts.items()
                      if str(k).isdigit() and isinstance(v, dict)
                      and not v.get("deleted") and not v.get("contaminated")
                      and not v.get("absent") and not v.get("created_at")
                      and (cut is None or int(k) < cut))
    # 오염 표시된 번호는 '구멍'으로도 잡지 않는다(키는 있으므로 원래도 안 잡히지만,
    # 나중에 키를 지우는 방식으로 바뀌어도 여기서 한 번 더 막힌다).
    bad = set(contaminated)
    gaps = [n for n in gaps if n not in bad]
    stale = [n for n in stale if n not in bad]
    return {"band": band, "range": (lo, hi), "cache_hi": hi_cache, "n": len(ks),
            "new": new, "gaps": gaps, "stale": stale, "deleted": dead,
            "dateless": dateless, "contaminated": contaminated,
            "deleted_known": deleted_known, "ghost": ghost,
            # 왜 위쪽을 그만큼만 보는지 — 목록을 받는 쪽이 사람에게 설명할 수 있어야 한다.
            "absent_cut": cut, "absent_why": absent_why, "probing": probing}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band", help="특정 밴드만")
    ap.add_argument("--floor", type=int, default=None,
                    help="이 번호부터 전부 대상(1이면 밴드 개설 이후 전량)."
                         " 안 주면 collect_scope.json 의 사람 결정을 쓴다")
    ap.add_argument("--ahead", type=int, default=None,
                    help="캐시 최대 번호 위로 새 글을 몇 개까지 찾아볼까(기본 40)")
    ap.add_argument("--limit", type=int, default=60,
                    help="한 번에 훑을 개수(글당 5초+ — 60이면 약 7분)")
    a = ap.parse_args()
    sc = scope()

    bands = [a.band] if a.band else sorted(
        f[:-5] for f in os.listdir(CACHE)
        if f.endswith(".json") and f[:-5].isdigit())
    for band in bands:
        posts = load(band)
        if posts is None:
            print(f"{band}: 캐시 없음")
            continue
        floor = a.floor if a.floor is not None else \
            int((sc.get("floor") or {}).get(band, 0) or 0)
        ahead = a.ahead if a.ahead is not None else int(sc.get("ahead") or 40)
        p = plan(band, posts, floor, ahead)
        if not p:
            # ★ 한 밴드가 나머지를 죽이면 안 된다 — 유령 밴드의 **빈 캐시**가 그렇다
            #   (make_oneclick 은 2026-08-08 에 같은 자리를 막았는데 여기는 안 막혀
            #   있어서 `python band/recheck_plan.py` 가 통째로 죽고 있었다).
            print(f"{band}: 계획을 세울 수 없다(빈 캐시일 수 있다)")
            continue
        print(f"밴드 {band}: 보유 {p['n']}건 ({p['range'][0]}~{p['range'][1]}) · "
              f"새 글 후보 {len(p['new'])} · 구멍 {len(p['gaps'])}(floor {floor}) · "
              f"재수집 전 {len(p['stale'])}"
              + (f" · 날짜없음 {len(p['dateless'])}" if p.get("dateless") else "")
              + (f" · 오염=삭제됨(안 훑음) {len(p['deleted_known'])}"
                 if p.get("deleted_known") else "")
              + (f" · 유령(없던 번호) {len(p['ghost'])}" if p.get("ghost") else "")
              + (f" · 삭제됨 {p['deleted']}" if p.get("deleted") else ""))
        # 위쪽을 왜 그만큼만 보는지 한 줄로 말한다 — 목록만 보고는 알 수 없다.
        print(f"  · 위쪽 근거: {p.get('absent_why')}"
              + (f" → 존재 확인용 {len(p['new'])}건만(최대 {PROBE_AHEAD})"
                 if p.get("probing") else " → 확인된 구간 아래만 훑는다"))
        if p.get("cache_hi", p["range"][1]) > p["range"][1] and _absent_from(band) is None:
            # 캐시가 실재보다 위에 있다는 사실 자체를 숨기지 않는다 — 이게 사고의 흔적이다.
            print(f"  ※ 캐시 최대 {p['cache_hi']} > 믿을 수 있는 최대 {p['range'][1]}"
                  f" — 그 사이는 유령이다. python band/real_latest.py 로 확인·표시할 것")
        # ★ 순서: **새 글 → 구멍(최근부터) → 재수집**.
        #   대표 보고가 쓰는 것은 최신분이다. 과거글을 먼저 훑으면 오늘 숫자가 계속 틀린다.
        #   구멍도 큰 번호(최근)부터 간다 — 옛날로 갈수록 업무 가치가 떨어진다.
        # 날짜없음은 구멍 다음이다 — 본문은 이미 있으니 아주 급하진 않지만,
        # 날짜가 채워지기 전까지는 대조에서 통째로 빠져 있다는 점은 같다.
        todo = (p["new"] + sorted(p["gaps"], reverse=True)
                + sorted(p.get("dateless") or [], reverse=True)
                + sorted(p["stale"], reverse=True))[:a.limit]
        if a.band and todo:
            print("다음 배치(JS 그대로 붙여넣기):")
            print(f"__grabStart({band}, {json.dumps(todo)})")
        elif todo:
            print(f"  다음 {min(a.limit, len(todo))}건: {todo[:10]}{' …' if len(todo) > 10 else ''}")


if __name__ == "__main__":
    main()
