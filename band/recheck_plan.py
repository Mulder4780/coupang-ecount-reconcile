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

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
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


def _confirmed_quiet(band, hi, today=None):
    """`hi` 바로 위가 '없음'으로 확인됐고, 그 확인이 아직 최근인가.

    근거는 `reports/밴드_확인시각.json` — convert_dump 가 덤프의 missing·notime 지문에서
    만든다. session_handoff 의 '조용함' 판정과 **같은 파일**을 본다(두 곳이 어긋나면
    한쪽은 긁으라 하고 다른 쪽은 조용하다고 해서 사람이 무엇을 믿을지 모르게 된다).
    """
    p = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "reports", "밴드_확인시각.json")
    try:
        rec = (json.load(open(p, encoding="utf-8")) or {}).get(str(band)) or {}
    except Exception:
        return False
    if int(rec.get("없음확인") or 0) != int(hi) + 1:
        return False
    seen = str(rec.get("확인시각") or "")[:10]
    if not seen:
        return False
    try:
        day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
        age = (datetime.strptime(day, "%Y-%m-%d") - datetime.strptime(seen, "%Y-%m-%d")).days
    except ValueError:
        return False
    return 0 <= age <= QUIET_LIMIT_DAYS


def plan(band, posts, floor=0, ahead=0):
    ks = sorted(int(k) for k in posts if str(k).isdigit())
    if not ks:
        return None
    lo, hi = ks[0], ks[-1]
    have = set(ks)
    # ★ 캐시 **위쪽**(새 글)을 먼저 본다. 예전에는 구멍만 봐서, 마지막 수집 이후 올라온
    #   글이 영원히 대상 밖이었다 — 2026-08-06 쿠팡AS 밴드가 8/4 에 멈춰 있는데도
    #   "구멍 0" 이라 아무도 몰랐고, 8/5 돌발AS 가 1건으로 보고됐다.
    new = list(range(hi + 1, hi + 1 + int(ahead or 0)))
    # ★ 방금 '없다'고 확인한 번호를 또 훑지 않는다 (2026-08-07 지시).
    #   글 번호는 이어지므로 **hi+1 이 없으면 그 위도 전부 없다.** 그런데도 매 회차
    #   40번을 헛짚었고(글당 5초 → 3분), 그 헛짚음이 바로 오늘 오염 사고의 입구였다.
    #   확인이 오래되면 그 사이 새 글이 올라왔을 수 있으므로 그때는 다시 훑는다 —
    #   판정 기준은 session_handoff 의 '조용함' 과 같은 근거 파일 하나다.
    if new and _confirmed_quiet(band, hi):
        new = []
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
    dateless = sorted(int(k) for k, v in posts.items()
                      if str(k).isdigit() and isinstance(v, dict)
                      and not v.get("deleted") and not v.get("created_at"))
    return {"band": band, "range": (lo, hi), "n": len(ks), "new": new,
            "gaps": gaps, "stale": stale, "deleted": dead, "dateless": dateless}


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
        print(f"밴드 {band}: 보유 {p['n']}건 ({p['range'][0]}~{p['range'][1]}) · "
              f"새 글 후보 {len(p['new'])} · 구멍 {len(p['gaps'])}(floor {floor}) · "
              f"재수집 전 {len(p['stale'])}"
              + (f" · 날짜없음 {len(p['dateless'])}" if p.get("dateless") else "")
              + (f" · 삭제됨 {p['deleted']}" if p.get("deleted") else ""))
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
