# -*- coding: utf-8 -*-
"""
real_latest.py — 밴드의 **진짜 최신 글 번호**를 확인해 기록한다 (2026-08-07 지시)

왜 필요한가
  캐시가 가진 가장 큰 번호(`hi`)를 "여기까지 있다"로 믿으면 안 된다.
  2026-08-07 오염 사고로 **존재하지 않는 번호**가 캐시에 들어앉았고, 그 유령이 `hi` 를
  밀어 올렸다. 그러면 recheck_plan 이 `hi+1` 부터 새 글 후보를 만드는데 —
  그 번호들은 **처음부터 없는 번호**다. 밴드는 없는 번호에도 200 과 앱 껍데기를
  주므로, 긁으면 직전 화면 본문이 또 잡힌다. 즉 **사고가 사고를 먹여 살린다.**

    90610953: 캐시 hi 5464 · 실제 최신 5437 → 5438~5464 스물두 개가 유령
              그 위 5465~5504 마흔 개를 매 회차 헛짚고 있었다

  그래서 CLAUDE.md 가 못박는다 — **"밴드를 긁기 전에 먼저 그 밴드의 실제 최신 글
  번호를 확인한다."** 이 도구가 그 확인을 받아 적는 자리다.

무엇을 하나
  ① `reports/밴드_확인시각.json` 에 `없음확인 = 최신+1` 로 적는다.
     convert_dump 가 덤프 지문에서 만드는 것과 **같은 파일·같은 열쇠**다.
     session_handoff 의 '조용함' 과 recheck_plan 의 '새 글 후보'가 이 한 장을 같이 본다
     (두 곳이 어긋나면 한쪽은 긁으라 하고 다른 쪽은 조용하다고 해서 사람이 헷갈린다).
  ② 캐시에서 **최신 위쪽**에 있는 항목을 `absent: true` 로 표시한다.
     지우지 않는다 — 지우면 recheck_plan 이 '구멍'으로 보고 영원히 다시 훑는다.

안전장치 (진짜 글은 절대 건드리지 않는다)
  `created_at` 이 있는 항목은 **표시하지 않는다.** 작성시각이 있다는 것은 실제로
  열려서 수확된 진짜 글이라는 뜻이다. 그런 것이 최신 위에 있으면 내가 읽은 '최신'
  쪽이 틀린 것이므로, 표시 대신 **경고를 띄우고 멈춘다.**

최신 번호는 어떻게 얻나 (사람/AI 가 브라우저에서 한 줄)
    // https://www.band.us/band/<밴드>/post 를 연 뒤
    Math.max(...[...document.querySelectorAll('a[href*="post/"]')]
      .map(e => +((e.getAttribute('href')||'').match(/post\\/(\\d+)/)||[0,0])[1]))

사용
  python band/real_latest.py --band 90610953 --latest 5437           # 무엇이 바뀌는지만
  python band/real_latest.py --band 90610953 --latest 5437 --apply   # 기록·표시
  python band/real_latest.py                                          # 지금 기록 상태
  python band/real_latest.py --heal --apply    # 추월된 근거를 스스로 무효로(워치독 30분)

★ 사람 손은 `--latest` 한 줄뿐이고, **틀린 근거를 되돌리는 일은 이제 자동이다**
  (`heal()` · 2026-08-11). 근거가 "N 부터 없다"는데 N 이상이 이미 수확돼 있으면
  그것은 모순이므로 기계가 '모름'으로 되돌린다 — 없는 근거를 지어내지는 않는다.
"""
import argparse
import json
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE = os.path.join(HERE, "cache")
SEEN = os.path.join(ROOT, "reports", "밴드_확인시각.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _load(path, default):
    try:
        return json.load(open(path, encoding="utf-8")) or default
    except Exception:
        return default


def _save_atomic(path, doc):
    """정본은 원자적으로 갈아끼운다(사고 24) — 쓰다 죽어도 반쪽 파일이 남지 않는다."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _rp():
    """'믿을 수 있는 수확'의 정의는 **판정하는 자리에서 빌린다**(`recheck_plan`).

    여기서 따로 세면 언젠가 갈린다 — 실제로 갈려 있었다. 이 파일은 `absent` 만 걸렀고
    `trusted_hi` 는 `contaminated`·`deleted` 까지 거른다. 오염 항목도 작성시각을 갖는
    일이 있어(남의 본문을 통째로 베껴 왔으므로) 이쪽 계산만 위로 떠올랐다.
    **정정하는 쪽과 판정하는 쪽이 서로 다른 최대값을 보면 고쳐도 안 고쳐진다.**
    """
    if HERE not in sys.path:
        sys.path.insert(0, HERE)
    import recheck_plan
    return recheck_plan


def _posts(band):
    return (_load(os.path.join(CACHE, f"{band}.json"), {}) or {}).get("posts") or {}


def _collected_max(band):
    """캐시에서 **실제로 본문을 받아 둔** 가장 큰 글 번호(유령·오염·시각없음 제외)."""
    return _rp().trusted_hi(_posts(band))


def heal(apply=False, today=None):
    """**추월된 근거**를 기계가 스스로 무효로 만든다 (2026-08-11 지시: "자동화 시켜").

    `[217]` 은 추월된 근거를 **읽는 쪽 둘**(수집 계획·인계 문서)이 거르게 했다.
    그런데 근거 자체는 틀린 채로 남아 매 회차 같은 거름질을 다시 시키고, 바로잡는
    길은 사람이 밴드 피드를 열어 `--latest` 를 적어 주는 것뿐이었다.
    **그 한 줄이 이 사고에 남은 마지막 사람 몫이었다.**

    ★ **지어낼 것이 없다.** 근거가 "N 부터 없다"는데 캐시에 N 이상의 **진짜 글**
      (작성시각이 있고 죽지 않은 수확)이 있으면 그 근거는 틀렸다 — 이건 판단이 아니라
      **모순**이다. 그래서 기계가 지울 수 있다. 낡은 근거는 여기서 안 건드린다:
      낡은 것은 틀린 것이 아니라 **다시 물어봐야 하는 것**이고, 그 물음은 다음 회차의
      `PROBE_AHEAD` 다섯 건이 싸게 한다.
    ★ **없음확인을 위로 올리지 않는다.** `top+1` 부터 없다는 근거는 아무 데도 없다.
      '모른다'로 되돌릴 뿐이다 — 그러면 다음 회차가 다섯 건으로 물어보고 사다리는
      저절로 오른다. 없는 근거를 지어내면 `[217]` 을 손수 다시 만드는 것이다.
    ★ **캐시는 한 글자도 안 고친다.** 고치는 것은 근거 한 장뿐이다. 틀린 근거로 수확에
      `absent` 를 찍으면 실재하는 글이 유령이 된다 — 되돌릴 수 없는 쪽이다.
    ★ 옛 값을 버리지 않는다(`이전없음확인`) — "그때 무엇을 믿고 있었나"를 잃으면
      같은 사고를 또 봐도 못 알아본다.

    돌려주는 값: 정정된 밴드 목록 `[{밴드, 이전, 실제수확}]`. 고칠 것이 없으면 빈 목록.
    """
    seen = _load(SEEN, {})
    if not isinstance(seen, dict):
        return []
    day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
    hi_of = _rp().trusted_hi
    fixed = []
    for band, rec in sorted(seen.items()):
        if not isinstance(rec, dict):
            continue
        try:
            n = int(rec.get("없음확인") or 0)
        except (TypeError, ValueError):
            n = 0
        if n <= 0:
            continue                       # 이미 '모름' — 고칠 것이 없다
        top = hi_of(_posts(band))
        if top is None or top < n:
            continue                       # 모순 없음
        fixed.append({"밴드": band, "이전": n, "실제수확": top})
        seen[band] = dict(rec, 없음확인=0, 수집최대=top,
                          이전없음확인=n, 이전확인시각=rec.get("확인시각"),
                          확인시각=day,
                          근거="추월돼 무효 — %s 부터 없다 했으나 %s 가 실제로 수확됨"
                               " (자동 정정 real_latest.heal)" % (n, top))
    if fixed and apply:
        _save_atomic(SEEN, seen)
    return fixed


def survey(band, latest):
    """`latest` 위쪽 캐시 항목을 갈래별로 센다."""
    posts = _posts(band)
    above = [int(k) for k in posts
             if str(k).isdigit() and int(k) > int(latest)]
    real, ghost, already = [], [], []
    for n in sorted(above):
        v = posts.get(str(n)) or {}
        if not isinstance(v, dict):
            continue
        if v.get("absent"):
            already.append(n)
        elif v.get("created_at"):
            real.append(n)                 # ★ 진짜 글이 최신 위에 있다 = 내 '최신'이 틀렸다
        else:
            ghost.append(n)
    return {"real": real, "ghost": ghost, "already": already}


def run(band=None, latest=None, apply=False, today=None):
    seen = _load(SEEN, {})
    if band is None or latest is None:                    # 상태만 보여 준다
        if not seen:
            print("기록 없음 — --band/--latest 로 확인 결과를 적는다")
        for b, rec in sorted(seen.items()):
            print(f"{b}: 없음확인 {rec.get('없음확인')} · {rec.get('확인시각')}"
                  f" · {rec.get('근거') or ''}")
        return 0

    band, latest = str(band), int(latest)
    s = survey(band, latest)
    print(f"밴드 {band} · 실제 최신 {latest}")
    print(f"  최신 위 캐시 항목: 유령 {len(s['ghost'])} · 이미표시 {len(s['already'])}"
          f" · 작성시각 있는 진짜 글 {len(s['real'])}")

    if s["real"]:
        # 여기서 멈추는 것이 핵심이다. 진짜 글을 '없음'으로 덮으면 되돌릴 수 없다.
        print(f"  ✗ 중단 — 최신보다 큰 번호에 **작성시각이 있는 글**이 있다: "
              f"{s['real'][:8]}{' …' if len(s['real']) > 8 else ''}")
        print("     읽어 온 '최신 글 번호'가 틀렸을 가능성이 높다. 피드를 다시 확인할 것.")
        return 2

    if not apply:
        print("  (쓰지 않음 — 실제로 기록하려면 --apply)")
        return 0

    day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
    # ★ `수집최대` 를 빠뜨리지 말 것 (2026-08-07). session_handoff 의 '(조용함)' 줄이
    #   이 값을 읽어 "N번까지 수집 완료" 를 찍는다. 없으면 그 자리에 None 이 박히고,
    #   다음 세션은 어디까지 모았는지를 인계 문서에서 못 읽는다. 예전 기록을 덮어쓸 때
    #   조용히 사라졌다 — 그래서 여기서 매번 다시 계산해 넣는다.
    seen[band] = {"없음확인": latest + 1, "확인시각": day,
                  "수집최대": _collected_max(band),
                  "근거": "밴드 피드 최상단 글 번호 직접 확인(real_latest.py)"}
    _save_atomic(SEEN, seen)
    print(f"  → 없음확인 {latest + 1} · {day} 기록")

    if s["ghost"]:
        path = os.path.join(CACHE, f"{band}.json")
        doc = _load(path, None)
        posts = (doc or {}).get("posts") or {}
        for n in s["ghost"]:
            v = posts.get(str(n)) or {}
            # ★ 원래 표시를 지우지 말 것. `contaminated` 를 떨어뜨렸더니 그 번호들이
            #   이번엔 '날짜없음'으로 다시 잡혀 재수집 목록에 되돌아왔다(실측).
            #   표시를 바꾸는 도구는 **더하기만** 한다.
            keep = {k: v[k] for k in ("contaminated", "deleted", "why")
                    if isinstance(v, dict) and v.get(k)}
            keep.update({"absent": True,
                         "captured_at": (v or {}).get("captured_at") or 0,
                         "why": f"실제 최신({latest})보다 큰 번호 — 존재하지 않는 글"})
            posts[str(n)] = keep
        _save_atomic(path, doc)
        print(f"  → 유령 {len(s['ghost'])}건 absent 표시 · {band}.json")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="밴드의 진짜 최신 글 번호를 기록한다")
    ap.add_argument("--band")
    ap.add_argument("--latest", type=int)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--heal", action="store_true",
                    help="추월된 근거를 스스로 무효로 만든다(워치독 30분 회차가 부른다)")
    a = ap.parse_args(argv)
    if a.heal:
        fixed = heal(apply=a.apply)
        if not fixed:
            print("밴드 근거 정상 — 추월된 것 없음")
        for f in fixed:
            print("%s: 없음확인 %s → 모름 (%s 가 실제로 수확됨)%s"
                  % (f["밴드"], f["이전"], f["실제수확"],
                     "" if a.apply else "  (쓰지 않음 — --apply)"))
        return 0
    return run(a.band, a.latest, a.apply)


if __name__ == "__main__":
    sys.exit(main())
