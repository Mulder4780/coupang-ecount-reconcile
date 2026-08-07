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


def _collected_max(band):
    """캐시에서 **실제로 본문을 받아 둔** 가장 큰 글 번호.

    유령(`absent`)·시각 없는 항목은 세지 않는다 — 그것들은 '모은 것'이 아니다.
    """
    posts = (_load(os.path.join(CACHE, f"{band}.json"), {}) or {}).get("posts") or {}
    got = [int(k) for k, v in posts.items()
           if str(k).isdigit() and isinstance(v, dict)
           and v.get("created_at") and not v.get("absent")]
    return max(got) if got else None


def survey(band, latest):
    """`latest` 위쪽 캐시 항목을 갈래별로 센다."""
    posts = (_load(os.path.join(CACHE, f"{band}.json"), {}) or {}).get("posts") or {}
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
    a = ap.parse_args(argv)
    return run(a.band, a.latest, a.apply)


if __name__ == "__main__":
    sys.exit(main())
