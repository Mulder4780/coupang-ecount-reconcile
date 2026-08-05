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


def plan(band, posts, floor=0):
    ks = sorted(int(k) for k in posts if str(k).isdigit())
    if not ks:
        return None
    lo, hi = ks[0], ks[-1]
    have = set(ks)
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
    return {"band": band, "range": (lo, hi), "n": len(ks),
            "gaps": gaps, "stale": stale, "deleted": dead}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band", help="특정 밴드만")
    ap.add_argument("--floor", type=int, default=0,
                    help="이 번호부터 전부 대상(1이면 밴드 개설 이후 전량)")
    ap.add_argument("--limit", type=int, default=60,
                    help="한 번에 훑을 개수(글당 5초+ — 60이면 약 7분)")
    a = ap.parse_args()

    bands = [a.band] if a.band else sorted(
        f[:-5] for f in os.listdir(CACHE)
        if f.endswith(".json") and f[:-5].isdigit())
    for band in bands:
        posts = load(band)
        if posts is None:
            print(f"{band}: 캐시 없음")
            continue
        p = plan(band, posts, a.floor)
        print(f"밴드 {band}: 보유 {p['n']}건 ({p['range'][0]}~{p['range'][1]}) · "
              f"구멍 {len(p['gaps'])} · 재수집 전 {len(p['stale'])}"
              + (f" · 삭제됨 {p['deleted']}" if p.get("deleted") else ""))
        # 최신 글부터 재확인한다 — 수정은 최근 글에서 일어난다.
        todo = (p["gaps"] + sorted(p["stale"], reverse=True))[:a.limit]
        if a.band and todo:
            print("다음 배치(JS 그대로 붙여넣기):")
            print(f"__grabStart({band}, {json.dumps(todo)})")
        elif todo:
            print(f"  다음 {min(a.limit, len(todo))}건: {todo[:10]}{' …' if len(todo) > 10 else ''}")


if __name__ == "__main__":
    main()
