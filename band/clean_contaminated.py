# -*- coding: utf-8 -*-
"""
clean_contaminated.py — 캐시에 들어앉은 **가짜 글 기록**을 걷어낸다 (2026-08-07 지시)

무엇이 가짜인가
  밴드는 `/band/<밴드>/post/<번호>` 를 iframe 으로 열면 **피드로 되돌린다**(2026-08-07 실측).
  그러면 껍데기에 직전 화면(피드 맨 위 글) 본문이 그대로 남고, 수집기가 그것을 뜯어
  `ok` 로 저장했다. 그래서 서로 다른 번호 수백 개가 **같은 본문**을 갖게 됐다.
  · 84789192: 날짜없음 98건 → 서로 다른 본문 **2종**
  · 90610953: 날짜없음 523건 → 서로 다른 본문 **7종**
  진짜라면 621종이어야 한다. 즉 621건 전부가 남의 본문을 베껴온 기록이다.

  ★ 예전 진단("밴드가 본문을 먼저 칠하고 시각을 뒤에 채워서 날짜가 빈다")은 **틀렸다.**
    그 설명대로면 본문은 제각각이어야 한다. 실제로는 본문까지 남의 것이었다.

판정 기준 (좁게 잡는다 — 진짜 글을 지우면 되돌릴 수 없다)
  ① 작성일이 없다(created_at 없음)  **그리고**
  ② 같은 본문을 가진 '작성일 없는 글'이 2건 이상이다
  하나뿐인 것은 진짜 글이 늦게 그려진 경우일 수 있으므로 **건드리지 않는다.**

무엇을 남기나
  지우지 않고 `contaminated: true` 로 **표시**한다. 키를 지우면 recheck_plan 이 그 번호를
  '구멍'으로 보고 매 회차 다시 훑는데, 지금 수집 경로로는 절대 못 가져오므로 영원히 헛돈다.
  표시만 해 두면 "모은 것도 아니고 다시 훑지도 않는" 상태로 **보이는 채** 남는다.
  가짜 본문·글쓴이는 지운다 — 남겨 두면 대조가 그것을 진짜로 쓴다.

  python band/clean_contaminated.py            # 무엇이 걸리는지만 (쓰지 않음)
  python band/clean_contaminated.py --apply    # 표시하고 저장
"""
import collections
import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def find(posts):
    """가짜로 판정된 번호 → 그 본문 지문."""
    dateless = {k: v for k, v in posts.items()
                if isinstance(v, dict) and not v.get("deleted")
                and not v.get("created_at") and not v.get("contaminated")}
    sig = {k: (v.get("content") or "")[:120] for k, v in dateless.items()}
    n = collections.Counter(s for s in sig.values() if s)
    return {k: s for k, s in sig.items() if n[s] >= 2}


def run(apply=False):
    total = 0
    for path in sorted(glob.glob(os.path.join(CACHE, "*.json"))):
        band = os.path.basename(path)[:-5]
        if not band.isdigit():                     # raw_* 같은 중간 파일은 건드리지 않는다
            continue
        doc = json.load(open(path, encoding="utf-8"))
        posts = doc.get("posts") or {}
        bad = find(posts)
        kinds = len(set(bad.values()))
        alive = sum(1 for v in posts.values()
                    if isinstance(v, dict) and not v.get("deleted")
                    and not v.get("contaminated"))
        print(f"{band}: 가짜 {len(bad)}건 (본문 {kinds}종) · 정상 보유 {alive - len(bad)}건")
        total += len(bad)
        if not (apply and bad):
            continue
        for k in bad:
            # 번호는 남기고 내용만 지운다 — 다시 훑지 않게, 그러나 없던 일로도 하지 않게.
            posts[k] = {"contaminated": True,
                        "captured_at": posts[k].get("captured_at") or 0,
                        "why": "iframe 리다이렉트로 피드 본문이 잡힌 가짜 기록(2026-08-07)"}
        tmp = path + ".tmp"                        # 정본은 원자적으로 갈아끼운다(사고 24)
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False)
        os.replace(tmp, path)
        print(f"  → 표시 완료 · {os.path.basename(path)}")
    if not apply:
        print(f"\n합계 {total}건 — 실제로 표시하려면 --apply")
    return 0


if __name__ == "__main__":
    sys.exit(run("--apply" in sys.argv[1:]))
