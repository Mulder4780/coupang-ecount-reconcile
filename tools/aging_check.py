# -*- coding: utf-8 -*-
"""
aging_check.py — 미청구가 얼마나 늙었는지 매일 본다 (GitHub Actions 전용)
================================================================================
왜 사무실 PC가 아니라 여기서 도는가:
  출장으로 **PC를 며칠 못 켜도 미청구는 계속 늙는다.** PC가 꺼져 있으면 아무도
  그걸 세지 않는다. 경과일은 새 데이터 없이 날짜만으로 계산되므로, 클라우드에서
  매일 세는 게 맞다. GitHub Actions는 공개 저장소에서 무료다.

무엇을 보는가:
  docs/aging.json — **날짜만** 들어 있다(금액·PO번호·현장명 없음, 저장소가 공개라서).
  상세는 잠긴 사본(data.enc)에 있고 폰에서 PIN을 넣어야 보인다.

어떻게 알리는가:
  기준을 넘으면 **일부러 실패로 끝낸다(exit 1)**. 그러면 GitHub가 저장소 주인에게
  자동으로 메일을 보낸다 — 메일 비밀번호나 외부 서비스가 전혀 필요 없다.

  python tools/aging_check.py            # 판정(넘으면 exit 1)
  python tools/aging_check.py --dry      # 판정만 하고 항상 exit 0
"""
import sys, os, json
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "aging.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

LABEL = {"po": "미청구 PO(계산서 미발행)"}


def load(path=SRC):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def evaluate(doc, today=None):
    """(경고 목록, 최대 경과일, 기준). 오늘 날짜만으로 계산한다."""
    today = today or date.today()
    warn, crit = int(doc.get("warn_days", 90)), int(doc.get("crit_days", 120))
    rows = []
    for it in doc.get("items", []):
        try:
            d = datetime.strptime(it["since"], "%Y-%m-%d").date()
        except Exception:
            continue
        days = (today - d).days
        if days >= warn:
            rows.append({"kind": it.get("k", "?"), "since": it["since"], "days": days,
                         "level": "심각" if days >= crit else "경고"})
    rows.sort(key=lambda r: -r["days"])
    return rows, (max((r["days"] for r in rows), default=0)), (warn, crit)


def main():
    if not os.path.exists(SRC):
        print(f"::warning::{os.path.relpath(SRC, ROOT)} 이 없습니다 — 사무실 PC가 아직 한 번도 올리지 않았습니다.")
        return 0

    doc = load()
    rows, worst, (warn, crit) = evaluate(doc)
    total = len(doc.get("items", []))
    stale = (date.today() - datetime.strptime(doc.get("generated", "1970-01-01"),
                                              "%Y-%m-%d").date()).days

    print(f"기준일 {date.today()} · 자료 갱신 {doc.get('generated')} ({stale}일 전)")
    print(f"미청구 {total}건 · 그중 {warn}일 넘은 것 {len(rows)}건 (최장 {worst}일)")
    for r in rows:
        print(f"  [{r['level']}] {LABEL.get(r['kind'], r['kind'])} — {r['since']} 발행 · {r['days']}일 경과")

    # 자료가 너무 오래됐으면 그 자체가 문제다(PC가 2주 넘게 안 켜졌다는 뜻)
    if stale > 14:
        print(f"::warning::자료가 {stale}일째 갱신되지 않았습니다 — 사무실 PC에서 한 번 돌려 주세요.")

    if not rows:
        print("\n기준을 넘은 건이 없습니다.")
        return 0

    print(f"\n금액·현장은 여기 담지 않습니다(공개 저장소). 상세는 앱에서 PIN 넣고 확인하세요:")
    print("  https://mulder4780.github.io/coupang-ecount-reconcile/")

    if "--dry" in sys.argv:
        return 0
    # ★ 일부러 실패시킨다 — GitHub가 주인에게 메일을 보내게 하는 유일한 무료 경로다.
    print(f"\n::error::미청구 {len(rows)}건이 {warn}일을 넘었습니다 (최장 {worst}일). 앱에서 확인하세요.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
