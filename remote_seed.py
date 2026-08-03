# -*- coding: utf-8 -*-
"""
remote_seed.py — 리모컨 재고표·입출고 관리대장(사람 작성분)을 DB 정본으로 옮긴다
================================================================================
출처: 사용자 제공 엑셀 2장 (2026-08-04 지시)
  ① 리모컨 재고 및 보유 현황 — 기준일 2026-07-29 / 07-31
  ② 리모컨 입·출고 및 사용 관리대장 — 2026-07-29 ~ 07-31 12행

**왜 스크립트로 남기나**: 손으로 한 번 넣고 끝내면 다음 세션이 근거를 못 찾는다.
여기 적힌 표가 곧 "그때 사람이 무엇을 셌는가"의 사본이고, 재실행해도 두 번 들어가지
않는다(같은 note 표식이 있으면 건너뛴다).

이후의 입·출고는 **앱 업무센터에서 입력**한다 — 이 파일은 개시 잔량 전용이다.

실행
  python remote_seed.py            # 미반영이면 넣고, 이미 있으면 아무것도 안 한다
  python remote_seed.py --status   # 현재 DB 상태만 출력
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import ledger_db as L

TAG = "[재고표 2026-07-29 이관]"          # 재실행 판별 표식
WHO = "류지영"                            # 재고표 작성·취합 주체

# ── ① 지점 기초 재고(2026-07-29 오전 기준) ─────────────────────────────
# 증평은 버전이 갈린다: 기존형 22 + (07-31 입고) VER.4 50.
BRANCH_OPEN = [
    ("부산", 21, "미확인", "2026-07-29", "기초재고 — 7월 29일 오전 기준"),
    ("시화", 0, "미확인", "2026-07-29", "기초재고 — 재고 없음"),
    ("증평", 22, "기존형", "2026-07-29", "기초재고 — 7월 29일 오전 기준"),
]

# ── ② 개인(AS 담당자) 개시 보유 ────────────────────────────────────────
# 김필우·김준형의 10개는 재고표에 '시화 입고분 지급 추정'으로 적혀 있다. 추정이라는
# 사실을 지우지 않고 note 에 그대로 남긴다 — 확인되면 사람이 고칠 근거가 된다.
TECH_OPEN = [
    ("차동호", 4, "2026-07-29", "", "기초 보유 — 7월 29일 오전 기준"),
    ("김필우", 10, "2026-07-29", "시화", "기초 보유 — 시화공장 20개 입고분 지급 추정"),
    ("김준형", 10, "2026-07-29", "시화", "기초 보유 — 시화공장 20개 입고분 지급 추정"),
]

# ── ③ 07-30~07-31 실제 움직임 ──────────────────────────────────────────
# 담당자 손을 거친 건 납품(보유 차감), 공장에서 바로 나간 건 재고 차감으로 나눈다.
DELIVERIES = [
    ("김필우", "UJ2601191", "구리3MB(배양리)", 1, "2026-07-30", "사용", "미확인",
     "리모컨 1EA 사용 — 개인 보유 10→9"),
    ("김준형", "UJ2601330", "남양주4MB(이패동)", 1, "2026-07-30", "납품", "미확인",
     "02호기 납품 — 개인 보유 10→9"),
    ("차동호", "UJ2601360", "창원1MB(두동)", 1, "2026-07-31", "교체", "미확인",
     "리모컨 1EA 교체 — 개인 보유 4→3"),
]
STOCK_MOVES = [
    ("증평", -2, "기존형", "2026-07-30",
     "현장출고(택배) 제주1Sub-hub(장전로) UJ2601354 — 건전지 교체 후 작동 불가"),
    ("증평", 50, "VER.4", "2026-07-31", "신규 리모컨 50EA 입고(업체→증평공장 자재실)"),
    ("증평", -1, "VER.4", "2026-07-31", "샘플출고 — 케이스 제작처 송부"),
]


def already_loaded():
    with L.conn() as c:
        n = c.execute(
            "SELECT COUNT(*) FROM remote_issue WHERE note LIKE ?", (f"%{TAG}%",)).fetchone()[0]
        m = c.execute(
            "SELECT COUNT(*) FROM remote_stock WHERE reason LIKE ?", (f"%{TAG}%",)).fetchone()[0]
    return n + m > 0


def load():
    for branch, qty, version, day, why in BRANCH_OPEN:
        if qty:
            L.remote_stock_adjust(branch, qty, "add", f"{why} {TAG}", WHO,
                                  version=version, moved_on=day)
        else:
            # 0개도 '세어 봤다'는 기록이 필요하다 — 미등록과 재고0은 다른 상태다.
            L.remote_stock_adjust(branch, 1, "add", f"{why} {TAG}", WHO,
                                  version=version, moved_on=day)
            L.remote_stock_adjust(branch, -1, "add", f"{why}(0개 확정) {TAG}", WHO,
                                  version=version, moved_on=day)
    for tech, qty, day, branch, why in TECH_OPEN:
        L.remote_open_balance(tech, qty, day, f"{why} {TAG}", WHO,
                              branch=branch, version="미확인")
    for tech, prj, camp, qty, day, kind, version, why in DELIVERIES:
        L.remote_deliver(tech, prj, camp, qty, day, f"{why} {TAG}", WHO,
                         kind=kind, version=version)
    for branch, delta, version, day, why in STOCK_MOVES:
        L.remote_stock_adjust(branch, delta, "add", f"{why} {TAG}", WHO,
                              version=version, moved_on=day)


def show():
    s = L.remote_status()
    print(f"보유 합계 {s['totals']['holding']} · 지점 재고 {s['totals']['stock']}"
          f" · 전체 {s['totals']['all']}")
    for tech, h in sorted(s["holdings"].items()):
        flag = "  ← 한도 초과" if h["holding"] > s["limit"] else ""
        print(f"  {tech}: 보유 {h['holding']} (받음 {h['issued']} · 나감 {h['delivered']}){flag}")
    for br, b in s["branch_stock"].items():
        vers = " · ".join(f"{k} {v}" for k, v in (b["versions"] or {}).items())
        print(f"  {b['label']}: {b['stock']}개" + (f" ({vers})" if vers else ""))


def main():
    if "--status" in sys.argv:
        show()
        return
    if already_loaded():
        print(f"이미 반영됨 {TAG} — 아무것도 하지 않았습니다")
        show()
        return
    load()
    print(f"재고표·관리대장 반영 완료 {TAG}")
    show()


if __name__ == "__main__":
    main()
