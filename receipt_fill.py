# -*- coding: utf-8 -*-
"""receipt_fill.py — 입금(수금) 자동입력.

왜: 폰 앱 상세의 '빈 항목 입력' 칸(PO번호·PO발행일·청구일·지급예정일·입금일·입금액) 중
PO 두 칸은 `po_reconcile` 이 이미 자동으로 채운다. **입금일·입금액만 사람이 손으로** 넣고
있었다. 근거 자료(거래처별계정별원장)가 들어오면 사람 손을 거치지 않게 한다
(사용자 지시 2026-07-28: "자료 확인되면 알아서 입력 자동화").

무엇을 채우나 — **입금일·입금액만** 채운다.
  · 근거: 거래처별계정별원장(외상매출금)의 **대변 = 입금**. 차변은 매출 발생이라 건드리지 않는다.
  · 청구일·지급예정일은 **채우지 않는다.** 사내에서 무엇을 '청구일'로 보는지(세금계산서 발행일인지
    아리바 업로드일인지), 지급예정일 산정기준이 무엇인지 확정된 규칙이 없다. 규칙 없이 날짜를
    만들어 넣으면 미수 집계가 통째로 틀어진다(AGENTS.md 절대규칙 10 — 원자료에 없는 값 임의 채움 금지).
    확정되면 여기에 한 줄 추가하면 된다.

어떻게 맞추나 — po_reconcile 과 같은 **양방향 유일 일치** 원칙.
  입금액이 06시트의 세금계산서합계(없으면 거래명세서합계)와 딱 맞고, 그런 후보가 **한 건뿐**이며,
  그 입금건과 금액이 같은 다른 입금도 없을 때만 자동입력한다. 쿠팡 입금은 여러 건을 묶어
  한 번에 넣는 일이 잦아 금액이 겹치면 엉뚱한 정산행에 붙는다 — 겹치면 사람에게 넘긴다.

실행
  python receipt_fill.py            # 미리보기(큐 적재 없음)
  python receipt_fill.py --queue    # ledger_writer 큐 적재(원장 반영은 --apply 가 한다)
"""
import os
import re
import sys
from datetime import date, datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")
sys.path.insert(0, BASE_DIR)

AMOUNT_TOL = 1          # 원 단위 반올림 차이만 허용


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _day(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.search(r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})", str(v or ""))
    if not m:
        return None
    try:
        return date(*(int(x) for x in m.groups()))
    except ValueError:
        return None


def parse_receipts(path):
    """거래처별계정별원장 → [{일자, 전표, 적요, 금액}] (대변 = 입금)

    합계/이월 행은 건너뛴다 — '월 계'·'누 계'·'전기이월' 이 금액을 갖고 있어
    그대로 두면 수천만 원짜리 가짜 입금이 하나 생긴다."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out, has_credit = [], False
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i, idx = None, {}
        for i, r in enumerate(rows[:25]):
            names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
            if any("적요" in n for n in names) and any("대변" in n for n in names):
                hdr_i = i
                for n, j in names.items():
                    if "일자" in n or "날짜" in n:
                        idx.setdefault("date", j)
                    if "No" in n or "번호" in n:
                        idx.setdefault("slip", j)
                    if "적요" in n:
                        idx["remark"] = j
                    if "대변" in n:
                        idx["credit"] = j
                break
        if hdr_i is None or "credit" not in idx:
            continue
        has_credit = True
        for r in rows[hdr_i + 1:]:
            if r is None or all(c is None for c in r):
                continue
            joined = " ".join(str(c) for c in r if c is not None)
            if re.search(r"(월|누|합)\s*계|이\s*월|소\s*계", joined):
                continue                                  # 합계·이월 행
            amt = _num(r[idx["credit"]])
            if not amt:
                continue
            d = _day(r[idx["date"]]) if idx.get("date") is not None else None
            if d is None:
                d = _day(joined)
            if d is None:
                continue                                  # 날짜 없는 입금은 쓸 수 없다
            out.append({
                "일자": d, "금액": amt,
                "전표": str(r[idx["slip"]] or "") if idx.get("slip") is not None else "",
                "적요": str(r[idx["remark"]] or "") if idx.get("remark") is not None else "",
                "출처": os.path.basename(path),
            })
    wb.close()
    return out, has_credit


def open_settlements(master):
    """입금일이 비어 있는 06시트 정산행 — 청구액(세금계산서합계 우선)과 함께."""
    from ecount_reconcile import read_ledger
    out = []
    for sid, r in read_ledger(master).items():
        if r.get("원장_입금일"):
            continue
        billed = r.get("원장_세금계산서합계") or r.get("원장_거래명세서합계")
        if not billed:
            continue
        out.append({
            "정산ID": sid, "청구액": float(billed),
            "프로젝트NO": r.get("프로젝트NO", ""), "캠프명": r.get("캠프명", ""),
            "발행일": _day(r.get("원장_세금계산서발행일") or r.get("원장_거래명세서발행일")),
        })
    return out


def match(receipts, rows):
    """양방향 유일 일치만 자동입력. 겹치면 사람에게 넘긴다."""
    paired, spare = [], []
    for rc in receipts:
        cands = [s for s in rows
                 if abs(s["청구액"] - rc["금액"]) <= AMOUNT_TOL
                 and (s["발행일"] is None or s["발행일"] <= rc["일자"])]
        rivals = [q for q in receipts if abs(q["금액"] - rc["금액"]) <= AMOUNT_TOL]
        if len(cands) == 1 and len(rivals) == 1:
            paired.append((rc, cands[0]))
        else:
            spare.append((rc, "후보 %d건 · 같은 금액 입금 %d건" % (len(cands), len(rivals))))
    return paired, spare


def main():
    queue_mode = "--queue" in sys.argv
    from ecount_reconcile import load_config, resolve_master
    from inbox_scan import pick
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])

    files = pick("ledger")
    if not files:
        print("거래처별계정별원장이 아직 없습니다 — 들어오면 자동으로 잡습니다.")
        print("  넣는 곳: '0. 원본 자료/1. ERP 내보내기' 또는 inbox/ (파일명은 아무거나)")
        print("  뽑는 곳: 회계 I > 출력물 > 장부 > 거래처별계정별원장")
        print("           거래처=쿠팡로지스틱스 · 계정=1089(외상매출금) · ★개별거래처기준 · 기간지정 → Excel")
        return 0

    receipts, usable = [], []
    for f in files:
        part, ok = parse_receipts(f)
        receipts += part
        if ok:
            usable.append(f)
    if not usable:
        # 여기서 멈추지 않으면 '입금 0건' 이라고 조용히 보고해 자료가 있는 줄 알게 된다.
        print("★ 계정별원장 형식(적요+대변)이 있는 파일이 없습니다 — %d개를 봤지만 전부 다른 표입니다."
              % len(files))
        for f in files:
            print("   -", os.path.basename(f))
        print("  회계 I > 출력물 > 장부 > 거래처별계정별원장 에서 ★개별거래처기준★ 으로 뽑아 주세요.")
        return 0
    rows = open_settlements(master)
    paired, spare = match(receipts, rows)

    print("입금 대조 — 원장 입금행 %d건 / 입금일 빈 정산 %d건 → 유일매칭 %d건 · 보류 %d건"
          % (len(receipts), len(rows), len(paired), len(spare)))
    for rc, s in paired[:20]:
        print("  · %s %s %s %s원 → %s (%s)"
              % (s["정산ID"], s["프로젝트NO"], s["캠프명"][:14],
                 format(int(rc["금액"]), ","), rc["일자"], rc["적요"][:20]))
    if spare:
        print("  [보류] %d건 — 묶음 입금이거나 금액이 겹칩니다(사람 확인)" % len(spare))
    if not receipts:
        print("  ※ 대변(입금) 행이 0건입니다 — '대표거래처로 합산' 이 켜져 있으면 캠프별 거래처의")
        print("     입금이 빠집니다. 검색 조건을 '개별거래처기준' 으로 다시 뽑아 주세요.")

    if not queue_mode:
        print("\n미리보기 — 실제 적재: python receipt_fill.py --queue")
        return 0

    items = []
    for rc, s in paired:
        ev = "거래처별계정별원장 대변 %s %d원 (금액 유일매칭)" % (rc["일자"], int(rc["금액"]))
        for sheet in ("06_거래서류청구수금", "16_입금수금관리"):
            items.append({"sheet": sheet, "key_col": "정산ID", "key": s["정산ID"],
                          "col": "입금일", "value": rc["일자"].isoformat(), "vtype": "date",
                          "evidence": ev, "only_if_empty": True})
            items.append({"sheet": sheet, "key_col": "정산ID", "key": s["정산ID"],
                          "col": "입금액", "value": str(int(rc["금액"])), "vtype": "number",
                          "evidence": ev, "only_if_empty": True})
    if not items:
        print("적재할 항목 없음")
        return 0
    from ledger_writer import queue_add
    print("큐 적재:", queue_add(items), "개 셀 → python ledger_writer.py --apply 로 원장 반영")
    return 0


if __name__ == "__main__":
    sys.exit(main())
