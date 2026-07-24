# -*- coding: utf-8 -*-
"""
po_reconcile.py — 쿠팡 발행 PO ↔ 관리대장 자동 대조
=====================================================
쿠팡(아리바/포털)에서 내려받은 PO 목록 엑셀을 inbox/에 넣으면(파일명에 'PO' 포함)
관리대장 06(PO번호·발행일)·21(PO정산연결)과 대조해 4유형을 검출한다.

  [유형A] 쿠팡PO 있는데 원장 미등록      : 받은 PO를 관리대장에 안 옮김 (누락)
  [유형B] 원장 PO번호가 쿠팡 목록에 없음  : 번호 오기입 가능성
  [유형C] 금액 불일치                    : 쿠팡 PO 금액 ≠ 원장 정산 공급가액 합
  [유형D] 미연결 제안                    : PO필요·번호없는 정산 ↔ 미등록 PO 금액 매칭 후보
          └ 금액이 양방향 유일 일치하면 자동입력 큐(06 PO번호·PO발행일, 빈 칸만)에 적재

파서는 머리글 자동탐지 + 전체 셀에서 PO번호 패턴(PO+숫자) 추출이라
쿠팡 내보내기 형식이 달라도 동작한다. 원장은 read-only. 결과는 reports/.

실행:  python po_reconcile.py                       # inbox에서 'PO' 파일 자동 탐지
       python po_reconcile.py --file 경로 [--master 경로]   # 테스트용
"""
import sys, os, re, csv, glob, json
from datetime import datetime
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_reconcile import read_ledger, load_config, _num, _d

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.path.join(BASE_DIR, "inbox")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")
PO_PAT = re.compile(r"\bPO[\s-]?(\d{3,})\b", re.I)


def norm_po(s):
    m = PO_PAT.search(str(s or ""))
    return f"PO{m.group(1)}" if m else ""


def parse_po_export(path):
    """쿠팡 PO 목록 엑셀 → [{po, date, amount, desc}] (형식 유연 파싱)"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    pos = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        # 머리글 탐지(선택적): 금액·일자 열 위치 힌트
        j_amt = j_date = None
        for r in rows[:15]:
            names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
            for n, j in names.items():
                if j_amt is None and any(k in n for k in ("금액", "공급가", "합계", "Amount", "amount", "Total")):
                    j_amt = j
                if j_date is None and any(k in n for k in ("일자", "날짜", "발행", "Date", "date")):
                    j_date = j
            if j_amt is not None:
                break
        for r in rows:
            if r is None:
                continue
            blob = " ".join(str(c) for c in r if c is not None)
            po = norm_po(blob)
            if not po:
                continue
            amt = _num(r[j_amt]) if j_amt is not None and j_amt < len(r) else None
            if amt is None:  # 금액 열을 못 찾으면 행에서 가장 큰 숫자 추정
                nums = [_num(c) for c in r if _num(c) and _num(c) > 1000]
                amt = max(nums) if nums else None
            dt = ""
            if j_date is not None and j_date < len(r):
                dt = _d(r[j_date])[:10]
            if not dt:
                m = re.search(r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})", blob)
                dt = m.group(1).replace("/", "-").replace(".", "-") if m else ""
            cur = pos.get(po)
            if not cur or (amt and not cur.get("amount")):
                pos[po] = {"po": po, "date": dt, "amount": amt,
                           "desc": blob[:120].replace("\n", " ")}
    wb.close()
    return list(pos.values())


def ledger_po_view(master):
    """원장 측 PO 시각: PO번호 → [정산행], PO없는 유상 정산 목록"""
    recs = read_ledger(master)
    by_po, no_po = defaultdict(list), []
    for sid, r in sorted(recs.items()):
        po = norm_po(r.get("원장_PO번호"))
        if po:
            by_po[po].append(r)
        elif r.get("비용구분") == "유상" and r.get("원장_공급가액") and \
                str(r.get("원장_PO필요여부", "")).startswith("필요"):
            no_po.append(r)
    return by_po, no_po


def main():
    args = sys.argv[1:]
    cfg = load_config()
    master = cfg["reconcile"]["master_xlsx"]
    if "--master" in args:
        master = args[args.index("--master") + 1]
    if "--file" in args:
        files = [args[args.index("--file") + 1]]
    else:
        files = [f for f in glob.glob(os.path.join(INBOX_DIR, "*.xlsx"))
                 if "PO" in os.path.basename(f).upper() and not os.path.basename(f).startswith("~$")
                 and "원장" not in os.path.basename(f)]
    if not files:
        sys.exit("inbox/ 에 쿠팡 PO 목록 파일이 없습니다. 파일명에 'PO'를 포함해 넣어주세요. (예: 쿠팡PO목록_7월.xlsx)")

    coupang = []
    for f in files:
        part = parse_po_export(f)
        coupang += part
        print(f"'{os.path.basename(f)}': PO {len(part)}건 파싱")
    cp_by_no = {p["po"]: p for p in coupang}

    by_po, no_po = ledger_po_view(master)
    tol = cfg["reconcile"].get("금액허용오차", 0)

    A, B, C, OK, D = [], [], [], [], []
    for po, p in sorted(cp_by_no.items()):
        lrows = by_po.get(po)
        if not lrows:
            A.append({"PO번호": po, "쿠팡금액": p["amount"], "발행일": p["date"], "내용": p["desc"][:60],
                      "판정": "원장 미등록 — 관리대장에 옮겨지지 않은 PO"})
            continue
        led_sum = sum((r.get("원장_공급가액") or 0) for r in lrows)
        ids = ",".join(r["정산ID"] for r in lrows)
        if p["amount"] is not None and abs(led_sum - p["amount"]) > tol:
            C.append({"PO번호": po, "정산ID": ids, "쿠팡금액": p["amount"], "원장공급가액합": led_sum,
                      "차액": round((p["amount"] or 0) - led_sum)})
        else:
            OK.append(po)
    for po, lrows in sorted(by_po.items()):
        if po not in cp_by_no:
            B.append({"PO번호": po, "정산ID": ",".join(r["정산ID"] for r in lrows),
                      "원장공급가액합": sum((r.get("원장_공급가액") or 0) for r in lrows),
                      "판정": "쿠팡 목록에 없음 — 번호 오기입 또는 목록 기간 밖"})

    # 유형D: 미등록 PO ↔ PO없는 유상 정산 — 금액 유일 매칭이면 자동입력 후보
    unmatched_po = [p for p in coupang if p["po"] not in by_po and p["amount"]]
    queue_items = []
    for p in unmatched_po:
        cands = [r for r in no_po if abs((r.get("원장_공급가액") or 0) - p["amount"]) <= tol]
        if cands:
            unique = (len(cands) == 1 and
                      sum(1 for q in unmatched_po if q["amount"] and
                          abs(q["amount"] - (cands[0].get("원장_공급가액") or 0)) <= tol) == 1)
            D.append({"PO번호": p["po"], "쿠팡금액": p["amount"], "발행일": p["date"],
                      "후보정산": ",".join(r["정산ID"] for r in cands[:5]),
                      "판정": "유일 매칭 → 자동입력" if unique else f"후보 {len(cands)}건 — 수동 확인"})
            if unique:
                sid = cands[0]["정산ID"]
                queue_items.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": sid,
                                    "col": "PO번호", "value": p["po"], "vtype": "text",
                                    "evidence": f"쿠팡 PO목록 금액 유일매칭({p['amount']:,})", "only_if_empty": True})
                if p["date"]:
                    queue_items.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": sid,
                                        "col": "PO발행일", "value": p["date"], "vtype": "date",
                                        "evidence": "쿠팡 PO목록 발행일", "only_if_empty": True})
    if queue_items and "--no-queue" not in args:
        try:
            from ledger_writer import queue_add
            print("자동입력 큐 적재:", queue_add(queue_items), "건 (PO번호·발행일)")
        except Exception as e:
            print("(큐 적재 실패:", e, ")")

    os.makedirs(REPORT_DIR, exist_ok=True)
    base = os.path.join(REPORT_DIR, f"PO대조_{datetime.now():%Y%m%d_%H%M}")
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("# 쿠팡 PO ↔ 관리대장 대조\n\n")
        f.write(f"- 생성 {datetime.now():%Y-%m-%d %H:%M} / 쿠팡 PO {len(coupang)}건 · 원장 PO {len(by_po)}건 · 정상 일치 {len(OK)}건\n")
        f.write(f"- PO필요·번호없는 유상 정산: {len(no_po)}건\n")
        for title, rows_ in [("A. 원장 미등록 PO (누락 ★)", A), ("B. 쿠팡 목록에 없는 원장 PO (오기입 의심)", B),
                             ("C. 금액 불일치", C), ("D. 연결 제안 (미등록 PO ↔ 무PO 정산)", D)]:
            f.write(f"\n## {title} — {len(rows_)}건\n\n")
            if rows_:
                cols = list(rows_[0].keys())
                f.write("| " + " | ".join(cols) + " |\n|" + "---|" * len(cols) + "\n")
                for r in rows_:
                    f.write("| " + " | ".join(str(r[c]) for c in cols) + " |\n")
    allrows = ([{"유형": "A", **r} for r in A] + [{"유형": "B", **r} for r in B]
               + [{"유형": "C", **r} for r in C] + [{"유형": "D", **r} for r in D])
    if allrows:
        keys = []
        for r in allrows:
            for k in r:
                if k not in keys:
                    keys.append(k)
        with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=keys); w.writeheader(); w.writerows(allrows)
    print(f"A(원장미등록) {len(A)} / B(쿠팡목록에없음) {len(B)} / C(금액불일치) {len(C)} / D(연결제안) {len(D)} / 정상 {len(OK)}")
    print("리포트:", base + ".md")


if __name__ == "__main__":
    main()
