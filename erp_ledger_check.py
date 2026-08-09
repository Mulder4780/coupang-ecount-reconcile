# -*- coding: utf-8 -*-
"""
erp_ledger_check.py — ERP 거래처별계정별원장 ↔ 관리대장 정밀 대조
===================================================================
이카운트 [거래처별계정별원장](쿠팡로지스틱스, 4049 제품매출) 엑셀 내보내기를 원본 자료나 inbox/에 넣으면
관리대장과 전표 단위로 대조해 다음 4가지 문제 유형을 자동 검출한다.

  [유형A] ERP에만 있는 전표      : 회계반영됐지만 관리대장에 근거 작업 없음
                                   → 실제 설치·작업 여부 현장 확인 필요 (예: 2026/03/25-1 인천8MB 26,690,000)
  [유형B] 원장에만 있는 유상건    : 작업완료·유상인데 ERP 미반영 (매출 누락)
  [유형C] 회계반영O·세금계산서X  : ERP 전표는 있는데 관리대장 15시트 세금계산서 미발행
  [유형D] 금액불일치             : 같은 전표번호인데 ERP 차변금액 ≠ 원장 합계금액

전표 매칭 키: 관리대장 06 거래명세서번호("2026/07/01-4") = ERP 일자-No.("2026/07/01 -4") — 공백 정규화.
월합계 비교(ERP 월계 vs 원장 유상합계)도 함께 출력.

실행:
    python erp_ledger_check.py                      # 원본 자료+inbox에서 내용을 보고 자동 탐지
    python erp_ledger_check.py --file 경로 [--master 경로]   # 파일 직접 지정(테스트용)
"""
import sys, os, re, csv, json, glob
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
STATUS_JSON = os.path.join(REPORT_DIR, "ERP원장대조_상태.json")


def norm_slip(s):
    """'2026/03/25 -1' / '2026/03/25-1' / '2026-03-25-1' → '2026/03/25-1'"""
    s = str(s or "").strip().replace(" ", "")
    s = re.sub(r"^(\d{4})[/-](\d{2})[/-](\d{2})", r"\1/\2/\3", s)
    return s


def ledger_record_date(row, slip=""):
    """원장 행을 ERP 조회기간과 비교할 대표 날짜(YYYY-MM-DD)."""
    normalized = norm_slip(slip)
    if re.match(r"^\d{4}/\d{2}/\d{2}-\d+$", normalized):
        return normalized[:10].replace("/", "-")
    for key in ("원장_거래명세서발행일", "작업완료일"):
        day = _d(row.get(key))
        if day:
            return day
    return ""


def in_erp_period(row, slip, start, end):
    """날짜를 아는 행은 ERP 내보내기 조회기간 안에서만 'ERP 미확인'으로 판정한다."""
    day = ledger_record_date(row, slip)
    return not (start and end and day) or start <= day <= end


def key_looks_wrong(matched, unique_slips, floor=10):
    """짝이 이만큼 안 지어지면 'A·B 가 많다'가 아니라 **열쇠가 안 맞는다**로 읽는다.

    자료가 정말 없어서 0건인 것과 열쇠가 안 맞아 0건인 것은 겉이 똑같다. 가르는 기준은
    비율이다 — 열쇠가 맞으면 몇 건은 반드시 걸린다. 전표가 `floor` 건도 안 되면
    비율을 말할 수 없으므로 아무 말도 하지 않는다(막 시작한 회차를 겁주지 않는다).
    """
    if unique_slips < floor:
        return False
    return matched * 10 < unique_slips


def project_sale_match(row, sales):
    """원장 행과 ERP 판매조회 전표를 프로젝트·PO·금액으로 직접 맞춘다.

    거래명세서번호와 회계 ``일자-No.``가 서로 다른 순번인 환경에서는 전표번호 불일치가
    곧 ERP 미등록을 뜻하지 않는다. 프로젝트가 같고, 두 쪽에 PO가 있으면 PO도 같고,
    공급가액/부가세포함 금액 중 하나가 원 단위로 같을 때만 직접 근거로 인정한다.
    """
    po_text = str(row.get("원장_PO번호") or "").upper()
    ledger_amounts = {round(_num(row.get(key)) or 0)
                      for key in ("원장_공급가액", "원장_합계", "원장_거래명세서합계")}
    ledger_amounts.discard(0)
    matched_po = []
    for sale in sales or []:
        sale_po = str(sale.get("po") or "").upper()
        if sale_po and po_text and sale_po not in po_text:
            continue
        matched_po.append(sale)
    erp_amounts = {round(_num(sale.get(key)) or 0)
                   for sale in matched_po for key in ("supply", "total")}
    erp_amounts.discard(0)
    return {
        "present": bool(matched_po),
        "amount_match": bool(ledger_amounts and erp_amounts and ledger_amounts & erp_amounts),
        "ledger_amounts": sorted(ledger_amounts),
        "erp_amounts": sorted(erp_amounts),
        "sales": matched_po,
    }


def parse_erp_export(path):
    """거래처별계정별원장 엑셀 파싱 → [{slip, date, remark, amount}]  (차변=매출)"""
    import openpyxl
    # 이카운트 내보내기는 실제 데이터가 있어도 dimension을 A1:A1로 잘못 쓰는 경우가 있다.
    # read_only=True는 그 메타데이터를 믿고 1셀만 읽으므로 일반 모드로 열되 저장하지 않는다.
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    slips, totals = [], {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i, idx = None, {}
        for i, r in enumerate(rows[:20]):
            names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
            if any("적요" in n for n in names) and any("대변" in n for n in names):
                hdr_i = i
                for n, j in names.items():
                    if "일자" in n or "No" in n: idx.setdefault("slip", j)
                    if "적요" in n: idx["remark"] = j
                    if "차변" in n: idx["debit"] = j
                break
        if hdr_i is None:
            continue
        for r in rows[hdr_i + 1:]:
            if r is None or all(c is None for c in r):
                continue
            slip_raw = r[idx["slip"]] if idx.get("slip") is not None else None
            remark = str(r[idx["remark"]] or "") if idx.get("remark") is not None else ""
            # 외상매출금 계정별원장은 매출 발생이 차변, 입금이 대변이다.
            # 대변을 매출로 읽으면 실제 2026-01~06 원본의 전표가 0건이 되고
            # 입금 행을 매출로 오인한다(2026-07-30 실데이터에서 발견).
            amt = _num(r[idx["debit"]]) if idx.get("debit") is not None else None
            joined = " ".join(str(c) for c in r if c is not None)
            mtot = re.search(r"(\d{4}/\d{2})\s*계|월\s*계", joined)
            if mtot and amt:
                key = mtot.group(1) or "월계"
                totals[key] = amt
                continue
            slip = norm_slip(slip_raw)
            if not re.match(r"^\d{4}/\d{2}/\d{2}-\d+$", slip) or amt is None:
                continue
            slips.append({"slip": slip, "date": slip[:10], "remark": remark, "amount": amt})
    wb.close()
    return slips, totals


def main():
    args = sys.argv[1:]
    cfg = load_config()
    master = cfg["reconcile"]["master_xlsx"]
    if "--master" in args:
        master = args[args.index("--master") + 1]
    if "--file" in args:
        files = [args[args.index("--file") + 1]]
    else:
        # 파일명이 아니라 **내용**으로 고른다 — 이카운트 다운로드는 이름이 무작위다
        from inbox_scan import pick
        files = pick("ledger")
    if not files:
        sys.exit("원본 자료와 inbox/ 에 거래처별계정별원장이 없습니다. "
                 "이카운트 [거래처별계정별원장]을 엑셀로 내려받아 넣어주세요"
                 "(파일명은 아무거나 괜찮습니다).")

    slips, totals = [], {}
    for f in files:
        s, t = parse_erp_export(f)
        slips += s; totals.update(t)
        print(f"ERP 파일 '{os.path.basename(f)}': 전표 {len(s)}건")

    recs = read_ledger(master)
    # 원장: 전표번호별 합계(한 전표에 여러 정산이 묶일 수 있음)
    by_slip = defaultdict(list)
    for sid, r in recs.items():
        sl = norm_slip(r.get("원장_거래명세서번호"))
        if re.match(r"^\d{4}/\d{2}/\d{2}-\d+$", sl):
            by_slip[sl].append(r)

    erp_by_slip = {s["slip"]: s for s in slips}
    A, B, C, D, OK = [], [], [], [], []
    erp_days = sorted(s["date"].replace("/", "-") for s in slips if s.get("date"))
    period_start = erp_days[0] if erp_days else ""
    period_end = erp_days[-1] if erp_days else ""

    # 유형A/D/C: ERP 전표 순회
    for sl, s in sorted(erp_by_slip.items()):
        lrows = by_slip.get(sl)
        if not lrows:
            A.append({"전표": sl, "적요": s["remark"][:60], "ERP금액": s["amount"],
                      "판정": "ERP에만 존재 — 근거작업·실제설치 확인 필요"})
            continue
        led_sum = sum((r.get("원장_합계") or 0) for r in lrows)
        ids = ",".join(r["정산ID"] for r in lrows)
        if abs(led_sum - s["amount"]) > 1:
            D.append({"전표": sl, "정산ID": ids, "ERP금액": s["amount"], "원장합계": led_sum,
                      "차액": round(s["amount"] - led_sum)})
        else:
            OK.append(sl)
        for r in lrows:
            issued = _d(r.get("원장_세금계산서실제발행일")) or _d(r.get("원장_세금계산서발행일"))
            if not issued:
                C.append({"전표": sl, "정산ID": r["정산ID"], "캠프명": r.get("캠프명"),
                          "ERP금액": s["amount"], "판정": "회계반영O·세금계산서 미발행"})

    # 전표번호 키의 건강도는 프로젝트 직접매칭을 섞기 전에 잰다. 이 값이 낮다는 사실이
    # A·B를 개별 업무로 지시하면 안 된다는 근거다.
    slip_matched = len(OK) + len(D)
    try:
        from settlement_completion import erp_sales_index
        project_sales = erp_sales_index()
    except Exception:
        project_sales = {}
    project_ok, project_amount_gap = [], []

    # 유형B: 원장 유상인데 ERP에 전표 없음
    for sid, r in sorted(recs.items()):
        if r.get("비용구분") != "유상" or not r.get("원장_공급가액"):
            continue
        sl = norm_slip(r.get("원장_거래명세서번호"))
        if not in_erp_period(r, sl, period_start, period_end):
            continue
        if sl and sl in erp_by_slip:
            continue
        project = str(r.get("프로젝트NO") or "").strip().upper()
        direct = project_sale_match(r, project_sales.get(project, []))
        if direct["present"] and direct["amount_match"]:
            project_ok.append(sid)
            continue
        if direct["present"]:
            project_amount_gap.append(sid)
            D.append({"전표": "(프로젝트 직접매칭)", "정산ID": sid,
                      "프로젝트NO": project, "ERP금액": ",".join(map(str, direct["erp_amounts"])),
                      "원장합계": ",".join(map(str, direct["ledger_amounts"])),
                      "차액": "직접 비교 필요"})
            continue
        B.append({"정산ID": sid, "프로젝트NO": project,
                  "캠프명": r.get("캠프명"), "명세서번호": sl or "(없음)",
                  "원장공급가액": r.get("원장_공급가액"),
                  "판정": "ERP 원장에서 전표 미확인" + ("" if sl else " (명세서번호도 없음 — 미청구)")})

    os.makedirs(REPORT_DIR, exist_ok=True)
    base = os.path.join(REPORT_DIR, f"ERP원장대조_{datetime.now():%Y%m%d_%H%M}")
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("# ERP 거래처별계정별원장 ↔ 관리대장 대조\n\n")
        f.write(f"- 생성 {datetime.now():%Y-%m-%d %H:%M} / ERP 전표 {len(slips)}건, 원장 매칭 정상 {len(OK)}건\n")
        if period_start:
            f.write(f"- ERP 조회기간 {period_start} ~ {period_end}; B(원장에만)는 이 기간의 원장 행만 집계\n")
        if totals:
            f.write(f"- ERP 월계: {json.dumps(totals, ensure_ascii=False)}\n")
        # ★ **매칭이 거의 없으면 그건 A·B 가 아니라 열쇠 이야기다** (2026-08-08 실측).
        #   이 리포트는 전표번호로 짝을 짓는데, 열쇠가 안 맞으면 짝이 하나도 안 지어지고
        #   그 결과가 "A 296건 · B 56건" 이라는 **그럴듯한 경보**로 나온다. 정상 0건이
        #   유일한 신호인데 머리글 한 줄에 작게 적혀 아무도 안 봤다(1,856건이던 시절에도
        #   0건이었다). 실측: ERP 302 전표 대 원장 명세서번호 65개 중 겹침 6 — 06시트
        #   거래명세서번호와 회계 전표번호는 **서로 다른 순번**이다.
        matched = slip_matched
        key_wrong = key_looks_wrong(matched, len(erp_by_slip))
        if key_wrong:
            f.write(
                f"\n> ⚠ **짝이 지어진 전표가 {matched}건뿐입니다"
                f"(ERP 고유 전표 {len(erp_by_slip)}건).** 아래 A·B 는 '없는 것'이 아니라\n"
                "> **열쇠가 안 맞아 못 찾은 것**일 수 있습니다. 이 대조는 06시트\n"
                "> `거래명세서번호` = ERP `일자-No.` 를 전제로 하는데, 실측에서 그 둘은\n"
                "> 서로 다른 순번이었습니다. 현장 확인을 지시하기 전에 열쇠부터 확인하세요.\n")
        for title, rows_ in [("A. ERP에만 있는 전표 (설치·작업 근거 확인 필요 ★)", A),
                             ("B. 원장 유상건인데 ERP 미확인 (매출 누락 위험)", B),
                             ("C. 회계반영O·세금계산서 미발행", C),
                             ("D. 금액 불일치", D)]:
            f.write(f"\n## {title} — {len(rows_)}건\n\n")
            if rows_:
                cols = list(rows_[0].keys())
                f.write("| " + " | ".join(cols) + " |\n|" + "---|" * len(cols) + "\n")
                for r in rows_:
                    f.write("| " + " | ".join(str(r[c]) for c in cols) + " |\n")
    allrows = ([{"유형": "A", **r} for r in A] + [{"유형": "B", **r} for r in B]
               + [{"유형": "C", **r} for r in C] + [{"유형": "D", **r} for r in D])
    csv_path = ""
    if allrows:
        keys = []
        for r in allrows:
            for k in r:
                if k not in keys:
                    keys.append(k)
        csv_path = base + ".csv"
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=keys); w.writeheader(); w.writerows(allrows)
    key_wrong = key_looks_wrong(slip_matched, len(erp_by_slip))
    status = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_csv": os.path.basename(csv_path) if csv_path else "",
        "key_looks_wrong": key_wrong,
        "erp_unique_slips": len(erp_by_slip),
        "matched_by_slip": slip_matched,
        "matched_by_project": len(project_ok),
        "project_amount_gap": len(project_amount_gap),
        "counts": {"A": len(A), "B": len(B), "C": len(C), "D": len(D)},
    }
    os.makedirs(REPORT_DIR, exist_ok=True)
    tmp = STATUS_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATUS_JSON)
    print(f"A(ERP에만) {len(A)} / B(원장에만) {len(B)} / C(계산서미발행) {len(C)} / "
          f"D(금액불일치) {len(D)} / 전표키 정상 {len(OK)} / 프로젝트 직접확인 {len(project_ok)}")
    print("리포트:", base + ".md")
    return {"A": A, "B": B, "C": C, "D": D, "OK": OK,
            "PROJECT_OK": project_ok, "key_looks_wrong": key_wrong}


if __name__ == "__main__":
    main()
