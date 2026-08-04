# -*- coding: utf-8 -*-
"""
erp_docs_check.py — ERP 매출서류 현황 ↔ 관리대장 대조 (월별 정합성)
================================================================================
이카운트에서 내려받는 3종을 함께 읽어 관리대장과 맞춘다.
  · 매출(세금)계산서현황   : 일자-No. · 거래처 · 프로젝트명 · 공급가액
  · 거래명세서 현황        : 일자-번호 · 거래처 · 공급가액
  · 회계거래조회/현황      : 전표번호 · 금액 · 적요

■ 왜 건별 1:1 대조가 아니라 월별인가 (중요)
  ERP는 여러 작업을 **한 장의 계산서로 묶어** 발행한다.
  예) `2026/07/25-11  8,600,100  "돌발AS_군포1Sub-Hub …"` 한 건이 실제로는
      여러 캠프의 돌발AS 십수 건을 합친 금액이다. 대표 캠프명만 제목에 남는다.
  묶인 내역은 이카운트 화면의 '내역보기'에만 있고 이 엑셀에는 없다.
  → 그래서 건별 대조는 불가능하고, **월별·유형별 총액**으로 맞춘다.
  → 건별까지 맞추려면 품목 단위 판매현황(내역) 엑셀이 추가로 필요하다.

산출
  reports/ERP서류대조_YYYYMMDD_HHMM.md   월별 비교표·미등록 추정 목록
  reports/ERP서류대조_YYYYMMDD_HHMM.csv  계산서 전량(유형 분류 포함)

실행
  python erp_docs_check.py                 # inbox에서 자동 탐지
  python erp_docs_check.py --tax 파일.xlsx  # 파일 직접 지정
"""
import sys, os, re, csv, collections
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")
CUST = "쿠팡"
KINDS = ["돌발AS", "정기점검", "신규납품", "계단", "철거", "기타"]
SLIP_RE = re.compile(r"^\d{4}/\d{2}/\d{2}\s*-\s*\d+$")


def norm_slip(v):
    return re.sub(r"\s+", "", str(v or ""))


def work_kind(title):
    """계산서 제목(프로젝트명·적요)에서 업무 유형을 뽑는다 — 순수 함수(합성검증 대상)"""
    s = str(title or "")
    if "돌발" in s or re.search(r"\bAS\b", s, re.I):
        return "돌발AS"
    if "정기점검" in s or "분기" in s:
        return "정기점검"
    if "철거" in s:
        return "철거"
    if "계단" in s:
        return "계단"
    if "신규" in s or "납품" in s or re.search(r"\d+\s*EA\b", s, re.I):
        return "신규납품"
    return "기타"


def read_doc(path, cust_col, amt_col, title_col=None):
    """이카운트 현황 엑셀 → [{slip, cust, amt, title}] (쿠팡만)"""
    import openpyxl
    # read_only 금지 — 이카운트 파일은 <dimension>이 잘못돼 1행만 읽히는 사고가 있었다
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                r = [("" if c is None else str(c).strip()) for c in r]
                if not r or len(r) <= max(cust_col, amt_col):
                    continue
                if not SLIP_RE.match(r[0] or "") or CUST not in (r[cust_col] or ""):
                    continue
                a = r[amt_col].replace(",", "")
                out.append({"slip": norm_slip(r[0]), "cust": r[cust_col],
                            "amt": int(float(a)) if re.fullmatch(r"-?\d+(\.\d+)?", a) else 0,
                            "title": r[title_col] if title_col is not None and len(r) > title_col else ""})
    finally:
        wb.close()
    return out


def month_sum(docs):
    """월 → 유형 → (금액, 건수)"""
    amt = collections.defaultdict(lambda: collections.Counter())
    cnt = collections.defaultdict(lambda: collections.Counter())
    for d in docs:
        mo, k = d["slip"][:7], work_kind(d["title"])
        amt[mo][k] += d["amt"]
        cnt[mo][k] += 1
    return amt, cnt


def ledger_months():
    """관리대장 06시트 월별 유상 금액 (작업완료일 기준).

    ★ 2026-08-04 오진 수정: 예전에는 `원장_공급가액`(06시트 '실제작업공급가액')만 봤다.
      그 칸은 **2026-07부터 채우기 시작**한 것이라 1~6월이 전부 0으로 잡혔고,
      리포트가 "06시트에 정산 행이 없음 — 4억 2천만원 미등록"이라고 **잘못 경고**했다.
      실제로는 같은 달 `거래명세서합계`가 111/111·85/85처럼 거의 다 채워져 있다.
      → 실제작업공급가액이 비면 **거래명세서합계로 대신 본다**(둘 다 06시트 유상 금액).
      실제작업공급가액이 비어 있다는 사실 자체는 '금액 재계산 대기' 상태가 따로 잡는다.
    """
    from ecount_reconcile import read_ledger, load_config
    recs = read_ledger(load_config()["reconcile"]["master_xlsx"])
    out = collections.defaultdict(lambda: collections.Counter())
    for sid, r in recs.items():
        mo = str(r.get("작업완료일") or "")[:7].replace("-", "/")
        amt = r.get("원장_공급가액") or r.get("원장_거래명세서합계")
        if mo and amt:
            out[mo][r.get("업무구분") or "기타"] += int(amt or 0)
    return out, recs


def main():
    args = sys.argv[1:]
    from inbox_scan import pick
    tax = args[args.index("--tax") + 1] if "--tax" in args else (pick("tax") or [None])[0]
    stmt = args[args.index("--stmt") + 1] if "--stmt" in args else (pick("stmt") or [None])[0]
    if not tax:
        sys.exit("inbox/ 에 [매출(세금)계산서현황] 엑셀이 없습니다. 이카운트에서 내려받아 넣어주세요"
                 "(파일명은 아무거나 괜찮습니다).")

    taxd = read_doc(tax, 1, 4, 3)                     # 일자-No. | 거래처 | 이메일 | 프로젝트명 | 공급가액
    stmtd = read_doc(stmt, 1, 3) if stmt else []      # 일자-번호 | 거래처 | 이메일 | 공급가액
    amt, cnt = month_sum(taxd)
    lm, recs = ledger_months()

    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    csv_p = os.path.join(REPORT_DIR, f"ERP서류대조_{stamp}.csv")
    with open(csv_p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["일자-No.", "월", "유형", "공급가액", "프로젝트명"])
        for d in sorted(taxd, key=lambda x: x["slip"]):
            w.writerow([d["slip"], d["slip"][:7], work_kind(d["title"]), d["amt"], d["title"]])

    L = []
    L.append(f"# ERP 매출서류 ↔ 관리대장 대조 ({datetime.now():%Y-%m-%d %H:%M})\n")
    L.append(f"- 세금계산서 **{len(taxd)}건** / 거래명세서 **{len(stmtd)}건** (쿠팡)")
    L.append(f"- 계산서 공급가액 합계 **{sum(d['amt'] for d in taxd):,}원**\n")
    L.append("## ERP 세금계산서 — 월별·유형별 공급가액\n")
    L.append("| 월 | " + " | ".join(KINDS) + " | 합계 | 건수 |")
    L.append("|---|" + "---|" * (len(KINDS) + 2))
    for mo in sorted(amt):
        cells = " | ".join(f"{amt[mo][k]:,}" if amt[mo][k] else "-" for k in KINDS)
        L.append(f"| {mo} | {cells} | **{sum(amt[mo].values()):,}** | {sum(cnt[mo].values())} |")

    L.append("\n## AS·정기점검만 — ERP vs 관리대장\n")
    L.append("| 월 | ERP(계산서 발행) | 관리대장 06시트 | 차이 | 판정 |")
    L.append("|---|---|---|---|---|")
    gaps = []
    for mo in sorted(set(list(amt) + list(lm))):
        e = amt[mo]["돌발AS"] + amt[mo]["정기점검"]
        l = sum(lm[mo].values())
        diff = e - l
        if l == 0 and e > 0:
            verdict = "★ 대장 미등록"
            gaps.append((mo, e))
        elif diff == 0:
            verdict = "일치"
        elif diff > 0:
            verdict = "대장 누락 추정"
        else:
            verdict = "계산서 미발행분 존재"
        L.append(f"| {mo} | {e:,} | {l:,} | {diff:,} | {verdict} |")

    if gaps:
        L.append(f"\n### ★ 관리대장 06시트에 정산이 없는 달: {len(gaps)}개월 "
                 f"(AS·점검 합계 **{sum(a for _, a in gaps):,}원**)\n")
        for mo, a in gaps:
            L.append(f"- **{mo}** {a:,}원 — ERP 계산서는 발행 완료, 대장에는 정산 행이 없음")

    L.append("\n## 확인 필요\n")
    L.append("- ERP는 여러 작업을 **한 장으로 묶어** 발행한다(예: 계산서 1건 = 돌발AS 15건). "
             "묶인 내역은 이카운트 '내역보기'에만 있어 이 엑셀에는 없다.")
    L.append("- 따라서 **건별 금액 배분은 이 자료만으로 불가**하다. "
             "건별까지 채우려면 **품목 단위 판매현황(내역) 엑셀**이 추가로 필요하다.")
    lnums = {norm_slip(r.get("원장_거래명세서번호")) for r in recs.values()
             if str(r.get("원장_거래명세서번호") or "").strip()}
    enums = {d["slip"] for d in stmtd} or {d["slip"] for d in taxd}
    if lnums:
        L.append(f"- 관리대장 거래명세서번호 {len(lnums)}개 중 ERP 번호와 일치 **{len(lnums & enums)}개** "
                 f"— 대장의 번호 체계가 ERP 발행번호와 다르다(확인 필요).")

    md_p = os.path.join(REPORT_DIR, f"ERP서류대조_{stamp}.md")
    open(md_p, "w", encoding="utf-8").write("\n".join(L) + "\n")

    # 관리대장에 원본을 남긴다 — 앱·보고가 2026년 전체 매출을 실제 자료로 보게
    if "--sheet" in args:
        from findings_sheet import upsert, build_generic_sheet
        from ecount_reconcile import load_config as _lc, resolve_master
        HDR = ["일자-No.", "월", "유형", "공급가액", "부가세", "합계", "거래처", "프로젝트명"]
        W = [14, 9, 10, 14, 12, 14, 26, 60]
        smap = {d["slip"]: d for d in stmtd}
        rows = []
        for d in sorted(taxd, key=lambda x: x["slip"]):
            vat = round(d["amt"] * 0.1)
            rows.append([d["slip"], d["slip"][:7], work_kind(d["title"]), d["amt"], vat,
                         d["amt"] + vat, d["cust"], d["title"]])
        xml = build_generic_sheet(
            "25_ERP매출서류", HDR, W, rows,
            "이카운트 [매출(세금)계산서현황] 원본. ERP는 여러 작업을 한 장으로 묶어 발행하므로 "
            "이 시트의 1행 = 작업 1건이 아니다(건별 배분은 품목 단위 판매현황 필요). "
            "erp_docs_check.py 가 자동 갱신 — 수기 입력 금지.")
        master = resolve_master(_lc()["reconcile"]["master_xlsx"])
        dst, msg = upsert(master, xml, sheet_name="25_ERP매출서류", headers=HDR)
        print(f"25_ERP매출서류 시트: {msg}")

    print(f"ERP 계산서 {len(taxd)}건 · 명세서 {len(stmtd)}건 (쿠팡) | "
          f"대장 미등록 {len(gaps)}개월 {sum(a for _, a in gaps):,}원")
    for mo in sorted(amt):
        e = amt[mo]["돌발AS"] + amt[mo]["정기점검"]
        print(f"  {mo}  AS·점검 {e:>12,} | 대장 {sum(lm[mo].values()):>12,}")
    print(f"리포트: {os.path.basename(md_p)} / {os.path.basename(csv_p)}")


if __name__ == "__main__":
    main()
