# -*- coding: utf-8 -*-
"""
erp_bundle.py — ERP 계산서 1장에 **어떤 프로젝트가 묶여 있는지** 추정해 표기
================================================================================
ERP는 여러 작업을 한 장으로 묶어 발행한다(최대 15건). 그래서 25_ERP매출서류의
1행은 작업 1건이 아니다. 품목 단위 판매현황이 들어오기 전까지는 건별 배분을 못 하지만,
**캠프·유형·기간**으로 후보를 좁히고 **금액 합이 맞는지**로 확신도를 매길 수는 있다.

판정
  확정   후보 작업들의 공급가액 합 = 계산서 공급가액 (±1원)   → 구성이 확실하다
  유력   합이 ±3% 이내                                        → 거의 맞다
  추정   캠프·유형·기간은 맞지만 금액이 안 맞음                → 사람이 확인
  미상   후보를 못 찾음                                        → 사람이 확인

**추정을 확정처럼 쓰지 않는다** — 판정 열을 함께 적어 어디까지 믿을지 보이게 한다.

실행
  python erp_bundle.py             # 리포트만
  python erp_bundle.py --sheet     # + 26_계산서구성 시트 반영(vN+1)
"""
import sys, os, re, csv
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")

_CAMP_RE = re.compile(r"[가-힣A-Za-z]+\d*(?:BMB|MB|캠프|Sub-?FC|Sub-?hub|FC)(?:\([^)]*\))?", re.I)
_PAREN_RE = re.compile(r"\([^)]*\)")


def camp_of(title):
    """계산서 제목·캠프명에서 캠프를 뽑는다"""
    m = _CAMP_RE.search(str(title or ""))
    if m:
        return m.group()
    t = re.sub(r"^(쿠팡\S*|돌발AS|정기점검)[_\s-]*", "", str(title or "")).strip()
    return (t or str(title or ""))[:28]


def camp_key(name):
    """'송파1MB(감일동)' 과 '송파1MB' 를 같은 캠프로 본다 — 괄호 안 지명은 표기 흔들림이 크다"""
    s = _PAREN_RE.sub("", str(name or ""))
    return re.sub(r"[\s_·\-]", "", s).lower()


def kind_of(title):
    """계산서 제목 → 업무 유형 (erp_docs_check.work_kind 와 같은 규칙)"""
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


def load_erp():
    """25_ERP매출서류 → [{slip, month, kind, amt, title, camp}]"""
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    if "25_ERP매출서류" not in wb.sheetnames:
        wb.close()
        return master, []
    ws = wb["25_ERP매출서류"]
    hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
    idx = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        g = lambda c: (r[idx[c]] if c in idx and idx[c] < len(r) else None)
        slip = str(g("일자-No.") or "").strip()
        if not slip:
            continue
        title = str(g("프로젝트명") or "")
        out.append({"slip": slip, "month": str(g("월") or "")[:7],
                    "kind": str(g("유형") or "") or kind_of(title),
                    "amt": int(g("공급가액") or 0), "title": title,
                    "camp": camp_of(title)})
    wb.close()
    return master, out


def load_works(master):
    """02·04·06 → 후보 작업 [{prj, camp, kind, date, amt, src}]"""
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    out = []
    spec = [("02_돌발AS접수", "돌발AS", "접수일자", "작업완료일", None),
            ("04_정기점검", "정기점검", "점검예정일", "실제점검일", None),
            ("05_신규납품설치", None, "요청일", None, "견적금액"),
            ("06_거래서류청구수금", None, "작업완료일", None, "실제작업공급가액")]
    for sheet, kind, dcol, dcol2, amtcol in spec:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        idx = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
        for r in ws.iter_rows(min_row=5, values_only=True):
            g = lambda c: (r[idx[c]] if c in idx and idx[c] < len(r) else None)
            prj = str(g("프로젝트NO") or "").strip()
            camp = str(g("캠프명") or "").strip()
            if not prj or not camp:
                continue
            d = g(dcol2) or g(dcol)
            out.append({"prj": prj, "camp": camp,
                        "kind": kind or str(g("업무구분") or ""),
                        "date": str(d or "")[:10],
                        "amt": int(g(amtcol) or 0) if amtcol else 0,
                        "src": sheet})
    wb.close()
    return out


_QTR_RE = re.compile(r"(\d{2})년\s*(\d)\s*분기")


def window(doc):
    """후보로 볼 기간 (시작월, 끝월).

    ERP는 작업이 끝난 뒤 묶어서 끊으므로 발행월보다 앞선다 → 기본은 발행월 이전 3개월.
    다만 제목에 '25년 4분기'처럼 분기가 적혀 있으면 그 분기를 그대로 쓴다
    (분기 점검은 발행이 한참 늦어져 3개월 창을 벗어난다)."""
    mo = doc["month"].replace("/", "-")
    if len(mo) < 7:
        return "", ""
    q = _QTR_RE.search(doc["title"])
    if q:
        y, n = 2000 + int(q.group(1)), int(q.group(2))
        return "%04d-%02d" % (y, n * 3 - 2), "%04d-%02d" % (y, n * 3)
    y, m = int(mo[:4]), int(mo[5:7])
    lo = "%04d-%02d" % (y - 1, m + 9) if m <= 3 else "%04d-%02d" % (y, m - 3)
    return lo, mo


def bundle(doc, works):
    """계산서 1장 → (후보 프로젝트 목록, 판정, 합계)

    후보를 못 찾으면 판정 자리에 **왜 못 찾았는지**를 적는다.
    '미상' 한 마디로는 사람이 무엇을 해야 할지 알 수 없다.
    """
    ck, kind = camp_key(doc["camp"]), doc["kind"]
    lo, hi = window(doc)
    if not ck or not lo:
        return [], "미상(캠프 못 읽음)", 0
    same_camp = [w for w in works if camp_key(w["camp"]) == ck]
    if not same_camp:
        return [], "미상(그 캠프 작업이 대장에 없음)", 0
    same_kind = [w for w in same_camp
                 if not kind or kind == "기타" or kind in w["kind"] or w["kind"] in kind]
    if not same_kind:
        return [], f"미상({kind} 자료 없음)", 0
    hits = [w for w in same_kind if w["date"] and lo <= w["date"][:7] <= hi]
    if not hits:
        got = sorted({w["date"][:7] for w in same_kind if w["date"]})
        return [], f"미상(기간 {lo}~{hi} 밖 · 보유 {','.join(got[:3])})", 0
    seen, prjs = set(), []
    for w in hits:
        if w["prj"] not in seen:
            seen.add(w["prj"])
            prjs.append(w["prj"])
    tot = sum(w["amt"] for w in hits)
    if doc["amt"] and tot:
        if abs(tot - doc["amt"]) <= 1:
            v = "확정"
        elif abs(tot - doc["amt"]) <= doc["amt"] * 0.03:
            v = "유력"
        else:
            v = "추정"
    else:
        v = "추정"
    return prjs, v, tot


def main():
    args = sys.argv[1:]
    master, docs = load_erp()
    if not docs:
        sys.exit("25_ERP매출서류 시트가 비어 있습니다 — 먼저 erp_docs_check.py --sheet 를 돌리세요")
    works = load_works(master)
    rows, stat, why = [], {}, {}
    for d in sorted(docs, key=lambda x: x["slip"]):
        prjs, verdict, tot = bundle(d, works)
        stat[verdict.split("(")[0]] = stat.get(verdict.split("(")[0], 0) + 1
        why[verdict] = why.get(verdict, 0) + 1
        rows.append([d["slip"], d["month"], d["kind"], d["camp"], d["amt"],
                     len(prjs), ", ".join(prjs[:15]), tot, verdict, d["title"]])

    os.makedirs(REPORT_DIR, exist_ok=True)
    base = os.path.join(REPORT_DIR, f"계산서구성_{datetime.now():%Y%m%d_%H%M}")
    HDR = ["일자-No.", "월", "유형", "캠프명", "계산서공급가액", "포함건수",
           "포함프로젝트NO", "후보합계", "판정", "프로젝트명"]
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(HDR)
        w.writerows(rows)
    print(f"ERP 계산서 {len(rows)}장 — " + " · ".join(f"{k} {v}" for k, v in sorted(stat.items())))
    print("  미상 사유: " + " · ".join(f"{k.split('(',1)[-1].rstrip(')')} {v}"
                                       for k, v in sorted(why.items(), key=lambda x: -x[1])
                                       if k.startswith("미상"))[:300])
    print("리포트:", base + ".csv")

    if "--sheet" in args:
        from findings_sheet import upsert, build_generic_sheet
        W = [14, 9, 10, 22, 15, 9, 62, 14, 8, 46]
        xml = build_generic_sheet(
            "26_계산서구성", HDR, W, rows,
            "ERP 계산서 1장에 묶인 작업을 캠프·유형·기간으로 좁혀 추정한 것. "
            "판정이 '확정'(금액 합 일치)일 때만 그대로 믿고, '추정'·'미상'은 사람이 확인한다. "
            "품목 단위 판매현황이 들어오면 건별로 정확히 배분된다. erp_bundle.py 자동 갱신 — 수기 입력 금지.")
        dst, msg = upsert(master, xml, sheet_name="26_계산서구성", headers=HDR)
        print(f"26_계산서구성 시트: {msg}")


if __name__ == "__main__":
    main()
