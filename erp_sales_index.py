# -*- coding: utf-8 -*-
"""
erp_sales_index.py — ERP 판매조회를 **프로젝트NO 색인**으로 굳힌다 (읽기 전용)

왜(2026-08-05):
 ① 앱 정산 금액이 틀렸다. 실제작업공급가액(03시트 수식)이 비어 거래명세서합계로 대신
    보여 줬는데, **명세서합계는 부가세 포함액**이다(UJ2600050: 476,300 = 433,000×1.1).
    "부가세 별도"라고 적힌 자리에 포함액이 들어가 있었다. → ERP 공급가액을 정본으로 쓴다.
 ② "세금계산서 미발행을 찾아 완료 처리" — 발행 여부의 객관 근거는 ERP 진행상태다
    (6.세금계산서발행 / 7.수금완료). 그 상태를 프로젝트NO 로 바로 찾게 만든다.

산출: reports/ERP판매_프로젝트색인.json
  { "UJ2600050": {"supply":433000,"vat":43300,"total":476300,"state":"4.세금계산서발행대기",
                  "date":"2026/01/07","po":"","cust":"강서1MB(가양A)","rows":1}, ... }
같은 UJ 가 여러 행이면 금액은 합산하고 상태는 **가장 앞선 단계**를 남긴다
(한 건이라도 미발행이면 그 프로젝트는 아직 안 끝난 것으로 본다 — 낙관 금지).
"""
import glob
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT = os.path.join(ROOT, "reports", "ERP판매_프로젝트색인.json")
UJ = re.compile(r"UJ\d{7}")
# 진행 단계 순서 — 작은 값이 덜 진행된 것. 합칠 때 **가장 덜 진행된 상태**를 남긴다.
ORDER = {"1.미확인": 1, "확인": 2, "2.메일발송": 3, "3.오더처리": 4,
         "4.세금계산서발행대기": 5, "5.": 6, "6.세금계산서발행": 7, "7.수금완료": 8,
         "8.무상납품완료": 9}
ISSUED = ("6.세금계산서발행", "7.수금완료")


def rank(state):
    for k, v in ORDER.items():
        if str(state or "").startswith(k):
            return v
    return 0


def newest_sales():
    """가장 최근 판매조회 엑셀 경로 — 머리글로 직접 판정한다(파일명이 무작위)."""
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import source_dirs as S
    cands = [p for p in glob.glob(os.path.join(S.ERP_DIR, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(p).startswith(("~$", "ESD007E"))]
    cands.sort(key=os.path.getmtime, reverse=True)
    for p in cands[:60]:
        try:
            wb = openpyxl.load_workbook(p, read_only=False, data_only=True)
            ws = wb.active
            head = [str(c or "") for c in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))]
            if "프로젝트코드코드" in "|".join(head) and "진행상태" in "|".join(head):
                return p, ws, head, wb
            wb.close()
        except Exception:
            continue
    return None, None, None, None


def build():
    path, ws, head, wb = newest_sales()
    if not ws:
        return {}, None
    idx = {h: i for i, h in enumerate(head)}

    def num(r, key):
        s = (r[idx[key]] if key in idx and len(r) > idx[key] else "").replace(",", "")
        return int(float(s)) if re.fullmatch(r"-?\d+(\.\d+)?", s) else 0

    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        r = ["" if c is None else str(c).strip() for c in r]
        m = UJ.search(" ".join(r))
        if not m:
            continue
        uj = m.group(0)
        cur = out.setdefault(uj, {"supply": 0, "vat": 0, "total": 0, "state": "",
                                  "date": "", "po": "", "cust": "", "rows": 0})
        cur["supply"] += num(r, "공급가액합계")
        cur["vat"] += num(r, "부가세합계")
        cur["total"] += num(r, "금액합계")
        cur["rows"] += 1
        st = r[idx["진행상태"]] if "진행상태" in idx else ""
        if not cur["state"] or rank(st) < rank(cur["state"]):
            cur["state"] = st
        if not cur["date"]:
            cur["date"] = (r[0] or "")[:10]
        for k, col in (("po", "PO번호"), ("cust", "거래처명")):
            if not cur[k] and col in idx and len(r) > idx[col]:
                cur[k] = r[idx[col]]
    wb.close()
    return out, os.path.basename(path)


def main():
    idx, src = build()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump({"src": src, "count": len(idx), "index": idx},
              open(OUT, "w", encoding="utf-8"), ensure_ascii=False)
    issued = sum(1 for v in idx.values() if str(v["state"]).startswith(ISSUED))
    print(f"ERP 판매 색인 {len(idx)}개 프로젝트 (원본 {src}) · 발행 이상 단계 {issued} → "
          f"reports/ERP판매_프로젝트색인.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
