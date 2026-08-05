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


def sales_exports(limit=60):
    """판매조회 엑셀을 **전부** 찾는다(새 것부터). 파일명이 무작위라 머리글로 판정한다.

    ★ 2026-08-05 — 예전에는 '가장 최근 것 하나'만 읽었다. 그런데 2025 자료를 받으려고
      2025 판매조회를 내려받는 순간 그것이 '가장 최근 파일'이 되어 색인이 통째로
      2025 로 바뀌었고, 2026 명세서 793건이 **전부 짝 없음**이 됐다(실측).
      한 해를 받으면 다른 해가 사라지는 구조 자체가 틀렸다 — 모두 읽어 합친다.
    """
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import source_dirs as S
    cands = [p for p in glob.glob(os.path.join(S.ERP_DIR, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(p).startswith(("~$", "ESD007E"))]
    cands.sort(key=os.path.getmtime, reverse=True)
    found = []
    for p in cands[:limit]:
        try:
            wb = openpyxl.load_workbook(p, read_only=False, data_only=True)
            ws = wb.active
            head = [str(c or "") for c in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))]
            if "프로젝트코드코드" in "|".join(head) and "진행상태" in "|".join(head):
                found.append((p, ws, head, wb))
            else:
                wb.close()
        except Exception:
            continue
    return found


def build_one(ws, head):
    """엑셀 한 개 → {UJ: 집계}. 한 파일 안에서만 합산한다(파일 간 합산은 중복이다)."""
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
    return out


def build():
    """여러 판매조회를 합친다. **같은 UJ 는 새 파일 것이 이긴다** — 더하지 않는다.
    (더하면 두 내보내기가 겹치는 기간에서 금액이 두 배가 된다)"""
    merged, srcs = {}, []
    found = sales_exports()
    for path, ws, head, wb in found:                 # 새 것 → 옛 것 순서
        try:
            one = build_one(ws, head)
        finally:
            wb.close()
        if not one:
            continue
        srcs.append(os.path.basename(path))
        for uj, v in one.items():
            merged.setdefault(uj, v)                 # 이미 있으면(=더 새 파일) 그대로 둔다
    return merged, srcs


def main():
    idx, srcs = build()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump({"src": srcs, "count": len(idx), "index": idx},
              open(OUT, "w", encoding="utf-8"), ensure_ascii=False)
    issued = sum(1 for v in idx.values() if str(v["state"]).startswith(ISSUED))
    years = {}
    for v in idx.values():
        years[str(v.get("date", ""))[:4]] = years.get(str(v.get("date", ""))[:4], 0) + 1
    print(f"ERP 판매 색인 {len(idx)}개 프로젝트 (원본 {len(srcs)}개) · 발행 이상 단계 {issued} · "
          + " ".join(f"{y}:{n}" for y, n in sorted(years.items()) if y)
          + " → reports/ERP판매_프로젝트색인.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
