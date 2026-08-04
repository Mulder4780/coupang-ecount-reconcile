# -*- coding: utf-8 -*-
"""
stmt_docs.py — ERP **거래명세서 인쇄본(ESD007E)** 파서·원장 대조

왜 필요한가(2026-08-04 실측):
  거래명세서 785건을 받아 Z: 에 넣었는데 `erp_docs_check` 는 "거래명세서 0건"으로 봤다.
  같은 'stmt' 로 분류되지만 **레이아웃이 전혀 다른 두 종류**가 있기 때문이다.
    · 거래명세서 **현황**(표) — 일자-번호 | 거래처 | 이메일 | 공급가액  ← 기존 도구가 읽는 것
    · 거래명세서 **인쇄본**(문서, ESD007E) — 전표 블록이 세로로 반복  ← 아무도 안 읽고 있었다
  게다가 인쇄본의 거래처는 '쿠팡로지스틱스'가 아니라 **캠프명**(송파5캠프·시흥2MB…)이라,
  쿠팡 필터(CUST)로 거르면 0건이 된다. 그래서 캠프명 기준으로 따로 읽는다.

블록 구조(실측):
    거래명세표
    전표번호 : 20260126-7          ← 원장 06시트 거래명세서번호("2026/01/26-7")와 같은 키
    부서-거래처 : 송파5캠프(…)
    출고창고 : 부산공장
    금액 : 153,351,000
    일자 / 품목명[규격] / 수량 / 단가 / 공급가액 / 부가세 / 적요
    01/14 / … (품목 행 여러 개)

산출: reports/거래명세서_상세.json (건별) + reports/거래명세서_대조.md (요약)
읽기 전용 — 엑셀에 쓰지 않는다. 원장 반영이 필요하면 판단은 사람/후속 도구가 한다.

  python stmt_docs.py            # 전체 스캔·대조
  python stmt_docs.py --json     # JSON 경로만 출력
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

OUT_JSON = os.path.join(ROOT, "reports", "거래명세서_상세.json")
OUT_MD = os.path.join(ROOT, "reports", "거래명세서_대조.md")
def norm_slip(raw):
    """전표번호를 한 모양으로 — `2026/01/26-7`.

    ERP 인쇄본은 `20260126-7`, 원장 06시트는 `2026/07/01-4` 로 적는다.
    게다가 원장에는 `2026/07-01-2` 처럼 구분자가 밀린 손입력도 섞여 있다.
    구분자를 신뢰하지 말고 **숫자만 뽑아** 앞 8자리를 날짜, 마지막 덩어리를 번호로 본다.
    """
    parts = re.findall(r"\d+", str(raw or ""))
    if len(parts) < 2:
        return ""
    d, n = "".join(parts[:-1]), parts[-1]
    if len(d) != 8:
        return ""
    return f"{d[:4]}/{d[4:6]}/{d[6:8]}-{int(n)}"


def _files():
    """ESD007E 인쇄본만 모은다 — 날짜 폴더 전체를 재귀로 훑는다."""
    import source_dirs as S
    out = []
    for base in (S.ERP_DIR,):
        out += glob.glob(os.path.join(base, "**", "ESD007E*.xlsx"), recursive=True)
    return sorted(set(out))


def parse_file(path):
    """한 파일에서 전표 블록들을 뜯는다. 실패한 블록은 건너뛰고 계속 간다."""
    import openpyxl
    # read_only 금지 — 이카운트 파일은 <dimension>이 틀려 1행만 읽히는 사고가 있었다.
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    docs, cur = [], None
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [("" if c is None else str(c).strip()) for c in row]
                joined = " ".join(x for x in cells if x)
                if not joined:
                    continue
                head = cells[0] if cells else ""
                if head.startswith("전표번호"):
                    if cur and cur.get("slip"):
                        docs.append(cur)
                    cur = {"slip": norm_slip(joined), "cust": "", "warehouse": "",
                           "amount": 0, "items": [], "src": os.path.basename(path)}
                    continue
                if cur is None:
                    continue
                if head.startswith("부서-거래처"):
                    cur["cust"] = joined.split(":", 1)[-1].strip()
                elif head.startswith("출고창고"):
                    cur["warehouse"] = joined.split(":", 1)[-1].strip()
                elif head.startswith("금액"):
                    v = joined.split(":", 1)[-1].strip().replace(",", "")
                    cur["amount"] = int(float(v)) if re.fullmatch(r"-?\d+(\.\d+)?", v) else 0
                elif re.fullmatch(r"\d{1,2}/\d{1,2}", head):
                    # 일자 / 품목 / 수량 / 단가 / 공급가액 / 부가세 / 적요
                    def num(i):
                        s = (cells[i] if len(cells) > i else "").replace(",", "")
                        return int(float(s)) if re.fullmatch(r"-?\d+(\.\d+)?", s) else 0
                    cur["items"].append({"date": head, "name": cells[1] if len(cells) > 1 else "",
                                         "qty": num(2), "supply": num(4), "vat": num(5),
                                         "note": cells[6] if len(cells) > 6 else ""})
        if cur and cur.get("slip"):
            docs.append(cur)
    finally:
        wb.close()
    return docs


def collect():
    seen, docs = set(), []
    for f in _files():
        for d in parse_file(f):
            # 같은 전표가 여러 배치 파일에 겹쳐 나온다(25건 배치로 나눠 받았다).
            if d["slip"] in seen:
                continue
            seen.add(d["slip"])
            docs.append(d)
    return docs


def main():
    docs = collect()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    json.dump({"count": len(docs), "docs": docs},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    if "--json" in sys.argv:
        print(OUT_JSON)
        return 0

    # 원장 06시트의 거래명세서번호와 맞춰 본다 — 어느 쪽에만 있는지가 조치 대상이다.
    try:
        from ecount_reconcile import read_ledger, load_config
        recs = read_ledger(load_config()["reconcile"]["master_xlsx"])
        lnums = {norm_slip(r.get("원장_거래명세서번호")) for r in recs.values()
                 if str(r.get("원장_거래명세서번호") or "").strip()}
        lnums.discard("")
    except Exception as e:
        recs, lnums = {}, set()
        print(f"! 원장 읽기 실패 — 대조 생략: {str(e)[:60]}")

    enums = {d["slip"] for d in docs}
    both, only_erp, only_ledger = enums & lnums, enums - lnums, lnums - enums
    total = sum(d["amount"] for d in docs)

    months = {}
    for d in docs:
        mo = d["slip"][:7]
        m = months.setdefault(mo, {"n": 0, "amt": 0})
        m["n"] += 1
        m["amt"] += d["amount"]

    L = [f"# ERP 거래명세서 인쇄본 ↔ 관리대장 대조", "",
         f"- 명세서 **{len(docs)}건** · 금액 합계 **{total:,}원** (원본 {len(_files())}개 파일)",
         f"- 원장 거래명세서번호 {len(lnums)}개 중 **일치 {len(both)}개**",
         f"- ERP 에만 있음 **{len(only_erp)}건** · 원장에만 있음 **{len(only_ledger)}건**", "",
         "## 월별", "", "| 월 | 건수 | 금액 |", "|---|---:|---:|"]
    for mo in sorted(months):
        L.append(f"| {mo} | {months[mo]['n']} | {months[mo]['amt']:,} |")
    if only_erp:
        L += ["", "## ERP 에만 있는 명세서(원장 미기재 후보, 최대 40건)", "",
              "| 전표번호 | 거래처(캠프) | 금액 |", "|---|---|---:|"]
        for d in sorted(docs, key=lambda x: x["slip"])[:400]:
            if d["slip"] in only_erp:
                L.append(f"| {d['slip']} | {d['cust'][:28]} | {d['amount']:,} |")
                if L.count("|") > 0 and len([x for x in L if x.startswith('| 2026')]) >= 40:
                    break
    L += ["", "> 읽기 전용 진단이다. 발행일·금액을 여기서 원장에 쓰지 않는다 —",
          "> 반영은 확정 근거가 있을 때 ledger_db 큐를 거친다."]
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(L))
    print(f"거래명세서 인쇄본 {len(docs)}건 · {total:,}원 | 원장 일치 {len(both)} · "
          f"ERP만 {len(only_erp)} · 원장만 {len(only_ledger)} → reports/거래명세서_대조.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
