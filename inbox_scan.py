# -*- coding: utf-8 -*-
"""
inbox_scan.py — inbox에 들어온 엑셀이 **무슨 자료인지 내용으로 판별**한다
==============================================================================
이카운트에서 내려받으면 파일명이 `8W1JR7MGB50PHOP.xlsx` 처럼 무작위다.
파일명으로 고르면(‘원장’ 포함 등) 이런 파일은 전부 무시돼 "파일이 없습니다"가 뜬다.
→ 파일을 열어 머리글을 보고 종류를 정한다. 사용자는 그냥 넣기만 하면 된다.

판별 종류
  ledger : 거래처별계정별원장  (적요 + 차변/대변)
  po     : 쿠팡 PO 목록        (PO번호 컬럼 또는 PO+숫자 값이 다수)
  sales  : 판매/세금계산서 내보내기 (공급가액 + 거래처/품목)
  unknown: 판별 실패

사용
  python inbox_scan.py                 # inbox 폴더 목록·판별 결과 출력
  from inbox_scan import pick          # pick("ledger") → 해당 종류 파일 경로 리스트
"""
import sys, os, re, glob

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.environ.get("COUPANG_INBOX_DIR") or os.path.join(BASE_DIR, "inbox")
PO_VAL = re.compile(r"\bPO[\s-]?\d{3,}\b", re.I)


def _cells(path, sheets=4, rows=30):
    """앞부분만 읽는다 — 판별에 필요한 건 머리글 근처뿐.
    read_only=True는 쓰지 않는다: 이카운트가 내보낸 파일은 <dimension>이 'A1:A1'로
    잘못 적혀 있어 read_only 모드가 1행만 읽고 멈춘다(판별 전량 실패의 원인이었다)."""
    import openpyxl
    out = []
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    try:
        for ws in wb.worksheets[:sheets]:
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i >= rows:
                    break
                out.append([("" if c is None else str(c)).strip() for c in r])
    finally:
        wb.close()
    return out


def classify_rows(rows):
    """머리글·값 패턴으로 종류 결정 (순수 함수 — 합성검증 대상)"""
    flat = [c for r in rows for c in r if c]
    joined = " ".join(flat)
    has = lambda *ks: any(k in c for c in flat for k in ks)

    # 원장: '적요'와 차변/대변이 같은 표에 있다
    for r in rows:
        names = [c for c in r if c]
        if any("적요" in n for n in names) and any(("대변" in n or "차변" in n) for n in names):
            return "ledger"
    if "계정별원장" in joined or "거래처별계정별" in joined:
        return "ledger"

    # 이카운트 매출 관련 현황 3종 (파일명이 무작위라 머리글로만 구분된다)
    for r in rows:
        n = [c for c in r if c]
        if any("일자" in x and "No" in x.replace(" ", "") for x in n) or any(x.replace(" ", "") == "일자-번호" for x in n):
            if any("매출부가세" in x or "매출합계" in x for x in n):
                return "tax"          # 매출(세금)계산서현황
            if any("부가세" in x for x in n) and any("공급가액" in x for x in n):
                return "stmt"         # 거래명세서 현황
        if any("전표번호" in x for x in n) and any("거래처명" in x for x in n):
            return "slips"            # 회계거래조회 / 회계거래현황
    if "매출(세금)계산서현황" in joined:
        return "tax"
    if "거래명세서" in joined and has("공급가액"):
        return "stmt"

    # PO: 컬럼명에 PO가 있거나, PO12345 형태 값이 여러 개
    if has("PO번호", "PO No", "PO NO", "P/O", "발주번호", "Purchase Order"):
        return "po"
    if len(PO_VAL.findall(joined)) >= 3:
        return "po"

    # 판매/세금계산서 내보내기
    if has("공급가액") and has("거래처", "거래처명", "품목", "품목명"):
        return "sales"
    return "unknown"


def classify(path):
    try:
        return classify_rows(_cells(path))
    except Exception:
        return "unknown"


def scan(folder=None):
    """[(경로, 종류)] — 임시파일(~$)은 제외"""
    folder = folder or INBOX_DIR
    out = []
    for p in sorted(glob.glob(os.path.join(folder, "*.xls*"))):
        if os.path.basename(p).startswith("~$"):
            continue
        out.append((p, classify(p)))
    return out


def pick(kind, folder=None):
    """해당 종류 파일 경로. 내용 판별이 우선이고, 판별 실패 시 파일명으로 한 번 더 본다.

    ★ 폴더를 지정하지 않으면 **원본 자료 폴더 + 로컬 inbox** 를 모두 본다.
      사용자 지시(2026-07-28)로 원본은 '0. 원본 자료' 폴더에 모으지만, 급할 때
      PC inbox 에 바로 떨어뜨리는 경우가 있어 둘 다 훑는다."""
    if folder is None:
        try:
            from source_dirs import excel_dirs
            got, seen = [], set()
            for d in excel_dirs():
                for p_, k in scan(d):
                    if k == kind and os.path.basename(p_) not in seen:
                        seen.add(os.path.basename(p_))
                        got.append(p_)
            if got:
                return got
        except Exception:
            pass
    got = [p for p, k in scan(folder) if k == kind]
    if got:
        return got
    hints = {"ledger": ("원장", "계정"), "po": ("PO",), "sales": ("판매", "매출"),
             "tax": ("계산서",), "stmt": ("명세서",), "slips": ("전표", "거래조회")}
    return [p for p, k in scan(folder) if k == "unknown"
            and any(h.lower() in os.path.basename(p).lower() for h in hints.get(kind, ()))]


LABEL = {"ledger": "거래처별계정별원장", "po": "쿠팡 PO 목록",
         "sales": "판매·세금계산서 내보내기", "tax": "매출(세금)계산서현황",
         "stmt": "거래명세서 현황", "slips": "회계거래(전표) 현황", "unknown": "판별 실패"}

if __name__ == "__main__":
    rows = scan()
    if not rows:
        print(f"inbox 비어 있음 — {INBOX_DIR}")
    # ★ 파일이 있어도 **내용이 비어 있으면** 대조가 안 된다. 2026-07-27에 ERP 내보내기
    #   3개가 '회사명' 한 줄만 있는 빈 파일이었는데, 크기(13KB·20KB)만 보면 정상처럼 보여
    #   아무도 몰랐다. 그래서 행 수를 세어 같이 보여준다.
    def rowcount(path):
        try:
            import openpyxl
            w = openpyxl.load_workbook(path, read_only=True, data_only=True)
            n = 0
            for sn in w.sheetnames:
                n += sum(1 for r in w[sn].iter_rows(values_only=True)
                         if sum(1 for x in r if x not in (None, "")) >= 3)
            w.close()
            return n
        except Exception:
            return -1
    empty = []
    for p, k in rows:
        n = rowcount(p)
        mark = ""
        if n == 0:
            mark = "  ★ 비어 있음 — 다시 내보내세요"
            empty.append(os.path.basename(p))
        elif n > 0:
            mark = f"  {n}행"
        print(f"  [{LABEL[k]:14s}] {os.path.basename(p)}  ({os.path.getsize(p)//1024}KB){mark}")
    if empty:
        print(f"\n★ 내용이 없는 파일 {len(empty)}개: " + ", ".join(empty))
        print("  이카운트에서 조회 조건(기간·거래처)을 넣고 **화면에 행이 보이는 상태**에서 내보내세요.")
    if rows:
        from collections import Counter
        c = Counter(k for _, k in rows)
        print("집계:", ", ".join(f"{LABEL[k]} {n}건" for k, n in c.items()))
