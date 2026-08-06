# -*- coding: utf-8 -*-
"""
inbox_scan.py — inbox에 들어온 엑셀이 **무슨 자료인지 내용으로 판별**한다
==============================================================================
이카운트에서 내려받으면 파일명이 `8W1JR7MGB50PHOP.xlsx` 처럼 무작위다.
파일명으로 고르면(‘원장’ 포함 등) 이런 파일은 전부 무시돼 "파일이 없습니다"가 뜬다.
→ 파일을 열어 머리글을 보고 종류를 정한다. 사용자는 그냥 넣기만 하면 된다.

판별 종류
  ledger    : 거래처별계정별원장  (적요 + 차변/대변)
  acctledger: 계정별원장          (거래처 구분이 없는 계정 단위 장부)
  journal   : 분개장              (전표번호 + 계정명 — 전 계정이 섞여 있다)
  cashbook  : 현금출납장
  po     : 쿠팡 PO 목록        (PO번호 컬럼 또는 PO+숫자 값이 다수)
  sales  : 판매/세금계산서 내보내기 (공급가액 + 거래처/품목)
  unknown: 판별 실패

★ 분개장·현금출납장을 'ledger' 로 잡으면 안 된다 (2026-08-06 실측 사고)
  예전 첫 규칙은 "같은 행에 '적요' 와 '차변|대변' 이 있으면 ledger" 였는데
  **분개장·현금출납장 머리글도 그 조건을 그대로 만족한다.**
  실측: `1. ERP 내보내기` 의 xlsx 117개 중 ledger 로 잡힌 6개 =
        거래처별계정별원장 3 + 계정별원장(빈 파일) 1 + **분개장 1 + 현금출납장 1**
  그래서 `erp_ledger_check` · `receipt_fill` 이 `pick("ledger")` 로 분개장을 집을 수
  있었다. 분개장은 **전 계정이 섞여** 있어 외상매출금 차변/대변 판정이 오염된다 —
  그 값이 입금 자동입력의 근거라 조용히 틀린 입금이 들어간다.

사용
  python inbox_scan.py                 # inbox 폴더 목록·판별 결과 출력
  from inbox_scan import pick          # pick("ledger") → 해당 종류 파일 경로 리스트
"""
import sys, os, re, glob, time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.environ.get("COUPANG_INBOX_DIR") or os.path.join(BASE_DIR, "inbox")
PO_VAL = re.compile(r"\bPO[\s-]?\d{3,}\b", re.I)


def _cells(path, sheets=4, rows=30, with_sheets=False):
    """앞부분만 읽는다 — 판별에 필요한 건 머리글 근처뿐.
    read_only=True는 쓰지 않는다: 이카운트가 내보낸 파일은 <dimension>이 'A1:A1'로
    잘못 적혀 있어 read_only 모드가 1행만 읽고 멈춘다(판별 전량 실패의 원인이었다).

    with_sheets=True 면 (행들, 시트이름들) 을 준다. 기본값은 예전 그대로 행들만
    돌려준다 — `upload_intake.py` 가 이 함수를 쓰고 있어 반환형을 바꾸지 않는다."""
    import openpyxl
    out, titles = [], []
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    try:
        for ws in wb.worksheets[:sheets]:
            titles.append(ws.title)
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i >= rows:
                    break
                out.append([("" if c is None else str(c)).strip() for c in r])
    finally:
        wb.close()
    return (out, titles) if with_sheets else out


# 이카운트는 **시트 이름에 화면 이름을 그대로** 쓴다(2026-08-06 실측: 파일명이
# `hzwL4LeZiaH84jBr.xlsx` 처럼 무작위여도 시트명은 '분개장' 이었다). 머리글보다
# 확실한 지문이라 가장 먼저 본다.
# ★ 순서가 곧 정확도다 — '거래처별계정별원장' 이 '계정별원장' 을 글자로 품고 있어
#   긴 이름을 먼저 보지 않으면 거래처별원장이 계정별원장으로 떨어진다.
SHEET_KIND = (
    ("거래처별계정별원장", "ledger"),
    ("거래처별계정별", "ledger"),
    ("계정별원장", "acctledger"),
    ("현금출납장", "cashbook"),
    ("분개장", "journal"),
)


def classify_rows(rows, sheets=None):
    """머리글·값 패턴으로 종류 결정 (순수 함수 — 합성검증 대상)

    sheets: 시트 이름 목록(있으면 가장 먼저·가장 세게 믿는다)."""
    flat = [c for r in rows for c in r if c]
    joined = " ".join(flat)
    has = lambda *ks: any(k in c for c in flat for k in ks)

    # 0) 시트 이름이 가장 정확한 지문이다 — 있으면 여기서 끝낸다.
    for title in sheets or ():
        t = "".join(str(title).split())
        for name, kind in SHEET_KIND:
            if name in t:
                return kind

    # 1) 분개장을 **원장보다 먼저** 가려낸다.
    #    분개장 머리글은 전표번호/계정명/거래처/차변/대변/적요 라서 아래 원장 규칙
    #    ('적요'+차변/대변)을 그대로 만족한다. 순서를 바꾸면 다시 ledger 로 샌다.
    #    구분점은 **'계정명'** 이다 — 거래처별계정별원장은 계정이 이미 고정돼 있어
    #    행마다 계정명을 달지 않는다.
    for r in rows:
        n = [c for c in r if c]
        if any("전표번호" in x for x in n) and any("계정명" in x for x in n):
            return "journal"

    # 2) 표 제목 줄에 화면 이름이 찍혀 나오는 경우(시트명을 못 읽었을 때의 보루).
    for name, kind in SHEET_KIND:
        if name in joined:
            return kind

    # 3) 원장: '적요'와 차변/대변이 같은 표에 있다.
    #    여기까지 왔다는 건 분개장·현금출납장·계정별원장이 아니라는 뜻이다.
    for r in rows:
        names = [c for c in r if c]
        if any("적요" in n for n in names) and any(("대변" in n or "차변" in n) for n in names):
            return "ledger"

    # 홈택스 전자(세금)계산서 리스트 — '승인번호'가 있으면 이것 말고 없다.
    # (회계 I > 전자(세금)계산서 > 홈택스자료조회 에서 Excel 로 내려받은 것)
    if has("승인번호") and has("공급자사업자번호", "공급받는자사업자번호", "공급자상호"):
        return "hometax"

    # 회계 I 출력물 [매출(세금)계산서현황] — 일자-No. + 매출부가세/매출합계 조합이
    # 정확한 지문이다. ★ 2026-08-03 실측: 쿠팡매출청구서현황 양식에도 '내역보기' 열이
    # 있어서 아래 taxinv 휴리스틱이 먼저 삼켰다 — 이 검사가 반드시 taxinv 보다 앞서야
    # erp_docs_check 가 원본을 찾는다.
    for r in rows:
        n = [c for c in r if c]
        if (any("일자" in x and "No" in x.replace(" ", "") for x in n)
                and any("매출부가세" in x or "매출합계" in x for x in n)):
            return "tax"

    # 매출(세금)계산서조회(재고) — 재고 I > 영업관리 > 판매일괄회계반영.
    # 거래명세서 현황과 머리글이 겹쳐(공급가액+부가세) 예전엔 'stmt' 로 잘못 잡혔다.
    if "매출(세금)계산서조회" in joined:
        return "taxinv"
    if has("내역보기") and has("공급가액") and has("부가세"):
        return "taxinv"

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

    # 입금/수금 원본 — 파일명이 무작위여도 날짜+금액 머리글 조합이면 독립 원천이다.
    if has("입금일", "수금일", "입금일자", "수금일자") and has("입금액", "수금액") \
            and has("거래처", "프로젝트", "캠프"):
        return "receipt"

    # ERP 판매조회에는 PO번호 열이 함께 있어 아래 PO 규칙이 먼저 잡으면 쿠팡 목록으로
    # 잘못 분류된다. 판매 진행상태+공급가액+거래처/품목 조합은 ERP 판매가 우선이다.
    if has("진행상태") and has("공급가액") and has("거래처", "거래처명", "품목", "품목명"):
        return "sales"

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
        rows, titles = _cells(path, with_sheets=True)
        return classify_rows(rows, titles)
    except Exception:
        return "unknown"


# ── 분류 결과 캐시 ────────────────────────────────────────────
#   ★ 2026-07-30 실측: classify() 는 파일을 **열어서** 내용으로 종류를 정한다.
#     scan() 이 '0. 원본 자료' 트리의 엑셀 전부를 열고, /api/status 가 pick() 을
#     5번(ledger·slips·po·tax·stmt) 부르므로 같은 파일을 5번 열었다. Z: 네트워크
#     드라이브라 대시보드 첫 로딩이 **280초** 걸렸다(측정값).
#     파일 내용이 바뀌지 않았으면 다시 열 이유가 없다 — (크기, 수정시각) 으로 판정한다.
#   ★ 2026-08-06: 캐시 열쇠에 **판별기 판번호**를 넣었다. 예전 열쇠는 (크기, 수정시각)
#     뿐이라, 판별 규칙을 고쳐도 **파일이 그대로면 옛 판정이 그대로 나왔다.**
#     분개장을 원장으로 잡던 버그를 고쳐도 이미 캐시에 'ledger' 로 박힌 파일은
#     영원히 원장으로 남는다 — 고친 것이 실제로는 안 고쳐지는 조용한 실패다.
#     규칙을 바꿀 때마다 이 번호를 올린다.
CLS_VER = 2                                     # 2: 분개장·현금출납장·계정별원장 분리
_CLS_MEM = {}                                   # 프로세스 안: 경로 → (크기, mtime, 종류)
_CLS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "reports", "inbox_classify.json")
_CLS_DISK = None                                # 프로세스 간: 재시작해도 유지


def _cls_load():
    global _CLS_DISK
    if _CLS_DISK is None:
        try:
            import json as _json
            _CLS_DISK = _json.load(open(_CLS_FILE, encoding="utf-8"))
        except Exception:
            _CLS_DISK = {}
    return _CLS_DISK


def _cls_save():
    try:
        import json as _json
        os.makedirs(os.path.dirname(_CLS_FILE), exist_ok=True)
        _json.dump(_CLS_DISK, open(_CLS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass


def classify_cached(path):
    """내용이 그대로면 캐시값을 쓴다. 바뀌었으면 그때만 다시 열어 본다."""
    try:
        st = os.stat(path)
        sig = [st.st_size, int(st.st_mtime)]
    except OSError:
        return classify(path)
    hit = _CLS_MEM.get(path)
    if hit and hit[0] == sig[0] and hit[1] == sig[1]:
        return hit[2]
    disk = _cls_load().get(path)
    # 판번호가 다르면(=판별 규칙이 바뀌었으면) 캐시를 믿지 않고 다시 연다.
    if disk and disk[0] == sig[0] and disk[1] == sig[1] \
            and len(disk) > 3 and disk[3] == CLS_VER:
        _CLS_MEM[path] = (sig[0], sig[1], disk[2])
        return disk[2]
    kind = classify(path)
    _CLS_MEM[path] = (sig[0], sig[1], kind)
    _cls_load()[path] = [sig[0], sig[1], kind, CLS_VER]
    _cls_save()
    return kind


_SCAN_MEM = {}                                  # 폴더 → (만료시각, [(경로, 종류)])
SCAN_TTL = 60.0                                 # 초. 한 요청 안의 반복 호출을 합치는 용도


def scan(folder=None, ttl=None):
    """[(경로, 종류)] — 임시파일(~$)은 제외. 분류는 캐시, 폴더 훑기는 짧게 메모이즈한다.

    ★ 분류를 캐시한 뒤에도 대시보드가 느렸다. `pick()` 이 5종(ledger·slips·po·tax·stmt)
      마다 **네트워크 드라이브를 다시 훑었기** 때문이다(재귀 glob + 파일마다 os.stat).
      종류가 다르다고 폴더가 달라지는 게 아니므로 훑기 결과를 잠깐 재사용한다.
      TTL 을 짧게(60초) 두어 새 파일이 들어오면 다음 분에는 잡힌다."""
    folder = folder or INBOX_DIR
    ttl = SCAN_TTL if ttl is None else ttl
    now = time.monotonic()
    hit = _SCAN_MEM.get(folder)
    if hit and hit[0] > now:
        return hit[1]
    out = []
    for p in sorted(glob.glob(os.path.join(folder, "**", "*.xls*"), recursive=True)):
        if os.path.basename(p).startswith("~$"):
            continue
        out.append((p, classify_cached(p)))
    _SCAN_MEM[folder] = (now + ttl, out)
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
    # ★ 'ledger' 의 파일명 힌트에서 '계정' 을 뺐다 — 분개장·계정별원장 파일명까지
    #   빨아들여, 내용 판별로 갈라 놓은 것을 파일명이 도로 뭉개고 있었다.
    hints = {"ledger": ("거래처별계정별원장",), "acctledger": ("계정별원장",),
             "journal": ("분개장",), "cashbook": ("현금출납장", "출납"),
             "po": ("PO",), "sales": ("판매", "매출"),
             "tax": ("계산서",), "stmt": ("명세서",), "slips": ("전표", "거래조회"),
             "receipt": ("입금", "수금", "송금")}
    return [p for p, k in scan(folder) if k == "unknown"
            and any(h.lower() in os.path.basename(p).lower() for h in hints.get(kind, ()))]


LABEL = {"ledger": "거래처별계정별원장", "acctledger": "계정별원장",
         "journal": "분개장", "cashbook": "현금출납장", "po": "쿠팡 PO 목록",
         "sales": "판매·세금계산서 내보내기", "tax": "매출(세금)계산서현황",
         "stmt": "거래명세서 현황", "slips": "회계거래(전표) 현황",
         "taxinv": "매출(세금)계산서조회(재고)", "hometax": "홈택스 전자(세금)계산서",
         "receipt": "입금·수금 내역",
         "unknown": "판별 실패"}

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
