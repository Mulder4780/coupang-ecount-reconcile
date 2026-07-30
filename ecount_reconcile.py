# -*- coding: utf-8 -*-
"""
ecount_reconcile.py — 관리대장 ↔ 이카운트(ECOUNT) 대조기
==========================================================
관리대장(원장)에 기록된 "예상"과 이카운트에 실제 등록된 전표를 3개 문서층에서 대조:

  ① 거래명세서/판매전표  : 원장 06(정산) O·P·Q 금액  ↔  이카운트 판매전표
  ② 전자세금계산서        : 원장 15 실제발행일·발행금액·승인번호  ↔  이카운트 세금계산서
  ③ 수금/입금             : 원장 16 입금일·입금액  ↔  이카운트 수금

판정: 일치 · 금액불일치 · 이카운트미등록(원장에만) · 원장누락(이카운트에만)

원칙:
- 관리대장(마스터 xlsx)은 read-only 로만 연다. (openpyxl save 는 차트·도형·검증을 손상시키므로 절대 저장하지 않음)
- 결과는 별도 리포트로만 출력: reports/ 폴더에 CSV·Markdown·독립 xlsx.
- 이카운트 미연결(설정 미완/ IP 미등록)이어도 "원장 기준 대조 준비표"를 먼저 만들어 즉시 활용 가능.

실행:
    python ecount_reconcile.py            # 이카운트 연결 시도 후 대조, 실패하면 원장측만
    python ecount_reconcile.py --offline  # 이카운트 호출 없이 원장측 준비표만
    python ecount_reconcile.py --selftest # 설정/인증만 점검
"""
import sys, os, csv, json, time
from datetime import datetime, date
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
from ecount_client import EcountClient, load_config, days_until_expiry, EcountError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(BASE_DIR, "reports")

HDR_ROW = 4          # 관리대장 공통 상수 — 절대 변경 금지(수천 수식이 의존)
FIRST_DATA = 5


# ───────────────────────── 원장 읽기 ─────────────────────────
def _num(v):
    if v in (None, "", "-"):
        return None
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def _money(v):
    """금액 표시는 모든 보고서에서 소수점 없이 천 단위 쉼표로 통일한다."""
    if v in (None, "", "-"):
        return ""
    try:
        return f"{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return str(v)


def _d(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return "" if v is None else str(v)


def _latest_in(folder):
    """폴더에서 v번호가 가장 큰 관리대장. 구 버전 정리의 기준점으로만 쓴다."""
    import glob, re as _re
    best, bv = None, -1
    for p in glob.glob(os.path.join(folder, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx")):
        m = _re.search(r"_v(\d+)\.xlsx$", p)
        if m and "~$" not in p and int(m.group(1)) > bv:
            best, bv = p, int(m.group(1))
    return best or folder


_MASTER_CACHE = {}          # xlsx_path -> (잰시각, 결과경로)
_MASTER_TTL = 30.0          # 초


def resolve_master(xlsx_path):
    """최신 관리대장을 찾되 **30초 안에는 다시 찾지 않는다**(2026-07-31 속도 개선).

    이 함수는 네트워크 드라이브에서 glob 을 돌리고 autoprune 까지 부른다. 그런데
    API 한 번에 7번씩 불리고 있었다(`real_works` 측정). 어느 파일인지만 고르는 일이라
    30초 캐시로 충분하다 — 내용이 바뀌었는지는 read_ledger 가 mtime 으로 따로 본다.
    새 프로세스는 항상 차갑게 시작하므로 낡은 경로를 오래 붙들 일은 없다.
    """
    now = time.monotonic()
    hit = _MASTER_CACHE.get(xlsx_path)
    if hit is not None and now - hit[0] < _MASTER_TTL:
        return hit[1]
    result = _resolve_master_uncached(xlsx_path)
    _MASTER_CACHE[xlsx_path] = (now, result)
    return result


def _resolve_master_uncached(xlsx_path):
    """설정된 경로가 없으면 같은 폴더에서 최신 v번호 파일을 자동 탐지한다
    (관리대장은 v19→v20→v21…로 계속 버전업되며 구버전은 OLD로 이동됨)."""
    import glob, re as _re
    # 구 버전을 OLD 로 접는다 — 사용자 지시(2026-07-28) "말 안 해도 들어가게".
    # 저장하는 쪽이 11군데라 여기(찾는 쪽) 한 곳에만 건다. 실패해도 본 작업은 그대로 간다.
    # autoprune 은 관리대장 이름이 아니면 스스로 아무것도 하지 않는다(합성검증 파일 보호).
    try:
        from ledger_versions import autoprune
        autoprune(xlsx_path if os.path.exists(xlsx_path)
                  else _latest_in(os.path.dirname(xlsx_path)))
    except Exception:
        pass
    if os.path.exists(xlsx_path):
        return xlsx_path
    folder = os.path.dirname(xlsx_path)
    cands = glob.glob(os.path.join(folder, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx"))
    def ver(p):
        m = _re.search(r"_v(\d+)\.xlsx$", p)
        return int(m.group(1)) if m else -1
    cands = [c for c in cands if ver(c) >= 0 and "~$" not in c]
    if not cands:
        raise FileNotFoundError(f"관리대장을 찾을 수 없음: {xlsx_path} (폴더에 v*.xlsx 없음)")
    best = max(cands, key=ver)
    print(f"i 관리대장 최신본 자동 탐지: {os.path.basename(best)}")
    return best


_MASTER_BYTES = {}          # abspath -> ((mtime_ns, size), bytes)


def master_stream(xlsx_path):
    """관리대장을 **네트워크가 아니라 메모리에서** 열게 해 준다(2026-07-31 속도 개선).

    앱은 한 화면을 그리는 동안 `openpyxl.load_workbook(master, ...)` 를 9군데에서
    각각 부른다. 그때마다 **Z: 네트워크 드라이브에서 수십 MB 를 다시 끌어온다** — 느린
    것의 대부분이 파싱이 아니라 이 전송이다. 한 번 읽어 두고 BytesIO 로 건네면 파싱만 남는다.

    · 무효화는 mtime+크기. vN+1 이 생기면 키가 달라져 자동으로 다시 읽는다.
    · 새 버전을 읽을 때 캐시를 **비운다** — 안 비우면 버전이 쌓일수록 메모리를 계속 먹는다.
    · 파일을 못 재면 원래 경로를 그대로 돌려준다(호출부는 경로든 스트림이든 받는다).
    """
    import io
    key = os.path.abspath(xlsx_path)
    try:
        st = os.stat(key)
        sig = (st.st_mtime_ns, st.st_size)
    except OSError:
        return xlsx_path
    hit = _MASTER_BYTES.get(key)
    if hit is None or hit[0] != sig:
        with open(key, "rb") as f:
            data = f.read()
        _MASTER_BYTES.clear()
        _MASTER_BYTES[key] = (sig, data)
        hit = _MASTER_BYTES[key]
    return io.BytesIO(hit[1])


_LEDGER_CACHE = {}          # abspath -> ((mtime_ns, size), 원장dict)


def read_ledger(xlsx_path):
    """원장을 읽되, **파일이 그대로면 다시 열지 않는다**(2026-07-31 속도 개선).

    관리대장은 **네트워크 드라이브(Z:)** 에 있고 한 번 여는 데 ~1.0초가 든다. 그런데
    캐시가 없어서 API 한 번에 여러 번 다시 열고 있었다 — `real_works` 는 한 호출에
    7번 열었다(측정: settlements 2.5초, works 3.5초).

    무효화는 **파일의 mtime+크기**로 한다. 시간(TTL)으로 하면 vN+1 이 새로 생겼는데도
    남은 시간 동안 옛 숫자를 보여 주게 된다 — 원장은 그러면 안 된다. 파일이 바뀌는
    순간 키가 달라져 자동으로 다시 읽는다.

    돌려줄 때 **사본**을 준다. 호출자가 받은 dict 를 고쳐도 캐시가 오염되지 않아야 한다.
    """
    key = os.path.abspath(xlsx_path)
    try:
        st = os.stat(key)
        sig = (st.st_mtime_ns, st.st_size)
    except OSError:
        return _read_ledger_uncached(xlsx_path)      # 못 재면 그냥 읽는다
    hit = _LEDGER_CACHE.get(key)
    if hit is not None and hit[0] == sig:
        return _copy_ledger(hit[1])
    data = _read_ledger_uncached(xlsx_path)
    _LEDGER_CACHE[key] = (sig, data)
    return _copy_ledger(data)


def _copy_ledger(d):
    """원장은 {정산ID: {열: 값}} 2단이고 값은 전부 단순형이라 2단 복사면 충분하다.
       copy.deepcopy 는 datetime 까지 새로 만들어 느리다."""
    return {k: dict(v) for k, v in d.items()}


def _read_ledger_uncached(xlsx_path):
    """정산ID 기준으로 원장 예상값(3개 문서층)을 모은다."""
    import openpyxl
    xlsx_path = resolve_master(xlsx_path)
    wb = openpyxl.load_workbook(master_stream(xlsx_path), read_only=True, data_only=True)

    def col_index(ws, names):
        row = next(ws.iter_rows(min_row=HDR_ROW, max_row=HDR_ROW, values_only=True))
        idx = {}
        for i, h in enumerate(row):
            if h is not None:
                idx[str(h).strip()] = i
        return {n: idx.get(n) for n in names}

    recs = {}   # 정산ID -> dict

    # 06 거래서류청구수금 : 기준 원장
    ws = wb["06_거래서류청구수금"]
    c = col_index(ws, ["정산ID", "업무구분", "원천업무ID", "프로젝트NO", "캠프명", "작업완료일",
                       "비용구분", "실제작업공급가액", "실제작업부가세", "실제작업합계",
                       "거래명세서번호", "거래명세서발행일", "거래명세서공급가액", "거래명세서합계",
                       "세금계산서발행일", "세금계산서합계", "입금일", "입금액",
                       "청구일", "지급예정일", "PO필요여부", "PO번호", "PO발행일"])
    for row in ws.iter_rows(min_row=FIRST_DATA, values_only=True):
        sid = row[c["정산ID"]] if c["정산ID"] is not None else None
        if not sid:
            continue
        recs[str(sid)] = {
            "정산ID": str(sid),
            "업무구분": _d(row[c["업무구분"]]),
            "원천업무ID": _d(row[c["원천업무ID"]]),
            "프로젝트NO": _d(row[c["프로젝트NO"]]),
            "캠프명": _d(row[c["캠프명"]]),
            "작업완료일": _d(row[c["작업완료일"]]),
            "비용구분": _d(row[c["비용구분"]]),
            "원장_공급가액": _num(row[c["실제작업공급가액"]]),
            # 부가세는 **원장에 적힌 값**을 그대로 쓴다. 공급가액×10%로 계산하면 반올림 때문에
            # 원장·계산서와 1원씩 어긋날 수 있고, 화면 숫자가 서류와 달라지면 신뢰를 잃는다.
            "원장_부가세": _num(row[c["실제작업부가세"]]) if "실제작업부가세" in c else None,
            "원장_합계": _num(row[c["실제작업합계"]]),
            "원장_거래명세서번호": _d(row[c["거래명세서번호"]]),
            "원장_거래명세서발행일": _d(row[c["거래명세서발행일"]]),
            "원장_거래명세서합계": _num(row[c["거래명세서합계"]]),
            "원장_세금계산서발행일": _d(row[c["세금계산서발행일"]]),
            "원장_세금계산서합계": _num(row[c["세금계산서합계"]]),
            "원장_입금일": _d(row[c["입금일"]]),
            "원장_입금액": _num(row[c["입금액"]]),
            "원장_청구일": _d(row[c["청구일"]]) if c["청구일"] is not None else "",
            "원장_지급예정일": _d(row[c["지급예정일"]]) if c["지급예정일"] is not None else "",
            "원장_PO필요여부": _d(row[c["PO필요여부"]]) if c["PO필요여부"] is not None else "",
            "원장_PO번호": _d(row[c["PO번호"]]) if c["PO번호"] is not None else "",
            "원장_PO발행일": _d(row[c["PO발행일"]]) if c["PO발행일"] is not None else "",
        }

    # 15 세금계산서관리 : 승인번호·실제발행일 보강
    ws = wb["15_세금계산서관리"]
    c = col_index(ws, ["정산ID", "실제발행일", "발행금액", "승인번호"])
    for row in ws.iter_rows(min_row=FIRST_DATA, values_only=True):
        sid = row[c["정산ID"]] if c["정산ID"] is not None else None
        if sid and str(sid) in recs:
            r = recs[str(sid)]
            r["원장_세금계산서실제발행일"] = _d(row[c["실제발행일"]])
            r["원장_세금계산서승인번호"] = _d(row[c["승인번호"]])
            if _num(row[c["발행금액"]]):
                r["원장_세금계산서합계"] = _num(row[c["발행금액"]]) or r.get("원장_세금계산서합계")

    # 16 입금수금관리 : 입금 보강
    ws = wb["16_입금수금관리"]
    c = col_index(ws, ["정산ID", "입금일", "입금액", "미수금액"])
    for row in ws.iter_rows(min_row=FIRST_DATA, values_only=True):
        sid = row[c["정산ID"]] if c["정산ID"] is not None else None
        if sid and str(sid) in recs:
            r = recs[str(sid)]
            if _d(row[c["입금일"]]):
                r["원장_입금일"] = _d(row[c["입금일"]])
            if _num(row[c["입금액"]]) is not None:
                r["원장_입금액"] = _num(row[c["입금액"]])
            r["원장_미수금액"] = _num(row[c["미수금액"]])

    wb.close()
    return recs


# ───────────────────────── 이카운트 값 정규화 ─────────────────────────
# 이카운트 조회 응답의 필드명은 계정/버전에 따라 다르므로 후보군에서 유연 추출
CAND = {
    "금액":   ["SUPPLY_AMT", "SUPPLY_AMT_KRW", "AMT", "TOTAL_AMT", "PRICE", "SUP_AMT", "TTL_AMT"],
    "일자":   ["IO_DATE", "BASE_DATE", "IO_DATE_KRW", "DATE", "WR_DATE", "REG_DATE"],
    "적요":   ["REMARKS", "E_TEXT", "U_MEMO", "WH_DES", "REMARKS_WIN", "PRJ_CD"],
    "거래처": ["CUST_DES", "CUST", "CUST_NM", "EMP_NM"],
    "번호":   ["IO_NO", "DOC_NO", "TAX_NO", "IV_NO", "SLIP_NO", "APPROVE_NO", "NTS_CONFIRM_NUM"],
}


def has_statement(r):
    """거래명세서가 실제로 있는가 — 번호 하나로 판단하지 않는다.

    billing_fill.py 가 ERP 판매조회를 근거로 채운 건들은 **발행일·금액은 있는데 번호가
    없다**(629건). 번호만 보면 이 전부가 '미청구(전표 없음)'로 잡히는데, 명세서는 실제로
    나갔고 번호만 원장에 안 옮겨진 것이라 거짓 경보다. 발행일+금액이 함께 있으면 있는 것으로 본다.
    """
    if str(r.get("원장_거래명세서번호") or "").strip():
        return True
    return bool(r.get("원장_거래명세서발행일")) and bool(r.get("원장_거래명세서합계"))


def settle_status(r):
    """정산 1건의 상태를 **한 곳에서** 판정한다.

    화면(webapp/app_server.py)과 엑셀(findings_export.py)이 각자 같은 if 사다리를 들고
    있어서 한쪽만 고치면 두 숫자가 어긋났다. 여기로 모은다.

    ★ 금액(06시트 I열 실제작업공급가액) 판정에 조건이 두 개 붙는 이유:
      ① I열은 750행 중 684행이 **수식**이고, 그 수식은
         `IF(LEFT(원천업무ID,3)="AS-", SUMIF('03_현장작업실적'!B:B, 원천업무ID, ...!Q:Q), "")`
         이다. 즉 **돌발AS(AS-)가 아니면 설계상 영원히 빈칸**이다. 정기점검(PM-) 305건을
         '금액 미입력'으로 잡던 것은 오탐이었고, 그 오탐이 if 사다리 맨 위에 있어서
         그 305건의 **진짜 문제(세금계산서 미발행)를 가리고 있었다.**
      ② AS- 인데 비어 있는 건은, 03시트에 실적 행이 아직 안 생겨서 수식이 0인 것이다.
         03시트 행 생성과 재계산은 지금 차단된 작업이다(AGENTS.md 0번). 거래명세서 금액이라는
         **근거가 있으면** '금액 재계산 대기'로 따로 부른다 — 사람이 손으로 넣을 일이 아니라
         재계산이 풀리면 저절로 채워질 건이라, 조치 목록에 섞어 두면 안 된다.
      ★ I열에 값을 직접 써넣지 말 것. 수식을 덮으면 AE(작업대비거래명세서차액)가 항상 0이 되어
        이 시트의 존재 이유인 '작업금액과 명세서금액의 차이 드러내기'가 죽는다.
    """
    if r.get("비용구분") != "유상":
        return "무상/보험"
    src = str(r.get("원천업무ID") or "")
    if not r.get("원장_공급가액") and src.startswith("AS-"):
        return "금액 재계산 대기" if r.get("원장_거래명세서합계") else "금액 미입력"
    if not has_statement(r):
        return "미청구(전표 없음)"
    if not (r.get("원장_세금계산서실제발행일") or r.get("원장_세금계산서발행일")):
        return "세금계산서 미발행"
    if not r.get("원장_입금일"):
        return "입금 대기"
    return "정상"


def pick(row, kind):
    for k in CAND[kind]:
        if isinstance(row, dict) and row.get(k) not in (None, ""):
            return row[k]
    return None


def norm_ecount(rows):
    out = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        amt = pick(r, "금액")
        try:
            amt = round(float(str(amt).replace(",", "")), 2) if amt is not None else None
        except ValueError:
            amt = None
        out.append({
            "금액": amt,
            "일자": _d(pick(r, "일자")),
            "적요": _d(pick(r, "적요")),
            "거래처": _d(pick(r, "거래처")),
            "번호": _d(pick(r, "번호")),
            "_raw": r,
        })
    return out


def match_project(ledger_rec, ecount_rows, tol_amt, filter_cust):
    """프로젝트NO(적요/번호 내 포함) 1순위, 금액 2순위로 이카운트 전표 1건 매칭."""
    prj = (ledger_rec.get("프로젝트NO") or "").strip()
    exp_amt = ledger_rec.get("원장_공급가액")
    cust_ok = lambda e: (not filter_cust) or (filter_cust in (e.get("거래처") or "") or not e.get("거래처"))
    # 1순위: 프로젝트NO 문자열 포함
    if prj:
        for e in ecount_rows:
            blob = f"{e.get('적요','')} {e.get('번호','')} {json.dumps(e.get('_raw',{}), ensure_ascii=False)}"
            if prj in blob and cust_ok(e):
                return e, "프로젝트NO"
    # 2순위: 금액 근접
    if exp_amt is not None:
        for e in ecount_rows:
            if e.get("금액") is not None and abs(e["금액"] - exp_amt) <= tol_amt and cust_ok(e):
                return e, "금액"
    return None, None


# ───────────────────────── inbox(수동 내보내기) 로더 ─────────────────────────
INBOX_DIR = os.path.join(BASE_DIR, "inbox")

def load_inbox(cfg):
    """이카운트 화면에서 내려받은 판매현황/세금계산서 엑셀을 inbox/에서 읽어
    API 조회 결과와 동일한 형태(CAND 키)로 정규화한다.
    파일 분류: 파일명에 '판매'→sale, '세금계산서' 또는 '계산서'→tax_invoice."""
    import openpyxl, glob
    # ★ 2026-07-30 수정 이유: 여기는 `INBOX_DIR/*.xlsx` **비재귀** 글롭이라
    #   (1) 원본 자료 폴더(`0. 원본 자료/1. ERP 내보내기/2026/07/2026-07-25/...`)를 아예 못 봤고
    #   (2) 로컬 inbox 의 하위 폴더도 못 봤다. 실제로 판매조회(898행)가 원본 자료 폴더에만
    #   있어서 이 대조기는 그 자료를 한 번도 읽지 못했다.
    #   경로는 source_dirs 한 곳에서만 정한다는 원칙(AGENTS.md)에 맞춰 excel_dirs() 를 쓴다.
    from source_dirs import excel_dirs
    from billing_fill import dedupe_files
    from inbox_scan import classify
    cands, seen = [], set()
    for d in excel_dirs() or [INBOX_DIR]:
        for f in glob.glob(os.path.join(d, "**", "*.xlsx"), recursive=True):
            if os.path.basename(f).startswith("~$"):
                continue
            key = os.path.normcase(os.path.abspath(f))
            if key in seen:
                continue
            seen.add(key)
            cands.append(f)
    # ★ 같은 내용의 사본을 여러 번 읽지 않는다. 판매조회가 SHA256 동일한 3벌 있었고
    #   (2026-07-30 3배 합산 사고) 파일명(`__dup_`)으로 거르면 다음번에 다른 이름으로 뚫린다.
    files = dedupe_files(cands)
    if not files:
        return None
    out = {"sale": [], "tax_invoice": []}
    used = []
    for f in files:
        name = os.path.basename(f)
        # 파일명이 무작위인 이카운트 내보내기가 섞여 있다. **내용**으로 종류를 먼저 본다.
        # 원장(차변/대변)은 매출 후보로 쓰면 엉뚱한 금액이 매칭돼 거짓 '일치'가 되므로 뺀다.
        # ★ `po` 는 빼지 않는다: ERP 판매조회에 PO번호 열이 있어 classify() 가 이 파일을
        #   `po` 로 준다(2026-07-30 확인). 여기서 po 를 빼면 정작 필요한 판매 자료가 빠진다.
        #   쿠팡 PO 목록(오더 번호·고객·금액)은 아래 머리글 조건에 '공급가액/합계금액' 이
        #   없어 자연히 걸러진다 — 종류 판정에 이중으로 의존하지 않는다.
        ck = classify(f)
        if ck == "ledger":
            continue
        if ck in ("tax", "taxinv", "hometax"):
            kind = "tax_invoice"
        elif ck in ("sales", "stmt", "slips"):
            kind = "sale"
        else:
            kind = "tax_invoice" if ("세금계산서" in name or "계산서" in name) else "sale"
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            # 머리글 행 탐색: '공급가액' 또는 '합계' + '거래처'류가 있는 행
            hdr_i, hdr = None, None
            for i, r in enumerate(rows[:15]):
                cells = [str(c) for c in r if c is not None]
                joined = " ".join(cells)
                if ("공급가액" in joined or "합계금액" in joined) and len(cells) >= 3:
                    hdr_i, hdr = i, r
                    break
            if hdr_i is None:
                continue
            def col(*keys):
                for j, h in enumerate(hdr):
                    if h is None:
                        continue
                    hs = str(h)
                    if any(k in hs for k in keys):
                        return j
                return None
            j_amt  = col("공급가액")
            j_tot  = col("합계금액", "합계")
            j_date = col("일자", "날짜", "작성일")
            j_no   = col("승인번호", "전표", "No", "번호")
            j_cust = col("거래처")
            for r in rows[hdr_i + 1:]:
                if r is None or all(c is None for c in r):
                    continue
                amt = r[j_amt] if j_amt is not None else (r[j_tot] if j_tot is not None else None)
                if amt is None or _num(amt) is None:
                    continue
                blob = " ".join(str(c) for c in r if c is not None)
                if "합계" in blob[:10]:      # 합계행 제외
                    continue
                out[kind].append({
                    "SUPPLY_AMT": _num(amt),
                    "IO_DATE": _d(r[j_date]) if j_date is not None else "",
                    "IO_NO": _d(r[j_no]) if j_no is not None else "",
                    "CUST_DES": _d(r[j_cust]) if j_cust is not None else "",
                    "REMARKS": blob,
                })
        wb.close()
        used.append(f"{name}→{kind}")
    if not (out["sale"] or out["tax_invoice"]):
        return None
    out["_files"] = used
    return out


# ───────────────────────── 대조 로직 ─────────────────────────
def reconcile(recs, ecount, cfg):
    rc = cfg["reconcile"]
    tol = rc.get("금액허용오차", 0)
    fc = rc.get("거래처_필터", "")
    sale = norm_ecount(ecount.get("sale"))
    tax = norm_ecount(ecount.get("tax_invoice"))
    online = ecount.get("_online", False)

    results = []
    for sid, r in sorted(recs.items()):
        # 원장이 유상이 아니면 세금계산서/판매 대조 제외 안내
        유상 = (r.get("비용구분") == "유상")

        # ① 판매/거래명세서 층
        if online:
            m, how = match_project(r, sale, tol, fc)
            if m:
                if r.get("원장_공급가액") is not None and m.get("금액") is not None and abs(m["금액"] - r["원장_공급가액"]) <= tol:
                    s1 = "일치"
                elif m.get("금액") is None:
                    s1 = "이카운트금액없음"
                else:
                    s1 = f"금액불일치(원장 {_money(r.get('원장_공급가액'))} / EC {_money(m.get('금액'))})"
                ec_sale_no = m.get("번호")
            else:
                s1 = "이카운트미등록" if 유상 else "해당없음(무상)"
                ec_sale_no = ""
        else:
            s1 = "원장기준-EC미조회"
            ec_sale_no = ""

        # ② 세금계산서 층
        if online:
            m2, _ = match_project(r, tax, tol, fc)
            if m2:
                ec_tax_no = m2.get("번호")
                led_no = r.get("원장_세금계산서승인번호") or ""
                if led_no and ec_tax_no and led_no != ec_tax_no:
                    s2 = f"승인번호불일치(원장 {led_no} / EC {ec_tax_no})"
                else:
                    s2 = "일치"
            else:
                led_issued = r.get("원장_세금계산서실제발행일") or ""
                if led_issued:
                    s2 = "원장발행-EC미확인"
                else:
                    s2 = "양측미발행" if 유상 else "해당없음(무상)"
                ec_tax_no = ""
        else:
            s2 = "원장기준-EC미조회"
            ec_tax_no = ""

        results.append({
            "정산ID": sid,
            "업무구분": r.get("업무구분"),
            "프로젝트NO": r.get("프로젝트NO"),
            "캠프명": r.get("캠프명"),
            "비용구분": r.get("비용구분"),
            "작업완료일": r.get("작업완료일"),
            "원장_공급가액": r.get("원장_공급가액"),
            "원장_거래명세서번호": r.get("원장_거래명세서번호"),
            "EC_판매번호": ec_sale_no,
            "①판매/명세서_판정": s1,
            "원장_세금계산서발행일": r.get("원장_세금계산서실제발행일") or r.get("원장_세금계산서발행일"),
            "원장_세금계산서승인번호": r.get("원장_세금계산서승인번호"),
            "EC_세금계산서번호": ec_tax_no,
            "②세금계산서_판정": s2,
            "원장_입금액": r.get("원장_입금액"),
            "원장_미수금액": r.get("원장_미수금액"),
        })
    return results


# ───────────────────────── 리포트 출력 ─────────────────────────
def write_reports(results, cfg, meta):
    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base = os.path.join(REPORT_DIR, f"이카운트대조_{stamp}")

    # CSV
    cols = list(results[0].keys()) if results else []
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(results)

    # 집계
    from collections import Counter
    c1 = Counter(r["①판매/명세서_판정"].split("(")[0] for r in results)
    c2 = Counter(r["②세금계산서_판정"].split("(")[0] for r in results)

    # Markdown
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write(f"# 이카운트 ↔ 관리대장 대조 리포트\n\n")
        f.write(f"- 생성: {datetime.now():%Y-%m-%d %H:%M}\n")
        f.write(f"- 이카운트 연결: {'예(실시간 조회)' if meta.get('online') else '아니오(원장 기준 준비표)'}\n")
        f.write(f"- 대상 정산 건수: {len(results)}\n")
        if meta.get("expiry_days") is not None:
            f.write(f"- 인증키 만료까지: {meta['expiry_days']}일\n")
        if meta.get("note"):
            f.write(f"- 비고: {meta['note']}\n")
        f.write("\n## ① 판매/거래명세서 대조 집계\n\n")
        for k, v in c1.most_common():
            f.write(f"- {k}: {v}건\n")
        f.write("\n## ② 세금계산서 대조 집계\n\n")
        for k, v in c2.most_common():
            f.write(f"- {k}: {v}건\n")
        f.write("\n## 상세(불일치·미등록 우선)\n\n")
        f.write("| 정산ID | 프로젝트NO | 캠프 | 공급가액 | ①판매 | ②세금계산서 |\n")
        f.write("|---|---|---|--:|---|---|\n")
        def rank(r):
            bad = ("불일치", "미등록", "미확인", "미발행")
            return 0 if any(b in r["①판매/명세서_판정"] or b in r["②세금계산서_판정"] for b in bad) else 1
        for r in sorted(results, key=rank):
            f.write(f"| {r['정산ID']} | {r['프로젝트NO']} | {r['캠프명']} | "
                    f"{_money(r['원장_공급가액'])} | {r['①판매/명세서_판정']} | {r['②세금계산서_판정']} |\n")

    # 독립 xlsx(마스터와 무관한 별도 파일)
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "이카운트대조"
        ws.append(cols)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="305496")
        red = PatternFill("solid", fgColor="FFC7CE")
        grn = PatternFill("solid", fgColor="E2EFDA")
        for r in results:
            ws.append([r.get(c) for c in cols])
            row = ws[ws.max_row]
            j1 = cols.index("①판매/명세서_판정"); j2 = cols.index("②세금계산서_판정")
            for jj in (j1, j2):
                val = str(row[jj].value or "")
                row[jj].fill = grn if val.startswith("일치") or "해당없음" in val else red
        ws.freeze_panes = "A2"
        wb.save(base + ".xlsx")
    except Exception as e:
        print("  (xlsx 생성 생략:", e, ")")

    return base


# ───────────────────────── 메인 ─────────────────────────
def main():
    args = set(sys.argv[1:])
    cfg = load_config()
    expiry = days_until_expiry(cfg)
    if expiry is not None and expiry <= cfg["auth"].get("만료경고_남은일수", 30):
        print(f"⚠ 인증키 만료까지 {expiry}일 — 이카운트에서 유효기간 연장 필요")

    # 설정 완성도
    need = [k for k in ("COM_CODE", "USER_ID") if not (cfg["auth"].get(k) or "").strip()]

    if "--selftest" in args:
        print("=== 설정 점검 ===")
        print("인증키:", "설정됨" if cfg["auth"].get("API_CERT_KEY") else "없음")
        print("COM_CODE:", cfg["auth"].get("COM_CODE") or "(비어있음 — 필수)")
        print("USER_ID :", cfg["auth"].get("USER_ID") or "(비어있음 — 필수)")
        print("도메인  :", "테스트(sboapi)" if cfg["auth"].get("IS_TEST") else "실서비스(oapi)")
        if need:
            print("→ 미설정:", ", ".join(need)); return
        try:
            cli = EcountClient(cfg); cli.login()
            print("Zone:", cli.zone, "/ 로그인 성공, SESSION_ID 발급됨")
        except EcountError as e:
            print("로그인 실패:", e)
        return

    # 원장 읽기
    recs = read_ledger(cfg["reconcile"]["master_xlsx"])
    print(f"원장 정산 건수: {len(recs)}")

    ecount = {"_online": False}
    note = ""
    # 1순위: inbox/ 의 이카운트 수동 내보내기 파일 (판매·세금계산서 조회 API는 이카운트가 미제공)
    inbox = None if "--no-inbox" in args else load_inbox(cfg)
    if inbox:
        ecount["sale"] = inbox["sale"]
        ecount["tax_invoice"] = inbox["tax_invoice"]
        ecount["_online"] = True
        note = "이카운트 내보내기 파일(inbox) 기준 대조: " + ", ".join(inbox["_files"])
        print("inbox 로드:", note)
        print(f"  판매 {len(inbox['sale'])}건 / 세금계산서 {len(inbox['tax_invoice'])}건")
    elif "--offline" in args or need:
        note = ("이카운트 미조회 — " +
                ("COM_CODE/USER_ID 미설정" if need else "offline 모드") +
                ". 원장 기준 준비표만 생성.")
        print("i", note)
    else:
        note = ("이카운트 OAPI는 판매·세금계산서 '조회' API를 제공하지 않음(2026-07 확인). "
                "이카운트 화면에서 판매현황/세금계산서 목록을 엑셀로 내려받아 inbox/ 폴더에 넣으면 자동 대조됩니다. "
                "이번 실행은 원장 기준 준비표.")
        print("i", note)

    results = reconcile(recs, ecount, cfg)
    base = write_reports(results, cfg, {"online": ecount["_online"], "expiry_days": expiry, "note": note})
    print("리포트 생성:")
    for ext in (".md", ".csv", ".xlsx"):
        if os.path.exists(base + ext):
            print("  -", base + ext)


if __name__ == "__main__":
    main()
