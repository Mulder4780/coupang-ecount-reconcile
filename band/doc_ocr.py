# -*- coding: utf-8 -*-
"""
doc_ocr.py — 밴드에 올라온 거래명세서·세금계산서 **이미지**를 읽어 관리대장과 대조
================================================================================
밴드 글에는 금액이 글로 안 적히고 명세서/계산서 사진만 올라오는 경우가 많다.
그 사진에서 번호·날짜·금액을 뽑아 관리대장(06 시트)과 맞춰본다.

핵심 원칙
  · OCR은 **Windows 내장 엔진**만 쓴다(band/ocr_win.ps1). 외부 API·업로드 없음 —
    거래명세서·세금계산서는 금융 문서이므로 PC 밖으로 내보내지 않는다.
  · OCR은 반드시 틀린다("UJ2601138"→"U]2601138", "2,000,000"→"2|000,000").
    그래서 라벨:값 인접이 아니라 **패턴 + 오인식 보정 + 금액 정합성(공급가+세액=합계)**
    으로 뽑고, 검증을 통과한 값만 신뢰한다.
  · 기본은 **제안만** 한다. 금액까지 일치해 확실한 건만 --apply로 빈칸에 채운다.

사용
  python band/doc_ocr.py --scan                 # 기본 폴더 스캔 → 리포트
  python band/doc_ocr.py --scan --apply         # 확실한 건은 대장 빈칸까지 입력
  python band/doc_ocr.py --text-only 파일.txt    # OCR 없이 파서만(테스트용)

이미지를 넣는 곳:  ecount/band/docs_inbox/   (밴드에서 저장한 사진을 그대로 복사)
"""
import os, re, sys, csv, glob, json, subprocess
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
INBOX = os.path.join(HERE, "docs_inbox")
PS1 = os.path.join(HERE, "ocr_win.ps1")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
IMG_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")

# ── OCR ────────────────────────────────────────────────────────────────
OCR_CACHE = os.path.join(HERE, "ocr_cache")


def _cache_path(path):
    """같은 사진을 두 번 읽지 않는다. 사진 1,458장 OCR에 25분이 걸려
    daily_run의 600초 제한에 걸렸다(2026-07-26). 파일이 바뀌면 키도 바뀐다."""
    try:
        st = os.stat(path)
        key = f"{os.path.basename(path)}|{st.st_size}|{int(st.st_mtime)}"
    except OSError:
        key = os.path.basename(path)
    import hashlib
    return os.path.join(OCR_CACHE, hashlib.md5(key.encode()).hexdigest() + ".txt")


def ocr_image(path, lang="ko", timeout=120):
    """Windows 내장 OCR → 줄 단위 텍스트. 실패하면 빈 문자열."""
    cp = _cache_path(path)
    if os.path.exists(cp):
        try:
            return open(cp, encoding="utf-8").read()
        except OSError:
            pass
    text = _ocr_run(path, lang, timeout)
    try:
        os.makedirs(OCR_CACHE, exist_ok=True)
        open(cp, "w", encoding="utf-8").write(text)
    except OSError:
        pass
    return text


def _ocr_run(path, lang="ko", timeout=120):
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-File", PS1, "-Path", path, "-Lang", lang],
                           capture_output=True, timeout=timeout)
        if r.returncode != 0:
            return ""
        return r.stdout.decode("utf-8", "replace")
    except Exception:
        return ""


# ── 파싱 (OCR 오인식을 전제로 설계 — 순수 함수라 합성검증 대상) ──────────────
# OCR이 콤마를 |, ., ' 등으로 자주 오인식한다. 단 **줄바꿈은 넣지 않는다** —
# \s를 쓰면 '1,850,000⏎150,000'이 한 덩어리(1850000150000)로 붙어버린다.
_SEP = r"[,|.·'’` \t]"
_AMT_RE = re.compile(r"\d{1,3}(?:" + _SEP + r"\d{3})+")
_DATE_RE = re.compile(r"(20\d{2})[-./년]\s?(\d{1,2})[-./월]\s?(\d{1,2})")
# 프로젝트NO: U 다음 글자(J)가 ], ), I, 1, L 등으로 깨져도 잡는다
_PRJ_RE = re.compile(r"U[^0-9]{0,2}(2\d{6})")
_APPROVE_RE = re.compile(r"(?<!\d)\d{8}-\d{8}-\d{8}(?!\d)")
_SLIP_RE = re.compile(r"(?<![A-Z0-9])[A-Z]{2}-\d{4}-\d{3,5}(?!\d)")
_BIZNO_RE = re.compile(r"(?<!\d)\d{3}-\d{2}-\d{5}(?!\d)")


def _despace(t):
    return re.sub(r"\s+", "", t or "")


def _lines(t):
    """줄 단위로 공백만 제거. 문서 전체를 이어붙이면 옆 칸 숫자와 들러붙어
    'SL-2026-0712' + '123-81-45678' → 'SL-2026-07121' 처럼 잘못 읽힌다."""
    return [_despace(l) for l in (t or "").splitlines() if l.strip()]


def _first(rx, text):
    for l in _lines(text):
        m = rx.search(l)
        if m:
            return m.group()
    return ""


def doc_type(text):
    """'거 래 명 세 서'처럼 자간이 벌어져도 잡히게 공백을 죽이고 판정"""
    d = _despace(text)
    if "세금계산서" in d or "계산서" in d and "명세서" not in d:
        return "세금계산서"
    if "거래명세서" in d or "명세서" in d:
        return "거래명세서"
    return "미상"


def amounts(text):
    """문서에 등장하는 금액 후보(중복 제거·등장 순서 유지)"""
    out, seen = [], set()
    for line in (text or "").splitlines():
        for m in _AMT_RE.finditer(line):
            v = int(re.sub(_SEP, "", m.group()))
            if v >= 1000 and v not in seen:
                seen.add(v)
                out.append(v)
    return out


def infer_amounts(vals, tol=2):
    """공급가액·세액·합계를 **정합성으로** 확정한다.
    세액 = 공급가액의 10%, 합계 = 공급가액 + 세액 을 모두 만족하는 조합만 인정 →
    OCR이 자릿수를 틀리면 조합이 성립하지 않으므로 자연스럽게 걸러진다."""
    s = set(vals)
    best = None
    for sup in vals:
        vat = round(sup * 0.1)
        for t in (vat, vat - 1, vat + 1):
            if t in s and (sup + t) in s:
                if best is None or sup > best[0]:
                    best = (sup, t, sup + t)
    if best:
        return best
    return (None, None, None)


def dates(text):
    return ["%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
            for m in _DATE_RE.finditer(text or "")]


def projects(text):
    """UJ+7자리. 오인식된 두 번째 글자(J→], ), I, 1, L …)를 J로 되돌린다."""
    out, seen = [], set()
    for l in _lines(text):
        for m in _PRJ_RE.finditer(l):
            p = "UJ" + m.group(1)
            if p not in seen:
                seen.add(p)
                out.append(p)
    return out


def parse_doc(text, src=""):
    """OCR 텍스트 → 구조화된 문서 레코드"""
    sup, vat, tot = infer_amounts(amounts(text))
    ds = dates(text)
    prj = projects(text)
    slip = _first(_SLIP_RE, text)
    kind = doc_type(text)
    biz = sorted({m.group() for l in _lines(text) for m in _BIZNO_RE.finditer(l)})
    rec = {
        "파일": os.path.basename(src), "유형": kind,
        "발행일": ds[0] if ds else "",
        "명세서번호": slip if kind == "거래명세서" else "",
        "승인번호": _first(_APPROVE_RE, text),
        "프로젝트NO": prj[0] if prj else "",
        "프로젝트후보": ",".join(prj[1:]),
        "공급가액": sup or "", "세액": vat or "", "합계": tot or "",
        "사업자번호": ",".join(biz),
    }
    # 신뢰도 — 자동 입력을 허용할지 판단하는 기준
    score = sum([bool(rec["발행일"]), bool(rec["프로젝트NO"]), bool(sup),
                 bool(rec["명세서번호"] or rec["승인번호"])])
    rec["신뢰도"] = {4: "높음", 3: "보통"}.get(score, "낮음")
    return rec


# ── 관리대장 대조 ────────────────────────────────────────────────────────
def load_ledger():
    sys.path.insert(0, ROOT)
    from ecount_reconcile import read_ledger, load_config
    return read_ledger(load_config()["reconcile"]["master_xlsx"])


def match(rec, recs):
    """문서 1건 ↔ 대장. 프로젝트NO로 찾고 금액으로 확인한다."""
    prj = rec.get("프로젝트NO")
    if not prj:
        return None, "프로젝트NO를 못 읽음 — 사람이 확인 필요"
    hit = [(sid, r) for sid, r in recs.items() if str(r.get("프로젝트NO") or "").strip() == prj]
    if not hit:
        return None, f"{prj} 가 관리대장에 없음"
    sid, r = hit[0]
    led = r.get("원장_공급가액")
    doc = rec.get("공급가액")
    if doc and led and int(led) != int(doc):
        return sid, f"금액 불일치 — 대장 {int(led):,} / 문서 {int(doc):,}"
    if doc and not led:
        return sid, f"대장 공급가액 비어 있음 → 문서값 {int(doc):,} 제안"
    return sid, "일치"


def proposals(rec, sid, recs):
    """대장 빈칸 중 이 문서로 채울 수 있는 항목"""
    r = recs.get(sid) or {}
    out = {}
    if rec["유형"] == "거래명세서":
        if rec["명세서번호"] and not str(r.get("원장_거래명세서번호") or "").strip():
            out["거래명세서번호"] = rec["명세서번호"]
        if rec["발행일"] and not r.get("원장_거래명세서발행일"):
            out["거래명세서발행일"] = rec["발행일"]
    else:
        if rec["승인번호"] and not str(r.get("원장_세금계산서승인번호") or "").strip():
            out["세금계산서승인번호"] = rec["승인번호"]
        # 06시트의 실제 열 이름은 '세금계산서발행일'이다('…실제발행일'이 아니다).
        # 이름이 틀리면 ledger_writer가 "열 없음"으로 조용히 걸러낸다.
        if rec["발행일"] and not r.get("원장_세금계산서발행일"):
            out["세금계산서발행일"] = rec["발행일"]
    if rec["공급가액"] and not r.get("원장_공급가액"):
        out["공급가액"] = rec["공급가액"]
    return out


# ── 실행 ────────────────────────────────────────────────────────────────
def photo_dirs(folder=None):
    """사진을 찾을 곳. 원본은 서버('0. 원본 자료/4. 밴드 원본/문서사진')에 두고,
    PC 로컬 inbox 도 계속 본다 — 새로 받은 사진을 급히 떨어뜨리는 자리이고,
    서버가 끊겨도 그건 읽을 수 있어야 한다(사용자 지시 2026-07-28: 원본은 서버로)."""
    if folder:
        return [folder]
    try:
        sys.path.insert(0, ROOT)
        from source_dirs import doc_photo_dirs
        dirs = doc_photo_dirs()
    except Exception:
        dirs = []
    return dirs or [INBOX]


def scan(folder=None, apply=False):
    dirs = photo_dirs(folder)
    imgs, seen = [], set()
    for d in dirs:
        for p in sorted(glob.glob(os.path.join(d, "**", "*"), recursive=True)):
            if not p.lower().endswith(IMG_EXT):
                continue
            name = os.path.basename(p)
            if name in seen:          # 서버·로컬 양쪽에 같은 파일이 있으면 한 번만 본다
                continue
            seen.add(name)
            imgs.append(p)
    if not imgs:
        os.makedirs(INBOX, exist_ok=True)
        print("이미지 없음 — 아래 중 한 곳에 밴드에서 저장한 명세서·계산서 사진을 넣으세요")
        for d in dirs:
            print("   ", d)
        return []
    if len(dirs) > 1:
        print("사진 %d장 (%d곳에서: %s)" % (len(imgs), len(dirs), " · ".join(dirs)))
    try:
        recs = load_ledger()
    except Exception as e:
        print(f"관리대장을 읽지 못함: {e}")
        recs = {}
    rows = []
    for p in imgs:
        rec = parse_doc(ocr_image(p), p)
        sid, verdict = match(rec, recs) if recs else (None, "대장 미로드")
        rec["정산ID"] = sid or ""
        rec["판정"] = verdict
        rec["제안"] = json.dumps(proposals(rec, sid, recs), ensure_ascii=False) if sid else ""
        rows.append(rec)

    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    csv_p = os.path.join(REPORT_DIR, f"밴드문서OCR_{stamp}.csv")
    cols = ["파일", "유형", "발행일", "프로젝트NO", "정산ID", "명세서번호", "승인번호",
            "공급가액", "세액", "합계", "신뢰도", "판정", "제안", "프로젝트후보", "사업자번호"]
    with open(csv_p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})

    ok = sum(1 for r in rows if r["판정"] == "일치")
    bad = [r for r in rows if "불일치" in r["판정"]]
    print(f"밴드 문서 이미지 {len(rows)}건 — 일치 {ok} · 불일치 {len(bad)} · "
          f"확인필요 {len(rows)-ok-len(bad)}  → {os.path.basename(csv_p)}")
    for r in bad:
        print(f"  ★ {r['파일']} {r['프로젝트NO']} {r['판정']}")

    if apply:
        applied = apply_proposals(rows)
        print(f"대장 빈칸 입력: {applied}건")
    return rows


VTYPE = {"거래명세서발행일": "date", "세금계산서발행일": "date", "공급가액": "number"}


def build_updates(rows):
    """신뢰도 '높음' + 판정 '일치'인 건만 입력 후보로 만든다(순수 함수 — 합성검증 대상).
    실제 기록은 ledger_writer가 **빈칸일 때만** 수행하므로 기존 값은 덮이지 않는다."""
    todo = []
    for r in rows:
        if r.get("신뢰도") != "높음" or r.get("판정") != "일치" or not r.get("정산ID"):
            continue
        prop = r.get("제안")
        prop = json.loads(prop) if isinstance(prop, str) else (prop or {})
        for k, v in prop.items():
            todo.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID",
                         "key": r["정산ID"], "col": k, "value": v,
                         "vtype": VTYPE.get(k, "text"),
                         "근거": f"밴드 문서 이미지 OCR ({r['파일']})"})
    return todo


def apply_proposals(rows):
    sys.path.insert(0, ROOT)
    from ledger_writer import queue_add
    todo = build_updates(rows)
    return queue_add(todo) if todo else 0


if __name__ == "__main__":
    if "--text-only" in sys.argv:
        p = sys.argv[sys.argv.index("--text-only") + 1]
        print(json.dumps(parse_doc(open(p, encoding="utf-8").read(), p),
                         ensure_ascii=False, indent=1))
    else:
        scan(apply="--apply" in sys.argv)
