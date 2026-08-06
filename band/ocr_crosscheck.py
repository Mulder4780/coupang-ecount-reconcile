# -*- coding: utf-8 -*-
"""
ocr_crosscheck.py — 문서 사진을 **두 번 읽어 맞춰 보고**, 원장과 **항목별로** 대조
==============================================================================
2026-08-06 지시: "문서 스캔해서 텍스트 확인하는 최고의 무료 도구 확인 해서
비교 대조 할 수 있는 알고리즘 추가해"

왜 만들었나
  · OCR 은 반드시 틀린다. **한 엔진의 답만 보면 틀린 줄을 모른다** — 조용히 틀린 값이
    원장에 들어가는 것이 이 프로젝트에서 제일 비싼 사고다.
  · 그래서 서로 다른 엔진에게 같은 사진을 읽히고 **값이 겹칠 때만** 믿는다.
    겹치지 않으면(충돌) 자동입력을 막고 사람에게 넘긴다.
  · 그리고 문서 ↔ 원장을 **항목별로** 대조한다. 기존 `doc_ocr.match()` 는
    프로젝트NO 로 찾아 공급가액 하나만 봤다 — 발행일·문서번호·세액·합계가 틀려도 통과했다.

★ 두 번 읽는 것은 **필요한 것만**이다 (사진 1,816장을 매일 두 번 읽을 수는 없다)
  `needs_second_opinion()` 이 재검 대상을 고른다: 금액 정합성이 깨졌거나, 핵심 항목을
  못 읽었거나, **원장에 값을 쓰려는 순간**. 나머지는 1엔진 결과를 그대로 쓴다.

엔진 — 전부 무료 · 전부 로컬 · 업로드 없음(금융 문서는 PC 밖으로 나가지 않는다)
  paddle     PaddleOCR 한국어 (기본)   무료권에서 한중일·표 인식률 최상위. 이미 설치됨.
  windows    Windows.Media.Ocr        Win10+ 내장이라 설치가 필요 없다 → 둘째 의견의 기본값.
  tesseract  Tesseract 5 (kor)        있으면 쓴다. 한글 정확도는 낮지만 **다른 방식으로
                                      틀려서** 교차검증 표로는 값이 있다.
  (surya 등 추가는 band/OCR_ENGINES.md 참고 — ENGINES 에 한 줄 추가하면 붙는다)
  설치는 하지 않는다. **있는 것만 골라 쓴다.**

사용
  python band/ocr_crosscheck.py --status          # 이 PC 에서 쓸 수 있는 엔진
  python band/ocr_crosscheck.py --scan            # 재검 대상만 교차검증 → 리포트
  python band/ocr_crosscheck.py --scan --all      # 전부 두 번 읽기(느리다)
  python band/ocr_crosscheck.py --scan --apply    # '합치' + 원장 빈칸만 입력 큐로
"""
import os, re, sys, csv, glob, json, shutil, hashlib, subprocess, tempfile
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
for _p in (HERE, ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import doc_ocr  # 같은 폴더 — 파서·엔진 호출·사진 경로를 그대로 재사용한다

REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
XCACHE = os.path.join(doc_ocr.OCR_CACHE, "cross")


# ── 엔진 ────────────────────────────────────────────────────────────────
def _cache_path(engine, path):
    """엔진마다 따로 저장한다. 같은 사진을 같은 엔진으로 두 번 읽지 않는다."""
    try:
        st = os.stat(path)
        key = "%s|%s|%s|%s" % (engine, os.path.abspath(path), st.st_size, int(st.st_mtime))
    except OSError:
        key = "%s|%s" % (engine, os.path.basename(path))
    return os.path.join(XCACHE, hashlib.md5(key.encode()).hexdigest() + ".txt")


def _run_paddle(paths, timeout):
    return doc_ocr._paddle_batch(paths, timeout)


def _run_windows(paths, timeout):
    return {p: doc_ocr._ocr_run(p, "ko", timeout) for p in paths}


def _run_tesseract(paths, timeout):
    exe = shutil.which("tesseract")
    if not exe:
        return {}
    out = {}
    for p in paths:
        try:
            r = subprocess.run([exe, p, "stdout", "-l", "kor+eng", "--psm", "6"],
                               capture_output=True, timeout=timeout)
            out[p] = r.stdout.decode("utf-8", "replace") if r.returncode == 0 else ""
        except Exception:
            out[p] = ""
    return out


ENGINES = {
    # 이름: (설명, 사용 가능한가, 실행 함수, 등급) — 등급은 리포트 표시용이다.
    "paddle": ("PaddleOCR 한국어(로컬)",
               lambda: os.path.isfile(doc_ocr.PADDLE_PY) and os.path.isfile(doc_ocr.PADDLE_WORKER),
               _run_paddle, "정밀"),
    "windows": ("Windows 내장 OCR",
                lambda: os.path.isfile(doc_ocr.PS1) and os.name == "nt",
                _run_windows, "보조"),
    "tesseract": ("Tesseract 5 (kor)",
                  lambda: bool(shutil.which("tesseract")),
                  _run_tesseract, "보조"),
}
ORDER = ["paddle", "windows", "tesseract"]


def engine_status():
    """이 PC 에서 무엇을 쓸 수 있나 — 설치는 하지 않고 확인만 한다."""
    out = []
    for name in ORDER:
        desc, ready, _fn, grade = ENGINES[name]
        try:
            ok = bool(ready())
        except Exception:
            ok = False
        out.append({"엔진": name, "설명": desc, "등급": grade, "사용가능": ok})
    return out


def available(want=None):
    names = [e["엔진"] for e in engine_status() if e["사용가능"]]
    if want:
        names = [n for n in names if n in want]
    return names


def read_texts(engine, paths, timeout=120):
    """엔진 하나로 여러 장 → {경로: 텍스트}. 캐시 먼저, 없는 것만 실제로 읽는다.

    paddle 은 doc_ocr 가 이미 1,800장분을 캐시해 두었다 — 그 자리를 그대로 읽는다.
    다시 읽으면 사진 한 장에 28초씩 걸려 하루가 간다.
    """
    out, pending = {}, []
    for p in paths:
        if engine == "paddle":
            t = doc_ocr._read_cache(p)
            if t is not None:
                out[p] = t
                continue
        cp = _cache_path(engine, p)
        if os.path.exists(cp):
            try:
                out[p] = open(cp, encoding="utf-8").read()
                continue
            except OSError:
                pass
        pending.append(p)
    if pending:
        fn = ENGINES[engine][2]
        try:
            got = fn(pending, timeout) or {}
        except Exception:
            got = {}
        os.makedirs(XCACHE, exist_ok=True)
        for p in pending:
            t = got.get(p, "") or ""
            out[p] = t
            try:
                open(_cache_path(engine, p), "w", encoding="utf-8").write(t)
            except OSError:
                pass
    return out


# ── 교차검증 (순수 함수 — 합성검증 대상) ──────────────────────────────────
VOTE_FIELDS = ("유형", "발행일", "명세서번호", "승인번호", "프로젝트NO",
               "공급가액", "세액", "합계", "사업자번호")
CORE_FIELDS = ("발행일", "프로젝트NO", "공급가액")


def vote(values):
    """엔진들이 낸 같은 항목의 값 → (확정값, 판정)

    · 두 엔진 이상이 **같은 값** → '합치'  (믿는다. 자동입력 허용)
    · 값을 낸 엔진이 **하나뿐**  → '단독'  (제안만. 자동입력 금지)
    · 서로 **다른 값**           → '충돌'  (사람에게. 값을 정하지 않는다)
    · 아무도 못 읽음             → '없음'
    금액·날짜는 근사 비교를 하지 않는다 — 한 자리만 틀려도 다른 값이다.
    """
    vals = [v for v in values if v not in (None, "", 0)]
    if not vals:
        return "", "없음"
    tally = {}
    for v in vals:
        k = str(v).strip()
        tally[k] = tally.get(k, 0) + 1
    top = max(tally.values())
    winners = sorted(k for k, c in tally.items() if c == top)
    if len(winners) > 1:
        return "", "충돌"
    if top >= 2:
        return winners[0], "합치"
    return (winners[0], "단독") if len(tally) == 1 else ("", "충돌")


def _amt_ok(rec, tol=1):
    """공급가액 + 세액 = 합계 인가. 셋 중 하나라도 비면 판정하지 않는다(True)."""
    try:
        s, v, t = int(rec.get("공급가액") or 0), int(rec.get("세액") or 0), int(rec.get("합계") or 0)
    except (TypeError, ValueError):
        return True
    if not (s and v and t):
        return True
    return abs(s + v - t) <= tol


def drop_dependent(texts):
    """{엔진: 원문} 에서 **원문이 똑같은** 엔진은 하나로 친다 (순수 함수).

    doc_ocr 는 paddle 이 실패하면 Windows 결과를 paddle 자리에 넣어 캐시한다.
    그 자리를 둘로 세면 "두 엔진이 합치했다"는 **거짓 근거**가 만들어진다 —
    실은 같은 답을 두 번 본 것뿐이다. 빈 텍스트도 엔진으로 세지 않는다.
    """
    seen, out = set(), {}
    for e in sorted(texts):
        key = re.sub(r"\s+", "", texts[e] or "")
        if not key or key in seen:
            continue
        seen.add(key)
        out[e] = texts[e]
    return out


def merge_records(by_engine):
    """{엔진: parse_doc 결과} → 합쳐진 1건 + 항목별 교차판정

    반환의 `교차` 는 {항목: 합치|단독|충돌|없음} 이고, `신뢰도` 는 **교차 결과로**
    다시 매긴다 — 한 엔진 안에서만 그럴듯한 값은 '높음'이 될 수 없다.
    """
    live = {e: (r or {}) for e, r in by_engine.items() if r}
    merged, cross = {}, {}
    for f in VOTE_FIELDS:
        v, how = vote([live[e].get(f) for e in live])
        merged[f], cross[f] = v, how
    merged["엔진"] = ",".join(sorted(live))
    merged["교차"] = cross

    # 합쳐 놓고 보니 금액이 안 맞으면, 겹쳤더라도 금액은 믿지 않는다.
    if not _amt_ok(merged):
        for f in ("공급가액", "세액", "합계"):
            if cross.get(f) in ("합치", "단독"):
                cross[f] = "충돌"
        merged["금액정합"] = "깨짐"
    else:
        merged["금액정합"] = "정상"

    n_no = sum(1 for f in VOTE_FIELDS if cross[f] == "충돌")
    core_ok = all(cross[f] == "합치" for f in CORE_FIELDS)
    has_no = bool(merged.get("명세서번호") or merged.get("승인번호"))
    if len(live) < 2:
        merged["신뢰도"] = "낮음(1엔진)"
    elif n_no:
        merged["신뢰도"] = "낮음"
    elif core_ok and has_no:
        merged["신뢰도"] = "높음"
    elif core_ok:
        merged["신뢰도"] = "보통"
    else:
        merged["신뢰도"] = "낮음"
    merged["교차요약"] = "합치 %d · 단독 %d · 충돌 %d" % (
        sum(1 for f in VOTE_FIELDS if cross[f] == "합치"),
        sum(1 for f in VOTE_FIELDS if cross[f] == "단독"),
        n_no)
    return merged


def needs_second_opinion(rec, for_write=False):
    """한 엔진의 답만으로 끝내도 되나. 되면 '' , 아니면 **재검 이유**.

    전량을 두 번 읽지 않기 위한 문지기다. 원장에 값을 쓰려는 건은 무조건 재검한다 —
    쓰는 순간이 되돌리기 제일 비싼 지점이기 때문이다.
    """
    if not rec:
        return "읽은 값 없음"
    if not _amt_ok(rec):
        return "금액 정합성 깨짐"
    if not rec.get("프로젝트NO"):
        return "프로젝트NO 못 읽음"
    if not rec.get("공급가액"):
        return "공급가액 못 읽음"
    if not rec.get("발행일"):
        return "발행일 못 읽음"
    if not (rec.get("명세서번호") or rec.get("승인번호")):
        return "문서번호 못 읽음"
    if str(rec.get("신뢰도") or "") != "높음":
        return "신뢰도 %s" % (rec.get("신뢰도") or "?")
    if for_write:
        return "원장 입력 후보 — 두 번 읽어 확인"
    return ""


# ── 문서 ↔ 원장 항목별 대조 (순수 함수) ────────────────────────────────────
# (문서항목, 원장 열 후보, 형)  — 원장 열은 앞에서부터 값이 있는 것을 쓴다
LEDGER_MAP = {
    "거래명세서": [("발행일", ("원장_거래명세서발행일",), "date"),
                   ("명세서번호", ("원장_거래명세서번호",), "text"),
                   ("공급가액", ("원장_공급가액",), "money"),
                   ("세액", ("원장_부가세",), "money"),
                   ("합계", ("원장_거래명세서합계", "원장_합계"), "money")],
    "세금계산서": [("발행일", ("원장_세금계산서실제발행일", "원장_세금계산서발행일"), "date"),
                   ("승인번호", ("원장_세금계산서승인번호",), "text"),
                   ("공급가액", ("원장_공급가액",), "money"),
                   ("세액", ("원장_부가세",), "money"),
                   ("합계", ("원장_세금계산서합계", "원장_합계"), "money")],
}
# 원장에 실제로 써도 되는 항목만 (세액·합계는 수식·집계라 건드리지 않는다)
WRITE_COL = {
    "거래명세서": {"발행일": "거래명세서발행일", "명세서번호": "거래명세서번호",
                   "공급가액": "공급가액"},
    "세금계산서": {"발행일": "세금계산서발행일", "승인번호": "세금계산서승인번호",
                   "공급가액": "공급가액"},
}
VTYPE = {"거래명세서발행일": "date", "세금계산서발행일": "date", "공급가액": "number"}


def _txt(v):
    return re.sub(r"[\s\-·.]", "", str(v or "")).upper()


def _date(v):
    s = str(v or "").strip()[:10].replace("/", "-").replace(".", "-")
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def _money(v):
    try:
        return int(round(float(re.sub(r"[,\s원]", "", str(v)))))
    except (TypeError, ValueError):
        return None


def compare_ledger(rec, led, tol=1):
    """문서 1건 ↔ 원장 1행을 **항목마다** 견준다 → [{항목,문서,원장,판정}]"""
    kind = rec.get("유형") if rec.get("유형") in LEDGER_MAP else "거래명세서"
    rows = []
    for label, keys, typ in LEDGER_MAP[kind]:
        lv = ""
        for k in keys:
            if (led or {}).get(k) not in (None, "", 0):
                lv = led[k]
                break
        dv = rec.get(label)
        if typ == "money":
            d, l = _money(dv), _money(lv)
            ds = "" if d is None else str(d)
            ls = "" if l is None else str(l)
            same = d is not None and l is not None and abs(d - l) <= tol
        elif typ == "date":
            ds, ls = _date(dv), _date(lv)
            same = bool(ds) and ds == ls
        else:
            ds, ls = str(dv or "").strip(), str(lv or "").strip()
            same = bool(ds) and _txt(ds) == _txt(ls)
        if not ds and not ls:
            verdict = "양쪽 빈칸"
        elif not ds:
            verdict = "문서에서 못 읽음"
        elif not ls:
            verdict = "원장 빈칸"
        elif same:
            verdict = "일치"
        else:
            verdict = "불일치"
        rows.append({"항목": label, "문서": ds, "원장": ls, "판정": verdict})
    return rows


def ledger_verdict(rows):
    """항목별 대조 → 한 줄 판정. 불일치가 하나라도 있으면 그것이 결론이다."""
    bad = [r["항목"] for r in rows if r["판정"] == "불일치"]
    if bad:
        return "불일치(%s)" % "·".join(bad)
    fill = [r["항목"] for r in rows if r["판정"] == "원장 빈칸"]
    ok = sum(1 for r in rows if r["판정"] == "일치")
    if fill:
        return "원장 빈칸 %d항목%s" % (len(fill), (" — " + "·".join(fill)))
    if ok:
        return "일치(%d항목)" % ok
    return "대조할 값 없음"


def writable_now(rec, rows, sid):
    """지금 이 건이 **원장 빈칸을 채우려 하는가** (순수 함수).
    채우려 한다면 반드시 두 번 읽는다 — 쓰는 순간이 되돌리기 제일 비싼 지점이다."""
    if not sid or not rows or any(r["판정"] == "불일치" for r in rows):
        return False
    kind = rec.get("유형") if rec.get("유형") in WRITE_COL else None
    if not kind:
        return False
    return any(r["판정"] == "원장 빈칸" and WRITE_COL[kind].get(r["항목"]) and r["문서"]
               for r in rows)


def recheck_reason(rec, led_rows, can_write):
    """이 사진을 두 번 읽어야 하나 → (우선순위, 이유). 아니면 (None, '').
    숫자가 작을수록 급하다. 0 은 예산과 상관없이 무조건 재검한다."""
    if can_write:
        return 0, "원장 입력 후보"
    if any(r["판정"] == "불일치" for r in (led_rows or [])):
        return 1, "원장과 불일치"
    if not _amt_ok(rec):
        return 2, "금액 정합성 깨짐"
    why = needs_second_opinion(rec)
    return (3, why) if why else (None, "")


def recheck_plan(cands, budget=0):
    """[(키, 우선순위)] → (이번에 다시 읽을 키, 미룬 수)  — 순수 함수

    사진이 1,816장이고 둘째 엔진이 한 장에 5초다. 전량 재검은 하루가 간다.
    그래서 급한 것부터 예산만큼만 읽는다. 결과는 엔진별로 캐시되므로
    남은 것은 다음 회차에 이어서 읽힌다 — 며칠이면 밀린 것이 빠진다.
    **미룬 수는 반드시 알린다.** 조용히 잘라내면 '다 봤다'로 읽힌다.
    """
    must = [k for k, pr in cands if pr == 0]
    rest = [k for k, pr in sorted((c for c in cands if c[1] != 0), key=lambda c: c[1])]
    if budget and len(rest) > budget:
        return must + rest[:budget], len(rest) - budget
    return must + rest, 0


def build_updates(items):
    """원장에 넣을 것을 고른다 (순수 함수 — 합성검증 대상)

    통과 조건을 전부 만족해야 한다:
      ① 교차검증 '합치'  ② 원장이 빈칸  ③ 다른 항목에 '불일치'가 없음
      ④ 정산ID 가 잡혔음  ⑤ 쓸 수 있는 열(WRITE_COL)일 것
    ②는 ledger_writer 도 다시 확인하지만 여기서도 막는다 — 덮어쓰기는 복구가 안 된다.
    """
    todo = []
    for it in items:
        rec, sid = it.get("문서") or {}, it.get("정산ID")
        rows = it.get("대조") or []
        if not sid or any(r["판정"] == "불일치" for r in rows):
            continue
        kind = rec.get("유형") if rec.get("유형") in WRITE_COL else None
        if not kind:
            continue
        cross = rec.get("교차") or {}
        for r in rows:
            if r["판정"] != "원장 빈칸":
                continue
            col = WRITE_COL[kind].get(r["항목"])
            if not col or cross.get(r["항목"]) != "합치":
                continue
            todo.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": sid,
                         "col": col, "value": r["문서"], "vtype": VTYPE.get(col, "text"),
                         "근거": "문서사진 OCR 교차검증 합치 (%s / %s)" % (
                             rec.get("파일", ""), rec.get("엔진", ""))})
    return todo


# ── 실행 ────────────────────────────────────────────────────────────────
def _images(folder=None):
    imgs, seen = [], set()
    for d in doc_ocr.photo_dirs(folder):
        for p in sorted(glob.glob(os.path.join(d, "**", "*"), recursive=True)):
            if not p.lower().endswith(doc_ocr.IMG_EXT):
                continue
            n = os.path.basename(p)
            if n in seen:
                continue
            seen.add(n)
            imgs.append(p)
    return imgs


def _find_row(prj, recs):
    for sid, r in (recs or {}).items():
        if str(r.get("프로젝트NO") or "").strip() == str(prj or "").strip() and prj:
            return sid, r
    return None, {}


BUDGET = 300   # 한 회차에 둘째 엔진으로 다시 읽을 최대 장수(약 25분). 0 이면 무제한


def crosscheck(folder=None, engines=None, all_docs=False, apply=False, limit=0,
               timeout=120, budget=BUDGET):
    st = engine_status()
    usable = [e["엔진"] for e in st if e["사용가능"]]
    if engines:
        usable = [e for e in usable if e in engines]
    if not usable:
        print("쓸 수 있는 OCR 엔진이 없습니다 — python band/ocr_crosscheck.py --status")
        return []
    primary = usable[0]
    others = usable[1:]

    imgs = _images(folder)
    if limit:
        imgs = imgs[:limit]
    if not imgs:
        print("이미지 없음 — 문서 사진 폴더가 비어 있습니다")
        return []
    try:
        recs = doc_ocr.load_ledger()
    except Exception as e:
        print("관리대장을 읽지 못함: %s" % e)
        recs = {}

    # 1) 기본 엔진으로 전량 (캐시가 있으면 즉시)
    base_txt = read_texts(primary, imgs, timeout)
    first = {p: doc_ocr.parse_doc(base_txt.get(p, ""), p) for p in imgs}

    # 2) 재검 대상을 **급한 것부터** 고른다 — "두 번 읽되 다 읽지는 않는다"의 핵심
    cands, why = [], {}
    for p in imgs:
        if all_docs:
            cands.append((p, 3))
            why[p] = "전량 재검(--all)"
            continue
        sid0, led0 = _find_row(first[p].get("프로젝트NO"), recs)
        rows0 = compare_ledger(first[p], led0) if sid0 else []
        pr, reason = recheck_reason(first[p], rows0, writable_now(first[p], rows0, sid0))
        if pr is not None:
            cands.append((p, pr))
            why[p] = reason
    recheck, deferred = recheck_plan(cands, 0 if all_docs else budget)

    # 3) 둘째·셋째 의견
    second = {}
    if others and recheck:
        for e in others:
            t = read_texts(e, recheck, timeout)
            for p in recheck:
                second.setdefault(p, {})[e] = (t.get(p, ""), doc_ocr.parse_doc(t.get(p, ""), p))

    # 4) 합치고 · 원장과 항목별로 대조
    items = []
    for p in imgs:
        texts = {primary: base_txt.get(p, "")}
        for e, (raw, _r) in (second.get(p) or {}).items():
            texts[e] = raw
        # 원문이 똑같은 엔진은 하나로 친다 — 같은 답을 두 번 본 것을 합치라 부르지 않는다
        live = drop_dependent(texts)
        by = {}
        if primary in live:
            by[primary] = first[p]
        for e in live:
            if e != primary:
                by[e] = second[p][e][1]
        by = by or {primary: first[p]}
        merged = merge_records(by) if len(by) > 1 else dict(
            first[p], **{"엔진": primary,
                         "교차": {f: ("단독" if first[p].get(f) else "없음") for f in VOTE_FIELDS},
                         "금액정합": "정상" if _amt_ok(first[p]) else "깨짐",
                         "신뢰도": "낮음(1엔진)",
                         "교차요약": "1엔진" + ("(재검 미룸)" if p in why and p not in recheck
                                                else "(재검 불필요)" if p not in why else "(둘째 엔진이 못 읽음)")})
        merged["파일"] = os.path.basename(p)
        sid, led = _find_row(merged.get("프로젝트NO"), recs)
        rows = compare_ledger(merged, led) if sid else []
        items.append({"경로": p, "문서": merged, "정산ID": sid or "",
                      "재검사유": why.get(p, ""), "대조": rows,
                      "판정": ledger_verdict(rows) if sid else (
                          "프로젝트NO 못 읽음" if not merged.get("프로젝트NO")
                          else "%s 가 관리대장에 없음" % merged["프로젝트NO"])})

    _write_reports(items, st, primary, others, len(recheck))
    if deferred:
        print("※ 재검 %d장을 다음 회차로 미뤘습니다(한 회차 %d장). 급한 것부터 읽었고 "
              "읽은 결과는 캐시되므로 며칠이면 밀린 것이 빠집니다." % (deferred, budget))

    n_conf = sum(1 for it in items if "불일치" in it["판정"])
    n_cross = sum(1 for it in items if "충돌" in (it["문서"].get("교차요약") or "")
                  and not (it["문서"].get("교차요약") or "").endswith("충돌 0"))
    todo = build_updates(items)
    print("문서 %d장 — 재검 %d장(%s) · 원장 불일치 %d · OCR 충돌 %d · 입력후보 %d"
          % (len(items), len(recheck), "+".join(others) or "둘째 엔진 없음",
             n_conf, n_cross, len(todo)))
    for it in items:
        if "불일치" in it["판정"]:
            print("  ★ %s %s %s" % (it["문서"].get("파일"), it["문서"].get("프로젝트NO"), it["판정"]))

    if apply:
        if todo:
            from ledger_writer import queue_add
            print("원장 빈칸 입력 큐: %d건" % queue_add(todo))
        else:
            print("원장 빈칸 입력 큐: 0건 (교차검증을 통과한 빈칸 없음)")
    return items


def _write_reports(items, st, primary, others, n_recheck):
    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    doc_p = os.path.join(REPORT_DIR, "문서OCR교차검증_%s.csv" % stamp)
    cols = ["파일", "유형", "프로젝트NO", "정산ID", "발행일", "명세서번호", "승인번호",
            "공급가액", "세액", "합계", "엔진", "교차요약", "금액정합", "신뢰도",
            "재검사유", "판정"]
    with open(doc_p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for it in items:
            d = it["문서"]
            w.writerow(dict({c: d.get(c, "") for c in cols},
                            **{"정산ID": it["정산ID"], "재검사유": it["재검사유"],
                               "판정": it["판정"]}))
    diff_p = os.path.join(REPORT_DIR, "문서OCR항목대조_%s.csv" % stamp)
    with open(diff_p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["파일", "정산ID", "항목", "문서", "원장",
                                          "판정", "교차판정"])
        w.writeheader()
        for it in items:
            cross = it["문서"].get("교차") or {}
            for r in it["대조"]:
                if r["판정"] in ("양쪽 빈칸",):
                    continue
                w.writerow({"파일": it["문서"].get("파일"), "정산ID": it["정산ID"],
                            "항목": r["항목"], "문서": r["문서"], "원장": r["원장"],
                            "판정": r["판정"], "교차판정": cross.get(r["항목"], "")})
    try:
        json.dump({"시각": datetime.now().isoformat(timespec="seconds"),
                   "엔진": st, "기본": primary, "둘째": others,
                   "문서수": len(items), "재검": n_recheck,
                   "불일치": sum(1 for it in items if "불일치" in it["판정"])},
                  open(os.path.join(REPORT_DIR, "문서OCR교차검증_최근.json"), "w",
                       encoding="utf-8"), ensure_ascii=False, indent=1)
    except OSError:
        pass


def print_status():
    print("문서 스캔 엔진 — 이 PC 에서 쓸 수 있는 것 (전부 무료·로컬·업로드 없음)")
    for e in engine_status():
        print("  %s %-9s %-24s %s" % ("O" if e["사용가능"] else "-",
                                      e["엔진"], e["설명"], e["등급"]))
    use = available()
    if len(use) >= 2:
        print("\n교차검증 가능: %s → 값이 겹칠 때만 원장에 넣습니다." % " + ".join(use))
    elif use:
        print("\n엔진이 %s 하나뿐입니다 — 교차검증이 안 되므로 자동입력을 하지 않습니다." % use[0])
        print("둘째 의견을 붙이는 방법은 band/OCR_ENGINES.md 를 보세요.")
    else:
        print("\n쓸 수 있는 엔진이 없습니다 — band/OCR_ENGINES.md 를 보세요.")


if __name__ == "__main__":
    a = sys.argv[1:]
    if "--status" in a or not a:
        print_status()
    else:
        eng = None
        if "--engines" in a:
            eng = [s.strip() for s in a[a.index("--engines") + 1].split(",") if s.strip()]
        lim = int(a[a.index("--limit") + 1]) if "--limit" in a else 0
        bud = int(a[a.index("--budget") + 1]) if "--budget" in a else BUDGET
        crosscheck(engines=eng, all_docs="--all" in a, apply="--apply" in a,
                   limit=lim, budget=bud)
