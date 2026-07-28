# -*- coding: utf-8 -*-
"""
app_server.py — Coupang Service Operations System 앱 서버 (반응형 웹앱 백엔드)
============================================================
PC에서 실행하면 같은 와이파이의 휴대폰·다른 PC가 브라우저로 접속하는 ERP형 앱.
표준 라이브러리만 사용(설치 0). 데이터는 전부 사내 PC에 남는다(클라우드 전송 없음).

  실행:  python webapp/app_server.py            # 실서비스 (첫 실행 시 PIN 자동 생성)
         python webapp/app_server.py --demo     # 합성데이터 데모 (PIN 0000)
  접속:  PC      → http://localhost:8899
         휴대폰  → http://<PC IP>:8899   (같은 와이파이, 방화벽 허용 필요)

보안: 4자리 PIN(첫 요청 시 입력, 기기에 저장). 사내 LAN 전용 설계 — 외부 인터넷 개방 금지.
"""
import sys, os, re, json, glob, time, threading, random, subprocess, hashlib
from collections import deque
from datetime import datetime, date, timedelta
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)
PY = sys.executable
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}

DEMO = "--demo" in sys.argv
PORT = int(sys.argv[sys.argv.index("--port") + 1]) if "--port" in sys.argv else 8899
WEBCFG = os.path.join(ROOT, "config", "webapp.json")

# 폰 홈 화면 아이콘이 가리켜야 할 **바뀌지 않는 주소**.
# 터널 주소(trycloudflare)는 띄울 때마다 새로 받으므로 아이콘에 박으면 안 된다.
FIXED_ENTRY = "https://mulder4780.github.io/coupang-ecount-reconcile/"


def load_pin():
    if DEMO:
        return "0000"
    try:
        return json.load(open(WEBCFG, encoding="utf-8"))["pin"]
    except Exception:
        pin = str(random.SystemRandom().randint(1000, 9999))
        os.makedirs(os.path.dirname(WEBCFG), exist_ok=True)
        json.dump({"pin": pin, "port": PORT}, open(WEBCFG, "w", encoding="utf-8"))
        return pin


PIN = load_pin()

# 브루트포스 차단: IP당 로그인 5회 실패 → 10분 잠금 (외부 터널 공개 대비)
_fails = {}
def _locked(ip):
    c, until = _fails.get(ip, (0, 0))
    return c >= 5 and time.time() < until
def _fail(ip):
    c, _ = _fails.get(ip, (0, 0))
    _fails[ip] = (c + 1, time.time() + 600)
def _ok_login(ip):
    _fails.pop(ip, None)

# ───────────────────────── 작업 러너 ─────────────────────────
TASKS = {
    "daily":         ("전체 대조 실행", [os.path.join(ROOT, "daily_run.py")]),
    "synthetic":     ("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")]),
    "writer_prev":   ("자동입력 미리보기", [os.path.join(ROOT, "ledger_writer.py")]),
    "writer_apply":  ("자동입력 반영", [os.path.join(ROOT, "ledger_writer.py"), "--apply"]),
    "upload_dry":    ("전표 전송대기 확인", [os.path.join(ROOT, "ecount_upload.py")]),
    "upload_post":   ("전표 실전송", [os.path.join(ROOT, "ecount_upload.py"), "--post"]),
    "kakao":         ("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]),
    "erp_ledger":    ("ERP원장 대조", [os.path.join(ROOT, "erp_ledger_check.py")]),
    "po":            ("쿠팡 PO 대조", [os.path.join(ROOT, "po_reconcile.py")]),
    "erp_docs":      ("ERP 매출서류 대조", [os.path.join(ROOT, "erp_docs_check.py")]),
    "band_ingest":   ("밴드 수집분 반영(24시트+백필)",
                      [os.path.join(ROOT, "band", "ingest.py"), "--sheet", "--backfill"]),
    "band_docs":     ("밴드 문서 이미지 대조", [os.path.join(ROOT, "band", "doc_ocr.py"), "--scan"]),
    "band_docs_apply": ("밴드 문서 → 대장 입력", [os.path.join(ROOT, "band", "doc_ocr.py"), "--scan", "--apply"]),
}
runner = {"busy": False, "task": "", "log": deque(maxlen=3000), "done_at": None}
_rlock = threading.Lock()


_codes_cache = {"t": 0, "v": None}


def get_codes():
    """드롭다운 선택지를 **10_코드관리 시트에서** 읽어 온다.

    화면에 목록을 박아 두면 사람이 바뀔 때마다 코드를 고쳐야 하고, 결국 시트와 어긋난다.
    시트가 진실이므로 거기서 읽는다(류지영 매니저가 시트만 고치면 앱도 따라간다).
    관리자검증상태는 10시트에 없어 기본값을 함께 준다.
    """
    if _codes_cache["v"] and time.time() - _codes_cache["t"] < 300:
        return _codes_cache["v"]
    out = {"관리자검증상태": ["일치", "추가작업발생", "작업내용누락", "확인필요"]}
    try:
        import openpyxl
        from ecount_reconcile import load_config, resolve_master
        wb = openpyxl.load_workbook(resolve_master(load_config()["reconcile"]["master_xlsx"]),
                                    read_only=True, data_only=True)
        ws = wb["10_코드관리"]
        rows = list(ws.iter_rows(min_row=4, values_only=True))
        if rows:
            hdr = [str(h).strip() if h else "" for h in rows[0]]
            for i, name in enumerate(hdr):
                if not name:
                    continue
                vals = []
                for r in rows[1:]:
                    v = r[i] if i < len(r) else None
                    if v not in (None, "") and str(v).strip() not in vals:
                        vals.append(str(v).strip())
                if vals:
                    out[name] = vals
        wb.close()
    except Exception as e:
        out["_error"] = str(e)[:80]
    _codes_cache.update({"t": time.time(), "v": out})
    return out


def enqueue_codes(codes):
    """폰이 예약한 프로젝트 코드를 실제 원장 행으로 등록한다.
    쓰기는 전부 ledger_writer(빈 칸만·근거 필수·vN+1)를 거치므로 기존 값은 덮이지 않는다."""
    import project_resolve as P
    ev = P.evidence()
    items, done, skip = [], [], []
    for c in codes:
        r = P.resolve(c, ev)
        if not r.get("ok"):
            skip.append({"code": c, "why": r.get("reason", "형식 오류")})
        elif r["state"] == "등록됨":
            skip.append({"code": c, "why": f"이미 {r['sheet']} {r.get('row')}행에 있습니다"})
        else:
            items += P.row_items(r, ev)
            done.append(c)
            # 같은 요청에 두 건이 오면 뒤엣것이 같은 행을 노린다 — 자리를 미리 물린다
            ev["tail"][r["sheet"]] = r["row"]
    if not items:
        return {"ok": True, "applied": 0, "skipped": skip}
    import ledger_writer as L
    L.queue_add(items)
    p = subprocess.run([PY, os.path.join(ROOT, "ledger_writer.py"), "--apply"],
                       cwd=ROOT, env=ENV, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    ok = p.returncode == 0
    runner["log"].append(f"[폰 예약] {len(done)}건 등록 시도 — {'성공' if ok else '실패'}")
    return {"ok": ok, "applied": len(done) if ok else 0, "codes": done, "skipped": skip,
            "msg": (p.stdout or "").strip().splitlines()[-1:] or [""]}


def start_task(key):
    with _rlock:
        if runner["busy"]:
            return False, "다른 작업 실행 중"
        if key not in TASKS:
            return False, "알 수 없는 작업"
        if DEMO:
            runner["log"].append(f"[데모] '{TASKS[key][0]}' — 합성 환경에서는 실행을 시뮬레이션합니다.")
            return True, "demo"
        runner["busy"], runner["task"] = True, TASKS[key][0]
        runner["log"].clear()

    def work():
        title, args = TASKS[key]
        runner["log"].append(f"===== {title} 시작 {datetime.now():%H:%M:%S} =====")
        try:
            p = subprocess.Popen([PY] + args, cwd=ROOT, env=ENV, stdout=subprocess.PIPE,
                                 stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
            for ln in p.stdout:
                if "UserWarning" not in ln and "warn(msg)" not in ln:
                    runner["log"].append(ln.rstrip())
            p.wait()
            runner["log"].append(f"===== 종료 (코드 {p.returncode}) =====")
        except Exception as e:
            runner["log"].append(f"오류: {e}")
        finally:
            runner["busy"], runner["done_at"] = False, datetime.now().isoformat()
    threading.Thread(target=work, daemon=True).start()
    return True, "started"


# ───────────────────────── 데이터 ─────────────────────────
_cache = {"t": 0, "settle": None, "status": None}
_readlock = threading.RLock()  # Z:드라이브 엑셀 동시 읽기 직렬화(스레드 충돌 방지)
# ★ RLock이어야 한다: 정산 조회가 락을 쥔 채 업무 조회(대표번호 색인용)를 부르므로
#   일반 Lock이면 같은 스레드에서 자기 자신을 기다리다 멈춘다(실제로 응답 없음 발생)


# ── 날짜 정렬 공통 규칙 ────────────────────────────────────────
# 앱·리포트 어디서나 **과거가 맨 위, 최근이 맨 아래**(오름차순)로 통일한다.
# 새로 추가되는 행도 반드시 이 함수를 거치므로 따로 정렬해 줄 필요가 없다.
DATE_KEYS = {
    "settle": ("완료일", "계산서발행일", "명세서발행일", "입금일"),
    "as":     ("접수일자", "작업완료일", "방문예정일"),
    "pm":     ("점검예정일", "실제점검일"),
}
_DATE_RE = re.compile(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})")


def norm_date(v):
    """'2026.6.3' · '2026-06-03 00:00' → '2026-06-03' (문자열 비교로 시간순이 되게)"""
    m = _DATE_RE.search(str(v or ""))
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def row_date(rec, keys=()):
    """행의 대표 날짜. 지정 키를 우선 보고, 없으면 아무 날짜 값이나 찾아 쓴다."""
    for k in keys:
        d = norm_date(rec.get(k))
        if d:
            return d
    for v in rec.values():
        d = norm_date(v)
        if d:
            return d
    return ""


def sort_by_date(rows, kind, idkey=None):
    """과거 → 최근. 날짜가 없는 행은 맨 뒤(=가장 최근으로 취급), 동률은 ID순."""
    keys = DATE_KEYS.get(kind, ())
    return sorted(rows, key=lambda r: ((d := row_date(r, keys)) == "", d,
                                       str(r.get(idkey) or "") if idkey else ""))


def demo_settlements():
    camps = ["송파5MB(감일동)", "울산2캠프", "인천7MB(마곡동)", "부천3(BUC3)", "대전1캠프",
             "구리1캠프", "제주1Sub-hub", "창원1MB(팔용동)", "군포1Sub-Hub", "광주2Sub-hub"]
    rows = []
    rnd = random.Random(42)
    for i in range(1, 16):
        amt = rnd.choice([380000, 418000, 470800, 760000, 1230000, 1472500])
        st = rnd.choice(["정상", "세금계산서 미발행", "ERP 미확인", "미청구", "입금 대기"])
        d = (date(2026, 7, 1) + timedelta(days=i)).isoformat()
        rows.append({"정산ID": f"JS-2607-{i:03d}", "업무구분": rnd.choice(["돌발AS", "정기점검"]),
                     "캠프명": camps[i % len(camps)], "프로젝트NO": f"UJ26{1000+i}",
                     "원천업무ID": f"AS-2607-{i:03d}",
                     "공급가액": amt, "부가세": int(amt * 0.1), "합계": int(amt * 1.1),
                     "명세서": "있음" if st != "미청구" else "없음",
                     "명세서번호": f"2026/07/{i:02d}-1" if st != "미청구" else "",
                     "명세서발행일": d if st != "미청구" else "",
                     "계산서": "발행" if st == "정상" else "미발행",
                     "계산서발행일": d if st == "정상" else "", "승인번호": "",
                     "입금일": d if st == "정상" else "", "입금액": int(amt * 1.1) if st == "정상" else 0,
                     "미수금": 0 if st == "정상" else int(amt * 1.1), "비용구분": "유상",
                     "상태": st, "완료일": d})
    return sort_by_date(rows, "settle", "정산ID")


def real_settlements():
    from ecount_reconcile import read_ledger, load_config
    cfg = load_config()
    recs = read_ledger(cfg["reconcile"]["master_xlsx"])
    rows = []
    for sid, r in sorted(recs.items()):
        issued = r.get("원장_세금계산서실제발행일") or r.get("원장_세금계산서발행일")
        has_stmt = bool(str(r.get("원장_거래명세서번호") or "").strip())
        if r.get("비용구분") != "유상":
            st = "무상/보험"
        elif not r.get("원장_공급가액"):
            st = "금액 미입력"
        elif not has_stmt:
            st = "미청구(전표 없음)"
        elif not issued:
            st = "세금계산서 미발행"
        elif not r.get("원장_입금일"):
            st = "입금 대기"
        else:
            st = "정상"
        rows.append({"정산ID": sid, "업무구분": r.get("업무구분"), "캠프명": r.get("캠프명"),
                     "프로젝트NO": r.get("프로젝트NO"), "원천업무ID": r.get("원천업무ID"),
                     "공급가액": r.get("원장_공급가액") or 0, "합계": r.get("원장_합계") or 0,
                     "부가세": r.get("원장_부가세"),
                     "명세서": "있음" if has_stmt else "없음",
                     "명세서번호": r.get("원장_거래명세서번호") or "",
                     "명세서발행일": str(r.get("원장_거래명세서발행일") or "")[:10],
                     "계산서": "발행" if issued else "미발행",
                     "계산서발행일": str(issued or "")[:10],
                     "승인번호": r.get("원장_세금계산서승인번호") or "",
                     "청구일": str(r.get("원장_청구일") or "")[:10],
                     "지급예정일": str(r.get("원장_지급예정일") or "")[:10],
                     "입금일": str(r.get("원장_입금일") or "")[:10],
                     "입금액": r.get("원장_입금액") or 0,
                     "미수금": r.get("원장_미수금액") if r.get("원장_미수금액") is not None else "",
                     "비용구분": r.get("비용구분"),
                     "PO필요": r.get("원장_PO필요여부") or "",
                     "PO번호": r.get("원장_PO번호") or "",
                     "PO발행일": str(r.get("원장_PO발행일") or "")[:10],
                     "상태": st, "완료일": str(r.get("작업완료일") or "")[:10]})
    return sort_by_date(rows, "settle", "정산ID")


def real_works():
    """02 돌발AS·04 정기점검 현황 (앱 '업무' 데이터)"""
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    out = {"as": [], "pm": []}
    spec = {
        # 뒤쪽 3개는 '확인 완료' 표시 — 관리자가 검증한 건인지 카드에서 바로 보이게 한다.
        "02_돌발AS접수": ("as", ["접수ID", "프로젝트NO", "캠프명", "접수일자", "담당기사", "진행상태",
                                "작업완료일", "유상·무상·보험", "신청내용", "긴급도", "방문예정일",
                                "관리자검증상태", "최종확인일", "사진등록",
                                "검증결과", "검증문제코드"]),
        "04_정기점검": ("pm", ["점검ID", "프로젝트NO", "캠프명", "점검예정일", "실제점검일", "점검상태",
                              "담당기사", "이상발견여부", "돌발AS전환여부",
                              "최종확인일(유현민 체크)", "점검사진",
                              "검증결과", "검증문제코드"]),
    }
    for sheet, (key, cols) in spec.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        idx = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
        for row in ws.iter_rows(min_row=5, values_only=True):
            # ID 열은 수식 — 새로 추가된 행은 엑셀을 열기 전까지 캐시값이 없다.
            # 프로젝트NO를 대체 키로 써서 백필 행도 앱에 표시한다.
            rid = row[idx[cols[0]]] if cols[0] in idx else None
            if not rid and "프로젝트NO" in idx and idx["프로젝트NO"] < len(row):
                rid = row[idx["프로젝트NO"]]
            if not rid:
                continue
            # 행 확장(expand_rows)으로 만들어 둔 **빈 예비행**은 ID 수식이 값을 내므로
            # rid만으로는 걸러지지 않는다. 날짜도 캠프명도 없으면 실제 업무가 아니다.
            _d = row[idx[cols[3]]] if cols[3] in idx and idx[cols[3]] < len(row) else None
            _c = row[idx["캠프명"]] if "캠프명" in idx and idx["캠프명"] < len(row) else None
            if not _d and not _c:
                continue
            rec = {}
            for c in cols:
                v = row[idx[c]] if c in idx and idx[c] < len(row) else None
                rec[c] = str(v)[:10] if hasattr(v, "year") else ("" if v is None else str(v))
            derive_status(rec, key)
            out[key].append(rec)
    wb.close()
    try:
        out["as"] += erp_work_rows(out["as"], "as")
        out["pm"] += erp_work_rows(out["pm"], "pm")
        idx = build_prj_index(out)
        apply_rep_no(out["as"], idx, "접수ID")
        apply_rep_no(out["pm"], idx, "점검ID")
    except Exception:
        pass
    out["as"] = sort_by_date(out["as"], "as", "접수ID")
    out["pm"] = sort_by_date(out["pm"], "pm", "점검ID")
    return out


def derive_status(rec, kind):
    """상태 열은 **수식**이라 새로 넣은 행은 엑셀을 한 번 열기 전까지 캐시값이 없다(None).
    그대로 두면 완료된 점검 90여 건이 전부 '미점검'으로 보인다 → 원본 열로 직접 판정한다.
    (판정 규칙은 시트 수식과 동일: 완료일 있으면 완료, 예정일이 지났으면 미점검)"""
    today = date.today().isoformat()
    if kind == "pm":
        if str(rec.get("점검상태") or "").strip():
            return
        if str(rec.get("실제점검일") or "").strip():
            rec["점검상태"] = "완료"
        elif str(rec.get("돌발AS전환여부") or "").strip():
            rec["점검상태"] = "AS전환"
        elif not str(rec.get("점검예정일") or "").strip():
            rec["점검상태"] = ""
        else:
            rec["점검상태"] = "미점검" if str(rec["점검예정일"])[:10] < today else "예정"
    else:
        if str(rec.get("진행상태") or "").strip():
            return
        rec["진행상태"] = "작업완료" if str(rec.get("작업완료일") or "").strip() else "접수"


def demo_works():
    rnd = random.Random(7)
    camps = ["송파5MB(감일동)", "울산2캠프", "인천7MB(마곡동)", "대전1캠프", "구리1캠프"]
    techs = ["김준형", "권오철", "김필우", "차동호"]
    a = [{"접수ID": f"AS-2607-{i:03d}", "캠프명": rnd.choice(camps),
          "접수일자": (date(2026, 7, 1) + timedelta(days=i)).isoformat(), "담당기사": rnd.choice(techs),
          "진행상태": rnd.choice(["접수", "방문예정", "작업중", "작업완료"]),
          "작업완료일": "", "유상·무상·보험": rnd.choice(["유상", "무상"]),
          "신청내용": "도어 센서 교체 외", "긴급도": rnd.choice(["보통", "긴급"])} for i in range(1, 11)]
    p = [{"점검ID": f"PM-2607-{i:03d}", "캠프명": rnd.choice(camps),
          "점검예정일": (date(2026, 7, 5) + timedelta(days=i * 2)).isoformat(),
          "실제점검일": "" if i % 3 == 0 else (date(2026, 7, 5) + timedelta(days=i * 2)).isoformat(),
          "점검상태": "예정" if i % 3 == 0 else "완료", "담당기사": rnd.choice(techs),
          "이상발견여부": rnd.choice(["없음", "있음"]), "돌발AS전환여부": "미전환"} for i in range(1, 8)]
    return {"as": a, "pm": p}


def _master_mtime():
    try:
        from ecount_reconcile import load_config, resolve_master
        return os.path.getmtime(resolve_master(load_config()["reconcile"]["master_xlsx"]))
    except Exception:
        return 0


def _fresh(key):
    """엑셀이 바뀌면(mtime) 즉시 + 120초 TTL로 캐시 무효화 — '엑셀·대조 변경 → 앱 자동 반영'"""
    mt = _master_mtime()
    if _cache.get("mt") != mt or time.time() - _cache.get("ts", 0) > 120:
        _cache.clear()
        _cache["mt"], _cache["ts"] = mt, time.time()
    return _cache.get(key)


def get_works():
    if DEMO:
        return demo_works()
    with _readlock:
        w = _fresh("works")
        if w:
            return w
        w = real_works()
        _cache["works"] = w
        return w


def _fmtv(v):
    """01시트 값 표시용: 부동소수 오차 정리·천단위·날짜"""
    if v is None or v == "":
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = round(float(v), 2)
        n = int(n) if abs(n - round(n)) < 0.01 else n
        return f"{n:,}"
    return str(v).strip()


def read_exec_report(master):
    """01_대표보고 시트를 구조 그대로 읽는다(엑셀 수식이 곧 집계 로직 — 앱에서 재계산하지 않음)."""
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    if "01_대표보고" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["01_대표보고"]
    rows = [r for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60), values_only=True)]
    wb.close()
    out = {"meta": {}, "summary": [], "sections": []}
    cur = None
    for i, row in enumerate(rows, 1):
        g = lambda j: row[j] if j < len(row) else None
        a = _fmtv(g(0))
        if i == 4:                                     # 보고일·집계기준일·보고자
            for li, vi in ((0, 1), (3, 4), (6, 7)):
                lab, val = _fmtv(g(li)), _fmtv(g(vi))
                if lab:
                    out["meta"][lab] = val
            continue
        if a.startswith("■"):
            out["summary"].append(a)
            continue
        # 섹션 헤더는 "2.  당일 업무 실적"처럼 숫자 뒤에 한글이 온다.
        # "1. [돌발AS] …" 같은 TOP5 항목은 대괄호로 시작하므로 헤더가 아니다.
        if re.match(r"^\d+\.\s*[가-힣]", a):
            cur = {"title": re.sub(r"\s+", " ", a), "items": [], "lines": []}
            out["sections"].append(cur)
            continue
        if not cur or a.startswith("※"):
            continue
        # 블록 헤더 행: "[ 돌발 AS · 현장 ]  [ 정기점검 ]  [ 거래서류 · 청구 ]"
        # → 열 위치별 그룹을 만들어 이후 행의 항목을 각 그룹에 담는다(AS/점검 구분 유지)
        heads = {li: _fmtv(g(li)).strip("[] ") for li in (0, 3, 6) if _fmtv(g(li)).startswith("[")}
        if heads:
            cur["colgroups"] = {}
            for li, name in heads.items():
                grp = {"name": name, "items": []}
                cur.setdefault("groups", []).append(grp)
                cur["colgroups"][li] = grp
            continue
        for li, vi in ((0, 1), (3, 4), (6, 7)):        # 3열 그룹: 라벨|값
            lab, val = _fmtv(g(li)), _fmtv(g(vi))
            if not lab or lab.startswith("["):
                continue
            if val == "" and len(lab) > 20:            # 값 없는 긴 문장 = 서술형(TOP5 등)
                cur["lines"].append(lab)
            else:
                grp = (cur.get("colgroups") or {}).get(li)
                (grp["items"] if grp else cur["items"]).append([lab, val])
    for s in out["sections"]:
        s.pop("colgroups", None)                       # 내부 매핑은 응답에서 제외
    return out


def get_exec_report():
    if DEMO:
        return {"meta": {"보고일": "2026-07-25", "집계기준일": "2026-07-24", "보고자": "유현민"},
                "summary": ["■ 데모 요약"], "sections": [
                    {"title": "2. 당일 업무 실적", "items": [["신규 접수", "3"], ["작업 완료", "1"]], "lines": []}]}
    with _readlock:
        r = _fresh("exec")
        if r:
            return r
        from ecount_reconcile import load_config, resolve_master
        r = read_exec_report(resolve_master(load_config()["reconcile"]["master_xlsx"]))
        _cache["exec"] = r
        return r


def get_issues():
    """07_불일치누락현황 — 엑셀의 '검증 안 된·확인해야 할' 항목 그대로"""
    if DEMO:
        return {"rows": [{"문제유형": "세금계산서 미발행", "업무ID": "JS-2607-002", "캠프명": "울산2캠프",
                          "문제내용": "명세서 발행 후 계산서 미발행", "담당자": "유현민"}], "cols": []}
    r = _fresh("issues")
    if r:
        return r
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    rows = []
    # 1순위: 관리대장 통합 시트 23_확인필요현황 (에이전트가 매일 갱신 — 단일 엑셀 관리)
    if "23_확인필요현황" in wb.sheetnames:
        ws = wb["23_확인필요현황"]
        hdr = [str(h).strip() for h in next(ws.iter_rows(min_row=4, max_row=4, values_only=True)) if h]
        for row in ws.iter_rows(min_row=5, values_only=True):
            vals = {hdr[i]: ("" if i >= len(row) or row[i] is None else str(row[i]))
                    for i in range(len(hdr))}
            if any(v for v in vals.values()):
                rows.append(vals)
        wb.close()
        out = {"rows": sort_by_date(apply_rep_no(rows), "check"), "cols": hdr, "source": "23_확인필요현황"}
        _cache["issues"] = out
        return out
    if "07_불일치누락현황" in wb.sheetnames:
        ws = wb["07_불일치누락현황"]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        heads = [(i, str(h).strip()) for i, h in enumerate(hdr) if h is not None]
        for row in ws.iter_rows(min_row=5, values_only=True):
            vals = {h: ("" if i >= len(row) or row[i] is None else
                        (str(row[i])[:10] if hasattr(row[i], "year") else str(row[i])))
                    for i, h in heads}
            if any(v for v in vals.values()):
                rows.append(vals)
            if len(rows) >= 300:
                break
    wb.close()
    # 대조 결과(밴드·카톡·ERP원장·쿠팡PO) 통합 — 07시트 위에 얹어 한 화면에서 전부 확인
    merged = []
    try:
        from findings_export import latest_csv
        for src, pat, filt in (("밴드 게시 미확인", "밴드대조_*.csv", lambda r: r.get("밴드게시") == "미확인"),
                               ("카톡 보고 미확인", "카톡대조_*.csv", lambda r: r.get("카톡보고") == "미확인"),
                               ("ERP원장 문제", "ERP원장대조_*.csv", lambda r: True),
                               ("쿠팡PO 문제", "PO대조_*.csv", lambda r: True)):
            for r in latest_csv(pat):
                if filt(r):
                    merged.append({"문제유형": src + (f"({r['유형']})" if r.get("유형") else ""),
                                   "업무ID": r.get("ID") or r.get("정산ID") or r.get("전표") or r.get("PO번호") or "",
                                   "캠프명": r.get("캠프명", ""), "담당자": r.get("담당기사", ""),
                                   "문제내용": (r.get("판정") or r.get("내용") or
                                                f"완료 {r.get('완료일','')}" ) [:100]})
    except Exception:
        pass
    rows = sort_by_date(apply_rep_no(merged + rows), "check")
    cols = []
    for r in rows[:50]:
        for k in r:
            if k not in cols:
                cols.append(k)
    out = {"rows": rows, "cols": cols}
    _cache["issues"] = out
    return out


def build_id():
    """index.html이 바뀌면 값이 달라진다. 폰에 열려 있는 앱이 구버전인지 판별하는 기준."""
    try:
        st = os.stat(os.path.join(BASE, "index.html"))
        return hashlib.md5(f"{int(st.st_mtime)}-{st.st_size}".encode()).hexdigest()[:10]
    except Exception:
        return "0"


def brand_logo():
    """webapp/brand/ 에 넣어 둔 고객사 로고 파일명. 없으면 빈 문자열(기본 CSOS 마크 사용).
    파일은 gitignore 대상 — 상표 자산을 공개 저장소에 올리지 않기 위해서다."""
    d = os.path.join(BASE, "brand")
    if not os.path.isdir(d):
        return ""
    for f in sorted(os.listdir(d)):
        if os.path.splitext(f)[1].lower() in (".png", ".svg", ".jpg", ".jpeg", ".webp"):
            return f
    return ""


def get_erpdocs():
    """25_ERP매출서류 — 이카운트 매출(세금)계산서 원본(2026년 전체).
    ERP는 여러 작업을 한 장으로 묶어 발행하므로 1행 = 작업 1건이 아니다."""
    if DEMO:
        return {"rows": [], "months": {}, "total": 0}
    r = _fresh("erpdocs")
    if r:
        return r
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    out = {"rows": [], "months": {}, "total": 0, "kinds": {}}
    try:
        wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
        if "25_ERP매출서류" in wb.sheetnames:
            for row in wb["25_ERP매출서류"].iter_rows(min_row=5, values_only=True):
                if not row or not row[0]:
                    continue
                slip, mo, kind, sup = row[0], row[1], row[2], int(row[3] or 0)
                out["rows"].append({"전표": str(slip), "월": str(mo), "유형": str(kind or ""),
                                    "공급가액": sup, "거래처": str(row[6] or ""),
                                    "프로젝트명": str(row[7] or "")})
                m = out["months"].setdefault(str(mo), {"합계": 0, "건수": 0})
                m["합계"] += sup
                m["건수"] += 1
                m[str(kind)] = m.get(str(kind), 0) + sup
                out["kinds"][str(kind)] = out["kinds"].get(str(kind), 0) + sup
                out["total"] += sup
        # 26_계산서구성 — 계산서 1장에 어떤 프로젝트가 묶였는지(추정). 전표번호로 붙인다.
        if "26_계산서구성" in wb.sheetnames:
            comp = {}
            for row in wb["26_계산서구성"].iter_rows(min_row=5, values_only=True):
                if not row or not row[0]:
                    continue
                comp[str(row[0])] = {"포함건수": int(row[5] or 0),
                                     "포함프로젝트": str(row[6] or ""),
                                     "후보합계": int(row[7] or 0),
                                     "판정": str(row[8] or "")}
            for r in out["rows"]:
                r.update(comp.get(r["전표"], {"포함건수": 0, "포함프로젝트": "",
                                             "후보합계": 0, "판정": "미상"}))
        wb.close()
    except Exception as e:
        out["error"] = str(e)
    out["rows"] = sort_by_date(out["rows"], "erpdocs")
    _cache["erpdocs"] = out
    return out


def get_checks():
    """최근 카톡·밴드·ERP원장·쿠팡PO 대조 CSV를 ID별로 조인 — 4원천 검증 배지"""
    import csv as _csv
    out = {}
    def latest(pat):
        fs = sorted(glob.glob(os.path.join(ROOT, "reports", pat)))
        return fs[-1] if fs else None
    f = latest("카톡대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            out.setdefault(r.get("ID", ""), {})["kakao"] = r.get("카톡보고", "")
    f = latest("밴드대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            out.setdefault(r.get("ID", ""), {})["band"] = r.get("밴드게시", "")
    f = latest("ERP원장대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            sid = r.get("정산ID", "") or r.get("전표", "")
            for one in str(sid).split(","):
                if one.strip():
                    out.setdefault(one.strip(), {})["erp"] = r.get("유형", "") + " " + r.get("판정", "")
    f = latest("PO대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            sid = r.get("정산ID", "") or r.get("ID", "")
            for one in str(sid).split(","):
                if one.strip():
                    out.setdefault(one.strip(), {})["po"] = (r.get("판정", "") or r.get("유형", "")).strip()
    if DEMO and not out:
        out = {"JS-2607-002": {"kakao": "확인", "band": "미확인", "erp": "D 금액불일치",
                               "po": "PO 미발행"}}
    return out


_UJ_RE = re.compile(r"(?<![A-Za-z0-9])UJ\d{6,}(?![0-9])")
# 계산서 제목에서 캠프명만 뽑는다: '쿠팡신규_송파1MB(감일동)-이동식…' → '송파1MB(감일동)'
_CAMP_RE = re.compile(r"[가-힣A-Za-z]+\d*(?:BMB|MB|캠프|Sub-?FC|Sub-?hub|FC)(?:\([^)]*\))?",
                      re.I)   # sub-hub / Sub-Hub 표기가 섞여 있어 대소문자 무시


def camp_of(title):
    m = _CAMP_RE.search(str(title or ""))
    if m:
        return m.group()
    t = re.sub(r"^(쿠팡\S*|돌발AS|정기점검)[_\s-]*", "", str(title or "")).strip()
    return (t or str(title or ""))[:28]


def erp_settlement_rows(ledger_rows):
    """관리대장에 정산 행이 **아예 없는 달**만 ERP 계산서로 채워 넣는다.

    왜 이렇게 하나
      · 06시트는 '작업 1건 = 1행' 구조다(업무구분이 원천업무ID 기반 수식).
        ERP 계산서는 여러 작업을 묶은 것이라 그 시트에 그대로 넣으면 수식이 어긋난다.
      · 그렇다고 1~6월을 비워두면 앱에서 그 달 매출이 0으로 보인다(사실과 다름).
      → 대장에 자료가 있는 달은 대장 우선, 없는 달만 ERP로 보완하고 출처를 표시한다.
        (같은 달을 양쪽에서 세지 않으므로 이중 계상이 없다)
    """
    have = {str(r.get("완료일") or "")[:7].replace("-", "/") for r in ledger_rows
            if r.get("공급가액")}
    docs = get_erpdocs()
    out = []
    for d in docs.get("rows", []):
        mo = d.get("월") or ""
        if not mo or mo in have:
            continue
        slip = d.get("전표") or ""
        iso = slip[:10].replace("/", "-")
        title = d.get("프로젝트명") or ""
        prj = (_UJ_RE.search(title) or [""])
        prj = prj.group() if hasattr(prj, "group") else ""

        sup = int(d.get("공급가액") or 0)
        out.append({
            "정산ID": "ERP-" + slip.replace("/", "").replace("-", "-"),
            "업무구분": d.get("유형") or "", "캠프명": camp_of(title),
            "프로젝트NO": prj, "원천업무ID": "",
            "공급가액": sup, "합계": sup + round(sup * 0.1),
            "명세서": "있음", "명세서번호": slip, "명세서발행일": iso,
            "계산서": "발행", "계산서발행일": iso, "승인번호": "",
            "청구일": "", "지급예정일": "", "입금일": "", "입금액": 0, "미수금": "",
            "비용구분": "유상", "PO필요": "", "PO번호": "", "PO발행일": "",
            "상태": "ERP 계산서(묶음)", "완료일": iso,
            "출처": "ERP", "적요": title})
    return out


# ── 대표 프로젝트NO ────────────────────────────────────────────
# 모든 건이 번호로 식별되게 한다. 우선순위:
#   1) 행에 이미 있는 프로젝트NO
#   2) 행 안 어딘가(내용·근거 등)에 적힌 UJ 번호
#   3) 같은 캠프·같은 달의 실제 작업에서 찾은 대표 번호
#   4) 그래도 없으면 ERP 전표번호 기반 식별자(UJ처럼 보이지 않게 'ERP-' 접두)
#      — 없는 UJ 번호를 지어내면 실제 번호와 헷갈리므로 절대 만들지 않는다.
def _camp_key(v):
    return re.sub(r"[\s()·]", "", str(v or "")).lower()[:14]


def build_prj_index(works):
    idx = {}
    for kind, dk in (("as", "접수일자"), ("pm", "점검예정일")):
        for r in works.get(kind, []):
            if r.get("출처") == "ERP":
                continue
            prj = str(r.get("프로젝트NO") or "").strip()
            if not prj:
                continue
            mo = norm_date(r.get(dk) or r.get("작업완료일") or r.get("실제점검일"))[:7]
            idx.setdefault((_camp_key(r.get("캠프명")), mo), []).append(prj)
    return idx


def rep_no(rec, idx=None, slip=""):
    """대표 프로젝트NO를 정한다(순수 함수 — 합성검증 대상)"""
    cur = str(rec.get("프로젝트NO") or "").strip()
    if cur:
        return cur, ""
    for v in rec.values():                       # 내용·근거 등 본문에 적힌 UJ
        m = _UJ_RE.search(str(v or ""))
        if m:
            return m.group(), "본문"
    if idx:
        mo = ""
        for k in ("완료일", "접수일자", "점검예정일", "일자", "작업완료일"):
            mo = norm_date(rec.get(k))[:7]
            if mo:
                break
        hits = idx.get((_camp_key(rec.get("캠프명")), mo))
        if hits:
            return sorted(hits)[0], "동일캠프·동월"
    if slip:
        m = re.match(r"(\d{2})(\d{2})/(\d{2})/(\d{2})\s*-\s*(\d+)", str(slip))
        if m:
            return f"ERP-{m.group(2)}{m.group(3)}{m.group(4)}-{m.group(5)}", "전표"
        return "ERP-" + re.sub(r"[^0-9A-Za-z-]", "", str(slip))[-10:], "전표"
    # 최후: 그 행 자신의 ID를 대표번호로 쓴다 — 번호 없는 행이 하나도 남지 않게
    for k in ("정산ID", "접수ID", "점검ID", "업무ID", "ID"):
        v = str(rec.get(k) or "").strip()
        if v:
            return v, "자체ID"
    return "", ""


def apply_rep_no(rows, idx=None, slipkey=None):
    for r in rows:
        no, how = rep_no(r, idx, str(r.get(slipkey) or "") if slipkey else "")
        if no and not str(r.get("프로젝트NO") or "").strip():
            r["프로젝트NO"] = no
            r["대표번호출처"] = how
    return rows


def erp_work_rows(existing, kind):
    """02/04 시트에 자료가 **아예 없는 달**만 ERP 계산서로 보완한다(정산과 같은 규칙).
    ERP 계산서 1장 = 작업 여러 건 묶음이므로 '건수'가 아니라 '그 달에 이런 업무가 있었다'는
    사실을 보여주는 용도다. 자료가 있는 달은 건드리지 않아 이중 계상이 없다."""
    dk = {"as": ("접수일자", "작업완료일"), "pm": ("점검예정일", "실제점검일")}[kind]
    want = {"as": "돌발AS", "pm": "정기점검"}[kind]
    have = {norm_date(r.get(dk[0]) or r.get(dk[1]))[:7] for r in existing}
    have = {h.replace("-", "/") for h in have if h}
    out = []
    for d in get_erpdocs().get("rows", []):
        if (d.get("유형") or "") != want:
            continue
        mo = d.get("월") or ""
        if not mo or mo in have:
            continue
        slip = d.get("전표") or ""
        iso = slip[:10].replace("/", "-")
        title = d.get("프로젝트명") or ""
        prj = _UJ_RE.search(title)
        base = {"프로젝트NO": prj.group() if prj else "", "캠프명": camp_of(title),
                "담당기사": "", "유상·무상·보험": "유상", "비고": "ERP 계산서 기준(작업 묶음)",
                "출처": "ERP", "적요": title}
        if kind == "as":
            base.update({"접수ID": "ERP-" + slip.replace("/", ""), "접수일자": iso,
                         "작업완료일": iso, "진행상태": "작업완료", "신청내용": title,
                         "긴급도": "", "방문예정일": ""})
        else:
            base.update({"점검ID": "ERP-" + slip.replace("/", ""), "점검예정일": iso,
                         "실제점검일": iso, "점검상태": "완료",
                         "이상발견여부": "", "돌발AS전환여부": ""})
        out.append(base)
    return out


def get_settlements():
    if DEMO:
        return demo_settlements()
    with _readlock:
        r = _fresh("settle")
        if r:
            return r
        rows = real_settlements()
        try:
            rows = rows + erp_settlement_rows(rows)
            idx = build_prj_index(get_works())
            apply_rep_no(rows, idx, "명세서번호")
            rows = sort_by_date(rows, "settle", "정산ID")
        except Exception:
            pass
        _cache["settle"] = rows
        return rows


def get_status():
    if DEMO:
        return {"master": "쿠팡_통합업무_일일보고_관리대장_v23.xlsx (데모)", "fork": [],
                "agent_last": "2026-07-24 09:50", "steps": [
                    {"n": "합성검증", "s": "ok"}, {"n": "판매·세금계산서 대조", "s": "ok"},
                    {"n": "ERP원장 4유형 대조", "s": "ok"}, {"n": "밴드 수집·대조", "s": "skip"},
                    {"n": "카톡 대조", "s": "ok"}, {"n": "관리대장 자동입력", "s": "ok"},
                    {"n": "전표 전송대기", "s": "ok"}],
                "pending_updates": 2, "inbox": 1, "kakao": 2, "band": False, "demo": True}
    try:
        from coupang_workbench import get_status as ws
        st = ws()
        # 동기화 백본: 에이전트가 쓴 agent_status.json 우선 (없으면 md 리포트 파싱)
        steps, rt = [], ""
        try:
            aj = json.load(open(os.path.join(ROOT, "reports", "agent_status.json"), encoding="utf-8"))
            steps = aj.get("steps", [])
            rt = aj.get("time", "")[:16].replace("T", " ")
        except Exception:
            for s in st.get("report_summary", []):
                mark = "ok" if "✅" in s else ("skip" if "스킵" in s else "fail")
                steps.append({"n": re.sub(r"[✅❌⏭]|스킵|실패", "", s).strip(), "s": mark})
            rt = st.get("report_time", "")
            if rt:
                rt = f"{rt[:4]}-{rt[4:6]}-{rt[6:8]} {rt[9:11]}:{rt[11:13]}"
        tunnel = ""
        try:
            tunnel = open(os.path.join(ROOT, "reports", "tunnel_url.txt"), encoding="utf-8").read().strip()
        except Exception:
            pass
        # 3원천 검증 실집계 (최신 대조 CSV) — 보고서·대시보드가 하드코딩 없이 실데이터를 쓰도록
        srcs = {}
        try:
            from findings_export import latest_csv
            for key, pat, col, okv in (("band", "밴드대조_*.csv", "밴드게시", "확인"),
                                       ("kakao", "카톡대조_*.csv", "카톡보고", "확인")):
                rows = latest_csv(pat)
                if rows:
                    ok = sum(1 for r in rows if r.get(col) == okv)
                    miss = [r for r in rows if r.get(col) != okv]
                    srcs[key] = {"total": len(rows), "ok": ok, "miss": len(miss),
                                 "miss_prj": [r.get("프로젝트NO") or r.get("ID") for r in miss[:8]]}
            erp = latest_csv("ERP원장대조_*.csv")
            if erp:
                srcs["erp"] = {"total": len(erp), "ok": 0, "miss": len(erp),
                               "miss_prj": [r.get("정산ID") or r.get("전표") for r in erp[:8]]}
            po = latest_csv("PO대조_*.csv")
            if po:
                srcs["po"] = {"total": len(po), "ok": 0, "miss": len(po),
                              "miss_prj": [r.get("PO번호") or r.get("정산ID") for r in po[:8]]}
        except Exception:
            pass

        # ★ 대조 리포트만 보면 '자료가 있는데 왜 없다고 하냐'는 말이 나온다 — 리포트가
        #   없는 이유가 (1) 파일을 안 넣었다 (2) 넣었는데 **파일이 비어 있다**
        #   (3) 대조를 안 돌렸다 로 갈리기 때문이다. 2026-07-27에 실제로 ERP 파일 3개가
        #   회사명 한 줄만 있는 빈 파일이었고, 아무도 그걸 몰랐다. 그래서 여기서 같이 본다.
        try:
            from inbox_scan import pick
            import openpyxl as _ox
            for key, kinds in (("erp", ("ledger", "slips")), ("po", ("po",)),
                               ("tax", ("tax",)), ("stmt", ("stmt",))):
                files = []
                for kd in kinds:
                    files += pick(kd) or []
                info = {"files": len(files), "rows": 0, "empty": []}
                for f in files[:6]:
                    try:
                        w = _ox.load_workbook(f, read_only=True, data_only=True)
                        n = 0
                        for sn in w.sheetnames:
                            n += sum(1 for r in w[sn].iter_rows(values_only=True)
                                     if sum(1 for x in r if x not in (None, "")) >= 3)
                        w.close()
                        info["rows"] += n
                        if n < 2:                      # 회사명 한 줄만 있는 '빈 내보내기'
                            info["empty"].append(os.path.basename(f))
                    except Exception:
                        pass
                if files:
                    srcs.setdefault(key, {})["inbox"] = info
        except Exception:
            pass

        return {"master": os.path.basename(st.get("master", "") or "") + "  " + st.get("master_label", ""),
                "fork": st.get("fork", []), "agent_last": rt or "기록 없음", "steps": steps,
                "pending_updates": st["pending_updates"], "inbox": st["inbox"],
                "kakao": st["kakao"], "band": st["band_auth"], "demo": False, "tunnel": tunnel,
                "sources": srcs, "build": build_id()}
    except Exception as e:
        return {"error": str(e)}


def latest_reports():
    out = []
    for pat, name in [("종합리포트_*.md", "종합"), ("카톡대조_*.md", "카톡"), ("밴드대조_*.md", "밴드"),
                      ("ERP원장대조_*.md", "ERP원장"), ("이카운트대조_*.md", "판매·계산서")]:
        fs = sorted(glob.glob(os.path.join(ROOT, "reports", pat)))
        if fs:
            out.append({"kind": name, "file": os.path.basename(fs[-1]),
                        "text": open(fs[-1], encoding="utf-8").read()[:20000]})
    if DEMO and not out:
        out = [{"kind": "종합", "file": "demo.md",
                "text": "# 데모 리포트\n\n| 단계 | 결과 |\n|---|---|\n| 합성검증 | ✅ |\n| 카톡 대조 | ✅ |\n\n## 문제 예시\n- JS-2607-002 금액불일치(원장 620,000 / EC 500,000)"}]
    return out


# ───────────────────────── HTTP ─────────────────────────
class H(BaseHTTPRequestHandler):
    def handle_one_request(self):
        # 어떤 예외도 소켓을 조용히 죽이지 않게(ERR_EMPTY_RESPONSE 방지) 전역 가드
        try:
            super().handle_one_request()
        except (ConnectionError, BrokenPipeError):
            pass
        except Exception as e:
            try:
                self._send(500, {"error": str(e)[:300]})
            except Exception:
                pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.end_headers()
        self.wfile.write(data)

    def _auth(self):
        # 잠금 카운트는 /api/login 에서만 증가 — 구 PIN이 저장된 브라우저의
        # 자동 폴링이 잠금을 유발하던 문제(자기 잠금) 방지
        if _locked(self.client_address[0]):
            return False
        return self.headers.get("X-Pin", "") == PIN

    def do_OPTIONS(self):
        """브라우저 수집기(band.us 페이지)에서 보내는 사전 요청 허용 — 로컬에서만 쓴다"""
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Pin")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        # Chrome의 Private Network Access: 공개 사이트(https)에서 로컬 주소로 보낼 때 필요
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html"):
            html = open(os.path.join(BASE, "index.html"), encoding="utf-8").read()
            # ★★ 터널 주소로 들어온 경우에는 **설치 가능하게 만들지 않는다**.
            #   터널 호스트는 띄울 때마다 바뀌는데, 여기서 [설치]를 하면 그 임시 호스트가
            #   앱 아이콘에 **영구히 박힌다**. 주소가 바뀌는 순간 그 아이콘은 영영
            #   'ERR_FAILED'만 띄운다 — PC·폰 둘 다 그렇게 죽었다(2026-07-28 실사고).
            #   설치는 오직 고정 주소에서만 되게 하고, 여기서는 매니페스트·서비스워커를 뺀다.
            host = (self.headers.get("Host") or "").lower()
            if "trycloudflare.com" in host:
                html = html.replace('<link rel="manifest" href="/manifest.json">', "")
                html = html.replace("navigator.serviceWorker.register('/sw.js')",
                                    "Promise.reject()")
            return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
        if p.startswith("/brand/"):                    # 고객사 CI(쿠팡 로고) — 로컬 파일만 서빙
            fn = os.path.basename(p)
            fp = os.path.join(BASE, "brand", fn)
            ext = os.path.splitext(fn)[1].lower()
            ct = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                  ".svg": "image/svg+xml", ".webp": "image/webp"}.get(ext)
            if not ct or not os.path.exists(fp):
                return self._send(404, {"error": "no brand asset"})
            return self._send(200, open(fp, "rb").read(), ct)
        if p == "/api/brief":
            # 대표 보고용 '내용' 브리핑. 화면·PC 리포트·폰 사본이 **같은 문장**을 쓰도록
            # daily_brief 하나만 출처로 삼는다(따로 만들면 숫자가 갈린다).
            try:
                import daily_brief as DB
                day = None
                m = re.search(r"[?&]date=(\d{4}-\d{2}-\d{2})", self.path)
                if m:
                    day = m.group(1)
                b = DB.brief(day, DB.load()[0])
                return self._send(200, {"ok": True, "text": DB.text(b), **b})
            except Exception as e:
                return self._send(200, {"ok": False, "error": str(e)[:200]})
        if p == "/api/codes":
            return self._send(200, get_codes())
        if p == "/api/brand":
            return self._send(200, {"logo": brand_logo()})
        if re.fullmatch(r"/icon(?:-\d+)?\.(svg|png)", p):      # 아이콘(벡터/래스터 공용)
            try:
                return self._send(200, open(os.path.join(BASE, p.lstrip("/")), "rb").read(),
                                  "image/svg+xml" if p.endswith(".svg") else "image/png")
            except Exception:
                return self._send(404, {"error": "no icon"})
        if p == "/sw.js":
            # 크롬이 [설치 및 바로가기 만들기]로 진짜 앱 설치를 해 주려면
            # fetch 핸들러를 가진 서비스 워커가 필요하다. 캐시는 하지 않는다 —
            # 이 앱은 매일 바뀌는 실데이터를 보여주므로 옛 화면이 남으면 안 된다.
            js = ("self.addEventListener('install',e=>self.skipWaiting());\n"
                  "self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));\n"
                  "self.addEventListener('fetch',e=>{e.respondWith(fetch(e.request));});\n")
            # ★ _send 는 str을 받으면 JSON으로 감싼다 — 반드시 bytes로 넘겨야 스크립트가 된다
            return self._send(200, js.encode("utf-8"), "application/javascript; charset=utf-8")
        if p == "/manifest.json":                      # 홈 화면에 추가 시 앱처럼 보이게
            # ★ start_url 은 **이 페이지와 같은 출처**여야 한다. 다른 도메인을 넣으면
            #   크롬이 매니페스트를 통째로 무시해 [설치 및 바로가기 만들기]가 먹통이 된다
            #   (2026-07-27에 고정 주소를 넣었다가 실제로 설치가 안 됐다).
            #   그래서 여기는 "/" 로 두고, **오래 쓸 아이콘은 고정 주소에서 설치**한다.
            #   터널 주소로 들어온 사람에게는 index.html이 배너로 그 사실을 알린다.
            return self._send(200, {
                "name": "Coupang Service Operations System", "short_name": "CSOS",
                "start_url": "/", "scope": "/", "display": "standalone",
                "background_color": "#060D2B", "theme_color": "#060D2B",
                "icons": [
                    {"src": "/icon.svg", "sizes": "any", "type": "image/svg+xml", "purpose": "any"},
                    {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
                    {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
                    {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}]},
                "application/manifest+json")
        if p == "/api/ping":
            return self._send(200, {"app": "coupang-work", "demo": DEMO, "build": build_id()})
        if not self._auth():
            return self._send(401, {"error": "PIN"})
        if p == "/api/status":
            return self._send(200, get_status())
        if p == "/api/settlements":
            return self._send(200, {"rows": get_settlements()})
        if p == "/api/works":
            return self._send(200, get_works())
        if p == "/api/issues":
            return self._send(200, get_issues())
        if p == "/api/exec_report":
            return self._send(200, get_exec_report())
        if p == "/api/erpdocs":
            return self._send(200, get_erpdocs())
        if p == "/api/checks":
            return self._send(200, get_checks())
        if p == "/api/reports":
            return self._send(200, {"reports": latest_reports()})
        if p == "/api/tasklog":
            return self._send(200, {"busy": runner["busy"], "task": runner["task"],
                                    "log": list(runner["log"])[-300:]})
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        p = self.path.split("?")[0]
        ip = self.client_address[0]
        if p == "/api/login":
            if _locked(ip):
                return self._send(429, {"ok": False, "error": "시도 초과 — 10분 후 다시"})
            ln = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(ln) or b"{}")
            ok = body.get("pin", "") == PIN
            (_ok_login if ok else _fail)(ip)
            return self._send(200 if ok else 401, {"ok": ok})
        if p == "/api/band_dump":
            return self._band_dump()
        if not self._auth():
            return self._send(401, {"error": "PIN"})
        if p == "/api/enqueue":
            # 폰이 **PC 꺼진 동안 예약해 둔** 프로젝트 코드를 받아 원장에 등록한다.
            # 오프라인 앱이 PC가 살아난 걸 확인하는 즉시 스스로 보낸다(사람 개입 없음).
            ln = int(self.headers.get("Content-Length", 0))
            codes = (json.loads(self.rfile.read(ln) or b"{}").get("codes") or [])[:50]
            if DEMO:
                return self._send(200, {"ok": True, "applied": 0, "msg": "데모"})
            return self._send(200, enqueue_codes(codes))
        m = re.match(r"^/api/run/(\w+)$", p)
        if m:
            ok, msg = start_task(m.group(1))
            return self._send(200 if ok else 409, {"ok": ok, "msg": msg})
        if p == "/api/open":
            # 워크벤치 대체: 관리대장·폴더 열기. 원격(터널)에서는 의미가 없고 위험하므로
            # 서버가 도는 PC에서 접속했을 때만 허용한다.
            if ip not in ("127.0.0.1", "::1", "localhost"):
                return self._send(403, {"ok": False, "error": "이 기능은 사무실 PC에서만 됩니다"})
            ln = int(self.headers.get("Content-Length", 0))
            what = json.loads(self.rfile.read(ln) or b"{}").get("what", "")
            try:
                from ecount_reconcile import load_config, resolve_master
                master = resolve_master(load_config()["reconcile"]["master_xlsx"])
            except Exception:
                master = ""
            targets = {"master": master, "master_dir": os.path.dirname(master),
                       "inbox": os.path.join(ROOT, "inbox"),
                       "kakao": os.path.join(ROOT, "kakao", "inbox"),
                       "band_docs": os.path.join(ROOT, "band", "docs_inbox"),
                       "band_cache": os.path.join(ROOT, "band", "cache"),
                       "reports": os.path.join(ROOT, "reports")}
            path = targets.get(what)
            if not path or not os.path.exists(path):
                return self._send(404, {"ok": False, "error": f"경로 없음: {what}"})
            try:
                os.startfile(path)
                return self._send(200, {"ok": True, "opened": os.path.basename(path) or path})
            except Exception as e:
                return self._send(500, {"ok": False, "error": str(e)[:200]})
        if p == "/api/set_dates":
            # 보고일·집계기준일 → 00_대시보드 B3·B4 (일일 갱신 입력칸 — 덮어쓰기 허용 화이트리스트)
            ln = int(self.headers.get("Content-Length", 0))
            b = json.loads(self.rfile.read(ln) or b"{}")
            items = []
            for cell, key, label in (("B3", "보고일", "보고일"), ("B4", "집계기준일", "집계기준일")):
                v = str(b.get(key, "")).strip()
                if v:
                    if not re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                        return self._send(400, {"ok": False, "error": f"{label} 형식 오류(YYYY-MM-DD)"})
                    items.append({"sheet": "00_대시보드", "cell": cell, "key": cell, "key_col": "-",
                                  "col": label, "value": v, "vtype": "date",
                                  "evidence": f"앱 기준일 설정({ip})", "only_if_empty": False})
            if not items:
                return self._send(400, {"ok": False, "error": "날짜 없음"})
            if DEMO:
                return self._send(200, {"ok": True, "demo": True})
            from ledger_writer import queue_add
            queue_add(items)
            ok, msg = start_task("writer_apply")     # 즉시 반영(vN+1)
            return self._send(200, {"ok": True, "applying": ok, "msg": msg})
        if p == "/api/input":
            # 앱 → 엑셀 입력: ledger_writer 큐에 적재(빈 칸만 정책은 반영 단계에서 강제)
            ln = int(self.headers.get("Content-Length", 0))
            b = json.loads(self.rfile.read(ln) or b"{}")
            ALLOW = {"02_돌발AS접수", "04_정기점검", "06_거래서류청구수금",
                     "15_세금계산서관리", "16_입금수금관리"}
            if b.get("sheet") not in ALLOW or not b.get("key") or not b.get("col") or b.get("value") in (None, ""):
                return self._send(400, {"ok": False, "error": "sheet/key/col/value 필요"})
            if b.get("vtype") not in ("text", "date", "number"):
                b["vtype"] = "text"
            if DEMO:
                return self._send(200, {"ok": True, "queued": 1, "demo": True})
            from ledger_writer import queue_add, load_queue
            n = queue_add([{"sheet": b["sheet"], "key_col": b.get("key_col", "정산ID"), "key": b["key"],
                            "col": b["col"], "value": b["value"], "vtype": b["vtype"],
                            "evidence": f"앱 입력({ip}) {datetime.now():%m-%d %H:%M}", "only_if_empty": True}])
            return self._send(200, {"ok": True, "queued": n, "pending": len(load_queue())})
        return self._send(404, {"error": "not found"})

    def _band_dump(self):
        """브라우저 수집기가 밴드 게시글 원본을 직접 전송(no-cors POST) — PIN 쿼리로 보호"""
        from urllib.parse import parse_qs, urlparse
        q = parse_qs(urlparse(self.path).query)
        if (q.get("pin") or [""])[0] != PIN:
            return self._send(401, {"ok": False})
        ln = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(ln)
        try:
            d = json.loads(raw.decode("utf-8"))
            band = re.sub(r"\D", "", str(d.get("band", ""))) or "unknown"
            os.makedirs(os.path.join(ROOT, "band", "cache"), exist_ok=True)
            path = os.path.join(ROOT, "band", "cache", f"dump_{band}.json")
            open(path, "w", encoding="utf-8").write(json.dumps(d, ensure_ascii=False))
            return self._send(200, {"ok": True, "saved": len(d.get("posts", {}))})
        except Exception as e:
            return self._send(400, {"ok": False, "error": str(e)[:200]})

    def log_message(self, *a):
        pass


def lan_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


PUBLISH_EVERY = 3 * 3600      # 폰이 보는 사본을 몇 초마다 새로 올릴지


def publish_loop():
    """PC가 켜져 있는 동안 **주기적으로** 폰용 사본을 올린다.

    ★ PC가 꺼져도 폰이 쓰이려면 사본이 최신이어야 한다. 예전에는 daily_run 이 돌 때만
      올려서, 아침에 한 번 돌리고 저녁에 PC를 끄면 폰은 **아침 숫자**를 보게 됐다.
      그러면 '꺼져도 된다'는 말이 사실이 아니게 된다. 그래서 3시간마다 올린다.
      (사본은 잠겨 있고 60~80KB라 부담이 없다)
    """
    if DEMO:
        return
    time.sleep(120)                        # 기동 직후 혼잡할 때는 피한다
    while True:
        from operation_window import is_input_window
        if is_input_window():
            time.sleep(60)
            continue
        try:
            r = subprocess.run([PY, os.path.join(ROOT, "cloud_publish.py"), "--push"],
                               cwd=ROOT, env=ENV, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=900)
            tail = [l for l in (r.stdout or "").splitlines() if l.strip()][-1:] or [""]
            runner["log"].append(f"[사본 자동 게시] {tail[0][:120]}")
        except Exception as e:
            runner["log"].append(f"[사본 자동 게시] 실패 {type(e).__name__}")
        time.sleep(PUBLISH_EVERY)


class _Server(ThreadingHTTPServer):
    """★ 사고 #16의 진짜 원인 — 윈도우에서 SO_REUSEADDR 이 켜져 있으면 **이미 남이 쓰는
    포트에도 바인드가 성공한다.** 그래서 새 서버가 '시작됨'을 찍고도 요청은 계속 옛
    프로세스가 받아 갔다. 코드를 고쳐도 화면이 안 바뀌는 증상의 정체가 이것이다.
    재사용을 끄면 두 번째 서버는 조용히 뜨는 대신 **에러로 죽는다** — 그게 옳다."""
    allow_reuse_address = False


def main():
    try:
        srv = _Server(("0.0.0.0", PORT), H)
    except OSError:
        print(f"★ {PORT} 포트를 이미 다른 앱 서버가 쓰고 있습니다. 새로 뜨지 않았습니다.")
        print("  옛 서버가 계속 응답하므로 **코드를 고쳐도 화면이 안 바뀝니다.**")
        print("  먼저 정리하세요 (PowerShell):")
        print("    Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" |"
              " ? { $_.CommandLine -like '*app_server.py*' } | % { Stop-Process -Id $_.ProcessId -Force }")
        return 1
    threading.Thread(target=publish_loop, daemon=True).start()
    mode = "데모(합성데이터)" if DEMO else "실서비스"
    print(f"Coupang Service Operations System 앱 서버 [{mode}] 시작")
    print(f"  PC:      http://localhost:{PORT}")
    print(f"  휴대폰:  http://{lan_ip()}:{PORT}   (같은 와이파이)")
    print(f"  PIN:     {PIN}")
    srv.serve_forever()


if __name__ == "__main__":
    main()
