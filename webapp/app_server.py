# -*- coding: utf-8 -*-
"""
app_server.py — 쿠팡 통합업무 앱 서버 (반응형 웹앱 백엔드)
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
}
runner = {"busy": False, "task": "", "log": deque(maxlen=3000), "done_at": None}
_rlock = threading.Lock()


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
_readlock = threading.Lock()   # Z:드라이브 엑셀 동시 읽기 직렬화(스레드 충돌 방지)


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
                     "공급가액": amt, "합계": int(amt * 1.1),
                     "명세서": "있음" if st != "미청구" else "없음",
                     "명세서번호": f"2026/07/{i:02d}-1" if st != "미청구" else "",
                     "명세서발행일": d if st != "미청구" else "",
                     "계산서": "발행" if st == "정상" else "미발행",
                     "계산서발행일": d if st == "정상" else "", "승인번호": "",
                     "입금일": d if st == "정상" else "", "입금액": int(amt * 1.1) if st == "정상" else 0,
                     "미수금": 0 if st == "정상" else int(amt * 1.1), "비용구분": "유상",
                     "상태": st, "완료일": d})
    return rows


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
    return rows


def real_works():
    """02 돌발AS·04 정기점검 현황 (앱 '업무' 데이터)"""
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    out = {"as": [], "pm": []}
    spec = {
        "02_돌발AS접수": ("as", ["접수ID", "프로젝트NO", "캠프명", "접수일자", "담당기사", "진행상태",
                                "작업완료일", "유상·무상·보험", "신청내용", "긴급도", "방문예정일"]),
        "04_정기점검": ("pm", ["점검ID", "프로젝트NO", "캠프명", "점검예정일", "실제점검일", "점검상태",
                              "담당기사", "이상발견여부", "돌발AS전환여부"]),
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
            rec = {}
            for c in cols:
                v = row[idx[c]] if c in idx and idx[c] < len(row) else None
                rec[c] = str(v)[:10] if hasattr(v, "year") else ("" if v is None else str(v))
            out[key].append(rec)
    wb.close()
    return out


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
        out = {"rows": rows, "cols": hdr, "source": "23_확인필요현황"}
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
    rows = merged + rows
    cols = []
    for r in rows[:50]:
        for k in r:
            if k not in cols:
                cols.append(k)
    out = {"rows": rows, "cols": cols}
    _cache["issues"] = out
    return out


def get_checks():
    """최근 카톡·밴드·ERP원장 대조 CSV를 ID별로 조인 — 3원천 검증 배지"""
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
    if DEMO and not out:
        out = {"JS-2607-002": {"kakao": "확인", "band": "미확인", "erp": "D 금액불일치"}}
    return out


def get_settlements():
    if DEMO:
        return demo_settlements()
    with _readlock:
        r = _fresh("settle")
        if r:
            return r
        rows = real_settlements()
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
        return {"master": os.path.basename(st.get("master", "") or "") + "  " + st.get("master_label", ""),
                "fork": st.get("fork", []), "agent_last": rt or "기록 없음", "steps": steps,
                "pending_updates": st["pending_updates"], "inbox": st["inbox"],
                "kakao": st["kakao"], "band": st["band_auth"], "demo": False, "tunnel": tunnel}
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
        self.end_headers()
        self.wfile.write(data)

    def _auth(self):
        # 잠금 카운트는 /api/login 에서만 증가 — 구 PIN이 저장된 브라우저의
        # 자동 폴링이 잠금을 유발하던 문제(자기 잠금) 방지
        if _locked(self.client_address[0]):
            return False
        return self.headers.get("X-Pin", "") == PIN

    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html"):
            html = open(os.path.join(BASE, "index.html"), encoding="utf-8").read()
            return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
        if p == "/api/ping":
            return self._send(200, {"app": "coupang-work", "demo": DEMO})
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
        m = re.match(r"^/api/run/(\w+)$", p)
        if m:
            ok, msg = start_task(m.group(1))
            return self._send(200 if ok else 409, {"ok": ok, "msg": msg})
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


def main():
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), H)
    mode = "데모(합성데이터)" if DEMO else "실서비스"
    print(f"쿠팡 통합업무 앱 서버 [{mode}] 시작")
    print(f"  PC:      http://localhost:{PORT}")
    print(f"  휴대폰:  http://{lan_ip()}:{PORT}   (같은 와이파이)")
    print(f"  PIN:     {PIN}")
    srv.serve_forever()


if __name__ == "__main__":
    main()
