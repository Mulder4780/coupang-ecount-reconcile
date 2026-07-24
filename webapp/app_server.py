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


def demo_settlements():
    camps = ["송파5MB(감일동)", "울산2캠프", "인천7MB(마곡동)", "부천3(BUC3)", "대전1캠프",
             "구리1캠프", "제주1Sub-hub", "창원1MB(팔용동)", "군포1Sub-Hub", "광주2Sub-hub"]
    rows = []
    rnd = random.Random(42)
    for i in range(1, 16):
        amt = rnd.choice([380000, 418000, 470800, 760000, 1230000, 1472500])
        st = rnd.choice(["정상", "세금계산서 미발행", "ERP 미확인", "미청구", "입금 대기"])
        rows.append({"정산ID": f"JS-2607-{i:03d}", "업무구분": rnd.choice(["돌발AS", "정기점검"]),
                     "캠프명": camps[i % len(camps)], "프로젝트NO": f"UJ26{1000+i}",
                     "공급가액": amt, "합계": int(amt * 1.1),
                     "명세서": "있음" if st != "미청구" else "없음",
                     "계산서": "발행" if st == "정상" else "미발행",
                     "상태": st, "완료일": (date(2026, 7, 1) + timedelta(days=i)).isoformat()})
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
                     "프로젝트NO": r.get("프로젝트NO"),
                     "공급가액": r.get("원장_공급가액") or 0, "합계": r.get("원장_합계") or 0,
                     "명세서": "있음" if has_stmt else "없음",
                     "계산서": "발행" if issued else "미발행",
                     "상태": st, "완료일": str(r.get("작업완료일") or "")[:10]})
    return rows


def get_settlements():
    if DEMO:
        return demo_settlements()
    if time.time() - _cache["t"] < 60 and _cache["settle"]:
        return _cache["settle"]
    rows = real_settlements()
    _cache["settle"], _cache["t"] = rows, time.time()
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
        steps = []
        for s in st.get("report_summary", []):
            mark = "ok" if "✅" in s else ("skip" if "스킵" in s else "fail")
            steps.append({"n": re.sub(r"[✅❌⏭]|스킵|실패", "", s).strip(), "s": mark})
        rt = st.get("report_time", "")
        if rt:
            rt = f"{rt[:4]}-{rt[4:6]}-{rt[6:8]} {rt[9:11]}:{rt[11:13]}"
        return {"master": os.path.basename(st.get("master", "") or "") + "  " + st.get("master_label", ""),
                "fork": st.get("fork", []), "agent_last": rt or "기록 없음", "steps": steps,
                "pending_updates": st["pending_updates"], "inbox": st["inbox"],
                "kakao": st["kakao"], "band": st["band_auth"], "demo": False}
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
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _auth(self):
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
        if p == "/api/reports":
            return self._send(200, {"reports": latest_reports()})
        if p == "/api/tasklog":
            return self._send(200, {"busy": runner["busy"], "task": runner["task"],
                                    "log": list(runner["log"])[-300:]})
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        p = self.path.split("?")[0]
        if p == "/api/login":
            ln = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(ln) or b"{}")
            ok = body.get("pin", "") == PIN
            return self._send(200 if ok else 401, {"ok": ok})
        if not self._auth():
            return self._send(401, {"error": "PIN"})
        m = re.match(r"^/api/run/(\w+)$", p)
        if m:
            ok, msg = start_task(m.group(1))
            return self._send(200 if ok else 409, {"ok": ok, "msg": msg})
        return self._send(404, {"error": "not found"})

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
