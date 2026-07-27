# -*- coding: utf-8 -*-
"""
cloud_server.py — 클라우드에서 24시간 뜨는 앱 (PC가 꺼져도 폰에서 열린다)
================================================================================
사내 app_server.py 와 **같은 화면(index.html)** 을 쓰되, 엑셀 대신
cloud_export.py 가 올려 준 데이터 한 덩어리(bundle.json)만 읽는다.
Z: 드라이브도, 이카운트도, 엑셀도 이 서버에는 없다.

  사내 PC ──(매일 09:50, HTTPS PUSH)──▶ 클라우드  ──▶ 폰(24시간 조회)
                    ▲                                    │
                    └────(입력 큐를 PC가 가져가 반영)──────┘

할 수 있는 것 : 조회 전부(정산·업무·확인필요·대표보고·계산서구성)
할 수 없는 것 : 대조 실행·엑셀 직접 수정 → 이건 PC가 한다.
                폰에서 넣은 입력은 **큐에 쌓였다가** PC가 가져가 관리대장에 반영한다.

보안 : 조회는 PIN, 업로드/큐 회수는 토큰(X-Token). 둘 다 환경변수로 준다.
  CSOS_PIN=1234  CSOS_TOKEN=긴-임의-문자열  PORT=8080  python webapp/cloud_server.py
"""
import sys, os, json, gzip, time, threading, hmac
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.environ.get("CSOS_DATA") or os.path.join(BASE, "data")
os.makedirs(DATA, exist_ok=True)
BUNDLE = os.path.join(DATA, "bundle.json")
QUEUE = os.path.join(DATA, "input_queue.json")

PIN = os.environ.get("CSOS_PIN", "")
TOKEN = os.environ.get("CSOS_TOKEN", "")
PORT = int(os.environ.get("PORT", "8080"))
if not PIN or not TOKEN:
    sys.exit("CSOS_PIN 과 CSOS_TOKEN 환경변수를 설정하세요(둘 다 필수).")

_lock = threading.Lock()
_cache = {"mtime": 0, "data": {}}
_fails = {}


def bundle():
    """올라온 데이터. 파일이 바뀌면 다시 읽는다."""
    try:
        mt = os.path.getmtime(BUNDLE)
    except OSError:
        return {}
    if _cache["mtime"] != mt:
        with _lock:
            try:
                _cache["data"] = json.load(open(BUNDLE, encoding="utf-8"))
                _cache["mtime"] = mt
            except Exception:
                pass
    return _cache["data"]


def locked(ip):
    c, until = _fails.get(ip, (0, 0))
    return c >= 5 and time.time() < until


class H(BaseHTTPRequestHandler):
    server_version = "CSOS"

    def log_message(self, *a):
        pass

    # ── 응답 도우미 ──────────────────────────────────────────────
    def send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        # 사내 자료다 — 검색엔진·캐시에 남기지 않는다
        self.send_header("X-Robots-Tag", "noindex, nofollow")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        self.wfile.write(body)

    def authed(self):
        return self.headers.get("X-Pin", "") == PIN

    def tokened(self):
        got = self.headers.get("X-Token", "")
        return bool(got) and hmac.compare_digest(got, TOKEN)

    # ── 조회 ────────────────────────────────────────────────────
    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            f = os.path.join(BASE, "index.html")
            return self.send(200, open(f, "rb").read(), "text/html; charset=utf-8")
        if path == "/api/ping":
            b = bundle()
            return self.send(200, {"ok": True, "생성": b.get("생성", ""),
                                   "원본": b.get("원본", ""), "mode": "cloud"})
        if not path.startswith("/api/"):
            return self.send(404, {"error": "not found"})
        if not self.authed():
            return self.send(401, {"error": "PIN"})

        b = bundle()
        key = path[len("/api/"):]
        if key in ("settlements", "works", "issues", "exec_report", "erpdocs", "checks"):
            return self.send(200, b.get(key) or {})
        if key == "status":
            return self.send(200, {"mode": "cloud", "생성": b.get("생성", ""),
                                   "원본": b.get("원본", ""),
                                   "안내": "클라우드 사본입니다. 대조 실행·엑셀 수정은 사내 PC에서 합니다.",
                                   "steps": [], "busy": False})
        if key == "brand":
            return self.send(200, {"logo": ""})
        if key in ("reports", "tasklog"):
            return self.send(200, {"rows": [], "안내": "리포트는 사내 PC에서만 봅니다."})
        return self.send(200, {})

    # ── 입력(큐에 쌓아 두고 PC가 가져간다) ────────────────────────
    def do_POST(self):
        path = self.path.split("?")[0]
        ip = self.client_address[0]
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n) if n else b""
        if self.headers.get("Content-Encoding") == "gzip":
            try:
                raw = gzip.decompress(raw)
            except Exception:
                return self.send(400, {"error": "gzip 해제 실패"})

        if path == "/api/push":                       # 사내 PC → 클라우드
            if not self.tokened():
                return self.send(401, {"error": "token"})
            try:
                json.loads(raw.decode("utf-8"))        # 형식 확인 후에만 덮어쓴다
            except Exception:
                return self.send(400, {"error": "JSON 아님"})
            tmp = BUNDLE + ".tmp"
            open(tmp, "wb").write(raw)
            os.replace(tmp, BUNDLE)
            return self.send(200, {"ok": True, "bytes": len(raw)})

        if path == "/api/pull_queue":                  # 클라우드 → 사내 PC
            if not self.tokened():
                return self.send(401, {"error": "token"})
            q = []
            with _lock:
                if os.path.exists(QUEUE):
                    try:
                        q = json.load(open(QUEUE, encoding="utf-8"))
                    except Exception:
                        q = []
                    os.remove(QUEUE)                   # 가져간 건 비운다
            return self.send(200, {"items": q})

        if path == "/api/login":
            if locked(ip):
                return self.send(429, {"error": "잠시 후 다시"})
            try:
                pin = json.loads(raw.decode("utf-8")).get("pin", "")
            except Exception:
                pin = ""
            if hmac.compare_digest(str(pin), PIN):
                _fails.pop(ip, None)
                return self.send(200, {"ok": True})
            c, _ = _fails.get(ip, (0, 0))
            _fails[ip] = (c + 1, time.time() + 600)
            return self.send(401, {"error": "PIN"})

        if not self.authed():
            return self.send(401, {"error": "PIN"})

        if path == "/api/input":                       # 폰에서 넣은 값
            try:
                item = json.loads(raw.decode("utf-8"))
            except Exception:
                return self.send(400, {"error": "JSON 아님"})
            item["보낸시각"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            with _lock:
                q = []
                if os.path.exists(QUEUE):
                    try:
                        q = json.load(open(QUEUE, encoding="utf-8"))
                    except Exception:
                        q = []
                q.append(item)
                json.dump(q, open(QUEUE, "w", encoding="utf-8"), ensure_ascii=False)
            return self.send(200, {"ok": True, "대기": len(q),
                                   "안내": "사내 PC가 다음 실행 때 관리대장에 반영합니다."})

        # 대조 실행 같은 건 클라우드에서 할 수 없다 — 조용히 실패시키지 않고 이유를 준다
        return self.send(400, {"error": "클라우드에서는 실행할 수 없습니다",
                               "안내": "대조·엑셀 수정은 사내 PC에서 합니다."})


def main():
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), H)
    b = bundle()
    print(f"CSOS 클라우드 · 포트 {PORT} · 데이터 {DATA}")
    print("올라온 데이터:", b.get("생성", "(아직 없음)"), b.get("원본", ""))
    srv.serve_forever()


if __name__ == "__main__":
    main()
