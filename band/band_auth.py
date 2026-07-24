# -*- coding: utf-8 -*-
"""
band_auth.py — 네이버 밴드 Open API 최초 1회 인증 (사용자는 브라우저에서 [동의] 클릭만)
========================================================================================
흐름: 이 스크립트 실행 → 브라우저 자동 오픈(auth.band.us 동의화면) → 동의 클릭
      → localhost 콜백으로 code 수신 → access_token 교환 → .band_token.json 저장.

사전에 config/ecount_config.json 의 band.client_id / band.client_secret 이 채워져 있어야 함
(developers.band.us 에서 앱 등록 후 발급 — README '밴드 연동' 절 참고).
"""
import sys, os, json, base64, urllib.request, urllib.parse, webbrowser, ssl
from http.server import HTTPServer, BaseHTTPRequestHandler

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(os.path.dirname(BASE_DIR), "config", "ecount_config.json")
TOKEN_PATH = os.path.join(BASE_DIR, ".band_token.json")

cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
band = cfg.get("band", {})
CID = (band.get("client_id") or "").strip()
CSECRET = (band.get("client_secret") or "").strip()
PORT = int(band.get("redirect_port", 8976))
REDIRECT = f"http://localhost:{PORT}/callback"

if not CID or not CSECRET:
    sys.exit("band.client_id / band.client_secret 이 비어 있습니다. developers.band.us 앱 등록 후 config에 입력하세요.")

code_box = {}

class CB(BaseHTTPRequestHandler):
    def do_GET(self):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        code_box["code"] = (q.get("code") or [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write("<h2>인증 완료 — 이 창을 닫으셔도 됩니다.</h2>".encode("utf-8"))
    def log_message(self, *a):
        pass

auth_url = ("https://auth.band.us/oauth2/authorize?response_type=code"
            f"&client_id={CID}&redirect_uri={urllib.parse.quote(REDIRECT, safe='')}")
print("브라우저에서 밴드 로그인 후 [동의]를 눌러주세요...")
print("(자동으로 안 열리면 직접 열기):", auth_url)
webbrowser.open(auth_url)

srv = HTTPServer(("localhost", PORT), CB)
srv.handle_request()          # 동의 1회 수신까지 대기
srv.server_close()

code = code_box.get("code")
if not code:
    sys.exit("code 수신 실패 — 동의 화면에서 취소되었거나 redirect URI가 앱 설정과 다릅니다.")

token_url = f"https://auth.band.us/oauth2/token?code={urllib.parse.quote(code)}&grant_type=authorization_code"
basic = base64.b64encode(f"{CID}:{CSECRET}".encode()).decode()
req = urllib.request.Request(token_url, headers={"Authorization": f"Basic {basic}"})
with urllib.request.urlopen(req, timeout=30, context=ssl.create_default_context()) as r:
    tok = json.loads(r.read().decode("utf-8"))

if "access_token" not in tok:
    sys.exit(f"토큰 교환 실패: {json.dumps(tok, ensure_ascii=False)[:300]}")

json.dump(tok, open(TOKEN_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("access_token 저장 완료 →", TOKEN_PATH)
print("다음: python band_sync.py  (밴드 목록·게시글 자동 수집)")
