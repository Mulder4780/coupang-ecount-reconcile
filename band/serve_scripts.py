# -*- coding: utf-8 -*-
"""serve_scripts.py — 수집 스크립트를 밴드 탭이 **직접 받아 가게** 하는 임시 서버

왜 필요한가 (2026-08-06)
  수집기(grab_posts.js·grab_all.js)는 9KB 가 넘는다. 이것을 자동화 도구의 인자로
  매번 밀어 넣으면 대화가 그만큼 부풀고, 밤새 여러 번 다시 주입할 때마다 또 붙는다.
  그래서 로컬에 잠깐 띄워 두고 밴드 탭이 `fetch` 로 받아 가게 한다 —
  주입 명령이 두 줄로 줄어든다.

왜 http.server 를 그냥 안 쓰나
  band.us(https) 에서 부르는 요청이라 **CORS 헤더가 없으면 브라우저가 막는다.**
  (localhost 는 mixed content 예외라 http 여도 요청 자체는 나간다)

  python band/serve_scripts.py            # 8123 포트
  python band/serve_scripts.py --port 9123
"""
import os
import sys
from http.server import HTTPServer, SimpleHTTPRequestHandler

HERE = os.path.dirname(os.path.abspath(__file__))
ALLOW = {"grab_posts.js", "grab_all.js", "collect_band.js"}


class H(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def do_GET(self):
        name = os.path.basename(self.path.split("?")[0])
        if name not in ALLOW:                      # 이 폴더에는 캐시·토큰도 있다
            self.send_error(404, "not shared")
            return
        try:
            body = open(os.path.join(HERE, name), "rb").read()
        except OSError:
            self.send_error(404, "no file")
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/javascript; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):                     # 콘솔을 조용히
        pass


def main():
    port = 8123
    if "--port" in sys.argv:
        port = int(sys.argv[sys.argv.index("--port") + 1])
    print(f"수집 스크립트 서버: http://localhost:{port}/grab_posts.js  (Ctrl+C 로 종료)")
    HTTPServer(("127.0.0.1", port), H).serve_forever()


if __name__ == "__main__":
    main()
