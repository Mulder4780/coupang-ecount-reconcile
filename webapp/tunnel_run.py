# -*- coding: utf-8 -*-
"""
tunnel_run.py — 외부 접속 터널 (Cloudflare Quick Tunnel)
==========================================================
휴대폰이 와이파이 밖(LTE/5G)에서도 접속하도록 공개 HTTPS 주소를 만든다.
포트포워딩·계정 불필요. 주소는 터널을 새로 시작할 때마다 바뀌며
reports/tunnel_url.txt 에 저장되어 앱 대시보드에 표시된다.

사전 1회: winget install Cloudflare.cloudflared
실행:     python webapp/tunnel_run.py   (터널이 죽으면 자동 재시작)

★ 보안: 공개 주소 + PIN 4자리 구조다. 로그인 5회 실패 시 10분 잠금이 있지만,
  장기적으로는 Tailscale(사설망) 방식이 더 안전하다 — README 참고.
"""
import sys, os, re, time, subprocess, shutil

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URL_FILE = os.path.join(ROOT, "reports", "tunnel_url.txt")
PORT = 8899


def find_cloudflared():
    # 1순위: 포터블(webapp/cloudflared.exe — 관리자 권한 불필요)
    local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cloudflared.exe")
    if os.path.exists(local):
        return local
    p = shutil.which("cloudflared")
    if p:
        return p
    for c in [os.path.expandvars(r"%ProgramFiles(x86)%\cloudflared\cloudflared.exe"),
              os.path.expandvars(r"%ProgramFiles%\cloudflared\cloudflared.exe"),
              os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WinGet\Links\cloudflared.exe")]:
        if os.path.exists(c):
            return c
    return None


def main():
    exe = find_cloudflared()
    if not exe:
        sys.exit("cloudflared 미설치. 먼저:  winget install Cloudflare.cloudflared")
    os.makedirs(os.path.dirname(URL_FILE), exist_ok=True)
    while True:
        print(f"터널 시작... (대상 http://localhost:{PORT})")
        p = subprocess.Popen([exe, "tunnel", "--url", f"http://localhost:{PORT}"],
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace")
        url = None
        for ln in p.stdout:
            m = re.search(r"(https://[a-z0-9-]+\.trycloudflare\.com)", ln)
            if m and not url:
                url = m.group(1)
                open(URL_FILE, "w", encoding="utf-8").write(url)
                print(f"★ 외부 접속 주소: {url}  (앱 대시보드에도 표시됨)")
        p.wait()
        open(URL_FILE, "w", encoding="utf-8").write("")
        print("터널 종료됨 — 5초 후 재시작")
        time.sleep(5)


if __name__ == "__main__":
    main()
