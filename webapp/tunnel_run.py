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


MAX_PER_HOUR = 4
STAMP = os.path.join(ROOT, "reports", "tunnel_starts.txt")


def _throttle():
    """1시간에 MAX_PER_HOUR번까지만 새 터널을 만든다.

    주소를 새로 받을 때마다 폰·PC 북마크가 가리키는 곳이 바뀐다. 자주 만들면
    고정 주소가 따라가기도 전에 또 바뀌고, 심하면 **DNS에 등록되지 않은 주소**를
    받는다(2026-07-27 실사고). 한도를 넘으면 가장 오래된 기록이 1시간을 넘길 때까지 기다린다.
    """
    now = time.time()
    try:
        hist = [float(x) for x in open(STAMP, encoding="utf-8").read().split()]
    except Exception:
        hist = []
    hist = [t for t in hist if now - t < 3600]
    while len(hist) >= MAX_PER_HOUR:
        wait = int(3600 - (now - hist[0])) + 5
        print(f"터널을 너무 자주 새로 만들고 있습니다 — {wait}초 기다립니다"
              f" (최근 1시간 {len(hist)}회)")
        time.sleep(min(wait, 300))
        now = time.time()
        hist = [t for t in hist if now - t < 3600]
    hist.append(now)
    try:
        os.makedirs(os.path.dirname(STAMP), exist_ok=True)
        open(STAMP, "w", encoding="utf-8").write("\n".join(str(t) for t in hist))
    except OSError:
        pass


def publish(url):
    """고정 주소(GitHub Pages)가 지금 주소를 가리키게 한다.

    폰 북마크는 고정 주소 하나만 알면 되고, 실제 주소가 바뀌어도 여기서 따라간다.
    (trycloudflare 무료 터널은 **주소를 지정할 수 없다** — 매번 새로 받는다)
    """
    try:
        subprocess.run([sys.executable, os.path.join(ROOT, "publish_endpoint.py")],
                       cwd=ROOT, timeout=120, capture_output=True,
                       env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        print("   고정 주소 갱신 완료")
    except Exception as e:
        print(f"   고정 주소 갱신 실패: {type(e).__name__} {e}")


def watch(proc, url):
    """주소가 살아 있는지 1분마다 확인해 죽으면 cloudflared를 내린다.

    ★ 이게 없으면 **cloudflared 프로세스는 살아 있는데 주소만 만료된** 상태로 방치된다.
      바깥에서는 '사이트에 연결할 수 없습니다'인데 안에서는 정상으로 보여 아무도 못 고친다
      (2026-07-27 실사고 — 이틀 동안 접속 불가였다).
      프로세스를 내리면 바깥 while 루프가 새 주소로 다시 띄운다.
    """
    import threading, urllib.request

    def ping(u, t=20):
        try:
            with urllib.request.urlopen(u, timeout=t) as r:
                return r.status == 200
        except Exception:
            return False

    def loop():
        fail = 0
        while proc.poll() is None:
            time.sleep(90)
            if ping(url + "/api/ping"):
                fail = 0
                continue
            # 앱 자체가 죽었으면 터널을 갈아도 소용없다 — 주소만 쓸데없이 바뀐다.
            # (주소가 바뀔 때마다 직접 북마크한 사람은 다시 못 들어온다)
            if not ping(f"http://localhost:{PORT}/api/ping", 8):
                print("앱이 응답하지 않음 — 터널은 그대로 두고 기다립니다")
                fail = 0
                continue
            fail += 1
            print(f"터널 응답 없음 {fail}/3")
            if fail >= 3:                      # 90초 간격 3회(=약 4분) 연속 실패에만 교체
                print("터널이 죽은 것으로 판단 — 새 주소로 다시 띄웁니다")
                try:
                    proc.kill()
                except Exception:
                    pass
                return
    threading.Thread(target=loop, daemon=True).start()


def main():
    # 싱글톤 락: 이미 다른 tunnel_run이 돌고 있으면 즉시 종료(중복 터널·주소 회전 방지)
    import socket
    global _lock_sock
    _lock_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        _lock_sock.bind(("127.0.0.1", 8977))
    except OSError:
        print("이미 실행 중 — 종료")
        return
    exe = find_cloudflared()
    if not exe:
        sys.exit("cloudflared 미설치. 먼저:  winget install Cloudflare.cloudflared")
    # 터널을 너무 자주 새로 만들면 주소만 계속 바뀌고(폰 북마크가 무용지물),
    # 새로 받은 주소가 아예 등록되지 않는 일도 생긴다(2026-07-27에 한 시간에 12개를
    # 만들다가 DNS에 없는 주소를 받았다). 1시간에 4번까지만 새로 만든다.
    _throttle()
    os.makedirs(os.path.dirname(URL_FILE), exist_ok=True)
    while True:
        print(f"터널 시작... (대상 http://localhost:{PORT})")
        p = subprocess.Popen([exe, "tunnel", "--url", f"http://localhost:{PORT}"],
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace")
        url = None
        for ln in p.stdout:
            # ★ cloudflared 로그에는 우리 터널 주소 말고 **api.trycloudflare.com**(내부 API)도
            #   나온다. 그냥 잡으면 그게 먼저 걸려 고정 주소가 엉뚱한 곳을 가리키고,
            #   폰에서는 '사이트에 연결할 수 없습니다'만 뜬다(2026-07-27 실사고).
            #   무료 터널 주소는 언제나 '단어-단어-단어-단어' 꼴이라 하이픈으로 걸러낸다.
            m = re.search(r"(https://[a-z0-9]+(?:-[a-z0-9]+){2,}\.trycloudflare\.com)", ln)
            if m and not url:
                url = m.group(1)
                open(URL_FILE, "w", encoding="utf-8").write(url)
                print(f"★ 외부 접속 주소: {url}  (앱 대시보드에도 표시됨)")
                publish(url)      # 고정 주소(GitHub Pages)가 새 주소를 가리키게
                watch(p, url)     # 주소가 죽으면 감지해 cloudflared를 내린다
        p.wait()
        open(URL_FILE, "w", encoding="utf-8").write("")
        print("터널 종료됨 — 5초 후 재시작")
        time.sleep(5)


if __name__ == "__main__":
    main()
