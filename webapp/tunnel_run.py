# -*- coding: utf-8 -*-
"""
tunnel_run.py — 외부 접속 터널 (Cloudflare Quick Tunnel)
==========================================================
휴대폰이 와이파이 밖(LTE/5G)에서도 접속하도록 공개 HTTPS 주소를 만든다.
포트포워딩·계정 불필요. 주소는 터널을 새로 시작할 때마다 바뀌며
reports/tunnel_url.txt 에 저장되어 앱 대시보드에 표시된다.

사전 1회: winget install Cloudflare.cloudflared
실행:     python webapp/tunnel_run.py   (자동으로 창 없는 pythonw 로 전환)
디버깅:   python webapp/tunnel_run.py --console

★ 보안: 공개 주소 + PIN 4자리 구조다. 로그인 5회 실패 시 10분 잠금이 있지만,
  장기적으로는 Tailscale(사설망) 방식이 더 안전하다 — README 참고.
"""
import sys, os, re, json, time, subprocess, shutil
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from operation_window import is_input_window

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URL_FILE = os.path.join(ROOT, "reports", "tunnel_url.txt")
ENDPOINT = os.path.join(ROOT, "docs", "endpoint.json")
LOG_FILE = os.path.join(ROOT, "reports", "tunnel_run.log")
PORT = 8899
# ★ 대상은 반드시 **127.0.0.1**. 'localhost'로 주면 윈도우에서 IPv6(::1)로 먼저 풀리는데
#   앱은 IPv4(0.0.0.0)로만 듣기 때문에 연결이 거부돼 폰에는 **HTTP 530**만 돌아온다.
#   터널은 살아 있는데 아무것도 안 열리는 그 증상의 정체다(2026-07-27 원인 확정).


def _background_entry():
    """자동 실행은 어느 입구로 들어와도 콘솔을 남기지 않는다.

    예전 실행 파일 두 개가 `python.exe` 로 이 장기 실행 서비스를 불렀다. cloudflared
    자식만 CREATE_NO_WINDOW 로 숨겨도 **부모 Python 콘솔**에는 주소 게시 재시도 문구가
    계속 쌓여 화면을 가렸다(2026-08-14 실사고). 호출자마다 고치면 새 바로가기에서
    다시 생기므로, 서비스 입구가 스스로 pythonw 로 갈아탄다. 사람이 로그를 볼 때만
    `--console` 을 명시한다.
    """
    if os.name != "nt" or "--console" in sys.argv:
        return False
    if os.path.basename(sys.executable).lower() == "pythonw.exe":
        return False
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pythonw):
        return False
    env = dict(os.environ)
    env["CSOS_TUNNEL_BACKGROUND"] = "1"
    subprocess.Popen(
        [pythonw, "-u", os.path.abspath(__file__), *sys.argv[1:]],
        cwd=ROOT,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=(getattr(subprocess, "CREATE_NO_WINDOW", 0)
                       | getattr(subprocess, "DETACHED_PROCESS", 0)),
        env=env,
    )
    return True


def _background_log():
    """창을 없애도 실패 이유는 잃지 않도록 1MB 단위 파일 로그를 연결한다."""
    background = (os.environ.get("CSOS_TUNNEL_BACKGROUND") == "1"
                  or os.path.basename(sys.executable).lower() == "pythonw.exe"
                  or sys.stdout is None)
    if not background:
        return
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    try:
        if os.path.getsize(LOG_FILE) >= 1024 * 1024:
            old = LOG_FILE + ".1"
            if os.path.exists(old):
                os.remove(old)
            os.replace(LOG_FILE, old)
    except OSError:
        pass
    try:
        stream = open(LOG_FILE, "a", encoding="utf-8", buffering=1)
        sys.stdout = stream
        sys.stderr = stream
    except OSError:
        pass


def _local_app_alive(timeout=5):
    """Return True only when the local CSOS API is actually responding."""
    import urllib.request
    try:
        with urllib.request.urlopen(
                f"http://127.0.0.1:{PORT}/api/ping", timeout=timeout) as response:
            return b"coupang-work" in response.read()
    except Exception:
        return False


def ensure_local_app():
    """Start the app server when the tunnel exists but its origin is down."""
    if _local_app_alive():
        return True
    py = sys.executable
    pythonw = py.replace("python.exe", "pythonw.exe")
    if os.path.exists(pythonw):
        py = pythonw
    app = os.path.join(ROOT, "webapp", "app_server.py")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.Popen(
            [py, "-u", app],
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=flags,
        )
    except OSError as exc:
        print(f"앱 서버 자동 시작 실패: {exc}")
        return False
    for _ in range(30):
        time.sleep(0.5)
        if _local_app_alive(2):
            print(f"앱 서버 자동 복구 완료: http://127.0.0.1:{PORT}")
            return True
    print("앱 서버를 시작했지만 응답 확인에 실패했습니다.")
    return False


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


MAX_PER_HOUR = 4       # 터널이 **살아 있을 때** 갈아치우는 한도
HARD_MAX = 10          # 죽어 있어도 이 이상은 안 만든다(폭주 방지)
STAMP = os.path.join(ROOT, "reports", "tunnel_starts.txt")


def _ping(u, t=8):
    """회사 DNS가 trycloudflare를 막고 있어 직접 찔러 보면 늘 실패한다 — 공개 DNS로 우회."""
    sys.path.insert(0, ROOT)
    from net_probe import probe
    return probe(u, t)[0]


def _endpoint_alive():
    """지금 고정 주소가 가리키는 터널이 실제로 응답하는가."""
    try:
        u = json.load(open(ENDPOINT, encoding="utf-8")).get("url", "")
    except Exception:
        u = ""
    return bool(u) and _ping(u.rstrip("/") + "/api/ping")


def _throttle():
    """새 터널 생성 한도. 주소가 바뀔 때마다 폰·PC 북마크가 가리키는 곳이 달라지므로
    살아 있는 터널을 자주 갈아치우면 안 된다(심하면 DNS에 없는 주소를 받는다).

    ★ 단, 한도는 '주소가 자꾸 바뀌는 것'을 막으려는 것이지 **아무 서비스도 없는 상태**를
      지키려는 게 아니다. 지금 게시된 주소가 죽어 있으면 기다릴 이유가 없다 — 기다리는
      동안 폰은 아예 못 들어온다(2026-07-27: 재시작 몇 번에 한도가 차서 접속 주소 없이
      멈춰 있었다). 그래서 **살아 있을 때만** 기다리고, 죽어 있으면 즉시 새로 만든다.
      그래도 HARD_MAX를 넘으면 진짜 폭주이므로 그때는 기다린다.
    """
    now = time.time()
    try:
        hist = [float(x) for x in open(STAMP, encoding="utf-8").read().split()]
    except Exception:
        hist = []
    hist = [t for t in hist if now - t < 3600]

    down = len(hist) >= MAX_PER_HOUR and not _endpoint_alive()
    if down and len(hist) < HARD_MAX:
        print(f"한도({MAX_PER_HOUR}회/시간)를 넘었지만 **접속 주소가 죽어 있어** "
              f"바로 새로 만듭니다 (최근 1시간 {len(hist)}회)")
    else:
        while len(hist) >= MAX_PER_HOUR:
            wait = int(3600 - (now - hist[0])) + 5
            why = "폭주로 판단" if len(hist) >= HARD_MAX else "너무 자주 새로 만들고 있습니다"
            print(f"터널 {why} — {wait}초 기다립니다 (최근 1시간 {len(hist)}회)")
            time.sleep(min(wait, 300))
            now = time.time()
            hist = [t for t in hist if now - t < 3600]

    hist.append(now)
    try:
        os.makedirs(os.path.dirname(STAMP), exist_ok=True)
        open(STAMP, "w", encoding="utf-8").write("\n".join(str(t) for t in hist))
    except OSError:
        pass


def published_url():
    """고정 주소가 지금 가리키는 곳(끝 슬래시 제거)."""
    try:
        return (json.load(open(ENDPOINT, encoding="utf-8")).get("url") or "").rstrip("/")
    except Exception:
        return ""


def publish(url):
    """고정 주소(GitHub Pages)가 지금 주소를 가리키게 한다.

    폰 북마크는 고정 주소 하나만 알면 되고, 실제 주소가 바뀌어도 여기서 따라간다.
    (trycloudflare 무료 터널은 **주소를 지정할 수 없다** — 매번 새로 받는다)
    """
    if is_input_window():
        print("입력 보호시간 — 고정 주소 게시 생략")
        return False
    # ★ 예전에는 결과를 안 보고 무조건 "갱신 완료"를 찍었다. publish_endpoint 가 게시를
    #   거절해도(주소가 아직 안 붙었을 때 흔하다) 성공처럼 보였고, 고정 주소는 **옛 죽은
    #   주소를 그대로 가리킨 채** 남았다. 폰은 다음 날 아침 접속이 안 됐다(2026-07-28).
    #   이제 실제로 게시됐는지 endpoint.json 으로 확인하고, 안 됐으면 다시 시도한다.
    for attempt in range(1, 4):
        try:
            subprocess.run([sys.executable, os.path.join(ROOT, "publish_endpoint.py")],
                           cwd=ROOT, timeout=180, capture_output=True,
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"},
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            print(f"   고정 주소 갱신 오류: {type(e).__name__} {e}")
        if published_url() == url.rstrip("/"):
            print("   고정 주소 갱신 완료")
            return True
        wait = attempt * 20
        print(f"   고정 주소가 아직 새 주소를 가리키지 않음 — {wait}초 후 재시도 ({attempt}/3)")
        time.sleep(wait)
    print("   ★ 고정 주소 갱신 실패 — 폰이 옛 주소로 들어가게 됩니다")
    return False


def watch(proc, url):
    """주소가 살아 있는지 1분마다 확인해 죽으면 cloudflared를 내린다.

    ★ 이게 없으면 **cloudflared 프로세스는 살아 있는데 주소만 만료된** 상태로 방치된다.
      바깥에서는 '사이트에 연결할 수 없습니다'인데 안에서는 정상으로 보여 아무도 못 고친다
      (2026-07-27 실사고 — 이틀 동안 접속 불가였다).
      프로세스를 내리면 바깥 while 루프가 새 주소로 다시 띄운다.
    """
    import threading, urllib.request

    def ping(u, t=20):
        # ★ 사내망은 *.trycloudflare.com 을 안 풀어 준다 — 그대로 믿으면 멀쩡한 터널을
        #   '죽었다'며 계속 갈아치운다(그게 폰이 못 들어온 진짜 이유였다).
        sys.path.insert(0, ROOT)
        from net_probe import probe
        return probe(u, t)[0]

    def loop():
        fail = 0
        while proc.poll() is None:
            time.sleep(90)
            if is_input_window():
                fail = 0
                continue
            if ping(url + "/api/ping"):
                fail = 0
                # ★ 터널은 멀쩡한데 **고정 주소만 딴 데를 가리키는** 경우가 실제로 있었다
                #   (게시가 조용히 실패). 매 점검마다 맞는지 보고 어긋나면 다시 게시한다.
                if published_url() != url.rstrip("/"):
                    print("고정 주소가 현재 터널과 다릅니다 — 다시 게시합니다")
                    publish(url)
                continue
            # 앱 자체가 죽었으면 터널을 갈아도 소용없다 — 주소만 쓸데없이 바뀐다.
            # (주소가 바뀔 때마다 직접 북마크한 사람은 다시 못 들어온다)
            if not ping(f"http://127.0.0.1:{PORT}/api/ping", 8):
                print("앱이 응답하지 않음 — 로컬 앱 서버를 자동 복구합니다")
                ensure_local_app()
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
    if is_input_window():
        print("입력 보호시간 — 새 터널을 시작하지 않습니다.")
        return
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
        if not ensure_local_app():
            print("앱 서버 복구 대기 — 10초 후 다시 시도합니다")
            time.sleep(10)
            continue
        print(f"터널 시작... (대상 http://127.0.0.1:{PORT})")
        # ★ 창을 달지 않는다(2026-08-13 지시 · [248] 과 같은 자리).
        #   tunnel_run 자신은 pythonw(콘솔 없음)로 뜨므로, **콘솔 앱**인 cloudflared 를
        #   깃발 없이 띄우면 윈도우가 새 콘솔을 할당한다. 윈도우 11 기본 콘솔 호스트가
        #   터미널이라 제목이 'webapp\cloudflared.exe' 인 검은 창이 사람 화면을 덮었다.
        #   터널이 끊기면 아래 while 이 5초마다 다시 띄우므로 **창이 계속 쌓인다.**
        #   ⚠ 여기서 잃는 로그는 없다 — cloudflared 출력은 아래 for 문이 PIPE 로 그대로
        #   읽어 주소를 뽑고 URL_FILE 에 적는다. 그래서 그 창에는 애초에 아무것도
        #   안 찍혔다(순수한 방해라 고장으로 안 보였다). 창만 없애고 읽기는 그대로다.
        p = subprocess.Popen([exe, "tunnel", "--url", f"http://127.0.0.1:{PORT}"],
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace",
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
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
    if _background_entry():
        raise SystemExit(0)
    _background_log()
    main()
