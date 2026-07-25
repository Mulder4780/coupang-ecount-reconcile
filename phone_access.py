# -*- coding: utf-8 -*-
"""
phone_access.py — 휴대폰 접속 안내 페이지 (QR + 주소 + PIN)
==============================================================
외부접속 터널(trycloudflare)은 재시작할 때마다 주소가 바뀐다.
이 스크립트는 **현재 살아있는 주소**를 찾아 QR과 함께 브라우저에 띄운다.
폰으로 QR만 찍으면 되므로 주소가 바뀌어도 타이핑할 필요가 없다.

동작:
  1. 터널 주소(reports/tunnel_url.txt)를 읽고 실제 접속되는지 확인
  2. 죽어 있으면 터널을 다시 띄우고 새 주소를 기다림
  3. QR(SVG) + 주소 + PIN을 담은 안내 페이지를 만들어 기본 브라우저로 연다

실행:  python phone_access.py       (바탕화면 "폰 접속 QR" 바로가기와 동일)
"""
import sys, os, io, json, time, subprocess, webbrowser, urllib.request

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(BASE, "reports", "tunnel_url.txt")
OUT_HTML = os.path.join(BASE, "reports", "폰접속.html")
PORT = 8899


def alive(url, timeout=12):
    try:
        with urllib.request.urlopen(url.rstrip("/") + "/api/ping", timeout=timeout) as r:
            return b"coupang-work" in r.read()
    except Exception:
        return False


def read_url():
    try:
        return open(URL_FILE, encoding="utf-8").read().strip()
    except Exception:
        return ""


def ensure_tunnel():
    """살아있는 터널 주소를 반환. 없으면 띄우고 최대 40초 대기."""
    url = read_url()
    if url and alive(url):
        return url, False
    py = sys.executable.replace("python.exe", "pythonw.exe")
    py = py if os.path.exists(py) else sys.executable
    subprocess.Popen([py, os.path.join(BASE, "webapp", "tunnel_run.py")],
                     cwd=BASE, creationflags=0x00000008)
    for _ in range(20):
        time.sleep(2)
        u = read_url()
        if u and u != url and alive(u, 6):
            return u, True
    return read_url(), True


def local_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]; s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def qr_svg(data):
    """외부 이미지 없이 QR을 SVG로 (pillow 불필요)"""
    import qrcode
    import qrcode.image.svg as svg
    img = qrcode.make(data, image_factory=svg.SvgPathImage, box_size=11, border=2)
    buf = io.BytesIO(); img.save(buf)
    return buf.getvalue().decode("utf-8")


def pin():
    try:
        return json.load(open(os.path.join(BASE, "config", "webapp.json"), encoding="utf-8"))["pin"]
    except Exception:
        return "----"


FIXED = "https://mulder4780.github.io/coupang-ecount-reconcile/"   # 바뀌지 않는 진입점


def main():
    print("터널 확인 중…")
    url, restarted = ensure_tunnel()
    if not url:
        sys.exit("터널을 띄우지 못했습니다. webapp/cloudflared.exe 존재 여부를 확인하세요.")
    ok = alive(url)
    # 현재 주소를 고정 진입점에 게시(변경 시에만 커밋)
    try:
        subprocess.run([sys.executable, os.path.join(BASE, "publish_endpoint.py")],
                       cwd=BASE, capture_output=True, timeout=120)
    except Exception:
        pass
    lan = f"http://{local_ip()}:{PORT}"
    html = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>폰 접속</title>
<style>body{{margin:0;font-family:"Malgun Gothic",sans-serif;background:#0E1B3F;color:#fff;
 display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}}
.card{{background:#fff;color:#101828;border-radius:20px;padding:28px;max-width:430px;width:100%;text-align:center;
 box-shadow:0 20px 60px rgba(0,0,0,.4)}}
h1{{font-size:16px;margin:0 0 5px;letter-spacing:-.2px}} .sub{{font-size:12.5px;color:#667085;margin-bottom:18px}}
.qr{{background:#fff;border:1px solid #E4E9F0;border-radius:14px;padding:10px;display:inline-block}}
.qr svg{{width:230px;height:230px;display:block}}
.pin{{font-size:34px;font-weight:800;letter-spacing:10px;color:#2452E6;margin:14px 0 2px}}
.u{{font-size:11.5px;word-break:break-all;background:#F4F6FA;border-radius:9px;padding:9px;margin-top:12px;color:#475467}}
.b{{display:inline-block;margin-top:12px;background:#2452E6;color:#fff;text-decoration:none;
 padding:11px 18px;border-radius:11px;font-weight:700;font-size:14px}}
.s{{margin-top:14px;font-size:11.5px;color:#98A2B3;line-height:1.6}}
.warn{{background:#FEF1E2;color:#B54708;border-radius:9px;padding:9px;font-size:12px;margin-top:10px}}</style></head>
<body><div class="card">
<h1>Coupang Service Operations System</h1>
<div class="sub">📱 QR을 한 번만 찍으세요 · 이 주소는 <b>바뀌지 않습니다</b></div>
<div class="qr">{qr_svg(FIXED)}</div>
<div style="font-size:12px;color:#667085;margin-top:14px">접속 PIN</div>
<div class="pin">{pin()}</div>
<div class="u">{FIXED}</div>
<a class="b" href="{FIXED}" target="_blank">이 PC에서 열기</a>
{'' if ok else '<div class="warn">현재 접속 주소가 응답하지 않습니다. 잠시 후 다시 실행해 주세요.</div>'}
<div class="s">폰에서 열면 <b>현재 주소로 자동 연결</b>됩니다 — 홈 화면에 추가해 두고 쓰세요.<br>
지금 연결 주소: {url[:46]}…<br>사내 와이파이 전용: <b>{lan}</b></div>
</div></body></html>"""
    os.makedirs(os.path.dirname(OUT_HTML), exist_ok=True)
    open(OUT_HTML, "w", encoding="utf-8").write(html)
    print(f"주소: {url}  (응답 {'정상' if ok else '확인필요'}{', 터널 재시작함' if restarted else ''})")
    print(f"PIN : {pin()}")
    webbrowser.open("file:///" + OUT_HTML.replace("\\", "/"))


if __name__ == "__main__":
    main()
