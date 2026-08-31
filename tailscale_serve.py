# -*- coding: utf-8 -*-
"""tailscale_serve.py — 전체 기능 앱을 **바뀌지 않는 주소**로 띄운다.

사용자 지시(2026-07-28): "전체 기능 사용할 수 있는 앱 고정주소가 필요해."

왜 이게 답인가
  trycloudflare 무료 터널은 **주소를 지정할 수 없다** — 띄울 때마다 새로 받는다.
  그래서 폰 북마크·홈아이콘·창 복원이 재부팅마다 죽었다(2026-07-28 종일 겪음).
  Tailscale 은 기기마다 **영구 이름**(MagicDNS)을 준다. 이 PC 는 `mulder.tailf14aae.ts.net`.
  이 이름은 재부팅해도, IP 가 바뀌어도, 회선을 갈아도 그대로다.

두 가지 방식
  serve   : 내 tailnet 안에서만 열린다. 폰에 Tailscale 을 깔고 같은 계정으로 1회 로그인.
            인터넷에 전혀 노출되지 않는다(지금 trycloudflare 보다 안전).
  funnel  : 인터넷에 공개한다. 폰에 아무것도 안 깔아도 된다.
            ★ tailnet 관리자 페이지에서 **한 번 허용**해야 켜진다(계정 로그인이라 사람이 해야 함).

주소는 둘 다 같다 — `https://<MagicDNS이름>/`. 방식만 바뀐다.

실행
  python tailscale_serve.py --status
  python tailscale_serve.py --serve      # 사설망 전용
  python tailscale_serve.py --funnel     # 인터넷 공개
  python tailscale_serve.py --off
"""
import argparse
import json
import os
import socket
import ssl
import subprocess
import sys
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
PORT = 8899
# 사용자 확정 정본. 장치 이름이나 자동 탐지 결과가 달라져도 새 주소를
# 자동 게시하지 않는다. 불일치는 주소 변경이 아니라 관리자 확인 대상이다.
FIXED_HOST = "mulder.tailf14aae.ts.net"
FIXED_URL = "https://mulder.tailf14aae.ts.net/"
EXE_CANDIDATES = [
    r"C:\Program Files\Tailscale\tailscale.exe",
    r"C:\Program Files (x86)\Tailscale\tailscale.exe",
]


def exe():
    for p in EXE_CANDIDATES:
        if os.path.exists(p):
            return p
    from shutil import which
    return which("tailscale")


def run(*args, timeout=90):
    """★ `funnel` 은 tailnet 에서 아직 허용되지 않았으면 **끝나지 않고 기다린다**.
    타임아웃을 예외로 터뜨리면 도구가 통째로 죽으므로 코드 124 로 돌려준다."""
    e = exe()
    if not e:
        return 1, "", "tailscale.exe 를 찾지 못했습니다"
    try:
        r = subprocess.run([e] + list(args), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired as ex:
        out = ex.stdout or ""
        err = ex.stderr or ""
        if isinstance(out, bytes):
            out = out.decode("utf-8", "replace")
        if isinstance(err, bytes):
            err = err.decode("utf-8", "replace")
        return 124, out.strip(), err.strip()
    return r.returncode, (r.stdout or "").strip(), (r.stderr or "").strip()


def hostname():
    """MagicDNS 이름 — 이게 '바뀌지 않는 주소'의 정체다."""
    code, out, _err = run("status", "--json")
    if code:
        return ""
    try:
        d = json.loads(out)
    except ValueError:
        return ""
    name = (d.get("Self") or {}).get("DNSName") or ""
    return name.rstrip(".")


def public_ingress_ips(host: str, timeout: int = 8) -> list[str]:
    """Resolve Funnel through public DNS, not this PC's tailnet DNS.

    On the Tailscale PC the hostname resolves to its private 100.x address.
    That can answer normally even while the public Funnel relays are failing,
    which is exactly the state in which a phone without Tailscale cannot open
    the app.  Public DNS-over-HTTPS answers expose the public ingress IPs.
    Google is tried first and Cloudflare is the fallback; a brief resolver
    outage must not be mistaken for three minutes of Funnel failure.
    """
    if not host:
        return []
    urls = (
        "https://dns.google/resolve?name=%s&type=A" % host,
        "https://cloudflare-dns.com/dns-query?name=%s&type=A" % host,
    )
    for url in urls:
        try:
            request = urllib.request.Request(
                url,
                headers={"Accept": "application/dns-json", "User-Agent": "CSOS-Funnel-Check"},
            )
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception:
            continue
        found = []
        for answer in payload.get("Answer") or []:
            value = str(answer.get("data") or "").strip()
            if answer.get("type") == 1 and value and value not in found:
                found.append(value)
        if found:
            return found
    return []


def _public_ping_ip(host: str, ip: str, timeout: int = 8) -> bool:
    """TLS-ping one public Funnel relay while keeping SNI/cert host intact."""
    request = (
        f"GET /api/ping?public-funnel-check=1 HTTP/1.1\r\n"
        f"Host: {host}\r\nConnection: close\r\nUser-Agent: CSOS-Funnel-Check\r\n\r\n"
    ).encode("ascii")
    try:
        with socket.create_connection((ip, 443), timeout=timeout) as raw:
            raw.settimeout(timeout)
            with ssl.create_default_context().wrap_socket(raw, server_hostname=host) as tls:
                tls.sendall(request)
                chunks = []
                total = 0
                while total < 32768:
                    chunk = tls.recv(min(4096, 32768 - total))
                    if not chunk:
                        break
                    chunks.append(chunk)
                    total += len(chunk)
        response = b"".join(chunks)
        status = response.split(b"\r\n", 1)[0]
        return b" 200 " in status and b"coupang-work" in response
    except Exception:
        return False


def public_funnel_alive(host: str | None = None, timeout: int = 8) -> bool:
    """True only when at least one *public* Funnel ingress reaches the app."""
    host = host or FIXED_HOST
    ips = public_ingress_ips(host, timeout)
    return bool(ips) and any(_public_ping_ip(host, ip, timeout) for ip in ips)


def ensure_public_funnel(repair: bool = True) -> tuple[bool, bool]:
    """Verify the phone path and re-register the same fixed URL when stale.

    Returns ``(available, repaired)``.  Resetting Funnel does not change the
    MagicDNS hostname; it only refreshes the public relay/TLS route.
    """
    actual_host = hostname()
    if not actual_host or actual_host != FIXED_HOST:
        # Never follow a renamed device to a new public address silently.
        return False, False
    host = FIXED_HOST
    try:
        from webapp.tunnel_run import ensure_local_app
        if not ensure_local_app():
            return False, False
    except Exception:
        pass
    if public_funnel_alive(host):
        return True, False
    if not repair:
        return False, False

    # A stale relay can still be reported as "Funnel on".  Re-registering the
    # existing mapping refreshes its public TLS route without changing the URL.
    #
    # ★ 2026-08-31. 여기 있던 Funnel 전체 초기화 호출을 뺐다.
    #   `funnel reset` 은 **이 노드의 Funnel 설정을 통째로 지운다** — 이 앱 것만이
    #   아니다. 같은 노드(mulder)를 자금흐름 앱이 :10000·:8443 으로 나눠 쓰는데,
    #   이 함수가 60초마다 도는 server_guard 에서 불리므로 그 두 문이 곁불로
    #   사라졌다(2026-08 실측 55회). 그때 자금흐름 앱은 폰에서 통째로 안 열린다.
    #   ★ 지우지 않아도 되는 근거는 **바로 위 주석 자신**이다 —
    #     "Re-registering the existing mapping refreshes its public TLS route".
    #     고치는 것은 재등록이지 삭제가 아니다.
    #   ★ 그리고 자금흐름 앱은 처음부터 reset 없이 제 문을 걸어 왔고 잘 된다.
    #     즉 재등록만으로 충분하다는 것이 이미 확인돼 있다.
    #   ★ tailscale 에는 **포트 하나만 끄는 명령이 없다**(serve·funnel 둘 다
    #     reset 뿐이다). 그래서 '자기 것만 끄기' 가 아니라 '지우지 않기' 로 했다.
    #   되돌리려면: 같은 폴더의 tailscale_serve.py.bak_20260831 로 바꾸면 된다
    #   (그 안에 옛 reset 호출이 그대로 있다). 여기에 그 명령을 글자로 적어 두지
    #   않는 까닭은, 이 앱 게이트가 그 글자가 **있는지**로 자동복구를 확인하기
    #   때문이다 — 주석에 남기면 검사가 우연히 통과하고, 그러면 그 검사는
    #   있으나 마나가 된다.
    code, _out, _err = run("funnel", "--bg", str(PORT), timeout=30)
    if code:
        return False, False
    for _ in range(12):
        time.sleep(1)
        if public_funnel_alive(host, timeout=5):
            return True, True
    return False, True


def status():
    actual_host = hostname()
    host = FIXED_HOST
    code, out, err = run("funnel", "status")
    print("MagicDNS 이름 :", actual_host or "(확인 실패 — Tailscale 로그인 상태를 보세요)")
    print("고정 주소     :", FIXED_URL)
    if actual_host and actual_host != FIXED_HOST:
        print("주소 불일치   : 자동 변경 금지 — Tailscale 장치 이름을 원복해야 합니다")
    print("현재 설정     :", (out or err or "(없음)").splitlines()[0] if (out or err) else "(없음)")
    print("휴대폰 외부경로:", "정상" if public_funnel_alive(FIXED_HOST) else "연결 실패")
    return 0


def enable(mode):
    actual_host = hostname()
    if not actual_host:
        print("Tailscale 이 로그인/실행 상태가 아닙니다. 먼저 Tailscale 을 켜 주세요.")
        return 1
    if actual_host != FIXED_HOST:
        print("Tailscale 장치 이름이 고정주소와 다릅니다. 주소를 바꾸지 않고 중단합니다.")
        print("  고정 주소:", FIXED_URL)
        print("  현재 이름:", actual_host)
        return 1
    code, out, err = run(mode, "--bg", str(PORT), timeout=25)
    msg = (out + "\n" + err).strip()
    if code == 124 and mode == "funnel":
        # 승인 대기로 멈춘 것이다. 매달려 있어 봐야 켜지지 않는다.
        print("Funnel 이 아직 tailnet 에서 허용되지 않았습니다 (명령이 승인을 기다리며 멈춥니다).")
        print("  관리자 페이지에서 Enable Funnel 을 누른 뒤 다시 실행하세요:")
        print("    python tailscale_serve.py --funnel")
        print("  링크는 `tailscale funnel 8899` 를 직접 실행하면 나옵니다.")
        return 2
    if "not enabled on your tailnet" in msg or "login.tailscale.com/f/funnel" in msg:
        # ★ 여기서 멈추는 게 맞다. 이 승인은 tailnet 관리자 계정 로그인이라 도구가 대신 못 한다.
        print("Funnel 이 아직 tailnet 에서 허용되지 않았습니다.")
        for ln in msg.splitlines():
            if "login.tailscale.com" in ln:
                print("  이 링크를 열어 Enable Funnel 을 눌러 주세요:\n   ", ln.strip())
        print("  누른 뒤 다시:  python tailscale_serve.py --funnel")
        return 2
    if code:
        print("실패:", msg or code)
        return 1
    url = FIXED_URL
    print("켰습니다 —", ("인터넷 공개(Funnel)" if mode == "funnel" else "내 tailnet 전용(Serve)"))
    print("  고정 주소:", url)
    print("  ※ 이 주소는 재부팅해도 바뀌지 않습니다. 폰·문서에 이걸 적으세요.")
    if mode == "funnel":
        ok, repaired = ensure_public_funnel(repair=True)
        print("  휴대폰 외부경로:", "정상" if ok else "연결 실패",
              "(공개 경로 재등록)" if repaired else "")
        if not ok:
            return 1
    return 0


def disable():
    code, out, err = run("serve", "reset")
    print("껐습니다." if not code else ("실패: " + (err or out or str(code))))
    return code


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--serve", action="store_true", help="내 tailnet 안에서만 열기")
    ap.add_argument("--funnel", action="store_true", help="인터넷에 공개")
    ap.add_argument("--repair", action="store_true", help="휴대폰 외부경로 검사·자동복구")
    ap.add_argument("--off", action="store_true")
    a = ap.parse_args()
    if a.off:
        return disable()
    if a.serve:
        return enable("serve")
    if a.funnel:
        return enable("funnel")
    if a.repair:
        ok, repaired = ensure_public_funnel(repair=True)
        print("휴대폰 외부경로:", "정상" if ok else "연결 실패",
              "(공개 경로 재등록)" if repaired else "")
        return 0 if ok else 1
    return status()


if __name__ == "__main__":
    sys.exit(main())
