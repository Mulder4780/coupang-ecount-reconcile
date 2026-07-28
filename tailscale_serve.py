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
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
PORT = 8899
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
                           encoding="utf-8", errors="replace", timeout=timeout)
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


def status():
    host = hostname()
    code, out, err = run("serve", "status")
    print("MagicDNS 이름 :", host or "(확인 실패 — Tailscale 로그인 상태를 보세요)")
    print("고정 주소     :", ("https://%s/" % host) if host else "-")
    print("현재 설정     :", (out or err or "(없음)").splitlines()[0] if (out or err) else "(없음)")
    return 0


def enable(mode):
    host = hostname()
    if not host:
        print("Tailscale 이 로그인/실행 상태가 아닙니다. 먼저 Tailscale 을 켜 주세요.")
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
    url = "https://%s/" % host
    print("켰습니다 —", ("인터넷 공개(Funnel)" if mode == "funnel" else "내 tailnet 전용(Serve)"))
    print("  고정 주소:", url)
    print("  ※ 이 주소는 재부팅해도 바뀌지 않습니다. 폰·문서에 이걸 적으세요.")
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
    ap.add_argument("--off", action="store_true")
    a = ap.parse_args()
    if a.off:
        return disable()
    if a.serve:
        return enable("serve")
    if a.funnel:
        return enable("funnel")
    return status()


if __name__ == "__main__":
    sys.exit(main())
