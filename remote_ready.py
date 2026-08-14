# -*- coding: utf-8 -*-
"""remote_ready.py — **폰에서 Claude Code 를 붙일 수 있는 상태인가**를 한 번에 본다.

사용자 지시(2026-07-31): "휴대폰에서 클로드 코드 리모트로 할 수 있게 정리."

폰에서 이 프로젝트를 이어서 하는 길은 두 가지이고, 막히는 지점이 서로 다르다.

  A. 웹 Claude Code (claude.ai/code)  — 폰 브라우저만 있으면 된다. PC 가 꺼져 있어도 된다.
     GitHub 저장소를 클라우드에서 열기 때문에 **푸시된 것까지만** 보인다.
     Z: 관리대장·이카운트·밴드에는 손이 닿지 않는다 → 코드·문서·검증 작업용.
  B. SSH 로 이 PC 접속 후 `claude` 실행 — 진짜 이 PC 다. Z: 도 ERP 도 다 된다.
     대신 **PC 가 깨어 있어야** 하고, OpenSSH 서버가 켜져 있어야 한다.

이 도구는 그 두 길의 준비 상태를 점검만 한다 — **아무것도 바꾸지 않는다.**
전원 설정·SSH 서버는 관리자 권한이 필요하고 시스템 설정이라, 명령만 알려 준다.
자세한 절차는 `REMOTE_PHONE.md`.
"""
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
REPORT = os.path.join(REPORT_DIR, "remote_ready.json")

# Claude Code 데스크톱이 함께 설치하는 CLI. 버전 폴더가 바뀌므로 상위에서 찾는다.
CLI_BASE = os.path.join(os.environ.get("APPDATA", ""), "Claude", "claude-code")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _decode(raw):
    """★ 한국어 Windows 콘솔은 cp949 로 낸다. utf-8 로 못박으면 powercfg 출력이
    통째로 깨져 '설정값을 못 찾음' 이 된다(처음에 실제로 그랬다)."""
    for enc in ("utf-8", "cp949", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _run(cmd, timeout=20):
    """실패해도 죽지 않는다 — 없는 명령·꺼진 서비스가 정상 상태일 수 있다."""
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=timeout, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return r.returncode, _decode(r.stdout or b"").strip(), _decode(r.stderr or b"").strip()
    except (OSError, subprocess.SubprocessError):
        return -1, "", "실행 불가"


def check_tailscale():
    """폰↔PC 를 잇는 사설망. 공유기 포트를 열지 않아도 되는 이유가 이것이다."""
    exe = shutil.which("tailscale")
    if not exe:
        return {"ok": False, "detail": "tailscale CLI 없음",
                "how": "Tailscale 설치 후 로그인 (https://tailscale.com/download)"}
    code, out, _ = _run([exe, "status"])
    if code != 0:
        return {"ok": False, "detail": "tailscale 로그인 안 됨", "how": "tailscale up"}
    ip = ""
    for line in out.splitlines():
        parts = line.split()
        if parts and parts[0].count(".") == 3 and parts[0].startswith("100."):
            ip = parts[0]
            break
    return {"ok": bool(ip), "detail": ("PC 사설망 주소 " + ip) if ip else "주소 미할당",
            "ip": ip, "how": "" if ip else "tailscale up"}


def check_sshd():
    """B 경로의 관문. Windows 는 Tailscale SSH 서버를 못 쓰므로 OpenSSH 서버가 필요하다."""
    code, out, _ = _run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                         "(Get-Service sshd -ErrorAction SilentlyContinue).Status"])
    status = (out or "").strip()
    if status == "Running":
        return {"ok": True, "detail": "OpenSSH 서버 실행 중"}
    if status:
        return {"ok": False, "detail": "OpenSSH 서버 설치됐으나 %s" % status,
                "how": "관리자 PowerShell: Start-Service sshd; Set-Service sshd -StartupType Automatic"}
    return {"ok": False, "detail": "OpenSSH 서버 없음",
            "how": "관리자 PowerShell: Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0"}


def check_sleep():
    """★ 실제로 여기서 막힌다. 폰에서 붙어도 PC 가 자면 그걸로 끝이다.

    2026-07-31 진단에서 '잠자기만 끄면 된다'고 적어 뒀으나 값은 그대로 5시간이었다.
    """
    code, out, _ = _run(["powercfg", "/query", "SCHEME_CURRENT", "SUB_SLEEP",
                         "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"])
    if code != 0:
        return {"ok": None, "detail": "전원 설정을 읽지 못함"}
    # ★ 언어에 기대지 않는다. 이 질의는 설정 하나만 내므로 0x 값의 순서가 고정이다:
    #   [최소, 최대, 증가, AC, DC] — 뒤에서 두 번째가 AC(전원 연결 시)다.
    #   한국어판은 '현재 AC 전원 설정 색인', 영문판은 'Current AC Power Setting Index' 라
    #   문구로 잡으면 한쪽에서 반드시 깨진다.
    values = []
    for line in out.splitlines():
        if "0x" in line:
            try:
                values.append(int(line.split("0x")[1].strip().split()[0], 16))
            except (ValueError, IndexError):
                pass
    ac = values[-2] if len(values) >= 2 else None
    if ac is None:
        return {"ok": None, "detail": "대기모드 설정값을 못 찾음"}
    if ac == 0:
        return {"ok": True, "detail": "전원 연결 시 대기모드 없음(항상 깨어 있음)", "seconds": 0}
    return {"ok": False, "seconds": ac,
            "detail": "전원 연결 상태에서도 %d분 뒤 대기모드 — 그때부터 폰이 못 붙는다" % (ac // 60),
            "how": "powercfg /change standby-timeout-ac 0"}


def check_cli():
    """B 경로에서 SSH 로 들어와 실행할 실행파일."""
    if not os.path.isdir(CLI_BASE):
        return {"ok": False, "detail": "Claude Code CLI 폴더 없음"}
    found = []
    for name in sorted(os.listdir(CLI_BASE)):
        exe = os.path.join(CLI_BASE, name, "claude.exe")
        if os.path.exists(exe):
            found.append((name, exe))
    if not found:
        return {"ok": False, "detail": "claude.exe 를 찾지 못함"}
    version, path = found[-1]
    return {"ok": True, "detail": "claude %s" % version, "path": path}


def check_repo():
    """A 경로(웹)는 **푸시된 것만** 본다. 미푸시가 있으면 폰에서 옛 코드를 보게 된다."""
    def git(*args):
        code, out, _ = _run(["git", "-C", ROOT] + list(args))
        return out if code == 0 else ""

    unpushed = [l for l in git("log", "origin/master..HEAD", "--oneline").splitlines() if l.strip()]
    dirty = [l for l in git("status", "--short").splitlines()
             if l.strip() and not l.strip().startswith("??")]
    url = git("remote", "get-url", "origin")
    if unpushed or dirty:
        return {"ok": False, "remote": url,
                "detail": "미푸시 커밋 %d개 · 미커밋 변경 %d개 — 폰(웹)에서는 안 보인다"
                          % (len(unpushed), len(dirty)),
                "how": "git commit 후 git push (비밀 스캔 먼저)"}
    return {"ok": True, "remote": url, "detail": "저장소가 원격과 같다 — 웹에서 최신으로 열린다"}


CHECKS = (
    ("사설망(Tailscale)", check_tailscale, "B"),
    ("OpenSSH 서버", check_sshd, "B"),
    ("PC 절전", check_sleep, "B"),
    ("Claude Code CLI", check_cli, "B"),
    ("저장소 동기", check_repo, "A"),
)


def collect():
    out = {}
    for name, fn, path in CHECKS:
        try:
            r = fn()
        except Exception as exc:                      # 점검 하나가 죽어도 나머지는 봐야 한다
            r = {"ok": None, "detail": "점검 실패: %s" % str(exc)[:120]}
        r["경로"] = path
        out[name] = r
    return out


def summarize(results):
    """A(웹)와 B(SSH)를 따로 판정한다 — 한쪽만 돼도 폰에서 할 수 있는 일이 있다."""
    def ready(path):
        return all(v.get("ok") for k, v in results.items() if v["경로"] == path)
    return {"웹_claude_ai": ready("A"), "SSH_이PC": ready("B")}


def main():
    results = collect()
    summary = summarize(results)
    os.makedirs(REPORT_DIR, exist_ok=True)
    payload = {"time": datetime.now().isoformat(), "요약": summary, "상세": results}
    tmp = REPORT + ".%d.tmp" % os.getpid()
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, REPORT)

    print("폰 원격 준비: 웹(claude.ai/code) %s · SSH(이 PC) %s"
          % ("가능" if summary["웹_claude_ai"] else "막힘",
             "가능" if summary["SSH_이PC"] else "막힘"))
    for name, r in results.items():
        mark = "OK" if r.get("ok") else ("?" if r.get("ok") is None else "!!")
        print("  [%s] %-16s %s" % (mark, name, r.get("detail", "")))
        if not r.get("ok") and r.get("how"):
            print("       → %s" % r["how"])
    print("  자세한 절차: ecount/REMOTE_PHONE.md")


if __name__ == "__main__":
    main()
