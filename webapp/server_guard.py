# -*- coding: utf-8 -*-
"""Keep the local app origin available without depending on the heavy watchdog round.

The 30-minute watchdog also reconciles sources, builds reports, and can legitimately run
for hours.  App liveness must therefore have its own tiny process.  This guard checks the
cheap /api/ping endpoint every ten seconds, restarts only after three consecutive failures,
and keeps a one-minute anti-flap window between restart attempts.

It also makes sure the tunnel supervisor process exists.  tunnel_run.py owns tunnel health;
this file only replaces that supervisor when the process itself disappeared.
"""
from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import time
import urllib.request
from datetime import datetime


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER = os.path.join(ROOT, "webapp", "app_server.py")
TUNNEL = os.path.join(ROOT, "webapp", "tunnel_run.py")
REPORTS = os.path.join(ROOT, "reports")
LOG = os.path.join(REPORTS, "server_guard.log")
STATUS = os.path.join(REPORTS, "server_guard_status.json")
PORT = 8899
LOCK_PORT = 8978
CHECK_SECONDS = 10
PING_TIMEOUT = 2.5
FAIL_LIMIT = 3
READY_TIMEOUT = 75
RESTART_COOLDOWN = 60
TUNNEL_CHECK_SECONDS = 60
HEARTBEAT_SECONDS = 30


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _atomic_json(path: str, value: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temp = path + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=1)
    os.replace(temp, path)


def _log(message: str, **extra) -> None:
    os.makedirs(REPORTS, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    try:
        if os.path.exists(LOG) and os.path.getsize(LOG) >= 1_000_000:
            old = LOG + ".1"
            if os.path.exists(old):
                os.remove(old)
            os.replace(LOG, old)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
    state = {"checked_at": _now(), "message": message, **extra}
    try:
        _atomic_json(STATUS, state)
    except OSError:
        pass


def _heartbeat(state: str) -> None:
    """Leave cheap proof that the guard itself is alive for the tunnel's mutual watch."""
    try:
        try:
            with open(STATUS, encoding="utf-8") as f:
                current = json.load(f)
        except Exception:
            current = {}
        current.update({"checked_at": _now(), "heartbeat_at": _now(),
                        "guard_pid": os.getpid(), "state": state or "healthy"})
        _atomic_json(STATUS, current)
    except OSError:
        pass


def ping(timeout: float = PING_TIMEOUT) -> bool:
    try:
        req = urllib.request.Request(
            f"http://127.0.0.1:{PORT}/api/ping?t={time.time_ns()}",
            headers={"Cache-Control": "no-cache"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.status == 200 and b"coupang-work" in response.read(2048)
    except Exception:
        return False


def _cmdline_pids(needle: str, *, production_server: bool = False):
    """Return matching Python PIDs, or None when process inspection itself failed."""
    # Process inspection runs only after a health failure and once per minute for the
    # tunnel.  Keeping it out of the 10-second healthy path makes the guard very cheap.
    condition = f"$c -like '*{needle}*'"
    if production_server:
        condition += (
            " -and $c -notmatch '(?i)--demo\\b'"
            " -and ($c -notmatch '(?i)--port\\s+\\d+'"
            f" -or $c -match '(?i)--port\\s+{PORT}\\b')"
        )
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe' or Name='pythonw.exe'\" | "
        "Where-Object { $c=$_.CommandLine; $c -and (" + condition + ") } | "
        "ForEach-Object { $_.ProcessId }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True, text=True, timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            return None
        return [int(x) for x in result.stdout.split() if x.isdigit()]
    except Exception:
        return None


def server_pids():
    return _cmdline_pids("app_server.py", production_server=True)


def tunnel_pids():
    return _cmdline_pids("tunnel_run.py")


def _pythonw() -> str:
    exe = sys.executable or "python"
    if os.name == "nt" and exe.lower().endswith("python.exe"):
        quiet = exe[:-10] + "pythonw.exe"
        if os.path.isfile(quiet):
            return quiet
    return exe


def _start_hidden(script: str) -> None:
    flags = (getattr(subprocess, "CREATE_NO_WINDOW", 0)
             | getattr(subprocess, "DETACHED_PROCESS", 0)
             | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0))
    subprocess.Popen(
        [_pythonw(), "-u", script], cwd=ROOT,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=flags, close_fds=True,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )


def _stop_server(pids) -> None:
    for pid in pids or []:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F"], capture_output=True, timeout=20,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            pass
    # Do not start a second process while the old one still owns the port.
    deadline = time.time() + 12
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", PORT), timeout=0.3):
                time.sleep(0.3)
                continue
        except OSError:
            return


def restart_server(reason: str) -> bool:
    old = server_pids()
    if old is None:
        _log("프로세스 확인 실패 — 잘못된 종료를 피하려 재시작 보류", state="inspect-failed")
        return False
    if old:
        _stop_server(old)
    _start_hidden(SERVER)
    started = time.time()
    while time.time() - started < READY_TIMEOUT:
        if ping():
            new = server_pids() or []
            _log(
                "앱 서버 자동 복구 완료",
                state="healthy", reason=reason, old_pids=old, new_pids=new,
                ready_seconds=round(time.time() - started, 1),
            )
            return True
        time.sleep(0.5)
    _log(
        "앱 서버 자동 복구 실패 — 냉각 후 다시 시도",
        state="restart-failed", reason=reason, old_pids=old,
        ready_timeout=READY_TIMEOUT,
    )
    return False


def ensure_tunnel() -> None:
    pids = tunnel_pids()
    if pids == []:
        _start_hidden(TUNNEL)
        _log("터널 감독 프로세스가 없어 자동 시작", state="healthy", tunnel_started=True)


def acquire_singleton():
    lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if os.name == "nt" and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
        lock.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
    try:
        lock.bind(("127.0.0.1", LOCK_PORT))
        lock.listen(1)
        return lock
    except OSError:
        lock.close()
        return None


def main() -> int:
    singleton = acquire_singleton()
    if singleton is None:
        return 0
    _log("앱 서버 상시 감시 시작", state="starting", interval_seconds=CHECK_SECONDS)
    failures = 0
    last_restart = 0.0
    next_tunnel_check = 0.0
    next_heartbeat = 0.0
    last_state = ""
    while True:
        try:
            now = time.time()
            if now >= next_tunnel_check:
                ensure_tunnel()
                next_tunnel_check = now + TUNNEL_CHECK_SECONDS
            if now >= next_heartbeat:
                _heartbeat(last_state or "starting")
                next_heartbeat = now + HEARTBEAT_SECONDS

            if ping():
                if failures or last_state != "healthy":
                    _log("앱 서버 정상", state="healthy", consecutive_failures=0)
                failures = 0
                last_state = "healthy"
            else:
                failures += 1
                pids = server_pids() if failures == 1 else None
                missing = pids == []
                if missing or failures >= FAIL_LIMIT:
                    if now - last_restart >= RESTART_COOLDOWN:
                        reason = "프로세스 없음" if missing else f"ping {failures}회 연속 실패"
                        last_restart = now
                        ok = restart_server(reason)
                        failures = 0 if ok else FAIL_LIMIT
                        last_state = "healthy" if ok else "restart-failed"
                    elif last_state != "cooldown":
                        _log(
                            "연속 실패 감지 — 재시작 과열 방지 대기",
                            state="cooldown", consecutive_failures=failures,
                            retry_after_seconds=round(RESTART_COOLDOWN - (now - last_restart)),
                        )
                        last_state = "cooldown"
                elif last_state != "degraded":
                    _log(
                        "앱 서버 응답 지연 감지 — 오탐 방지 재확인",
                        state="degraded", consecutive_failures=failures,
                    )
                    last_state = "degraded"
        except Exception as exc:
            _log("감시 루프 오류 — 감시는 계속함", state="guard-error",
                 error=f"{type(exc).__name__}: {exc}"[:240])
        time.sleep(CHECK_SECONDS)


if __name__ == "__main__":
    raise SystemExit(main())
