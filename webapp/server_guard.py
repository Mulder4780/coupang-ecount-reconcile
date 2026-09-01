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
import threading
import time
import urllib.request
import urllib.error
from datetime import datetime


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    # pythonw launches this file by its webapp/ path, so only webapp/ is on
    # sys.path.  Funnel repair lives at the project root and must also work in
    # the real no-console startup environment, not only when imported by tests.
    sys.path.insert(0, ROOT)
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
FIXED_FUNNEL_CHECK_SECONDS = 60
FUNNEL_FAIL_LIMIT = 3
FUNNEL_REPAIR_COOLDOWN = 300
STAFF_PATHS = ("/staff/ryu-jiyeong", "/staff/oh-jonghyeon")
STAFF_CHECK_SECONDS = 30
STAFF_FAIL_LIMIT = 3
STAFF_RESTART_COOLDOWN = 900
_FUNNEL_REPAIR_LOCK = threading.Lock()


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _atomic_json(path: str, value: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    # The public-Funnel repair runs in a worker so origin health checks never stop.
    # Give each writer its own temp file; a shared .tmp lets heartbeat and repair
    # race at os.replace and turns the guard's evidence into a false error.
    temp = path + f".{os.getpid()}.{threading.get_ident()}.tmp"
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


# ── 포트가 답하나 · 그리고 **답한 것이 우리 앱인가** (2026-09-01 실사고) ──
#
# 그날 09:12 에 옆 프로젝트 세션이 `python -m http.server 8899` 를 띄워 이 포트를
# **먼저** 잡았다.  윈도우는 같은 포트에 둘이 LISTEN 되게 두고 **먼저 잡은 쪽이
# 모든 요청을 받는다.**  그래서 09:29 에 우리 앱이 떠도 손님은 남의 파일 목록을
# 받았고, Tailscale Funnel 이 그것을 **공개 인터넷으로 그대로 내보냈다.**
#
# ★ 그때 이 감시자는 **이미 정확히 가르고 있었다** — `coupang-work` 를 본다.
#   잘못은 그 결과를 **'실패' 하나로 뭉갠 것**이다.  그러면 조치가 언제나
#   '앱 서버 재시작'인데, 재시작으로는 남이 쥔 포트를 되찾을 수 없다 —
#   실측 **17분간 12회 헛재시작**했고 그동안 담당자 화면도 계속 끊겼다.
#   조치는 갈래마다 달라야 한다([289]).
#
# 갈래 셋:
#   ok      우리 앱이 답했다
#   foreign 남이 이 자리를 잡았다 — 사람이 끊어야 풀린다
#   down    연결이 안 되거나 그 밖 — 예전 그대로 재시작 대상
#
# ★ `foreign` 의 근거는 둘뿐이다([172]):
#     (a) 200 인데 우리 표시가 없다
#     (b) **404/403** — 우리 앱이 살아 있으면 `/api/ping` 은 **반드시** 있다.
#         남의 정적 파일 서버가 그 경로에 주는 전형적인 답이 이것이다.
#   ⚠ (b) 는 실측이 가르쳐 줬다: 처음엔 (a) 만 봤는데, 오늘 사고를 흉내 내
#     `python -m http.server` 를 띄워 재니 **404 라서 `down` 으로 떨어졌다** —
#     곧 그 판정으로는 **오늘 난 사고를 못 잡는다.**  검사가 없었으면
#     '고쳤다'고 적고 넘어갔을 자리다([272]).
#   ⚠ 500·502·401 은 넣지 않는다.  우리 앱도 그것을 줄 수 있고, 넣으면
#     **진짜 죽은 서버를 안 살린다**(좁히는 것도 고장이지만 넓히는 것도 고장이다).
PORT_OK = "ok"
PORT_FOREIGN = "foreign"
PORT_DOWN = "down"
APP_MARK = b"coupang-work"


def probe(timeout: float = PING_TIMEOUT) -> str:
    try:
        req = urllib.request.Request(
            f"http://127.0.0.1:{PORT}/api/ping?t={time.time_ns()}",
            headers={"Cache-Control": "no-cache"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status != 200:
                return PORT_DOWN
            return PORT_OK if APP_MARK in response.read(2048) else PORT_FOREIGN
    except urllib.error.HTTPError as exc:
        # 404/403 은 '그 경로가 없다'는 뜻이다 — 우리 앱이면 있을 수 없다.
        return PORT_FOREIGN if exc.code in (403, 404) else PORT_DOWN
    except Exception:
        return PORT_DOWN


def ping(timeout: float = PING_TIMEOUT) -> bool:
    """옛 호출자는 한 글자도 안 바뀐다 — 갈래를 bool 로 접어 준다."""
    return probe(timeout) == PORT_OK


def port_holder():
    """이 포트를 잡은 **우리 앱이 아닌** 프로세스.  조회가 실패하면 None([169]).

    ★ **죽이지 않는다.**  이름으로 죽이면 그 글자를 명령줄에 담은 무관한
      프로세스까지 죽는다(2026-08-13 에 실제로 PowerShell 을 통째로 죽였다 · [399]).
      여기는 **누가 잡았는지 말하는 자리**까지이고, 끊는 것은 사람이 정한다.
    """
    ps = (
        "Get-NetTCPConnection -LocalPort %d -State Listen -EA SilentlyContinue | "
        "ForEach-Object { $q = Get-CimInstance Win32_Process "
        "-Filter \"ProcessId=$($_.OwningProcess)\" -EA SilentlyContinue; "
        "if ($q) { '{0}|{1}|{2}' -f $q.ProcessId, $q.Name, $q.CommandLine } }"
    ) % PORT
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True, text=True, timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            return None
    except Exception:
        return None
    out = []
    for line in (result.stdout or "").splitlines():
        parts = line.strip().split("|", 2)
        if len(parts) != 3 or not parts[0].isdigit():
            continue
        cmd = parts[2]
        if "app_server.py" in cmd:
            continue                      # 우리 앱이다 — 남이 아니다
        out.append({"pid": int(parts[0]), "name": parts[1], "cmd": cmd[:400]})
    return out


def _holder_text(holder) -> str:
    """사람이 읽는 한 줄.  **못 읽었으면 0곳이라 하지 않는다**([169])."""
    if holder is None:
        return "누가 잡았는지 확인 못 함"
    if not holder:
        return "그 포트를 잡은 남의 프로세스를 못 찾음(조회는 됐다)"
    return " / ".join("pid %s %s" % (x["pid"], x["cmd"][:160]) for x in holder)


def staff_centers_alive(timeout: float = PING_TIMEOUT) -> dict[str, bool]:
    """Probe both staff entry pages, not only the generic process ping.

    A process can answer /api/ping while a route or the HTML response is broken.
    Reading only the first 16 KiB keeps this cheap while proving that the real
    work-center shell reached 류지영 and 오종현's paths.
    """
    result = {}
    for path in STAFF_PATHS:
        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{PORT}{path}?guard={time.time_ns()}",
                headers={"Cache-Control": "no-cache", "User-Agent": "CSOS-Server-Manager"},
            )
            with urllib.request.urlopen(req, timeout=timeout) as response:
                body = response.read(16384)
                result[path] = (
                    response.status == 200
                    and b"Coupang Service Operations System" in body
                )
        except Exception:
            result[path] = False
    return result


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
        # * **살아 있는 서버를 죽이는 것인지 로그가 말하게 한다** (2026-08-23).
        #   실측 7일: 재시작 08-18 10회 · 08-20 15회 · **08-21 20회** 인데 그중
        #   몇 개가 오탐이었는지 **알 길이 없었다** — `old_pids` 는 성공 로그의
        #   구조화 칸으로만 나가고 그 칸은 STATUS 파일에 **마지막 것만** 남는다.
        #   그래서 과거 재시작을 되짚을 수 없다([228] 이 회차에 대해 말한 것과
        #   같은 자리다: 자국이 없으면 왜인지는 영영 모른다).
        # * 가르는 근거는 **프로세스가 살아 있었나** 하나다. `missing`(프로세스
        #   없음)은 진짜 죽음이라 재시작이 옳다. 그런데 프로세스가 **살아 있는데**
        #   ping 만 늦은 것은 죽음이 아니라 **늘어진 것**일 수 있다 — 이 프로젝트는
        #   회차가 Z:(SMB)를 훑는 동안 앱이 늘어진다. 그때 죽이면 회차도 끊기고
        #   사람 화면도 끊긴다([197] 실측 9.3초).
        # ! 동작은 **한 톨도 안 바꾼다** — 재시작은 예전 그대로 한다. 여기서
        #   문을 새로 달면 진짜 죽었을 때 못 살린다([172] · 되돌릴 수 없는 쪽).
        #   먼저 **재고**, 오탐이 실제로 많으면 그때 고친다.
        _log("앱 서버 재시작 — 직전 프로세스가 살아 있었다(pid %s) · %s"
             % (",".join(str(p) for p in old), reason),
             state="restarting", old_pids=old, reason=reason, was_alive=True)
        _stop_server(old)
    else:
        _log("앱 서버 재시작 — 직전 프로세스가 없었다(진짜 죽음) · %s" % reason,
             state="restarting", old_pids=[], reason=reason, was_alive=False)
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


def fixed_funnel_alive(timeout: int = 5) -> bool:
    """Check the public phone path, not this PC's private Tailscale route."""
    try:
        from tailscale_serve import FIXED_HOST, public_funnel_alive
        return public_funnel_alive(FIXED_HOST, timeout=timeout)
    except Exception:
        return False


def repair_fixed_funnel_async() -> bool:
    """Refresh the fixed Funnel without blocking the ten-second origin guard.

    tailscale_serve is the one canonical owner of public DNS/SNI probing and the
    reset/re-register sequence.  This guard only decides *when* to call it.
    """
    if not _FUNNEL_REPAIR_LOCK.acquire(blocking=False):
        return False

    def worker():
        try:
            from tailscale_serve import ensure_public_funnel
            ok, repaired = ensure_public_funnel(repair=True)
            if ok:
                _log(
                    "고정 Funnel 외부경로 자동 복구 완료",
                    state="healthy", funnel_ok=True, funnel_repaired=bool(repaired),
                )
            else:
                _log(
                    "고정 Funnel 외부경로 복구 실패 — 냉각 후 다시 시도",
                    state="funnel-repair-failed", funnel_ok=False,
                )
        except Exception as exc:
            _log(
                "고정 Funnel 복구 작업 오류 — 앱 서버 감시는 계속함",
                state="funnel-repair-error",
                error=f"{type(exc).__name__}: {exc}"[:240],
            )
        finally:
            _FUNNEL_REPAIR_LOCK.release()

    threading.Thread(target=worker, daemon=True, name="fixed-funnel-repair").start()
    return True


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
    next_fixed_funnel_check = 0.0
    funnel_failures = 0
    last_funnel_repair = 0.0
    next_staff_check = 0.0
    staff_failures = 0
    last_staff_restart = 0.0
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

            verdict = probe()
            local_ok = verdict == PORT_OK
            if local_ok:
                if failures or last_state != "healthy":
                    _log("앱 서버 정상", state="healthy", consecutive_failures=0)
                failures = 0
                last_state = "healthy"

                if now >= next_staff_check:
                    next_staff_check = now + STAFF_CHECK_SECONDS
                    staff_health = staff_centers_alive()
                    failed_staff = [path for path, ok in staff_health.items() if not ok]
                    if not failed_staff:
                        if staff_failures:
                            _log(
                                "류지영·오종현 업무센터 정상 복귀",
                                state="healthy", staff_centers=staff_health,
                                consecutive_staff_failures=0,
                            )
                        staff_failures = 0
                    else:
                        staff_failures += 1
                        if staff_failures == 1:
                            _log(
                                "담당자 업무센터 응답 지연 감지 — 오탐 방지 재확인",
                                state="staff-degraded", staff_centers=staff_health,
                                failed_staff_paths=failed_staff,
                                consecutive_staff_failures=staff_failures,
                            )
                        if (staff_failures >= STAFF_FAIL_LIMIT
                                and now - last_staff_restart >= STAFF_RESTART_COOLDOWN):
                            last_staff_restart = now
                            last_restart = now
                            ok = restart_server(
                                "류지영·오종현 업무센터 %d회 연속 실패" % staff_failures
                            )
                            staff_failures = 0 if ok else STAFF_FAIL_LIMIT
                            last_state = "healthy" if ok else "restart-failed"

                # The fixed URL is a different failure domain from the local origin.
                # A healthy 127.0.0.1 must never be accepted as proof that a phone can enter.
                if now >= next_fixed_funnel_check:
                    next_fixed_funnel_check = now + FIXED_FUNNEL_CHECK_SECONDS
                    if fixed_funnel_alive():
                        if funnel_failures:
                            _log("고정 Funnel 외부경로 정상 복귀", state="healthy",
                                 funnel_ok=True, consecutive_funnel_failures=0)
                        funnel_failures = 0
                    else:
                        funnel_failures += 1
                        if funnel_failures == 1:
                            _log("고정 Funnel 외부경로 지연 감지 — 오탐 방지 재확인",
                                 state="funnel-degraded", funnel_ok=False,
                                 consecutive_funnel_failures=funnel_failures)
                        if (funnel_failures >= FUNNEL_FAIL_LIMIT
                                and now - last_funnel_repair >= FUNNEL_REPAIR_COOLDOWN):
                            if repair_fixed_funnel_async():
                                last_funnel_repair = now
                                funnel_failures = 0
                                _log("고정 Funnel 공개경로 비동기 재등록 시작",
                                     state="funnel-repairing", funnel_ok=False)
            else:
                failures += 1
                if verdict == PORT_FOREIGN:
                    # ★ 남이 이 자리를 잡았다 — 재시작해도 절대 안 낫는다.
                    #   냉각 타이머(`last_restart`)도 안 건드린다: 남이 물러난
                    #   순간 곧바로 정상으로 돌아와야 한다.
                    #   ⚠ 10초마다 찍지 않는다 — 로그가 넘치면 아무도 안 읽는다([170]).
                    if last_state != "port-taken":
                        holder = port_holder()
                        _log(
                            "앱 포트 %d 를 남이 잡았다 — 재시작으로는 안 낫는다 · %s"
                            % (PORT, _holder_text(holder)),
                            state="port-taken", consecutive_failures=failures,
                            port_holders=holder,
                        )
                        last_state = "port-taken"
                    continue
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
