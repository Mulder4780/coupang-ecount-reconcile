# -*- coding: utf-8 -*-
"""앱 서버를 **끄고 다시 띄운다** (2026-08-08).

★ 왜 따로 필요한가 — `앱서버실행.bat` 은 **띄우기만** 한다. 이미 떠 있으면 새 서버는
  포트를 못 잡고 조용히 죽고, 옛 서버가 계속 응답한다. 그래서 **코드를 고쳐도 화면이
  안 바뀐다.** 2026-08-08 실측: 어제 20:48 에 뜬 서버가 그날 하루치 코드 변경
  (금액 기준·접수취소·갱신 개선·AS 기사 문)을 하나도 반영하지 못한 채 돌고 있었다.
  화면은 멀쩡히 숫자를 보여 주므로 아무도 옛 서버인 줄 몰랐다 — 조용한 사고다.

  띄우는 자리(app_server.main)가 이미 그 상황을 알아보고 안내를 찍지만, 그 안내는
  **새로 띄우려 한 사람만** 본다. 폰으로 앱을 쓰는 사람에게는 아무 표시가 없다.

쓰기:
  python webapp/restart_server.py          # 끄고 다시 띄운다
  python webapp/restart_server.py --status # 지금 뭐가 떠 있나만 본다
"""
import os, sys, time, subprocess, argparse

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER = os.path.join(ROOT, "webapp", "app_server.py")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# 서버가 물고 있는 코드 — 이 파일들 중 하나라도 서버보다 새것이면 옛 코드로 도는 것이다.
WATCHED = ("webapp/app_server.py", "webapp/index.html", "webapp/tech.html",
           "ecount_reconcile.py", "ledger_db.py")


def stale():
    """서버가 **옛 코드로 돌고 있나**. 돌고 있으면 (pid, 뜬시각, 더 새로운 파일들).

    ★ 이것이 오늘(2026-08-08) 반나절을 먹은 조용한 사고다. 서버는 멀쩡히 200 을 주고
      화면도 숫자를 보여 주는데, 그 코드가 어제 것이었다. 고친 사람만 모르고 있었다.
      `app_server.main` 도 이 상황을 알아보고 안내를 찍지만 **새로 띄우려 한 사람만**
      본다 — 폰으로 쓰는 사람에게도, 다음 세션에게도 아무 표시가 없었다.
    ★ 판단은 **파일 mtime 대 프로세스 시작시각**이다. git 커밋 시각이 아니다 —
      받아만 놓고 안 띄운 경우(pull 직후)까지 잡아야 한다.
    """
    cur = running()
    if not cur:
        return None
    pid, when = cur[0]
    started = _started_epoch(when)
    if started is None:
        return None
    newer = []
    for rel in WATCHED:
        p = os.path.join(ROOT, rel)
        try:
            if os.path.getmtime(p) > started + 5:      # 5초는 기동 중 저장 여유
                newer.append(rel)
        except OSError:
            pass
    return (pid, when, newer) if newer else None


def _started_epoch(when):
    """PowerShell 이 준 시각 문자열을 epoch 로. 형식이 기계마다 달라 여러 개를 시도한다."""
    from datetime import datetime
    s = str(when or "").strip()
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%m/%d/%Y %p %I:%M:%S", "%Y-%m-%d %p %I:%M:%S"):
        try:
            return datetime.strptime(s, fmt).timestamp()
        except ValueError:
            continue
    return None


def _port():
    """주소를 손으로 적지 않는다 — 서버가 정하는 포트를 그대로 읽는다.

    ★ 처음엔 8765 라고 적어 뒀는데 실제로는 8899 였다. 안내가 틀리면 사람은
      '서버가 안 떴다'고 결론짓는다(실제로 그렇게 됐다). 정본은 app_server.py 다.
    """
    try:
        for line in open(SERVER, encoding="utf-8"):
            if line.startswith("PORT ="):
                return line.rsplit("else", 1)[-1].strip().rstrip(")").strip()
    except Exception:
        pass
    return "8899"


def running():
    """지금 떠 있는 앱 서버 (pid, 뜬 시각) 목록. 나 자신은 빼고 센다."""
    me = os.getpid()
    ps = ("Get-CimInstance Win32_Process -Filter \"Name like '%python%'\" | "
          "Where-Object { $_.CommandLine -like '*app_server.py*' } | "
          "ForEach-Object { \"$($_.ProcessId)`t$($_.CreationDate)\" }")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=60).stdout
    except Exception:
        return []
    found = []
    for line in out.splitlines():
        parts = line.strip().split("\t")
        if len(parts) == 2 and parts[0].isdigit() and int(parts[0]) != me:
            found.append((int(parts[0]), parts[1]))
    return found


def stop(pids):
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=30)
        except Exception:
            pass
    # 포트가 풀릴 때까지 잠깐 — 안 기다리면 새 서버가 포트를 못 잡고 죽는다.
    for _ in range(20):
        if not running():
            return True
        time.sleep(0.5)
    return not running()


def start():
    exe = sys.executable or "python"
    # pythonw 로 띄우면 창이 안 뜬다(원래 이 서버가 그렇게 돌고 있었다).
    quiet = exe.replace("python.exe", "pythonw.exe")
    if not os.path.isfile(quiet):
        quiet = exe
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | \
        getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([quiet, "-u", SERVER], cwd=ROOT, env=env,
                     creationflags=flags, close_fds=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description="앱 서버 재시작")
    ap.add_argument("--status", action="store_true", help="지금 상태만 본다")
    a = ap.parse_args(argv)

    cur = running()
    if a.status:
        if not cur:
            print("앱 서버가 떠 있지 않습니다.")
        for pid, when in cur:
            print(f"  pid {pid} · 뜬 시각 {when}")
        s = stale()
        if s:
            print(f"★ 옛 코드로 돌고 있습니다 — {', '.join(s[2][:4])} 가 서버보다 새것입니다.")
            print("  고쳐도 화면이 안 바뀝니다:  python webapp/restart_server.py")
        elif cur:
            print("  코드는 최신입니다.")
        return 0

    if cur:
        print("끄는 중:", ", ".join(f"pid {p}(뜬 시각 {w})" for p, w in cur))
        if not stop([p for p, _ in cur]):
            print("★ 옛 서버가 안 꺼졌습니다. 안 끄고 새로 띄우면 포트를 못 잡습니다.")
            return 1
    else:
        print("떠 있는 서버 없음 — 새로 띄웁니다.")

    start()
    for _ in range(30):                    # 떴는지 확인하고 끝낸다
        time.sleep(0.5)
        now = running()
        if now:
            pid, when = now[0]
            print(f"올라왔습니다 — pid {pid} · 뜬 시각 {when}")
            print(f"  PC:  http://localhost:{_port()}   (폰은 같은 와이파이에서 PC 주소)")
            return 0
    print("★ 새 서버가 안 보입니다. 손으로 확인하세요: python webapp/app_server.py")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
