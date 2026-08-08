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
