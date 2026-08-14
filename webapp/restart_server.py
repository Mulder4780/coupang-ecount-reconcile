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
                             capture_output=True, text=True, timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
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
                           capture_output=True, timeout=30, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass
    # 포트가 풀릴 때까지 잠깐 — 안 기다리면 새 서버가 포트를 못 잡고 죽는다.
    for _ in range(20):
        if not running():
            return True
        time.sleep(0.5)
    return not running()


def answering(timeout=1.0):
    """포트가 실제로 **답을 주나.** 프로세스 목록에 있는 것만으로는 부족하다 —
    소켓을 잡기 전 몇 초 동안 터널은 502 를 돌려준다(2026-08-10 지시로 추가).
    PIN 은 넣지 않는다. 401 이 오면 그것도 '살아 있다'는 답이다."""
    import socket
    try:
        port = int(str(_port()).strip())
    except (TypeError, ValueError):
        return False
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=timeout):
            return True
    except OSError:
        return False


# ── 지금 사람이 쓰고 있나 (2026-08-13 지시: "입력중인데 서버 떨어지면 골치아프다") ──
#
# 이날 실사고: 15:53 에 이 스크립트를 돌렸는데 그때 오종현이 입력 중이었다. 재시작은
# 실측 5.7~9.3초이고 그동안 폰·PC 는 502 를 받는다([197]). 그런데 이 스크립트는
# **누가 쓰고 있는지 보지도 않고** 죽였다 — 고친 사람은 성공이라 읽고, 입력하던
# 사람만 화면이 무너지는 것을 본다.
IN_USE_MIN = 10          # 최근 이만큼 안에 화면을 만졌으면 '쓰는 중'으로 본다
DEFER_LOG = os.path.join(ROOT, "reports", "앱서버_재시작보류.json")


def _ts_epoch(s):
    """ux 표의 시각 한 칸을 epoch 로.

    ★ **두 형식이 섞여 있다**(실측 2026-08-13): 브라우저가 보낸
      `2026-08-13T07:01:56.977Z`(UTC) 39,600건 · `ux_add` 기본값인 로컬 naive
      `2026-08-11T21:43:42` 15건.
    ★ 그래서 **문자열로 비교하면 조용히 0건**이 된다 — UTC 07시가 로컬 16시보다
      작아서, **1.5분 전에 사람이 만진 화면까지 '아무도 없음'으로 읽힌다.**
      첫 판이 실제로 그랬다. 계기가 0 을 내면 아무도 의심하지 않는다([169]) —
      그러니 반드시 파싱해서 잰다. 못 읽는 칸은 세지 않고 '못 읽음'으로 남긴다.
    """
    from datetime import datetime
    t = str(s or "").strip()
    if not t:
        return None
    if t.endswith("Z"):
        t = t[:-1] + "+00:00"
    try:
        d = datetime.fromisoformat(t)
    except ValueError:
        return None
    if d.tzinfo is None:          # 로컬 시각으로 적힌 것
        d = d.astimezone()
    return d.timestamp()


def in_use(minutes=IN_USE_MIN):
    """최근 `minutes` 분 안에 앱을 만진 흔적이 있나.

    근거는 **이미 쌓이고 있는** `ledger_db` 의 `ux` 표다(실측 39,611건). 새 계기를
    만들지 않는다 — 만들면 그것이 안 채워지는 날 또 눈이 먼다([169]).

    돌려주는 것: {"읽음": bool, "건수": int, "분전": float|None, "왜": str}
    ★ **못 읽으면 '아무도 없다'가 아니다.** `읽음=False` 면 부르는 쪽은 안전한 쪽
      (멈춤)으로 간다 — DB 가 잠겼다는 이유로 남의 입력을 날려서는 안 된다.
    """
    try:
        if ROOT not in sys.path:
            sys.path.insert(0, ROOT)
        import ledger_db
        with ledger_db.conn() as c:
            rows = c.execute(
                "SELECT ts FROM ux WHERE kind IN ('view','tap','search','slow','error') "
                "ORDER BY id DESC LIMIT 400").fetchall()
        cut, newest, n = time.time() - minutes * 60, None, 0
        for (ts,) in rows:
            e = _ts_epoch(ts)
            if e is None:
                continue
            if newest is None or e > newest:
                newest = e
            if e >= cut:
                n += 1
        return {"읽음": True, "건수": n,
                "분전": None if newest is None else (time.time() - newest) / 60.0, "왜": ""}
    except Exception as exc:
        return {"읽음": False, "건수": 0, "분전": None, "왜": str(exc)[:120]}


def _unattended():
    """사람이 답할 수 있는 자리인가. 워치독·회차는 `pythonw` 로 돌아 stdin 이 없다."""
    if os.environ.get("COUPANG_UNATTENDED") == "1":
        return True
    try:
        return not sys.stdin.isatty()
    except Exception:
        return True


def _note_defer(why, u):
    """미룬 것을 **자국으로 남긴다.** 조용히 넘어가면 '왜 아직 옛 코드지'가 된다."""
    import json
    from datetime import datetime
    try:
        os.makedirs(os.path.dirname(DEFER_LOG), exist_ok=True)
        try:
            hist = json.load(open(DEFER_LOG, encoding="utf-8"))
        except Exception:
            hist = []
        hist = (hist if isinstance(hist, list) else [])[-49:]
        hist.append({"때": datetime.now().isoformat(timespec="seconds"),
                     "왜": why, "활동읽음": u.get("읽음"), "최근건수": u.get("건수"),
                     "분전": None if u.get("분전") is None else round(u["분전"], 1)})
        json.dump(hist, open(DEFER_LOG, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    except Exception:
        pass


def guard(force=False):
    """죽이기 **전에** 부른다. 내려도 되면 None, 미뤄야 하면 사람 말로 된 이유."""
    if force:
        return None
    u = in_use()
    if not u["읽음"]:
        return ("지금 누가 쓰고 있는지 확인하지 못했습니다(%s). "
                "확인 못 한 것을 '아무도 없다'로 치지 않습니다 — 그대로 둡니다." % (u["왜"] or "이유 모름"))
    if u["건수"] > 0:
        ago = "방금" if (u["분전"] or 0) < 1 else "%.0f분 전" % u["분전"]
        return ("%s까지 누가 앱을 쓰고 있었습니다(최근 %d분 안에 %d번). "
                "지금 내리면 그 사람 화면이 %d초쯤 끊깁니다."
                % (ago, IN_USE_MIN, u["건수"], 10))
    return None


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
    ap.add_argument("--force", action="store_true",
                    help="누가 쓰고 있어도 내린다(사람이 판단해서 쓴다)")
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
        u = in_use()
        if not u["읽음"]:
            print(f"  쓰는 사람 확인 못 함 — {u['왜']}")
        elif u["건수"]:
            print(f"  ★ 지금 누가 쓰고 있습니다 — 최근 {IN_USE_MIN}분 안에 {u['건수']}번"
                  + ("" if u["분전"] is None else f" (마지막 {u['분전']:.1f}분 전)"))
        else:
            print(f"  최근 {IN_USE_MIN}분 안에 쓴 흔적 없음 — 내려도 됩니다.")
        return 0

    # ★ 죽이기 **전에** 누가 쓰고 있는지 본다. 기본값은 **멈추는 것**이다.
    if cur:
        why = guard(a.force)
        if why:
            if _unattended():
                # 무인 회차는 사람이 답할 수 없다 — 물어보지 말고 미룬다.
                _note_defer("무인 호출 — " + why, in_use())
                print("미룹니다: " + why)
                print("  다음 회차가 다시 봅니다. 지금 꼭 내려야 하면: "
                      "python webapp/restart_server.py --force")
                return 3
            print("★ " + why)
            try:
                ans = input("  그래도 지금 내릴까요? (y = 내림 / 그 밖 = 그대로 둠): ")
            except (EOFError, KeyboardInterrupt):
                ans = ""            # 답을 못 받으면 **안전한 쪽**이다
            if ans.strip().lower() not in ("y", "yes"):
                _note_defer("사람이 미룸 — " + why, in_use())
                print("  그대로 두었습니다. 지금 꼭 내려야 하면 --force 를 붙이세요.")
                return 3

    if cur:
        print("끄는 중:", ", ".join(f"pid {p}(뜬 시각 {w})" for p, w in cur))
        if not stop([p for p, _ in cur]):
            print("★ 옛 서버가 안 꺼졌습니다. 안 끄고 새로 띄우면 포트를 못 잡습니다.")
            return 1
    else:
        print("떠 있는 서버 없음 — 새로 띄웁니다.")

    t0 = time.time()
    start()
    for _ in range(60):                    # 떴는지 확인하고 끝낸다
        time.sleep(0.5)
        now = running()
        if not now:
            continue
        # ★ **프로세스가 있는 것과 답을 주는 것은 다르다** (2026-08-10).
        #   여기서 곧장 '올라왔습니다'를 찍으면, 아직 소켓을 안 잡은 몇 초 동안
        #   폰(클라우드플레어 터널)은 **502** 를 받는다 — 고친 사람은 성공이라 읽고
        #   폰을 든 사람만 실패를 본다. 실패가 성공처럼 보이는 자리다.
        if not answering():
            continue
        pid, when = now[0]
        print(f"올라왔습니다 — pid {pid} · 뜬 시각 {when} "
              f"(응답까지 {time.time() - t0:.1f}초 — 그동안 폰은 502 를 받습니다)")
        print(f"  PC:  http://localhost:{_port()}   (폰은 같은 와이파이에서 PC 주소)")
        return 0
    print("★ 새 서버가 안 보이거나 답을 주지 않습니다. "
          "손으로 확인하세요: python webapp/app_server.py")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
