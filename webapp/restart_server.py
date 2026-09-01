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


# ⚠ 이 목록은 "서버가 물고 있는 코드" 가 **아니다 — 그 일부다.**
#    2026-08-18 실측: `app_server.py` 가 import 하는 프로젝트 모듈은 **36개**인데
#    여기 적힌 `.py` 는 **3개**다(app_server·ecount_reconcile·ledger_db). 나머지
#    33개(work_flow·system_audit·camp_contacts·local_ai·notify·error_book …)는
#    바뀌어도 **아무 화면에도 안 뜨고** 서버는 옛 코드로 계속 200 을 준다 —
#    `[156]` 이 반나절을 먹었던 바로 그 모양이 그 33개에 대해서는 아직 열려 있다.
#    (분담판 [112] 의 '앱 서버가 옛 코드로 돌고 있었다' 가 실제로 이 구멍이다)
#
#    ★ 그렇다고 36개를 다 넣으면 안 된다 — 이 목록은 워치독의 **자동 재시작**도
#      몰기 때문에 커밋마다 서버가 갈리고, 그때마다 화면이 반쪽으로 내려간다
#      (2026-08-18 실측 · 분담판 [117]). **판정을 넓히는 것과 재시작을 넓히는 것은
#      다른 결정이다** — 짐작으로 함께 넓히지 않는다. 분담판 [118].
#: 서버가 물고 있는 것의 **씨앗**.  여기서 시작해 import 를 따라간다 —
#: 손으로 적는 목록은 언제나 뒤처진다(`[162]`).  `.html` 은 import 가 아니라
#: 디스크에서 읽는 파일이라 씨앗에만 있고 따라갈 것이 없다.
WATCHED = ("webapp/app_server.py", "webapp/index.html", "webapp/tech.html",
           "ecount_reconcile.py", "ledger_db.py")

#: 서버가 **요청마다 디스크에서 다시 읽는** 파일 — 고쳐도 **재시작이 필요 없다**.
#:
#: ★ 2026-08-21 (분담판 `[191]`).  전에는 이 둘을 파이썬 모듈과 똑같이 취급해
#:   `stale()` 이 '옛 코드'를 올렸고, 워치독 `heal_stale_server` 가 그것을 믿고
#:   **멀쩡한 서버를 내렸다** — 담당자에게 8초씩 502 다(`[197]` 실측).  화면을
#:   고칠 때마다 바뀌므로 이 가짜 경보가 **제일 자주** 떴다.
#: ★ 근거는 셋이고 전부 실측이다:
#:   ① 요청 핸들러가 `open(...).read()` 한다 — `app_server.py` index 8827 · tech 8799
#:   ② `build_id()` 는 부를 때마다 `os.stat` 한다 — 프로세스에 캐시되는 값이 없다
#:   ③ 화면 `checkBuild()` 가 그 값이 달라지면 **스스로 갱신을 제안한다**
#:      (`index.html` 의 `offerUpdate`) — 사람에게 알리는 일까지 이미 돌고 있다.
#: ★ 그러므로 **경보에서 뺀다.  다만 조용히 빼지 않는다**(`[169]`) —
#:   `live_changed()` 가 그대로 돌려주고 `--status` 가 '새로고침하면 됩니다'라 적는다.
#: ⚠ 이 목록에 넣기 전에 **그 파일을 프로세스가 물고 있지 않은지** 확인한다.
#:   물고 있는데 넣으면 `[156]` 의 그 사고(옛 코드로 200 을 주는 서버)가 되살아난다.
LIVE_FILES = ("webapp/index.html", "webapp/tech.html")

#: import 를 찾을 폴더(프로젝트 루트 기준).  `app_server` 는 `sys.path` 에
#: 루트와 `webapp` 을 넣고 돌므로 그 둘이 먼저다.
_SEARCH_DIRS = ("", "webapp/", "band/")


def watched_files(seed=None, start=None, root=None):
    """서버가 **실제로 물고 있는** 파일들 · 못 읽은 것 수.

    인자는 **검증이 임시 파일로 재기 위한 것**이다 — 실측 증거(`app_server.py`)를
    건드리지 않고 '새 모듈이 저절로 따라오나'를 잴 수 있어야 한다(`[247]`).

    ★ 2026-08-20 (분담판 `[118]`).  전에는 손으로 적은 다섯 개뿐이라
      **나머지를 고치면 아무 화면에도 안 떴다** — 서버는 옛 코드를 메모리에 물고
      200 을 주고 화면은 숫자를 보여 준다.  고친 사람만 모른다(`[156]` 의 그 사고가
      다섯 파일 밖에서는 그대로 살아 있었다는 뜻이다).
    ★ 목록을 **늘려 적지 않고 따라간다** — `app_server.py` 에서 시작해 이 프로젝트
      안의 import 를 재귀로 좇는다.  모듈이 늘면 저절로 따라온다.
    ★ **못 읽은 파일을 조용히 넘기지 않는다**(`[169]`) — 파싱이 깨진 파일은 감시
      밖인데, 그 사실을 안 적으면 '이상 없음'과 구별되지 않는다.  수를 같이 준다.
    """
    import ast
    base_root = root or ROOT
    out, unread = set(seed if seed is not None else WATCHED), 0
    seen, stack = set(), list(start or ["webapp/app_server.py"])
    while stack:
        rel = stack.pop()
        if rel in seen:
            continue
        seen.add(rel)
        try:
            with open(os.path.join(base_root, rel), encoding="utf-8") as fh:
                tree = ast.parse(fh.read())
        except (OSError, SyntaxError, ValueError, UnicodeDecodeError):
            unread += 1
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and not node.level and node.module:
                names = [node.module]
            else:
                continue
            for name in names:
                head = name.split(".")[0] + ".py"
                for base in _SEARCH_DIRS:
                    cand = base + head
                    if os.path.exists(os.path.join(base_root, cand)):
                        out.add(cand)
                        stack.append(cand)
                        break
    return sorted(out), unread


def newer_than_server():
    """서버가 뜬 **뒤에 바뀐 파일**을 두 갈래로 재서 한 번에 돌려준다.

    돌려주는 것: `(pid, 뜬시각, 재시작필요[], 새로고침이면됨[], 못읽은수)` · 없으면 None.

    ★ 재는 자리는 **여기 하나**다(`[162]`).  `stale()` 과 `live_changed()` 는 이
      결과를 갈라 보여 줄 뿐이다 — 각자 mtime 을 다시 재면 같은 순간에 두 답이 나온다.
    """
    cur = running()
    if not cur:
        return None
    pid, when = cur[0]
    started = _started_epoch(when)
    if started is None:
        return None
    newer, live = [], []
    # ★ 목록은 **따라가서 만든다**([162]·분담판 [118]) — 손으로 적은 다섯 개만 보던
    #   때는 나머지를 고쳐도 아무 화면에 안 떴다(실측 5 → 89개).
    files, unread = watched_files()
    for rel in files:
        p = os.path.join(ROOT, rel)
        try:
            if os.path.getmtime(p) > started + 5:      # 5초는 기동 중 저장 여유
                (live if rel in LIVE_FILES else newer).append(rel)
        except OSError:
            pass
    return (pid, when, newer, live, unread)


def stale_longrunner(mark, seed, start=None, root=None):
    """오래 사는 프로세스가 **제 코드보다 낡았나** — (pid, 뜬시각, 새파일[], 못읽은수).

    없거나 못 재면 `None`.  **모르면 안 간다**([169]) — 프로세스를 못 찾거나
    시작시각을 못 읽으면 갈지 않는다.  잘못 갈면 그 순간 앱이 무방비가 된다.

    ★ 왜 `newer_than_server()` 로 안 되나: 그것은 **앱 서버 전용**이다(씨앗이
      `app_server.py` 이고 프로세스도 그것만 찾는다).  그런데 이 프로젝트에는
      앱 서버 말고도 오래 사는 python 프로세스가 있고 — 서버 보호자
      (`webapp/server_guard.py`) · 터널 감시자(`webapp/tunnel_run.py`) —
      **그것들은 아무도 안 갈아 준다.**
    ★ 실측 2026-08-23 (형님 지시 "앱 구동에 문제되는 거 전부 찾아서"):
      보호자 파일은 13:24 인데 그 프로세스는 **08-22 20:17** 에 떴다 — 17시간 된
      코드다.  그래서 하루 전에 붙인 자국(`was_alive`)이 **한 줄도 안 찍혔고**,
      "왜 재시작이 잦은가"를 재려던 계기가 통째로 눈이 멀어 있었다.
      `[156]`(서버가 옛 코드로 200 을 준다)의 **보호자 판**이고, 자리만 다르다.
    ★ 더 나쁜 것은 **저절로는 절대 안 갈린다**는 점이다: `heal_server_guard` 는
      heartbeat 만 보고 "정상"이라 하고, 새로 띄우려 해도 singleton(8978)이
      막는다.  즉 사람이 손으로 죽이기 전에는 영원히 옛 코드다.
    ★ 파일 목록은 **따라가서 만든다**([162]·[118]) — 씨앗에서 import 를 재귀로
      좇으므로 모듈이 늘면 저절로 따라온다.  손으로 적으면 늘 뒤처진다.
    """
    cur = running(mark)
    if not cur:
        return None
    pid, when = cur[0]
    started = _started_epoch(when)
    if started is None:
        return None
    files, unread = watched_files(seed=seed, start=start or list(seed), root=root)
    base_root = root or ROOT
    newer = []
    for rel in files:
        try:
            if os.path.getmtime(os.path.join(base_root, rel)) > started + 5:
                newer.append(rel)
        except OSError:
            pass
    return (pid, when, newer, unread)


def stale():
    """서버가 **옛 코드로 돌고 있나**. 돌고 있으면 (pid, 뜬시각, 더 새로운 파일들).

    ★ 이것이 2026-08-08 반나절을 먹은 조용한 사고다. 서버는 멀쩡히 200 을 주고
      화면도 숫자를 보여 주는데, 그 코드가 어제 것이었다. 고친 사람만 모르고 있었다.
      `app_server.main` 도 이 상황을 알아보고 안내를 찍지만 **새로 띄우려 한 사람만**
      본다 — 폰으로 쓰는 사람에게도, 다음 세션에게도 아무 표시가 없었다.
    ★ 판단은 **파일 mtime 대 프로세스 시작시각**이다. git 커밋 시각이 아니다 —
      받아만 놓고 안 띄운 경우(pull 직후)까지 잡아야 한다.
    ★ **요청마다 읽는 파일은 여기 안 담는다**(`LIVE_FILES` · 분담판 `[191]`).
      담으면 '재시작해야 한다'가 **거짓말**이 되고, 그 거짓말을 워치독이 믿고
      멀쩡한 서버를 내린다 — 담당자에게 8초씩 502 다.  안 담는 대신
      `live_changed()` 가 말한다(`[169]` — 조용히 빼지 않는다).
    """
    m = newer_than_server()
    if not m:
        return None
    pid, when, newer, _live, unread = m
    if unread and newer:
        # ★ **못 읽은 것을 조용히 넘기지 않는다**([169]).  경보가 이미 섰을 때만
        #   덧붙인다 — 아무 일 없는 날까지 말하면 아무도 안 읽는다([170]).
        newer.append("(그 밖에 %d개는 못 읽어 감시 밖)" % unread)
    return (pid, when, newer) if newer else None


def live_changed():
    """**새로고침이면 되는** 화면 파일이 서버보다 새것인가 — 경보가 아니라 사실이다.

    ★ 경보로 올리지 않는다(`[170]`).  화면이 이미 `checkBuild()` 로 사람에게
      묻고 있으므로 인계 문서까지 매번 같은 말을 하면 진짜 경보가 묻힌다.
      대신 **묻는 사람에게는 답한다** — `--status` 가 이것을 적는다(`[169]`).
    """
    m = newer_than_server()
    return list(m[3]) if m else []


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


def running(mark="app_server.py"):
    """명령줄에 `mark` 가 든 python 프로세스 (pid, 뜬 시각) 목록. 나 자신은 뺀다.

    ★ 기본값이 앱 서버라 **옛 호출자는 한 글자도 안 바뀐다**.  인자를 둔 이유는
      상시 프로세스(서버 보호자·터널 감시자)도 **같은 자리**에서 찾게 하려는
      것이다([162]) — 찾는 법을 저마다 적으면 한쪽만 고쳐지고, 갈린 뒤에는
      어느 쪽이 맞는지 아무도 모른다.
    """
    me = os.getpid()
    ps = ("Get-CimInstance Win32_Process -Filter \"Name like '%python%'\" | "
          "Where-Object { $_.CommandLine -like '*" + mark + "*' } | "
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
    """포트가 실제로 **답을 주나 — 그리고 답한 것이 우리 앱인가.**

    프로세스 목록에 있는 것만으로는 부족하다 — 소켓을 잡기 전 몇 초 동안
    터널은 502 를 돌려준다(2026-08-10 지시로 추가).  PIN 은 넣지 않는다.

    ★ **소켓만 봐서는 남이 그 자리를 잡아도 '올라왔습니다'를 찍는다**
      (2026-09-01 실사고).  그날 옆 프로젝트 세션의 파일 서버가 8899 를 먼저
      잡았는데, 이 함수는 소켓이 열렸다는 이유로 **거짓 성공**을 찍었고
      고친 사람은 성공이라 읽었다 — 담당자만 남의 폴더 목록을 봤다.
    ★ 판정은 `server_guard.probe()` **한 곳**에서 빌린다([162]).  여기에 또
      적으면 사본이 되고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
    ★ **좁히지 않는다**([172]): 200 인데 우리 표시가 없을 때만 거짓이고,
      401 처럼 200 이 아닌 응답은 예전대로 '살아 있다'로 본다 —
      그것까지 거짓으로 치면 멀쩡한 재시작이 매번 실패로 끝난다.
    """
    import socket
    try:
        port = int(str(_port()).strip())
    except (TypeError, ValueError):
        return False
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=timeout):
            pass
    except OSError:
        return False
    # 소켓은 열렸다.  그 자리에 앉은 것이 우리 앱인지 한 번 더 묻는다.
    try:
        import sys as _sys
        _here = os.path.dirname(os.path.abspath(__file__))
        if _here not in _sys.path:
            _sys.path.insert(0, _here)
        from server_guard import probe, PORT_FOREIGN
        return probe(timeout=max(timeout, 1.5)) != PORT_FOREIGN
    except Exception:
        # 못 물어봤으면 예전 그대로 '살아 있다' — 모름을 실패로 치지 않는다([169]).
        return True


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


# ── 막혀 있는 중인가 ────────────────────────────────────────────────────────
# ★ **쓰는 중과 막혀 있는 중은 다른 사실이다** (2026-08-24 오종현 실사고).
#   서버가 옛 코드로 4시간 반을 돌았고 그 사이 담당자가 리모컨 불출을 못 했다.
#   워치독은 `guard()` 의 '누가 쓰고 있다' 하나로 재시작을 미뤘는데 —
#   **그 사람이 쓰고 있던 것이 아니라 막혀서 같은 단추를 계속 누르고 있었다.**
#   그러면 '쓰는 중' 신호가 오히려 세지고 → 더 안 갈리고 → 더 막힌다.
#   **스스로를 강화하는 고장**이라 시간이 갈수록 확신에 차서 틀린다.
#   실측 기록: 11:38 `/api/remote/request · HTTP_ERROR:400 · "…한도 3개를 넘습니다"`
#   (뒤에 길안내가 없다 = 옛 코드) · 코드는 11:05 에 고쳐져 있었다.
BLOCK_CODES = ("400", "409", "422", "500")
# ★ **502·503·504·네트워크 끊김은 절대 안 센다.** 그것은 재시작 **자체가** 만드는
#   오류다([197] 실측 6.7초). 세는 순간 재시작 → 오류 → 또 재시작 이 되어
#   담당자 화면이 영원히 끊긴다 — 고치려던 것보다 나쁘다([172]).
# ★ 401·403 도 안 센다 — 권한은 코드를 갈아도 안 풀린다([290]).
# * 회차를 띄우는 길(`/api/run/...`)도 안 센다 — 거기 409 는 "이미 다른 작업이
#   돌고 있다"는 뜻이고, 그 잠금은 **다른 프로세스**가 쥔 것이라 앱 서버를
#   갈아도 안 풀린다. 세면 회차가 도는 동안 담당자 화면만 끊는다([172]).
#   실측 2026-08-25 15:44 — 표본이 `/api/run/daily HTTP_ERROR:409 · 다른 작업
#   실행 중` 하나뿐인데 그것으로 재시작 갈래가 섰다.
BLOCK_SKIP_TARGETS = ("/api/live-state", "/api/ping", "/api/sync-health",
                      "/api/error_help", "/api/run/")
BLOCK_MIN = 15           # 최근 이만큼 안에 맞은 거절만 센다
FORCE_COOLDOWN_MIN = 20  # 이 갈래로 한 번 갈면 이만큼은 다시 안 간다


def _http_code(detail):
    """`HTTP_ERROR:400 · …` 에서 코드만. 못 읽으면 None(모름 · [169])."""
    s = str(detail or "")
    i = s.find("HTTP_ERROR:")
    if i < 0:
        return None
    d = s[i + 11:i + 14]
    return d if d.isdigit() else None


def blocked_now(minutes=BLOCK_MIN):
    """담당자가 **쓰는 길에서 거절을 맞고 있나**.

    돌려주는 것: {"읽음": bool, "건수": int, "표본": [str], "왜": str}
    ★ **못 읽으면 '안 막혔다'가 아니다**([169]) — `읽음=False` 면 부르는 쪽은
      예전대로 미룬다. 모름을 근거로 남의 화면을 끊지 않는다.
    """
    try:
        if ROOT not in sys.path:
            sys.path.insert(0, ROOT)
        import ledger_db
        with ledger_db.conn() as c:
            rows = c.execute(
                "SELECT ts,target,detail FROM ux WHERE kind='error' "
                "ORDER BY id DESC LIMIT 300").fetchall()
        cut, n, sample = time.time() - minutes * 60, 0, []
        for ts, target, detail in rows:
            e = _ts_epoch(ts)
            if e is None or e < cut:
                continue
            tgt = str(target or "")
            if any(tgt.startswith(x) for x in BLOCK_SKIP_TARGETS):
                continue
            if _http_code(detail) not in BLOCK_CODES:
                continue
            n += 1
            if len(sample) < 3:
                sample.append(("%s %s" % (tgt, str(detail or "")))[:150])
        return {"읽음": True, "건수": n, "표본": sample, "왜": ""}
    except Exception as exc:
        return {"읽음": False, "건수": 0, "표본": [], "왜": str(exc)[:120]}


def _forced_recently(minutes=FORCE_COOLDOWN_MIN):
    """이 갈래로 방금 갈았나 — 냉각. 못 읽으면 **갈았다고 친다**(안전한 쪽)."""
    import json
    from datetime import datetime
    try:
        hist = json.load(open(DEFER_LOG, encoding="utf-8"))
        if not isinstance(hist, list):
            return False
        for rec in reversed(hist[-30:]):
            if not isinstance(rec, dict) or rec.get("갈래") != "막힘":
                continue
            t = datetime.fromisoformat(str(rec.get("때")))
            return (datetime.now() - t).total_seconds() < minutes * 60
        return False
    except FileNotFoundError:
        return False
    except Exception:
        return True


def _note_forced(b, st, u):
    """미루지 **않은** 것도 자국으로 남긴다 — 왜 그 사람 화면이 끊겼는지."""
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
                     "갈래": "막힘",
                     "왜": "쓰는 중이지만 거절을 맞고 있고 서버가 옛 코드다 — 미루지 않았다",
                     "거절건수": b.get("건수"), "표본": b.get("표본"),
                     "옛코드": list(st[2])[:6] if st and len(st) > 2 else [],
                     "최근활동": u.get("건수")})
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
        # ★ **막혀 있는 사람에게 미루는 것은 오히려 해롭다** (2026-08-24 오종현 실사고).
        #   두 조건이 같이 설 때만 간다 — 하나만으로는 근거가 안 선다:
        #     ① 사람이 **쓰는 길에서 거절을 맞고 있다**(400/409/422/500)
        #     ② 서버가 **옛 코드**다 — 갈면 달라질 수 있다는 뜻
        #   ②가 없으면 갈아도 같은 거절이 난다 — 그때 내리면 **멀줦한 사람 화면만
        #   8초 끊기고 문제는 그대로 남는다**([172] 의 틀린 지목).
        b = blocked_now()
        st = stale()
        if b["읽음"] and b["건수"] > 0 and st and not _forced_recently():
            _note_forced(b, st, u)
            return None
        ago = "방금" if (u["분전"] or 0) < 1 else "%.0f분 전" % u["분전"]
        return ("%s까지 누가 앱을 쓰고 있었습니다(최근 %d분 안에 %d번). "
                "지금 내리면 그 사람 화면이 %d초쯤 끊깁니다."
                % (ago, IN_USE_MIN, u["건수"], 10))
    return None


def _identity_env_keys():
    """앱 서버 자식에게 **물려주면 안 되는** 세션 신분 환경변수.

    목록은 `ai_claim.SID_ENV` 한 곳에서 온다([162]) — 여기 손으로 적어 두면
    Codex 쪽 키가 늘어난 날 그 갈래만 조용히 새어 나간다. 못 읽으면 아는 것만이라도
    지운다(빈손으로 물려주는 것보다 낫다) — 그리고 못 읽었다는 사실은 숨기지 않는다.
    """
    keys = ["CLAUDE_PID"]
    try:
        if ROOT not in sys.path:
            sys.path.insert(0, ROOT)
        import ai_claim
        keys.extend(ai_claim.SID_ENV)
    except Exception:
        keys.extend(["CLAUDE_CODE_SESSION_ID", "CLAUDE_CODE_HOST_SESSION_ID",
                     "CODEX_THREAD_ID", "CODEX_SESSION_ID", "AI_SESSION_ID"])
    return keys


def start():
    exe = sys.executable or "python"
    # pythonw 로 띄우면 창이 안 뜬다(원래 이 서버가 그렇게 돌고 있었다).
    quiet = exe.replace("python.exe", "pythonw.exe")
    if not os.path.isfile(quiet):
        quiet = exe
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    # ★ 서버에게 **제 신분**을 준다 — 대화창의 sid 를 물려주지 않는다
    #   (2026-08-19 실측 · 분담판 [156]).
    #   사람이 대화 창에서 이 스크립트를 부르면 자식(앱 서버)이
    #   `CLAUDE_CODE_SESSION_ID` 를 물려받는다. 그러면 서버가 `publish` 를 잡을 때
    #   `who=server · sid=<그 대화창>` 으로 적히고, 그 창의 자동 마무리
    #   (`session_wrapup --free-all` — PreCompact·SessionEnd)가 `_is_mine` 을 참으로
    #   읽어 **게시 중인 서버의 잠금을 푼다.** 그 순간 워치독 3시간 `--push` 와
    #   앱 서버 10분 `--cloud` 가 같은 `docs/data.enc`·git 을 동시에 만질 수 있다.
    #   [104] 가 '남의 것은 못 놓는다'로 막은 사고의 사촌인데, 그때는 옆 **창**만
    #   보고 **sid 를 물려받은 자식 프로세스**는 안 봤다.
    #   ⚠ `CLAUDE_PID` 도 같이 지운다 — 남기면 `ai_claim._is_dead` 가 서버가 아니라
    #     **대화창의 pid** 로 생사를 판정해, 그 창을 닫는 순간 게시 중인 서버의 점유가
    #     '죽음'으로 읽힌다. 지우면 `agent_pid` 가 0 이라 점유에 적힌 **그 서버의
    #     pid** 가 증거가 된다(같은 함수의 폴백).
    #   지운 뒤 서버 sid 는 `sid_of("<호스트>/manual")` 이라 스케줄러가 띄운 서버와
    #   같아진다 — 같아야 맞다. '이 호스트의 앱 서버'는 어느 길로 떴든 한 배우다.
    for _k in _identity_env_keys():
        env.pop(_k, None)
    # ★ DETACHED_PROCESS 만으로는 창이 뜬다 — `quiet` 가 pythonw 를 못 찾아
    #   python.exe 로 떨어지면 새 콘솔이 할당된다 (2026-08-14, 검증 [272]).
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | \
        getattr(subprocess, "DETACHED_PROCESS", 0) | \
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
        live = live_changed()
        if live:
            # ★ 재시작이 아니라 **새로고침**이다(분담판 [191]) — 이것을 '옛 코드'라
            #   부르면 워치독이 멀쩡한 서버를 내린다(담당자에게 8초씩 502).
            print(f"  화면 파일이 서버보다 새것입니다 — {', '.join(live[:4])}")
            print("  서버 재시작은 필요 없습니다 — 브라우저 새로고침이면 반영됩니다"
                  "(앱이 스스로 갱신을 제안합니다).")
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
