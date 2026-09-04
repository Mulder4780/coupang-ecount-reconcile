# -*- coding: utf-8 -*-
"""
watchdog.py — 자가치유 워치독 (무인 운영의 심장)
===================================================
30분마다 작업 스케줄러가 실행. 사람이 손대지 않아도:
  1. 앱 서버 죽음 감지 → 자동 재시작
  2. 외부접속 터널 죽음/주소소실 감지 → 자동 재시작
  3. 30일 지난 리포트 자동 정리
관리대장 버전 정리는 `ledger_versions.py` 한 곳에서만 수행한다.
모든 조치는 reports/watchdog_log.txt 에 기록.

실행:  python watchdog.py            # 점검+복구 1회
       python watchdog.py --dry      # 점검만(복구·이동 없음)
"""
import sys, os, re, glob, json, time, subprocess, urllib.request
from datetime import datetime, timedelta
from operation_window import input_window_label, is_input_window
from proc_guard import run_tree

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PYW = PY.replace("python.exe", "pythonw.exe")
LOG = os.path.join(ROOT, "reports", "watchdog_log.txt")
SYNC_CONTRACT = os.path.join(ROOT, "reports", "화면동기화_감시.json")
PORT = 8899


def log(msg):
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    line = f"[{datetime.now():%m-%d %H:%M}] {msg}"
    # pythonw.exe에는 stdout이 없다. 무창 스케줄러에서 print가 먼저 죽으면 아래
    # 파일 자국까지 못 남는다 — 사람이 손으로 돌릴 때만 콘솔에도 보인다.
    if sys.stdout is not None:
        print(line)
    try:
        open(LOG, "a", encoding="utf-8").write(line + "\n")
    except Exception:
        pass


# ── 판단 로직 (합성 테스트 대상 — 순수 함수) ─────────────────────
def gap_note(last_line, now, expect_min=30, slack_min=15):
    """직전 기록과 지금 사이가 너무 벌어졌으면 그 사실을 남긴다.

    ★ 2026-07-28: 워치독이 16:27~20:43 **4시간 넘게 안 돌았고**, 그 사이 터널 주소가
      죽어 폰 접속이 끊겼다. 그런데 로그에는 '정상'만 줄줄이 남아 있어 아무도 몰랐다
      (작업 스케줄러가 놓친 실행을 따라잡지 않는 설정이었다 — StartWhenAvailable=False).
      쉰 것 자체보다 **쉰 걸 아무도 모르는 것**이 문제다. 공백은 로그에 적어 둔다."""
    if not last_line:
        return ""
    m = re.match(r"\[(\d{2})-(\d{2}) (\d{2}):(\d{2})\]", last_line)
    if not m:
        return ""
    mo, d, hh, mm = (int(x) for x in m.groups())
    try:
        prev = now.replace(month=mo, day=d, hour=hh, minute=mm, second=0, microsecond=0)
    except ValueError:
        return ""
    if prev > now:                      # 해가 바뀌면 작년 기록이다
        prev = prev.replace(year=prev.year - 1)
    gap = (now - prev).total_seconds() / 60.0
    if gap <= expect_min + slack_min:
        return ""
    return "★ 워치독이 %.0f분 쉬었다(예상 %d분) — 그 사이 터널이 죽어도 아무도 못 고친다" % (gap, expect_min)


def last_log_line(path=None):
    try:
        lines = [l for l in open(path or LOG, encoding="utf-8").read().splitlines() if l.strip()]
        return lines[-1] if lines else ""
    except Exception:
        return ""


def pick_archive(version_files, keep=1):
    """[(버전, 경로)] → OLD로 옮길 경로 목록 (현재 정책은 최신 1개 제외)"""
    s = sorted(version_files, key=lambda x: x[0], reverse=True)
    return [p for _, p in s[keep:]]


def pick_old_reports(files_with_mtime, days=30, protect=("agent_status.json", "tunnel_url.txt", "watchdog_log.txt")):
    """[(경로, mtime)] → 삭제 대상 (보호 파일 제외, days일 초과)"""
    cut = time.time() - days * 86400
    return [p for p, mt in files_with_mtime
            if mt < cut and os.path.basename(p) not in protect]


# ── 점검·복구 ─────────────────────────────────────────────
def ping():
    try:
        with urllib.request.urlopen(f"http://localhost:{PORT}/api/ping", timeout=5) as r:
            return b"coupang-work" in r.read()
    except Exception:
        return False


def proc_running(image):
    """PowerShell Get-Process 기반(로케일·wmic 무관, 신뢰성 우선)"""
    name = image.replace(".exe", "")
    try:
        out = run_tree(
            ["powershell", "-NoProfile", "-Command",
             f"(Get-Process {name} -ErrorAction SilentlyContinue).Count"],
            timeout=20, drain_timeout=5,
        ).stdout.strip()
        return out.isdigit() and int(out) > 0
    except Exception:
        return False


def start_hidden(script):
    # ★ `DETACHED_PROCESS`(0x8) 는 **창 없는 깃발이 아니다** — 부모 콘솔을 안 물려받는다는
    #   뜻일 뿐이라, `PYW` 가 없어 `PY`(python.exe)로 떨어지면 **새 콘솔이 뜬다.**
    #   창을 없애는 깃발은 `CREATE_NO_WINDOW` 다 (2026-08-14, 검증 [272]).
    subprocess.Popen([PYW if os.path.exists(PYW) else PY, os.path.join(ROOT, script)],
                     cwd=ROOT,
                     creationflags=(getattr(subprocess, "CREATE_NO_WINDOW", 0)
                                    | 0x00000008))  # + DETACHED_PROCESS


def heal_server(dry):
    if ping():
        return heal_stale_server(dry)
    if dry:
        return "서버 죽음(dry — 복구 생략)"
    # ★ **이름으로 죽이지 않는다** (2026-08-18 · 분담판 [115]).
    #   예전 `kill_by_cmdline("app_server.py")` 는 그 글자를 명령줄에 담은 **모든**
    #   파이썬을 죽였다 — 합성검증의 데모 서버(`--demo --port 18899`)·8898·8897 시험
    #   서버·사람이 띄운 개발 서버까지. 관문(`[6]`)이 아무 이유 없이 빨개질 수 있는
    #   자리였고, 죽인 쪽에는 아무 자국도 안 남는다. `kill_stale_tunnel` 이 `[85]`
    #   에서 배운 것과 같은 교훈이다 — **그 이름은 남의 명령줄에도 있다.**
    # ★ **판정은 한 곳이다**([162]). `server_guard` 가 이미 운영 서버만 고른다
    #   (`--demo` 제외 · `--port` 가 있으면 8899 일 때만). 여기서 그 규칙을 베껴
    #   쓰면 포트를 옮기는 날 한쪽만 고쳐진다.
    # ★ **4초는 짧다** — 실측 재시작은 9.3초다(`[197]`). 옛 코드는 성공한 재시작을
    #   매번 '실패(다음 주기 재시도)'로 적었다. `restart_server` 는 프로세스가 뜬 것이
    #   아니라 **진짜 ping 이 올 때까지** 본다.
    try:
        from webapp import server_guard
    except Exception as exc:
        return "서버 죽음 — 복구기를 못 불렀다: %s" % str(exc)[:60]
    # ★ **못 읽은 것을 '없다'로 치지 않는다**([169]). 프로세스 목록을 못 읽으면
    #   그것은 '재시작 실패'가 아니라 **'잘못 죽이지 않으려고 안 한 것'** 이다.
    #   뭉쳐 적으면 사람이 없는 서버 고장을 찾아 나선다([172]).
    if server_guard.server_pids() is None:
        return "서버 죽음 — 프로세스 목록을 못 읽어 재시작 보류(다음 주기 재시도)"
    if server_guard.restart_server("watchdog"):
        return "서버 재시작 → 성공"
    return "서버 재시작 → 실패(다음 주기 재시도)"


def _heal_stale_guard_code(dry):
    """보호자가 옛 코드로 돌면 **갈아 준다**.  갈 것이 없으면 `None`.

    ★ **앱 서버는 안 내린다** — 바뀌는 것은 감시자뿐이라 류지영·오종현 화면은
      한 순간도 안 끊긴다(항상 담당자 업무가 먼저다).  감시 공백은 죽이고
      곧바로 띄우는 몇 초뿐이고, 그동안 터널 감시자가 두 번째 줄로 받친다.
    ★ **pid 로만 죽인다** — 이름(`CommandLine -like`)으로 죽이면 그 글자를 명령줄에
      담은 **무관한 프로세스까지** 죽는다(2026-08-13 에 `kill_stale_tunnel` 이
      내 PowerShell 을 통째로 죽인 그 자리다).
    ★ **모르면 안 간다**([169]) — 프로세스를 못 찾거나 시각을 못 읽으면 그대로 둔다.
      멀쩡한 감시자를 헛되이 갈면 그 순간이 앱의 사각지대가 된다.
    """
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        import restart_server as _rs
        got = _rs.stale_longrunner("server_guard.py", ("webapp/server_guard.py",))
    except Exception as exc:                       # 판정 실패를 "정상"이라 하지 않는다
        return "서버 관리 에이전트 코드 나이 확인 못 함: %s" % str(exc)[:50]
    if not got:
        return None
    pid, when, newer, unread = got
    if not newer:
        return None
    if dry:
        return "서버 관리 에이전트가 옛 코드(%s 외 %d개 · dry)" % (newer[0], len(newer) - 1)
    try:
        subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"],
                       capture_output=True, timeout=30,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        from webapp.tunnel_run import ensure_server_guard
        # `max_age=-1` 이라야 곧바로 띄운다 — 방금 죽였어도 heartbeat 파일은
        # 아직 새것이라, 기본값(120)으로 부르면 "살아 있다"고 보고 안 띄운다.
        ok = ensure_server_guard(max_age=-1)
    except Exception as exc:
        return "서버 관리 에이전트 교체 실패: %s" % str(exc)[:50]
    return ("서버 관리 에이전트를 새 코드로 갈았다(%s 외 %d개)" % (newer[0], len(newer) - 1)
            if ok else "서버 관리 에이전트 교체 뒤 시작 실패")


def heal_server_guard(dry):
    """Third recovery line for the lightweight server manager.

    The guard and tunnel supervisor normally revive each other.  The scheduled
    30-minute watchdog is deliberately independent, so it can restore the guard
    even when both long-lived processes disappeared together.
    """
    status = os.path.join(ROOT, "reports", "server_guard_status.json")
    try:
        age = time.time() - os.path.getmtime(status)
        if age <= 120:
            # 살아 있다 — 그런데 **그 코드가 새것인가**를 여기서 한 번 더 묻는다.
            # 안 물으면 보호자를 고쳐도 **영영 반영이 안 된다**(2026-08-23 실측:
            # 파일 13:24 · 프로세스 08-22 20:17 · 하루 전에 붙인 자국이 0건).
            # heartbeat 만 보던 때는 "정상"이라 답하고, 새로 띄워도 singleton 이
            # 막으므로 사람이 손으로 죽이기 전에는 안 갈린다([156] 의 보호자 판).
            return _heal_stale_guard_code(dry) or "서버 관리 에이전트 정상"
    except OSError:
        age = None
    if dry:
        return "서버 관리 에이전트 heartbeat 없음(dry)"
    try:
        from webapp.tunnel_run import ensure_server_guard
        ok = ensure_server_guard(max_age=120)
        return "서버 관리 에이전트 자동 시작" if ok else "서버 관리 에이전트 시작 실패"
    except Exception as exc:
        return "서버 관리 에이전트 확인 오류: %s" % str(exc)[:60]


def sync_contract_checks(server_src, ui_src):
    """새 기능도 공통 저장→revision→재조회 계약에 들어오는지 싼 정적 계기로 잰다.

    함수·화면 이름을 전부 열거하지 않는다. 새 POST는 기본 편입, 새 자료 GET은
    registerDataSection 한 곳, 새 view는 전체 section fallback이라는 **구조**를 잰다.
    그래서 기능이 늘어도 감시 코드를 매번 고치지 않는다.
    """
    checks = {
        "새 자료 POST 기본 편입": (
            "새 POST는 기본적으로 동기화 대상" in server_src
            and '_mark_live_mutation(getattr(self, "path", ""))' in server_src
        ),
        "전 업무센터 공통 입력": (
            "STAFF_ENTRY_PERMISSIONS = {" in server_src
            and "for slug in STAFF_CENTERS" in server_src
            and "_ALL_STAFF_ENTRY_FIELDS" in server_src
        ),
        "대표 브리핑 AppStore overlay": (
            'data["as"] = list(works.get("as")' in server_src
            and 'data["pm"] = list(works.get("pm")' in server_src
            and "*_app_db_stamp()" in server_src
        ),
        "신규 자료기능 단일 등록문": (
            "function registerDataSection(key,def)" in ui_src
            and "typeof def.apply!=='function'" in ui_src
            and "DATA_SECTION_STATE[k]=" in ui_src
        ),
        "신규 화면 자동 흡수": (
            "return byView[view]||Object.keys(DATA_SECTION_DEFS);" in ui_src
        ),
        "담당자·달력 전용 재조회": (
            "refreshStaffCenter(targetRevision" in ui_src
            and "refreshCalendarData(targetRevision" in ui_src
        ),
        "전 업무센터 공통 카드": all(x in ui_src for x in (
            "if(staffSlug) injectOhUpload();",
            "if(staffSlug) injectRemoteCard();",
            "if(staffSlug) injectRyuTodo();",
        )),
        "5초 백그라운드 revision loop": (
            "const LIVE_STATE_POLL_MS=5000" in ui_src
            and "queueLiveViewCatchup(v)" in ui_src
        ),
        "무인 감시 revision 응답": (
            'if p == "/api/sync-health":' in server_src
            and '"live_write_seq": state.get("live_write_seq")' in server_src
            and "state = get_live_state()" in server_src
        ),
    }
    return checks


def sync_contract_status(probe=True):
    """기존 워치독이 읽을 앱 전체 동기화 계약 상태. 업무값은 기록하지 않는다."""
    try:
        server_src = open(os.path.join(ROOT, "webapp", "app_server.py"),
                          encoding="utf-8").read()
        ui_src = open(os.path.join(ROOT, "webapp", "index.html"),
                      encoding="utf-8").read()
        checks = sync_contract_checks(server_src, ui_src)
    except Exception as exc:
        checks = {"코드 계약 읽기": False}
        read_error = str(exc)[:180]
    else:
        read_error = ""
    live = {"확인": False, "ok": None, "revision": "", "live_write_seq": None}
    if probe:
        try:
            with urllib.request.urlopen(
                    f"http://localhost:{PORT}/api/sync-health?t={int(time.time())}",
                    timeout=6) as response:
                payload = json.loads(response.read().decode("utf-8", "replace"))
            live = {"확인": True, "ok": bool(payload.get("ok")),
                    "revision": str(payload.get("revision") or "")[:24],
                    "live_write_seq": payload.get("live_write_seq")}
            checks["실행 서버 revision 응답"] = bool(
                live["ok"] and live["revision"] and live["live_write_seq"] is not None)
        except Exception as exc:
            live = {"확인": False, "ok": False, "revision": "",
                    "live_write_seq": None, "왜": str(exc)[:180]}
            checks["실행 서버 revision 응답"] = False
    problems = [name for name, ok in checks.items() if not ok]
    return {"ok": not problems, "확인시각": datetime.now().astimezone().isoformat(timespec="seconds"),
            "검사": checks, "문제": problems, "서버": live,
            "읽기오류": read_error,
            "규칙": "새 POST는 자동 편입 · 새 GET은 registerDataSection · 새 view는 전체 section fallback"}


def watch_sync_contract(dry):
    """기존 30분 워치독에 동기화 구조 감시를 얹는다. 창·팝업·브라우저는 열지 않는다."""
    report = sync_contract_status(probe=not dry)
    if not dry:
        os.makedirs(os.path.dirname(SYNC_CONTRACT), exist_ok=True)
        tmp = SYNC_CONTRACT + ".%d.tmp" % os.getpid()
        try:
            with open(tmp, "w", encoding="utf-8") as out:
                json.dump(report, out, ensure_ascii=False, indent=2)
            os.replace(tmp, SYNC_CONTRACT)
        except Exception as exc:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            return "화면동기화 감시 기록 실패: %s" % str(exc)[:60]
    if report["ok"]:
        return "화면동기화 계약 정상(%d계기)" % len(report["검사"])
    return "★ 화면동기화 계약 깨짐: " + ", ".join(report["문제"][:3])


def _code_claim_holder():
    """다른 **살아 있는** 세션이 지금 '코드·설정 변경'(code) 을 잡고 있나 (2026-09-04, [368]).

    ★ 왜 mtime 만으로는 모자란가 — 아래 60초 문은 '편집이 끝났다' 와 '저장 사이라
      잠깐 조용하다' 를 **못 가른다**. 실측 2026-09-04: 옆 창(codex)의 저장 시각이
      08:55:06 → 08:58:12 → 09:00:18 곧 **약 3분 간격**이라, 60초 창은 그 사이로
      깨끗이 지나간다. 그러면 **반쯤 저장된 코드가 류지영·오종현 앱 서버로 올라간다**
      — 이 함수 독스트링이 스스로 "그게 더 나쁘다" 라고 적어 둔 바로 그 자리다.

    ★ 점유는 짐작이 아니라 **선언**이다. mtime 은 "방금 파일이 바뀌었다" 까지이고,
      점유판은 그 세션이 **제 손으로** "나는 지금 코드를 고치는 중" 이라고 적은 것이다.
      근거의 세기가 다르므로 이쪽을 같이 본다.

    ⚠ 파일별이 아니다 — `LOCKS` 의 `code` 는 자원 이름 **하나**이고 어느 파일인지는
      안 적힌다(실측: 이 저장소에 `code:<파일>` 형태 점유는 없다). 그래도 이 판단에는
      족하다 — 묻는 것이 "누가 지금 코드를 고치나" 하나이기 때문이다.

    ★ 판정을 새로 만들지 않는다([162]) — 생사는 `ai_claim._is_dead`, 주인은
      `ai_claim._is_mine` 을 그대로 빌린다. 여기서 다시 판정하면 `--take` 와
      이 자리가 같은 점유를 두고 서로 다른 답을 한다.
    ★ **내 세션 것은 안 센다**([172]) — 사람이 `code` 를 잡은 창에서 워치독을 손으로
      돌리면 제 점유에 막혀 **영영 못 고친다**. 좁히는 것도 고장이다.
    ★ **죽은 세션 것도 안 센다** — 크레딧이 소진돼 멈춘 창의 점유는 편집 중이 아니다.
    ★ **못 읽으면 `None`**([169]) — 그때는 예전처럼 60초 문 하나로 간다. 점유판 하나가
      깨졌다고 자가치유가 통째로 멈추면 옛 코드가 며칠을 사는 쪽이 더 나쁘다.

    돌려주는 것: (누가, sid, 몇분째) 또는 None.
    """
    try:
        sys.path.insert(0, ROOT)
        import ai_claim
        cur = ai_claim.load().get("code")
        if not isinstance(cur, dict):
            return None
        if ai_claim._is_mine(cur, cur.get("who") or ""):
            return None
        if ai_claim._is_dead(cur):
            return None
        mins = int((time.time() - cur.get("at", 0)) / 60)
        return (cur.get("who") or "?", cur.get("sid") or "옛형식", mins)
    except Exception:
        return None


def heal_stale_server(dry):
    """살아 있지만 **옛 코드로 도는** 서버를 스스로 새 코드로 올린다 (2026-08-08).

    ★ 죽은 서버는 원래 잡고 있었다. 그런데 **살아 있는데 옛 코드인 경우**는 아무도
      안 봤다 — 서버는 200 을 주고 화면도 숫자를 보여 주므로 정상으로 보인다.
      2026-08-08 하루에만 세 번 그 상태가 됐고(어제 20:48 서버가 하루치 변경을 통째로
      못 실은 채 돌았다), 그중 두 번은 **옆 세션이 파일을 고쳐서** 생겼다.
      사람이 고칠 때마다 기억해서 눌러야 하는 일은 결국 안 눌린다.

    ★ 안전장치 — 함부로 끄지 않는다:
      · 파일이 **1분 이상** 안정된 뒤에만 올린다. 편집 중간에 끄면 반쯤 저장된 코드로
        올라간다(그게 더 나쁘다).
      · 30분 주기 안에서 **한 번만** 시도한다. 실패가 반복되면 그냥 두고 인계에 남긴다
        — 계속 끄고 켜면 쓰는 사람이 아무것도 못 한다.
      · 재시작 뒤 응답을 확인한다. 못 살아나면 그 사실을 그대로 적는다.
    """
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        import restart_server
        s = restart_server.stale()
        if not s:
            return "서버 정상"
        _pid, _when, newer = s
        # 편집이 끝났는지 — 가장 최근에 바뀐 파일이 60초 넘게 조용해야 한다.
        newest = max((os.path.getmtime(os.path.join(ROOT, f))
                      for f in newer if os.path.isfile(os.path.join(ROOT, f))),
                     default=0)
        if time.time() - newest < 60:
            return "서버 옛코드 — 방금 편집 중이라 미룸(%s)" % ", ".join(newer[:2])
        # ★ [368] mtime 이 조용해도 **점유판이 편집 중이라 말하면** 미룬다.
        #   60초 문은 저장 사이의 침묵과 편집 끝을 못 가른다(실측 저장 간격 3분).
        holder = _code_claim_holder()
        if holder:
            return ("서버 옛코드 — %s 세션[%s]이 코드를 고치는 중이라 미룸(%d분째): %s"
                    % (holder[0], holder[1], holder[2], ", ".join(newer[:2])))
        if dry:
            return "서버 옛코드(dry — 재시작 생략): %s" % ", ".join(newer[:3])
        # ★ 사람이 쓰고 있으면 restart_server 가 **스스로 미룬다**(exit 3, 2026-08-13).
        #   입력 중에 서버를 내리면 그 사람 화면이 10초쯤 끊긴다 — 옛 코드로 도는 것보다
        #   나쁘다. 미룬 것을 **'실패'라고 적지 않는다**: 실패로 적으면 사람이 없는
        #   고장을 찾아 나서고, 진짜 실패와 구별이 안 된다([169]).
        # ★ 스케줄러가 보이는 콘솔에서 python.exe 로 이 파일을 부르면 stdin 이 TTY다.
        # restart_server 는 그 환경을 사람 실행으로 오해해 ``지금 내릴까요? (y)`` 를
        # 물었고, 30분 워치독이 답을 기다린 채 영원히 멈췄다(2026-08-14 실사고).
        # 워치독이 부른 재시작은 언제나 무인 회차임을 **호출 경계에서** 명시한다.
        old_unattended = os.environ.get("COUPANG_UNATTENDED")
        os.environ["COUPANG_UNATTENDED"] = "1"
        try:
            rc = restart_server.main([])
        finally:
            if old_unattended is None:
                os.environ.pop("COUPANG_UNATTENDED", None)
            else:
                os.environ["COUPANG_UNATTENDED"] = old_unattended
        if rc == 3:
            return ("서버 옛코드 — 지금 쓰는 사람이 있어 미룸(다음 주기 재시도): %s"
                    % ", ".join(newer[:3]))
        ok = ping()
        return ("서버 옛코드 → 재시작 %s (%s)"
                % ("성공" if (rc == 0 and ok) else "실패(다음 주기 재시도)",
                   ", ".join(newer[:3])))
    except Exception as exc:
        return "서버 코드나이 확인 실패: %s" % str(exc)[:40]


def heal_band_evidence(dry):
    """'없음 확인' 근거가 **추월됐으면** 스스로 무효로 만든다 (2026-08-11 지시).

    ★ `[217]` 은 읽는 쪽 둘(수집 계획·인계 문서)이 추월된 근거를 **거르게** 했다.
      거르는 것만으로는 근거가 틀린 채로 남아, 바로잡는 길이 사람이 밴드 피드를 열어
      `real_latest.py --latest` 를 적어 주는 것뿐이었다. 그 한 줄이 마지막 사람 몫이었다.
    ★ **붙여넣기 파일보다 먼저** 온다. 파일에 어떤 번호를 담을지를 정하는 것이 이
      근거이므로, 순서가 뒤집히면 그 회차는 **틀린 근거로 만든 목록**을 사람 손에 쥐여
      준다 — 없는 번호 한 개가 21초다.
    ★ **밴드에 접속하지 않는다.** 캐시와 근거 파일만 읽고 근거 한 장만 고친다 —
      수집이 아니므로 코딩 세션 금지 규칙과 어긋나지 않는다(`heal_stale_pastefiles` 와 같다).
    """
    band_dir = os.path.join(ROOT, "band")
    try:
        if band_dir not in sys.path:
            sys.path.insert(0, band_dir)
        import real_latest as RL
        fixed = RL.heal(apply=not dry)
    except Exception as exc:                        # 근거 하나 때문에 회차를 세우지 않는다
        return "밴드 근거 확인 실패: %s" % str(exc)[:60]
    if not fixed:
        return "밴드 근거 정상"
    return "밴드 근거 추월 정정%s: %s" % (
        "(dry)" if dry else "",
        ", ".join("%s(%s→모름)" % (f["밴드"], f["이전"]) for f in fixed[:3]))


def heal_stale_pastefiles(dry):
    """붙여넣기 파일이 **옛 수집 JS**를 담고 있으면 다시 만든다 (2026-08-08).

    ★ 서버가 옛 코드로 도는 것과 **똑같은 종류의 조용한 사고**다. 수집 규칙을 고쳐도
      사람 손에 가는 것은 디스크에 있는 그 파일이다. 2026-08-08 에 댓글 수집을
      붙였는데 band/*_붙여넣기_*.js 네 개는 전부 그 이전에 만들어진 것이었다 —
      그대로 붙여넣었으면 **댓글이 한 건도 안 들어오는데 수집은 성공으로 끝났다.**
      개수도 날짜도 멀쩡해서 아무도 몰랐을 것이다.
    ★ 판단은 **파일 mtime 대 만드는 쪽 mtime** 이다(내용 비교가 아니다) —
      회차 번호는 매번 달라지므로 내용은 원래 다르다.
    ★ **'만드는 쪽'은 수집기만이 아니다** (2026-08-11). 예전에는 `grab_posts.js` 하나만
      봤는데, 파일에 **어떤 번호를 담을지**를 정하는 것은 `recheck_plan`·`make_oneclick`
      이다. 그래서 없는 번호 40개를 담던 규칙을 고쳐도 **디스크의 그 파일은 그대로**
      남았다 — 사람은 여전히 옛 목록을 붙여넣고 14분을 버린다. [162] 와 똑같은 모양이
      한 겹 위에서 반복된 것이다. 이제 셋 중 **가장 최근**을 기준으로 본다.
    ★ 이것은 **수집이 아니라 파일 만들기**다. 캐시를 읽기만 하고 밴드에 접속하지
      않는다 — 코딩 세션이 해도 되는 일이다(CLAUDE.md 의 수집 금지와 어긋나지 않는다).
    """
    import glob as _g
    band_dir = os.path.join(ROOT, "band")
    js = os.path.join(band_dir, "grab_posts.js")
    try:
        js_mt = os.path.getmtime(js)
    except OSError:
        return "붙여넣기 확인 생략(grab_posts.js 없음)"
    # ★ **번호를 고르는 쪽 mtime 은 make_oneclick 이 만든 파일에만 댄다** (2026-08-11).
    #   첫판은 모든 붙여넣기 파일에 댔는데, `댓글채우기_*` 는 comment_backfill 이,
    #   `재수집_*` 은 recollect 가 만든다. 남의 파일까지 '낡음'으로 몰면 여기서 지워지고
    #   **make_oneclick 은 그 파일을 다시 만들지 않는다** — 실측으로 두 개가 그렇게
    #   사라졌다(다음 09:50 회차 전까지 사람 손에 아무것도 안 남는다).
    #   낡음의 기준은 **그 파일을 만드는 쪽**이어야 한다.
    mk_mt = js_mt
    for maker in ("make_oneclick.py", "recheck_plan.py"):
        try:
            mk_mt = max(mk_mt, os.path.getmtime(os.path.join(band_dir, maker)))
        except OSError:
            pass                       # 없는 파일은 기준이 못 된다(있는 것만으로 판단)
    # ★ **기준은 한 곳에서 정한다** (2026-08-18 실사고 — 분담판 [56]).
    #   전에는 낡음을 `mk_mt`(만드는 쪽 최신)로 고르고 **새로워졌는지는 `js_mt`**
    #   (grab_posts.js) 로만 봤다. 그래서 `recheck_plan.py`(08-18 10:28) 가
    #   `grab_posts.js`(08-12 00:53) 보다 새로우면 **같은 파일이 낡았으면서 동시에
    #   새로웠다** — 30분마다 골라서, 아무것도 안 바뀌었는데, "새로 1개"라고 적었다.
    #   실측: `수집_붙여넣기_84789192.js` 가 08-16 00:14 그대로인 채 13:58·14:34·
    #   14:58 세 회차 연속 성공으로 보고됐다. 워치독은 result=0 으로 끝난다.
    #   ★ 2026-08-11 에 **낡음 기준만** 셋으로 넓히고 성공 기준을 안 따라가게 둔 것이
    #     원인이다 — 같은 판단을 두 곳에서 하면 언젠가 갈린다(`[162]`).
    def _기준(path):
        return mk_mt if os.path.basename(path).startswith("수집_") else js_mt

    old = [p for p in _g.glob(os.path.join(band_dir, "*붙여넣기_*.js"))
           if os.path.getmtime(p) < _기준(p)]
    if not old:
        return "붙여넣기 파일 최신"
    names = ", ".join(os.path.basename(p) for p in old[:3])
    if dry:
        return "붙여넣기 옛 JS(dry — 생성 생략): %s" % names
    made, gone, failed = 0, 0, 0
    for p in old:
        base = os.path.basename(p)
        band = base.rsplit("_", 1)[-1][:-3]
        try:
            if not base.startswith("수집_"):
                # 남의 회차가 만든 파일이다 — 재수집은 08:00, 댓글채우기·댓글은 09:50
                # 회차가 **대상 번호를 정한다.** 여기서 대상까지 새로 고르면 그 회차의
                # 판단을 가로챈다(make_oneclick 을 불러 봤자 제 이름의 파일만 쓴다).
                # 그래서 **지운다.** 다음 회차가 새 JS 로 다시 만든다 —
                # 옛 JS 를 남겨 두는 것보다 없는 편이 낫다.
                os.unlink(p)
                gone += 1
                continue
            run_tree([sys.executable, os.path.join(band_dir, "make_oneclick.py"),
                      "--band", band], cwd=ROOT, timeout=180, drain_timeout=10)
            # ★ **끝난 코드로 성공을 판단하지 않는다.** 훑을 것이 없는 밴드에서는
            #   생성기가 아무 파일도 안 쓰고 정상 종료한다(0). 그걸 성공으로 세면
            #   낡은 파일이 그대로 남은 채 "새로 만들었다"고 적힌다 — 거짓 보고다.
            #   실제로 파일이 새로워졌는지 mtime 으로 본다.
            if os.path.exists(p) and os.path.getmtime(p) >= _기준(p):
                made += 1
            else:
                # 만들 것이 없다 = 이 파일은 있을 이유가 없다. 남겨 두면 사람이
                # 옛 JS 를 붙여넣는다. 지우면 필요할 때 다시 만들어진다.
                try:
                    os.unlink(p)
                    gone += 1
                except OSError:
                    failed += 1
        except Exception:
            failed += 1
    # ★ **지운 것을 '새로 만든 것'과 같은 숫자에 담지 않는다**(`[169]`).
    #   섞으면 "새로 2개"로 읽히는데 실제로는 "2개 지웠다"일 수 있다 — 사람은
    #   붙여넣을 파일이 준비된 줄 안다. 지운 것은 다음 회차가 새 규칙으로 만든다.
    부분 = ["새로 %d개" % made] if made else []
    if gone:
        부분.append("지움 %d개(만들 것이 없어 — 다음 회차가 새로 만든다)" % gone)
    if failed:
        부분.append("실패 %d개" % failed)
    if not 부분:
        부분.append("아무것도 못 함")
    return "붙여넣기 옛 JS %d개 → %s (%s)" % (len(old), " · ".join(부분), names)


def heal_fixed_funnel(dry):
    """Check the same public path a phone uses and refresh stale Funnel TLS."""
    try:
        from tailscale_serve import (
            FIXED_HOST, ensure_public_funnel, hostname, public_funnel_alive,
        )
        actual_host = hostname()
        if not actual_host:
            return "고정 Funnel 스킵 — Tailscale 로그인 없음"
        if actual_host != FIXED_HOST:
            return "고정 Funnel 주소 불일치 — 자동 주소변경 금지"
        if public_funnel_alive(FIXED_HOST):
            return "고정 Funnel 휴대폰 경로 정상"
        if dry:
            return "고정 Funnel 휴대폰 경로 죽음(dry)"
        ok, repaired = ensure_public_funnel(repair=True)
        if ok:
            return "고정 Funnel 공개경로 재등록 → 성공" if repaired else "고정 Funnel 정상"
        return "고정 Funnel 공개경로 재등록 → 실패(다음 주기 재시도)"
    except Exception as exc:
        return f"고정 Funnel 검사 오류: {str(exc)[:40]}"


def tunnel_alive(url):
    """프로세스 개수가 아니라 **실제 응답**으로 판정.
    (Get-Process는 1개일 때 Count가 비어 나와 살아있는 터널을 죽었다고 오판 →
     불필요한 재시작으로 공개 주소가 바뀌던 실사고가 있었다)"""
    if not url:
        return False
    try:
        from net_probe import probe
        return probe(url.rstrip("/") + "/api/ping", 12)[0]
    except Exception:
        return False


def heal_tunnel(dry):
    url_f = os.path.join(ROOT, "reports", "tunnel_url.txt")
    url = ""
    try:
        url = open(url_f, encoding="utf-8").read().strip()
    except Exception:
        pass
    if tunnel_alive(url):
        return f"터널 정상 ({url[8:40]}...)"
    if dry:
        return "터널 죽음(dry)"
    # ★ 좀비 정리를 먼저 한다.
    #   tunnel_run은 포트 8977 싱글톤 락을 쓴다. 예전 tunnel_run이 살아 있으면
    #   새로 띄워도 "이미 실행 중"으로 즉시 끝나 **아무것도 고쳐지지 않는다**.
    #   실제로 cloudflared는 살아 있는데 주소만 만료된 채 이틀 방치됐다(2026-07-27).
    killed = kill_stale_tunnel()
    if not os.path.exists(os.path.join(ROOT, "webapp", "cloudflared.exe")):
        return f"터널 스킵 — cloudflared 없음{killed}"
    start_hidden(os.path.join("webapp", "tunnel_run.py"))
    return f"터널 재시작 지시{killed} (새 주소는 고정주소가 자동으로 따라감)"


def kill_stale_tunnel():
    """이 프로젝트의 죽은 터널만 정리한다.

    프로세스 이름이 ``cloudflared.exe`` 라는 이유만으로 전부 죽이면 같은 PC에서 도는
    다른 서비스 터널까지 같이 끊긴다. 정확한 ``webapp/tunnel_run.py`` 명령행과 그
    자식, 또는 이 앱 포트(8899)를 가리키는 quick-tunnel만 대상으로 삼는다.

    ★ **경로가 명령줄에 있다는 것만으로는 근거가 안 된다** (2026-08-13 실사고 · 분담판
      [85]). 그 파일을 편집기로 열어 두거나, 그 경로를 인자로 넘긴 grep·검사 명령도
      명령줄에 같은 글자를 담는다. 실측으로 이 고침을 확인하던 PowerShell 이 통째로
      죽었다(exit 255 · 출력 0). **이름으로 죽일 때는 그 이름이 남의 명령줄에도
      있다고 생각한다** — 그래서 **python 이 실행 중인 것**만 주인으로 친다.
    """
    script = os.path.normcase(os.path.abspath(os.path.join(ROOT, "webapp", "tunnel_run.py")))
    script_ps = script.replace("'", "''")
    ps = (
        "$all = @(Get-CimInstance Win32_Process); "
        f"$owners = @($all | Where-Object {{ $_.CommandLine -and "
        f"$_.Name -and (@('python.exe','pythonw.exe') -contains $_.Name.ToLower()) -and "
        f"[IO.Path]::GetFullPath(('{script_ps}')) -and "
        f"$_.CommandLine.ToLower().Contains(('{script_ps}').ToLower()) }}); "
        "$ownerIds = @($owners | ForEach-Object { [int]$_.ProcessId }); "
        "$targets = @($owners); "
        "$targets += @($all | Where-Object { $_.Name -eq 'cloudflared.exe' -and ("
        "$ownerIds -contains [int]$_.ParentProcessId -or "
        "($_.CommandLine -and ($_.CommandLine -like '*127.0.0.1:8899*' -or "
        "$_.CommandLine -like '*localhost:8899*'))) }); "
        "$targets | Sort-Object ProcessId -Unique | ForEach-Object { "
        "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $_.ProcessId }"
    )
    try:
        r = run_tree(["powershell", "-NoProfile", "-Command", ps],
                     timeout=60, drain_timeout=10)
        n = len([x for x in (r.stdout or "").split() if x.strip().isdigit()])
        return f" (좀비 {n}개 정리)" if n else ""
    except Exception:
        return ""


def snapshot_handoff(dry):
    """세션 인계 스냅샷 — **세션이 갑자기 끊겨도 여기까지는 남는다.**

    종료 체크리스트는 끝낼 시간이 있을 때만 지켜진다. 컨텍스트가 차거나 크레딧이 끊기면
    그럴 기회가 없고, 그때 점유·큐·임시파일이 방치된 채 남는다. 워치독이 30분마다
    상태를 적어 두면 다음 세션은 최대 30분 전 상태에서 이어받는다.
    ★ 읽기 전용이다 — 무엇이 걸렸는지 적을 뿐, 점유를 풀거나 큐를 반영하지 않는다.
      상대 AI 가 일하는 중일 수 있어 함부로 가로채면 안 된다."""
    if dry:
        return "세션인계(dry)"
    try:
        r = run_tree([PY, os.path.join(ROOT, "session_handoff.py"), "--snapshot"],
                     cwd=ROOT, timeout=120, drain_timeout=10)
        line = [x for x in (r.stdout or "").splitlines() if "세션인계" in x]
        return line[-1] if line else "세션인계 갱신"
    except Exception as e:
        return f"세션인계 실패: {str(e)[:40]}"


def clean_reports(dry):
    files = [(p, os.path.getmtime(p)) for p in glob.glob(os.path.join(ROOT, "reports", "*"))
             if os.path.isfile(p)]
    targets = pick_old_reports(files, days=30)
    if not dry:
        for p in targets:
            try:
                os.remove(p)
            except Exception:
                pass
    return f"오래된 리포트 {len(targets)}개 정리{'(dry)' if dry else ''}"


def lock_ledger_archive(dry):
    """관리대장 보관본이 안 잠겨 있으면 잠근다 — **엑셀은 저장용이다**(2026-08-28 지시).

    ★ 왜 그물이 필요한가 — `[474]` 의 배선은 **반쪽이었다.**
      vN+1 을 만드는 도구가 다섯인데(`ledger_writer`·`expand_rows`·`fix_ids`·
      `dedupe_rows`·`findings_sheet`) 잠금을 건 것은 `ledger_writer` 하나뿐이다.
      그런데 **2026-08-27 v621 을 만든 것은 `expand_rows`** 였다 — 곧 카톡 보류로
      행을 늘릴 때마다 새 최신본이 열린 채로 남는다.
      다섯 곳에 각각 붙이면 사본이 다섯이 된다([162]) — 그래서 **여기 한 곳**이
      30분마다 "안 잠겼으면 잠근다". 어느 도구가 만들었든 걸린다.

    ★ **사람이 쓰는 중이면 안 잠근다**(`check_in_use=True`) — `~$` 잠금파일이나
      최근 저장이 보이면 물러난다. 쓰는 중인 파일을 잠그면 그 사람의 입력이
      날아간다([104] 류지영 우선). 다음 회차가 다시 본다.
    ★ **읽기는 한 톨도 안 좁아진다** — 읽기 전용은 여는 것을 막지 않는다.
    ★ 못 잠가도 **회차를 안 죽인다** — 그러나 이미 잠겨 있으면 **아무 말도 안 한다**([170])."""
    try:
        import archive_lock
        import workbook_patch
    except Exception as exc:
        return "보관본 잠금: 못 불렀다(%s)" % exc
    if not archive_lock.enabled():
        return ""                      # 사람이 껐다 — 고장이 아니다
    try:
        path, ver = workbook_patch.latest_master()
    except BaseException as exc:       # SystemExit 도 받는다(latest_master 는 exit 한다)
        return "보관본 잠금: 관리대장을 못 찾았다(%s)" % str(exc)[:60]
    st = archive_lock.is_locked(path)
    if st is True:
        return ""                      # 이미 잠겼다 — 정상까지 말하면 아무도 안 읽는다
    if st is None:
        return "보관본 잠금: 상태를 못 읽었다(v%s)" % ver
    if dry:
        return "보관본 잠금: v%s 가 열려 있다(미리보기 — 안 잠갔다)" % ver
    ok, why = archive_lock.lock(path, check_in_use=True)
    if ok:
        return "보관본 v%s 를 읽기 전용으로 잠갔다 — 값은 앱에서 고친다" % ver
    return "보관본 v%s 를 안 잠갔다: %s" % (ver, why)


def sweep_files(dry):
    """쓸데없는 파일을 지운다 — **되돌릴 수 있는 갈래만**(2026-08-27 지시).

    형님 지시: "DB 및 앱 관련 데이터는 … 여기서 계속 관리하고 **(용량 커지지 않게)**".

    ★ 빠져 있던 것은 도구가 아니라 **부르는 자리**였다([328]) — `cleanup_files.py` 는
      있는데 부르는 코드가 **한 곳도 없어서**(실측 grep 0곳) 회차 산출물이 다시
      312.9MB 쌓였다. 코드가 있는 것과 그것이 도는 것은 다른 말이다.
    ★ **무엇을 지울지는 여기서 안 정한다**([162]) — `cleanup_files.AUTO_RULES` 한 곳이다.
      여기에 갈래 이름을 적으면 사본이 되어 한쪽만 고쳐진다.
    ★ **새 스케줄 작업을 안 만들었다** — 이미 도는 자리에 한 단계를 더한다([297]).
    ★ 실측 2026-08-27: **3.2초 · 321.7MB / 5,781개** · Z: 를 한 번도 안 만진다([168]).
    """
    try:
        import cleanup_files as CFL
    except Exception as e:
        return "파일 정리 못 함(%s)" % type(e).__name__
    try:
        return CFL.sweep(dry=dry)
    except Exception as e:
        # 청소 하나로 회차를 죽이지 않는다 — 그러나 조용히 넘기지도 않는다([169]).
        return "파일 정리 실패(%s: %s)" % (type(e).__name__, str(e)[:60])


MIRROR_BUDGET_S = int(os.environ.get("COUPANG_MIRROR_BUDGET_S") or 240)


#: 조율 표에 적는 이름 — **한 곳**이다([293]).  양보와 완주가 다른 이름으로 적히면
#: `audit()` 이 짝을 못 찾아 **한 건도 안 걸리면서 오류도 안 난다**([165]).
MIRROR_JOB = "원본이전"


def _mark_mirror_run(상태, 왜):
    """조율 표에 '이 회차에 돌았다'를 남긴다([293]).

    ★ **양보만 넘기고 완주를 안 넘기면 연속이 영원히 안 풀린다.**  실측 2026-08-27:
      17:56 부터 실제로 돌아 79개를 복사하고 있었는데 표는 `마지막 돎: 없음 ·
      5회 연속 양보` 라고 확언했고, 그 거짓이 매일 인계 맨 위에 올라왔다([170]).
      한쪽만 적는 표는 시간이 갈수록 더 확신에 차서 틀린다.
    ★ **자국을 못 남겨도 회차는 안 죽인다** — 이전 하나로 30분 회차를 세우지 않는다.
    """
    try:
        import coordinate as CO
        CO.record_run(MIRROR_JOB, 상태, 왜)
    except Exception:
        pass


def mirror_originals(dry):
    """원본 자료를 새 정본 자리로 **복사**한다 — 예산 안에서 (2026-08-27 지시).

    형님 지시: "`…\\2. CSOS DATA` 이 폴더에 원본데이터 폴더 번호 순서대로 계속
    저장하고 **기존 자료도 이쪽으로 전부 복사**해서 가져와서 관리해".

    ★ **한 회차에 다 안 한다** — 회선이 느려(실측 폴더 확인 1.3~28.9초 · ERP 폴더
      훑기만 102.2초) 한 번에 하려다 끊기면 **진도가 0** 이 된다([388]·[406]·[427]).
      예산이 다 되면 멈추고 다음 회차가 잇는다 — 일감은 유한하므로 **수렴한다**.
    ★ **도는 회차에는 양보한다**([313]) — 09:35 원본정리·09:50 대조가 Z: 를 통째로
      훑는 동안 같이 물면 **양쪽이 다 느려진다**(실측 2026-08-27: 겹쳤을 때 폴더 확인
      한 번이 **28.9초**, 한가할 때 **1.3초**). 물러나는 값은 회차 한 번이다.
    ★ **판정을 새로 만들지 않는다**([162]) — 무엇을 옮길지는 `data_mirror` 가 정하고
      도는 회차는 `coordinate.running()` 이 안다.
    ★ **원본을 한 글자도 안 지운다** — 형님이 "복사" 라 하셨다.
    """
    try:
        import data_mirror as DM
    except Exception as e:
        return "원본 이전 못 함(%s)" % type(e).__name__
    try:
        import coordinate as CO
        busy = CO.running()
    except Exception:
        busy = None
    if busy:
        # ★ 양보는 **주장이므로 자국을 남긴다**([293]) — 매일 양보만 하는 단계는
        #   없는 단계와 같고, 굶주림 판정이 그것을 잡는다.
        try:
            CO.record_yield(MIRROR_JOB, " · ".join(sorted(busy))[:60],
                            "Z: 를 같이 긁지 않는다")
        except Exception:
            pass
        return "원본 이전 양보(%s 도는 중)" % " · ".join(sorted(busy))[:40]
    try:
        res = DM.run(apply=not dry, budget_s=MIRROR_BUDGET_S)
        msg = DM.line(res) + ("(dry)" if dry else "")
        # ★ **돌았으면 돌았다고 적는다**([293]) — 예산 끝도 **완주**다(설계된
        #   이어하기다: 진도가 남고 다음 회차가 잇는다).  Z: 에 못 닿은 것만
        #   실패로 적는다 — 그때는 일이 안 됐다.
        # ★ **dry 는 안 적는다**([169]) — 미리보기는 그 일을 한 것이 아니다.
        if not dry:
            _mark_mirror_run("실패" if res.get("왜못함") else "완주",
                             str(res.get("왜못함") or msg)[:80])
        return msg
    except Exception as e:
        # 이전 하나로 회차를 안 죽인다 — 그러나 조용히 넘기지도 않는다([169]).
        if not dry:
            _mark_mirror_run("실패", "%s: %s" % (type(e).__name__, str(e)[:60]))
        return "원본 이전 실패(%s: %s)" % (type(e).__name__, str(e)[:60])


def publish_endpoint(dry):
    """터널 주소가 바뀌면 고정 주소(GitHub Pages)에 자동 게시 — 폰 북마크 불변"""
    if dry:
        return "게시(dry)"
    try:
        r = run_tree([PY, os.path.join(ROOT, "publish_endpoint.py")], cwd=ROOT,
                     timeout=120, drain_timeout=10)
        return (r.stdout or "").strip().splitlines()[-1][:60] if r.stdout.strip() else "게시 무응답"
    except Exception as e:
        return f"게시 오류: {str(e)[:40]}"


def sync_cloud_queue(dry):
    """PC 재가동 뒤 휴대폰 예약을 관리대장에 반영한다."""
    if dry:
        return "클라우드 예약 반영(dry)"
    try:
        r = run_tree(
            [PY, os.path.join(ROOT, "cloud_queue_sync.py")],
            cwd=ROOT,
            timeout=180,
            drain_timeout=10,
        )
        line = (r.stdout or r.stderr or "").strip().splitlines()
        return ("클라우드 예약 " + line[-1][:80]) if line else "클라우드 예약 확인"
    except Exception as e:
        return f"클라우드 예약 오류: {str(e)[:40]}"


def sync_uploads(dry):
    """단일 투입함을 30분마다 정본 분류하고, 새 원본이면 전체 대조를 한 번 깨운다."""
    if dry:
        return "업로드 투입함(dry)"
    try:
        r = run_tree(
            [PY, os.path.join(ROOT, "upload_intake.py"), "--apply"],
            cwd=ROOT, timeout=300, drain_timeout=15,
        )
        line = (r.stdout or r.stderr or "").strip().splitlines()
        summary = line[-1] if line else "업로드 분류 무응답"
        m = re.search(r"업로드 원본 분류:\s*(\d+)건", summary)
        moved = int(m.group(1)) if m else 0
        if r.returncode == 0 and moved:
            # daily_run의 프로세스 잠금이 이미 실행 중인 중복 기동을 안전하게 막는다.
            start_hidden("daily_run.py")
            return f"{summary} → 전체 대조 시작"
        return summary[:100]
    except Exception as e:
        return f"업로드 투입함 오류: {str(e)[:50]}"


def resume_deferred_apply(dry):
    """엑셀 열림으로 연기된 반영의 30분 안전망(2026-08-03 지시).

    닫힘 감시자(ledger_db --resume-watch)가 죽었어도 마커가 남아 있으면
    여기서 재개하거나 감시자를 다시 띄운다."""
    try:
        import ledger_db
        if not ledger_db._defer_state().get("slots"):
            return "연기 회차 없음"
        if dry:
            return "연기 회차 있음(dry — 재개 생략)"
        r = ledger_db.resume_check()
        return "연기 반영 점검 → " + str(r.get("상태") or r)
    except Exception as exc:
        return f"연기 반영 점검 실패({type(exc).__name__})"


def heal_autopilot(dry):
    """공용 자원 장애로 미뤄 둔 안전 작업을 30분마다 조금씩 이어 간다."""
    try:
        import autopilot
        result = autopilot.heal(limit=2, budget_seconds=600, dry=dry)
        actions = result.get("actions") or []
        # ★ 이 회차 예산 밖이라 **안 부른 것**을 조용히 넘기지 않는다([169]·[436]) —
        #   안 적으면 다 건너뛴 회차가 '대기 없음'으로 보인다.
        over = result.get("예산밖") or []
        tail = (" · 이 회차 예산 밖 %d건(예약 회차·사람 몫)" % len(over)) if over else ""
        if not actions:
            return ("자율복구 대기 없음" if not result.get("active") else (
                "자율복구 %d건 대기(재시도 시각 전·인증 대기)" % result.get("active", 0))) + tail
        done = sum(1 for x in actions if x.get("result") == "done")
        return "자율복구 %d건 실행 · 완료 %d · 남음 %d%s" % (
            len(actions), done, result.get("active", 0), tail)
    except Exception as exc:
        return "자율복구 점검 실패(%s)" % type(exc).__name__


def sync_worklog(dry):
    """일지 원본이 바뀌면 **기다리지 않고** 바로 대조해 큐에 넣는다 (2026-08-09 지시).

    사용자 지시: "돌발 AS 일지랑 정기점검 일지 엑셀과 앱에 자동 반영 ... 담당자 손댈 필요 없이".

    ★ 조각은 이미 다 있었다 — `upload_intake` 가 정본 자리로 옮기고, 09:50 회차가
      `work_log_sync --queue` 를 돌리고, 11:00·15:00 이 엑셀에 쓴다.
      빠진 것은 **속도**다: 오후에 올린 일지는 다음 날 09:50 까지 아무 데도 안 반영된다.
      그 사이 화면은 옛 숫자를 멀쩡히 보여 주므로 **올린 사람만 반영된 줄 안다.**
    ★ 파일이 안 바뀌었으면 아무것도 안 한다. 매 30분 대조를 돌리면 Z: 를 계속 훑어
      앱까지 같이 느려진다([168] · 사고 #29). 판정 근거는 **정본 폴더 최신 mtime** 하나다.
    """
    if dry:
        return "일지 감시(dry)"
    stamp = os.path.join(ROOT, "reports", ".worklog_seen.json")
    try:
        from source_dirs import work_log_dirs
        newest, newest_p = 0.0, ""
        for d in work_log_dirs():
            if not os.path.isdir(d):
                continue
            for f in glob.glob(os.path.join(d, "**", "*.xls*"), recursive=True):
                if os.path.basename(f).startswith("~$"):
                    continue
                m = os.path.getmtime(f)
                if m > newest:
                    newest, newest_p = m, f
        if not newest:
            return "일지 원본 없음"
        seen = 0.0
        if os.path.exists(stamp):
            try:
                seen = float(json.load(open(stamp, encoding="utf-8")).get("mtime") or 0)
            except Exception:
                seen = 0.0
        if newest <= seen:
            return "일지 그대로"
        r = run_tree([PY, os.path.join(ROOT, "work_log_sync.py"), "--queue"],
                     cwd=ROOT, timeout=900, drain_timeout=30, output_limit=120_000)
        # ★ 실패했으면 자국을 남기지 않는다 — 남기면 '봤다'가 되어 **다시는 안 본다.**
        #   그 일지는 영영 반영되지 않고 오류도 안 난다.
        if r.returncode != 0 or r.timed_out:
            why = "시간초과" if r.timed_out else "rc=%d" % r.returncode
            return "일지 대조 실패(%s) — 다음 회차에 다시 시도" % why
        with open(stamp, "w", encoding="utf-8", newline="") as f:
            json.dump({"mtime": newest, "파일": os.path.basename(newest_p),
                       "본때": datetime.now().isoformat(timespec="seconds")},
                      f, ensure_ascii=False)
        tail = [l for l in (r.stdout or "").strip().splitlines() if l.strip()]
        return "새 일지 반영 대기열 적재 · %s" % (tail[-1][:70] if tail else "요약 없음")
    except Exception as e:
        return "일지 감시 오류: %s" % str(e)[:50]


def run_incremental_pipeline(dry):
    """Five-minute task의 안전망 — **시작만 요청하고 기다리지 않는다.**

    별도 작업 스케줄러가 주 경로다. 워치독에도 맨 앞에 두는 이유는 스케줄러가
    비활성화되거나 PC가 예약 시각에 꺼져 있었어도 다음 30분 회차에서 복구하기
    위해서다. 파이프라인 자체 PID 잠금과 지문 멱등성이 중복 실행을 막는다.

    ★ 2026-08-24 실사고: 워치독이 이 자리에서 파이프라인 종료를 기다리다 **130분**
      멈췄다. 그동안 서버·터널·회차 감시도 함께 멈췄다. 감시자는 일을 시키되 그 일과
      같이 갇히면 안 된다. Task Scheduler의 ``IgnoreNew`` 에 맡기고 바로 다음 눈으로 간다.
    """
    if dry:
        return "증분 자동화 시작 예행(dry)"
    result = run_tree(
        ["schtasks.exe", "/Run", "/TN", "CSOS_AutomationPipeline"],
        cwd=ROOT, timeout=20, drain_timeout=5, output_limit=20_000,
    )
    if result.timed_out:
        return "증분 자동화 시작 요청 시간초과 — 다음 워치독이 재시도"
    if result.returncode != 0:
        tail = [line for line in (result.stdout or "").splitlines() if line.strip()][-1:]
        return "증분 자동화 시작 실패(rc=%d)%s" % (
            result.returncode,
            (" · " + tail[0][:80]) if tail else "",
        )
    return "증분 자동화 시작 요청 전달(완료는 5분 회차 자국이 확인)"


def resume_parked(dry):
    """세워 둔 일이 풀렸으면 알리고(아무도 없으면 AI 에게 넘기고), 보류된 푸시를 민다.

    ★ 2026-08-11 지시 "이 세션도 완전 자동화시켜". 그날 실측 둘 — 옆 세션이 `code` 를
      놓았는데 **아무 화면도 "이제 된다"고 말하지 않아** 사람이 "하던 작업 진행" 을 두 번
      쳤고, 자동 마무리가 보류한 푸시는 그 세션이 사라진 뒤에도 **미는 사람이 없었다.**
      판단은 `worksplit_auto` 한 곳이 한다 — 여기서 다시 세면 두 벌이 된다.
    """
    try:
        import worksplit_auto
        return worksplit_auto.run(dry=dry)["한줄"]
    except Exception as exc:
        return "세션자동화 실패: %s: %s" % (type(exc).__name__, str(exc)[:80])


def watch_schedules(dry):
    """회차가 **정말 돌았나** — 스케줄러의 마지막 결과를 읽는다 (2026-08-12, `[228]`).

    ★ 이 프로젝트에서 자동화의 마지막 구멍이 여기였다. 지시문 '자동으로 도는 것'
      목록에 한 줄을 적으면 자동이 된 것처럼 보이지만, 실제로 도는지를 아는 것은 그
      목록이 아니라 **작업 스케줄러**다. 실측 2026-08-12 — 일일대조·원본정리가 매일
      제한시간에 걸려 강제 종료되고 정오회차는 등록조차 안 돼 한 번도 안 돌았는데
      **아무 화면에도 안 떴다**(`LastTaskResult` 를 읽는 코드가 한 줄도 없었다).
    ★ **`snapshot_handoff` 보다 먼저**다. 여기서 판정을 새로 써야 같은 회차의 인계
      문서가 그것을 읽는다 — 뒤에 두면 인계가 언제나 30분 전 판정을 싣는다.
    ★ **`dry` 여도 묻는다.** 읽기 전용이라 고치는 것이 없고, 안 물으면 `--dry` 로는
      이 눈이 도는지 확인할 길이 자체가 없다.
    """
    try:
        import schedule_watch
        st = schedule_watch.build()
    except Exception as exc:                       # 눈 하나 때문에 회차를 세우지 않는다
        return "스케줄러 감시 실패: %s" % str(exc)[:60]
    if st.get("조회실패"):
        # ★ '이상 없음'이 아니라 '확인 못 함'이다 — 감시자가 눈먼 채 정상을 말하면 안 된다.
        return "스케줄러 확인 못 함: %s" % st["조회실패"][:60]
    al = st.get("경보") or []
    if not al:
        # ★ '경보 0' 을 '정상'이라 적지 않는다 — 꺼진 회차는 알림이라 여기 안 들어온다.
        no = len(st.get("알림") or [])
        return "스케줄러 경보 없음(%d회차%s)" % (
            len(st.get("작업") or []), " · 알림 %d건" % no if no else "")
    return "스케줄러 경보 %d건: %s" % (
        len(al), ", ".join("%s %s" % (a["갈래"], a["작업"]) for a in al[:4]))


def recover_missed_schedules(dry):
    """작업 스케줄러가 놓친 일일 회차를 한 건씩 되살린다.

    판정은 바로 앞 ``watch_schedules``가 만든 파일을 읽고, 쓰는 손은 별도 모듈에 둔다.
    감시자의 읽기 전용 계약을 깨지 않으며 한 번에 하나만 시작한다.
    """
    try:
        import schedule_recover
        return schedule_recover.run(dry=dry)
    except Exception as exc:
        return "놓친 회차 자동복구 실패: %s: %s" % (type(exc).__name__, str(exc)[:70])


def watch_coordination(dry):
    """겹친 일이 **정말 되었나** (2026-08-17 지시, `[293]`).

    양보는 "저쪽이 그 일을 한다"는 주장이다. 주인마저 못 끝내면 그 일은 아무도 안 한
    것인데 — 양보는 실패가 아니라서 스케줄러도 회차 감시도 아무 말을 안 한다.
    ★ 읽기 전용이다. 회차를 다시 띄우지 않는다 — 지금 도는 중일 수 있고, 그때 띄우면
      같은 자료를 두 번 긁는다. 여기는 보고 말하는 자리다(`typo_watch` 와 같은 자리).
    """
    try:
        import coordinate
        d, why = coordinate._load()
        if not dry:
            coordinate.write_report()
        ns = coordinate.notices(d, why)
    except Exception as exc:
        return "조율 확인 실패: %s" % str(exc)[:60]
    if why:
        return "조율 확인 못 함: %s" % why[:60]
    if not ns:
        return "겹침 경보 없음(작업 %d개)" % len(d)
    return "겹침 알릴 것 %d건: %s" % (len(ns), ", ".join(sorted({n["갈래"] for n in ns})))


def watch_orgchart(dry):
    """조직도가 바뀌었고 **그것이 따라갔나** (2026-08-13 지시, `[297]`).

    ★ 반영하는 손은 이미 있다 — 조직도 **코드**가 바뀌면 이 회차의 `heal_stale_server`
      가 서버를 갈아 준다(`[156]`·`[265]`). 여기서 또 갈면 폰이 502 를 두 번 받는다.
    ★ 아무도 안 보던 것은 **코드가 아닌 변경**이다: 흐름 정의는 DB 표(`flow_step`)라
      사람이 앱에서 고치면 **파일 mtime 이 안 움직인다** — `heal_stale_server` 는
      "서버 정상"이라 말하고 스케줄러도 조용하다.
    ★ `heal_stale_server` **뒤**, `snapshot_handoff` **앞**이다(`watch_schedules` 와
      같은 이유 — 뒤에 두면 인계가 언제나 30분 전 판정을 싣는다).
    ★ **`dry` 여도 본다.** 읽기 전용이라 고치는 것이 없고, 안 보면 `--dry` 로 이 눈이
      도는지 확인할 길이 없다. 다만 자국은 안 남긴다.
    """
    try:
        import org_watch
        d = org_watch.build(save=not dry)
        return org_watch._line(d)
    except Exception as exc:
        return "조직도 확인 실패: %s: %s" % (type(exc).__name__, str(exc)[:60])


def _heal_pm_calendar(dry, kind):
    """달력 정기점검 예정이 **밀렸으면** 그 자리에서 다시 만든다 (2026-08-26).

    * 왜 여기까지 하나 — 같은 원본이 **둘**을 먹이는데(`[351]`) 고치는 값이 하늘과 땅이다:
      담당자 자료(`camp_contacts --write`)는 밴드 전체를 다시 파싱해 수십 초라 사람이
      명령한다. 달력(`pm_schedule_sync.py`)은 **실측 3.3초**다. 그런데 그것을 부르는
      회차가 **09:50 하루 한 번뿐**이다 — 실측 2026-08-26: 5분 증분 파이프라인 단계에
      그 이름이 없고, 정기점검 스케줄 원본은 신호 갈래(kakao/band/erp) **어디에도 안 든다**.
      그래서 류지영이 11:58 에 저장하면 다음 날 09:50 까지 **22시간** 동안 대표 보고
      달력이 **옛 예정**을 오류 없이 그대로 보여 준다(`[169]`·`[376]` 과 같은 모양).
    * **`--apply` 를 쓰지 않는다.** 그것은 관리대장 vN+1 을 만든다. 여기서 만드는 것은
      `reports/pm_schedule_sync.json` 하나이고 화면·대표 캡처는 그것을 읽는다.
      원장을 언제 쓸지는 **사람이 정한다**(주 1회 보관 · 2026-08-24 지시).
    * **밀렸을 때만** 한다(`[169]`) — 모름은 못 잰 것이지 밀린 것이 아니고, 정상까지
      돌리면 30분마다 Z: 를 헛되이 문다(`[168]`·`[170]`).
    * **담당자 자료는 안 건드린다**(`[172]`) — 값이 다르므로 조치도 다르다(`[289]`).
    * `subprocess.run(timeout=)` 을 쓰지 않는다(`[175]`) — Z: 대기에 걸리면 윈도우에서 안 끝난다.
    * 실패를 성공처럼 적지 않는다(`[169]`).
    """
    if kind != "밀림":
        return ""
    if dry:
        return "(dry — 안 고침)"
    try:
        r = run_tree([PY, os.path.join(ROOT, "pm_schedule_sync.py")],
                     cwd=ROOT, timeout=300, drain_timeout=15)
    except Exception as exc:
        return " -> 다시 만들기 실패: %s" % type(exc).__name__
    if getattr(r, "timed_out", False):
        return " -> 다시 만들기 시간초과(300초)"
    if r.returncode != 0:
        return " -> 다시 만들기 실패(코드 %s)" % r.returncode
    # * 돌았다고 고쳐졌다 하지 않는다(`[322]`) — **다시 재서** 말한다.
    try:
        import camp_contacts
        again = (camp_contacts.sched_stale().get("달력") or {}).get("갈래")
    except Exception:
        return " -> 다시 만들었다(그 뒤 상태는 못 쟀다)"
    return " -> 다시 만들어 정상" if again == "정상" else " -> 다시 만들었는데 %s" % again


def watch_camp_source(dry):
    """정기점검 스케줄 원본이 새로 왔는데 **앱 담당자 자료가 안 따라갔나**
    (2026-08-19 실사고, 분담판 `[151]` · `[328]`).

    ★ 이 사고는 **빈칸이 아니라 '틀린 사람'** 으로 나타난다([165]) — 원본이 갱신된
      뒤 09:50 회차 전까지 화면은 **옛 담당자 이름.전화**를 멀쩡히 보여 준다.
      실측 2026-08-19: 원본 16:44 대 자료 16:16 = 152캠프 754칸 중 **236칸**이
      옛 값이었고, 형님이 캡처를 들고 묻고서야 드러났다.
    ★ **여기서 다시 만들지 않는다**([168]) — `camp_contacts.build()` 는 밴드 전체를
      파싱한다(수십 초). 재는 것은 mtime 둘뿐이고 고치는 것은 사람이 명령한다.
    ★ **`snapshot_handoff` 앞**이다(`watch_schedules` 와 같은 이유 — 뒤에 두면
      인계가 언제나 30분 전 판정을 싣는다).
    ★ **`dry` 여도 본다** — 읽기 전용이라 고치는 것이 없다. 다만 자국은 안 남긴다.
    """
    try:
        import camp_contacts
        rec = camp_contacts.sched_stale() if dry else camp_contacts.stale_mark()
        갈래 = rec.get("갈래") or "모름"
        # ★ 같은 원본이 **달력 예정**도 먹인다(분담판 `[168]`) — 담당자 자료만 보고
        #   "최신"이라 적으면 달력이 옛 예정을 보여 주는 동안 회차가 조용하다(`[169]`).
        cal = (rec.get("달력") or {}).get("갈래")
        tail = ("" if cal in (None, "정상")
                else " · 달력 예정 %s%s" % (cal, _heal_pm_calendar(dry, cal)))
        if 갈래 == "정상":
            return "캠프 원본 최신" + tail
        if 갈래 == "밀림":
            return "캠프 원본 밀림 %d분" % (rec.get("늦은분") or 0) + tail
        return "캠프 원본 %s" % 갈래 + tail
    except Exception as exc:
        # ★ 못 본 것을 '정상'이라 하지 않는다([169]).
        return "캠프 원본 확인 실패: %s: %s" % (type(exc).__name__, str(exc)[:60])


def watch_takeover(dry):
    """다른 계정이 **언제든 이어받을 수 있는 상태인가** (2026-08-17 지시).

    ★ 크레딧 소진에는 훅이 없다. compact·`/clear`·종료는 셋 다 훅이 받아 인계를
      남기지만, 크레딧이 떨어진 창은 그대로 뜬 채 **대화기록만 멈춘다** —
      pid 가 살아 있으니 어떤 계기도 "이 창은 멈췄다"고 말하지 않는다.
    ★ **`snapshot_handoff` 보다 먼저**다(`watch_schedules` 와 같은 이유).
    ★ 읽기 전용이다 — 점유를 뺏지 않는다. 회수는 기존 규칙 몫이다.
    """
    try:
        import takeover
        rows, why = takeover.sessions()
        repo = takeover.repo_state()
        takeover.write(takeover.card(rows, why, repo))
    except Exception as exc:
        return "이어받기 카드 실패: %s" % str(exc)[:60]
    if rows is None:
        return "이어받기 확인 못 함: %s" % why[:60]
    n = len(takeover.notices(rows, why, repo))
    끊김 = sum(1 for r in rows if r["갈래"] == "끊긴듯")
    return "이어받기 준비됨(끊긴 창 %d · 알릴 것 %d)" % (끊김, n)


def watch_credit(dry):
    """크레딧 5시간 창을 **파일에 적어 둔다** (2026-08-22 형님 지시).

    ★ 이 한 줄이 형님 지시의 "다른 계정이나 다른 세션에서도 인식"이다. 대화기록은
      창마다 따로라 옆 창은 남의 소진을 못 본다. 그런데 이 파일은 **저장소가 공유**
      하므로, 무인 회차가 30분마다 적어 두면 **어느 계정으로 어느 창을 열든** 같은
      사실을 본다.
    ★ 무인 회차는 크레딧과 무관하다 — 파이썬이지 Claude 가 아니다. 그래서 소진 중에도
      이 단계는 그대로 돌고, 충전 시각이 지나면 다음 회차가 그것을 적는다.
    ★ **`snapshot_handoff` 보다 먼저**다(`watch_schedules` 와 같은 이유 · `[228]`).
    ★ 읽기 전용이다 — 아무것도 안 고친다. 실측 0.2초([168]).
    """
    try:
        import credit_window
        st = credit_window.note()
    except Exception as exc:
        return "크레딧 창 확인 실패: %s" % str(exc)[:60]
    갈 = st.get("갈래")
    if 갈 == "소진":
        return "크레딧 소진 — %d분 뒤 충전(AI 인계는 그때까지 안 만든다)" % (st.get("남은분") or 0)
    if 갈 == "제한":
        return credit_window.line(st)
    if 갈 == "모름":
        return "크레딧 창 확인 못 함: %s" % str(st.get("왜") or "")[:60]
    return "크레딧 창 여유 있음"


def watch_userscript(dry):
    """크롬 전용 수집이 **정말 돌고 있나** — 유저스크립트의 되보고를 읽는다 (2026-08-13).

    ★ `watch_schedules` 가 스케줄러 회차에 대해 하는 일을 여기서는 **브라우저**에
      대해 한다. 실측 2026-08-13: 유저스크립트는 2026-08-09 에 만들어져 검증 `[182]`
      까지 붙어 있었는데 **나흘 동안 한 번도 안 돌았다**(Tampermonkey 미설치).
      그런데 그 사실을 말해 주는 화면이 어디에도 없었다 — 자동 수집이 '있다'고 적힌
      채 아무 일도 안 일어났고 아무도 몰랐다.
    ★ **`snapshot_handoff` 보다 먼저**다(`watch_schedules` 와 같은 이유).
    ★ 읽기 전용이라 `dry` 여도 묻는다 — 안 물으면 `--dry` 로 이 눈을 확인할 길이 없다.
    """
    try:
        from band import userscript_watch
        # `--dry` 여도 정본 md 는 쓴다 — 읽기 전용 판정이라 잃는 것이 없고,
        # 안 쓰면 같은 회차의 인계가 읽을 것이 없다(`watch_schedules` 와 같다).
        st = userscript_watch.check(write=True)
    except Exception as exc:                       # 눈 하나 때문에 회차를 세우지 않는다
        return "크롬수집 감시 실패: %s" % str(exc)[:60]
    kind = str(st.get("갈래") or "")
    # ★ '못 읽음'을 '정상'이라 하지 않는다([169]) — 갈래 이름을 그대로 싣는다.
    if kind == "정상":
        return "크롬수집 정상(%d밴드)" % len(st.get("밴드") or {})
    return "크롬수집 %s: %s" % (kind or "?", (st.get("왜") or "")[:60])


def heal_band_bridge(dry):
    """사람 탭이 **없을 때** 전용 크롬이 대신 긁는다 (2026-08-27 지시).

    형님 지시: "매번 불편하게 계속 안잡힌다 뭐해라뭐해라 하지말고 **알아서 좀 해봐**".
    실측 2026-08-27: 매출처업무 밴드가 **6일째** 멈춰 있었고, 그동안 어느 회차도
    대신해 주지 않았다 — 다리를 부르는 코드가 **한 줄도 없었다**([328]).

    ★ **창을 내려 둔 채로 돈다**(2026-08-27 실측).  최소화된 창에서
      `Emulation.setFocusEmulationEnabled` 뒤 `document.hidden=false` ·
      3초 타이머가 **3009ms**(안 조여진다).  그러므로 형님 화면을 한 번도 안 뺏는다.
      ⚠ 속이는 것이 아니다 — 크롬 자신의 손잡이다([457]).

    ★ **사람 탭이 살아 있는 밴드는 안 건드린다.**  두 창이 같은 밴드를 긁으면
      캐시가 오염된다(사고 #27 — 되돌릴 수 없는 쪽).  그래서 가져가는 것은
      `끊김`(6시간 넘게 조용) · `안옴`(한 번도 안 옴) **둘뿐**이다.
      ⚠ `가려짐` 은 **일부러 뺐다** — 그 창의 스크립트는 살아 있어 형님이 탭을
        바꾸는 순간 시작한다.  겹치는 값이 되돌릴 수 없으므로 안 가져간다([172]).
        (한 창에서는 밴드 하나만 긁힌다 — 2026-08-25 실측.  그 자리를 메우려면
        양쪽이 함께 보는 잠금이 있어야 하고, 그것은 재고 나서 한다([67]).)

    ★ **판정을 새로 만들지 않는다**([162]) — `userscript_watch.check()` 가 이미
      밴드마다 갈래를 매긴다.  여기서 다시 재면 두 화면이 다른 답을 한다.

    ★ **한 회차에 한 밴드만** · 기다림은 짧게(`BRIDGE_STEP_WAIT_S`).  수집기는
      10건마다 저장하므로([388]) 짧은 창에서도 진도가 남고 다음 회차가 이어받는다.
      길게 잡으면 워치독 예산을 통째로 먹는다([436]).
    """
    # ★ 형님이 밴드 수집을 멈추라 하셨으면 **긁지 않는다** (2026-09-01 지시).
    #   판정은 `band.collect_switch` 한 곳이다([162]) — 여기서 다시 재면
    #   자동 경로 셋이 서로 다른 답을 한다.  못 읽으면 예전 그대로 긁는다([169]).
    try:
        from band import collect_switch as _CS
        _off, _why = _CS.stopped()
    except Exception:
        _off, _why = False, ""
    # ★ **--dry 를 먼저 본다** — 그것은 미리보기이고, 미리보기의 뜻은
    #   "지금 돌리면 무엇을 하나" 다.  중단 문구만 내면 *"--dry 인데도 긁었나"* 를
    #   알 수 없다.  그래서 **왜 안 긁는지 둘 다** 말한다([169]).
    #   ⚠ 순서를 뒤집으면 검증 [460] 이 잡는다(2026-09-01 실측 — 관문이 빨갰다).
    if dry:
        return ("밴드 다리: --dry 라 안 긁음"
                + ((" (지금은 %s)" % _why) if _off else ""))
    if _off:
        return "밴드 다리: %s" % (_why or "수집 중단")
    try:
        from band import userscript_watch
        st = userscript_watch.check(write=False)
    except Exception as exc:
        return "밴드 다리: 갈래를 못 읽음 — %s" % str(exc)[:60]

    TAKE = ("끊김", "안옴")
    cand = [b for b, v in (st.get("밴드") or {}).items()
            if str(v.get("갈래") or "") in TAKE]
    # ⚠ 대기열을 여기서 다시 보지 않는다([162]).  처음엔 "대기열에 있는데 보고가
    #   없으면 아무도 안 긁는 것" 이라는 예비 경로를 뒀는데, 그것이 곧 **판정을
    #   두 곳에서** 하는 것이라 **살아 있는 탭까지 가져갔다**(검증 [460] 이 잡았다).
    #   보고에 없는 밴드는 `userscript_watch` 가 `안옴` 으로 말해야 한다 — 안 말하면
    #   고칠 자리는 여기가 아니라 거기다.
    if not cand:
        return ""                 # 사람 탭이 다 살아 있다 — 조용하다([170])

    band = sorted(cand)[0]        # 한 회차에 하나만
    try:
        from band import browser_bridge as BB
    except Exception as exc:
        return "밴드 다리: 못 들여옴 — %s" % str(exc)[:60]
    if not BB.alive():
        # ★ 형님 지시(2026-08-27): **"새로 띄우지 말고 기존 크롬창 지금 떠있는
        #   크롬창으로해"**.  그래서 여기서 크롬을 **안 띄운다** — 창이 하나 더
        #   뜨는 것이 그 지시가 막는 것이다.  붙을 크롬이 없으면 **그 사실만 적고
        #   물러난다**([169] — 못 한 것을 한 것처럼 적지 않는다).
        # ⚠ 형님 크롬을 그대로 쓰는 길은 **막혔다**(2026-08-27 실측) — 크롬 136판부터
        #   **기본 프로필에서는 디버깅 문을 안 연다**.  그러니 여기서 그 길을
        #   가리키면 사람이 **되지도 않는 것을 하러 간다**([172]·[448]).
        return ("밴드 다리(%s): 디버깅 문이 열린 크롬이 없다 — "
                "`python band/browser_bridge.py --up` 으로 전용 크롬을 띄우고 "
                "그 창에서 **한 번만** 로그인하면 그 뒤로는 자동입니다" % band)
    wait = int(os.environ.get("BRIDGE_STEP_WAIT_S", "420"))
    try:
        out = BB.collect_band(band, wait_s=wait)
    except Exception as exc:
        return "밴드 다리(%s): 실패 — %s" % (band, str(exc)[:60])
    try:
        BB.note(out)
    except Exception:
        pass
    결과 = out.get("결과")
    if 결과 == "사람대기":
        # ★ **로그인은 대신 안 한다**(절대규칙 3) — 그 사실만 말한다([169]).
        return ("밴드 다리(%s): 전용 크롬 로그인 한 번 필요 — "
                "python band/browser_bridge.py --up" % band)
    if 결과 != "완주":
        # ★ **봉투를 버리지 않는다**([365]·[289]).  '실패' 다섯 글자만 남기면 겉은
        #   경보인데 **왜인지는 영영 알 수 없다** — 조치가 갈래마다 다르다.
        return "밴드 다리(%s): %s — %s" % (band, 결과, (out.get("왜") or "이유 없음")[:80])
    상태 = out.get("상태") or {}
    return "밴드 다리(%s): %s · 수확 %s · 실패 %s" % (
        band, 결과, 상태.get("ok"), 상태.get("failed"))


def _bridge_minimize(BB):
    """다리 창을 내려 둔다 — 못 내려도 수집은 돈다(focusEmulation 이 본체다)."""
    try:
        t = (BB.tabs() or [None])[0]
        if not t:
            return
        with BB._conn(t) as ws:
            w = BB._call(ws, "Browser.getWindowForTarget")
            BB._call(ws, "Browser.setWindowBounds",
                     {"windowId": w.get("windowId"),
                      "bounds": {"windowState": "minimized"}})
    except Exception:
        pass


def close_upload_notices(dry):
    """올린 것의 **결과**를 뒤따라 알린다 (2026-08-14 지시).

    ★ 접수 알림은 앱이 그 자리에서 보내고, 앱이 시작한 회차의 결과도 그 자리에서
      확인한다. 그런데 5분 스케줄러(`CSOS_AutomationPipeline`)가 먼저 집어가면
      **앱은 그 회차가 끝난 것을 모른다** — 그러면 '올렸다'만 가고 결과는 영영
      안 온다([169] 의 모양). 그 자리를 이 단계가 막는다.
    ★ 판정을 새로 만들지 않는다([162]) — `notify.sweep_uploads` 가 회차 자국을
      읽어서 하고, 여기는 부르기만 한다.
    ★ `dry` 면 알림을 만들지 않는다 — 알림은 사람에게 나가는 것이라 예행에서 보내면
      되돌릴 수 없다.
    """
    if dry:
        return "업로드 결과 알림 예행(보내지 않음)"
    try:
        import notify
        r = notify.sweep_uploads()
    except Exception as exc:                       # 눈 하나 때문에 회차를 세우지 않는다
        return "업로드 결과 알림 실패: %s" % str(exc)[:60]
    if not (r.get("끝") or r.get("실패") or r.get("확인못함")):
        return "업로드 결과 알림 없음(대기 %d건)" % int(r.get("남음") or 0)
    return "업로드 결과 알림 끝%d 실패%d 확인못함%d" % (
        int(r.get("끝") or 0), int(r.get("실패") or 0), int(r.get("확인못함") or 0))


#: 이 회차가 어디까지 갔나 — **죽어도 남는다. 그게 요점이다**([180]).
PROGRESS = os.path.join(ROOT, "reports", ".워치독_진행.json")
#: 단계가 터졌을 때의 자취 — `schedule_watch.traces()` 가 `*_오류.json` 을
#  글로브로 모으므로 목록을 손으로 안 적어도 인계에 실린다([304]·[324]).
TRACE = os.path.join(ROOT, "reports", "워치독_오류.json")
#: 회차 주기(30분)를 넘겼을 때만 말한다 — 정상까지 적으면 아무도 안 읽는다([170]).
ROUND_WARN_MIN = 30
_STEP_TIMES = []


def _note_progress(단계, 상태, 시작):
    """지금 어느 단계인지·단계마다 몇 초 걸렸는지를 그때그때 디스크에 적는다.

    ★ 자국 하나로 회차를 세우지 않는다 — 못 적어도 그냥 넘어간다.
    """
    try:
        느린 = sorted(_STEP_TIMES, key=lambda x: -x[1])[:5]
        d = {"단계": 단계, "상태": 상태, "pid": os.getpid(),
             "회차시작": datetime.fromtimestamp(시작).isoformat(timespec="seconds"),
             "경과분": round((time.time() - 시작) / 60.0, 1),
             "끝낸단계": len(_STEP_TIMES),
             "단계기록": [{"단계": n, "초": round(sec, 1)} for n, sec in _STEP_TIMES],
             "느린단계": [{"단계": n, "분": round(sec / 60.0, 1)}
                          for n, sec in 느린 if sec >= 60]}
        os.makedirs(os.path.dirname(PROGRESS), exist_ok=True)
        tmp = PROGRESS + ".tmp"
        open(tmp, "w", encoding="utf-8").write(json.dumps(d, ensure_ascii=False))
        os.replace(tmp, PROGRESS)
    except Exception:
        pass


def _leave_trace(단계, 왜):
    """터진 단계를 자취로 남긴다 — 완주하면 지운다([228])."""
    try:
        os.makedirs(os.path.dirname(TRACE), exist_ok=True)
        open(TRACE, "w", encoding="utf-8").write(json.dumps(
            {"때": datetime.now().isoformat(timespec="seconds"),
             "작업": "쿠팡업무_워치독", "단계": 단계, "무엇": 왜,
             "어떻게": "그 단계 코드를 본다 — 회차는 나머지를 끝까지 돌았다([175])."},
            ensure_ascii=False))
    except Exception:
        pass


def _clear_trace():
    """다음 회차가 끝까지 갔으면 옛 자취를 내린다 — 이미 고쳐진 고장을
    계속 보고하지 않는다([228])."""
    try:
        if os.path.exists(TRACE):
            os.remove(TRACE)
    except Exception:
        pass


def _run_step(fn, dry, 시작):
    """단계 하나를 돌리고 **걸린 시간을 남긴다.**

    ★ 왜 필요한가(2026-08-28 실측): 워치독 회차 555개 중 중앙값은 7.0분인데
      **19개(3%)가 30분(회차 주기)을 넘겼고** 최대 **354분**이었다.  그 사이
      서버·회차·인계 감시가 통째로 멈추는데 스케줄러는 '이미 실행 중'으로
      건너뛰고 **성공이라 적는다**([175]).  그런데 **어느 단계가 먹었는지 아는
      길이 한 줄도 없었다** — 결과를 끝에 한 번만 찍기 때문이다.  게다가
      **7회는 자국조차 통째로 사라졌다.**  `daily_run` 이 `[180]`·`[228]` 에서
      배운 것이 여기 안 와 있었다([300]).
    ★ 한 단계가 터져도 **회차를 세우지 않는다**([175]) — 나머지 감시가 통째로
      멈추는 값이 더 크다.  그러나 **삼키지도 않는다**([355]): 갈래와 터진 자리를
      적어 자취로 남기고, 그 자취가 인계 '먼저 처리할 것'에 오른다.
    ★ 이름은 함수에서 온다([162]·[340]) — 손으로 적으면 사본이 되어 뒤처진다.
    """
    이름 = getattr(fn, "__name__", "?")
    t0 = time.time()
    _note_progress(이름, "시작", 시작)
    try:
        out = fn(dry)
    except Exception as e:                                       # noqa: BLE001
        import traceback
        tb = traceback.extract_tb(sys.exc_info()[2])
        자리 = "%s:%d" % (os.path.basename(tb[-1].filename), tb[-1].lineno) if tb else "?"
        out = "%s 터짐 — %s(%s) @ %s" % (
            이름, type(e).__name__, (str(e) or "(사유 없음)")[:120], 자리)
        _leave_trace(이름, out)
    _STEP_TIMES.append((이름, time.time() - t0))
    _note_progress(이름, "끝", 시작)
    return out


def slow_note(step_times, 분, warn_min=None):
    """회차가 주기를 넘겼으면 **어느 단계가 먹었는지** 한 줄로 말한다.

    ★ 넘기지 않았으면 아무 말도 안 한다([170]) · 1분 미만 단계는 안 적는다
      (전부 적으면 아무도 안 읽는다).
    """
    if warn_min is None:
        warn_min = ROUND_WARN_MIN
    if 분 <= warn_min:
        return ""
    느린 = [(n, sec) for n, sec in sorted(step_times, key=lambda x: -x[1])[:3] if sec >= 60]
    꼬리 = (" — 오래 걸린 단계: "
            + ", ".join("%s %.0f분" % (n, sec / 60.0) for n, sec in 느린)) if 느린 else ""
    return "★ 이 회차가 %.0f분 걸렸다(주기 %d분)%s" % (분, warn_min, 꼬리)


def main():
    # 류지영 매니저 입력 중에는 로그 파일조차 갱신하지 않고 즉시 종료한다.
    if is_input_window():
        if sys.stdout is not None:
            print(f"입력 보호시간({input_window_label()}) — 워치독 무동작 종료")
        return
    dry = "--dry" in sys.argv
    # 원장 버전 정리는 daily_run의 ledger_versions.py 한 곳에서만 수행한다.
    # 워치독이 낮은 버전 포크를 OLD로 옮겨 증거를 숨기는 일을 막는다.
    gap = gap_note(last_log_line(), datetime.now())     # 기록은 healing 전에 읽는다
    # 긴 증분 회차가 도는 동안에도 감시자가 멈춘 것으로 오판하지 않게 시작 심박을 남긴다.
    # pythonw 무인 실행에서도 파일에만 기록하므로 콘솔 창은 생기지 않는다.
    log("워치독 회차 시작" + ("(dry)" if dry else ""))
    # ★ 순서가 뜻을 갖는다 — 아래 주석이 그 이유다.  이름은 함수에서 오므로
    #   손으로 적은 사본이 없다([162]·[340]).
    steps = [run_incremental_pipeline, sync_uploads, sync_worklog,
             sync_cloud_queue, heal_server_guard, heal_server,
             watch_sync_contract,
             heal_fixed_funnel,
             # ★ 근거 정정이 **붙여넣기 파일 만들기보다 먼저**다 — 목록에 담을 번호를
             #   정하는 것이 그 근거다(2026-08-11, `[223]`).
             heal_band_evidence,
             heal_stale_pastefiles,
             heal_autopilot,
             resume_parked,
             heal_tunnel, publish_endpoint, clean_reports,
             # ★ 쓸데없는 파일 정리 — `clean_reports`(30일 넘은 리포트) 뒤다.
             #   그쪽은 **날짜 도장이 없는 옛 파일**을 지우고 이쪽은 **도장이
             #   있는 것**을 묶음마다 남긴다 — 기준이 달라 둘 다 필요하다([172]).
             sweep_files,
             # ★ 엑셀은 저장용이다 — 어느 도구가 만든 새 보관본이든
             #   여기서 잠근다(2026-08-28 지시 · `[477]`).
             lock_ledger_archive,
             # ★ 원본을 새 정본 자리로 복사 — 예산 안에서 조금씩,
             #   도는 회차에는 양보한다(2026-08-27 지시 · `[464]`).
             mirror_originals,
             # ★ 스케줄러 판정이 **인계 스냅샷보다 먼저**다 — 뒤에 두면 인계 문서가
             #   언제나 30분 전 판정을 싣는다(2026-08-12, `[228]`).
             watch_schedules,
             # ★ StartWhenAvailable=True 인데도 08-24 일일 세 회차가 빠졌다.
             #   판정을 만든 **뒤** 한 건만 시작하고 다음 감시를 계속한다.
             recover_missed_schedules,
             # ★ 브라우저 쪽 눈도 인계보다 먼저다 — 같은 이유(2026-08-13, `[247]`).
             watch_userscript,
             # ★ 사람 탭이 없는 밴드는 전용 크롬이 대신 긁는다 (2026-08-27).
             #   `watch_userscript` **뒤**다 — 그 갈래를 근거로 쓴다([162]).
             heal_band_bridge,
             # ★ 이어받기 준비도 인계보다 먼저다 — 크레딧이 떨어진 창은 훅이 없어
             #   스스로 인계를 못 남긴다(2026-08-17, `[291]`).
             # ★ 크레딧 창은 **`watch_takeover` 앞**이다 — 이어받기 카드가 '끊긴듯'
             #   창의 원인을 이 자국에서 읽는다. 뒤에 두면 카드가 언제나 30분 전
             #   자국을 싣는다(2026-08-22 지시 · `[228]` 과 같은 이유).
             watch_credit,
             watch_takeover,
             # ★ 겹쳐서 양보한 일이 정말 되었나 — 인계보다 먼저다(2026-08-17, `[293]`).
             watch_coordination,
             # ★ 조직도 변경이 따라갔나 — `heal_stale_server` 뒤(그가 서버를 갈고
             #   나서 봐야 한다) · 인계 앞이다(2026-08-13 지시, `[297]`).
             watch_orgchart,
             # ★ 정기점검 스케줄 원본이 새로 왔는데 앱 담당자가 안 따라갔나 —
             #   인계 앞이다(2026-08-19, `[328]`).
             watch_camp_source,
             # ★ 올린 것의 결과를 뒤따라 알린다 — 5분 스케줄러가 집어간 회차는
             #   앱이 끝난 줄 모른다(2026-08-14).
             close_upload_notices,
             snapshot_handoff, resume_deferred_apply]
    시작 = time.time()
    del _STEP_TIMES[:]
    try:
        results = [_run_step(fn, dry, 시작) for fn in steps]
    finally:
        # ★ 죽어도 여기까지는 남는다 — '아직 도는 중'과 '죽었다'를 가른다([180]).
        _note_progress("(회차 끝)", "종료", 시작)
    말 = slow_note(_STEP_TIMES, (time.time() - 시작) / 60.0)
    if 말:
        results.append(말)
    if gap:
        results.insert(0, gap)
    log(" | ".join(results))
    _clear_trace()   # 끝까지 갔으면 옛 자취를 내린다([228])


if __name__ == "__main__":
    main()
