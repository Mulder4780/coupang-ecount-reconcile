# -*- coding: utf-8 -*-
"""
ledger_db.py — 모든 반영을 **DB에 모았다가 하루 두 번만 엑셀에 쓴다**
===============================================================================
사용자 지시(2026-07-30):
  "앱이나 클로드코드 명령, 코덱스 명령으로 반영은 모두 DB로 저장했다가 엑셀에 한 번에 반영"
  "엑셀 반영 시점은 오전 11시, 오후 3시 하루에 딱 두 번"

## 왜 바꾸나 (지금 방식의 문제)
지금은 도구가 무언가 채울 때마다 `ledger_writer --apply` 가 곧바로 vN+1 을 만든다.
그래서 하루에도 관리대장 버전이 수십 개씩 늘고(오늘 하루 v311→v327), 사람이 파일을 열어
작업하는 도중에도 새 버전이 생겨 **어느 것이 정본인지 흔들린다.**
모아 두었다가 정해진 시각에 한 번만 쓰면 버전은 하루 두 개, 정본은 언제나 분명하다.

## 왜 SQLite 인가
· **표준 라이브러리**다(`sqlite3`) — 이 프로젝트의 "새 의존성 금지" 원칙을 지킨다.
· 앱·Claude·Codex 세 곳이 동시에 넣어도 트랜잭션으로 안전하다.
  지금의 JSON 큐는 두 프로세스가 동시에 쓰면 한쪽이 통째로 사라진다(실제 위험).
· 중간에 죽어도 남는다. "무엇을 언제 누가 왜 넣었는지" 를 질의할 수 있다.

## 반영 시각 — 하루 두 번
  11:00 · 15:00 (한국시간). 각 시각 뒤 `GRACE_MIN` 분 안에 실행되면 그 회차로 친다.
  ★ 놓친 회차를 그냥 버리지 않는다 — PC가 꺼져 있었으면 다음 실행 때 밀린 회차를 처리한다.
  ★ 입력 보호시간(08:00~09:30)과 겹치지 않는다. 사람이 입력하는 동안 원장을 건드리지 않는다.

## 흐름
    앱/Claude/Codex ──enqueue()──▶ SQLite(pending)
                                      │  11:00 · 15:00 에만
                                      ▼
                              ledger_writer ──▶ 관리대장 vN+1
    그 사이 어느 시점에 무엇이 밀려 있는지는 앱이 항상 보여 준다(다음 반영까지 남은 시간).

사용
  python ledger_db.py --status          # 대기 건수·다음 반영 시각
  python ledger_db.py --intake          # updates/pending_updates.json 을 DB로 흡수
  python ledger_db.py --apply           # 지금이 반영 시각이면 반영(아니면 아무것도 안 함)
  python ledger_db.py --apply --force   # 시각을 무시하고 즉시(긴급용, 이유가 기록된다)
  python ledger_db.py --self-test
"""
import sys, os, json, sqlite3, subprocess, glob, tempfile, time
from datetime import datetime, timedelta, time as dtime
from contextlib import contextmanager

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

DB_DIR = os.path.join(ROOT, "db")
DB_PATH = os.path.join(DB_DIR, "ledger_queue.db")
JSON_QUEUE = os.path.join(ROOT, "updates", "pending_updates.json")
REPORT_DIR = os.path.join(ROOT, "reports")
STATUS_CACHE = os.path.join(ROOT, "reports", "반영대기.json")
APPLY_LOCK = os.path.join(ROOT, "reports", ".ledger_db_apply.lock")

# ★ 사용자 확정: 하루 딱 두 번
WINDOWS = (dtime(11, 0), dtime(15, 0))
GRACE_MIN = 45          # 작업 스케줄러가 조금 늦게 시작해도 같은 회차로 인정

SCHEMA = """
CREATE TABLE IF NOT EXISTS pending(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,              -- 넣은 시각
  source TEXT NOT NULL,          -- app | claude | codex | tool 이름
  sheet TEXT NOT NULL,
  key_col TEXT, key TEXT, cell TEXT, col TEXT,
  value TEXT, vtype TEXT DEFAULT 'text',
  evidence TEXT,
  only_if_empty INTEGER DEFAULT 1,
  ingest_key TEXT,                 -- JSON staging 파일+순번(중단 후 재시도 중복 방지)
  status TEXT NOT NULL DEFAULT 'pending',   -- pending | applied | skipped
  batch_id INTEGER,
  applied_at TEXT
);
CREATE INDEX IF NOT EXISTS ix_pending_status ON pending(status);
CREATE TABLE IF NOT EXISTS batch(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  slot TEXT NOT NULL,            -- 어느 회차인가(2026-07-30 11:00)
  started TEXT, finished TEXT,
  cells INTEGER, ok INTEGER, note TEXT, forced INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS ux(              -- 앱 사용 기록(다음 개선의 근거)
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL, kind TEXT NOT NULL,     -- view | tap | search | error | slow
  target TEXT, detail TEXT, ms INTEGER
);
CREATE INDEX IF NOT EXISTS ix_ux_kind ON ux(kind);
CREATE TABLE IF NOT EXISTS handoff(         -- 19_AI작업인수인계 예약
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,
  title TEXT NOT NULL,
  detail TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending',
  applied_at TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS ix_handoff_pending
  ON handoff(title,detail) WHERE status='pending';
"""


@contextmanager
def conn():
    """★ sqlite3 의 `with` 는 **트랜잭션**만 끝낼 뿐 연결을 닫지 않는다.
    윈도우에서는 열린 연결이 파일을 물고 있어 DB 파일을 지울 수 없다(자체검증이 여기서 실패했다).
    그래서 커밋과 닫기를 함께 책임지는 컨텍스트 매니저를 따로 둔다."""
    os.makedirs(DB_DIR, exist_ok=True)
    c = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c.execute("PRAGMA journal_mode=WAL")     # 동시에 읽고 써도 막히지 않는다
        c.executescript(SCHEMA)
        # 기존 DB도 안전하게 올린다. CREATE TABLE IF NOT EXISTS만으로는 새 열이 생기지 않는다.
        cols = {row[1] for row in c.execute("PRAGMA table_info(pending)").fetchall()}
        if "ingest_key" not in cols:
            try:
                c.execute("ALTER TABLE pending ADD COLUMN ingest_key TEXT")
            except sqlite3.OperationalError as exc:
                if "duplicate column" not in str(exc).lower():
                    raise
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ix_pending_ingest"
                  " ON pending(ingest_key) WHERE ingest_key IS NOT NULL")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ix_batch_done_slot"
                  " ON batch(slot) WHERE ok=1")
        yield c
        c.commit()
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def _pid_alive(pid):
    try:
        os.kill(int(pid), 0)
        return True
    except (OSError, ValueError):
        return False


@contextmanager
def apply_lock():
    """11시·15시 작업이 겹쳐 같은 vN+1을 두 번 만들지 않게 한다."""
    os.makedirs(os.path.dirname(APPLY_LOCK), exist_ok=True)
    owned = False
    for _ in range(2):
        try:
            fd = os.open(APPLY_LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{os.getpid()} {datetime.now().isoformat()}".encode("ascii"))
            os.close(fd)
            owned = True
            break
        except FileExistsError:
            try:
                pid = int(open(APPLY_LOCK, encoding="ascii").read().split()[0])
            except Exception:
                pid = 0
            if pid and _pid_alive(pid):
                raise RuntimeError(f"원장 DB 일괄반영이 이미 실행 중입니다(PID {pid})")
            try:
                os.unlink(APPLY_LOCK)
            except FileNotFoundError:
                pass
    if not owned:
        raise RuntimeError("원장 DB 일괄반영 잠금을 만들 수 없습니다")
    try:
        yield
    finally:
        try:
            os.unlink(APPLY_LOCK)
        except FileNotFoundError:
            pass


@contextmanager
def json_queue_lock(path, timeout=30):
    """기존 ledger_writer.queue_add와 같은 `.lock`을 사용해 JSON 인계를 원자화한다."""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    lock = path + ".lock"
    started = time.monotonic()
    fd = None
    while fd is None:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{os.getpid()} {datetime.now().isoformat()}".encode("ascii"))
        except FileExistsError:
            if time.monotonic() - started >= timeout:
                raise TimeoutError(f"JSON 큐 잠금 대기 초과: {lock}")
            time.sleep(0.1)
    try:
        yield
    finally:
        os.close(fd)
        try:
            os.unlink(lock)
        except FileNotFoundError:
            pass


# ── 시각 판정 (순수 함수 — 합성 검증 대상) ──────────────────────
def slot_of(now, windows=WINDOWS, grace=GRACE_MIN):
    """지금이 어느 반영 회차인가. 아니면 None.

    각 시각부터 grace 분 안이면 그 회차로 친다. 스케줄러가 조금 늦어도(또는 PC가 잠깐
    꺼져 있어도) 그 회차를 놓치지 않게 하려는 것이다."""
    for w in windows:
        start = now.replace(hour=w.hour, minute=w.minute, second=0, microsecond=0)
        if start <= now < start + timedelta(minutes=grace):
            return f"{start:%Y-%m-%d %H:%M}"
    return None


def next_window(now, windows=WINDOWS):
    """다음 반영 시각(넘어가면 내일 첫 회차)."""
    today = [now.replace(hour=w.hour, minute=w.minute, second=0, microsecond=0) for w in windows]
    for t in today:
        if t > now:
            return t
    first = windows[0]
    return (now + timedelta(days=1)).replace(hour=first.hour, minute=first.minute,
                                             second=0, microsecond=0)


def missed_slots(now, done_slots, windows=WINDOWS, days_back=2):
    """PC가 꺼져 있어 건너뛴 회차(표시용).

    실제 반영은 이 목록을 이유로 임의 시각에 실행하지 않는다. 대기 항목은 다음
    11:00/15:00 회차에 함께 처리해 '하루 두 번' 규칙을 지킨다.
    """
    out = []
    for d in range(days_back, -1, -1):
        day = now - timedelta(days=d)
        for w in windows:
            t = day.replace(hour=w.hour, minute=w.minute, second=0, microsecond=0)
            if t <= now:
                s = f"{t:%Y-%m-%d %H:%M}"
                if s not in set(done_slots or []):
                    out.append(s)
    return out


def eligible_slot(now, done_slots, force=False):
    """이번 실행이 실제 반영할 수 있는 회차인가."""
    if force:
        return f"{now:%Y-%m-%d %H:%M}(강제)"
    slot = slot_of(now)
    if not slot or slot in set(done_slots or []):
        return None
    return slot


# ── 적재 ─────────────────────────────────────────────────────
FIELDS = ("sheet", "key_col", "key", "cell", "col", "value", "vtype", "evidence")


def enqueue(items, source="tool", ingest_prefix=None):
    """반영할 셀을 DB에 넣는다. **여기서는 엑셀을 건드리지 않는다.**"""
    rows = []
    now = datetime.now().isoformat(timespec="seconds")
    for pos, it in enumerate(items or []):
        d = dict(it or {})
        v = d.get("value")
        rows.append((now, source, d.get("sheet") or "", d.get("key_col"), d.get("key"),
                     d.get("cell"), d.get("col"),
                     v if isinstance(v, str) else json.dumps(v, ensure_ascii=False, default=str),
                     d.get("vtype") or "text", d.get("evidence"),
                     1 if d.get("only_if_empty", True) else 0,
                     f"{ingest_prefix}:{pos}" if ingest_prefix else None))
    if not rows:
        return 0
    with conn() as c:
        before = c.total_changes
        c.executemany(
            "INSERT OR IGNORE INTO pending"
            "(ts,source,sheet,key_col,key,cell,col,value,vtype,evidence,only_if_empty,ingest_key)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?)", rows)
        return c.total_changes - before


def intake_json(path=JSON_QUEUE, source="tool"):
    """기존 도구들이 쓰는 JSON 큐를 DB로 흡수한다.

    ★ 도구를 전부 뜯어고치지 않고 갈아타기 위한 다리다. 도구는 지금처럼 `--queue` 로
      JSON 에 넣고, 이 함수가 그것을 DB 로 옮긴 뒤 JSON 을 비운다."""
    from ledger_writer import atomic_json_dump
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    # 1) 공용 큐를 잠근 채 staging으로 떼어 낸다. 이후 새 입력은 즉시 빈 공용 큐에 쌓인다.
    with json_queue_lock(path):
        try:
            items = json.load(open(path, encoding="utf-8"))
        except Exception:
            items = []
        if isinstance(items, list) and items:
            stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
            stage = f"{path}.intake.{stamp}.{os.getpid()}.json"
            os.replace(path, stage)
            atomic_json_dump([], path)

    # 2) 중간에 죽어 남은 staging도 다시 처리한다. ingest_key가 같은 행의 재흡수를 막는다.
    added = 0
    for stage in sorted(glob.glob(path + ".intake.*.json")):
        try:
            batch = json.load(open(stage, encoding="utf-8"))
            if not isinstance(batch, list):
                continue
            added += enqueue(batch, source=source, ingest_prefix=os.path.basename(stage))
            os.unlink(stage)
        except Exception:
            # 원문 staging을 남겨 다음 실행이 이어받게 한다.
            continue
    return added


def pending_rows():
    with conn() as c:
        cur = c.execute("SELECT id,sheet,key_col,key,cell,col,value,vtype,evidence,only_if_empty"
                        " FROM pending WHERE status='pending' ORDER BY id")
        return [dict(zip(("id", "sheet", "key_col", "key", "cell", "col", "value", "vtype",
                          "evidence", "only_if_empty"), r)) for r in cur.fetchall()]


def handoff_add(title, detail):
    """19시트 인수인계를 Excel 대신 DB에 예약한다."""
    title = str(title or "").strip()
    detail = str(detail or "").strip()
    if not title or not detail:
        raise ValueError("인수인계 제목과 상세가 모두 필요합니다")
    with conn() as c:
        before = c.total_changes
        c.execute(
            "INSERT OR IGNORE INTO handoff(ts,title,detail,status) VALUES(?,?,?,'pending')",
            (datetime.now().isoformat(timespec="seconds"), title[:500], detail[:4000]),
        )
        return c.total_changes - before


def pending_handoffs():
    with conn() as c:
        rows = c.execute(
            "SELECT id,title,detail FROM handoff WHERE status='pending' ORDER BY id"
        ).fetchall()
    return [{"id": r[0], "title": r[1], "detail": r[2]} for r in rows]


def counts():
    with conn() as c:
        p = c.execute("SELECT COUNT(*) FROM pending WHERE status='pending'").fetchone()[0]
        by = c.execute("SELECT source,COUNT(*) FROM pending WHERE status='pending'"
                       " GROUP BY source ORDER BY 2 DESC").fetchall()
        done = [r[0] for r in c.execute("SELECT slot FROM batch WHERE ok=1").fetchall()]
    return p, dict(by), done


def status(now=None):
    now = now or datetime.now()
    p, by, done = counts()
    nxt = next_window(now)
    handoffs = len(pending_handoffs())
    doc = {"확인": now.isoformat(timespec="seconds"), "대기": p, "인수인계대기": handoffs,
           "출처별": by,
           "다음반영": nxt.isoformat(timespec="minutes"),
           "남은분": max(0, int((nxt - now).total_seconds() // 60)),
           "지금회차": slot_of(now), "밀린회차": missed_slots(now, done),
           "반영시각": [f"{w.hour:02d}:{w.minute:02d}" for w in WINDOWS]}
    os.makedirs(os.path.dirname(STATUS_CACHE), exist_ok=True)
    json.dump(doc, open(STATUS_CACHE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return doc


# ── UX 기록 ──────────────────────────────────────────────────
def ux_add(events):
    """앱이 보내는 사용 기록. **개인정보가 아니라 화면 사용 흔적만** 담는다."""
    rows = []
    now = datetime.now().isoformat(timespec="seconds")
    for e in events or []:
        d = dict(e or {})
        rows.append((d.get("ts") or now, str(d.get("kind") or "tap")[:20],
                     str(d.get("target") or "")[:120], str(d.get("detail") or "")[:300],
                     int(d.get("ms") or 0)))
    if not rows:
        return 0
    with conn() as c:
        c.executemany("INSERT INTO ux(ts,kind,target,detail,ms) VALUES(?,?,?,?,?)", rows)
        cutoff = (datetime.now() - timedelta(days=90)).isoformat(timespec="seconds")
        c.execute("DELETE FROM ux WHERE ts < ?", (cutoff,))
    return len(rows)


def ux_summary(days=7, limit=15):
    since = (datetime.now() - timedelta(days=days)).isoformat(timespec="seconds")
    with conn() as c:
        def q(sql, *a):
            return c.execute(sql, a).fetchall()
        return {
            "기간": f"최근 {days}일",
            "화면별": q("SELECT target,COUNT(*) FROM ux WHERE kind='view' AND ts>=?"
                      " GROUP BY target ORDER BY 2 DESC LIMIT ?", since, limit),
            "많이누른것": q("SELECT target,COUNT(*) FROM ux WHERE kind='tap' AND ts>=?"
                        " GROUP BY target ORDER BY 2 DESC LIMIT ?", since, limit),
            "오류": q("SELECT target,detail,COUNT(*) FROM ux WHERE kind='error' AND ts>=?"
                    " GROUP BY target,detail ORDER BY 3 DESC LIMIT ?", since, limit),
            "느린화면": q("SELECT target,MAX(ms),COUNT(*) FROM ux WHERE kind='slow' AND ts>=?"
                      " GROUP BY target ORDER BY 2 DESC LIMIT ?", since, limit),
            "빈손검색": q("SELECT detail,COUNT(*) FROM ux WHERE kind='search' AND ms=0 AND ts>=?"
                      " GROUP BY detail ORDER BY 2 DESC LIMIT ?", since, limit),
        }


# ── 반영 ─────────────────────────────────────────────────────
def scheduled_workbook_maintenance(now=None):
    """11:00·15:00 회차 안에서만 구조 시트·수식 캐시를 갱신한다.

    확정 셀 입력뿐 아니라 23·24·25·27·28 시트와 Excel 재계산까지 같은 회차로 묶어,
    09:50 자동대조가 별도 vN+1을 만드는 우회 경로를 없앤다. 각 도구는 멱등이라 내용이
    같으면 버전을 만들지 않는다. 한 단계 실패가 이미 성공한 셀 입력을 되돌리지는 않으며
    다음 회차에서 다시 시도할 수 있도록 보고서에 단계별 결과를 남긴다.
    """
    from ledger_writer import atomic_json_dump
    try:
        from inbox_scan import pick
        has_tax = bool(pick("tax"))
    except Exception:
        has_tax = False

    env = {**os.environ, "PYTHONIOENCODING": "utf-8",
           "COUPANG_LEDGER_GATE": "1", "CSOS_AI": "scheduler"}
    jobs = []
    if has_tax:
        jobs.append(("25_ERP매출서류", [os.path.join(ROOT, "erp_docs_check.py"), "--sheet"]))
    jobs.extend([
        ("27_정기점검원본일정", [os.path.join(ROOT, "pm_schedule_sync.py"), "--apply"]),
        ("28_일지대조현황", [os.path.join(ROOT, "work_log_sync.py"), "--apply"]),
    ])
    band_cache = glob.glob(os.path.join(ROOT, "band", "cache", "*.json"))
    if any(not os.path.basename(p).startswith(("raw_", "dump_")) for p in band_cache):
        jobs.append(("24_밴드업무추출", [os.path.join(ROOT, "band_extract.py"), "--sheet"]))
    jobs.extend([
        ("23_확인필요현황", [os.path.join(ROOT, "findings_sheet.py")]),
        ("워크북 무결성 복구", [os.path.join(ROOT, "fix_workbook.py"), "--apply"]),
        ("Excel 수식 재계산", [os.path.join(ROOT, "excel_recalc.py"), "--run"]),
    ])

    results = []
    for name, cmd in jobs:
        try:
            r = subprocess.run(
                [sys.executable, *cmd], cwd=ROOT, capture_output=True, text=True,
                encoding="utf-8", errors="replace", timeout=1800, env=env,
            )
            lines = [x.strip() for x in ((r.stdout or "") + "\n" + (r.stderr or "")).splitlines()
                     if x.strip()]
            results.append({"단계": name, "성공": r.returncode == 0,
                            "메모": (lines[-1] if lines else "")[:240]})
        except Exception as exc:
            results.append({"단계": name, "성공": False,
                            "메모": f"{type(exc).__name__}: {exc}"[:240]})
    # 인수인계는 이 회차에서 만들어진 최종본의 19시트에 마지막으로 기록한다.
    for item in pending_handoffs():
        try:
            r = subprocess.run(
                [sys.executable, os.path.join(ROOT, "workbook_patch.py"),
                 "--b", item["title"], "--c", item["detail"]],
                cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=1800, env=env,
            )
            lines = [x.strip() for x in ((r.stdout or "") + "\n" + (r.stderr or "")).splitlines()
                     if x.strip()]
            ok = r.returncode == 0
            results.append({"단계": "19_AI작업인수인계", "성공": ok,
                            "메모": (lines[-1] if lines else "")[:240]})
            if ok:
                with conn() as c:
                    c.execute(
                        "UPDATE handoff SET status='applied',applied_at=?"
                        " WHERE id=? AND status='pending'",
                        (datetime.now().isoformat(timespec="seconds"), item["id"]),
                    )
        except Exception as exc:
            results.append({"단계": "19_AI작업인수인계", "성공": False,
                            "메모": f"{type(exc).__name__}: {exc}"[:240]})
    os.makedirs(REPORT_DIR, exist_ok=True)
    atomic_json_dump(
        {"시각": (now or datetime.now()).isoformat(timespec="seconds"), "결과": results},
        os.path.join(REPORT_DIR, "scheduled_workbook_maintenance.json"),
    )
    try:
        import ai_claim
        ai_claim.free("scheduler", "ledger")
    except Exception:
        pass
    return results


MASTER_LOCK_GLOB = "~$쿠팡_통합업무_일일보고_관리대장*.xlsx"
LOCK_STALE_HOURS = 24     # 이보다 오래된 잠금만 크래시 잔재로 본다
LOCK_POLL_SEC = 180       # 잠금이 풀리기를 기다리는 간격


def _master_folder():
    try:
        cfg = json.load(open(os.path.join(ROOT, "config", "ecount_config.json"),
                             encoding="utf-8"))
        return os.path.dirname(cfg["reconcile"]["master_xlsx"])
    except (OSError, KeyError, ValueError):
        return ""


def human_editing(folder=None):
    """사람이 관리대장을 열어 두었는가 — **네트워크 공유의 진실은 ~$ 잠금파일뿐이다.**

    2026-07-31 실사고: 류지영 매니저가 **다른 PC** 에서 v331 을 열어 입력하는 동안
    15:05 반영이 v336 을 만들어 그녀의 15:43 저장이 고아가 됐다. 그때 잠금파일이
    있었는데 '이 PC 에 EXCEL 프로세스가 없다'며 잔재로 잘못 판정했다 — 로컬 프로세스로는
    다른 PC 의 편집을 볼 수 없다. 잠금이 있으면 사람이 있다고 본다.
    (잘못 기다린 손해는 반영이 늦는 것뿐이지만, 잘못 진행한 손해는 사람 입력 유실이다.
     LOCK_STALE_HOURS 를 넘긴 잠금만 크래시 잔재로 보고 지나간다.)"""
    folder = folder or _master_folder()
    if not folder:
        return None
    locks = None
    for _attempt in range(3):                       # Z: 는 순간적으로 끊긴다 — 재시도
        try:
            locks = glob.glob(os.path.join(folder, MASTER_LOCK_GLOB))
            break
        except OSError:
            time.sleep(2)
    if not locks:
        return None
    out = []
    for p in locks:
        try:
            age_min = (time.time() - os.path.getmtime(p)) / 60
        except OSError:
            continue
        if age_min >= LOCK_STALE_HOURS * 60:
            continue
        who = ""
        try:
            raw = open(p, "rb").read(64)
            if raw and 0 < raw[0] < 60:
                who = raw[1:1 + raw[0]].decode("cp949", "replace").strip()
        except OSError:
            pass
        out.append({"잠금": os.path.basename(p), "소유자": who, "분": int(age_min)})
    return out or None


def _wait_editing_clear(now, slot_name):
    """사람이 열어 둔 동안은 쓰지 않는다 — 회차 유예(GRACE_MIN) 안에서 기다렸다가,
    끝내 안 풀리면 이 회차를 **건너뛴다**(batch 에 기록하지 않으므로 큐는 남고,
    missed_slots 가 다음 실행에서 마저 처리한다)."""
    deadline = datetime.strptime(slot_name, "%Y-%m-%d %H:%M") + timedelta(minutes=GRACE_MIN)
    while True:
        locks = human_editing()
        if not locks:
            return None
        if datetime.now() >= deadline:
            return locks
        left = int((deadline - datetime.now()).total_seconds() // 60)
        print(f"  사람이 관리대장을 열어 두었습니다({locks[0].get('소유자') or '?'}) — "
              f"잠금 해제 대기(남은 유예 {left}분)")
        time.sleep(LOCK_POLL_SEC)


def apply_now(force=False, now=None):
    """정해진 시각일 때만 엑셀에 쓴다. 실제 쓰기는 기존 ledger_writer 에 맡긴다."""
    now = now or datetime.now()
    from operation_window import is_input_window, input_window_label
    if is_input_window(now):
        return {"상태": "보류", "사유": f"입력 보호시간({input_window_label()})"}

    intake_json()                                    # 도구들이 넣어 둔 것 먼저 흡수
    p, by, done = counts()
    slot_name = eligible_slot(now, done, force)
    if not slot_name:
        nxt = next_window(now)
        current = slot_of(now)
        why = "이미 처리한 회차" if current and current in set(done) else "반영 시각이 아님"
        return {"상태": "대기", "사유": f"{why} — 다음 {nxt:%m-%d %H:%M}",
                "대기": p}
    # ★ 쓰기 직전 관문 — 사람이 관리대장을 열어 두었으면 이 회차를 양보한다(2026-07-31 실사고).
    #   구조 갱신(scheduled_workbook_maintenance)도 vN+1 을 만들므로 같이 막아야 한다.
    #   --force 도 뚫지 못한다: 강제의 용도는 '시각 밖 반영'이지 '사람을 밀어내기'가 아니다.
    locks = _wait_editing_clear(now, slot_name)
    if locks:
        who = locks[0].get("소유자") or "?"
        return {"상태": "보류", "회차": slot_name,
                "사유": f"사람이 관리대장 편집 중({who}, {locks[0]['분']}분째) — "
                        f"회차를 건너뛰고 다음 실행에서 처리", "대기": p}
    if p == 0:
        # 빈 회차도 완료로 기록해야 같은 시간대 재실행이 새 버전을 만들지 않는다.
        if not force:
            with conn() as c:
                c.execute(
                    "INSERT OR IGNORE INTO batch(slot,started,finished,cells,ok,note,forced)"
                    " VALUES(?,?,?,?,?,?,0)",
                    (slot_name, now.isoformat(timespec="seconds"),
                     now.isoformat(timespec="seconds"), 0, 1, "반영할 항목 없음"),
                )
        maintenance = scheduled_workbook_maintenance(now)
        status(now)
        return {"상태": "없음", "회차": slot_name, "사유": "확정 셀 입력 없음",
                "대기": 0, "구조갱신": maintenance}

    rows = pending_rows()
    payload = []
    for r in rows:
        d = {"sheet": r["sheet"], "value": r["value"], "vtype": r["vtype"],
             "evidence": r["evidence"], "only_if_empty": bool(r["only_if_empty"])}
        if r["cell"]:
            d["cell"] = r["cell"]
        else:
            d.update({"key_col": r["key_col"], "key": r["key"], "col": r["col"]})
        payload.append(d)
    with conn() as c:
        cur = c.execute("INSERT INTO batch(slot,started,cells,forced) VALUES(?,?,?,?)",
                        (slot_name, now.isoformat(timespec="seconds"), len(payload),
                         1 if force else 0))
        batch_id = cur.lastrowid

    # 공용 JSON 큐를 다시 쓰지 않는다. 전용 배치 파일을 ledger_writer의 --queue로 넘겨
    # 실패 후 재흡수 중복과, 반영 도중 들어온 새 입력의 유실을 모두 막는다.
    from ledger_writer import atomic_json_dump
    batch_queue = os.path.join(ROOT, "updates", f".ledger_db_batch_{batch_id}.json")
    atomic_json_dump(payload, batch_queue)
    try:
        r = subprocess.run(
            [sys.executable, os.path.join(ROOT, "ledger_writer.py"),
             "--queue", batch_queue, "--apply"],
            cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=1800,
            env={**os.environ, "PYTHONIOENCODING": "utf-8",
                 "COUPANG_LEDGER_GATE": "1", "CSOS_AI": "scheduler"},
        )
        ok = r.returncode == 0
        output = "\n".join(x for x in (r.stdout or "", r.stderr or "") if x)
    except Exception as exc:
        ok = False
        output = f"{type(exc).__name__}: {exc}"
    finally:
        try:
            os.unlink(batch_queue)
        except FileNotFoundError:
            pass
    tail = [line for line in output.splitlines() if line.strip()][-1:] or [""]
    with conn() as c:
        c.execute("UPDATE batch SET finished=?,ok=?,note=? WHERE id=?",
                  (datetime.now().isoformat(timespec="seconds"), 1 if ok else 0,
                   tail[0][:200], batch_id))
        if ok and rows:
            ids = [row["id"] for row in rows]
            marks = ",".join("?" for _ in ids)
            c.execute(
                f"UPDATE pending SET status='applied',batch_id=?,applied_at=?"
                f" WHERE status='pending' AND id IN ({marks})",
                (batch_id, datetime.now().isoformat(timespec="seconds"), *ids),
            )
    if ok:
        maintenance = scheduled_workbook_maintenance(now)
    else:
        maintenance = []
        try:
            import ai_claim
            ai_claim.free("scheduler", "ledger")
        except Exception:
            pass
    status(now)
    return {"상태": "반영" if ok else "실패", "회차": slot_name, "셀": len(payload),
            "메모": tail[0][:120], "구조갱신": maintenance}


def backup_to(dst):
    """WAL 사용 중에도 일관된 SQLite 복구본을 만든다."""
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=os.path.basename(dst) + ".", suffix=".tmp",
                               dir=os.path.dirname(dst) or ".")
    os.close(fd)
    try:
        with conn() as source, sqlite3.connect(tmp) as target:
            source.backup(target)
        os.replace(tmp, dst)
    finally:
        try:
            if os.path.exists(tmp):
                os.unlink(tmp)
        except OSError:
            pass
    return dst


# ── 합성 검증 ─────────────────────────────────────────────────
def self_test():
    bad = 0
    D = datetime
    # 11:00·15:00 회차 판정
    # ★ 유예(GRACE_MIN)는 조정될 수 있다(2026-07-30: 90→45분). 값을 고정해 시험하면
    #   설정을 바꿀 때마다 검증이 깨진다 — **설정값을 기준으로** 경계를 만든다.
    from datetime import timedelta as _td
    base = D(2026, 7, 30, 11, 0)
    for delta, want in ((0, True), (GRACE_MIN - 1, True), (GRACE_MIN, False),
                        (GRACE_MIN + 1, False), (-1, False)):
        n = base + _td(minutes=delta)
        got = slot_of(n) is not None
        if got != want:
            print(f"  [FAIL] slot_of 11:00+{delta}분({n:%H:%M}) → {got}"); bad += 1
    if slot_of(D(2026, 7, 30, 15, 5)) is None:
        print("  [FAIL] 15시 회차를 인식하지 못한다"); bad += 1
    # 다음 시각
    if next_window(D(2026, 7, 30, 9, 0)).hour != 11:
        print("  [FAIL] 다음 시각(오전)"); bad += 1
    if next_window(D(2026, 7, 30, 12, 0)).hour != 15:
        print("  [FAIL] 다음 시각(오후)"); bad += 1
    nd = next_window(D(2026, 7, 30, 16, 0))
    if nd.hour != 11 or nd.day != 31:
        print("  [FAIL] 다음 시각(내일)"); bad += 1
    # 놓친 회차는 버리지 않는다
    ms = missed_slots(D(2026, 7, 30, 16, 0), ["2026-07-30 11:00"], days_back=0)
    if ms != ["2026-07-30 15:00"]:
        print("  [FAIL] 밀린 회차", ms); bad += 1
    if missed_slots(D(2026, 7, 30, 16, 0), ["2026-07-30 11:00", "2026-07-30 15:00"], days_back=0):
        print("  [FAIL] 이미 한 회차를 또 하려 한다"); bad += 1
    if eligible_slot(D(2026, 7, 30, 13, 0), []) is not None:
        print("  [FAIL] 회차 밖 반영 허용"); bad += 1
    if eligible_slot(D(2026, 7, 30, 11, 5), ["2026-07-30 11:00"]) is not None:
        print("  [FAIL] 같은 회차 중복 허용"); bad += 1
    # 하루 두 번뿐이다
    if len(WINDOWS) != 2 or [w.hour for w in WINDOWS] != [11, 15]:
        print("  [FAIL] 반영 시각이 11·15시가 아니다"); bad += 1
    # DB 왕복
    global DB_PATH
    import tempfile
    old = DB_PATH
    with tempfile.TemporaryDirectory() as td:
        DB_PATH = os.path.join(td, "t.db")
        try:
            item = {"sheet": "02_돌발AS접수", "cell": "C9", "value": "AS-1",
                    "evidence": "테스트"}
            n = enqueue([item], source="claude", ingest_prefix="same")
            if n != 1 or len(pending_rows()) != 1:
                print("  [FAIL] DB 적재"); bad += 1
            if enqueue([item], source="claude", ingest_prefix="same") != 0:
                print("  [FAIL] staging 재시도 중복"); bad += 1
            if enqueue([], source="x") != 0:
                print("  [FAIL] 빈 목록"); bad += 1
            if ux_add([{"kind": "tap", "target": "정산"}]) != 1:
                print("  [FAIL] UX 기록"); bad += 1
            if handoff_add("테스트", "19시트 예약") != 1 or len(pending_handoffs()) != 1:
                print("  [FAIL] 인수인계 예약"); bad += 1
            if handoff_add("테스트", "19시트 예약") != 0:
                print("  [FAIL] 인수인계 예약 중복"); bad += 1
            p, by, _ = counts()
            if p != 1 or by.get("claude") != 1:
                print("  [FAIL] 집계", p, by); bad += 1
        finally:
            DB_PATH = old
    print("ledger_db self-test:", "OK" if not bad else f"{bad}건 실패")
    return bad == 0


def main():
    if "--self-test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    if "--intake" in sys.argv:
        print(f"JSON 큐 → DB 흡수 {intake_json()}건")
    if "--handoff" in sys.argv:
        try:
            title = sys.argv[sys.argv.index("--b") + 1]
            detail = sys.argv[sys.argv.index("--c") + 1]
        except (ValueError, IndexError):
            sys.exit("사용: python ledger_db.py --handoff --b \"제목\" --c \"상세\"")
        n = handoff_add(title, detail)
        print("19시트 인수인계 DB 예약:", "추가 1건" if n else "이미 같은 예약 있음")
        print("Excel 기록은 다음 11:00·15:00 회차 마지막에 수행")
        return
    if "--apply" in sys.argv:
        with apply_lock():
            r = apply_now(force="--force" in sys.argv)
        print(" · ".join(f"{k} {v}" for k, v in r.items()))
        if r.get("상태") == "실패":
            sys.exit(1)
        return
    d = status()
    print(f"반영 대기 {d['대기']}건 · 다음 반영 {d['다음반영']} (약 {d['남은분']}분 뒤)")
    if d["출처별"]:
        print("  출처:", ", ".join(f"{k} {v}건" for k, v in d["출처별"].items()))
    if d["인수인계대기"]:
        print(f"  19시트 인수인계 예약: {d['인수인계대기']}건")
    if d["밀린회차"]:
        print("  ★ 밀린 회차:", ", ".join(d["밀린회차"]))
    print(f"  반영 시각: 매일 {' · '.join(d['반영시각'])} (하루 두 번)")


if __name__ == "__main__":
    main()
