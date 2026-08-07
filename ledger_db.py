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
import sys, os, json, sqlite3, subprocess, glob, tempfile, time, re
from datetime import datetime, timedelta, time as dtime
from contextlib import contextmanager

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

# ★ 큐 DB 는 **본체 체크아웃 하나**만 쓴다 (2026-08-06).
#   워크트리에서 enqueue() 한 입력이 워크트리 안의 다른 DB 로 들어가면
#   11:00·15:00 반영(본체에서 돈다)이 그 값을 영영 못 본다 — 큐에 넣었으니
#   됐다고 믿는 동안 값이 사라진다. 본체에서는 경로가 예전과 글자 그대로 같다.
#   SQLite 라서 링크로 잇지 않는다(`-wal`·`-journal` 사이드카가 갈려 DB 가 깨진다).
try:
    from worktree_state import shared as _shared
    DB_DIR = _shared("db")
except Exception:
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
CREATE TABLE IF NOT EXISTS resolution(      -- 객관 입증으로 완료 처리한 건(사용자 지시 2026-07-31)
  id INTEGER PRIMARY KEY AUTOINCREMENT,     -- ★ 엑셀 셀에 백필하지 않는다 — DB 가 기록의 정본이다
  settle_id TEXT NOT NULL UNIQUE,           -- 06시트 정산ID
  project TEXT,
  status TEXT NOT NULL,                     -- 완료(ERP 수금확인) 등 — settle_status 판정값
  basis TEXT NOT NULL,                      -- 무엇이 입증했나(ERP 판매조회 진행상태=7.수금완료 등)
  first_seen TEXT NOT NULL,                 -- 처음 입증된 시각
  last_seen TEXT NOT NULL                   -- 마지막으로 같은 입증이 확인된 시각
);
CREATE TABLE IF NOT EXISTS work_resolution( -- AS·정기점검·설치의 객관 입증 완료(DB 정본)
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL,                       -- as | pm | install
  record_id TEXT NOT NULL,                  -- 접수ID | 점검ID | 업무ID (없으면 프로젝트NO)
  project TEXT,
  status TEXT NOT NULL,                     -- 작업완료 | 완료
  completed_on TEXT NOT NULL,               -- 원자료가 말한 실제 완료일
  basis TEXT NOT NULL,                      -- 밴드 완료글·검증 정상 등 근거
  first_seen TEXT NOT NULL,
  last_seen TEXT NOT NULL,
  UNIQUE(kind, record_id)
);
CREATE INDEX IF NOT EXISTS ix_work_resolution_project
  ON work_resolution(kind, project);
CREATE TABLE IF NOT EXISTS staff_resolution( -- 담당자별 객관 입증 완료(DB 정본)
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  owner TEXT NOT NULL,                      -- 류지영 | 오종현 | 유현민
  task_kind TEXT NOT NULL,                  -- field_as | settlement | po_source 등
  record_id TEXT NOT NULL,
  project TEXT,
  status TEXT NOT NULL,                     -- "류지영 완료"처럼 담당자까지 명시
  completed_on TEXT NOT NULL,
  basis TEXT NOT NULL,
  first_seen TEXT NOT NULL,
  last_seen TEXT NOT NULL,
  UNIQUE(owner, task_kind, record_id)
);
CREATE INDEX IF NOT EXISTS ix_staff_resolution_owner
  ON staff_resolution(owner, completed_on);
CREATE TABLE IF NOT EXISTS remote_issue(     -- 리모컨 불출 (2026-08-03 지시)
  id INTEGER PRIMARY KEY AUTOINCREMENT,      -- AS 담당자당 보유 3개 한도 · 부사장 승인 후 불출
  branch TEXT NOT NULL,                      -- 부산 | 시화 | 증평
  issuer TEXT NOT NULL,                      -- 불출 담당: 부산=오종현 · 시화=안은숙 · 증평=류지영
  technician TEXT NOT NULL,                  -- 받아 간 AS 담당자
  qty INTEGER NOT NULL,
  status TEXT NOT NULL,                      -- 승인대기 | 불출완료 | 반려
  requested_by TEXT NOT NULL,                -- 신청을 올린 사람(업무센터 로그인 주체)
  requested_at TEXT NOT NULL,
  approved_by TEXT,                          -- 부사장 승인자
  approved_at TEXT,
  note TEXT,
  issued_on TEXT,                            -- 불출 일자(공지 2026-08-04 — 입력일과 다를 수 있다)
  camp TEXT,                                 -- 투입 예정 캠프명(공지 2026-08-04)
  version TEXT                               -- 리모컨 버전(미확인·기존형·VER.3·VER.4)
);
CREATE TABLE IF NOT EXISTS remote_delivery(  -- 리모컨 납품 — 어느 프로젝트/캠프에 들어갔나
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  technician TEXT NOT NULL,
  project TEXT,                              -- 프로젝트NO(UJ…) — 캠프만 알아도 기록은 남긴다
  camp TEXT,
  qty INTEGER NOT NULL,
  delivered_on TEXT NOT NULL,
  note TEXT,
  created_by TEXT,
  created_at TEXT NOT NULL,
  kind TEXT,                                 -- 납품 | 사용 | 교체 | 샘플 | 택배출고 (2026-08-04)
  version TEXT
);
CREATE INDEX IF NOT EXISTS ix_remote_issue_tech ON remote_issue(technician, status);
CREATE INDEX IF NOT EXISTS ix_remote_delivery_tech ON remote_delivery(technician, delivered_on);
CREATE TABLE IF NOT EXISTS remote_stock(     -- 리모컨 지점 재고 조정 로그 (2026-08-03 지시)
  id INTEGER PRIMARY KEY AUTOINCREMENT,      -- 현재 재고 = 조정 합계 - 그 지점 불출 합계
  branch TEXT NOT NULL,                      -- 부산 | 시화 | 증평
  qty_delta INTEGER NOT NULL,                -- 입고 +N · 정정 ±N · 실사 맞춤도 델타로 기록
  reason TEXT,                               -- 입고 | 실사 | 정정 등
  created_by TEXT,
  created_at TEXT NOT NULL,
  version TEXT,                              -- 버전별 재고 구분(증평 기존형 20 · VER.4 49)
  moved_on TEXT                              -- 실제 입출고 일자(입력일과 다를 수 있다)
);
CREATE INDEX IF NOT EXISTS ix_remote_stock_branch ON remote_stock(branch, created_at);
CREATE TABLE IF NOT EXISTS remote_audit(     -- 리모컨 기록 수정·삭제 원장 (2026-08-06 지시)
  id INTEGER PRIMARY KEY AUTOINCREMENT,      -- 지운 것을 되돌릴 수 있어야 지우게 해 준다
  table_name TEXT NOT NULL,                  -- remote_issue | remote_delivery | remote_stock
  row_id INTEGER NOT NULL,
  action TEXT NOT NULL,                      -- 수정 | 삭제 | 복구
  before_json TEXT,                          -- 바꾸기 전 행 전체(복구의 근거)
  after_json TEXT,
  reason TEXT,
  forced INTEGER NOT NULL DEFAULT 0,         -- 한도·재고 경고를 무시하고 강제 저장했나
  actor TEXT,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_remote_audit_row ON remote_audit(table_name, row_id);
CREATE TABLE IF NOT EXISTS call_note(        -- 통화·회의 기록 (2026-08-07 지시: 민감 — DB 전용)
  id INTEGER PRIMARY KEY AUTOINCREMENT,      -- ★ Z: 공유 폴더에 원본을 두지 않는다. 여기가 정본이다.
  file TEXT NOT NULL UNIQUE,                 -- 원래 메모 파일 이름(식별자 — 같은 이름이면 갱신)
  on_date TEXT,                              -- 통화 일자
  whom TEXT,                                 -- 통화 상대
  body TEXT NOT NULL,                        -- 메모 본문(정본)
  todos_json TEXT,                           -- 뽑아낸 할 일(19시트 예약의 근거)
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_call_note_on ON call_note(on_date);
"""

# 리모컨 불출 규칙(2026-08-03 사용자 지시) — 지점별 불출 담당과 담당자당 보유 한도.
REMOTE_BRANCH_ISSUERS = {"부산": "오종현", "시화": "안은숙", "증평": "류지영"}
REMOTE_BRANCH_LABELS = {"부산": "부산공장", "시화": "시화공장", "증평": "증평본사"}
REMOTE_HOLD_LIMIT = 3
# 리모컨 버전(2026-08-06 지시: "버전 관리가 VER.3인지 VER.4인지 입력 및 확인 수정 가능하게").
# 재고표가 실제로 쓰는 이름 그대로다 — 화면 선택지도 서버도 이 목록 하나만 본다.
REMOTE_VERSIONS = ("미확인", "기존형", "VER.3", "VER.4")


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
        # 리모컨 공지(2026-08-04): 불출 일자·투입 예정 캠프명을 불출 기록에 남긴다.
        # 같은 날 2차: 실제 재고표가 버전(기존형·VER.3·VER.4)과 처리유형(사용·교체·
        # 샘플·택배출고)으로 관리되고 있어 세 표에 열을 더 붙인다.
        for table, cols in (("remote_issue", ("issued_on", "camp", "version")),
                            ("remote_delivery", ("kind", "version")),
                            ("remote_stock", ("version", "moved_on"))):
            have = {row[1] for row in c.execute(f"PRAGMA table_info({table})").fetchall()}
            for col in cols:
                if col not in have:
                    try:
                        c.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
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


def _dead_or_abandoned_lock(path, timeout):
    """PID가 죽었거나 소유자 없는 채 오래 남은 JSON 잠금만 회수한다."""
    try:
        words = open(path, encoding="ascii").read().split()
        if words:
            return not _pid_alive(int(words[0]))
    except (OSError, ValueError):
        pass
    try:
        return time.time() - os.path.getmtime(path) > timeout
    except OSError:
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
    owner = f"{os.getpid()} {datetime.now().isoformat()} {time.monotonic_ns()}"
    while fd is None:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, owner.encode("ascii"))
        except FileExistsError:
            if _dead_or_abandoned_lock(lock, timeout):
                try:
                    os.unlink(lock)
                except FileNotFoundError:
                    pass
                continue
            if time.monotonic() - started >= timeout:
                raise TimeoutError(f"JSON 큐 잠금 대기 초과: {lock}")
            time.sleep(0.1)
    try:
        yield
    finally:
        os.close(fd)
        try:
            if open(lock, encoding="ascii").read() == owner:
                os.unlink(lock)
        except OSError:
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


def resolution_sync(entries):
    """객관 입증 완료 건을 DB 에 기록한다(멱등 upsert).

    사용자 지시(2026-07-31): "객관적으로 입증이 되는 정보들은 모두 완료 처리해.
    최종 목적은 엑셀 파일에 저장 안 하고 DB 로만 관리하는 것."
    → 완료 상태는 엑셀 셀(06시트 발행일 등)에 써넣지 않는다. 원자료에 없는 값을
      지어내지 않기 위해서다(절대규칙 10). 대신 **언제부터 무엇이 입증했는지** 를
      이 표가 기억한다 — 입증이 사라져도(파일 교체) first_seen 이 남는다.
    entries: [{"settle_id","project","status","basis"}, ...]"""
    now = datetime.now().isoformat(timespec="seconds")
    n = 0
    with conn() as c:
        for e in entries or []:
            sid = str(e.get("settle_id") or "").strip()
            if not sid:
                continue
            c.execute(
                "INSERT INTO resolution(settle_id,project,status,basis,first_seen,last_seen)"
                " VALUES(?,?,?,?,?,?)"
                " ON CONFLICT(settle_id) DO UPDATE SET"
                "   project=excluded.project, status=excluded.status,"
                "   basis=excluded.basis, last_seen=excluded.last_seen",
                (sid, str(e.get("project") or ""), str(e.get("status") or ""),
                 str(e.get("basis") or ""), now, now))
            n += 1
    return n


def resolutions():
    """{정산ID: {status, basis, first_seen}} — 앱·보고서가 완료 근거를 보여줄 때 쓴다."""
    with conn() as c:
        return {row[0]: {"status": row[1], "basis": row[2], "first_seen": row[3]}
                for row in c.execute(
                    "SELECT settle_id,status,basis,first_seen FROM resolution")}


def resolution_retract(settle_ids):
    """현재 ERP에 완료·미완료 전표가 함께 확인된 정산만 정확한 ID로 철회한다.

    원천 파일이 잠시 사라진 경우에는 과거 근거를 보존한다. 이 함수는 호출자가 현재
    원본의 명시적 충돌을 확인한 ID만 넘길 때 사용하며, 담당자 정산 완료도 같은 키만
    함께 지워 두 정본이 어긋나지 않게 한다.
    """
    ids = sorted({str(value or "").strip() for value in settle_ids or []
                  if str(value or "").strip()})
    if not ids:
        return 0
    removed = 0
    with conn() as c:
        for settle_id in ids:
            cur = c.execute("DELETE FROM resolution WHERE settle_id=?", (settle_id,))
            removed += cur.rowcount
            c.execute(
                "DELETE FROM staff_resolution"
                " WHERE owner='류지영' AND task_kind='settlement' AND record_id=?",
                (settle_id,),
            )
    return removed


def work_resolution_sync(entries):
    """현장업무의 객관 완료 판정을 DB에 멱등 기록한다.

    완료 상태를 수식 셀에 덮어쓰지 않는다. 앱은 이 표를 원장보다 우선해 읽고,
    완료보고서·사진·ERP 누락은 기존 검증 열에서 계속 별도 경고한다.
    entries: [{kind,record_id,project,status,completed_on,basis}, ...]
    """
    now = datetime.now().isoformat(timespec="seconds")
    n = 0
    with conn() as c:
        for entry in entries or []:
            kind = str(entry.get("kind") or "").strip()
            record_id = str(entry.get("record_id") or "").strip()
            project = str(entry.get("project") or "").strip()
            completed_on = str(entry.get("completed_on") or "").strip()[:10]
            if not kind or not record_id or not completed_on:
                continue
            # A new pending row has no formula-generated AS/PM ID yet, so its
            # provisional record_id is the project number. Once Excel assigns
            # the canonical ID, migrate that row instead of counting it twice.
            if project and record_id == project:
                canonical = c.execute(
                    "SELECT record_id FROM work_resolution"
                    " WHERE kind=? AND project=? AND record_id<>? ORDER BY id LIMIT 1",
                    (kind, project, project),
                ).fetchone()
                if canonical:
                    record_id = canonical[0]
            elif project:
                provisional = c.execute(
                    "SELECT id,first_seen FROM work_resolution"
                    " WHERE kind=? AND record_id=?", (kind, project),
                ).fetchone()
                if provisional:
                    canonical = c.execute(
                        "SELECT id,first_seen FROM work_resolution"
                        " WHERE kind=? AND record_id=?", (kind, record_id),
                    ).fetchone()
                    if canonical:
                        first_seen = min(provisional[1], canonical[1])
                        c.execute("UPDATE work_resolution SET first_seen=? WHERE id=?",
                                  (first_seen, canonical[0]))
                        c.execute("DELETE FROM work_resolution WHERE id=?", (provisional[0],))
                    else:
                        c.execute("UPDATE work_resolution SET record_id=? WHERE id=?",
                                  (record_id, provisional[0]))
            c.execute(
                "INSERT INTO work_resolution(kind,record_id,project,status,completed_on,basis,first_seen,last_seen)"
                " VALUES(?,?,?,?,?,?,?,?)"
                " ON CONFLICT(kind,record_id) DO UPDATE SET"
                "   project=excluded.project, status=excluded.status,"
                "   completed_on=excluded.completed_on, basis=excluded.basis,"
                "   last_seen=excluded.last_seen",
                (kind, record_id, project,
                 str(entry.get("status") or ""), completed_on,
                 str(entry.get("basis") or ""), now, now))
            n += 1
    return n


def work_resolutions():
    """{(업무종류, ID 또는 프로젝트NO): 완료판정} — 앱 상태 오버레이용."""
    out = {}
    with conn() as c:
        rows = c.execute(
            "SELECT kind,record_id,project,status,completed_on,basis,first_seen"
            " FROM work_resolution").fetchall()
    for kind, record_id, project, status, completed_on, basis, first_seen in rows:
        value = {"status": status, "completed_on": completed_on,
                 "basis": basis, "first_seen": first_seen}
        out[(kind, record_id)] = value
        if project:
            out.setdefault((kind, project), value)
    return out


def staff_resolution_sync(entries):
    """객관 자료로 끝난 담당자별 업무를 SQLite 정본에 멱등 기록한다."""
    now = datetime.now().isoformat(timespec="seconds")
    allowed = {"류지영", "오종현", "유현민"}
    n = 0
    with conn() as c:
        for entry in entries or []:
            owner = str(entry.get("owner") or "").strip()
            task_kind = str(entry.get("task_kind") or "").strip()
            record_id = str(entry.get("record_id") or "").strip()
            completed_on = str(entry.get("completed_on") or "").strip()[:10]
            basis = str(entry.get("basis") or "").strip()
            if (owner not in allowed or not task_kind or not record_id
                    or not completed_on or not basis):
                continue
            try:
                completed_date = datetime.strptime(completed_on, "%Y-%m-%d").date()
            except ValueError:
                continue
            if completed_date > datetime.now().date():
                continue
            status = f"{owner} 완료"
            c.execute(
                "INSERT INTO staff_resolution"
                "(owner,task_kind,record_id,project,status,completed_on,basis,first_seen,last_seen)"
                " VALUES(?,?,?,?,?,?,?,?,?)"
                " ON CONFLICT(owner,task_kind,record_id) DO UPDATE SET"
                " project=excluded.project,status=excluded.status,"
                " completed_on=excluded.completed_on,basis=excluded.basis,"
                " last_seen=excluded.last_seen",
                (owner, task_kind, record_id, str(entry.get("project") or ""),
                 status, completed_on, basis, now, now),
            )
            n += 1
    return n


def staff_resolutions(owner="", limit=0):
    """담당자별 객관완료 목록을 최근 완료일부터 반환한다."""
    sql = ("SELECT owner,task_kind,record_id,project,status,completed_on,basis,first_seen,last_seen"
           " FROM staff_resolution")
    params = []
    if owner:
        sql += " WHERE owner=?"
        params.append(str(owner).strip())
    sql += " ORDER BY completed_on DESC,id DESC"
    if limit:
        sql += " LIMIT ?"
        params.append(int(limit))
    with conn() as c:
        rows = c.execute(sql, params).fetchall()
    keys = ("owner", "task_kind", "record_id", "project", "status",
            "completed_on", "basis", "first_seen", "last_seen")
    return [dict(zip(keys, row)) for row in rows]


def staff_resolution_retract(entries):
    """현재 대조가 비유일·충돌로 판정한 자동 완료만 정확한 키로 회수한다."""
    allowed = {"류지영", "오종현", "유현민"}
    removed = 0
    with conn() as c:
        before = c.total_changes
        for entry in entries or []:
            owner = str(entry.get("owner") or "").strip()
            task_kind = str(entry.get("task_kind") or "").strip()
            record_id = str(entry.get("record_id") or "").strip()
            if owner not in allowed or not task_kind or not record_id:
                continue
            c.execute(
                "DELETE FROM staff_resolution WHERE owner=? AND task_kind=? AND record_id=?",
                (owner, task_kind, record_id),
            )
        removed = c.total_changes - before
    return removed


def staff_resolution_summary():
    """세 담당자의 완료 건수와 최근 확인 시각을 빠르게 표시한다."""
    owners = ("류지영", "오종현", "유현민")
    with conn() as c:
        rows = {row[0]: {"completed": row[1], "last_seen": row[2]}
                for row in c.execute(
                    "SELECT owner,COUNT(*),MAX(last_seen) FROM staff_resolution GROUP BY owner")}
    return {owner: rows.get(owner, {"completed": 0, "last_seen": ""}) for owner in owners}


# ── 통화·회의 기록 (2026-08-07 지시: "통화_MD는 원본 자료에서 안 보이게, DB만 보관") ──
# ★ 왜 DB 만인가: 메모에는 사람 이름·평가·거래 조건이 섞인다. 예전에는 Z: 의
#   `0. 원본 자료/10. 통화·회의 기록/` 으로 복사해 두었는데 그 폴더는 **공유 폴더**라
#   앱 '원본 자료' 목록에 그대로 떴다(2026-08-07 실측: 통화_20260805_김준형.md 노출).
#   파일을 아예 두지 않으면 샐 곳이 없다 — 숨기는 것이 아니라 두지 않는 것이 조치다.

def call_note_save(file, body, whom="", on="", todos=None):
    """통화 메모 본문을 DB 에 보관한다. 같은 파일 이름이면 갱신한다."""
    now = datetime.now().isoformat(timespec="seconds")
    payload = json.dumps(todos or [], ensure_ascii=False)
    with conn() as c:
        cur = c.execute("UPDATE call_note SET on_date=?,whom=?,body=?,todos_json=?,updated_at=?"
                        " WHERE file=?", (on, whom, body, payload, now, file))
        if not cur.rowcount:
            c.execute("INSERT INTO call_note(file,on_date,whom,body,todos_json,created_at,updated_at)"
                      " VALUES(?,?,?,?,?,?,?)", (file, on, whom, body, payload, now, now))
    return True


def call_notes(limit=0, with_body=False):
    """보관된 통화 기록. **본문은 달라고 해야만 준다** — 목록·로그에 실수로 흘리지 않게."""
    cols = "file,on_date,whom,todos_json,created_at,updated_at" + (",body" if with_body else "")
    sql = f"SELECT {cols} FROM call_note ORDER BY on_date DESC, file DESC"
    if limit:
        sql += f" LIMIT {int(limit)}"
    names = cols.split(",")
    out = []
    with conn() as c:
        for row in c.execute(sql).fetchall():
            d = dict(zip(names, row))
            try:
                d["todos"] = json.loads(d.pop("todos_json") or "[]")
            except Exception:
                d["todos"] = []
            out.append(d)
    return out


def call_note_get(file):
    """본문까지 한 건. 없으면 None."""
    for n in call_notes(with_body=True):
        if n.get("file") == file:
            return n
    return None


# 보유로 잡는 불출 상태. '기초보유'는 2026-07-29 재고표에서 넘어온 개시 잔량이라
# 한도(3개) 검사를 받지 않는다 — 이미 손에 있는 물건을 없다고 할 수는 없기 때문이다.
# 대신 remote_status 가 한도 초과자를 따로 알려 준다(공지 4항의 회수·확인 대상).
REMOTE_HOLD_STATUSES = ("불출완료", "기초보유")


def _remote_version(v):
    """버전 표기를 하나로 모은다 — 사람은 'ver3'·'VER 3'·'v4'·'4' 를 다 쓴다.

    같은 물건이 표기만 달라 재고가 갈라지면 "VER.3 이 몇 개냐"에 답할 수 없다.
    빈 값은 빈 값으로 둔다(화면에서 '미확인'으로 보여 주는 것과 저장은 다른 일이다).
    """
    s = str(v or "").strip()
    if not s:
        return ""
    t = s.upper().replace(" ", "").replace("_", "").replace("-", "")
    if t in ("미확인", "UNKNOWN", "UNSURE"):
        return "미확인"
    if t in ("기존", "기존형", "OLD", "구형"):
        return "기존형"
    m = re.fullmatch(r"(?:VER\.?|V)?(\d+)", t)
    return f"VER.{m.group(1)}" if m else s


def _remote_holdings(c):
    """AS 담당자별 보유 = (불출 + 기초보유) 합 - 납품·사용 합.

    2026-08-06 지시("버전 관리가 VER.3인지 VER.4인지 … 확인 가능하게")로 **버전별
    내역**을 함께 준다. 합계만 있으면 "이 기사가 VER.3 을 몇 개 들고 있나"에 답할
    수 없었다 — 현장에서 맞는 버전을 안 들고 가면 다시 나가야 하는 일이 생긴다.
    버전을 안 적은 옛 기록은 '미확인'으로 모은다(빈칸으로 흩어 두지 않는다).
    """
    UNK = "COALESCE(NULLIF(version,''),'미확인')"
    issued = {row[0]: row[1] for row in c.execute(
        "SELECT technician,SUM(qty) FROM remote_issue"
        " WHERE status IN ('불출완료','기초보유') GROUP BY technician")}
    delivered = {row[0]: row[1] for row in c.execute(
        "SELECT technician,SUM(qty) FROM remote_delivery GROUP BY technician")}
    by_ver = {}
    for tech, ver, qty in c.execute(
            "SELECT technician,%s,SUM(qty) FROM remote_issue"
            " WHERE status IN ('불출완료','기초보유') GROUP BY technician,%s" % (UNK, UNK)):
        d = by_ver.setdefault(tech, {})
        d[ver] = d.get(ver, 0) + int(qty or 0)
    for tech, ver, qty in c.execute(
            "SELECT technician,%s,SUM(qty) FROM remote_delivery"
            " GROUP BY technician,%s" % (UNK, UNK)):
        d = by_ver.setdefault(tech, {})
        d[ver] = d.get(ver, 0) - int(qty or 0)
    out = {}
    for tech in set(issued) | set(delivered):
        got, used = int(issued.get(tech) or 0), int(delivered.get(tech) or 0)
        out[tech] = {"issued": got, "delivered": used, "holding": got - used,
                     # 0 이나 음수는 버린다 — 버전을 안 적고 납품만 잡힌 옛 기록 탓에
                     # 음수가 나오면 화면에서 "-2개 보유"로 읽혀 더 헷갈린다
                     "versions": {k: v for k, v in sorted((by_ver.get(tech) or {}).items())
                                  if v > 0}}
    return out


def remote_version_gaps(c, limit=40):
    """버전을 아직 안 적은 줄들 — 화면에서 바로 골라 채우게 하려고 모아 준다.

    2026-08-06 지시의 "입력 및 확인 수정 가능하게" 가 여기다. 버전이 비어 있으면
    재고가 '미확인' 더미로 쌓여, VER.3/VER.4 어느 쪽이 부족한지 알 수 없다.
    """
    out = []
    for kind, table, day_col, who_col in (
            ("issue", "remote_issue", "issued_on", "technician"),
            ("delivery", "remote_delivery", "delivered_on", "technician"),
            ("stock", "remote_stock", "moved_on", "branch")):
        qty_col = "qty_delta" if table == "remote_stock" else "qty"
        for rid, day, who, qty in c.execute(
                # 빈칸뿐 아니라 **'미확인'으로 저장된 줄도** 채울 대상이다.
                # 기초 재고표에서 넘어온 줄들이 그렇게 들어와 있어서, 빈칸만 찾으면
                # 24개가 미확인인데 "채울 것 4건"으로 보였다(2026-08-06 실측).
                "SELECT id,COALESCE(%s,''),COALESCE(%s,''),%s FROM %s"
                " WHERE version IS NULL OR TRIM(version)='' OR TRIM(version)='미확인'"
                " ORDER BY id DESC LIMIT ?" % (day_col, who_col, qty_col, table),
                (int(limit),)):
            out.append({"kind": kind, "id": rid, "on": day, "who": who,
                        "qty": int(qty or 0)})
    out.sort(key=lambda r: (r["on"] or "", r["id"]), reverse=True)
    return out[:limit]


def _remote_branch_stock(c):
    """지점별 현재 재고 = 재고 조정 합계 - 그 지점 불출 합계 (2026-08-03 지시).

    버전별 내역도 함께 준다(2026-08-04): 증평은 기존형과 VER.4 를 나눠 세고 있어
    합계만 보면 어느 물건이 남았는지 알 수 없다.
    """
    added = {row[0]: int(row[1] or 0) for row in c.execute(
        "SELECT branch,SUM(qty_delta) FROM remote_stock GROUP BY branch")}
    # 기초보유는 빼지 않는다 — 2026-07-29 기초 재고표의 지점 수량은 기사에게 이미
    # 지급하고 남은 net 값이라, 여기서 또 빼면 같은 물건을 두 번 차감하게 된다.
    issued = {row[0]: int(row[1] or 0) for row in c.execute(
        "SELECT branch,SUM(qty) FROM remote_issue"
        " WHERE status='불출완료' GROUP BY branch")}
    by_ver = {}
    for br, ver, delta in c.execute(
            "SELECT branch,COALESCE(NULLIF(version,''),'미확인'),SUM(qty_delta)"
            " FROM remote_stock GROUP BY branch,COALESCE(NULLIF(version,''),'미확인')"):
        by_ver.setdefault(br, {})[ver] = by_ver.get(br, {}).get(ver, 0) + int(delta or 0)
    for br, ver, qty in c.execute(
            "SELECT branch,COALESCE(NULLIF(version,''),'미확인'),SUM(qty)"
            " FROM remote_issue WHERE status='불출완료'"
            " GROUP BY branch,COALESCE(NULLIF(version,''),'미확인')"):
        if br in by_ver and ver in by_ver[br]:
            by_ver[br][ver] -= int(qty or 0)
    out = {}
    for br in REMOTE_BRANCH_ISSUERS:
        got, used = added.get(br, 0), issued.get(br, 0)
        out[br] = {"label": REMOTE_BRANCH_LABELS.get(br, br), "in": got,
                   "issued": used, "stock": got - used,
                   "versions": {k: v for k, v in sorted((by_ver.get(br) or {}).items())},
                   "tracked": br in added}          # 조정 기록이 있어야 재고 관리 대상
    return out


def remote_stock_adjust(branch, qty, mode="add", reason="", created_by="",
                        version="", moved_on=""):
    """지점 재고 등록/조정. mode='add'는 입고(±델타), 'set'은 실사값으로 맞춘다.

    version 을 주면 버전별 잔량까지 따라간다(증평 기존형/VER.4). 공장에서 현장으로
    바로 나간 택배 출고·샘플 송부처럼 담당자를 거치지 않는 출고도 음수 델타로 남긴다.
    """
    branch = str(branch or "").strip()
    if branch not in REMOTE_BRANCH_ISSUERS:
        raise ValueError("지점은 부산·시화·증평 중 하나여야 합니다")
    qty = int(qty)
    now = datetime.now().isoformat(timespec="seconds")
    with conn() as c:
        if mode == "set":
            current = _remote_branch_stock(c)[branch]["stock"]
            delta = qty - current
            if delta == 0:
                return current
            reason = reason or f"실사 {qty}개 맞춤"
        else:
            delta = qty
            if delta == 0:
                raise ValueError("수량이 0입니다")
            reason = reason or ("입고" if delta > 0 else "정정")
        after = _remote_branch_stock(c)[branch]["stock"] + delta
        if after < 0:
            raise ValueError(f"{REMOTE_BRANCH_LABELS[branch]} 재고가 음수({after})가 됩니다")
        c.execute(
            "INSERT INTO remote_stock(branch,qty_delta,reason,created_by,created_at,"
            "version,moved_on) VALUES(?,?,?,?,?,?,?)",
            (branch, delta, str(reason or ""), str(created_by or ""), now,
             _remote_version(version), str(moved_on or now[:10])[:10]))
        return after


def remote_open_balance(technician, qty, on="", note="", created_by="",
                        branch="", version=""):
    """개시 보유량 등록 — 2026-07-29 재고표에서 넘어온 '이미 갖고 있는 수량'.

    한도(3개)를 검사하지 않는다. 실제로 김필우·김준형 기사는 9개씩 들고 있었고,
    이를 거부하면 시스템이 현실과 어긋난 채로 남는다. 대신 상태를 '기초보유'로 남겨
    새 불출(remote_request)과 구분하고, remote_status 가 한도 초과자를 보고한다.
    지점 재고에서 빼지 않는다 — 기초 재고표가 이미 지급된 뒤의 숫자이기 때문이다.
    """
    technician = str(technician or "").strip()
    qty = int(qty or 0)
    if not technician:
        raise ValueError("AS 담당자 이름이 필요합니다")
    if qty < 1:
        raise ValueError("수량은 1개 이상이어야 합니다")
    now = datetime.now().isoformat(timespec="seconds")
    day = str(on or now[:10])[:10]
    with conn() as c:
        cur = c.execute(
            "INSERT INTO remote_issue(branch,issuer,technician,qty,status,"
            "requested_by,requested_at,note,issued_on,camp,version)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (str(branch or ""), REMOTE_BRANCH_ISSUERS.get(str(branch or ""), ""),
             technician, qty, "기초보유", str(created_by or ""), now,
             str(note or ""), day, "", str(version or "")))
        return cur.lastrowid


def remote_request(branch, technician, qty, requested_by, note="",
                   issued_on="", camp="", version="", issuer=""):
    """리모컨 불출을 즉시 기록한다(2026-08-03 지시 — 승인 단계 없음).

    한도: 보유 + 이번 수량이 담당자당 3개를 넘으면 거절한다.
    지점 재고를 등록해 둔 지점은 재고보다 많이 불출할 수 없다(재고 자동 차감).
    공지(2026-08-04): 불출 일자·투입 예정 캠프명도 함께 남긴다 — 류지영 매니저
    최종 취합의 원본이 이 표다.
    2026-08-06 지시: **버전**과 **지사 불출자 이름**을 받는다. 버전을 안 받으면
    지점 재고의 버전별 잔량이 '미확인' 으로 깎여 어느 물건이 나갔는지 알 수 없다
    (실제로 시화가 VER.3 2 · 미확인 -2 로 어긋나 있었다). 불출자는 지점 기본
    담당(부산=오종현·시화=안은숙·증평=류지영)이되, 대신 내준 사람이 있으면 그 이름을 남긴다.
    """
    branch = str(branch or "").strip()
    technician = str(technician or "").strip()
    qty = int(qty or 0)
    if branch not in REMOTE_BRANCH_ISSUERS:
        raise ValueError("불출 지점은 부산·시화·증평 중 하나여야 합니다")
    if not technician:
        raise ValueError("AS 담당자 이름이 필요합니다")
    if not 1 <= qty <= REMOTE_HOLD_LIMIT:
        raise ValueError(f"수량은 1~{REMOTE_HOLD_LIMIT}개여야 합니다")
    now = datetime.now().isoformat(timespec="seconds")
    with conn() as c:
        hold = _remote_holdings(c).get(technician) or {"holding": 0}
        if hold["holding"] + qty > REMOTE_HOLD_LIMIT:
            raise ValueError(
                f"{technician} 보유 {hold['holding']}개에 {qty}개를 더하면 "
                f"한도 {REMOTE_HOLD_LIMIT}개를 넘습니다")
        stock = _remote_branch_stock(c)[branch]
        if stock["tracked"] and stock["stock"] < qty:
            raise ValueError(
                f"{stock['label']} 재고 {stock['stock']}개 — {qty}개를 불출할 수 없습니다. "
                f"재고 등록(입고)을 먼저 하세요")
        day = str(issued_on or now[:10])[:10]
        cur = c.execute(
            "INSERT INTO remote_issue(branch,issuer,technician,qty,status,"
            "requested_by,requested_at,note,issued_on,camp,version)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (branch, str(issuer or "").strip() or REMOTE_BRANCH_ISSUERS[branch],
             technician, qty, "불출완료",
             str(requested_by or ""), now, str(note or ""), day,
             str(camp or "").strip(), _remote_version(version)))
        return cur.lastrowid


def remote_deliver(technician, project, camp, qty, delivered_on="", note="",
                   created_by="", kind="납품", version=""):
    """리모컨 납품 기록 — 어느 프로젝트/캠프에 몇 개가 들어갔는지 추적의 원본.

    kind 로 처리유형을 구분한다(2026-08-04 재고표 기준): 납품·사용·교체 모두
    기사 보유를 줄인다는 점은 같지만, 사람이 나중에 "왜 줄었나"를 물을 때 답이 다르다.
    """
    technician = str(technician or "").strip()
    qty = int(qty or 0)
    if not technician:
        raise ValueError("AS 담당자 이름이 필요합니다")
    if qty < 1:
        raise ValueError("수량은 1개 이상이어야 합니다")
    if not (str(project or "").strip() or str(camp or "").strip()):
        raise ValueError("프로젝트NO 또는 캠프명 중 하나는 필요합니다")
    now = datetime.now().isoformat(timespec="seconds")
    day = str(delivered_on or now[:10])[:10]
    with conn() as c:
        hold = _remote_holdings(c).get(technician) or {"holding": 0}
        if qty > hold["holding"]:
            raise ValueError(
                f"{technician} 보유 {hold['holding']}개보다 많은 {qty}개를 납품할 수 없습니다")
        cur = c.execute(
            "INSERT INTO remote_delivery(technician,project,camp,qty,delivered_on,"
            "note,created_by,created_at,kind,version) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (technician, str(project or "").strip().upper(), str(camp or "").strip(),
             qty, day, str(note or ""), str(created_by or ""), now,
             str(kind or "납품"), _remote_version(version)))
        return cur.lastrowid


# ── 리모컨 기록 고치기·지우기 (2026-08-06 지시) ────────────────────────────
# 왜 이 층이 따로 있나: 불출·납품·재고는 **수량이 서로 물려 있다.** 한 줄을 고치면
# 그 사람의 보유와 그 지점의 재고가 같이 움직인다. 그래서 "그냥 UPDATE" 로 두지 않고
#   ① 고치기 전 상태를 원장(remote_audit)에 남기고
#   ② 고친 뒤 숫자가 말이 되는지 다시 계산해 보고
#   ③ 말이 안 되면 막되, **강제(force)** 를 켜면 이유를 적고 통과시킨다
# 는 세 단계를 거친다. ③이 필요한 이유는 현실이 먼저이기 때문이다 — 실제로 기사들이
# 한도 3개를 넘겨 들고 있었고, 시스템이 거부하면 장부가 현실과 어긋난 채로 남는다.
REMOTE_TABLES = {
    "issue": ("remote_issue",
              ("branch", "issuer", "technician", "qty", "status",
               "issued_on", "camp", "version", "note")),
    "delivery": ("remote_delivery",
                 ("technician", "project", "camp", "qty", "delivered_on",
                  "kind", "version", "note")),
    "stock": ("remote_stock",
              ("branch", "qty_delta", "reason", "version", "moved_on")),
}


def _remote_table(kind):
    table = REMOTE_TABLES.get(str(kind or "").strip())
    if not table:
        raise ValueError("대상은 issue(불출)·delivery(납품)·stock(재고) 중 하나여야 합니다")
    return table


def _remote_row(c, table, rid):
    cur = c.execute(f"SELECT * FROM {table} WHERE id=?", (int(rid or 0),))
    row = cur.fetchone()
    if row is None:
        raise ValueError(f"{rid}번 기록이 없습니다 — 이미 지워졌을 수 있습니다")
    return dict(zip([d[0] for d in cur.description], row))


def _remote_problems(c):
    """지금 숫자에서 '말이 안 되는 것'을 이름→크기로 뽑는다."""
    out = {}
    for tech, h in _remote_holdings(c).items():
        if h["holding"] < 0:
            out[f"{tech} 보유가 음수"] = -h["holding"]
        elif h["holding"] > REMOTE_HOLD_LIMIT:
            out[f"{tech} 보유 한도({REMOTE_HOLD_LIMIT}개) 초과"] = h["holding"]
    for br, s in _remote_branch_stock(c).items():
        if s["tracked"] and s["stock"] < 0:
            out[f"{s['label']} 재고가 음수"] = -s["stock"]
    return out


def _remote_guard(before, after, force):
    """**이번 수정이 새로 만들거나 키운 문제**만 막는다.

    이미 어긋나 있는 장부(한도 초과 보유자가 실제로 있다)를 이유로 무관한 줄의
    수정까지 막으면 아무것도 고칠 수 없게 된다 — 그래서 전후를 비교한다.
    """
    worse = [f"{k}({v}개)" for k, v in after.items() if v > before.get(k, 0)]
    if worse and not force:
        raise ValueError("이 수정은 " + " · ".join(worse)
                         + " 를 만듭니다 — 그래도 저장하려면 '강제'를 켜고 사유를 적으세요")
    return worse


def _remote_audit(c, table, rid, action, before, after, reason, forced, actor):
    c.execute(
        "INSERT INTO remote_audit(table_name,row_id,action,before_json,after_json,"
        "reason,forced,actor,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
        (table, int(rid), action,
         json.dumps(before, ensure_ascii=False) if before else "",
         json.dumps(after, ensure_ascii=False) if after else "",
         str(reason or ""), 1 if forced else 0, str(actor or ""),
         datetime.now().isoformat(timespec="seconds")))


def _remote_clean(table, fields, allowed):
    """화면이 보낸 값을 표에 넣을 수 있는 형태로 다듬는다(허용된 열만)."""
    sets = {}
    for k, v in (fields or {}).items():
        if k not in allowed or v is None:
            continue
        if k in ("qty", "qty_delta"):
            sets[k] = int(v)
        elif k == "version":
            sets[k] = _remote_version(v)
        elif k in ("issued_on", "delivered_on", "moved_on"):
            sets[k] = str(v or "")[:10]
        elif k == "project":
            sets[k] = str(v or "").strip().upper()
        else:
            sets[k] = str(v or "").strip()
    if not sets:
        raise ValueError("바꿀 내용이 없습니다")
    if "branch" in sets and sets["branch"] not in REMOTE_BRANCH_ISSUERS:
        raise ValueError("지점은 부산·시화·증평 중 하나여야 합니다")
    if "qty" in sets and sets["qty"] < 1:
        raise ValueError("수량은 1개 이상이어야 합니다")
    if "qty_delta" in sets and sets["qty_delta"] == 0:
        raise ValueError("증감이 0이면 기록할 것이 없습니다")
    if table == "remote_issue" and "technician" in sets and not sets["technician"]:
        raise ValueError("AS 담당자 이름이 필요합니다")
    return sets


def remote_edit(kind, rid, fields, edited_by="", force=False, reason=""):
    """리모컨 기록 한 줄을 고친다. 되돌릴 근거를 remote_audit 에 먼저 남긴다."""
    table, allowed = _remote_table(kind)
    sets = _remote_clean(table, fields, allowed)
    if force and not str(reason or "").strip():
        raise ValueError("강제로 저장할 때는 사유를 적어야 합니다")
    with conn() as c:
        before = _remote_row(c, table, rid)
        problems0 = _remote_problems(c)
        c.execute(f"UPDATE {table} SET {','.join(k + '=?' for k in sets)} WHERE id=?",
                  (*sets.values(), int(rid)))
        forced_over = _remote_guard(problems0, _remote_problems(c), force)
        after = _remote_row(c, table, rid)
        _remote_audit(c, table, rid, "수정", before, after, reason,
                      bool(force and forced_over), edited_by)
    return {"row": after, "warnings": forced_over}


def remote_delete(kind, rid, deleted_by="", force=False, reason=""):
    """리모컨 기록 한 줄을 지운다 — 지운 내용 전체가 원장에 남아 복구할 수 있다."""
    table, _ = _remote_table(kind)
    if not str(reason or "").strip():
        raise ValueError("삭제할 때는 사유를 적어야 합니다")
    with conn() as c:
        before = _remote_row(c, table, rid)
        problems0 = _remote_problems(c)
        c.execute(f"DELETE FROM {table} WHERE id=?", (int(rid),))
        forced_over = _remote_guard(problems0, _remote_problems(c), force)
        _remote_audit(c, table, rid, "삭제", before, None, reason,
                      bool(force and forced_over), deleted_by)
    return {"row": before, "warnings": forced_over}


def remote_restore(audit_id, actor=""):
    """삭제를 되돌린다 — 원장의 before_json 을 그대로 다시 넣는다.

    삭제를 허용하려면 되돌리기가 있어야 한다. 없으면 사람이 무서워서 안 쓰거나,
    쓰고 나서 복구를 사람 손으로 해야 한다(그때 숫자가 또 어긋난다).
    """
    with conn() as c:
        row = c.execute(
            "SELECT table_name,row_id,action,before_json FROM remote_audit WHERE id=?",
            (int(audit_id or 0),)).fetchone()
        if not row:
            raise ValueError("복구할 기록을 찾지 못했습니다")
        table, row_id, action, before_json = row
        if action != "삭제" or not before_json:
            raise ValueError("삭제 기록만 복구할 수 있습니다")
        data = json.loads(before_json)
        if c.execute(f"SELECT 1 FROM {table} WHERE id=?", (int(row_id),)).fetchone():
            raise ValueError(f"{row_id}번은 이미 살아 있습니다")
        cols = [k for k in data if data[k] is not None]
        c.execute(f"INSERT INTO {table}({','.join(cols)})"
                  f" VALUES({','.join('?' for _ in cols)})",
                  tuple(data[k] for k in cols))
        _remote_audit(c, table, row_id, "복구", None, data, "삭제 되돌리기",
                      False, actor)
    return data


def remote_audit_list(limit=30):
    with conn() as c:
        return [dict(zip(("id", "table_name", "row_id", "action", "reason",
                          "forced", "actor", "created_at"), r))
                for r in c.execute(
                    "SELECT id,table_name,row_id,action,reason,forced,actor,created_at"
                    " FROM remote_audit ORDER BY id DESC LIMIT ?", (int(limit),))]


def remote_status(limit=60):
    """업무센터·대시보드·대표보고용 현황: 담당자별 보유와 최근 불출·납품 이력."""
    with conn() as c:
        holdings = _remote_holdings(c)
        branch_stock = _remote_branch_stock(c)
        issues = [dict(zip(("id", "branch", "issuer", "technician", "qty", "status",
                            "requested_at", "issued_on", "camp", "version"), row))
                  for row in c.execute(
                      "SELECT id,branch,issuer,technician,qty,status,requested_at,"
                      "issued_on,camp,version"
                      " FROM remote_issue ORDER BY id DESC LIMIT ?", (int(limit),))]
        deliveries = [dict(zip(("id", "technician", "project", "camp", "qty",
                                "delivered_on", "note", "kind", "version"), row))
                      for row in c.execute(
                          "SELECT id,technician,project,camp,qty,delivered_on,note,"
                          "kind,version"
                          " FROM remote_delivery ORDER BY id DESC LIMIT ?", (int(limit),))]
        moves = [dict(zip(("id", "branch", "qty_delta", "reason", "version",
                           "moved_on"), row))
                 for row in c.execute(
                     "SELECT id,branch,qty_delta,reason,version,moved_on"
                     " FROM remote_stock ORDER BY id DESC LIMIT ?", (int(limit),))]
        version_gaps = remote_version_gaps(c)
    # 공지 4항: 한도는 '기존 보유 포함'이다. 기초보유로 이미 넘긴 사람은 새 불출 대상이
    # 아니라 **회수·사용확인 대상**이므로 따로 뽑아 화면에 띄운다.
    over = {t: h["holding"] for t, h in holdings.items()
            if h["holding"] > REMOTE_HOLD_LIMIT}
    total_hold = sum(h["holding"] for h in holdings.values())
    total_stock = sum(b["stock"] for b in branch_stock.values())
    # 버전별 전체 합계 — "VER.3 이 지금 몇 개 남았나"에 한 줄로 답한다(2026-08-06 지시).
    # 개인 보유와 지점 재고는 서로 다른 물건이므로 나눠 세고 합을 함께 준다.
    ver_totals = {}
    for h in holdings.values():
        for k, n in (h.get("versions") or {}).items():
            ver_totals.setdefault(k, {"holding": 0, "stock": 0})["holding"] += n
    for b in branch_stock.values():
        for k, n in (b.get("versions") or {}).items():
            if n:
                ver_totals.setdefault(k, {"holding": 0, "stock": 0})["stock"] += n
    for k, v in ver_totals.items():
        v["all"] = v["holding"] + v["stock"]
    return {"limit": REMOTE_HOLD_LIMIT, "branches": REMOTE_BRANCH_ISSUERS,
            "branch_labels": REMOTE_BRANCH_LABELS, "branch_stock": branch_stock,
            "holdings": holdings, "issues": issues, "deliveries": deliveries,
            "moves": moves, "over_limit": over,
            # 화면이 버전 선택지를 스스로 만들지 않게 한다 — 서버와 목록이 갈리면
            # 'VER.3' 과 'ver3' 이 따로 세어진다(2026-08-06 지시).
            "versions": list(REMOTE_VERSIONS), "audit": remote_audit_list(20),
            # 버전별 합계와, 아직 버전을 안 적은 줄들(화면에서 바로 채운다)
            "version_totals": dict(sorted(ver_totals.items())),
            "version_gaps": version_gaps,
            "totals": {"holding": total_hold, "stock": total_stock,
                       "all": total_hold + total_stock},
            "rule": "AS 담당자당 최대 3개(기존 보유 포함) · 불출·납품 기록 · 지점 재고 자동 차감"}


def pending_work_completion_entries(today=None):
    """Return objectively completed AS/PM records still waiting for Excel.

    New Kakao/app records can remain in ``pending`` until the 11:00 or 15:00
    workbook slot. A real completion date with source evidence is already
    enough to persist the completion decision in SQLite. Repeated queue rows
    are collapsed by target sheet/row and the newest value per column wins.
    """
    today = today or datetime.now().date()
    if isinstance(today, datetime):
        today = today.date()
    if isinstance(today, str):
        try:
            today = datetime.strptime(today[:10], "%Y-%m-%d").date()
        except ValueError:
            today = datetime.now().date()

    with conn() as c:
        rows = c.execute(
            "SELECT id,sheet,cell,col,value,evidence FROM pending"
            " WHERE status='pending' AND sheet IN (?,?) ORDER BY id",
            ("02_돌발AS접수", "04_정기점검"),
        ).fetchall()

    grouped = {}
    for row_id, sheet, cell, col, value, evidence in rows:
        match = re.fullmatch(r"[A-Z]+(\d+)", str(cell or "").strip().upper())
        if not match:
            continue
        key = (sheet, int(match.group(1)))
        grouped.setdefault(key, {})[str(col or "").strip()] = {
            "id": row_id,
            "cell": str(cell or "").strip(),
            "value": "" if value is None else str(value).strip(),
            "evidence": "" if evidence is None else str(evidence).strip(),
        }

    specs = {
        "02_돌발AS접수": {
            "kind": "as", "project": "프로젝트NO", "completed": "작업완료일",
            "state": "진행상태", "conflicts": ("취소", "철회"), "status": "작업완료",
        },
        "04_정기점검": {
            "kind": "pm", "project": "프로젝트NO", "completed": "실제점검일",
            "state": "점검상태", "conflicts": ("AS전환", "점검불가", "취소", "철회"),
            "status": "완료",
        },
    }
    out = {}
    for (sheet, row_no), values in grouped.items():
        spec = specs[sheet]
        project_value = values.get(spec["project"], {})
        project = project_value.get("value", "")
        done = values.get(spec["completed"], {})
        state = values.get(spec["state"], {}).get("value", "")
        if not re.fullmatch(r"UJ\d{6,}", project, flags=re.IGNORECASE):
            continue
        try:
            completed = datetime.strptime(done.get("value", "")[:10], "%Y-%m-%d").date()
        except ValueError:
            continue
        if completed > today or any(word in state for word in spec["conflicts"]):
            continue
        if not done.get("evidence"):
            continue
        # A reused empty row may contain an older task's still-pending date.
        # Bind the date to the same source event (or an evidence string that
        # explicitly names this project) before treating it as objective proof.
        project_evidence = project_value.get("evidence", "")
        if (not project_evidence or
                (done["evidence"] != project_evidence
                 and project.lower() not in done["evidence"].lower())):
            continue
        entry = {
            "kind": spec["kind"],
            "record_id": project,
            "project": project,
            "status": spec["status"],
            "completed_on": completed.isoformat(),
            "basis": f"반영대기 {sheet}!{done['cell']} + {done['evidence']}",
        }
        out[(entry["kind"], project)] = entry
    return list(out.values())


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
    """사람이 열어 둔 동안은 쓰지 않는다 — 그리고 **기다리지도 않는다**(2026-08-03 지시).

    예전에는 유예 45분을 이 자리에서 자며 버텼다. 이제는 열림을 감지하는 즉시
    잠금 목록을 돌려주고, 호출자(apply_now)가 회차를 '연기'로 기록한 뒤 닫힘
    감시자(--resume-watch)를 띄운다 — 에이전트는 곧바로 다른 작업으로 전환하고,
    사람이 입력을 끝내고 저장(닫기)하면 감시자가 미룬 회차를 그 회차 이름으로
    마저 반영한다."""
    return human_editing()


def apply_now(force=False, now=None, resume_slot=None, ignore_input_window=False):
    """정해진 시각일 때만 엑셀에 쓴다. 실제 쓰기는 기존 ledger_writer 에 맡긴다.

    resume_slot: 사람이 관리대장을 열어 두어 '연기'된 회차의 재개(2026-08-03).
    임의 시각 반영이 아니라 **그 회차의 지연 실행**이라 회차 이름을 그대로 쓴다.
    """
    now = now or datetime.now()
    from operation_window import is_input_window, input_window_label
    if is_input_window(now) and not ignore_input_window:
        return {"상태": "보류", "사유": f"입력 보호시간({input_window_label()})"}
    if is_input_window(now) and ignore_input_window:
        # ★ 사용자가 "지금 반영"을 명시했을 때만 보호시간을 건너뛴다(2026-08-05 지시).
        #   보호시간의 목적은 **사람이 열어 둔 원장을 덮어쓰지 않는 것**이므로,
        #   엑셀 잠금 파일(~$…)이 있으면 지시가 있어도 하지 않는다 — 그건 실제 충돌이다.
        import glob as _g
        try:
            from ecount_reconcile import load_config as _lc, resolve_master as _rm
            _dir = os.path.dirname(_rm(_lc()["reconcile"]["master_xlsx"]))
            if _g.glob(os.path.join(_dir, "~$*.xlsx")):
                return {"상태": "보류", "사유": "관리대장이 열려 있음 — 닫은 뒤 다시"}
        except Exception:
            pass

    intake_json()                                    # 도구들이 넣어 둔 것 먼저 흡수
    p, by, done = counts()
    slot_name = eligible_slot(now, done, force)
    if not slot_name and resume_slot and resume_slot not in set(done):
        slot_name = resume_slot
    if not slot_name:
        nxt = next_window(now)
        current = slot_of(now)
        why = "이미 처리한 회차" if current and current in set(done) else "반영 시각이 아님"
        return {"상태": "대기", "사유": f"{why} — 다음 {nxt:%m-%d %H:%M}",
                "대기": p}
    # ★ 쓰기 직전 관문 — 사람이 관리대장을 열어 두었으면 즉시 '연기'한다(2026-08-03 지시).
    #   구조 갱신(scheduled_workbook_maintenance)도 vN+1 을 만들므로 같이 막아야 한다.
    #   force 로도 뚫지 못한다: 강제의 용도는 '시각 밖 반영'이지 '사람을 밀어내기'가 아니다.
    locks = _wait_editing_clear(now, slot_name)
    if locks:
        who = locks[0].get("소유자") or "?"
        watch = defer_apply(slot_name)
        return {"상태": "연기", "회차": slot_name,
                "사유": f"사람이 관리대장 편집 중({who}, {locks[0]['분']}분째) — "
                        f"다른 작업으로 전환, 저장(닫힘) 감지 후 자동 반영", "대기": p,
                "감시자": watch.get("watcher_pid")}
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


# ── 엑셀 열림 감지 → 연기 → 닫힘(저장) 후 자동 재개 (2026-08-03 지시, 상시) ──
DEFER_FLAG = os.path.join(DB_DIR, "apply_deferred.json")


def _defer_state():
    try:
        return json.load(open(DEFER_FLAG, encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _defer_write(state):
    os.makedirs(DB_DIR, exist_ok=True)
    tmp = DEFER_FLAG + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False)
    os.replace(tmp, DEFER_FLAG)


def defer_apply(slot_name, spawn=True):
    """미룬 회차를 기록하고 닫힘 감시자를 **하나만** 띄운다.

    감시자는 별도 프로세스라 스케줄러·에이전트는 바로 다른 작업으로 넘어간다.
    감시자가 죽어도 마커가 남아 30분 워치독(resume_check)이 다시 잇는다.
    """
    state = _defer_state()
    slots = sorted(set(state.get("slots") or []) | {slot_name})
    pid = state.get("watcher_pid")
    if spawn and not (pid and _pid_alive(pid)):
        try:
            proc = subprocess.Popen(
                [sys.executable, os.path.abspath(__file__), "--resume-watch"],
                cwd=ROOT, creationflags=0x00000008 if os.name == "nt" else 0,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL)
            pid = proc.pid
        except Exception:
            pid = None
    state = {"slots": slots, "watcher_pid": pid,
             "since": state.get("since") or datetime.now().isoformat(timespec="seconds")}
    _defer_write(state)
    return state


def resume_deferred(now=None):
    """닫힘이 확인된 뒤 미룬 회차를 그 회차 이름으로 반영한다."""
    now = now or datetime.now()
    state = _defer_state()
    slots = [s for s in (state.get("slots") or [])]
    if not slots:
        try:
            os.unlink(DEFER_FLAG)
        except OSError:
            pass
        return {"상태": "없음", "사유": "연기된 회차 없음"}
    with apply_lock():
        result = apply_now(now=now, resume_slot=slots[-1])
    if result.get("상태") in ("반영", "없음", "대기"):
        # 큐는 하나라 최신 회차 반영이 전부를 포함한다 — 나머지 연기 회차도 완료로 남긴다.
        stamp = datetime.now().isoformat(timespec="seconds")
        with conn() as c:
            for s in slots:
                c.execute(
                    "INSERT OR IGNORE INTO batch(slot,started,finished,cells,ok,note,forced)"
                    " VALUES(?,?,?,?,1,?,0)",
                    (s, stamp, stamp, 0, f"사람 편집 연기 — {slots[-1]} 재개분에 포함"))
        try:
            os.unlink(DEFER_FLAG)
        except OSError:
            pass
    return result


def resume_watch(max_hours=10):
    """엑셀 닫힘(저장)을 기다렸다가 재개한다 — defer_apply 가 띄우는 감시자 본체.

    입력 보호시간(08:00~09:30)에는 재개하지 않고 계속 기다린다. 시한이 지나면
    조용히 물러난다 — 큐는 그대로라 다음 11:00/15:00 회차가 마저 처리한다.
    """
    from operation_window import is_input_window
    deadline = time.time() + max_hours * 3600
    while time.time() < deadline:
        state = _defer_state()
        if not state.get("slots"):
            return {"상태": "없음"}
        if not human_editing() and not is_input_window(datetime.now()):
            result = resume_deferred()
            if result.get("상태") != "연기":       # 그 사이 다시 열렸으면 계속 기다린다
                return result
        time.sleep(LOCK_POLL_SEC)
    return {"상태": "시한초과", "사유": "다음 정규 회차가 처리"}


def resume_check():
    """워치독(30분)용 안전망: 마커가 있는데 감시자가 죽었으면 즉시 재개하거나 다시 띄운다."""
    state = _defer_state()
    if not state.get("slots"):
        return {"상태": "없음"}
    pid = state.get("watcher_pid")
    if pid and _pid_alive(pid):
        return {"상태": "감시중", "감시자": pid}
    from operation_window import is_input_window
    if not human_editing() and not is_input_window(datetime.now()):
        return resume_deferred()
    return defer_apply(state["slots"][-1])          # 아직 열려 있음 — 감시자만 재기동


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
            r = apply_now(force="--force" in sys.argv,
                          ignore_input_window="--now" in sys.argv)
        print(" · ".join(f"{k} {v}" for k, v in r.items()))
        if r.get("상태") == "실패":
            sys.exit(1)
        return
    if "--resume-watch" in sys.argv:
        # 엑셀 닫힘(저장) 감시자 — defer_apply 가 하나만 띄운다(2026-08-03).
        r = resume_watch()
        print(" · ".join(f"{k} {v}" for k, v in r.items()))
        return
    if "--resume-check" in sys.argv:
        # 워치독 30분 안전망 — 감시자가 죽어도 연기 회차를 잃지 않는다.
        r = resume_check()
        print(" · ".join(f"{k} {v}" for k, v in (r or {}).items()))
        return
    d = status()
    print(f"반영 대기 {d['대기']}건 · 다음 반영 {d['다음반영']} (약 {d['남은분']}분 뒤)")
    if d["출처별"]:
        print("  출처:", ", ".join(f"{k} {v}건" for k, v in d["출처별"].items()))
    if d["인수인계대기"]:
        print(f"  19시트 인수인계 예약: {d['인수인계대기']}건")
    if d["밀린회차"]:
        print("  ★ 밀린 회차:", ", ".join(d["밀린회차"]))
    deferred = _defer_state()
    if deferred.get("slots"):
        print("  ★ 엑셀 열림으로 연기된 회차:", ", ".join(deferred["slots"]),
              "— 저장(닫힘) 감지 후 자동 반영")
    print(f"  반영 시각: 매일 {' · '.join(d['반영시각'])} (하루 두 번)")


if __name__ == "__main__":
    main()
