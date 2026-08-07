# -*- coding: utf-8 -*-
"""
datalake.py — 전 자료의 **영구 보관소**. 1단계: 스키마 + 자산 흡수 + 증분 주사 (2026-08-08)

사용자 지시(2026-08-07): "모든 데이터는 Db화 해서 별도 보관하고 앞으로 들어오는 모든
데이터 포함 변경 및 로그 기록까지 같이 정리해. 그리고 인덱스 필터링 가능한 구조,
플로우 차트 연계 가능한 구조로 코딩하는 알고리즘 구성해."

설계 전문은 `ecount/DATALAKE.md`. 이 파일은 그 **1단계**다 —
표를 세우고, 원본 파일을 증분으로 흡수하고, 모든 일을 `event` 에 남긴다.
(`--find`/FTS5 는 2단계, `record` 는 4단계, `link`/mermaid 는 5단계)

  python datalake.py --scan            # Z: 를 증분으로 훑어 asset 갱신
  python datalake.py --scan --rescan   # 캐시 무시하고 전부 다시(느리다)
  python datalake.py --status          # 지금 무엇이 얼마나 들어 있나
  python datalake.py --log area=collect since=2026-08-07 [--fail-only]

★ 왜 큐 DB(`db/ledger_queue.db`)에 안 넣나
  큐 DB 는 11:00·15:00 엑셀 반영이 쓰기 트랜잭션으로 잠근다. 색인 흡수는 수만 행을
  쓴다. SQLite 는 **파일 단위** 쓰기 잠금이라, 한 파일에 두면 반영이 색인에 막혀
  회차를 통째로 놓친다. 수명도 다르다 — 큐는 "곧 엑셀로 나갈 것"만 담는 임시 통로고
  여기는 영구 보관소다. 이어 볼 일이 있으면 `ATTACH DATABASE` 로 충분하다.

★ 지우지 않는다
  파일이 사라져도 행을 지우지 않고 `gone_at` 만 찍는다. 지우면 **'있었다'는 사실**을
  잃는다. 2026-08-07 밴드 사고가 가르쳐 준 것이 이것이다 — 실패는 삭제의 증거가 아니다.

★ 밀림은 `mtime` 이 아니라 `biz_date` 로 잰다
  오늘 받은 8/4 자 자료는 '오늘 자료'가 아니다. mtime 으로 재면 매일 받기만 해도
  안 밀린 것처럼 보인다. `erp_grab` 이 정확히 여기서 헷갈렸다.
"""
import argparse
import hashlib
import json
import os
import sqlite3
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _shared(*parts):
    """공용 상태 경로. 워크트리에서도 **본체 하나**를 본다.

    DB 는 링크하면 안 된다 — `-wal` 사이드카가 갈리면 파일이 깨진다
    (CLAUDE.md 워크트리 규칙). 그래서 코드가 본체 경로를 직접 집는다.
    """
    try:
        from worktree_state import shared
        return shared(*parts)
    except Exception:
        return os.path.join(ROOT, *parts)


def db_path():
    return _shared("db", "datalake.db")


def who():
    """`claude:<sid>` — 창이 여러 개인 것이 기본이라 세션까지 적는다."""
    try:
        from ai_claim import session_id
        sid = session_id()
    except Exception:
        sid = "manual"
    tag = "codex" if os.environ.get("CODEX_SESSION_ID") else "claude"
    return f"{tag}:{sid}"


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


COMMIT_EVERY = 300      # 주사 중 커밋 간격(건). 길게 잡으면 남을 잠그고, 짧으면 느리다

SCHEMA = """
CREATE TABLE IF NOT EXISTS asset(
  id        INTEGER PRIMARY KEY AUTOINCREMENT,
  path      TEXT NOT NULL UNIQUE,
  kind      TEXT NOT NULL,
  bucket    TEXT DEFAULT '',
  mtime     REAL NOT NULL,
  size      INTEGER NOT NULL,
  sha1      TEXT,
  biz_date  TEXT,
  first_seen TEXT NOT NULL,
  last_seen  TEXT NOT NULL,
  gone_at    TEXT
);
CREATE INDEX IF NOT EXISTS ix_asset_kind_date ON asset(kind, biz_date DESC);
CREATE INDEX IF NOT EXISTS ix_asset_seen      ON asset(last_seen DESC);

CREATE TABLE IF NOT EXISTS asset_rev(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL REFERENCES asset(id),
  at TEXT NOT NULL, sha1 TEXT, size INTEGER, mtime REAL,
  note TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_rev_asset ON asset_rev(asset_id, at DESC);

CREATE TABLE IF NOT EXISTS record(
  id       INTEGER PRIMARY KEY AUTOINCREMENT,
  kind     TEXT NOT NULL,
  natural_key TEXT NOT NULL,
  asset_id INTEGER REFERENCES asset(id),
  biz_date TEXT,
  party    TEXT DEFAULT '',
  amount   INTEGER,
  status   TEXT DEFAULT '',
  payload  TEXT NOT NULL,
  hash     TEXT NOT NULL,
  created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
  UNIQUE(kind, natural_key)
);
CREATE INDEX IF NOT EXISTS ix_rec_date  ON record(biz_date DESC);
CREATE INDEX IF NOT EXISTS ix_rec_party ON record(party, biz_date DESC);
CREATE INDEX IF NOT EXISTS ix_rec_kind  ON record(kind, biz_date DESC);

CREATE TABLE IF NOT EXISTS record_rev(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  record_id INTEGER NOT NULL REFERENCES record(id),
  at TEXT NOT NULL, who TEXT NOT NULL,
  field TEXT NOT NULL, old TEXT, new TEXT,
  why TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_rrev_rec ON record_rev(record_id, at DESC);

CREATE TABLE IF NOT EXISTS event(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at   TEXT NOT NULL,
  who  TEXT NOT NULL,
  area TEXT NOT NULL,
  action TEXT NOT NULL,
  ok   INTEGER NOT NULL DEFAULT 1,
  ref_kind TEXT, ref_id INTEGER,
  detail TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS ix_ev_at   ON event(at DESC);
CREATE INDEX IF NOT EXISTS ix_ev_area ON event(area, at DESC);

CREATE TABLE IF NOT EXISTS link(
  src INTEGER NOT NULL REFERENCES record(id),
  dst INTEGER NOT NULL REFERENCES record(id),
  rel TEXT NOT NULL,
  conf REAL DEFAULT 1.0,
  by   TEXT DEFAULT '', at TEXT NOT NULL,
  PRIMARY KEY(src, dst, rel)
);
CREATE INDEX IF NOT EXISTS ix_link_dst ON link(dst);

-- ★ 로그는 고칠 수도 지울 수도 없다. 고쳐질 수 있으면 근거가 아니다.
--   "8/5 돌발AS 가 왜 1건이었나"를 되짚을 때 믿을 것이 이 표뿐이다.
CREATE TRIGGER IF NOT EXISTS ev_no_update BEFORE UPDATE ON event
BEGIN SELECT RAISE(ABORT,'event 는 고칠 수 없다'); END;
CREATE TRIGGER IF NOT EXISTS ev_no_delete BEFORE DELETE ON event
BEGIN SELECT RAISE(ABORT,'event 는 지울 수 없다'); END;
"""


def connect(path=None):
    """DB 를 열고(없으면 만들고) 스키마를 맞춘다.

    WAL 이라 **읽는 세션이 쓰는 세션을 막지 않는다** — 수집 세션과 코딩 세션이
    같이 떠 있는 것이 이 프로젝트의 기본이라 그것이 조건이다.
    """
    p = path or db_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    # ★ 넉넉히 기다린다. WAL 이라도 **쓰기는 한 번에 하나**다. Z: 주사가 도는 동안
    #   다른 도구가 로그 한 줄을 못 써서 죽는 일이 있었다(2026-08-08 실측).
    con = sqlite3.connect(p, timeout=120)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript(SCHEMA)
    con.commit()
    return con


def log(con, area, action, ok=True, ref_kind=None, ref_id=None, detail=None, actor=None):
    """무슨 일이 있었는지 한 줄. **실패도 반드시 남긴다**(ok=0)."""
    con.execute(
        "INSERT INTO event(at,who,area,action,ok,ref_kind,ref_id,detail)"
        " VALUES(?,?,?,?,?,?,?,?)",
        (now(), actor or who(), area, action, 1 if ok else 0, ref_kind, ref_id,
         json.dumps(detail, ensure_ascii=False) if detail is not None else ""))


def sha1_of(path, chunk=1 << 20):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(chunk), b""):
            h.update(blk)
    return h.hexdigest()


def ingest_asset(con, path, kind, bucket="", biz_date=None, note="", st=None,
                 want_sha1=True):
    """원본 파일 한 건을 흡수한다 → (asset_id, 상태)

    상태: `새것` | `바뀜` | `그대로` — 부르는 쪽이 세어 보고할 수 있게.

    ★ 증분의 핵심은 여기다. (mtime, size) 가 DB 와 같으면 **파일을 열지 않는다.**
      sha1 을 매번 계산하면 Z: 는 SMB 라 몇 초가 몇 시간이 된다.
    """
    st = st or os.stat(path)
    mtime, size = float(st.st_mtime), int(st.st_size)
    row = con.execute("SELECT * FROM asset WHERE path=?", (path,)).fetchone()
    ts = now()

    if row is None:
        digest = sha1_of(path) if want_sha1 else None
        cur = con.execute(
            "INSERT INTO asset(path,kind,bucket,mtime,size,sha1,biz_date,"
            "first_seen,last_seen) VALUES(?,?,?,?,?,?,?,?,?)",
            (path, kind, bucket, mtime, size, digest, biz_date, ts, ts))
        aid = cur.lastrowid
        con.execute("INSERT INTO asset_rev(asset_id,at,sha1,size,mtime,note)"
                    " VALUES(?,?,?,?,?,?)", (aid, ts, digest, size, mtime, note or "처음"))
        return aid, "새것"

    aid = row["id"]
    same_stat = (abs(float(row["mtime"]) - mtime) < 1e-6 and int(row["size"]) == size)
    if same_stat and not row["gone_at"]:
        # 그대로다 — 마지막 본 시각만 갱신한다. 파일은 열지 않는다.
        con.execute("UPDATE asset SET last_seen=?, kind=COALESCE(NULLIF(?,''),kind),"
                    " biz_date=COALESCE(?,biz_date) WHERE id=?",
                    (ts, kind, biz_date, aid))
        return aid, "그대로"

    digest = sha1_of(path) if want_sha1 else None
    changed = (digest != row["sha1"]) if want_sha1 else True
    con.execute("UPDATE asset SET kind=COALESCE(NULLIF(?,''),kind), bucket=?,"
                " mtime=?, size=?, sha1=?, biz_date=COALESCE(?,biz_date),"
                " last_seen=?, gone_at=NULL WHERE id=?",
                (kind, bucket or row["bucket"], mtime, size, digest, biz_date, ts, aid))
    if changed:
        # ★ 내용이 실제로 달라졌을 때만 이력을 쌓는다. 같은 파일을 백 번 봐도
        #   한 행도 안 늘어야 이력이 읽을 만한 것으로 남는다.
        con.execute("INSERT INTO asset_rev(asset_id,at,sha1,size,mtime,note)"
                    " VALUES(?,?,?,?,?,?)", (aid, ts, digest, size, mtime, note))
        return aid, "바뀜"
    return aid, "그대로"


def mark_gone(con, seen_paths, roots):
    """이번 주사에서 못 본 경로에 묘비를 세운다 → 세운 수.

    ★ 지우지 않는다. 그리고 **이번에 실제로 훑은 뿌리 아래 것만** 본다 —
      한 폴더만 훑고 전부 사라졌다고 적으면 그것이 더 큰 사고다.
    """
    n = 0
    rows = con.execute("SELECT id, path FROM asset WHERE gone_at IS NULL").fetchall()
    for r in rows:
        if r["path"] in seen_paths:
            continue
        if not any(r["path"].startswith(x) for x in roots):
            continue                      # 이번에 안 훑은 영역 — 판단하지 않는다
        con.execute("UPDATE asset SET gone_at=? WHERE id=?", (now(), r["id"]))
        n += 1
    return n


def _walk(top, SI):
    """`os.scandir` 로 훑는다 → (폴더, 이름, 확장자, 전체경로, stat).

    ★ **`os.stat()` 을 따로 부르지 않는 것이 핵심이다** (2026-08-08 실측).
      `os.walk` 로 이름만 받아 파일마다 `os.stat()` 을 부르면 Z:(SMB)에서 왕복이
      파일 수만큼 생긴다 — 밴드 폴더 한 곳(4만여 장)에서 5분이 넘어도 안 끝났다.
      `scandir` 이 주는 `entry.stat()` 은 윈도에서 **디렉터리 열거에 딸려 온 값**이라
      추가 왕복이 없다. 같은 폴더가 25초에 3만 개씩 훑힌다.
    """
    stack = [top]
    while stack:
        d = stack.pop()
        try:
            it = list(os.scandir(d))
        except OSError:
            continue
        for e in it:
            try:
                if e.is_dir(follow_symlinks=False):
                    if e.name not in SI.SKIP_DIRS:
                        stack.append(e.path)
                    continue
                ext = os.path.splitext(e.name)[1].lower()
                if ext in SI.SKIP_EXT or e.name.startswith("~$"):
                    continue
                yield d, e.name, ext, e.path, e.stat()
            except OSError:
                yield d, e.name, "", e.path, None


def scan(con, rescan=False, limit_roots=None, quiet=False):
    """Z: 원본 폴더를 증분으로 훑어 `asset` 을 맞춘다 → 집계 dict.

    `source_index.scan()` 과 **같은 자리를 같은 규칙으로** 본다(그쪽이 아직 정본이고,
    둘이 같은 답을 내는지 한동안 나란히 확인한 뒤에 물러난다 — DATALAKE.md).
    """
    import source_dirs as S
    import source_index as SI

    roots = []
    for attr in ("ERP_DIR", "BAND_DIR", "COUPANG_DIR", "KAKAO_DIR", "RECEIPT_DIR",
                 "DOC_DIR", "ORIGIN_ROOT"):
        p = getattr(S, attr, None)
        if p and os.path.isdir(p):
            roots.append(p)
    if limit_roots:
        roots = [r for r in roots if any(x in r for x in limit_roots)]
    tops = []
    for r in sorted(set(roots), key=len):        # 상위가 하위를 품으면 한 번만 훑는다
        if not any(r.startswith(x + os.sep) for x in tops):
            tops.append(r)

    try:
        from inbox_scan import classify
    except Exception:
        classify = None

    t0 = time.time()
    tally = {"본것": 0, "새것": 0, "바뀜": 0, "그대로": 0, "건너뜀": 0, "오류": 0}
    seen = set()
    for top in tops:
        for dirpath, fn, ext, p, st in _walk(top, SI):
            if SI.is_private(p, fn):     # 통화 메모 등 — 보관소에도 남기지 않는다
                tally["건너뜀"] += 1
                continue
            if st is None:
                tally["오류"] += 1
                continue
            seen.add(p)
            tally["본것"] += 1
            kind = SI.folder_kind(p) or "기타"
            try:
                prev = con.execute(
                    "SELECT mtime,size FROM asset WHERE path=?", (p,)).fetchone()
                fresh = rescan or prev is None or \
                    abs(float(prev["mtime"]) - st.st_mtime) > 1e-6 or \
                    int(prev["size"]) != int(st.st_size)
                # 내용 판별은 **새것·바뀐 것에만**. ERP 엑셀만 이름이 무작위다.
                if fresh and classify and ext == ".xlsx" and kind in ("ERP", "기타"):
                    try:
                        k2 = classify(p)
                        if k2 and k2 != "unknown":
                            kind = f"ERP:{k2}"
                    except Exception:
                        pass
                # ★ **처음 보는 파일에는 sha1 을 재지 않는다** (2026-08-08 실측).
                #   sha1 은 파일을 통째로 읽는 일이고 Z: 는 SMB 다. 첫 주사에서
                #   5만 개를 전부 읽으면 몇 시간이 걸린다(실측: 8분에 300건도 못 갔다).
                #   그리고 처음 보는 파일에는 **견줄 옛 지문이 없다** — 지금 재 봐야
                #   아무 판정에도 안 쓰인다. 지문이 필요한 순간은 '이미 아는 파일의
                #   mtime/size 가 달라졌을 때 내용이 진짜 바뀌었나' 하나뿐이다.
                #   나중에 채우려면 `--fill-sha1`.
                _aid, state = ingest_asset(
                    con, p, kind, bucket=SI.folder_kind(p),
                    biz_date=SI.guess_date(fn, st.st_mtime),
                    st=st, want_sha1=(fresh and prev is not None))
                tally[state] = tally.get(state, 0) + 1
            except Exception as e:
                tally["오류"] += 1
                log(con, "intake", "datalake.scan.file", ok=False,
                    detail={"path": p, "왜": str(e)[:200]})
            # ★ 자주 끊어 커밋한다. 뿌리 하나를 다 훑고 커밋하면 트랜잭션이
            #   수십 분 열려 있고, 그동안 **다른 도구가 로그 한 줄을 못 써서
            #   죽는다**(2026-08-08 실측: collect_all 이 'database is locked').
            #   중간에 끊겨도 여기까지는 남는다는 뜻이기도 하다.
            if tally["본것"] % COMMIT_EVERY == 0:
                con.commit()
        con.commit()

    tally["묘비"] = mark_gone(con, seen, tops)
    tally["초"] = round(time.time() - t0, 1)
    log(con, "intake", "datalake.scan", ok=tally["오류"] == 0, detail=tally)
    con.commit()
    if not quiet:
        print("자산 주사: 본것 {본것} · 새것 {새것} · 바뀜 {바뀜} · 그대로 {그대로}"
              " · 묘비 {묘비} · 오류 {오류} ({초}초)".format(**tally))
    return tally


def fill_sha1(con, limit=2000, quiet=False):
    """지문이 비어 있는 자산에 sha1 을 채운다 → 채운 수.

    첫 주사는 일부러 지문을 안 잰다(SMB 라 몇 시간이 걸린다). 지문이 실제로 쓰이는
    자리는 '아는 파일이 바뀌었나'뿐이라 **나중에 한가할 때 채워도 늦지 않다.**
    한 번에 다 하지 않고 상한을 두어 여러 회차에 나눠 채운다.
    """
    rows = con.execute("SELECT id, path FROM asset WHERE sha1 IS NULL"
                       " AND gone_at IS NULL LIMIT ?", (int(limit),)).fetchall()
    n = 0
    for r in rows:
        try:
            d = sha1_of(r["path"])
        except OSError:
            continue
        con.execute("UPDATE asset SET sha1=? WHERE id=?", (d, r["id"]))
        con.execute("UPDATE asset_rev SET sha1=? WHERE asset_id=? AND sha1 IS NULL",
                    (d, r["id"]))
        n += 1
        if n % COMMIT_EVERY == 0:
            con.commit()
    con.commit()
    log(con, "intake", "datalake.fill_sha1", detail={"채움": n, "남음": len(rows) - n})
    con.commit()
    if not quiet:
        남 = con.execute("SELECT COUNT(*) c FROM asset WHERE sha1 IS NULL"
                         " AND gone_at IS NULL").fetchone()["c"]
        print(f"지문 채움: {n}건 · 아직 없는 것 {남}건")
    return n


def status(con):
    """지금 무엇이 얼마나 들어 있나 → 사람이 읽는 요약."""
    out = {}
    out["자산"] = con.execute("SELECT COUNT(*) c FROM asset WHERE gone_at IS NULL").fetchone()["c"]
    out["사라짐"] = con.execute("SELECT COUNT(*) c FROM asset WHERE gone_at IS NOT NULL").fetchone()["c"]
    out["이력"] = con.execute("SELECT COUNT(*) c FROM asset_rev").fetchone()["c"]
    out["레코드"] = con.execute("SELECT COUNT(*) c FROM record").fetchone()["c"]
    out["로그"] = con.execute("SELECT COUNT(*) c FROM event").fetchone()["c"]
    out["종류"] = [(r["kind"], r["c"], r["last"]) for r in con.execute(
        "SELECT kind, COUNT(*) c, MAX(biz_date) last FROM asset"
        " WHERE gone_at IS NULL GROUP BY kind ORDER BY c DESC LIMIT 20")]
    return out


def events(con, area=None, since=None, action=None, fail_only=False, limit=40):
    q = "SELECT * FROM event WHERE 1=1"
    args = []
    if area:
        q += " AND area=?"; args.append(area)
    if action:
        q += " AND action LIKE ?"; args.append(f"%{action}%")
    if since:
        q += " AND at>=?"; args.append(since)
    if fail_only:
        q += " AND ok=0"
    q += " ORDER BY at DESC LIMIT ?"
    args.append(int(limit))
    return con.execute(q, args).fetchall()


def _kv(pairs):
    out = {}
    for p in pairs or []:
        if "=" in p:
            k, v = p.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="전 자료 영구 보관소 — 스키마·흡수·증분 주사")
    ap.add_argument("--scan", action="store_true", help="Z: 를 증분으로 훑어 asset 갱신")
    ap.add_argument("--rescan", action="store_true", help="캐시 무시하고 전부 다시(느리다)")
    ap.add_argument("--only", nargs="*", help="특정 뿌리만 (예: ERP 밴드)")
    ap.add_argument("--fill-sha1", type=int, nargs="?", const=2000, metavar="N",
                    help="지문이 빈 자산에 sha1 을 채운다(첫 주사는 일부러 안 잰다)")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--log", nargs="*", metavar="키=값", help="area=… since=… action=…")
    ap.add_argument("--fail-only", action="store_true")
    ap.add_argument("--db", help="DB 경로(합성검증용)")
    a = ap.parse_args(argv)

    con = connect(a.db)
    try:
        if a.scan:
            scan(con, rescan=a.rescan, limit_roots=a.only)
        if a.fill_sha1:
            fill_sha1(con, limit=a.fill_sha1)
        if a.log is not None:
            f = _kv(a.log)
            rows = events(con, area=f.get("area"), since=f.get("since"),
                          action=f.get("action"), fail_only=a.fail_only)
            for r in rows:
                mark = "✓" if r["ok"] else "✗"
                print(f"  {mark} {r['at']}  [{r['area']}] {r['action']}  {r['who']}"
                      f"  {(r['detail'] or '')[:90]}")
            if not rows:
                print("  (해당하는 로그 없음)")
        if a.status or not (a.scan or a.fill_sha1 or a.log is not None):
            s = status(con)
            print(f"보관소: {db_path()}")
            print(f"  자산 {s['자산']}건 (사라짐 {s['사라짐']}) · 변경이력 {s['이력']}"
                  f" · 레코드 {s['레코드']} · 로그 {s['로그']}")
            for kind, c, last in s["종류"]:
                print(f"    {kind:<22} {c:>6}건  최신 {last or '-'}")
        con.commit()
    finally:
        con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
