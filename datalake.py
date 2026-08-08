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
import re
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

-- ★ 어느 파일을 **어떤 모습일 때** 이미 뜯어 봤나 (worksplit #24).
--   ERP 엑셀은 100개가 넘고 하나 여는 데 수 초다. 매 회차 전부 다시 열면 회차가
--   못 끝난다. 그렇다고 '한 번 봤으면 끝'도 안 된다 — 같은 이름으로 **다시 받은**
--   파일이 내용만 다를 때 그것을 놓친다. 그래서 기억하는 것은 파일이 아니라
--   **그때의 모습(크기·수정시각)**이다. 모습이 달라지면 다시 뜯는다.
CREATE TABLE IF NOT EXISTS parsed(
  asset_id INTEGER PRIMARY KEY REFERENCES asset(id),
  fingerprint TEXT NOT NULL,
  rows INTEGER DEFAULT 0,
  at   TEXT NOT NULL
);

-- ★ 로그는 고칠 수도 지울 수도 없다. 고쳐질 수 있으면 근거가 아니다.
--   "8/5 돌발AS 가 왜 1건이었나"를 되짚을 때 믿을 것이 이 표뿐이다.
CREATE TRIGGER IF NOT EXISTS ev_no_update BEFORE UPDATE ON event
BEGIN SELECT RAISE(ABORT,'event 는 고칠 수 없다'); END;
CREATE TRIGGER IF NOT EXISTS ev_no_delete BEFORE DELETE ON event
BEGIN SELECT RAISE(ABORT,'event 는 지울 수 없다'); END;
"""


# FTS5 는 없는 빌드가 있다. 없다고 보관소 전체가 안 열리면 안 되므로 따로 세운다.
FTS_SCHEMA = """
CREATE VIRTUAL TABLE IF NOT EXISTS record_fts USING fts5(
  natural_key, party, body, content=''
);
CREATE TRIGGER IF NOT EXISTS rec_fts_ins AFTER INSERT ON record BEGIN
  INSERT INTO record_fts(rowid, natural_key, party, body)
  VALUES (new.id, new.natural_key, new.party, new.payload);
END;
CREATE TRIGGER IF NOT EXISTS rec_fts_del AFTER DELETE ON record BEGIN
  INSERT INTO record_fts(record_fts, rowid, natural_key, party, body)
  VALUES ('delete', old.id, old.natural_key, old.party, old.payload);
END;
CREATE TRIGGER IF NOT EXISTS rec_fts_upd AFTER UPDATE ON record BEGIN
  INSERT INTO record_fts(record_fts, rowid, natural_key, party, body)
  VALUES ('delete', old.id, old.natural_key, old.party, old.payload);
  INSERT INTO record_fts(rowid, natural_key, party, body)
  VALUES (new.id, new.natural_key, new.party, new.payload);
END;
"""


def has_fts(con):
    return bool(con.execute("SELECT name FROM sqlite_master WHERE type='table'"
                            " AND name='record_fts'").fetchone())


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
    try:
        con.executescript(FTS_SCHEMA)
    except sqlite3.OperationalError:
        # FTS5 없는 빌드 — 자유문 검색만 못 쓴다. 보관소는 그대로 열려야 한다.
        pass
    con.commit()
    return con


def log(con, area, action, ok=True, ref_kind=None, ref_id=None, detail=None, actor=None):
    """무슨 일이 있었는지 한 줄. **실패도 반드시 남긴다**(ok=0)."""
    con.execute(
        "INSERT INTO event(at,who,area,action,ok,ref_kind,ref_id,detail)"
        " VALUES(?,?,?,?,?,?,?,?)",
        (now(), actor or who(), area, action, 1 if ok else 0, ref_kind, ref_id,
         json.dumps(detail, ensure_ascii=False) if detail is not None else ""))


def note(area, action, ok=True, detail=None, ref_kind=None, ref_id=None, actor=None):
    """다른 도구가 **한 줄로** 로그를 남기는 길 (worksplit #20).

    ★ 이 함수는 **무슨 일이 있어도 예외를 내지 않는다.** 로그를 남기려다 본 작업을
      멈추면 그 순간 아무도 로그를 안 쓰게 된다. DB 가 잠겨 있어도 조용히 지나간다.
    ★ 로그는 append-only 다(트리거가 UPDATE·DELETE 를 막는다). 고쳐질 수 있으면
      근거가 아니다 — "8/5 돌발AS 가 왜 1건이었나"를 되짚을 때 믿을 것이 이 표다.
    """
    try:
        con = connect()
        try:
            log(con, area, action, ok=ok, ref_kind=ref_kind, ref_id=ref_id,
                detail=detail, actor=actor)
            con.commit()
        finally:
            con.close()
        return True
    except Exception:
        return False


# ── 기록(record) — 원본에서 뽑아낸 '한 건' (worksplit #21) ────────────────────
# asset 이 '파일'이라면 record 는 그 안의 **업무 한 건**이다. 밴드 글 하나, 정기점검
# 한 회, 돌발AS 한 건. 원본이 다시 수집돼 내용이 달라지면 record_rev 에 **무엇이
# 어떻게 바뀌었는지**를 남긴다 — 조용히 덮어쓰면 "어제와 숫자가 다르다"를 설명할 수 없다.
_REC_TRACK = ("biz_date", "party", "amount", "status")


def put_record(con, kind, natural_key, payload, biz_date=None, party="",
               amount=None, status="", asset_id=None, why="", actor=None,
               hash_on=None):
    """한 건을 넣거나 갱신한다. 바뀐 것이 없으면 아무것도 쓰지 않는다.

    `hash_on` 은 **'바뀌었나'를 판정할 때 볼 칸들**이다. 안 주면 payload 전체를 본다.
    ★ 왜 필요한가 (2026-08-08 실측): ERP 내보내기에는 **회차마다 달라지는 칸**이
      섞여 있다 — 계정별원장의 `잔액` 은 뽑은 기간이 어디서 시작하느냐에 따라 달라진다.
      전표 내용은 한 글자도 안 바뀌었는데 잔액이 달라서 '바뀜'이 됐고, 첫 흡수에서
      22,260번의 재회 중 **11,227번이 가짜 변경**으로 잡혔다. 그렇게 부푼 record_rev
      안에서는 진짜 변경(발행→완료, 금액 정정)을 영영 못 찾는다.
      그래서 판정은 **업무상 뜻이 있는 칸**으로만 한다. 나머지는 payload 에 그대로
      담아 화면에 보여 주되, 그것이 달라졌다고 변경으로 세지는 않는다.

    돌려주는 값은 (record_id, 'new'|'same'|'changed')."""
    body = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    if hash_on:
        core = {k: payload.get(k) for k in hash_on}
        core.update({"_날짜": biz_date, "_거래처": party, "_금액": amount, "_상태": status})
        seed = json.dumps(core, ensure_ascii=False, sort_keys=True, default=str)
    else:
        seed = body
    h = hashlib.sha1(seed.encode("utf-8")).hexdigest()
    ts = now()
    cur = con.execute("SELECT id,hash,biz_date,party,amount,status FROM record"
                      " WHERE kind=? AND natural_key=?", (kind, natural_key)).fetchone()
    if cur is None:
        rid = con.execute(
            "INSERT INTO record(kind,natural_key,asset_id,biz_date,party,amount,status,"
            "payload,hash,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (kind, natural_key, asset_id, biz_date, party or "", amount, status or "",
             body, h, ts, ts)).lastrowid
        return rid, "new"
    rid, old_hash = cur[0], cur[1]
    if old_hash == h:
        return rid, "same"
    olds = dict(zip(_REC_TRACK, cur[2:]))
    news = {"biz_date": biz_date, "party": party or "", "amount": amount,
            "status": status or ""}
    w = actor or who()
    for f in _REC_TRACK:
        if olds.get(f) != news.get(f):
            con.execute("INSERT INTO record_rev(record_id,at,who,field,old,new,why)"
                        " VALUES(?,?,?,?,?,?,?)",
                        (rid, ts, w, f, _s(olds.get(f)), _s(news.get(f)), why))
    # 본문이 바뀐 것도 남긴다(어느 칸인지 모를 때가 대부분이다 — 해시로 사실만 적는다)
    con.execute("INSERT INTO record_rev(record_id,at,who,field,old,new,why)"
                " VALUES(?,?,?,?,?,?,?)", (rid, ts, w, "payload", old_hash, h, why))
    con.execute("UPDATE record SET asset_id=COALESCE(?,asset_id),biz_date=?,party=?,"
                "amount=?,status=?,payload=?,hash=?,updated_at=? WHERE id=?",
                (asset_id, biz_date, party or "", amount, status or "", body, h, ts, rid))
    return rid, "changed"


def _s(v):
    return "" if v is None else str(v)


def band_day(v):
    """밴드 캐시의 `created_at` → `YYYY-MM-DD`. 못 읽으면 빈 문자열.

    ★ 캐시가 담는 값은 **밀리초 정수**다(convert_dump). 예전에는 이것을 문자열로
      바꿔 앞 열 자를 잘랐다 — `1766704935000` → `"1766704935"`. 날짜처럼 생기지도
      않은 값이 `biz_date` 에 7,782건 들어앉아 있었다(2026-08-08 발견).
      숫자라서 조용했다: 정렬도 되고 비교도 되니 아무 도구도 불평하지 않았고,
      "최근 30일"·"이번 달" 같은 **어떤 기간 질문에도 한 건도 안 걸렸다.**
      비어 있으면 눈에 띄지만 틀린 값은 안 띈다 — 그래서 여기 한 곳으로 모은다.
    """
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)) or str(v).isdigit():
        n = int(v)
        if n > 10 ** 12:                    # 밀리초
            n //= 1000
        if n <= 0:
            return ""
        try:
            return datetime.fromtimestamp(n).strftime("%Y-%m-%d")
        except (OverflowError, OSError, ValueError):
            return ""
    s = str(v).strip()
    return s[:10] if len(s) >= 10 and s[4] == "-" and s[7] == "-" else ""


def ingest_band(con, quiet=False, since=None, why="밴드 캐시 흡수"):
    """밴드 캐시(band/cache/*.json)의 글을 record 로 옮긴다.

    ★ **여기서 밴드를 긁지 않는다.** 수집은 'CSOS 리서치 및 자료 수집' 세션이 맡는다
      (CLAUDE.md). 이 함수는 이미 모여 있는 캐시 파일을 **읽기만** 한다.
    ★ 시각이 없는 글은 버린다 — 2026-08-07 실사고에서 밴드가 없는 글 번호에도 앱
      껍데기를 줘서 직전 글 본문이 마흔 건 복제됐다. 그 지문이 '시각 없음'이었다.

    `since`(YYYY-MM-DD) 를 주면 **그 날짜 이후 글만** 본다. 재수집 회차
    (`band/recollect.py`)가 최근 30일 창을 흡수할 때 쓴다 — 전량을 매번 다시
    비교하면 회차가 길어지고, 무엇이 이번 창에서 바뀌었는지도 흐려진다.

    돌려주는 값에 **무엇이** 바뀌었는지(`바뀐글`·`새글`)를 담는다. 개수만으로는
    인계 문서에 "3건 바뀜"밖에 못 적고, 사람은 결국 DB 를 다시 뒤져야 한다.
    """
    import glob as _g
    made = same = changed = skipped = junk = 0
    hit_new, hit_chg = [], []
    for fp in sorted(_g.glob(_shared("band", "cache", "*.json"))):
        base = os.path.basename(fp)
        if base.startswith("raw"):          # 원본 덤프는 convert_dump 가 다룬다
            continue
        try:
            d = json.load(open(fp, encoding="utf-8"))
        except Exception:
            continue
        band_id = os.path.splitext(base)[0]
        bname = str(d.get("band_name") or "")
        for no, post in sorted((d.get("posts") or {}).items()):
            if not str(no).isdigit() or not isinstance(post, dict):
                continue
            # ★ **삭제·오염·유령은 업무 기록이 아니다** (2026-08-08 발견).
            #   수집기·재수집·recheck_plan 은 셋 다 이것들을 걸렀는데 흡수기만 안 걸렀다.
            #   그래서 `contaminated`(남의 본문이 잡힌 기록)와 `absent`(처음부터 없던
            #   번호)가 record 표에 **진짜 업무 한 건**으로 앉았다. 캐시는 표시를
            #   달아 두는데 DB 는 그 표시를 안 읽으니, 표시해 둔 보람이 없었다.
            #   거르는 근거는 recheck_plan·recollect 와 **같은 네 가지**여야 한다 —
            #   한 곳만 다르면 화면마다 건수가 달라진다.
            if post.get("deleted") or post.get("contaminated") or post.get("absent"):
                junk += 1
                continue
            day = band_day(post.get("created_at"))
            if not day:
                skipped += 1                # 시각 없는 수확은 믿지 않는다
                continue
            if since and day < str(since):
                continue
            key = f"{band_id}/{no}"
            _rid, how = put_record(
                con, "band_post", key,
                {"밴드": bname, "밴드ID": band_id, "글번호": no,
                 "글쓴이": (post.get("author") or ""),
                 "본문": (post.get("content") or "")[:4000],
                 "사진수": post.get("photo_count"), "댓글수": post.get("comment_count"),
                 "수집시각": post.get("captured_at") or ""},
                biz_date=day, party=(post.get("author") or ""),
                status="", why=why)
            row = {"밴드": bname or band_id, "밴드ID": band_id, "글번호": no,
                   "작성일": day, "글쓴이": (post.get("author") or ""),
                   "요약": " ".join((post.get("content") or "").split())[:60]}
            if how == "new":
                made += 1
                hit_new.append(row)
            elif how == "changed":
                changed += 1
                hit_chg.append(row)
            else:
                same += 1
    con.commit()
    log(con, "band", "ingest_records", ok=True,
        detail={"신규": made, "변경": changed, "그대로": same, "시각없음버림": skipped,
                "삭제·오염·유령제외": junk, "창": since or "전체"})
    con.commit()
    if not quiet:
        print(f"  밴드 글 → 기록: 신규 {made} · 변경 {changed} · 그대로 {same}"
              f" · 시각 없어 버림 {skipped} · 삭제/오염/유령 제외 {junk}")
    return {"신규": made, "변경": changed, "그대로": same, "버림": skipped, "제외": junk,
            "새글": hit_new, "바뀐글": hit_chg}


# ── ERP 엑셀 → 기록 (worksplit #24) ──────────────────────────────────────────
# 화면마다 열 이름이 다르다. **어느 열이 그 화면의 신분증인가**만 여기 적어 두면
# 나머지(머리행 찾기·숫자 읽기·합계행 버리기)는 공통으로 처리된다.
#   키   = 자연키를 만드는 열들. 다시 받아도 같은 값이어야 한다(그래야 '바뀜'을 안다)
#   날짜 = 업무 날짜를 캐낼 열. `2026/07/01 -1` 처럼 번호가 붙어 있어도 앞 열 자를 쓴다
ERP_MAP = {
    # ★ 원장은 전표 하나에 **여러 줄**이다(차변·대변·적요별). `일자-No.` 만으로는
    #   8,820줄이 1,924키로 뭉개져, 같은 키를 서로 덮어쓰며 매번 '바뀜'이 됐다.
    "ERP:ledger":  [{"키": ["일자-No.", "거래처명", "적요"], "날짜": "일자-No.",
                     "거래처": "거래처명", "금액": "차변금액", "적요": "적요"},
                    # 분개장 — 같은 `ledger` 통이지만 전표번호가 `26/01/02-2-1`(두 자리 해)다
                    {"키": ["전표번호", "계정명", "거래처", "적요"], "날짜": "전표번호",
                     "거래처": "거래처", "금액": "차변", "적요": "적요"}],
    "ERP:tax":     {"키": ["일자-No."], "날짜": "일자-No.", "거래처": "거래처명",
                    "금액": "매출합계"},
    "ERP:taxstep": {"키": ["일자-No."], "날짜": "일자-No.", "거래처": "거래처명",
                    # ★ 열 이름은 **끝까지** 적어야 한다 — '진행'까지만 적었더니 상태가
                    #   통째로 빈칸이었다(발행/미발행을 못 가렸다. 2026-08-08).
                    "금액": "합계금액", "상태": "전자(세금)계산서 진행단계"},
    "ERP:slips":   [{"키": ["전표번호"], "날짜": "전표번호", "거래처": "거래처명",
                     "금액": "금액", "상태": "입력메뉴", "적요": "적요명"},
                    # 회계거래**현황** — 같은 `slips` 통인데 `입력메뉴` 가 없고 `거래유형` 이다.
                    # 이 모양을 안 적어 뒀더니 매출/매입을 못 갈라 53건이 통째로 빠졌다.
                    {"키": ["전표번호"], "날짜": "전표번호", "거래처": "거래처명",
                     "금액": "금액", "상태": "거래유형", "적요": "적요"}],
    "ERP:sales":   [{"키": ["일자", "PO번호", "거래처명", "품목명(요약)"], "날짜": "일자",
                     "거래처": "거래처명", "금액": "금액합계", "상태": "진행상태"},
                    # 주문서현황내역 — 판매현황과 같은 `sales` 통에 들어오지만 표가 다르다
                    {"키": ["발주일", "프로젝트코드", "프로젝트명"], "날짜": "발주일",
                     "거래처": "프로젝트명", "상태": "주문형태"}],
    # ★ 한 종류에 **표 모양이 둘 이상**일 수 있다 (2026-08-08). 거래명세서는
    #   낱장(상세: 일자·품목명[규격]·수량·단가)과 현황 목록(일자-No.·진행상태·
    #   금액합계)이 같은 `stmt` 로 분류된다. 앞의 것 하나만 적어 두면 뒤의 것은
    #   머리행을 못 찾아 **한 건도 안 읽힌다** — 그런데 파일은 있으니 아무도 모른다.
    "ERP:stmt":    [{"키": ["일자", "품목명[규격]"], "날짜": "일자",
                     "금액": "공급가액", "적요": "적요"},
                    {"키": ["일자-No.", "거래처명", "품목명"], "날짜": "일자-No.",
                     "거래처": "거래처명", "금액": "금액합계", "상태": "진행상태"}],
    "ERP:hometax": {"키": ["승인번호"], "날짜": "일자", "거래처": "공급받는자상호",
                    "금액": "공급가액", "상태": "계산서종류"},
    "ERP:taxinv":  {"키": ["일자 - 번호"], "날짜": "일자 - 번호", "거래처": "거래처명",
                    "금액": "합 계"},
    "ERP:quote":   {"키": ["일자-No.", "프로젝트코드코드", "품목명(요약)"], "날짜": "일자-No.",
                    "거래처": "거래처명", "금액": "견적금액합계", "상태": "진행상태",
                    "적요": "적요/수리내역/작업지시내용"},
}
_DATE_RE = re.compile(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})")


_DATE2_RE = re.compile(r"^(\d{2})[./-](\d{1,2})[./-](\d{1,2})\b")
# 내보내기 꼬리말의 인쇄 시각: `2026/08/05(수)오전11:11:15`
_FOOTER_RE = re.compile(r"\([월화수목금토일]\)\s*(오전|오후)")


def _erp_day(v):
    """`2026/07/01 -1` · `2026-07-01` · `26/01/02-2-1` · 엑셀 날짜 → `YYYY-MM-DD`.

    ★ 분개장의 전표번호는 **두 자리 해**다(`26/01/02-2-1`). 네 자리만 찾으면 이 화면이
      통째로 안 읽힌다 — 파일은 있는데 건수가 0인, 아무도 모르는 구멍이 된다.
    """
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    m = _DATE_RE.search(s)
    if m:
        return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    m = _DATE2_RE.match(s)
    if m:
        return "20%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return ""


def _erp_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v or ""))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _erp_header(rows, want):
    """머리행을 **이름으로** 찾는다.

    ★ '몇 번째 줄'로 정하면 안 된다 — 거래명세서는 6행, 나머지는 2행이고, 화면이
      바뀌면 또 달라진다. 찾는 열 이름이 가장 많이 보이는 줄이 머리행이다.
    """
    best, bi, cols = 0, None, {}
    for i, r in enumerate(rows[:15]):
        names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
        hit = sum(1 for w in want if w in names)
        if hit > best:
            best, bi, cols = hit, i, names
    return (bi, cols) if best >= 1 else (None, {})


def ingest_erp(con, quiet=False, only=None, force=False, limit=None):
    """ERP 내보내기 엑셀을 **한 건씩** record 로 옮긴다 (worksplit #24).

    ★ 여기서 ERP 를 내려받지 않는다. `erp_grab.py` 가 받아 Z: 에 둔 것을 읽기만 한다.
    ★ 이미 뜯어 본 파일은 건너뛴다(`parsed` 표). 판단 근거는 이름이 아니라
      **크기·수정시각**이라, 같은 이름으로 다시 받으면 다시 뜯는다.
    ★ 합계·소계 줄은 버린다 — 그것을 한 건으로 넣으면 금액이 두 번 세어진다.

    ★ **회차 안에서는 한 건을 한 번만 쓴다** (2026-08-08 실측). ERP 내보내기는 기간이
      서로 겹친다 — 같은 전표가 파일 열 개에 들어 있다. 파일마다 곧바로 쓰면 한 회차
      안에서 같은 건이 열 번 덮어써지고, 그 왕복이 전부 '바뀜'으로 기록된다.
      같은 파일들을 그대로 다시 뜯어도 **또 1,310건이 바뀌었다고 나왔다** — 아무것도
      안 변했는데. 그래서 먼저 전부 읽어 **오래된 것 → 새것** 순으로 겹쳐 최종값을
      만들고, 그 최종값만 한 번 쓴다. 그제야 '바뀜'이 진짜 바뀜을 뜻한다.
    """
    try:
        import openpyxl
    except ImportError:
        return {"오류": "openpyxl 없음"}
    kinds = [only] if only else list(ERP_MAP)
    # 오래된 것부터 — 뒤에 오는(더 새로운) 내보내기가 앞의 값을 덮는다.
    q = ("SELECT id,path,kind,mtime,size FROM asset WHERE gone_at IS NULL AND kind IN (%s)"
         " ORDER BY biz_date ASC, id ASC" % ",".join("?" * len(kinds)))
    todo = con.execute(q, kinds).fetchall()
    seen = {r["asset_id"]: r["fingerprint"] for r in con.execute("SELECT * FROM parsed")}
    made = same = changed = files = skipped = bad = 0
    hit_chg, noheader = [], []
    final = {}                  # 자연키 → 이 회차의 최종값. 다 읽은 뒤 한 번만 쓴다.
    for a in todo:
        # ★ 지문에 **종류**를 넣는다 (2026-08-08). 파일은 그대로인데 분류가 고쳐진
        #   경우가 있다 — 견적서조회 3장이 '거래명세서'로 앉아 있다가 바로잡혔다.
        #   크기·수정시각만 보면 "이미 뜯었다"고 넘어가, **규칙을 고쳐도 그 파일만
        #   영영 안 읽힌다.** 종류가 달라졌으면 다시 뜯어야 한다.
        fp = "%s:%s:%s" % (a["kind"], a["mtime"], a["size"])
        if not force and seen.get(a["id"]) == fp:
            skipped += 1
            continue
        if limit and files >= int(limit):
            break
        specs = ERP_MAP[a["kind"]]
        specs = specs if isinstance(specs, list) else [specs]
        try:
            wb = openpyxl.load_workbook(a["path"], data_only=True, read_only=False)
        except Exception:
            bad += 1
            continue
        n, seq = 0, {}
        try:
            for ws in wb.worksheets:
                rows = list(ws.iter_rows(values_only=True))
                # 모양이 여럿이면 **가장 잘 맞는 것**을 쓴다(칸 이름이 가장 많이 보이는 것)
                hi, cols, spec, want, best = None, {}, specs[0], [], 0
                for cand in specs:
                    w = [cand[k] for k in ("날짜", "거래처", "금액", "상태", "적요")
                         if cand.get(k)] + cand["키"]
                    h, c = _erp_header(rows, w)
                    score = sum(1 for x in w if x in c) if h is not None else 0
                    if score > best:
                        hi, cols, spec, want, best = h, c, cand, w, score
                # 변경 판정에 쓸 칸 — 이 화면에서 **업무상 뜻이 있다고 적어 둔 것**만.
                # 나머지(잔액·누계 같은 회차 의존 값)는 담아 보여 주되 변경으로 세지 않는다.
                core_cols = sorted(set(want))
                if hi is None or not all(k in cols for k in spec["키"][:1]):
                    # ★ **조용히 건너뛰지 않는다.** 머리행을 못 찾았다는 것은 대개
                    #   '이 파일이 그 종류가 아니다'라는 뜻이다(분류가 틀렸다).
                    #   말없이 넘기면 그 화면 건수가 영원히 모자란 채로 맞아 보인다.
                    #   ★ 다만 **시트 하나**가 안 맞는 것은 흔하다(빈 시트·요약 시트).
                    #     파일 안 어느 시트도 못 읽었을 때만 신고한다 — 안 그러면
                    #     정상 파일이 매 회차 경고에 올라 경고가 값을 잃는다.
                    continue
                get = lambda r, name: (r[cols[name]] if cols.get(name) is not None
                                       and cols[name] < len(r) else None)
                for r in rows[hi + 1:]:
                    if r is None or all(c is None for c in r):
                        continue
                    parts = [str(get(r, k) or "").strip() for k in spec["키"]]
                    if not parts[0]:
                        continue
                    joined = " ".join(str(c) for c in r if c is not None)
                    # 합계·소계·이월 — 한 건이 아니라 요약이다. 넣으면 금액이 겹친다.
                    if re.search(r"(월\s*계|누\s*계|소\s*계|합\s*계|이\s*월|총\s*계)", joined):
                        continue
                    day = _erp_day(get(r, spec["날짜"]))
                    if not day:
                        continue
                    # ★ 내보내기 **인쇄 꼬리말**을 한 건으로 세지 않는다 (2026-08-08).
                    #   `2026/08/05(수)오전11:11:15` — 날짜처럼 생겨 통과했고, 거래처·
                    #   금액이 빈 채로 record 에 앉아 대조표 '어긋남' 칸에 여덟 건이
                    #   올라왔다. 없는 일을 쫓게 만드는 줄이다.
                    #   ★ 처음엔 '내용이 하나도 없는 줄'로 넓게 걸렀다가 **진짜 줄
                    #     8건까지 잃었다**(적요를 안 적어 둔 화면에서는 늘 빈칸이라).
                    #     그래서 꼬리말의 **생김새 자체**만 집는다 — 인쇄 시각이다.
                    if _FOOTER_RE.search(str(get(r, spec["날짜"]) or "")):
                        continue
                    _party = str(get(r, spec.get("거래처")) or "").strip()
                    _amt = _erp_num(get(r, spec.get("금액")))
                    # ★ 그래도 남는 겹침에는 **그 전표 안에서 몇 번째 줄인가**를 붙인다.
                    #   열을 아무리 더해도 완전히 같은 줄이 두 번 나오는 화면이 있다
                    #   (원장의 차변·대변 짝). 순번을 안 붙이면 둘이 한 건으로 뭉개져
                    #   서로 덮어쓰고, 매 회차 '바뀜'으로 잡힌다 — 진짜 변경이 묻힌다.
                    #   순번은 **한 전표 안**에서만 세므로 뽑은 기간이 달라져도 안 흔들린다.
                    base = "%s/%s" % (a["kind"].split(":")[-1],
                                      "|".join(p for p in parts if p))
                    seq[base] = seq.get(base, 0) + 1
                    key = base if seq[base] == 1 else "%s#%d" % (base, seq[base])
                    payload = {str(name): (str(r[j]).strip() if j < len(r) and r[j] is not None
                                           else "")
                               for name, j in cols.items() if j is not None}
                    # ★ **어느 파일에서 왔는지를 payload 에 넣지 말 것** (2026-08-08 실측).
                    #   payload 는 해시로 '바뀌었나'를 판정하는 통이다. 파일명을 넣으면
                    #   같은 건을 다른 회차 파일에서 다시 만날 때마다 **내용이 똑같아도
                    #   '바뀜'** 이 된다 — 첫 시험에서 2,304행 중 1,669행이 그렇게 잡혔다.
                    #   그러면 record_rev 가 가짜로 부풀고 진짜 변경이 그 안에 묻힌다.
                    #   출처는 `asset_id` 가 이미 가리키고 있다.
                    final[key] = (a["kind"], payload, day, _party[:60], _amt,
                                  str(get(r, spec.get("상태")) or "")[:40],
                                  a["id"], core_cols)
                    n += 1
        finally:
            wb.close()
        if not n:
            noheader.append(os.path.basename(a["path"]))
        con.execute("INSERT OR REPLACE INTO parsed(asset_id,fingerprint,rows,at)"
                    " VALUES(?,?,?,?)", (a["id"], fp, n, now()))
        con.commit()
        files += 1

    for i, (key, v) in enumerate(sorted(final.items())):
        kind, payload, day, party, amount, status, aid, core_cols = v
        _rid, how = put_record(con, kind, key, payload, biz_date=day, party=party,
                               amount=amount, status=status, asset_id=aid,
                               why="ERP 엑셀 흡수", hash_on=core_cols)
        if how == "new":
            made += 1
        elif how == "changed":
            changed += 1
            hit_chg.append({"종류": kind, "건": key, "날짜": day})
        else:
            same += 1
        if i % 2000 == 1999:            # Z: 주사처럼 한 거래를 오래 붙들지 않는다
            con.commit()
    con.commit()
    log(con, "erp", "ingest_records", ok=not noheader,
        detail={"파일": files, "건너뜀(그대로)": skipped, "못연파일": bad,
                "신규": made, "변경": changed, "그대로": same,
                "머리행못찾음": sorted(set(noheader))[:20]})
    con.commit()
    if not quiet:
        print(f"  ERP 엑셀 → 기록: 파일 {files}개(건너뜀 {skipped}"
              + (f" · 못 연 것 {bad}" if bad else "") + ")"
              + f" · 신규 {made} · 변경 {changed} · 그대로 {same}")
        if noheader:
            u = sorted(set(noheader))
            print(f"  ※ 머리행을 못 찾은 파일 {len(u)}개 — 분류가 틀렸을 수 있다:"
                  f" {', '.join(u[:3])}" + (" …" if len(u) > 3 else ""))
    return {"파일": files, "건너뜀": skipped, "못연파일": bad, "신규": made,
            "변경": changed, "그대로": same, "바뀐건": hit_chg,
            "머리행못찾음": sorted(set(noheader))}


def repair_band_dates(con, quiet=False):
    """`band_post` 의 망가진 `biz_date`(에포크 초 문자열)를 날짜로 되돌린다.

    ★ 이것은 **원본이 바뀐 것이 아니라 우리가 잘못 적은 것**이다. 그래서 record_rev
      에 7,782줄을 남기지 않는다 — 남기면 진짜 변경(밴드 글 수정)이 그 더미에 묻혀
      영영 안 보인다. 대신 event 에 한 줄, 몇 건을 고쳤는지 사실만 적는다.
      되돌릴 근거는 캐시(원본)라서 언제든 다시 만들 수 있다.
    `--repair-band-date` 로 부른다. 이미 고쳐졌으면 0건이라 다시 돌려도 안전하다.
    """
    rows = con.execute("SELECT id,natural_key,biz_date FROM record WHERE kind='band_post'"
                       " AND (biz_date IS NULL OR biz_date NOT LIKE '____-__-__')").fetchall()
    fixed = dead = 0
    for r in rows:
        day = band_day(r["biz_date"])
        if day:
            con.execute("UPDATE record SET biz_date=? WHERE id=?", (day, r["id"]))
            fixed += 1
        else:
            dead += 1
    con.commit()
    log(con, "band", "repair_biz_date", ok=True,
        detail={"고침": fixed, "못고침": dead, "본것": len(rows)})
    con.commit()
    if not quiet:
        print(f"밴드 글 날짜 교정: {fixed}건 고침"
              + (f" · {dead}건은 값이 없어 못 고침(재수집 대상)" if dead else "")
              + ("" if rows else " — 고칠 것 없음"))
    return {"고침": fixed, "못고침": dead}


def record_changes(con, kind=None, at_since=None, limit=200):
    """record_rev 를 사람이 읽는 줄로. **언제·무엇이·어떻게** 바뀌었나.

    `at_since` 는 **바뀐 시각**(수집 시각)이지 업무 날짜가 아니다 — 재수집 회차가
    "이번 회차에 달라진 것"을 뽑을 때 쓰므로 기준이 회차 시각이어야 한다.
    본문 변경은 해시끼리 비교라 사람에게 뜻이 없다 → '본문 바뀜'으로 적는다.
    """
    q = ("SELECT v.at,v.who,v.field,v.old,v.new,v.why,r.kind,r.natural_key,r.biz_date,r.party"
         " FROM record_rev v JOIN record r ON r.id=v.record_id")
    w, args = [], []
    if kind:
        w.append("r.kind=?")
        args.append(kind)
    if at_since:
        w.append("v.at>=?")
        args.append(str(at_since))
    if w:
        q += " WHERE " + " AND ".join(w)
    q += " ORDER BY v.at DESC, v.id DESC LIMIT ?"
    args.append(int(limit))
    out = []
    for r in con.execute(q, args).fetchall():
        d = dict(r)
        if d["field"] == "payload":
            d["어떻게"] = "본문 바뀜"
        else:
            d["어떻게"] = "%s: %s → %s" % (d["field"], d["old"] or "(빈칸)",
                                           d["new"] or "(빈칸)")
        out.append(d)
    return out


# ── 이음(link) + 흐름 그림 (worksplit #22) ───────────────────────────────────
def link_records(con, src, dst, rel, conf=1.0, by=""):
    """건과 건을 잇는다(밴드 글 → 돌발AS → 정산 …).

    ★ `conf` 는 **얼마나 확실한가**다. 사람이 지정하면 1.0, 규칙이 추측하면 그보다
      낮게 둔다. 추측을 확정처럼 적으면 나중에 무엇을 다시 봐야 하는지 알 수 없다."""
    con.execute("INSERT OR REPLACE INTO link(src,dst,rel,conf,by,at)"
                " VALUES(?,?,?,?,?,?)",
                (int(src), int(dst), rel, float(conf), by or who(), now()))
    return True


def flow_mermaid(con, kind=None, since=None, limit=40):
    """이어진 건들을 Mermaid 흐름도로. 앱 [개발 사양]과 같은 문법을 쓴다."""
    q = ("SELECT l.src,l.dst,l.rel,l.conf,"
         " a.kind,a.natural_key,a.biz_date, b.kind,b.natural_key,b.biz_date"
         " FROM link l JOIN record a ON a.id=l.src JOIN record b ON b.id=l.dst")
    w, args = [], []
    if kind:
        w.append("(a.kind=? OR b.kind=?)")
        args += [kind, kind]
    if since:
        w.append("(a.biz_date>=? OR b.biz_date>=?)")
        args += [since, since]
    if w:
        q += " WHERE " + " AND ".join(w)
    q += " ORDER BY a.biz_date DESC, l.src LIMIT ?"
    args.append(int(limit))
    rows = con.execute(q, args).fetchall()
    if not rows:
        return 'flowchart LR' + chr(10) + '  E["이어진 건이 아직 없습니다"]'
    out, seen = ["flowchart LR"], set()
    def nid(rid):
        return "R%d" % int(rid)
    def label(k, key, d):
        t = f"{k}<br/>{key}" + (f"<br/>{d}" if d else "")
        return t.replace('"', "'")
    for (s, dd, rel, conf, ak, akey, ad, bk, bkey, bd) in rows:
        for rid, k, key, dt in ((s, ak, akey, ad), (dd, bk, bkey, bd)):
            if rid not in seen:
                seen.add(rid)
                out.append(f'  {nid(rid)}["{label(k, key, dt)}"]')
        tag = rel + ("" if conf >= 0.999 else f" ({int(conf*100)}%)")
        out.append(f"  {nid(s)} -->|{tag}| {nid(dd)}")
    return chr(10).join(out)


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
    # ★ **거친 종류가 정밀한 종류를 덮으면 안 된다** (2026-08-08 실측).
    #   내용 판별은 새것·바뀐 것에만 돌린다(느려서). 그래서 그다음 주사부터는
    #   폴더만 보고 온 'ERP' 가 들어오는데, 그것이 먼저 알아낸 'ERP:taxstep' 을
    #   지워 버렸다. 잔량을 애써 갈라 놓고 이틀 뒤에 도로 묻는 셈이다.
    if ":" in (row["kind"] or "") and ":" not in (kind or ""):
        kind = row["kind"]
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


def reclassify(con, limit=500, quiet=False, deep=False):
    """종류가 아직 **거친** 엑셀을 내용으로 다시 갈라 준다 → 바꾼 수.

    주사는 새것·바뀐 것에만 내용 판별을 돌린다(엑셀을 여는 일이라 느리다). 그래서
    판별 규칙을 **나중에 고치면** 이미 들어와 있는 파일은 옛 종류 그대로 남는다 —
    잔량(`taxstep`) 규칙을 새로 넣고도 101행짜리 파일이 계속 'ERP' 였던 이유다.
    이 함수가 그 뒤처리를 맡는다. 한 번에 다 하지 않고 상한을 둔다.

    `deep=True` 면 **이미 갈라진 `ERP:*` 도 다시 본다** (2026-08-08 추가).
    기본값이 '거친 것만'인 데에는 이유가 있다 — 매번 전부 다시 여는 것은 느리다.
    그런데 그 때문에 **판별이 틀렸던 것은 규칙을 고쳐도 영영 안 고쳐졌다**:
    견적서조회 3장이 '거래명세서'로 앉아 있었고, 규칙을 바로잡아도 여기 안 걸렸다.
    규칙을 고친 뒤에는 `--reclassify-deep` 을 한 번 돌린다.
    """
    try:
        from inbox_scan import classify
    except Exception:
        return 0
    where = ("kind LIKE 'ERP%'" if deep else "kind NOT LIKE '%:%'")
    rows = con.execute(
        "SELECT id, path, kind FROM asset WHERE gone_at IS NULL AND " + where +
        " AND path LIKE '%.xlsx' LIMIT ?", (int(limit),)).fetchall()
    n = 0
    for r in rows:
        try:
            k = classify(r["path"])
        except Exception:
            continue
        if k and k != "unknown" and f"ERP:{k}" != r["kind"]:
            con.execute("UPDATE asset SET kind=? WHERE id=?", (f"ERP:{k}", r["id"]))
            n += 1
    con.commit()
    log(con, "intake", "datalake.reclassify", detail={"바꿈": n, "본것": len(rows),
                                                      "깊게": bool(deep)})
    con.commit()
    if not quiet:
        print(f"내용 재판별: {n}건 갈라냄 (본 것 {len(rows)})")
    return n


def _cmp(spec):
    """`>1000000` · `<=5` · `2026-08-01` → (연산자, 값). 기본은 `=`."""
    s = str(spec).strip()
    for op in (">=", "<=", "!=", ">", "<", "="):
        if s.startswith(op):
            return op, s[len(op):].strip()
    return "=", s


def find(con, on="asset", kind=None, since=None, until=None, q=None, party=None,
         amount=None, bucket=None, gone=False, limit=50, order=None):
    """인덱스 필터 검색 — **CLI 도 앱도 이 함수 하나를 부른다.**

    두 벌로 만들면 결과가 갈리고, 갈린 것을 알아채는 데 또 며칠이 든다(설계서).

    `on='asset'` 은 원본 파일, `on='record'` 는 뽑아낸 업무 레코드다.
    `q` 는 자유문 — record 에서는 FTS5, asset 에서는 경로 부분일치로 받는다
    (asset 에는 본문이 없다. 없는 것을 있는 척하지 않는다).
    """
    args = []
    if on == "record":
        base = ("SELECT r.id, r.kind, r.natural_key, r.biz_date, r.party, r.amount,"
                " r.status FROM record r")
        where = ["1=1"]
        if q:
            if not has_fts(con):
                raise RuntimeError("이 파이썬의 SQLite 에 FTS5 가 없다 — 자유문 검색 불가")
            base += " JOIN record_fts f ON f.rowid = r.id"
            where.append("record_fts MATCH ?")
            args.append(q)
        col_date, col_kind = "r.biz_date", "r.kind"
        if party:
            where.append("r.party LIKE ?"); args.append(f"%{party}%")
        if amount:
            op, v = _cmp(amount)
            where.append(f"r.amount {op} ?"); args.append(int(v))
    else:
        base = ("SELECT id, kind, path, bucket, biz_date, size, last_seen, gone_at"
                " FROM asset")
        where = ["gone_at IS NOT NULL" if gone else "gone_at IS NULL"]
        col_date, col_kind = "biz_date", "kind"
        if q:
            where.append("path LIKE ?"); args.append(f"%{q}%")
        if bucket:
            where.append("bucket LIKE ?"); args.append(f"%{bucket}%")

    if kind:
        # `kind=ERP` 는 `ERP:tax` 까지 잡는다 — 사람은 큰 갈래로 먼저 묻는다
        where.append(f"({col_kind}=? OR {col_kind} LIKE ?)")
        args += [kind, f"{kind}:%"]
    if since:
        where.append(f"{col_date} >= ?"); args.append(since)
    if until:
        where.append(f"{col_date} <= ?"); args.append(until)

    sql = f"{base} WHERE {' AND '.join(where)} ORDER BY {order or col_date} DESC LIMIT ?"
    args.append(int(limit))
    return con.execute(sql, args).fetchall()


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
    ap.add_argument("--reclassify", type=int, nargs="?", const=500, metavar="N",
                    help="종류가 거친 엑셀을 내용으로 다시 가른다(판별 규칙을 고친 뒤)")
    ap.add_argument("--find", nargs="*", metavar="키=값",
                    help="on=record kind=… since=… until=… q=… party=… amount='>100'")
    ap.add_argument("--limit", type=int, default=50)
    ap.add_argument("--fill-sha1", type=int, nargs="?", const=2000, metavar="N",
                    help="지문이 빈 자산에 sha1 을 채운다(첫 주사는 일부러 안 잰다)")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--log", nargs="*", metavar="키=값", help="area=… since=… action=…")
    ap.add_argument("--fail-only", action="store_true")
    ap.add_argument("--db", help="DB 경로(합성검증용)")
    # ★ 아래 둘은 **모여 있는 캐시를 읽기만** 한다. 밴드를 새로 긁지 않는다
    #   (수집은 'CSOS 리서치 및 자료 수집' 세션이 맡는다 — CLAUDE.md).
    ap.add_argument("--band", action="store_true",
                    help="밴드 캐시(band/cache)의 글을 record 로 흡수(수집 아님)")
    ap.add_argument("--reclassify-deep", type=int, nargs="?", const=1000, metavar="N",
                    help="이미 갈라진 ERP:* 도 다시 판별한다(판별 규칙을 고친 뒤 한 번)")
    ap.add_argument("--erp", action="store_true",
                    help="ERP 내보내기 엑셀을 한 건씩 record 로 흡수(내려받기 아님)")
    ap.add_argument("--erp-force", action="store_true",
                    help="이미 뜯어 본 파일도 다시 뜯는다")
    ap.add_argument("--erp-kind", help="ERP 화면 한 종류만 (예: ERP:taxstep)")
    ap.add_argument("--repair-band-date", action="store_true",
                    help="밴드 글의 망가진 biz_date(에포크 초)를 날짜로 교정")
    ap.add_argument("--flow", nargs="*", metavar="키=값",
                    help="이어진 건들을 Mermaid 흐름도로 (kind=… since=…)")
    a = ap.parse_args(argv)

    con = connect(a.db)
    try:
        if a.scan:
            scan(con, rescan=a.rescan, limit_roots=a.only)
        if a.fill_sha1:
            fill_sha1(con, limit=a.fill_sha1)
        if a.reclassify:
            reclassify(con, limit=a.reclassify)
        if a.reclassify_deep:
            reclassify(con, limit=a.reclassify_deep, deep=True)
        if a.find is not None:
            f = _kv(a.find)
            on = f.pop("on", "asset")
            try:
                rows = find(con, on=on, limit=a.limit,
                            **{k: v for k, v in f.items()
                               if k in ("kind", "since", "until", "q", "party",
                                        "amount", "bucket")})
            except Exception as e:
                print(f"✗ {e}")
                rows = []
            # 경로는 **Z: 기준 상대경로**로 보인다. 전체 경로는 사람이 못 읽는다
            # (앞의 예순 자가 매 줄 똑같아서 정작 다른 부분이 화면 밖으로 밀린다).
            try:
                import source_dirs as _S
                base = _S.ORIGIN_ROOT + os.sep
            except Exception:
                base = ""
            for r in rows:
                d = dict(r)
                head = str(d.get("path") or d.get("natural_key") or "")
                if base and head.startswith(base):
                    head = head[len(base):]
                print(f"  {d.get('biz_date') or '-'}  [{d.get('kind')}]  {head}")
            print(f"  — {len(rows)}건" + (" (상한에 걸림 — --limit 을 올릴 것)"
                                          if len(rows) == a.limit else ""))
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
        if a.erp or a.erp_force:
            ingest_erp(con, only=a.erp_kind, force=a.erp_force)
        if a.repair_band_date:
            repair_band_dates(con)
        if a.band:
            ingest_band(con)
        if a.flow is not None:
            f = _kv(a.flow)
            print(flow_mermaid(con, kind=f.get("kind"), since=f.get("since"),
                               limit=int(f.get("limit") or a.limit)))
        if a.status or not (a.scan or a.fill_sha1 or a.reclassify or a.find is not None
                            or a.log is not None or a.band or a.flow is not None
                            or a.repair_band_date or a.erp or a.erp_force
                            or a.reclassify_deep):
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
