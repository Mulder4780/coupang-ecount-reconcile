# -*- coding: utf-8 -*-
"""
usb_snapshot.py : USB(ORICO) 백업용 가벼운 DB 사본을 만든다 (2026-11-20 형님 지시)

  python usb_snapshot.py      # db/usb_snapshot/ 에 사본을 만든다

왜
  · app_store.db 는 4.1GB 인데 업무 자료(work_item·work_field 등)는 수십 MB 다.
    나머지는 변경 기록(change_event)·중복 방지 열쇠(idempotency_key)·보낼 함(outbox)이다.
  · 형님 지시: "db만 백업 · 용량 크게 먹는 파일은 절대 백업하지 말고 · 다른 pc 에서 구동".
    그래서 표 구조는 전부 두고, 위 세 기록 표는 **구조만(줄 0)** 담는다.
    (원본은 한 글자도 안 건드린다 · 읽기 전용으로 연다)
  · 이 사본으로 앱은 뜬다. 다만 변경 이력(감사로그)은 사본에 없다 : 정본은 이 PC 의 원본이다.
"""
import os, sys, sqlite3, json, shutil
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "db", "usb_snapshot")
LOG_TABLES = ("change_event", "idempotency_key", "outbox")
KEEP_DAYS = 0   # 기록 표는 구조만 담는다(줄 0)
PLAIN = ("ledger_queue.db", "pm_content.db")   # datalake.db(300MB) 는 ERP 원본을 다시 읽어 만들 수 있어 뺀다


def _ro(p):
    return sqlite3.connect("file:%s?mode=ro" % p.replace("\\", "/"), uri=True, timeout=60)


def _time_col(con, t):
    cols = [r[1] for r in con.execute('pragma table_info("%s")' % t)]
    for c in ("created_at", "at", "ts", "updated_at", "occurred_at"):
        if c in cols:
            return c
    return None


def slim_app_store(dst):
    src = os.path.join(ROOT, "db", "app_store.db")
    tmp = dst + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)
    s = _ro(src)
    d = sqlite3.connect(tmp)
    since = (datetime.now() - timedelta(days=KEEP_DAYS)).isoformat()
    kept = {}
    for name, sql in s.execute("select name, sql from sqlite_master where type='table' and sql is not null "
                               "and name not like 'sqlite_%'").fetchall():
        d.execute(sql)
        cols = [r[1] for r in s.execute('pragma table_info("%s")' % name)]
        q = 'select * from "%s"' % name
        args = ()
        if name in LOG_TABLES:
            tc = _time_col(s, name)
            if tc:
                q += ' where "%s" >= ?' % tc
                args = (since,)
            else:
                q += " limit 0"
        rows = s.execute(q, args).fetchall()
        d.executemany('insert into "%s" values (%s)' % (name, ",".join("?" * len(cols))), rows)
        kept[name] = len(rows)
    for (sql,) in s.execute("select sql from sqlite_master where type in ('index','trigger','view') "
                            "and sql is not null").fetchall():
        try:
            d.execute(sql)
        except sqlite3.Error:
            pass
    d.commit()
    d.execute("vacuum")
    d.close()
    s.close()
    os.replace(tmp, dst)
    return kept


def copy_db(name):
    src = os.path.join(ROOT, "db", name)
    dst = os.path.join(OUT, name)
    if not os.path.exists(src):
        return None
    s = _ro(src)
    d = sqlite3.connect(dst + ".tmp")
    s.backup(d)          # WAL 까지 일관 사본
    d.close()
    s.close()
    os.replace(dst + ".tmp", dst)
    return os.path.getsize(dst)


def run():
    os.makedirs(OUT, exist_ok=True)
    meta = {"만든때": datetime.now().isoformat(timespec="seconds"), "최근일수": KEEP_DAYS}
    meta["app_store_줄수"] = slim_app_store(os.path.join(OUT, "app_store.db"))
    for n in PLAIN:
        meta[n] = copy_db(n)
    for n in sorted(os.listdir(OUT)):
        if n.endswith(".db"):
            meta.setdefault("크기MB", {})[n] = round(os.path.getsize(os.path.join(OUT, n)) / 1048576, 1)
    with open(os.path.join(OUT, "_meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print("USB 사본 → %s · %s" % (OUT, meta["크기MB"]))
    return meta


if __name__ == "__main__":
    run()
