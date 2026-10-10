# -*- coding: utf-8 -*-
"""UNI-CashFlow 와 자료를 나누는 다리 (2026-10-10 형님 지시).

형님 지시: "UNI-CashFlow 센션과 공유할 수 있는 데이터는 db로 공유해 앞으로 쭉"

★ 왜 DB 인가 — 그 앱은 2026-08-22 에 HTTP 열쇠를 **일부러 끊었다**(실측: 발급돼
  있었는데 호출 0회). 그러니 다시 API 로 엮지 않는다. 파일 하나를 사이에 두면
  서버가 떠 있지 않아도 되고, 어느 쪽이 죽어도 상대가 안 죽는다.

★ 서로의 정본을 **직접 열지 않는다**. 저쪽 db/treasury.db, 우리 db/app_store.db
  는 각자의 정본이고(2026-08-10 규칙) WAL·락·스키마가 제 앱 것이다. 남의 파일을
  열면 그 앱이 스키마를 고치는 날 **조용히 0건**이 된다([165]).
  그래서 가운데에 **공유 전용 DB** 하나를 둔다.

★ 규칙 셋 (이것이 계약이다)
  (1) **자기 출처앱 행만 쓴다.** 남의 행은 **읽기만** 한다 - 지우지도 고치지도 않는다.
  (2) **해석하지 않고 숫자를 그대로** 넘긴다. 실측으로 미청구액이 음수고
      발행금액에 0 이 많은데 그 0 이 '없음'인지 '안 봄'인지 **나는 모른다**([169]).
      뜻을 붙여 넘기면 저쪽이 그 해석을 사실로 믿는다 - 숫자와 **잰 때**만 준다.
  (3) **역수입 금지 그대로**(2026-08-10). 여기서 읽은 값을 우리 정본에 자동으로
      쓰지 않는다. 보여 주기만 하고, 업무 확정은 앱 API 를 다시 거친다.

★ 자리는 **로컬**이다 - Documents/_공유DB/bridge.db. 공유폴더(Z:)에 두지 않는다:
  저쪽 규칙이 "원장은 공유폴더로 안 간다 - 거기는 다른 사람도 연다" 라 적었고,
  실측으로 그 폴더는 잠들었다 깨는 데 69초다(SQLite 락이 버티지 못한다).

쓰는 법:
  python share_bridge.py              # 상태만 (누가 언제 무엇을 넣었나)
  python share_bridge.py --push       # 우리 정본에서 뽑아 공유 DB 에 쓴다
  python share_bridge.py --pull       # 저쪽이 넣은 것을 읽어 보여 준다
"""
from __future__ import annotations

import json
import os
import sqlite3
import sys
from datetime import datetime, timezone, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
# 두 저장소 **밖**이다 - 어느 git 에도 안 들어간다.
SHARE_DIR = os.environ.get(
    "ULH_SHARE_DB_DIR",
    os.path.abspath(os.path.join(HERE, "..", "..", "_공유DB")),
)
SHARE_DB = os.path.join(SHARE_DIR, "bridge.db")

ME = "coupang"          # 이 앱이 쓰는 출처앱 이름
PEER = "cashflow"       # 저쪽
APP_DB = os.path.join(HERE, "db", "app_store.db")

KST = timezone(timedelta(hours=9))

FIELDS = (
    "미청구액", "미수금액", "청구상태", "비용구분",
    "세금계산서대비입금차액", "발행금액", "입금액",
    "발행상태(자동)", "실제발행일", "청구일",
    "세금계산서발행일", "거래명세서발행일",
)


def now_text() -> str:
    return datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")


def _ro(path: str) -> sqlite3.Connection:
    return sqlite3.connect("file:" + path.replace("\\", "/") + "?mode=ro", uri=True)


# ---------------------------------------------------------------- 공유 DB

SCHEMA = """
CREATE TABLE IF NOT EXISTS shared_fact (
  출처앱   TEXT NOT NULL,
  갈래     TEXT NOT NULL,
  열쇠     TEXT NOT NULL,
  값_json  TEXT NOT NULL,
  잰때     TEXT NOT NULL,
  PRIMARY KEY (출처앱, 갈래, 열쇠)
);
CREATE TABLE IF NOT EXISTS share_meta (
  출처앱   TEXT NOT NULL,
  갈래     TEXT NOT NULL,
  쓴때     TEXT NOT NULL,
  건수     INTEGER NOT NULL,
  낸곳     TEXT NOT NULL DEFAULT '',
  말       TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (출처앱, 갈래)
);
"""


def open_share(readonly: bool = False) -> sqlite3.Connection:
    """공유 DB 를 연다. 없으면 만든다(쓰기일 때만)."""
    if readonly:
        if not os.path.exists(SHARE_DB):
            raise FileNotFoundError(SHARE_DB)
        con = _ro(SHARE_DB)
    else:
        os.makedirs(SHARE_DIR, exist_ok=True)
        con = sqlite3.connect(SHARE_DB, timeout=20)
        con.executescript(SCHEMA)
        con.commit()
    con.row_factory = sqlite3.Row
    return con


def write_facts(갈래: str, rows: dict, 낸곳: str, 말: str = "") -> int:
    """내 몫(출처앱=ME)만 갈아 쓴다. 남의 행은 한 글자도 안 건드린다."""
    con = open_share()
    try:
        t = now_text()
        with con:
            # 내 갈래만 비운다 - 출처앱 조건을 빼면 남의 것이 날아간다.
            con.execute("DELETE FROM shared_fact WHERE 출처앱=? AND 갈래=?", (ME, 갈래))
            con.executemany(
                "INSERT INTO shared_fact (출처앱,갈래,열쇠,값_json,잰때) VALUES (?,?,?,?,?)",
                [
                    (ME, 갈래, str(k), json.dumps(v, ensure_ascii=False), t)
                    for k, v in rows.items()
                ],
            )
            con.execute(
                "INSERT INTO share_meta (출처앱,갈래,쓴때,건수,낸곳,말)"
                " VALUES (?,?,?,?,?,?)"
                " ON CONFLICT(출처앱,갈래) DO UPDATE SET"
                " 쓴때=excluded.쓴때, 건수=excluded.건수,"
                " 낸곳=excluded.낸곳, 말=excluded.말",
                (ME, 갈래, t, len(rows), 낸곳, 말),
            )
        return len(rows)
    finally:
        con.close()


def read_facts(출처앱: str = PEER, 갈래=None) -> dict:
    """상대가 넣은 것을 읽는다. 못 읽으면 '모름' 이라 적는다 - 0 이라 하지 않는다."""
    try:
        con = open_share(readonly=True)
    except Exception as e:                      # 파일 없음·깨짐
        return {"모름": "공유 DB 를 못 읽었다: %s" % e}
    try:
        q = "SELECT 갈래,열쇠,값_json,잰때 FROM shared_fact WHERE 출처앱=?"
        a = [출처앱]
        if 갈래:
            q += " AND 갈래=?"
            a.append(갈래)
        out = {}
        for r in con.execute(q, a):
            out.setdefault(r["갈래"], {})[r["열쇠"]] = {
                "값": json.loads(r["값_json"]),
                "잰때": r["잰때"],
            }
        meta = {
            r["갈래"]: dict(r)
            for r in con.execute("SELECT * FROM share_meta WHERE 출처앱=?", (출처앱,))
        }
        return {"자료": out, "메타": meta}
    finally:
        con.close()


# ------------------------------------------------- 우리 정본에서 뽑는 자리

def _app_rows() -> list:
    """app_store.db 를 **읽기만** 한다."""
    con = _ro(APP_DB)
    con.row_factory = sqlite3.Row
    try:
        marks = ",".join("?" * len(FIELDS))
        return con.execute(
            "SELECT w.kind, w.public_id, w.project_no, w.camp_name, w.status,"
            "       f.field_key, f.value_json"
            "  FROM work_item w"
            "  JOIN work_field f ON f.work_id = w.id"
            " WHERE w.deleted_at IS NULL"
            "   AND f.field_key IN (" + marks + ")",
            FIELDS,
        ).fetchall()
    finally:
        con.close()


def _num(raw):
    """숫자만 돌려준다. 못 읽으면 None (0 으로 뭉개지 않는다 - [169])."""
    if raw is None:
        return None
    try:
        v = json.loads(raw)
    except Exception:
        v = raw
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.replace(",", "").replace("원", "").strip()
        try:
            return float(s) if "." in s else int(s)
        except Exception:
            return None
    return None


def _txt(raw):
    if raw is None:
        return ""
    try:
        v = json.loads(raw)
    except Exception:
        v = raw
    return v if isinstance(v, str) else ("" if v is None else str(v))


def build_push() -> dict:
    """공유할 것을 만든다. **해석을 붙이지 않는다** - 숫자와 건수만."""
    by_item = {}
    for r in _app_rows():
        key = (r["kind"], r["public_id"] or "", r["project_no"] or "", r["camp_name"] or "")
        d = by_item.setdefault(
            key,
            {
                "갈래": r["kind"],
                "공개번호": r["public_id"],
                "프로젝트NO": r["project_no"],
                "캠프명": r["camp_name"],
                "진행상태": r["status"],
            },
        )
        d[r["field_key"]] = r["value_json"]

    청구 = {}
    for d in by_item.values():
        if d["갈래"] != "정산":
            continue
        k = d["공개번호"] or ("무번호:" + (d["프로젝트NO"] or "?") + ":" + (d["캠프명"] or ""))
        청구[k] = {
            "프로젝트NO": d.get("프로젝트NO"),
            "캠프명": d.get("캠프명"),
            "비용구분": _txt(d.get("비용구분")),
            "청구상태": _txt(d.get("청구상태")),
            "미청구액": _num(d.get("미청구액")),
            "미수금액": _num(d.get("미수금액")),
            "입금차액": _num(d.get("세금계산서대비입금차액")),
            "거래명세서발행일": _txt(d.get("거래명세서발행일")),
            # ★ `입금액` 은 **정산**에 붙어 있다 — 처음에 세금계산서 갈래에서만
            #   꺼내서 62건이 **조용히 안 실렸다**(`[165]` — 안 읽은 칸은 빈칸과
            #   구별되지 않는다). 실측: 입금액 62건 전부 kind=정산.
            "입금액": _num(d.get("입금액")),
        }

    계산서 = {}
    n = 0
    for d in by_item.values():
        if d["갈래"] != "세금계산서":
            continue
        n += 1
        k = d["공개번호"] or ("세금계산서:%03d" % n)
        계산서[k] = {
            "프로젝트NO": d.get("프로젝트NO"),
            "캠프명": d.get("캠프명"),
            "발행금액": _num(d.get("발행금액")),
            "입금액": _num(d.get("입금액")),
            "발행상태": _txt(d.get("발행상태(자동)")),
            "실제발행일": _txt(d.get("실제발행일")),
            "법정기한": _txt(d.get("세금계산서발행일")),
        }

    # 건수 요약 - 저쪽 화면이 한 줄로 쓸 수 있게.
    요약 = {}
    con = _ro(APP_DB)
    try:
        for kind, c in con.execute(
            "SELECT kind, COUNT(*) FROM work_item WHERE deleted_at IS NULL GROUP BY kind"
        ):
            요약[kind] = c
    finally:
        con.close()

    return {"정산청구": 청구, "세금계산서": 계산서, "업무건수": 요약}


def _source_note() -> str:
    """금액이 **어느 관리대장 버전**에서 온 것인지 센다.

    ★ `잰때`(내가 공유 DB 에 쓴 시각)와 **원본이 얼마나 낡았나**는 다른 말이다
      (`[233]`). 실측 2026-10-10: 금액 칸의 출처가 전부 **v583** 인데 최신은
      **v635** 였다 — 쓴때만 보면 오늘 값처럼 보이지만 52개 버전 전 값이다.
      이 한 줄이 없으면 저쪽이 낡은 금액을 오늘 값으로 읽는다.
    ★ 못 세면 **지어내지 않는다**(`[169]`) — 빈 문자열."""
    import re
    import collections
    try:
        con = _ro(APP_DB)
        try:
            cnt = collections.Counter()
            for (ref,) in con.execute(
                "SELECT source_ref FROM work_field"
                " WHERE field_key IN ('미청구액','발행금액','입금액')"
                "   AND source_ref IS NOT NULL"
            ):
                m = re.search(r"_v(\d+)\.xlsx", ref or "")
                if m:
                    cnt[int(m.group(1))] += 1
        finally:
            con.close()
        if not cnt:
            return ""
        쓴판 = sorted(cnt)
        try:
            import workbook_patch
            _, 최신 = workbook_patch.latest_master()
            최신 = int(최신)
        except BaseException:          # latest_master 는 SystemExit 을 던진다
            최신 = None
        말 = "금액의 원본 관리대장 판: v%s" % ", v".join(str(v) for v in 쓴판)
        if 최신 is not None:
            뒤처짐 = 최신 - max(쓴판)
            말 += " (최신은 v%d" % 최신
            말 += " — %d판 뒤처졌다)" % 뒤처짐 if 뒤처짐 > 0 else ")"
        return 말
    except Exception:
        return ""


def push() -> dict:
    made = build_push()
    낡음 = _source_note()
    out = {}
    out["정산청구"] = write_facts(
        "정산청구",
        made["정산청구"],
        낸곳="app_store.db / work_item kind=정산",
        말=(
            # ★ 2026-10-10 에 관리대장 v635 의 **수식**을 읽어 뜻이 가려졌다.
            #   그 전까지 "모른다"고 보냈던 두 가지가 여기서 끝났다.
            "미청구액 = 실제작업합계(K) - 거래명세서합계(Q) 다. "
            "680건이 음수인 것은 돈이 모자란 뜻이 아니라 "
            "**실제작업합계 칸이 비어 있다**는 뜻이다(750건 중 값이 있는 것 68건뿐). "
            "미수금액·입금차액이 750건 전부 0 인 것도 그 수식이 참조하는 "
            "세금계산서합계(X)가 **750건 전부 비어 있어서**다 — '미수가 없다'가 아니라 "
            "**계산할 재료가 없다**. 그 둘은 업무 판단에 쓰지 말 것. "
            "입금액은 쓸 수 있다(앱 DB 62건 — 엑셀에는 8건뿐이라 앱이 더 많이 안다). "
            + 낡음
        ).strip(),
    )
    out["세금계산서"] = write_facts(
        "세금계산서",
        made["세금계산서"],
        낸곳="app_store.db / work_item kind=세금계산서",
        말=("발행금액 63건 전부 0 — **금액이 없는 게 아니라 안 채워졌다**. "
            "실측 2026-10-10: 관리대장 v635 의 세금계산서합계(X) 열이 "
            "750건 전부 비어 있다. 그러니 이 금액으로 집계하지 말 것. "
            "발행상태(발행기한 임박 57 · 발행 완료 6)는 쓸 수 있다. "
            + 낡음).strip(),
    )
    out["업무건수"] = write_facts(
        "업무건수",
        made["업무건수"],
        낸곳="app_store.db / work_item",
        말="지워지지 않은 건만 센 것.",
    )
    return out


# ------------------------------------------------------------------ 화면

def status_text() -> str:
    if not os.path.exists(SHARE_DB):
        return (
            "공유 DB 가 아직 없다: %s\n"
            "  -> python share_bridge.py --push 로 만든다." % SHARE_DB
        )
    lines = ["공유 DB: %s (%s bytes)" % (SHARE_DB, os.path.getsize(SHARE_DB))]
    con = open_share(readonly=True)
    try:
        got = False
        for r in con.execute("SELECT * FROM share_meta ORDER BY 출처앱, 갈래"):
            got = True
            mine = "  <- 내 것" if r["출처앱"] == ME else ""
            lines.append(
                "  [%s] %s - %s건 · %s%s"
                % (r["출처앱"], r["갈래"], r["건수"], r["쓴때"], mine)
            )
            if r["말"]:
                lines.append("        말: %s" % r["말"])
        if not got:
            lines.append("  (아직 아무도 아무것도 안 넣었다)")
        peer = con.execute(
            "SELECT COUNT(*) FROM shared_fact WHERE 출처앱=?", (PEER,)
        ).fetchone()[0]
        if not peer:
            lines.append(
                "  * 저쪽(%s)이 넣은 것은 0건이다 - 아직 그 앱에 내보내는 자리를"
                " 안 만든 것이지, 자료가 없다는 뜻이 아니다." % PEER
            )
        return "\n".join(lines)
    finally:
        con.close()


def main(argv) -> int:
    if "--push" in argv:
        try:
            out = push()
        except Exception as e:
            print("내보내기 실패: %s" % e)
            return 1
        for k, v in out.items():
            print("  %s: %s건 보냄" % (k, v))
        print(status_text())
        return 0
    if "--pull" in argv:
        got = read_facts(PEER)
        if "--json" in argv:
            print(json.dumps(got, ensure_ascii=False, indent=2)[:4000])
            return 0
        if "모름" in got:
            print("모름 - %s" % got["모름"])
            return 3
        if not got["자료"]:
            print(
                "저쪽(%s)이 넣은 것이 아직 없다. (없는 것과 못 읽은 것은 다른 말인데,"
                " 여기서는 '읽었고 0건' 이다)" % PEER
            )
            return 0
        for 갈래, rows in got["자료"].items():
            m = got["메타"].get(갈래, {})
            print("[%s] %s건 · 쓴때 %s" % (갈래, len(rows), m.get("쓴때", "?")))
            for k, v in list(rows.items())[:5]:
                print("   ", k, v["값"])
        return 0
    print(status_text())
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
