# -*- coding: utf-8 -*-
"""
mulder_export.py : Mulder Control Hub 로 1주일에 1회 정리된 쿠팡 업무 DB 를 낸다
==================================================================================
형님 지시(2026-11-17): "멀더 컨트롤 허브로 1주일에 1회씩 정리된 db 보내자 앞으로 쭉"

  python mulder_export.py            # 지금 한 벌 만든다
  python mulder_export.py --status   # 마지막으로 낸 것만 본다

무엇을 내나 (읽기 전용 사본 · SQLite 한 파일)
  · 업무_<종류>  : 앱 DB work_item + work_field 를 업무 종류별 넓은 표로 편 것
                    (돌발AS · 정기점검 · 신규납품설치 · 정산 · 세금계산서 · PO발주 · 현장작업)
  · 리모컨_출고 · 리모컨_재고 · 리모컨_불출
  · 집계        : 종류·상태별 건수
  · _meta       : 만든 때 · 원본 · 표마다 행 수 · 주의

계약
  ★ 우리 정본(db/app_store.db · db/ledger_queue.db)은 **읽기만** 한다.
  ★ 해석하지 않고 값을 그대로 낸다 : 빈칸이 '없음'인지 '안 봄'인지 우리도 모른다([169]).
    대표적인 것을 _meta.주의 에 적는다.
  ★ 비밀값(config/*.json)은 안 넣는다. 지운 업무(deleted_at)는 뺀다.
  ★ 자리 : Documents/_공유DB/쿠팡/쿠팡업무_YYYYMMDD.db (+ 같은 이름 .json 요약 · 읽는법.md)
    최근 2벌만 남긴다. 자리를 옮기려면 ULH_SHARE_DB_DIR 하나.
  ★ 회차 : 작업 스케줄러 `쿠팡업무_Mulder주간DB` (월 07:00 · 숨김). 설치 install_mulder_weekly.ps1
"""
import os, sys, json, sqlite3, glob
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
APP_DB = os.path.join(ROOT, "db", "app_store.db")
Q_DB = os.path.join(ROOT, "db", "ledger_queue.db")
SHARE_ROOT = os.environ.get("ULH_SHARE_DB_DIR") or os.path.join(
    os.path.dirname(os.path.dirname(ROOT)), "_공유DB")
OUT_DIR = os.path.join(SHARE_ROOT, "쿠팡")
KEEP = 2

NOTES = [
    "정산 표의 세금계산서합계·미수금액·입금차액은 대부분 비어 있다(관리대장 수식 재료가 없음). 0/빈칸을 '미수 없음'으로 읽지 말 것.",
    "입금액은 앱 DB 기준이라 쓸 수 있다.",
    "리모컨 재고는 2026-08-19 증평→부산 30개 이동이 재고표에 빠져 있어 부산 VER.4 가 음수로 보인다(불출 표 id12 참고).",
    "업무 값은 앱 DB(정본) 그대로다. 관리대장 엑셀 보관본과 다를 수 있다 : 정본은 이쪽이다.",
]


def _ro(path):
    return sqlite3.connect("file:%s?mode=ro" % path.replace("\\", "/"), uri=True, timeout=30)


def _val(raw):
    if raw is None:
        return None
    try:
        v = json.loads(raw)
    except Exception:
        return raw
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def build(path):
    out = sqlite3.connect(path)
    meta = {"만든때": datetime.now().isoformat(timespec="seconds"),
            "원본": "CSOS 쿠팡 통합업무 앱 DB(읽기 전용 사본)", "표": {}, "주의": NOTES}
    src = _ro(APP_DB)
    try:
        items = src.execute(
            "SELECT id, kind, public_id, project_no, camp_name, status, updated_at"
            "  FROM work_item WHERE deleted_at IS NULL").fetchall()
        fields = {}
        for wid, k, v in src.execute("SELECT work_id, field_key, value_json FROM work_field"):
            fields.setdefault(wid, {})[k] = _val(v)
    finally:
        src.close()
    by_kind = {}
    for it in items:
        by_kind.setdefault(it[1], []).append(it)
    base = ["앱_공개ID", "프로젝트NO", "캠프명", "앱_상태", "앱_갱신시각"]
    for kind, rows in sorted(by_kind.items()):
        keys = []
        for r in rows:
            for k in fields.get(r[0], {}):
                if k not in keys and k not in base:
                    keys.append(k)
        cols = base + keys
        t = "업무_" + kind
        out.execute('CREATE TABLE "%s" (%s)' % (t, ",".join('"%s"' % c for c in cols)))
        data = []
        for r in rows:
            f = fields.get(r[0], {})
            data.append([r[2], r[3], r[4], r[5], r[6]] + [f.get(k) for k in keys])
        out.executemany('INSERT INTO "%s" VALUES (%s)' % (t, ",".join("?" * len(cols))), data)
        meta["표"][t] = len(data)
    agg = [(r[1], r[5] or "", 1) for r in items]
    cnt = {}
    for k, s, n in agg:
        cnt[(k, s)] = cnt.get((k, s), 0) + n
    out.execute('CREATE TABLE "집계" ("종류","상태","건수")')
    out.executemany('INSERT INTO "집계" VALUES (?,?,?)', [(k, s, n) for (k, s), n in sorted(cnt.items())])
    meta["표"]["집계"] = len(cnt)

    q = _ro(Q_DB)
    try:
        for src_t, t in (("remote_delivery", "리모컨_출고"), ("remote_stock", "리모컨_재고"),
                         ("remote_issue", "리모컨_불출")):
            cur = q.execute("SELECT * FROM %s ORDER BY id" % src_t)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
            out.execute('CREATE TABLE "%s" (%s)' % (t, ",".join('"%s"' % c for c in cols)))
            out.executemany('INSERT INTO "%s" VALUES (%s)' % (t, ",".join("?" * len(cols))), rows)
            meta["표"][t] = len(rows)
    finally:
        q.close()
    out.execute('CREATE TABLE "_meta" ("열쇠","값")')
    out.executemany('INSERT INTO "_meta" VALUES (?,?)',
                    [(k, json.dumps(v, ensure_ascii=False) if not isinstance(v, str) else v)
                     for k, v in meta.items()])
    out.commit()
    out.close()
    return meta


README = """# 쿠팡 업무 주간 DB : 읽는 법

CSOS(쿠팡 통합업무) 앱이 **매주 월요일 07:00** 에 내는 읽기 전용 사본입니다. Mulder Control Hub 용.

- 파일: `쿠팡업무_YYYYMMDD.db` (SQLite) + 같은 이름 `.json` 요약. 최근 2벌만 남습니다.
- **읽기만** 하십시오. 고쳐도 원본(CSOS 앱 DB)에는 안 들어갑니다.
- 표
  - `업무_<종류>` : 돌발AS · 정기점검 · 신규납품설치 · 정산 · 세금계산서 · PO발주 · 현장작업. 한 줄이 업무 한 건, 칸은 앱 화면 칸 이름 그대로.
  - `리모컨_출고` · `리모컨_재고` · `리모컨_불출`
  - `집계` : 종류·상태별 건수
  - `_meta` : 만든 때 · 표마다 행 수 · **주의**(빈칸을 어떻게 읽어야 하는지)
- 빈칸은 '없음'이 아니라 '기록 안 됨'일 수 있습니다. `_meta.주의` 를 먼저 보십시오.
- 더 필요한 칸이나 형식은 CSOS 앱 코딩 창에 말씀해 주십시오.
"""


def run():
    os.makedirs(OUT_DIR, exist_ok=True)
    day = datetime.now().strftime("%Y%m%d")
    final = os.path.join(OUT_DIR, "쿠팡업무_%s.db" % day)
    tmp = final + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)
    meta = build(tmp)
    os.replace(tmp, final)                      # 반쯤 쓴 파일을 남이 읽지 않게
    with open(final[:-3] + ".json", "w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False, indent=1)
    with open(os.path.join(OUT_DIR, "읽는법.md"), "w", encoding="utf-8") as fh:
        fh.write(README)
    old = sorted(glob.glob(os.path.join(OUT_DIR, "쿠팡업무_*.db")))[:-KEEP]
    for p in old:
        for q in (p, p[:-3] + ".json"):
            try:
                os.remove(q)
            except OSError:
                pass
    print("Mulder 주간 DB → %s" % final)
    print("   " + " · ".join("%s %d" % (k, v) for k, v in meta["표"].items()))
    return final, meta


def status():
    files = sorted(glob.glob(os.path.join(OUT_DIR, "쿠팡업무_*.db")))
    if not files:
        print("아직 낸 것이 없습니다 : %s" % OUT_DIR)
        return
    for p in files:
        print(p, os.path.getsize(p))


if __name__ == "__main__":
    if "--status" in sys.argv:
        status()
    else:
        run()
