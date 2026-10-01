# -*- coding: utf-8 -*-
"""
ops_backup.py — **모든 DB와 그 과거 기록**을 형님이 지정한 폴더에 백업한다
===============================================================================
형님 지시(2026-10-02): *"모든건 메뉴얼로 만들고 메뉴얼대로 진행해 내가 지정한 폴더에
만들고 로그관리나 오류 등 전부 문서화해서 관리하고 그 문서 규칙에 따라 모든 db도
과거 기록도 다 남기고 내가 지정한 폴더에 백업하도록해"*
형님이 고르신 자리: `<2. CSOS DATA>/12. 운영 매뉴얼`

## 왜 새로 만들었나 — 재고 나서다([67])
실측 2026-10-02: `archive_keep` 는 **`ledger_queue.db` 하나만** 백업한다.
업무 정본인 **`app_store.db`(2.4GB · 돌발AS·정기점검·정산 전부)는 어디에도 백업되지
않았다.** `datalake.db`(303MB) · `pm_content.db` 도 마찬가지다. 그러니 이것은
있는 것을 고치는 일이 아니라 **없던 것을 만드는 일**이다([172]).

## 어떻게 뜨나 — 복사가 아니라 **SQLite 가 떠 주는 일관 백업**
그냥 파일을 복사하면 **쓰는 중에 찢긴 사본**이 나온다(WAL 이 따로 놀아 복구가 안 된다).
`sqlite3.Connection.backup()` 은 DB 가 열려 있어도 **한 시점의 온전한 사본**을 만든다.

## 무엇을 남기나
· `04_DB백업/YYYY-MM-DD/<이름>.db.gz`  — 압축본(기록이 JSON 이라 잘 줄어든다)
· `04_DB백업/YYYY-MM-DD/요약.json`     — 만든 시각 · 원본 · 바이트 · sha256 · 표별 행수
**비밀키(config/*.json)는 한 글자도 안 넣는다**(절대규칙 1). DB 만 뜬다.

## 같은 내용이면 다시 안 뜬다([168])
어제치와 sha256 이 같으면 **건너뛰고 그 사실을 적는다.** 2.4GB 를 날마다 쌓으면
한 달에 72GB 다 — 그러면 공유 서버가 먼저 죽는다.

## 세대 보관 — 판정은 빌린다([162])
`archive_keep.keep_days` 를 그대로 쓴다(최근 14일은 매일 · 그 앞은 달마다 하루).
여기서 다시 정하면 두 곳이 갈리고, 갈린 뒤엔 어느 쪽이 맞는지 아무도 모른다.

## Z: 가 끊기면 — **버리지 않고 쌓아 둔다**([169])
못 닿으면 로컬 대기함(`reports/_백업대기/`)에 뜨고, 다음 회차가 Z: 로 올린다.
*"못 닿았다"* 를 적지 **"백업했다"고 적지 않는다** — 그것이 가장 위험한 거짓이다.

## 사람이 쓰는 법
  python ops_backup.py              # 오늘치 백업 + 밀린 것 올림 + 세대 정리
  python ops_backup.py --print      # 무엇이 있나만 본다(아무것도 안 쓴다)
  python ops_backup.py --dry        # 할 일만 말하고 안 쓴다
  python ops_backup.py --restore <날짜> <이름> --out <경로>   # 되살린다

## 되살리는 법은 매뉴얼에 있다
`01_매뉴얼/MAN-001_백업과_복구.md` — 백업은 **되살려 본 적이 있을 때만** 백업이다.
"""
import argparse
import datetime
import glob
import gzip
import hashlib
import io
import json
import os
import shutil
import sqlite3
import sys
import tempfile

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

DB_DIR = os.path.join(BASE, "db")
WAIT_DIR = os.path.join(BASE, "reports", "_백업대기")
MANUAL_DIRNAME = "12. 운영 매뉴얼"
BACKUP_SUB = "04_DB백업"
LOG_SUB = "02_로그"
ERR_SUB = "03_오류"

# 백업할 DB — **손으로 적지 않는다**([340]). db/ 를 훑어 만든다.
SKIP_SUFFIX = ("-wal", "-shm", ".tmp", ".part")


MARK = os.path.join(BASE, "reports", ".DB백업_도장.json")


def 오늘했나(today=None):
    """오늘 이미 떴나 — **도장은 이 PC 에 찍는다**(Z: 가 끊겨도 읽힌다).

    ★ 못 읽으면 '안 했다'로 친다([169] 를 이 자리에 맞게 정한 것).
      한 번 더 뜨는 값은 13분이고, 안 떠서 그날 백업이 없는 값은 **되돌릴 수 없다.**
    """
    today = today or datetime.date.today().isoformat()
    try:
        d = json.load(io.open(MARK, encoding="utf-8"))
        return d.get("마지막") == today, d
    except Exception:
        return False, {}


def 도장찍기(today, 말, 대기):
    try:
        os.makedirs(os.path.dirname(MARK), exist_ok=True)
        with io.open(MARK, "w", encoding="utf-8") as fh:
            json.dump({"마지막": today, "적은때": datetime.datetime.now().isoformat(
                timespec="seconds"), "한말": 말, "대기함에떴나": 대기},
                fh, ensure_ascii=False, indent=1)
    except Exception:
        pass            # 도장을 못 찍어도 백업을 무르지 않는다


def manual_root():
    """형님이 지정하신 자리. 못 찾으면 None — 지어내지 않는다([169])."""
    try:
        import source_dirs
        csos = getattr(source_dirs, "CSOS_DATA_ROOT", None)
        if csos:
            return os.path.join(csos, MANUAL_DIRNAME)
    except Exception:
        pass
    return None


def db_files():
    """백업 대상 — 코드에게 묻는다([340])."""
    out = []
    for p in sorted(glob.glob(os.path.join(DB_DIR, "*.db"))):
        n = os.path.basename(p)
        if any(n.endswith(s) for s in SKIP_SUFFIX):
            continue
        out.append(p)
    return out


def table_rows(path):
    """표마다 몇 줄인가 — 되살린 뒤 **같은지 대 볼 근거**다."""
    out = {}
    try:
        c = sqlite3.connect("file:%s?mode=ro" % path.replace("\\", "/"), uri=True)
        for (t,) in c.execute("select name from sqlite_master where type='table'"):
            try:
                out[t] = c.execute('select count(*) from "%s"' % t).fetchone()[0]
            except Exception:
                out[t] = None          # 못 센 것을 0 으로 적지 않는다([169])
        c.close()
    except Exception as e:
        return {"__못읽음": str(e)[:200]}
    return out


def consistent_copy(src, dst):
    """SQLite 가 떠 주는 **한 시점의 온전한 사본**. 파일 복사가 아니다."""
    s = sqlite3.connect("file:%s?mode=ro" % src.replace("\\", "/"), uri=True)
    try:
        d = sqlite3.connect(dst)
        try:
            s.backup(d)
        finally:
            d.close()
    finally:
        s.close()


def sha256_of(path, chunk=1 << 20):
    h = hashlib.sha256()
    with io.open(path, "rb") as fh:
        while True:
            b = fh.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def gz_to(src, dst):
    with io.open(src, "rb") as fi, gzip.open(dst, "wb", compresslevel=6) as fo:
        shutil.copyfileobj(fi, fo, 1 << 20)


def 어제까지의지문(root):
    """이미 떠 둔 것의 sha256 — 같은 내용을 또 안 뜨기 위해서다([168]).

    ★ **대기함도 같이 본다.** 안 보면 Z: 가 끊긴 날마다 2.4GB 를 새로 뜬다 —
      내용이 어제와 같은데도. 2026-10-02 에 그대로 밟았다.
    """
    seen = {}
    bases = [os.path.join(root, BACKUP_SUB)] if root else []
    bases.append(WAIT_DIR)                       # 대기함이 **나중**이다 = 더 최신
    for base in bases:
        if not os.path.isdir(base):
            continue
        for day in sorted(os.listdir(base)):
            p = os.path.join(base, day, "요약.json")
            try:
                d = json.load(io.open(p, encoding="utf-8"))
            except Exception:
                continue
            for name, info in (d.get("DB") or {}).items():
                k = info.get("sha256")
                if k:
                    seen[name] = (k, day, info.get("겉모습"))
    return seen


def 겉모습(src):
    """**뜨지 않고** 바뀌었는지 가늠할 값 — 본체와 `-wal` 의 크기·수정시각.

    ★ 지문(sha256)은 2.4GB 를 통째로 떠야 나온다(실측 13분 30초). 그러니 그 전에
      싸게 한 번 거른다([168]). 이것은 **같다는 증거이지 다르다는 증거가 아니다** —
      어긋나면 그때 진짜로 떠서 지문으로 판정한다.
    """
    out = {}
    for suf in ("", "-wal"):
        p = src + suf
        try:
            st = os.stat(p)
            out[suf or "본체"] = [st.st_size, int(st.st_mtime)]
        except OSError:
            out[suf or "본체"] = None
    return out


def one_backup(src, out_dir, 이미, dry=False):
    """DB 하나를 뜬다. 돌려주는 것: 사람이 읽을 한 줄 + 기록."""
    name = os.path.basename(src)
    겉 = 겉모습(src)
    before = 이미.get(name)
    # ★ 겉모습이 그대로면 **뜨지도 않는다** — 2.4GB 를 복사·압축하는 13분을 아낀다.
    if before and before[2] and before[2] == 겉 and not dry:
        return ("그대로", {"이름": name, "sha256": before[0], "건너뜀": True,
                         "같은날": before[1], "겉모습": 겉, "표": None,
                         "어떻게": "겉모습(크기·수정시각)이 그대로라 뜨지 않았다",
                         "원본바이트": os.path.getsize(src)})
    # ★ 미리보기는 **뜨지 않는다**([168]). 2.4GB 를 떠서 지문까지 내면 미리보기가
    #   진짜 백업만큼 걸린다 — 실측 2026-10-02 에 `--dry` 가 2분을 넘겼다.
    #   그래서 지문을 못 내는데, **그것을 '바뀌었다'로 적지 않는다**([169]).
    if dry:
        return ("뜰것", {"이름": name, "sha256": None, "표": table_rows(src),
                        "겉모습": 겉,
                        "원본바이트": os.path.getsize(src),
                        "왜모르나": "미리보기라 지문을 안 냈다 — 같은 내용인지는 모른다"})
    tmpd = tempfile.mkdtemp(prefix="ops_bak_")
    try:
        raw = os.path.join(tmpd, name)
        consistent_copy(src, raw)
        key = sha256_of(raw)
        rows = table_rows(raw)
        if before and before[0] == key:
            return ("그대로", {"이름": name, "sha256": key, "건너뜀": True,
                             "같은날": before[1], "표": rows, "겉모습": 겉,
                             "어떻게": "떠 보니 지문이 같았다",
                             "원본바이트": os.path.getsize(src)})
        os.makedirs(out_dir, exist_ok=True)
        dst = os.path.join(out_dir, name + ".gz")
        gz_to(raw, dst + ".part")
        os.replace(dst + ".part", dst)
        return ("떴다", {"이름": name, "sha256": key, "표": rows, "겉모습": 겉,
                        "원본바이트": os.path.getsize(src),
                        "압축바이트": os.path.getsize(dst),
                        "파일": os.path.basename(dst)})
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def run(root=None, today=None, dry=False, once=False):
    """오늘치 백업. 돌려주는 것: 사람이 읽을 한 줄.

    ★ `once=True` 면 **하루 한 번**이다. 실측 2026-10-02 에 2,739MB 를 뜨는 데
      13분 30초가 걸렸다 — 30분 회차에 그냥 달면 매번 예산을 먹는다.
      그래서 회차는 이 깃발로 부르고, 이미 떴으면 **한 줄로 바로 끝난다.**
    """
    today = today or datetime.date.today().isoformat()
    if once:
        했나, d = 오늘했나(today)
        if 했나:
            꼬리 = " (그때 공유 폴더에 못 닿아 대기함에 떴다)" if d.get("대기함에떴나") else ""
            return "DB 백업: 오늘 이미 떴다%s" % 꼬리, {"건너뜀": True, "도장": d}
    root = root if root is not None else manual_root()
    닿음 = bool(root) and os.path.isdir(os.path.dirname(root or "") or ".")
    if root and not os.path.isdir(root):
        try:
            if not dry:
                os.makedirs(root, exist_ok=True)
        except Exception:
            닿음 = False
    if root and not os.path.isdir(root) and not dry:
        닿음 = False
    대기 = not 닿음
    목적 = (WAIT_DIR if 대기 else os.path.join(root, BACKUP_SUB))
    out_dir = os.path.join(목적, today)

    이미 = 어제까지의지문(root if not 대기 else None)
    기록 = {"만든시각": datetime.datetime.now().isoformat(timespec="seconds"),
           "자리": out_dir, "대기함": 대기, "DB": {}}
    말 = []
    for p in db_files():
        갈래, info = one_backup(p, out_dir, 이미, dry=dry)
        기록["DB"][info["이름"]] = info
        if 갈래 == "그대로":
            말.append("%s 그대로(%s)" % (info["이름"], info["같은날"]))
        elif 갈래 == "뜰것":
            말.append("%s 뜰것" % info["이름"])
        else:
            말.append("%s %.0fMB→%.0fMB"
                      % (info["이름"], info["원본바이트"] / 1048576.0,
                         info["압축바이트"] / 1048576.0))
    if not dry and 기록["DB"]:
        os.makedirs(out_dir, exist_ok=True)
        with io.open(os.path.join(out_dir, "요약.json"), "w", encoding="utf-8") as fh:
            json.dump(기록, fh, ensure_ascii=False, indent=1)

    머리 = "DB 백업"
    if 대기:
        머리 += " — **공유 폴더에 못 닿아 이 PC 대기함에 떴다**(다음 회차가 올린다)"
    한줄 = "%s: %s" % (머리, " · ".join(말) or "뜰 DB가 없다")
    if not dry and 기록["DB"]:
        도장찍기(today, 한줄, 대기)
    return 한줄, 기록


def 밀린것올리기(root=None, dry=False):
    """대기함에 쌓인 것을 공유 폴더로 올린다. 못 닿으면 그대로 둔다([169])."""
    root = root if root is not None else manual_root()
    if not root or not os.path.isdir(root):
        n = len(glob.glob(os.path.join(WAIT_DIR, "*")))
        return "밀린 백업 %d벌 — 공유 폴더에 아직 못 닿았다" % n if n else ""
    옮김 = 0
    for day_dir in sorted(glob.glob(os.path.join(WAIT_DIR, "*"))):
        if not os.path.isdir(day_dir):
            continue
        dst = os.path.join(root, BACKUP_SUB, os.path.basename(day_dir))
        if dry:
            옮김 += 1
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if os.path.isdir(dst):
            for f in os.listdir(day_dir):
                shutil.move(os.path.join(day_dir, f), os.path.join(dst, f))
            shutil.rmtree(day_dir, ignore_errors=True)
        else:
            shutil.move(day_dir, dst)
        옮김 += 1
    return "밀린 백업 %d벌 올림" % 옮김 if 옮김 else ""


def 세대정리(root=None, today=None, dry=False):
    """오래된 벌을 솎는다 — **판정은 archive_keep 에서 빌린다**([162])."""
    root = root if root is not None else manual_root()
    if not root or not os.path.isdir(os.path.join(root, BACKUP_SUB)):
        return ""
    try:
        import archive_keep as A
    except Exception as e:
        return "세대 정리 못함(판정을 못 빌렸다: %s)" % str(e)[:60]
    base = os.path.join(root, BACKUP_SUB)
    days = []
    for n in os.listdir(base):
        d = A.parse_day(n)
        if d:
            days.append(d)
    if not days:
        return ""
    t = datetime.date.fromisoformat(today) if today else None
    남길 = A.keep_days(days, today=t)
    지움 = [d for d in days if d not in 남길]
    if dry:
        return "세대 정리: %d벌 지울 것" % len(지움) if 지움 else ""
    for d in 지움:
        shutil.rmtree(os.path.join(base, d.isoformat()), ignore_errors=True)
    return "세대 정리: %d벌 지움(남은 %d벌)" % (len(지움), len(남길)) if 지움 else ""


def restore(day, name, out, root=None):
    """되살린다. **백업은 되살려 본 적이 있을 때만 백업이다.**"""
    root = root if root is not None else manual_root()
    if not root:
        return "공유 폴더를 못 찾았다"
    src = os.path.join(root, BACKUP_SUB, day, name + ".gz")
    if not os.path.isfile(src):
        src2 = os.path.join(WAIT_DIR, day, name + ".gz")
        if os.path.isfile(src2):
            src = src2
        else:
            return "그 날짜(%s)의 %s 백업이 없다" % (day, name)
    with gzip.open(src, "rb") as fi, io.open(out, "wb") as fo:
        shutil.copyfileobj(fi, fo, 1 << 20)
    rows = table_rows(out)
    요약 = os.path.join(os.path.dirname(src), "요약.json")
    맞나 = "대 볼 요약이 없다"
    try:
        d = json.load(io.open(요약, encoding="utf-8"))
        적힌 = ((d.get("DB") or {}).get(name) or {}).get("표") or {}
        다름 = [t for t in set(적힌) | set(rows) if 적힌.get(t) != rows.get(t)]
        맞나 = "표 %d개 모두 같다" % len(rows) if not 다름 else "다른 표: %s" % 다름[:5]
    except Exception:
        pass
    return "되살림: %s → %s (%s)" % (src, out, 맞나)


def _print():
    root = manual_root()
    print("지정 폴더:", root or "(못 찾음)")
    print("  닿나:", os.path.isdir(root) if root else False)
    base = os.path.join(root, BACKUP_SUB) if root else None
    if base and os.path.isdir(base):
        days = sorted(os.listdir(base))
        print("  백업 %d벌: %s" % (len(days), (days[:3] + ["…"] + days[-2:]) if len(days) > 5 else days))
    n = len([p for p in glob.glob(os.path.join(WAIT_DIR, "*")) if os.path.isdir(p)])
    print("  이 PC 대기함:", n, "벌")
    print("  백업 대상 DB:")
    for p in db_files():
        print("    %-22s %8.1f MB" % (os.path.basename(p), os.path.getsize(p) / 1048576.0))


def main(argv=None):
    ap = argparse.ArgumentParser(description="모든 DB를 지정 폴더에 백업한다")
    ap.add_argument("--print", dest="show", action="store_true", help="무엇이 있나만 본다")
    ap.add_argument("--dry", action="store_true", help="할 일만 말하고 안 쓴다")
    ap.add_argument("--root", help="지정 폴더를 손으로 준다(시험용)")
    ap.add_argument("--today", help="날짜를 손으로 준다(시험용)")
    ap.add_argument("--once", action="store_true",
                    help="하루 한 번만 — 오늘 이미 떴으면 바로 끝낸다(회차가 쓴다)")
    ap.add_argument("--restore", nargs=2, metavar=("날짜", "이름"), help="되살린다")
    ap.add_argument("--out", help="되살릴 자리")
    a = ap.parse_args(argv)

    if a.show:
        _print()
        return 0
    if a.restore:
        if not a.out:
            print("--out 으로 되살릴 자리를 주십시오")
            return 2
        print(restore(a.restore[0], a.restore[1], a.out, root=a.root))
        return 0

    말, _ = run(root=a.root, today=a.today, dry=a.dry, once=a.once)
    print(말)
    for 한줄 in (밀린것올리기(root=a.root, dry=a.dry),
               세대정리(root=a.root, today=a.today, dry=a.dry)):
        if 한줄:
            print(한줄)
    return 0


if __name__ == "__main__":
    sys.exit(main())
