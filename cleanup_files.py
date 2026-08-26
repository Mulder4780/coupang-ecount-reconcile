# -*- coding: utf-8 -*-
"""쓸데없는 파일을 지운다 — 갈래 표 한 곳 · 지운 것은 DB 에 남긴다 (2026-08-27 지시).

형님 지시 둘: **"로그 오류 db 기록등은 모두 db로 기록해놓고 필요없는 파일은 정리해"** ·
**"쓸데 없는 파일들은 지우는 알고리즘 구현해"**

실측 2026-08-27 로 시작했다 — 프로젝트 폴더가 **285 GB** 였고 그중
`tmp/archive_spool/exports` 하나가 **280.9 GB**(보관본 530개)였다.
**지우는 코드가 한 줄도 없었다.** 273.5 GB 를 정리해 디스크가
110 GB → **383 GB**(11.8% → 41.2%)가 됐다.

## 이 파일이 하는 일
* **갈래 표(`RULES`)가 한 곳**이다([162]) — 무엇을·얼마나 남길지·왜 안전한지가 한 줄씩.
  갈래를 늘리려면 여기 한 줄을 더한다. 흩어 놓으면 무엇이 지워지는지 아무도 못 센다.
* **기본은 미리보기**다. `--apply` 를 **사람이 직접 줄 때만** 지운다.
* 지운 것은 **DB(`file_cleanup` 표)에 남긴다** — 형님 지시대로다. 몇 달 뒤
  "그 파일 어디 갔지"를 물을 사람이 그 표를 본다([228]).

## 절대 안 건드리는 것 ([172] — 좁히는 것도 고장이지만, 여기서는 넓히는 것이 되돌릴 수 없다)
* `inbox/` · `outputs/` · `0. 원본 자료`(Z:) — **업무 원본이다.**
* `db/` — 정본이다. 캐시(`source_index_cache.json`)도 안 지운다 — 지우면 다음 색인이
  Z: 를 통째로 다시 훑는다([198]).
* `.git/` · 워크트리(`.claude/worktrees`·`.codex-worktrees`) — 남의 세션 것일 수 있다([104]).
* `band/cache/` — 밴드 캐시는 **되돌릴 수 없다**(가짜 묘비 · 2026-08-19).
* **`*_오류.json`** — `schedule_watch.traces()` 가 글로브로 읽는다([228]·[304]).
  지우면 감시자가 눈이 먼다.
* **`.` 으로 시작하는 상태 파일** — `.daily_run.progress.json` 처럼 회차가 읽는다.

사람이 보는 명령:
    python cleanup_files.py                  # 무엇이 얼마나 쌓였나 (아무것도 안 지운다)
    python cleanup_files.py --apply          # 실제로 지운다
    python cleanup_files.py --only 회차산출물  # 한 갈래만
"""
import argparse
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")

# 이름 안의 **날짜 도장** — 이것이 있으면 그 회차의 산출물이라 옛것은 아무도 안 읽는다.
#   ★ `20\d{6}` 로 좁힌다([172]) — 넓히면 **뜻을 가진 숫자**까지 묶인다.
#     실측: 밴드번호 `84789192`·프로젝트 `UJ2601394` 가 그렇다. 그것들이 한 묶음이 되면
#     "묶음마다 최근 5개" 가 **한 밴드 것만 남기는** 규칙이 되어 남의 자료를 지운다.
#   ★ 묶음 키에서는 **날짜 뒤에 붙는 시각까지** 함께 지운다 — 안 그러면
#     `..._20260810_1102` 와 `..._20260811_1505` 가 서로 다른 묶음이 되어
#     묶음마다 몇 개씩만 쌓이고 **아무것도 안 지워진다**(실측: 333 MB 중 2.9 MB 만 잡혔다).
STAMP = re.compile(r"20[0-9]{6}")
GROUP_STAMP = re.compile(r"20[0-9]{6}([_\-]?[0-9]{2,6})?")

# 회차 산출물이라도 이 이름은 **감시자가 읽는다** — 절대 안 지운다.
NEVER = re.compile(r"_오류\.json$|^\.|세션인계|cloud_continuity|worksplit|ai_claims")

FRESH_DAYS = 7.0        # 이보다 새것은 손대지 않는다
KEEP_PER_GROUP = 5      # 같은 이름 묶음에서 남길 개수


def _size(path):
    if os.path.isfile(path):
        try:
            return os.path.getsize(path)
        except OSError:
            return 0
    total = 0
    for d, _dn, fn in os.walk(path):
        for f in fn:
            try:
                total += os.path.getsize(os.path.join(d, f))
            except OSError:
                pass
    return total


def _within(path, root):
    try:
        a = os.path.realpath(path)
        b = os.path.realpath(root)
        return a == b or a.startswith(b + os.sep)
    except Exception:
        return False


# ── 갈래 ① 보관본 창고 — 판정은 그 도구를 빌린다([162]) ──────────────────
def _rule_spool(keep):
    try:
        import archive_spool_prune as SP
    except Exception as e:
        return [], "보관본 도구를 못 읽었습니다(%s)" % type(e).__name__
    p = SP.plan(keep=10)
    if p["왜못함"]:
        return [], p["왜못함"]
    root = SP.EXPORTS
    return [{"경로": os.path.join(root, v["이름"]), "크기": v["크기"],
             "왜": v["갈래"]} for v in p["지울것"]], ""


# ── 갈래 ② 회차 산출물 — 이름에 날짜 도장이 있고 묶음마다 최근 N개만 남긴다 ──
def _rule_round_output(keep):
    if not os.path.isdir(REPORT_DIR):
        return [], "reports 폴더가 없습니다"
    now = time.time()
    groups = {}
    for d, _dn, fn in os.walk(REPORT_DIR):
        for f in fn:
            if NEVER.search(f):
                continue                       # 감시자가 읽는 것·상태 파일
            if not STAMP.search(f):
                continue                       # 도장이 없으면 '최신'을 가리키는 파일이다
            p = os.path.join(d, f)
            try:
                st = os.stat(p)
            except OSError:
                continue
            if (now - st.st_mtime) / 86400.0 < FRESH_DAYS:
                continue                       # 최근 것은 안 건드린다
            key = (d, GROUP_STAMP.sub("N", f))
            groups.setdefault(key, []).append((st.st_mtime, st.st_size, p))
    victims = []
    for _key, rows in groups.items():
        rows.sort()                            # 오래된 것부터
        for mt, sz, p in rows[:-keep] if keep > 0 else rows:
            victims.append({"경로": p, "크기": sz, "왜": "옛 회차 산출물"})
    return victims, ""


# ── 갈래 ③ 파이썬 캐시 — 언제나 다시 만들어진다 ──────────────────────────
def _rule_pycache(keep):
    victims = []
    skip = (".git", "worktrees", ".codex-worktrees", "node_modules")
    for d, dn, _fn in os.walk(BASE):
        dn[:] = [x for x in dn if x not in skip]
        if os.path.basename(d) == "__pycache__":
            victims.append({"경로": d, "크기": _size(d), "왜": "파이썬 캐시"})
            dn[:] = []
    return victims, ""


# ── 갈래 ④ 찌꺼기 — 죽은 회차가 남긴 임시 파일 ───────────────────────────
def _rule_scrap(keep):
    now = time.time()
    victims = []
    roots = [REPORT_DIR, os.path.join(BASE, "tmp")]
    for root in roots:
        if not os.path.isdir(root):
            continue
        for d, dn, fn in os.walk(root):
            dn[:] = [x for x in dn if x != "exports"]     # 보관본은 갈래 ①이 맡는다
            for f in fn:
                if not (f.endswith(".tmp") or ".stale-" in f):
                    continue
                p = os.path.join(d, f)
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                if (now - st.st_mtime) / 3600.0 < 2.0:
                    continue                   # 지금 만드는 중일 수 있다([171])
                victims.append({"경로": p, "크기": st.st_size, "왜": "죽은 회차의 찌꺼기"})
    return victims, ""


RULES = (
    # (이름, 함수, 얼마나 남길까, 왜 안전한가)
    ("보관본", _rule_spool, 10,
     "지금 쓰는 것(last-good)과 최근 10개는 남는다 — 되돌리기 관문에 쓰인다"),
    ("회차산출물", _rule_round_output, KEEP_PER_GROUP,
     "이름에 날짜 도장이 있는 것만 · 묶음마다 최근 5개 · 7일 안은 안 건드림 · 오류자국 제외"),
    ("파이썬캐시", _rule_pycache, 0, "다시 만들어진다"),
    ("찌꺼기", _rule_scrap, 0, "만들다 만 파일 — 2시간 넘은 것만"),
)


def plan(only=None):
    """무엇을 지울지 고른다 — **고르기만** 한다."""
    out = {"갈래": [], "합": 0, "개수": 0, "못읽음": []}
    for name, fn, keep, why in RULES:
        if only and name != only:
            continue
        try:
            victims, err = fn(keep)
        except Exception as e:
            out["못읽음"].append((name, "%s: %s" % (type(e).__name__, str(e)[:60])))
            continue
        if err:
            out["못읽음"].append((name, err))
            continue
        # 안전: 이 프로젝트 폴더 밖은 무슨 일이 있어도 안 담는다
        victims = [v for v in victims if _within(v["경로"], BASE)]
        s = sum(v["크기"] for v in victims)
        out["갈래"].append({"이름": name, "왜안전": why, "지울것": victims,
                            "크기": s, "개수": len(victims)})
        out["합"] += s
        out["개수"] += len(victims)
    return out


def _remember(rows):
    """지운 것을 **DB 에 남긴다**(형님 지시). 못 남겨도 지운 것을 되돌리지는 않는다."""
    try:
        import ledger_db as L
        with L.conn() as c:
            c.execute("""CREATE TABLE IF NOT EXISTS file_cleanup(
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            ts TEXT, 갈래 TEXT, 경로 TEXT, 바이트 INTEGER, 왜 TEXT)""")
            c.executemany(
                "INSERT INTO file_cleanup(ts,갈래,경로,바이트,왜) VALUES(?,?,?,?,?)",
                [(datetime.now().isoformat(timespec="seconds"), r["갈래"],
                  os.path.relpath(r["경로"], BASE), int(r["크기"]), r["왜"]) for r in rows])
        return len(rows), ""
    except Exception as e:
        return 0, "%s: %s" % (type(e).__name__, str(e)[:80])


def apply(p):
    """실제로 지운다. **프로젝트 폴더 밖은 절대 안 건드린다.**"""
    freed = 0
    gone = 0
    failed = []
    done = []
    for g in p["갈래"]:
        for v in g["지울것"]:
            path = v["경로"]
            if not _within(path, BASE):
                failed.append((path, "프로젝트 밖 경로"))
                continue
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                freed += v["크기"]
                gone += 1
                done.append(dict(v, 갈래=g["이름"]))
            except Exception as e:
                failed.append((path, "%s: %s" % (type(e).__name__, str(e)[:60])))
    noted, note_err = _remember(done) if done else (0, "")
    return freed, gone, failed, noted, note_err


def _mb(n):
    return n / 1024.0 ** 2


def notice():
    """인계 '먼저 처리할 것' 에 올릴 한 줄 — **넘칠 때만** 말한다([170])."""
    try:
        p = plan()
    except Exception:
        return None
    if p["합"] / 1024.0 ** 3 < 5:
        return None
    return ("쓸데없는 파일이 %.0f GB 쌓였다 — %d개(%s)"
            % (p["합"] / 1024.0 ** 3, p["개수"],
               " · ".join("%s %d" % (g["이름"], g["개수"]) for g in p["갈래"] if g["개수"])),
            "python cleanup_files.py            # 먼저 보기 · 실제로 지우려면 --apply")


def main(argv=None):
    ap = argparse.ArgumentParser(description="쓸데없는 파일 정리(기본은 읽기 전용)")
    ap.add_argument("--only", help="한 갈래만 (%s)" % ", ".join(r[0] for r in RULES))
    ap.add_argument("--apply", action="store_true", help="실제로 지운다(사람이 직접 줄 때만)")
    a = ap.parse_args(argv)

    p = plan(only=a.only)
    print("쓸데없는 파일 — 지금 상태")
    for g in p["갈래"]:
        print("  %-10s %5d개 %9.1f MB   (%s)" % (g["이름"], g["개수"], _mb(g["크기"]), g["왜안전"]))
        for v in g["지울것"][:3]:
            print("       %8.1f MB  %s" % (_mb(v["크기"]), os.path.relpath(v["경로"], BASE)[:66]))
        if g["개수"] > 3:
            print("       … 그 밖 %d개" % (g["개수"] - 3))
    for name, why in p["못읽음"]:
        # ★ 조용히 빼지 않는다([169]) — '0개'와 '못 봤다'는 다른 사실이다.
        print("  %-10s 확인 못 함 — %s" % (name, why))
    print("  ------------------------------------------")
    print("  합 %d개 · %.1f MB" % (p["개수"], _mb(p["합"])))
    try:
        t, _u, f = shutil.disk_usage(BASE)
        print("  디스크: 남은 것 %.0f GB / %.0f GB (%.1f%%)"
              % (f / 1024 ** 3, t / 1024 ** 3, f / t * 100))
    except Exception:
        pass

    if not p["개수"]:
        return 0
    if not a.apply:
        print("\n아무것도 안 지웠습니다 — 실제로 지우려면 뒤에 --apply 를 붙입니다.")
        return 0

    freed, gone, failed, noted, note_err = apply(p)
    print("\n지웠습니다: %d개 · %.1f MB" % (gone, _mb(freed)))
    print("  DB 기록: %d줄%s" % (noted, (" — 못 남김: " + note_err) if note_err else ""))
    if failed:
        print("  못 지운 것 %d개:" % len(failed))
        for path, why in failed[:5]:
            print("   ", os.path.relpath(path, BASE)[:50], "—", why)
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
