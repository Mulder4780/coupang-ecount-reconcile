# -*- coding: utf-8 -*-
"""보관본 창고가 무한정 커진다 — 재고 말하고, 지우는 것은 사람이 정한다 (2026-08-27).

형님 물음: **"이 폴더 용량이 왜이렇게 커?"**

실측 2026-08-27: `COUPANG_INTEGRATED_WORK_AGENT` 가 **285 GB** 인데 그중
`ecount/tmp/archive_spool/exports` 하나가 **280.9 GB**(보관본 530개)다.
디스크는 931 GB 중 **110 GB(11.8%)** 만 남아 있었다 — 곧 이 창고가 디스크의 **30%**다.

★ **왜 이렇게 됐나** — 보관본 하나에 앱 DB 통째(`db-snapshot.sqlite3`)가 들어가
  **210~360 MB** 인데 **지우는 코드가 한 줄도 없었다**(실측 `prune`·`retention` 0곳).
  `archive_export._export_once` 는 실패하면 `shutil.rmtree(stage)` 로 치우지만,
  **성공한 것을 정리하는 자리는 아무 데도 없다.**
  8/10~8/23 은 하루 20~30 GB 씩 늘었다(8/24 부터 주 1회로 바뀌어 크게 줄었다 — [417]).

★ **읽기 전용이 기본이다**(`typo_watch`·`truth_watch` 와 같은 자리). 무엇을 남길지는
  업무 판단이다 — 2026-08-10 정본 규칙의 **백업 관문**이 이 파일들에 걸려 있고,
  되돌리기(rollback) 리허설도 여기를 쓴다. 기계가 정할 일이 아니다.
  `--apply` 를 **사람이 직접 줄 때만** 지운다.

★ **지금 쓰는 것은 절대 안 지운다** — `last-good.json` · `worker-status.json` 이
  가리키는 보관본은 무슨 일이 있어도 남긴다. **그 파일을 못 읽으면 아무것도 안
  지운다**([169] — 모름을 '안 쓴다'로 치면 되돌릴 수 없다).

★ **도는 회차의 것도 안 건드린다** — 잠금(`archive-worker.lock`)이 살아 있거나
  만든 지 `FRESH_HOURS` 안이면 남긴다([171] 이 캐시 갈아끼우기에서 배운 그 규칙).

사람이 보는 명령:
    python archive_spool_prune.py                  # 지금 얼마나 쌓였나 (아무것도 안 지운다)
    python archive_spool_prune.py --keep 10        # 최근 10개만 남기면 얼마가 비나
    python archive_spool_prune.py --keep 10 --apply   # 실제로 지운다
    python archive_spool_prune.py --tmp-only --apply  # 만들다 만 것만 치운다(제일 안전)
"""
import argparse
import json
import os
import shutil
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
SPOOL = os.path.join(BASE, "tmp", "archive_spool")
EXPORTS = os.path.join(SPOOL, "exports")
LOCK = os.path.join(SPOOL, "archive-worker.lock")

KEEP_DEFAULT = 10          # 되돌리기 리허설에 쓸 만큼은 남긴다(주 1회 보관이면 10주치)
FRESH_HOURS = 2.0          # 이보다 새것은 도는 회차의 것일 수 있다
TMP_STALE_HOURS = 2.0      # 만들다 만 것을 '죽은 회차의 찌꺼기'로 보는 나이


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
    """spool 밖은 무슨 일이 있어도 안 건드린다."""
    try:
        a = os.path.realpath(path)
        b = os.path.realpath(root)
        return a == b or a.startswith(b + os.sep)
    except Exception:
        return False


def in_use():
    """지금 쓰는 보관본 이름들. **못 읽으면 None** — 그때는 아무것도 안 지운다([169])."""
    names = set()
    read_any = False
    for fn in ("last-good.json", "worker-status.json"):
        p = os.path.join(SPOOL, fn)
        if not os.path.exists(p):
            continue
        try:
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
        except Exception:
            return None                     # 깨진 파일 — 지금 무엇을 쓰는지 모른다
        read_any = True
        for key in ("artifact_dir", "archive_path", "command_plan"):
            for holder in (d, d.get("export") if isinstance(d.get("export"), dict) else {}):
                v = (holder or {}).get(key)
                if not v:
                    continue
                v = str(v)
                # 파일이면 그 부모가 보관본 폴더다
                cand = v if os.path.basename(v).startswith("exp-") else os.path.dirname(v)
                base = os.path.basename(cand.rstrip("\\/"))
                if base.startswith("exp-"):
                    names.add(base)
    return names if read_any else None


def worker_alive():
    """보관 회차가 지금 도는가. **모르면 '돈다'로 친다**(안전한 쪽)."""
    if not os.path.exists(LOCK):
        return False
    try:
        age = (time.time() - os.path.getmtime(LOCK)) / 3600.0
    except OSError:
        return True
    return age < FRESH_HOURS


def survey():
    """창고를 재기만 한다. 아무것도 안 고친다."""
    out = {"있음": os.path.isdir(EXPORTS), "정식": [], "임시": [], "합": 0,
           "쓰는것": in_use(), "회차중": worker_alive()}
    if not out["있음"]:
        return out
    now = time.time()
    for name in os.listdir(EXPORTS):
        p = os.path.join(EXPORTS, name)
        try:
            mt = os.path.getmtime(p)
        except OSError:
            continue
        rec = {"이름": name, "크기": _size(p), "시각": mt,
               "나이시간": (now - mt) / 3600.0}
        (out["임시"] if name.endswith(".tmp") else out["정식"]).append(rec)
        out["합"] += rec["크기"]
    out["정식"].sort(key=lambda r: r["시각"])
    out["임시"].sort(key=lambda r: r["시각"])
    return out


def plan(keep=KEEP_DEFAULT, tmp_only=False):
    """무엇을 지울지 고른다 — **고르기만** 한다.

    ★ 모르면 아무것도 안 고른다([169]).
    """
    s = survey()
    if not s["있음"]:
        return {"지울것": [], "왜못함": "보관본 창고가 없습니다", "잰것": s}
    if s["쓰는것"] is None:
        return {"지울것": [], "왜못함":
                "지금 어느 보관본을 쓰는지 못 읽었습니다(last-good.json·worker-status.json)"
                " — 모르면 아무것도 안 지웁니다", "잰것": s}
    if s["회차중"]:
        return {"지울것": [], "왜못함":
                "보관 회차가 도는 중입니다 — 끝난 뒤에 다시 봅니다", "잰것": s}

    victims = []
    # ① 만들다 만 것 — 코드가 원래 지우기로 한 것이다(죽어서 못 지웠을 뿐).
    for r in s["임시"]:
        if r["나이시간"] >= TMP_STALE_HOURS:
            victims.append(dict(r, 갈래="만들다 만 것"))
    # ② 오래된 정식 보관본 — **최근 keep 개는 남긴다**
    if not tmp_only:
        old = s["정식"][:-keep] if keep > 0 else list(s["정식"])
        for r in old:
            if r["이름"] in s["쓰는것"]:
                continue                    # 지금 쓰는 것은 절대 안 지운다
            if r["나이시간"] < FRESH_HOURS:
                continue                    # 방금 만든 것도 안 건드린다
            victims.append(dict(r, 갈래="오래된 보관본"))
    return {"지울것": victims, "왜못함": "", "잰것": s}


def apply(victims):
    """실제로 지운다. **spool 밖은 절대 안 건드린다.**"""
    freed = 0
    gone = 0
    failed = []
    for v in victims:
        p = os.path.join(EXPORTS, v["이름"])
        if not _within(p, EXPORTS):
            failed.append((v["이름"], "창고 밖 경로"))
            continue
        try:
            if os.path.isdir(p):
                shutil.rmtree(p)
            else:
                os.remove(p)
            freed += v["크기"]
            gone += 1
        except Exception as e:
            failed.append((v["이름"], "%s: %s" % (type(e).__name__, str(e)[:60])))
    return freed, gone, failed


def _gb(n):
    return n / 1024.0 ** 3


def notice(keep=KEEP_DEFAULT):
    """인계 '먼저 처리할 것' 에 올릴 한 줄 — **넘칠 때만** 말한다([170])."""
    p = plan(keep=keep)
    if p["왜못함"]:
        return None                          # 모르면 아무 말도 안 한다([169])
    v = p["지울것"]
    if not v:
        return None
    free = sum(x["크기"] for x in v)
    if _gb(free) < 5:
        return None                          # 몇 GB 로는 말 안 한다 — 매일 뜨면 아무도 안 본다
    return ("보관본 창고가 %.0f GB 다 — 지금 쓰는 것 말고 %d개(%.0f GB)를 지울 수 있다"
            " (지우는 것은 사람이 정한다)" % (_gb(p["잰것"]["합"]), len(v), _gb(free)),
            "python archive_spool_prune.py --keep %d          # 먼저 보기 · 실제로 지우려면 --apply" % keep)


def main(argv=None):
    ap = argparse.ArgumentParser(description="보관본 창고 재기·정리(기본은 읽기 전용)")
    ap.add_argument("--keep", type=int, default=KEEP_DEFAULT, help="최근 몇 개를 남길까")
    ap.add_argument("--tmp-only", action="store_true", help="만들다 만 것만 치운다")
    ap.add_argument("--apply", action="store_true", help="실제로 지운다(사람이 직접 줄 때만)")
    a = ap.parse_args(argv)

    p = plan(keep=a.keep, tmp_only=a.tmp_only)
    s = p["잰것"]
    if not s["있음"]:
        print("보관본 창고가 없습니다:", EXPORTS)
        return 0

    print("보관본 창고: %s" % EXPORTS)
    print("  정식 %d개 · 만들다 만 것 %d개 · 합 %.1f GB"
          % (len(s["정식"]), len(s["임시"]), _gb(s["합"])))
    if s["쓰는것"]:
        print("  지금 쓰는 것: %s" % ", ".join(sorted(s["쓰는것"])))
    try:
        t, _u, f = shutil.disk_usage(BASE)
        print("  디스크: 남은 것 %.0f GB / %.0f GB (%.1f%%)" % (_gb(f), _gb(t), f / t * 100))
    except Exception:
        pass

    if p["왜못함"]:
        print("\n지울 것을 못 골랐습니다 — %s" % p["왜못함"])
        return 0

    v = p["지울것"]
    if not v:
        print("\n지울 것이 없습니다.")
        return 0

    free = sum(x["크기"] for x in v)
    tmp_n = sum(1 for x in v if x["갈래"] == "만들다 만 것")
    print("\n지울 수 있는 것: %d개 · %.1f GB (만들다 만 것 %d · 오래된 보관본 %d)"
          % (len(v), _gb(free), tmp_n, len(v) - tmp_n))
    for x in v[:5]:
        print("   %8.0f MB · %5.0f시간 전 · %s · %s"
              % (x["크기"] / 1024.0 ** 2, x["나이시간"], x["갈래"], x["이름"][:52]))
    if len(v) > 5:
        print("   … 그 밖 %d개" % (len(v) - 5))

    if not a.apply:
        print("\n아무것도 안 지웠습니다 — 실제로 지우려면 뒤에 --apply 를 붙입니다.")
        return 0

    freed, gone, failed = apply(v)
    print("\n지웠습니다: %d개 · %.1f GB" % (gone, _gb(freed)))
    if failed:
        print("  못 지운 것 %d개:" % len(failed))
        for name, why in failed[:5]:
            print("   ", name[:50], "—", why)
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
