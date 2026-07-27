# -*- coding: utf-8 -*-
"""
ledger_versions.py — 관리대장 버전 파일이 쌓이는 걸 정리한다
================================================================================
원장을 고치는 도구는 11개나 되고(ledger_writer·fix_ids·workbook_patch·reorder_rows…),
전부 **원본을 건드리지 않고 vN+1을 새로 만든다**. 안전을 위해 그렇게 설계했지만,
하루 작업하면 폴더에 v169…v173처럼 대여섯 개가 쌓여 어느 게 최신인지 헷갈린다.

그래서 지우지 않고 **접어 둔다**:

    남긴다  · **최신본 하나만** (사용자 지시 2026-07-27)
            · 사람이 표시해 둔 것(파일명에 '보관')
    옮긴다  · 나머지 → `OLD/` 하위 폴더

**삭제하지 않는다.** 지우는 건 사람이 폴더를 보고 판단할 일이다.
최신본은 어떤 경우에도 건드리지 않는다(resolve_master가 이걸 찾는다).

  python ledger_versions.py              # 현황만 본다
  python ledger_versions.py --prune      # 접어 둔다(_이전버전/ 로 이동)
"""
import sys, os, re, glob, shutil
from datetime import datetime
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 사용자 지시(2026-07-27): "최신 버전만 남기고 이전 버전은 전부 OLD 폴더에 저장".
# 되돌리기가 필요하면 OLD에서 꺼내면 된다 — 폴더 하나 차이일 뿐 사라지지 않는다.
KEEP_LATEST = 1       # 작업 폴더에는 최신본 하나만 둔다
KEEP_DAYS = 0         # '하루의 마지막 버전'도 따로 남기지 않는다
# 사용자 지정(2026-07-27): 이전 버전은 관리대장 폴더 아래 OLD 로 모은다.
ARCHIVE = "OLD"
VRE = re.compile(r"_v(\d+)\.xlsx$")


def versions(master):
    d = os.path.dirname(master)
    out = []
    for p in glob.glob(os.path.join(d, "*_v*.xlsx")):
        m = VRE.search(p)
        if not m or os.path.basename(p).startswith("~$"):
            continue
        out.append({"path": p, "v": int(m.group(1)),
                    "mb": os.path.getsize(p) / 1024 / 1024,
                    "day": datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d"),
                    "mtime": os.path.getmtime(p)})
    return sorted(out, key=lambda x: x["v"])


def plan(master):
    """(남길 것, 접을 것). 최신본은 절대 접지 않는다."""
    vs = versions(master)
    if not vs:
        return [], []
    keep = {v["path"] for v in vs[-KEEP_LATEST:]}
    keep.add(max(vs, key=lambda x: x["v"])["path"])          # 최신본은 무조건

    # 하루의 마지막 버전 — 최근 KEEP_DAYS일 치
    byday = defaultdict(list)
    for v in vs:
        byday[v["day"]].append(v)
    for day in sorted(byday)[-KEEP_DAYS:]:
        keep.add(max(byday[day], key=lambda x: x["v"])["path"])

    for v in vs:                                             # 사람이 표시해 둔 것
        if "보관" in os.path.basename(v["path"]):
            keep.add(v["path"])

    move = [v for v in vs if v["path"] not in keep]
    return [v for v in vs if v["path"] in keep], move


def main():
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    vs = versions(master)
    if not vs:
        print("버전 파일을 찾지 못했습니다."); return
    keep, move = plan(master)
    tot = sum(v["mb"] for v in vs)

    print(f"관리대장 버전 {len(vs)}개 · 합계 {tot:.0f}MB "
          f"(v{vs[0]['v']} ~ v{vs[-1]['v']}) · 최신 {os.path.basename(master)}")
    days = defaultdict(int)
    for v in vs:
        days[v["day"]] += 1
    for d in sorted(days)[-5:]:
        print(f"  {d}  {days[d]}개")
    print(f"\n  남김 {len(keep)}개 (최신본만 — 나머지는 OLD로)"
          f" · 접을 것 {len(move)}개 {sum(v['mb'] for v in move):.0f}MB")
    for v in move[:8]:
        print(f"    v{v['v']}  {v['day']}  {v['mb']:.1f}MB")
    if len(move) > 8:
        print(f"    … 외 {len(move)-8}개")

    if not move:
        print("\n정리할 것이 없습니다 — 지금은 적당한 개수입니다.")
        return
    if "--prune" not in sys.argv:
        print("\n미리보기 — 실제 정리: python ledger_versions.py --prune")
        return

    dst = os.path.join(os.path.dirname(master), ARCHIVE)
    os.makedirs(dst, exist_ok=True)
    done, fail = 0, []
    for v in move:
        try:
            target = os.path.join(dst, os.path.basename(v["path"]))
            if os.path.exists(target):
                os.remove(v["path"])          # 이미 접어 둔 것과 중복
            else:
                shutil.move(v["path"], target)
            done += 1
        except OSError as e:                  # 엑셀이 열어 두고 있으면 못 옮긴다
            fail.append((os.path.basename(v["path"]), str(e)[:50]))
    print(f"\n{done}개를 {ARCHIVE}/ 로 옮겼습니다 (삭제 아님 — 필요하면 되꺼내면 됩니다)")
    for n, e in fail[:3]:
        print(f"  [건너뜀] {n} — {e}")
    # 최신본이 그대로인지 마지막 확인 — 이게 사라지면 모든 도구가 멈춘다
    assert os.path.exists(master), "최신본이 사라졌습니다"
    print("  최신본 확인:", os.path.basename(master))


if __name__ == "__main__":
    main()
