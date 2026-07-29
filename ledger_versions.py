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
  python ledger_versions.py --prune      # 접어 둔다(지정된 OLD/ 로 이동)
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
LEGACY_ARCHIVES = ("_이전버전",)  # 과거 도구가 쓴 보관 폴더 — 실행 시 OLD로 합친다
# 자동 정리는 **이 이름에만** 손댄다. 넓게 잡으면 남의 파일을 옮긴다(2026-07-28 실사고).
MASTER_RE = re.compile(r"^쿠팡_통합업무_일일보고_관리대장_v(\d+)\.xlsx$")
# 과거 보관본에는 파일명 뒤에 설명이 붙을 수 있다. 관리대장 접두어가 정확한 것만 다룬다.
ARCHIVED_MASTER_RE = re.compile(
    r"^쿠팡_통합업무_일일보고_관리대장_v(\d+).*\.xlsx$"
)


def archive_folder(folder):
    """보관 위치는 관리대장 폴더 바로 아래의 OLD 하나뿐이다."""
    root = os.path.abspath(folder)
    dst = os.path.abspath(os.path.join(root, ARCHIVE))
    if os.path.dirname(dst) != root or os.path.basename(dst) != "OLD":
        raise RuntimeError(f"잘못된 구버전 보관 경로: {dst}")
    return dst


def _collision_safe_target(dst, source, origin="중복"):
    """같은 이름이 있어도 덮어쓰거나 지우지 않고 보존할 새 이름을 만든다."""
    name = os.path.basename(source)
    target = os.path.join(dst, name)
    if not os.path.exists(target):
        return target
    stem, ext = os.path.splitext(name)
    tag = re.sub(r"[^0-9A-Za-z가-힣_-]+", "_", origin).strip("_") or "중복"
    n = 1
    while True:
        suffix = f"__from_{tag}" if n == 1 else f"__from_{tag}_{n}"
        target = os.path.join(dst, stem + suffix + ext)
        if not os.path.exists(target):
            return target
        n += 1


def _move_preserving(source, dst, origin="중복"):
    """한 파일만 OLD로 옮긴다. 기존 파일은 절대 덮어쓰거나 삭제하지 않는다."""
    if not os.path.isfile(source) or os.path.basename(source).startswith("~$"):
        return False
    os.makedirs(dst, exist_ok=True)
    shutil.move(source, _collision_safe_target(dst, source, origin))
    return True


def _legacy_files(folder):
    """예전 보관 폴더에 남은 관리대장만 찾는다(다른 엑셀은 건드리지 않는다)."""
    found = []
    for legacy_name in LEGACY_ARCHIVES:
        legacy = os.path.join(folder, legacy_name)
        if not os.path.isdir(legacy):
            continue
        for p in glob.glob(os.path.join(legacy, "*.xlsx")):
            if ARCHIVED_MASTER_RE.fullmatch(os.path.basename(p)) and not os.path.basename(p).startswith("~$"):
                found.append((p, legacy_name))
    return found


def versions(master):
    d = os.path.dirname(master)
    out = []
    for p in glob.glob(os.path.join(d, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx")):
        m = MASTER_RE.fullmatch(os.path.basename(p))
        if not m or os.path.basename(p).startswith("~$"):
            continue
        try:
            stat = os.stat(p)
        except FileNotFoundError:
            continue  # 다른 프로세스의 자동 정리가 방금 OLD로 옮긴 구버전
        out.append({"path": p, "v": int(m.group(1)),
                    "mb": stat.st_size / 1024 / 1024,
                    "day": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d"),
                    "mtime": stat.st_mtime})
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
    if KEEP_DAYS > 0:
        for day in sorted(byday)[-KEEP_DAYS:]:
            keep.add(max(byday[day], key=lambda x: x["v"])["path"])

    move = [v for v in vs if v["path"] not in keep]
    return [v for v in vs if v["path"] in keep], move


_AUTODONE = False       # 한 프로세스에서 한 번만 — resolve_master 는 여러 번 불린다


def _archive_old_versions(master, quiet=True):
    """최신본을 제외한 현역 구버전과 예전 보관 폴더의 파일을 OLD로 모은다."""
    if not master or not os.path.isfile(master):
        return 0
    current = MASTER_RE.fullmatch(os.path.basename(master))
    if not current:
        return 0

    folder = os.path.dirname(master)
    dst = archive_folder(folder)
    old = []
    for p in glob.glob(os.path.join(folder, "*.xlsx")):
        b = os.path.basename(p)
        m = ARCHIVED_MASTER_RE.fullmatch(b)
        if not m or b.startswith("~$") or os.path.abspath(p) == os.path.abspath(master):
            continue
        if int(m.group(1)) < int(current.group(1)):
            old.append((p, "작업폴더"))

    old.extend(_legacy_files(folder))
    done, failed = 0, []
    for source, origin in old:
        try:
            if _move_preserving(source, dst, origin):
                done += 1
        except OSError as exc:
            failed.append((os.path.basename(source), str(exc)[:80]))

    if done and not quiet:
        print(f"i 구 버전 {done}개를 {ARCHIVE}/ 로 옮겼습니다")
    if failed and not quiet:
        for name, why in failed[:5]:
            print(f"  [건너뜀] {name} — {why}")
    if not os.path.exists(master):
        raise RuntimeError("최신본이 사라졌습니다")
    return done


def autoprune(master, quiet=True):
    """★ 사용자 지시(2026-07-28): "구 버전은 말 안 해도 OLD 폴더로 들어가게."

    관리대장을 찾는 길목(resolve_master·latest_master)에서 저절로 불린다.
    도구가 11개라 저장하는 쪽마다 붙이면 반드시 하나를 빠뜨린다 — **찾는 쪽**에 한 번만 건다.

    조심할 것
      · **절대 지우거나 덮어쓰지 않는다.** 같은 이름이면 출처 꼬리표를 붙여 둘 다 보존한다.
      · 엑셀이 열어 둔 파일은 못 옮긴다 → 조용히 건너뛴다. 다음 실행에 다시 시도한다.
      · 최신본은 어떤 경우에도 손대지 않는다 — 사라지면 모든 도구가 멈춘다.
      · 과거 `_이전버전/`에 남은 관리대장도 지정된 `OLD/`로 합친다.
    """
    global _AUTODONE
    if os.environ.get("CSOS_SYNTHETIC") == "1":
        return 0                              # 합성검증은 실데이터·실폴더를 절대 옮기지 않는다
    if _AUTODONE:
        return 0
    # ★ 2026-07-28: 처음엔 plan() 을 그대로 썼는데 그게 `*_v*.xlsx` 를 통째로 잡아
    #   **합성검증용 임시 파일(합성대장F_v1.xlsx)까지 옮겨** 시험이 깨졌다.
    #   자동으로 도는 것은 반드시 **관리대장 이름에만** 손대야 한다. 사람이 부르는
    #   --prune 과 달리 여기는 아무도 안 보고 있다.
    if not master or not os.path.isfile(master) or not MASTER_RE.search(os.path.basename(master)):
        return 0
    _AUTODONE = True
    try:
        return _archive_old_versions(master, quiet=quiet)
    except Exception:
        return 0                              # 정리는 부수 작업이다 — 본 작업을 막지 않는다


def main():
    from ecount_reconcile import load_config, _latest_in
    configured = load_config()["reconcile"]["master_xlsx"]
    master = _latest_in(os.path.dirname(configured))
    if not os.path.isfile(master):
        print("관리대장 최신본을 찾지 못했습니다.")
        return
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

    legacy = _legacy_files(os.path.dirname(master))
    if legacy:
        print(f"  예전 보관 폴더에서 OLD로 합칠 것 {len(legacy)}개")
    if not move and not legacy:
        print("\n정리할 것이 없습니다 — 지금은 적당한 개수입니다.")
        return
    if "--prune" not in sys.argv:
        print(f"\n미리보기 — 보관 위치: {archive_folder(os.path.dirname(master))}")
        print("실제 정리: python ledger_versions.py --prune")
        return

    done = _archive_old_versions(master, quiet=False)
    print(f"\n{done}개 정리 완료 — 보관 위치: {archive_folder(os.path.dirname(master))}")
    # 최신본이 그대로인지 마지막 확인 — 이게 사라지면 모든 도구가 멈춘다
    assert os.path.exists(master), "최신본이 사라졌습니다"
    print("  최신본 확인:", os.path.basename(master))


if __name__ == "__main__":
    main()
