# -*- coding: utf-8 -*-
"""
source_organizer.py — ``0. 원본 자료``를 장기 보관 구조로 정리한다.

기본은 미리보기이며 ``--apply`` 때만 같은 원본 보관소 안에서 파일을 이동한다.
파일은 삭제하거나 덮어쓰지 않는다. 이름 충돌은 짧은 내용 해시를 붙여 모두 보존하고,
실제 이동 이력은 ``0. 정리이력.csv``에 남긴다.

구조 원칙
  1~4, 7번 자료: 자료유형 / YYYY / MM / YYYY-MM-DD / 파일
  밴드 사진:     4. 밴드 원본 / 문서사진 / YYYY / MM / YYYY-MM-DD / 파일
  밴드 JSON:     4. 밴드 원본 / 수집본 / YYYY / MM / YYYY-MM-DD / 파일
  정기점검:      최신 편집본은 5번 폴더 바로 아래, 이전본만 보관/YYYY/MM/날짜
  PO:            6. PO 원본 / YYYY / PO번호 / 파일

PO 한 장에 여러 캠프·프로젝트가 묶이는 경우가 있어 UJ 프로젝트번호보다 PO번호가
더 안정적인 보관 키다. UJ번호가 명확하고 PO번호가 없는 자료만 UJ번호로 묶는다.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import os
import re
import shutil
import sys
import time
from dataclasses import dataclass
from datetime import datetime

from source_dirs import (
    BAND_DIR,
    COUPANG_DIR,
    ERP_DIR,
    KAKAO_DIR,
    NEW_PROJECT_FLOW_DIR,
    ORIGIN_ROOT,
    PM_SCHEDULE_DIR,
    PO_DIR,
    RECEIPT_DIR,
    UPLOAD_DIR,
    WORK_LOG_DIR,
    LEGACY_WORK_LOG_DIR,
)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

GUIDE = os.path.join(ORIGIN_ROOT, "0. 수집안내.txt")
RULES = os.path.join(ORIGIN_ROOT, "0. 정리규칙.txt")
HISTORY = os.path.join(ORIGIN_ROOT, "0. 정리이력.csv")
LOCK = os.path.join(ORIGIN_ROOT, ".source_organizer.lock")

LEGACY_PO_DIR = os.path.join(ORIGIN_ROOT, "6. 26년도 PO 모음")
LEGACY_RECEIPT_DIR = os.path.join(ORIGIN_ROOT, "5. 입금내역")
MISC_DIR = os.path.join(ORIGIN_ROOT, "9. 미분류")

PO_RE = re.compile(r"(?i)PO\s*[-_]?\s*(\d{6})(?!\d)")
UJ_RE = re.compile(r"(?i)UJ\s*[-_]?\s*(\d{7})(?!\d)")
BAND_DATE_RE = re.compile(r"(?i)^band\d+_(\d{2})(\d{2})(\d{2})_")
KAKAO_DATE_RE = re.compile(r"(?i)^KakaoTalk_(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])_")


@dataclass(frozen=True)
class Move:
    src: str
    dst: str
    reason: str


def _inside(path: str, root: str) -> bool:
    """경로가 정확히 지정 보관소 안인지 확인한다."""
    try:
        p = os.path.normcase(os.path.abspath(path))
        r = os.path.normcase(os.path.abspath(root))
        return os.path.commonpath([p, r]) == r
    except (OSError, ValueError):
        return False


def _file_day(path: str) -> datetime:
    """파일명 날짜를 우선하고, 없으면 원본 수정시각을 보관일로 쓴다."""
    name = os.path.basename(path)
    m = BAND_DATE_RE.search(name)
    if m:
        try:
            return datetime(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    # 카카오톡 표준 내보내기 이름의 첫 날짜는 내보낸 날이라 보관일로 신뢰할 수 있다.
    # 일반 파일의 YYYYMMDD는 조회 시작일·종료일일 수도 있다(예: 판매조회_2026_0102-0726).
    # 그런 값으로 분류하면 7월 수집본이 1월에 들어가므로 일반 파일은 수정시각을 쓴다.
    m = KAKAO_DATE_RE.search(name)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    try:
        return datetime.fromtimestamp(os.path.getmtime(path))
    except OSError:
        return datetime.now()


def dated_dir(base: str, path: str, day: datetime | None = None) -> str:
    day = day or _file_day(path)
    return os.path.join(base, f"{day.year:04d}", f"{day.month:02d}", day.strftime("%Y-%m-%d"))


def _key_from_path(path: str) -> tuple[str, str]:
    joined = " ".join(os.path.normpath(path).split(os.sep))
    po = PO_RE.search(joined.replace(" ", ""))
    if po:
        return f"PO{po.group(1)}", "PO번호"
    uj = UJ_RE.search(joined.replace(" ", ""))
    if uj:
        return f"UJ{uj.group(1)}", "프로젝트번호"
    return "미분류", "번호 미확정"


def po_dir_for(path: str) -> str:
    key, _ = _key_from_path(path)
    joined = os.path.normpath(path)
    m = re.search(r"(?<!\d)(20\d{2})년도", joined)
    if m:
        year = int(m.group(1))
    elif "26년도" in joined:
        year = 2026
    else:
        year = _file_day(path).year
    return os.path.join(PO_DIR, f"{year:04d}", key)


class TimeBudgetExceeded(RuntimeError):
    """정해진 시간 안에 못 끝냈다. 반쪽 결과를 내는 대신 소리 내어 멈춘다."""


# ★ 스스로 끊는 상한 (2026-08-07 실사고 — 이날 하루를 통째로 잃었다)
#   09:35 회차가 **10시간 6분** 돌았다. 작업 스케줄러는 3시간 제한대로 12:35 에
#   종료를 시도했고 결과도 0x41306(강제종료)으로 남았는데, 정작 이 python
#   프로세스는 20:15 까지 살아 있었다 — 스케줄러가 죽이는 것은 **제가 띄운
#   껍데기**지 그 아래 손자 프로세스가 아니다. 즉 밖에서 거는 제한은 믿을 수 없다.
#   그 10시간 동안 Z: 를 계속 두드려 09:50 일일대조가 매일 조용히 건너뛰었고,
#   앱의 에이전트 날짜가 08-06 에 멈춰 있었다(사용자 지적: "지금 날짜로 반영이
#   안되었어"). 그래서 **안에서** 끊는다.
#   시간을 넘기면 반쪽 정리를 남기지 않고 **실패로 끝낸다.** 반쪽은 '멀쩡해 보이는
#   거짓말'이라 사고를 더 늦게 발견하게 만든다 — 이 프로젝트가 1순위로 막는 것이다.
BUDGET_SEC = int(os.environ.get("SOURCE_ORGANIZER_BUDGET_SEC", "7200"))
_DEADLINE = 0.0
_BUDGET = 0
_SCANNED = 0


def start_clock(budget_sec: int | None = None) -> None:
    global _DEADLINE, _BUDGET, _SCANNED
    _BUDGET = int(budget_sec or BUDGET_SEC)
    _DEADLINE = time.time() + _BUDGET
    _SCANNED = 0


def check_clock(where: str = "") -> None:
    if _DEADLINE and time.time() > _DEADLINE:
        raise TimeBudgetExceeded(
            "원본 정리가 제한 %d분을 넘겼다(%s · 훑은 파일 %d개) — Z: 가 느리거나 "
            "폴더가 너무 커졌다. 반쪽으로 남기지 않고 중단한다."
            % (_BUDGET // 60, where or "훑는 중", _SCANNED))


def _iter_files(folder: str):
    global _SCANNED
    if not os.path.isdir(folder):
        return
    for base, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if d not in (".source_organizer.guard",)]
        for name in files:
            if name.startswith("~$") or name in ("Thumbs.db", ".DS_Store"):
                continue
            _SCANNED += 1
            if _SCANNED % 200 == 0:          # SMB 왕복이 비싸 자주 재지 않는다
                check_clock(base)
            yield os.path.join(base, name)


def _already_dated(path: str, base: str) -> bool:
    try:
        rel = os.path.relpath(path, base).split(os.sep)
    except ValueError:
        return False
    return (len(rel) >= 4 and re.fullmatch(r"20\d{2}", rel[0] or "") is not None
            and re.fullmatch(r"\d{2}", rel[1] or "") is not None
            and re.fullmatch(r"20\d{2}-\d{2}-\d{2}", rel[2] or "") is not None)


def _same_path(a: str, b: str) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _move_for(src: str, dst_dir: str, reason: str) -> Move | None:
    dst = os.path.join(dst_dir, os.path.basename(src))
    if _same_path(src, dst):
        return None
    return Move(src, dst, reason)


def planned_moves(root: str = ORIGIN_ROOT) -> list[Move]:
    """현재 구조에서 필요한 이동만 계산한다. 파일시스템은 바꾸지 않는다."""
    root = os.path.abspath(root)
    category = {
        os.path.join(root, "1. ERP 내보내기"): "ERP 내보내기",
        os.path.join(root, "2. 쿠팡 목록"): "쿠팡 목록",
        os.path.join(root, "3. 카카오톡 내보내기"): "카카오톡 내보내기",
        os.path.join(root, "7. 입금내역"): "입금내역",
        os.path.join(root, "5. 입금내역"): "입금내역(옛 폴더)",
    }
    moves: list[Move] = []

    for base, label in category.items():
        if not os.path.isdir(base):
            continue
        canonical = os.path.join(root, "7. 입금내역") if "입금내역" in label else base
        for src in _iter_files(base):
            if _already_dated(src, canonical):
                continue
            m = _move_for(src, dated_dir(canonical, src), f"{label} 날짜별 보관")
            if m:
                moves.append(m)

    band = os.path.join(root, "4. 밴드 원본")
    if os.path.isdir(band):
        for src in _iter_files(band):
            rel = os.path.relpath(src, band).split(os.sep)
            if rel[0] == "문서사진":
                base = os.path.join(band, "문서사진")
                if _already_dated(src, base):
                    continue
                target = dated_dir(base, src)
                reason = "밴드 문서사진 게시일별 보관"
            else:
                base = os.path.join(band, "수집본")
                if rel[0] == "수집본" and _already_dated(src, base):
                    continue
                target = dated_dir(base, src)
                reason = "밴드 API 원문 수집일별 보관"
            m = _move_for(src, target, reason)
            if m:
                moves.append(m)

    # PO는 기존 폴더와 새 폴더를 모두 받아 새 정본 구조로 합친다.
    for base in (os.path.join(root, "6. 26년도 PO 모음"), os.path.join(root, "6. PO 원본")):
        if not os.path.isdir(base):
            continue
        for src in _iter_files(base):
            rel = os.path.relpath(src, os.path.join(root, "6. PO 원본")).split(os.sep)
            if (base.endswith("6. PO 원본") and len(rel) >= 3
                    and re.fullmatch(r"20\d{2}", rel[0] or "")
                    and (PO_RE.fullmatch(rel[1]) or UJ_RE.fullmatch(rel[1]) or rel[1] == "미분류")):
                continue
            key, why = _key_from_path(src)
            m = _move_for(src, os.path.join(root, "6. PO 원본", f"{_file_day(src).year:04d}", key),
                          f"PO 원본 {why}별 보관")
            if m:
                moves.append(m)

    # 사람이 계속 편집하는 기준 원본은 최신 파일 하나를 폴더 바로 아래에 둔다.
    # 신규 프로젝트 업무 흐름도는 앱 비표시 DB 동기화 원본이므로, 이전본만 보관함으로 옮긴다.
    for current_dir, label in (
        (os.path.join(root, "5. 정기점검 스케쥴 원본"), "정기점검"),
        (os.path.join(root, "50. 쿠팡 신규 프로젝트 업무 흐름도"), "신규 프로젝트 업무 흐름도"),
        (os.path.join(root, "8. 정기점검, 돌발AS 일지(미실시건)"), "정기점검·돌발AS 일지"),
        # 이전 폴더는 호환 보관만 한다. 새 원본은 위 8번 폴더를 우선한다.
        (os.path.join(root, "7. 정기점검, 돌발AS 일지"), "정기점검·돌발AS 일지(이전)"),
    ):
        if not os.path.isdir(current_dir):
            continue
        candidates = [
            p for p in _iter_files(current_dir)
            if p.lower().endswith((".xlsx", ".xlsm"))
            and os.path.relpath(p, current_dir).split(os.sep)[0] != "보관"
        ]
        latest = max(candidates, key=os.path.getmtime) if candidates else None
        for src in candidates:
            if _same_path(src, latest) and os.path.dirname(src) == current_dir:
                continue
            m = _move_for(src, dated_dir(os.path.join(current_dir, "보관"), src),
                          f"{label} 이전본 날짜별 보관")
            if m:
                moves.append(m)

    # 사람이 루트에 바로 놓은 파일도 잃지 않고 미분류 날짜함으로 옮긴다.
    reserved = {os.path.basename(GUIDE), os.path.basename(RULES), os.path.basename(HISTORY),
                os.path.basename(LOCK)}
    if os.path.isdir(root):
        for name in os.listdir(root):
            src = os.path.join(root, name)
            if not os.path.isfile(src) or name in reserved or name.startswith("~$"):
                continue
            m = _move_for(src, dated_dir(os.path.join(root, "9. 미분류"), src),
                          "루트 투입자료 날짜별 임시보관")
            if m:
                moves.append(m)

    # 같은 원본이 둘 이상의 분기에 잡히지 않게 한다.
    unique = {}
    for m in moves:
        unique[os.path.normcase(os.path.abspath(m.src))] = m
    return sorted(unique.values(), key=lambda x: x.src.lower())


def _sha8(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:8]


def _collision_target(src: str, dst: str) -> str:
    """이름이 겹칠 때 갈 자리를 정한다. 내용이 같으면 **이미 있는 파일**을 가리킨다.

    ★ 2026-07-30 실사고: 목적지에 파일이 있으면 내용을 보지 않고 `__dup_<해시>` 를 붙여
      복사했다. 그래서 같은 자료를 다시 정리할 때마다 사본이 늘어나
      판매조회가 SHA256 동일한 **3벌**이 됐고, billing_fill 이 전부 읽어 공급가액이
      36.2억 → 108.6억으로 **3배** 합산됐다(verification_sync 도 같은 파일을 3번 읽었다).
      "삭제·덮어쓰기하지 않고 모두 보존" 은 지킨다 — 다만 **내용이 같은 건 애초에 다른
      파일이 아니다.** 진짜로 내용이 다를 때만 사본 이름을 만든다.
    """
    if not os.path.exists(dst):
        return dst
    src_hash = _sha8(src)
    if src_hash == _sha8(dst):
        return dst                      # 같은 내용 — 이미 정리돼 있다(복사 자체를 하지 않는다)
    stem, ext = os.path.splitext(dst)
    candidate = f"{stem}__dup_{src_hash}{ext}"
    n = 2
    while os.path.exists(candidate):
        if os.path.exists(candidate) and _sha8(candidate) == src_hash:
            return candidate            # 같은 내용의 사본이 이미 있다 — 또 만들지 않는다
        candidate = f"{stem}__dup_{src_hash}_{n}{ext}"
        n += 1
    return candidate


def _append_history(rows: list[tuple[str, str, str]], root: str = ORIGIN_ROOT):
    history = os.path.join(root, os.path.basename(HISTORY))
    new_file = not os.path.exists(history)
    os.makedirs(os.path.dirname(history), exist_ok=True)
    with open(history, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["정리시각", "원래경로", "새경로", "기준"])
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for src, dst, reason in rows:
            w.writerow([stamp, src, dst, reason])


def _remove_empty_dirs(root: str):
    """파일이 하나도 없는 하위 폴더만 정리한다. 카테고리 뿌리는 유지한다."""
    keep = {
        os.path.abspath(root), os.path.abspath(ERP_DIR), os.path.abspath(COUPANG_DIR),
        os.path.abspath(KAKAO_DIR), os.path.abspath(BAND_DIR), os.path.abspath(PM_SCHEDULE_DIR),
        os.path.abspath(NEW_PROJECT_FLOW_DIR),
        os.path.abspath(PO_DIR), os.path.abspath(RECEIPT_DIR), os.path.abspath(WORK_LOG_DIR),
        os.path.abspath(LEGACY_WORK_LOG_DIR),
        os.path.abspath(MISC_DIR),
        os.path.abspath(UPLOAD_DIR),
    }
    for base, _dirs, _files in os.walk(root, topdown=False):
        if os.path.abspath(base) in keep:
            continue
        try:
            if not os.listdir(base):
                os.rmdir(base)
        except OSError:
            pass


def write_rules(root: str = ORIGIN_ROOT):
    lines = [
        "쿠팡 원본 자료 자동 정리 기준",
        f"갱신: {datetime.now():%Y-%m-%d %H:%M}",
        "",
        "1. ERP·쿠팡목록·카카오톡·입금내역 = 자료유형 / 연도 / 월 / 수집일",
        "2. 밴드 문서사진 = 게시일, 밴드 JSON = 수집일 기준",
        "3. PO = 연도 / PO번호. PO번호가 없고 UJ번호만 확실하면 프로젝트번호로 분류",
        "4. 정기점검·업무일지·신규 프로젝트 흐름도 = 최신 편집본은 해당 폴더 바로 아래, 이전본은 보관/연도/월/날짜",
        "5. 분류할 단서가 없는 루트 파일 = 9. 미분류/연도/월/날짜",
        "6. 파일은 삭제·덮어쓰기하지 않으며, 내용이 다를 때만 __dup_내용해시를 붙여 모두 보존",
        "7. 이동 이력은 0. 정리이력.csv에서 원래 위치까지 확인 가능",
        "8. 100. 업로드용 자료 = 단일 투입함. upload_intake가 내용 판별 후 위 정본 폴더로 이동",
        "",
        "새 자료는 100. 업로드용 자료에 넣으면 자동 분류됩니다.",
    ]
    path = os.path.join(root, os.path.basename(RULES))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def apply_moves(moves: list[Move], root: str = ORIGIN_ROOT) -> tuple[int, list[str]]:
    root = os.path.abspath(root)
    done_count = 0
    history_batch: list[tuple[str, str, str]] = []
    errors: list[str] = []
    for m in moves:
        if not (_inside(m.src, root) and _inside(m.dst, root)):
            errors.append(f"보관소 밖 경로 차단: {m.src} -> {m.dst}")
            continue
        if not os.path.isfile(m.src):
            continue
        dst = _collision_target(m.src, m.dst)
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.move(m.src, dst)
            history_batch.append((m.src, dst, m.reason))
            done_count += 1
            # 네트워크 드라이브 대량 이동은 오래 걸린다. 중간에 PC가 꺼져도 원래 경로를
            # 대부분 복원할 수 있도록 25개마다 이력을 확정한다.
            if len(history_batch) >= 25:
                _append_history(history_batch, root)
                history_batch.clear()
                # 이력을 확정한 직후에만 시계를 본다 — 여기서 멈추면 이미 옮긴 것은
                # 전부 이력에 남아 있어 되돌릴 수 있다(파일 하나가 뜨는 일이 없다).
                check_clock("이동 %d개째" % done_count)
        except OSError as e:
            errors.append(f"{m.src}: {e}")
    if history_batch:
        _append_history(history_batch, root)
    _remove_empty_dirs(root)
    write_rules(root)
    return done_count, errors


def _lock_acquire() -> bool:
    os.makedirs(ORIGIN_ROOT, exist_ok=True)
    try:
        fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        try:
            if time.time() - os.path.getmtime(LOCK) > 30 * 60:
                os.unlink(LOCK)
                return _lock_acquire()
        except OSError:
            pass
        return False
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(f"{os.getpid()} {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    return True


def _lock_release():
    try:
        os.unlink(LOCK)
    except OSError:
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="실제 정리")
    ap.add_argument("--limit", type=int, default=0, help="테스트용 처리 한도")
    ap.add_argument("--budget-min", type=int, default=BUDGET_SEC // 60,
                    help="이 시간을 넘기면 반쪽으로 두지 않고 중단한다(기본 120분)")
    args = ap.parse_args()
    if not os.path.isdir(ORIGIN_ROOT):
        print("원본 자료 폴더에 접근할 수 없습니다:", ORIGIN_ROOT)
        return 2
    start_clock(max(1, args.budget_min) * 60)
    try:
        moves = planned_moves()
    except TimeBudgetExceeded as e:
        # 조용히 성공한 척하지 않는다 — 이게 안 보여서 열 시간을 잃었다.
        print("중단:", e)
        return 4
    if args.limit > 0:
        moves = moves[:args.limit]
    print(f"{'정리 실행' if args.apply else '미리보기'} — 이동 대상 {len(moves)}개")
    reasons = {}
    for m in moves:
        reasons[m.reason] = reasons.get(m.reason, 0) + 1
    for reason, n in sorted(reasons.items()):
        print(f"  {reason}: {n}개")
    for m in moves[:20]:
        print(" ", os.path.relpath(m.src, ORIGIN_ROOT), "→", os.path.relpath(m.dst, ORIGIN_ROOT))
    if len(moves) > 20:
        print(f"  ... 외 {len(moves) - 20}개")
    if not args.apply:
        print("\n실제 정리: python source_organizer.py --apply")
        return 0
    if not _lock_acquire():
        print("다른 원본 정리 작업이 실행 중이라 이번 실행을 건너뜁니다.")
        return 3
    try:
        done, errors = apply_moves(moves)
    except TimeBudgetExceeded as e:
        print("중단:", e)
        return 4
    finally:
        _lock_release()
    print(f"완료: {done}개 이동 · 오류 {len(errors)}개")
    for e in errors[:20]:
        print("  오류:", e)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
