# -*- coding: utf-8 -*-
"""100. 업로드용 자료 → 0. 원본 자료의 유형별 정본 폴더로 안전하게 분류한다.

투입함은 재귀적으로 전부 훑는다. 아는 자료는 내용 우선으로 분류하고, 확정할 수 없는
파일도 버리지 않고 ``9. 미분류``에 보관한다. 같은 이름을 덮어쓰지 않으며 동일 내용은
정본 한 벌만 남긴다. 실제 이동은 ``--apply`` 때만 수행한다.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import time
from dataclasses import dataclass
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from source_organizer import dated_dir  # noqa: E402

REPORT = os.path.join(ROOT, "reports", "upload_intake.json")
MIN_STABLE_SECONDS = 30
LOCK_NAME = ".upload_intake.lock"
PO_RE = re.compile(r"(?i)\bPO\s*[-_]?\s*(\d{5,})\b")
UJ_RE = re.compile(r"(?i)\bUJ\s*[-_]?\s*(\d{7})\b")
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic"}
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
TEXT_EXT = {".txt", ".csv", ".json", ".xml", ".md"}
IGNORE_NAMES = {"thumbs.db", ".ds_store"}

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


@dataclass(frozen=True)
class Intake:
    src: str
    dst_dir: str
    kind: str
    reason: str


def _paths(root: str) -> dict[str, str]:
    return {
        "upload": os.path.join(root, "100. 업로드용 자료"),
        "erp": os.path.join(root, "1. ERP 내보내기"),
        "coupang": os.path.join(root, "2. 쿠팡 목록"),
        "kakao": os.path.join(root, "3. 카카오톡 내보내기"),
        "band": os.path.join(root, "4. 밴드 원본"),
        "pm": os.path.join(root, "5. 정기점검 스케쥴 원본"),
        "po": os.path.join(root, "6. PO 원본"),
        "receipt": os.path.join(root, "7. 입금내역"),
        "worklog": os.path.join(root, "8. 정기점검, 돌발AS 일지(미실시건)"),
        "misc": os.path.join(root, "9. 미분류"),
        "flow": os.path.join(root, "50. 쿠팡 신규 프로젝트 업무 흐름도"),
    }


def _inside(path: str, root: str) -> bool:
    try:
        p, r = os.path.abspath(path), os.path.abspath(root)
        return os.path.commonpath([p, r]) == r
    except (OSError, ValueError):
        return False


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _text_head(path: str, limit: int = 512 * 1024) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXT:
        try:
            from inbox_scan import _cells
            rows = _cells(path, sheets=8, rows=20)
            titles = []
            try:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                titles = list(wb.sheetnames)
                wb.close()
            except Exception:
                pass
            return ("\n".join(titles) + "\n" +
                    "\n".join(" | ".join(row) for row in rows))[:limit]
        except Exception:
            return ""
    if ext == ".pdf":
        try:
            from pypdf import PdfReader
            reader = PdfReader(path)
            return "\n".join((page.extract_text() or "") for page in reader.pages[:5])[:limit]
        except Exception:
            return ""
    if ext not in TEXT_EXT:
        return ""
    try:
        with open(path, "rb") as fh:
            raw = fh.read(limit)
    except OSError:
        return ""
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _po_dir(base: str, path: str, text: str) -> str:
    blob = os.path.basename(path) + " " + text
    po = PO_RE.search(blob.replace("-", ""))
    uj = UJ_RE.search(blob.replace("-", ""))
    if po:
        key = "PO" + po.group(1)
    elif uj:
        key = "UJ" + uj.group(1)
    else:
        key = "미분류"
    try:
        year = datetime.fromtimestamp(os.path.getmtime(path)).year
    except OSError:
        year = datetime.now().year
    return os.path.join(base, f"{year:04d}", key)


def classify_target(path: str, root: str) -> tuple[str, str, str]:
    """(목적지 폴더, 분류키, 근거). 모르면 반드시 미분류로 보낸다."""
    p = _paths(root)
    name = os.path.basename(path)
    low = name.lower()
    ext = os.path.splitext(low)[1]
    text = _text_head(path)
    blob = re.sub(r"\s+", " ", f"{name} {text}").strip()

    # 사람이 계속 편집하는 정본 3종은 최신 파일이 유형 폴더 바로 아래에 있어야 한다.
    if ext in EXCEL_EXT and "정기점검" in blob and any(k in blob for k in ("스케줄", "스케쥴", "일정")):
        return p["pm"], "pm_schedule", "정기점검 일정 원본"
    if ext in EXCEL_EXT and ("신규 프로젝트" in blob and any(k in blob for k in ("흐름도", "프로세스"))):
        return p["flow"], "new_project_flow", "신규 프로젝트 업무 흐름도"
    if ext in EXCEL_EXT and "일지" in blob and any(k in blob for k in ("돌발", "정기점검", "미실시")):
        return p["worklog"], "work_log", "정기점검·돌발AS 일지"

    kakao_form = (low.startswith("kakaotalk") or
                  ("카카오톡 대화" in blob and re.search(r"\[(오전|오후)?\s*\d{1,2}:\d{2}\]", blob)))
    if ext == ".txt" and kakao_form:
        return dated_dir(p["kakao"], path), "kakao", "카카오톡 대화 내보내기"

    if ext == ".json" and (low.startswith(("dump_", "raw_")) or
                            '"posts"' in text and ('"band"' in text or '"band_name"' in text)):
        return dated_dir(os.path.join(p["band"], "수집본"), path), "band_dump", "밴드 원문 JSON"

    if ext in EXCEL_EXT:
        try:
            from inbox_scan import classify
            kind = classify(path)
        except Exception:
            kind = "unknown"
        if kind == "po":
            return dated_dir(p["coupang"], path), "coupang_po", "엑셀 내용: 쿠팡 PO 목록"
        if kind == "receipt":
            return dated_dir(p["receipt"], path), "receipt", "엑셀 내용: 입금·수금 내역"
        if kind in {"ledger", "sales", "tax", "stmt", "slips", "taxinv", "hometax"}:
            return dated_dir(p["erp"], path), f"erp_{kind}", f"엑셀 내용: {kind}"

    if any(k in blob for k in ("입금내역", "입금 내역", "수금내역", "수금 내역", "송금명세")):
        return dated_dir(p["receipt"], path), "receipt", "입금·수금 원본"

    if PO_RE.search(blob.replace("-", "")) or any(k in blob for k in ("구매 오더", "구매오더", "견적서", "발주서")):
        return _po_dir(p["po"], path, text), "po_document", "PO·견적 원본"

    if ext in IMAGE_EXT:
        return dated_dir(os.path.join(p["band"], "문서사진"), path), "document_image", "문서·현장 사진"

    if any(k in blob for k in ("거래명세서", "세금계산서", "판매조회", "계정별원장")):
        return dated_dir(p["erp"], path), "erp_document", "ERP·거래서류 원본"

    return dated_dir(p["misc"], path), "unknown", "자동 판별 불가 — 미분류 보존"


def plan(root: str | None = None, min_age: int = MIN_STABLE_SECONDS) -> list[Intake] | None:
    from source_dirs import ORIGIN_ROOT
    root = os.path.abspath(root or ORIGIN_ROOT)
    upload = _paths(root)["upload"]
    if not os.path.isdir(root):
        return None
    if not os.path.isdir(upload):
        return []
    now = time.time()
    jobs: list[Intake] = []
    for base, dirs, files in os.walk(upload):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for name in sorted(files):
            if name.lower() in IGNORE_NAMES or name.startswith(("~$", ".")):
                continue
            src = os.path.join(base, name)
            try:
                before = os.stat(src)
                if min_age > 0 and now - before.st_mtime < min_age:
                    continue  # 복사 중인 파일은 다음 회차에서 처리
            except OSError:
                continue
            dst, kind, reason = classify_target(src, root)
            try:
                after = os.stat(src)
                if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
                    continue  # 분류 중에도 크기가 바뀌면 아직 업로드 중이다.
            except OSError:
                continue
            jobs.append(Intake(src, dst, kind, reason))
    return jobs


def _load_index(path: str) -> dict[str, str]:
    try:
        value = json.load(open(path, encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def _save_json(path: str, value) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(value, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(tmp, path)


def _collision(src: str, dst: str, digest: str) -> tuple[str, bool]:
    if not os.path.exists(dst):
        return dst, False
    try:
        if _sha256(dst) == digest:
            return dst, True
    except OSError:
        pass
    stem, ext = os.path.splitext(dst)
    candidate = f"{stem}__dup_{digest[:8]}{ext}"
    n = 2
    while os.path.exists(candidate):
        try:
            if _sha256(candidate) == digest:
                return candidate, True
        except OSError:
            pass
        candidate = f"{stem}__dup_{digest[:8]}_{n}{ext}"
        n += 1
    return candidate, False


def apply(jobs: list[Intake], root: str | None = None,
          report_path: str = REPORT, index_path: str | None = None) -> tuple[list[dict], list[dict]]:
    from source_dirs import ORIGIN_ROOT
    root = os.path.abspath(root or ORIGIN_ROOT)
    upload = _paths(root)["upload"]
    index_path = index_path or os.path.join(upload, ".upload_intake_index.json")
    index = _load_index(index_path)
    done, failed = [], []
    history = []
    for job in jobs:
        if not (_inside(job.src, upload) and _inside(job.dst_dir, root)):
            failed.append({"파일": job.src, "오류": "허용된 원본 보관소 밖 경로"})
            continue
        try:
            digest = _sha256(job.src)
            prior = index.get(digest)
            if prior and _inside(prior, root) and os.path.isfile(prior) and _sha256(prior) == digest:
                os.remove(job.src)
                done.append({"파일": os.path.basename(job.src), "분류": job.kind,
                             "근거": job.reason, "목적지": prior, "처리": "동일 원본 통합"})
                history.append((job.src, prior, job.reason + " / 동일 원본 통합"))
                continue
            os.makedirs(job.dst_dir, exist_ok=True)
            target, duplicate = _collision(job.src, os.path.join(job.dst_dir, os.path.basename(job.src)), digest)
            if duplicate:
                os.remove(job.src)
                action = "동일 원본 통합"
                history.append((job.src, target, job.reason + " / 동일 원본 통합"))
            else:
                shutil.move(job.src, target)
                action = "원본 이동"
                history.append((job.src, target, job.reason))
            index[digest] = target
            done.append({"파일": os.path.basename(job.src), "분류": job.kind,
                         "근거": job.reason, "목적지": target, "처리": action})
        except OSError as exc:
            failed.append({"파일": job.src, "오류": str(exc)[:180]})

    if history:
        try:
            from source_organizer import _append_history
            _append_history(history, root)
        except Exception as exc:
            failed.append({"파일": "0. 정리이력.csv", "오류": str(exc)[:180]})
    _save_json(index_path, index)
    payload = {"time": datetime.now().astimezone().isoformat(timespec="seconds"),
               "이동": done, "실패": failed,
               "미분류": [row for row in done if row.get("분류") == "unknown"]}
    _save_json(report_path, payload)
    return done, failed


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="실제 분류 이동")
    ap.add_argument("--min-age", type=int, default=MIN_STABLE_SECONDS,
                    help="복사 완료로 볼 최소 수정 후 경과 초")
    args = ap.parse_args()
    from source_dirs import ORIGIN_ROOT, UPLOAD_DIR
    if args.apply:
        if not os.path.isdir(ORIGIN_ROOT):
            print("원본 자료 폴더(Z:)에 닿지 않습니다 — 이번 회차는 건너뜁니다.")
            return 0
        os.makedirs(UPLOAD_DIR, exist_ok=True)
    lock = os.path.join(UPLOAD_DIR, LOCK_NAME)
    locked = False
    if args.apply:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write(f"{os.getpid()} {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            locked = True
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(lock) > 30 * 60:
                    os.unlink(lock)
                    return main()
            except OSError:
                pass
            print("다른 업로드 분류 작업이 실행 중이라 이번 회차를 건너뜁니다.")
            return 0
    try:
        jobs = plan(min_age=max(0, args.min_age))
        if jobs is None:
            print("원본 자료 폴더(Z:)에 닿지 않습니다 — 이번 회차는 건너뜁니다.")
            return 0
        if not args.apply:
            print(f"업로드 투입함 미리보기: {len(jobs)}건")
            for job in jobs[:20]:
                print(f"  [{job.kind}] {os.path.basename(job.src)} → {job.dst_dir}")
            return 0
        done, failed = apply(jobs)
        unknown = sum(1 for row in done if row.get("분류") == "unknown")
        print(f"업로드 원본 분류: {len(done)}건 · 미분류 {unknown}건 · 실패 {len(failed)}건")
        return 1 if failed else 0
    finally:
        if locked:
            try:
                os.unlink(lock)
            except OSError:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
