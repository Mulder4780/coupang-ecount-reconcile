"""Produce a verified, local-only Excel archive from the canonical app DB.

The application database is the system of record.  This worker asks
``ArchiveExporter`` for an immutable snapshot and command plan, applies that
plan to the exporter's *local byte-copy* of the workbook, and lets the exporter
promote the result to ``last-good`` only after both structural and semantic
verification pass.

Safety contract
---------------
* Z: and UNC paths may be read as templates, but are never write targets.
* The supplied template and the exporter's template-copy are hash checked and
  never opened for writing.
* Workbook changes are ZIP/XML patches.  ``openpyxl.save`` is not used.
* Formula cells are preserved; the workbook is marked for full recalculation.
* A snapshot/template pair is idempotent because ``ArchiveExporter`` derives a
  deterministic export id and returns an already verified artifact unchanged.
* A local process lock prevents two scheduled workers from finalizing the same
  artifact concurrently.

CLI examples::

    python archive_worker.py --run
    python archive_worker.py --run --template C:\\local\\template.xlsx
    python archive_worker.py --status
    python archive_worker.py --self-test
"""

from __future__ import annotations

import argparse
import contextlib
import copy
import hashlib
import html
import json
import os
import re
import sys
import tempfile
import time
import traceback
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Mapping, MutableMapping, Optional, Sequence, Tuple

import ledger_writer
import pid_alive
import proc_guard
from app_store import AppStore, SHEET_SPECS, canonical_json, default_store, sha256_json
from archive_export import (
    ArchiveExporter,
    ArchiveSafetyError,
    ArchiveVerificationError,
    PLAN_FORMAT,
    sha256_file,
)


WORKER_STATUS_FORMAT = "csos-archive-worker-status/v1"
WORKER_PROOF_FORMAT = "csos-archive-worker-proof/v1"
SIDECAR_SHEET = "99_AppDB_미매칭보관"
SIDECAR_FORMAT = "csos-appdb-sidecar/v1"
# 15k Unicode code points remains below Excel's 32,767 UTF-16-unit cell cap
# even when a payload consists entirely of non-BMP characters.
SIDECAR_CHUNK_SIZE = 15_000
HEADER_ROW = int(getattr(ledger_writer, "HDR_ROW", 4))
FIRST_DATA_ROW = int(getattr(ledger_writer, "FIRST", 5))
LOCK_STALE_SECONDS = 6 * 60 * 60

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DATE_HEADER = re.compile(r"(?:^|)(?:일|일자|날짜|예정일|완료일|등록일|확인일|발행일|수금일)$")
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_CELL_REF = re.compile(r"^(?P<col>[A-Z]{1,4})(?P<row>\d+)$")
# Put the self-closing alternative first.  A pattern shaped like
# ``[^>]*(?:/>|>.*?</row>)`` can consume the slash in ``<row .../>``, take the
# second alternative and then swallow the following non-empty row.  The same
# trap exists for cells and would make a blank cell appear to contain the next
# cell's formula.
_ROW_TAG = re.compile(r"<row\b[^>]*?/>|<row\b[^>]*>.*?</row>", re.S)
_CELL_TAG = re.compile(r"<c\b[^>]*?/>|<c\b[^>]*>.*?</c>", re.S)


class ArchiveWorkerError(RuntimeError):
    """Base error for deterministic worker failures."""


class ArchiveWorkerBusy(ArchiveWorkerError):
    """Another live local worker owns the spool lock."""


class ArchiveRenderError(ArchiveWorkerError):
    """The DB snapshot cannot be represented safely in the workbook."""


class ArchiveSourceError(ArchiveWorkerError):
    """A read-only source could not be staged locally within a bounded time."""


# Canonical DB audit attributes deliberately have no legacy workbook columns.
# Keep the unknown-field guard strict for every other name so a typo in a real
# workbook field still fails closed instead of silently disappearing.
DB_ONLY_ARCHIVE_FIELDS = frozenset(
    {
        "객관완료여부",
        "객관완료일",
        "객관완료상태",
        "객관완료근거",
        # 접수취소·원격해결의 감사 근거는 앱 DB 정본 필드다. 기존 02/04 시트에는
        # 대응 열이 없으므로 진행/점검상태만 주 시트에 반영하고, 아래 필드들은
        # snapshot.json·SQLite backup에 원문 그대로 보존한다. 임의의 기존 열에
        # 밀어 넣으면 오히려 관리대장 의미가 깨진다.
        "접수취소여부",
        "접수취소사유",
        "접수취소확인일",
        "처리구분",
        "접수취소근거",
        # '공급가액'(일반)은 06시트에 단일 열이 없다 — 실제작업/거래명세서/세금계산서
        # 공급가액으로 나뉜다. 그 특정 열들이 같은 기록에 이미 있어 이 일반값은
        # 중복 파생값이다. 어느 열인지 짐작해 쓰면 엉뚱한 칸에 박히므로 DB에만 둔다.
        "공급가액",
        # ★ 미처리 사유(담당자) — 02/04 시트에 대응 열이 **없다** (2026-08-19 지시).
        #   담당자가 대표 캡처에서 그 자리에 적는 사유이며 정본은 앱 DB 다.
        #   여기 안 올리면 `cannot archive unknown field` 로 **보관본 회차가 통째로
        #   죽는다** — 그런데 저장 자체는 성공하므로 적은 사람은 모르고 11:00·15:00
        #   회차만 조용히 실패한다. 위 접수취소 갈래와 같은 자리다.
        "미처리사유(담당자)",
        # ★ 청구 제외(다녀왔지만 이 건으로는 청구 안 함) — 02/04 시트에 대응 열이
        #   **없다**(v608 실측: 02시트 44열·04시트 32열 머리글 4행에 둘 다 없음).
        #   2026-08-13 류지영 요청으로 담당자 입력칸이 생겼는데 여기 안 올려서,
        #   실측 2026-08-20 보관본 회차가 통째로 죽고 있었다:
        #   `ArchiveRenderError: 02_돌발AS접수:AS-2606-092 cannot archive unknown
        #   field '청구제외'`. 저장은 성공하므로 **적은 사람은 모르고** 11:00·15:00
        #   회차만 조용히 실패한다 — 바로 위 갈래가 글로 경고해 둔 그 자리다.
        "청구제외",
        "청구제외사유",
    }
)


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def _is_network_or_z(path: os.PathLike[str] | str) -> bool:
    raw = str(path).strip().replace("/", "\\")
    return raw.startswith("\\\\") or bool(re.match(r"(?i)^Z:\\", raw))


def _require_local_write(path: os.PathLike[str] | str, label: str) -> Path:
    """Reject every write target outside this machine and source drive."""

    if _is_network_or_z(path):
        raise ArchiveSafetyError(f"{label} cannot be on Z:/UNC: {path}")
    candidate = Path(path).resolve()
    module_drive = Path(__file__).resolve().drive.upper()
    if candidate.drive.upper() == "Z:":
        raise ArchiveSafetyError(f"{label} cannot be on Z:: {candidate}")
    if module_drive and candidate.drive and candidate.drive.upper() != module_drive:
        raise ArchiveSafetyError(
            f"{label} must stay on local drive {module_drive}: {candidate}"
        )
    return candidate


def _atomic_json(path: Path, value: Mapping[str, Any]) -> None:
    path = _require_local_write(path, "worker status")
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent)
    )
    temp = Path(temp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def _copy_source_to_local(
    source: os.PathLike[str] | str,
    destination: os.PathLike[str] | str,
) -> Dict[str, Any]:
    """Read a workbook once into a new local file and return copy proof.

    This function is also the isolated child-process entrypoint used for a
    Z:/UNC source.  It deliberately never resolves or opens ``source`` for
    writing.  A failed child removes its own partial file when it can; a parent
    whose child remains stuck preserves the stage path as diagnostic evidence.
    """

    raw_source = str(source)
    target = _require_local_write(destination, "staged template")
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        raise FileExistsError(f"staged template already exists: {target}")
    digest = hashlib.sha256()
    size = 0
    completed = False
    try:
        with open(raw_source, "rb") as reader, open(target, "xb") as writer:
            while True:
                chunk = reader.read(4 * 1024 * 1024)
                if not chunk:
                    break
                writer.write(chunk)
                digest.update(chunk)
                size += len(chunk)
            writer.flush()
            os.fsync(writer.fileno())
        with zipfile.ZipFile(target, "r") as archive:
            bad = archive.testzip()
            if bad:
                raise ArchiveVerificationError(f"staged workbook ZIP CRC failed: {bad}")
        if sha256_file(target) != digest.hexdigest():
            raise ArchiveVerificationError("staged workbook hash changed after copy")
        completed = True
        return {
            "ok": True,
            "source": raw_source,
            "destination": str(target),
            "bytes": size,
            "sha256": digest.hexdigest(),
            "read_only_source": True,
            "external_write_performed": False,
        }
    finally:
        if not completed and target.exists():
            try:
                target.unlink()
            except OSError:
                pass


def _source_timeout_seconds() -> float:
    raw = str(os.environ.get("COUPANG_ARCHIVE_SOURCE_TIMEOUT") or "180").strip()
    try:
        return max(10.0, min(float(raw), 1800.0))
    except ValueError:
        return 180.0


@contextlib.contextmanager
def _staged_template(
    source: os.PathLike[str] | str,
    spool_dir: Path,
    *,
    force_child: bool = False,
) -> Iterator[Tuple[Path, Mapping[str, Any]]]:
    """Yield a local stable template, bounding every Z:/UNC read in a child.

    Local inputs are used read-only in place.  Network inputs are copied by
    :func:`proc_guard.run_tree`; if SMB enters an uninterruptible wait the main
    worker still returns and records the stuck PID instead of holding its lock
    forever.
    """

    raw = str(source)
    if not force_child and not _is_network_or_z(raw):
        local = Path(raw).resolve()
        if not local.is_file():
            raise FileNotFoundError(f"archive template not found: {local}")
        yield local, {
            "mode": "local-read-only",
            "source": str(local),
            "staged": False,
            "read_only_source": True,
            "external_write_performed": False,
        }
        return

    stage_root = _require_local_write(spool_dir / "source-stage", "source stage")
    stage_root.mkdir(parents=True, exist_ok=True)
    stage_dir = stage_root / uuid.uuid4().hex
    stage_dir.mkdir(parents=False, exist_ok=False)
    source_name = Path(raw).name or "template.xlsx"
    if not source_name.lower().endswith(".xlsx"):
        source_name += ".xlsx"
    destination = stage_dir / source_name
    command = [
        sys.executable,
        "-B",
        str(Path(__file__).resolve()),
        "--_copy-source",
        raw,
        "--_copy-dest",
        str(destination),
    ]
    try:
        result = proc_guard.run_tree(
            command,
            cwd=str(Path(__file__).resolve().parent),
            timeout=_source_timeout_seconds(),
            drain_timeout=20,
            output_limit=100_000,
        )
    except Exception as exc:
        try:
            destination.unlink(missing_ok=True)
            stage_dir.rmdir()
        except OSError:
            pass
        raise ArchiveSourceError(f"could not start source-stage child: {exc}") from exc
    proof: Dict[str, Any] = {}
    if result.returncode == 0 and not result.timed_out:
        try:
            proof = json.loads(result.stdout)
            if not proof.get("ok") or not destination.is_file():
                raise ArchiveSourceError(
                    "source-stage child reported success without a local copy"
                )
            if sha256_file(destination) != str(proof.get("sha256") or ""):
                raise ArchiveVerificationError("source-stage proof hash mismatch")
        except (TypeError, ValueError, json.JSONDecodeError) as exc:
            try:
                destination.unlink(missing_ok=True)
                stage_dir.rmdir()
            except OSError:
                pass
            raise ArchiveSourceError(
                f"source-stage child returned invalid proof: {exc}"
            ) from exc
        except (ArchiveWorkerError, ArchiveVerificationError):
            try:
                destination.unlink(missing_ok=True)
                stage_dir.rmdir()
            except OSError:
                pass
            raise
    else:
        detail = (result.stderr or result.stdout or "").strip()[-2_000:]
        message = (
            "source staging timed out"
            if result.timed_out
            else f"source staging failed (exit={result.returncode})"
        )
        message += f"; stuck_pid={result.stuck_pid}; stage={destination}"
        if detail:
            message += f"; detail={detail}"
        if not result.stuck_pid:
            try:
                destination.unlink(missing_ok=True)
                stage_dir.rmdir()
            except OSError:
                pass
        raise ArchiveSourceError(message)
    try:
        yield destination, {
            **proof,
            "mode": "bounded-read-only-child-copy",
            "staged": True,
            "stuck_pid": 0,
        }
    finally:
        try:
            destination.unlink(missing_ok=True)
            stage_dir.rmdir()
        except OSError:
            pass


@contextlib.contextmanager
def _worker_lock(spool_dir: Path, wait_seconds: float = 0.0) -> Iterator[None]:
    """Own one local worker lock; only dead/stale locks are moved aside."""

    spool_dir = _require_local_write(spool_dir, "archive spool")
    spool_dir.mkdir(parents=True, exist_ok=True)
    lock = spool_dir / "archive-worker.lock"
    token = {
        **pid_alive.identity(),
        "created_at": _utcnow(),
        "token": uuid.uuid4().hex,
    }
    deadline = time.monotonic() + max(0.0, float(wait_seconds))
    fd: Optional[int] = None
    while fd is None:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, canonical_json(token).encode("utf-8"))
            os.fsync(fd)
        except FileExistsError:
            owner: Dict[str, Any] = {}
            try:
                owner = json.loads(lock.read_text(encoding="utf-8"))
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                owner = {}
            try:
                lock_mtime = lock.stat().st_mtime
                age = max(0.0, time.time() - lock_mtime)
            except OSError:
                continue
            owner_pid = owner.get("pid")
            # 잠금이 쓰인 시각보다 뒤에 태어난 프로세스는 주인이 아니다 —
            # pid 재사용 오판 방지(검증 [210]).
            live = pid_alive.owner_alive(
                owner_pid,
                pid_started_at=owner.get("pid_started_at"),
                born_before=lock_mtime,
            )
            # Never reclaim a fresh, half-written or otherwise unidentifiable
            # claim.  Only a definitely dead PID, or an unidentifiable claim
            # older than the stale horizon, is safe to move aside.
            # A valid PID with an unavailable creation-time lookup remains
            # occupied conservatively.  Only unidentifiable/corrupt legacy
            # claims use the stale-age escape hatch.
            occupied = (
                live is not False
                if owner_pid
                else age < LOCK_STALE_SECONDS
            )
            if occupied:
                if time.monotonic() >= deadline:
                    raise ArchiveWorkerBusy(
                        f"archive worker already running (pid={owner.get('pid')})"
                    )
                time.sleep(0.1)
                continue
            # Preserve a dead lock as diagnostic evidence instead of deleting it.
            aside = spool_dir / (
                f"archive-worker.lock.stale-{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"
            )
            try:
                os.replace(lock, aside)
            except FileNotFoundError:
                continue
    try:
        yield
    finally:
        if fd is not None:
            os.close(fd)
        try:
            current = json.loads(lock.read_text(encoding="utf-8"))
            if current.get("token") == token["token"]:
                lock.unlink()
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            pass


def _row_fragments(xml: str) -> List[Tuple[int, int, int, str]]:
    rows: List[Tuple[int, int, int, str]] = []
    for match in _ROW_TAG.finditer(xml):
        number = re.search(r'\br="(\d+)"', match.group(0))
        if number:
            rows.append((int(number.group(1)), match.start(), match.end(), match.group(0)))
    return rows


def _cell_pattern(ref: str) -> re.Pattern[str]:
    return re.compile(
        r'(?:<c\b(?=[^>]*\br="'
        + re.escape(ref)
        + r'")[^>]*?/>|<c\b(?=[^>]*\br="'
        + re.escape(ref)
        + r'")[^>]*>.*?</c>)',
        re.S,
    )


def _cell_fragment(row_fragment: str, ref: str) -> str:
    match = _cell_pattern(ref).search(row_fragment)
    return match.group(0) if match else ""


def _unescape_text(value: str) -> str:
    return html.unescape(re.sub(r"<[^>]+>", "", value or ""))


def _cell_value(fragment: str, shared_strings: Sequence[str]) -> Any:
    if not fragment:
        return ""
    inline = re.findall(r"<t\b[^>]*>(.*?)</t>", fragment, re.S)
    if inline and ('t="inlineStr"' in fragment or "<is" in fragment):
        return "".join(html.unescape(v) for v in inline)
    value_match = re.search(r"<v\b[^>]*>(.*?)</v>", fragment, re.S)
    raw = html.unescape(value_match.group(1)) if value_match else ""
    if 't="s"' in fragment:
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if 't="b"' in fragment:
        return raw == "1"
    return raw


def _shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    raw = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="strict")
    values: List[str] = []
    for match in re.finditer(r"<si\b[^>]*>(.*?)</si>", raw, re.S):
        texts = re.findall(r"<t\b[^>]*>(.*?)</t>", match.group(1), re.S)
        values.append("".join(html.unescape(v) for v in texts))
    return values


def _formula(fragment: str) -> bool:
    return bool(re.search(r"<f(?:\s|>|/)", fragment or ""))


def _canonical_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    return str(value).strip()


def _value_type(header: str, value: Any, meta: Mapping[str, Any]) -> str:
    declared = str(meta.get("value_type") or "").lower()
    if isinstance(value, bool) or declared == "bool":
        return "bool"
    if isinstance(value, (int, float, Decimal)) or declared == "number":
        return "number"
    text = str(value or "").strip()
    if _ISO_DATE.fullmatch(text) and (
        _DATE_HEADER.search(str(header or ""))
        or any(word in str(header or "") for word in ("예정", "완료", "발행", "입금", "점검"))
    ):
        return "date"
    return "text"


def _values_equal(current: Any, desired: Any, value_type: str) -> bool:
    if value_type == "date":
        expected = str(ledger_writer.date_serial(str(desired)[:10]))
        if _canonical_text(current) == str(desired)[:10]:
            return True
        try:
            return Decimal(_canonical_text(current)) == Decimal(expected)
        except InvalidOperation:
            return False
    if value_type == "number":
        try:
            return Decimal(_canonical_text(current).replace(",", "")) == Decimal(
                _canonical_text(desired).replace(",", "")
            )
        except InvalidOperation:
            return _canonical_text(current) == _canonical_text(desired)
    if value_type == "bool":
        current_bool = str(current).strip().lower() in {"1", "true", "yes", "y", "예"}
        desired_bool = str(desired).strip().lower() in {"1", "true", "yes", "y", "예"}
        return current_bool == desired_bool
    return _canonical_text(current) == _canonical_text(desired)


def _translate_formula(formula: str, delta: int) -> str:
    """Translate relative A1 row references when cloning one worksheet row."""

    def repl(match: re.Match[str]) -> str:
        col, row_lock, row = match.group(1), match.group(2), int(match.group(3))
        if row_lock == "$":
            return match.group(0)
        return f"{col}{row + delta}"

    return re.sub(r"(\$?[A-Z]{1,4})(\$?)(\d+)", repl, formula)


def _empty_cloned_cell(fragment: str, old_row: int, new_row: int) -> str:
    old_ref = re.search(r'\br="([A-Z]{1,4})\d+"', fragment)
    if not old_ref:
        return ""
    col = old_ref.group(1)
    style = re.search(r'\bs="([^"]+)"', fragment)
    style_attr = f' s="{style.group(1)}"' if style else ""
    formula = re.search(r"<f\b([^>]*)>(.*?)</f>|<f\b([^>]*)/>", fragment, re.S)
    if not formula:
        return f'<c r="{col}{new_row}"{style_attr}/>'
    attrs = (formula.group(1) if formula.group(1) is not None else formula.group(3)) or ""
    if re.search(r'\bt="shared"', attrs):
        raise ArchiveRenderError(
            "cannot extend a worksheet row containing a shared formula; expand the template"
        )
    body = formula.group(2) or ""
    translated = _translate_formula(body, new_row - old_row)
    return f'<c r="{col}{new_row}"{style_attr}><f{attrs}>{translated}</f></c>'


def _append_blank_row(xml: str) -> Tuple[str, int]:
    rows = _row_fragments(xml)
    candidates = [row for row in rows if row[0] >= HEADER_ROW]
    if not candidates:
        raise ArchiveRenderError("worksheet has no row that can provide styles for append")
    old_row, _start, _end, fragment = candidates[-1]
    new_row = old_row + 1
    if fragment.endswith("/>"):
        new_fragment = f'<row r="{new_row}"/>'
    else:
        open_tag = re.match(r"<row\b[^>]*>", fragment)
        if not open_tag:
            raise ArchiveRenderError("cannot parse row used for append")
        head = re.sub(r'\br="\d+"', f'r="{new_row}"', open_tag.group(0), count=1)
        cells = [
            _empty_cloned_cell(match.group(0), old_row, new_row)
            for match in _CELL_TAG.finditer(fragment)
        ]
        new_fragment = head + "".join(cells) + "</row>"
    if "</sheetData>" not in xml:
        raise ArchiveRenderError("worksheet has no sheetData terminator")
    xml = xml.replace("</sheetData>", new_fragment + "</sheetData>", 1)

    def grow_dimension(match: re.Match[str]) -> str:
        first_col, first_row, last_col, last_row = (
            match.group(1), int(match.group(2)), match.group(3), int(match.group(4))
        )
        return f'<dimension ref="{first_col}{first_row}:{last_col}{max(last_row, new_row)}"/>'

    xml = re.sub(
        r'<dimension\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>',
        grow_dimension,
        xml,
        count=1,
    )
    return xml, new_row


@dataclass
class SheetState:
    name: str
    member: str
    xml: str
    shared_strings: Sequence[str]
    headers: Dict[str, str] = field(default_factory=dict)
    header_by_col: Dict[str, str] = field(default_factory=dict)
    reserved: Dict[int, str] = field(default_factory=dict)
    indexes: Dict[str, Dict[str, List[int]]] = field(default_factory=dict)
    rows: Dict[int, str] = field(default_factory=dict)
    column_styles: Dict[str, Optional[str]] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self._refresh_rows()
        row = self.row_fragment(HEADER_ROW)
        if not row:
            raise ArchiveRenderError(f"{self.name}: header row {HEADER_ROW} missing")
        for cell in _CELL_TAG.finditer(row):
            ref = re.search(r'\br="([A-Z]{1,4})\d+"', cell.group(0))
            if not ref:
                continue
            value = _canonical_text(_cell_value(cell.group(0), self.shared_strings))
            if not value:
                continue
            col = ref.group(1)
            if value in self.headers and self.headers[value] != col:
                raise ArchiveRenderError(f"{self.name}: duplicate header {value!r}")
            self.headers[value] = col
            self.header_by_col[col] = value

    def _refresh_rows(self) -> None:
        self.rows = {
            number: fragment
            for number, _start, _end, fragment in _row_fragments(self.xml)
        }
        self.indexes.clear()
        self.column_styles.clear()

    def row_fragment(self, row_number: int) -> str:
        return self.rows.get(int(row_number), "")

    def cell_fragment(self, row_number: int, col: str) -> str:
        return _cell_fragment(self.row_fragment(row_number), f"{col}{row_number}")

    def value(self, row_number: int, col: str) -> Any:
        return _cell_value(self.cell_fragment(row_number, col), self.shared_strings)

    def row_numbers(self) -> List[int]:
        return sorted(self.rows)

    def style_for(self, col: str) -> Optional[str]:
        if col not in self.column_styles:
            self.column_styles[col] = ledger_writer.find_col_style(self.xml, col)
        return self.column_styles[col]

    def set_row(self, row_number: int, fragment: str) -> None:
        if int(row_number) not in self.rows:
            raise ArchiveRenderError(f"{self.name}: row {row_number} is not allocated")
        self.rows[int(row_number)] = fragment

    def materialize(self) -> str:
        """Apply all cached row edits to the worksheet XML in one linear pass."""

        fragments = _row_fragments(self.xml)
        parts: List[str] = []
        cursor = 0
        for number, start, end, original in fragments:
            parts.append(self.xml[cursor:start])
            parts.append(self.rows.get(number, original))
            cursor = end
        parts.append(self.xml[cursor:])
        self.xml = "".join(parts)
        return self.xml

    def index(self, col: str) -> Dict[str, List[int]]:
        if col not in self.indexes:
            built: Dict[str, List[int]] = {}
            for row in self.row_numbers():
                if row < FIRST_DATA_ROW:
                    continue
                value = _canonical_text(self.value(row, col))
                if value:
                    built.setdefault(value, []).append(row)
            self.indexes[col] = built
        return self.indexes[col]

    def find(self, col: str, value: Any) -> List[int]:
        return list(self.index(col).get(_canonical_text(value), []))

    def note_index(self, col: str, value: Any, row: int) -> None:
        text = _canonical_text(value)
        if text:
            values = self.index(col)
            if row not in values.setdefault(text, []):
                values[text].append(row)

    def ensure_capacity(self) -> int:
        self.materialize()
        self.xml, row = _append_blank_row(self.xml)
        self._refresh_rows()
        return row

    def allocate(self, owner: str, occupancy_cols: Sequence[str]) -> int:
        rows = [row for row in self.row_numbers() if row >= FIRST_DATA_ROW]
        for row in rows:
            if row in self.reserved:
                continue
            occupied = any(_canonical_text(self.value(row, col)) for col in occupancy_cols if col)
            if not occupied:
                self.reserved[row] = owner
                return row
        row = self.ensure_capacity()
        self.reserved[row] = owner
        return row


@dataclass
class LocatedRecord:
    record: Mapping[str, Any]
    sheet: SheetState
    row: int
    key_col: str
    inserted: bool


@dataclass
class CellExpectation:
    sheet: str
    member: str
    row: int
    col: str
    header: str
    value: Any
    value_type: str
    formula_preserved: bool = False


def _reverse_sheet_specs() -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    for sheet, spec in SHEET_SPECS.items():
        out.setdefault(str(spec.get("kind") or ""), []).append(sheet)
    return out


def _record_sheet(record: Mapping[str, Any], available: Mapping[str, str]) -> str:
    target = record.get("target") or {}
    if target.get("sheet"):
        sheet = str(target["sheet"])
        if sheet not in available:
            raise ArchiveRenderError(f"target worksheet missing: {sheet}")
        return sheet
    choices = [s for s in _reverse_sheet_specs().get(str(record.get("kind") or ""), []) if s in available]
    if len(choices) != 1:
        raise ArchiveRenderError(
            f"cannot route kind={record.get('kind')!r} to one worksheet; choices={choices}"
        )
    return choices[0]


def _translated_fields(record: Mapping[str, Any], sheet_name: str) -> Tuple[Dict[str, Any], Dict[str, Mapping[str, Any]], List[str]]:
    raw = dict(record.get("fields") or {})
    meta = dict(record.get("field_meta") or {})
    spec = SHEET_SPECS.get(sheet_name, {})
    translated: Dict[str, Any] = {}
    translated_meta: Dict[str, Mapping[str, Any]] = {}
    warnings: List[str] = []
    # Workbook-shaped fields are loaded first.  Canonical core columns then win
    # if an old shadow field retained a conflicting value.
    for name, value in raw.items():
        if name in {"public_id", "project_no", "camp_name", "status"} or str(name).startswith("__"):
            continue
        translated[str(name)] = value
        translated_meta[str(name)] = dict(meta.get(name) or {})
    for core in ("public_id", "project_no", "camp_name", "status"):
        header = str(spec.get(core) or "")
        if not header:
            continue
        value = raw.get(core)
        if header in translated and not _values_equal(translated[header], value, "text"):
            warnings.append(f"{record.get('business_key')}:{header} core value overrides shadow field")
        translated[header] = value
        translated_meta[header] = dict(meta.get(header) or {})
    return translated, translated_meta, warnings


def _business_key_column(record: Mapping[str, Any], state: SheetState) -> str:
    target = record.get("target") or {}
    requested = str(target.get("business_key_col") or "")
    if requested and requested in state.headers:
        return requested
    fields = dict(record.get("fields") or {})
    key = _canonical_text(record.get("business_key"))
    for header, value in fields.items():
        if header in state.headers and _canonical_text(value) == key:
            return str(header)
    spec = SHEET_SPECS.get(state.name, {})
    public_header = str(spec.get("public_id") or "")
    if public_header and public_header in state.headers:
        return public_header
    if "정산ID" in state.headers and key.startswith("JS-"):
        return "정산ID"
    project_header = str(spec.get("project_no") or "프로젝트NO")
    if project_header in state.headers:
        return project_header
    raise ArchiveRenderError(f"{state.name}: no business-key column for {key}")


def _candidate_rows(record: Mapping[str, Any], state: SheetState, key_col: str) -> List[int]:
    # 식별 probe(키·공개ID·프로젝트NO)는 한 업무를 유일하게 가리킨다.
    # 캠프명은 여러 행이 공유하는 **속성**이라 identity 를 좁히기만(narrow) 하고
    # 절대 넓히지(union) 않는다 — 그렇지 않으면 같은 캠프의 서로 다른 건들이
    # 한 키로 뭉쳐 '여러 행에 일치'로 보관본 생성이 통째로 막힌다(2026-08-10 실사고).
    spec = SHEET_SPECS.get(state.name, {})
    fields = dict(record.get("fields") or {})
    identity_probes: List[Tuple[str, Any]] = [(key_col, record.get("business_key"))]
    for core, raw_name in (
        ("public_id", "public_id"),
        ("project_no", "project_no"),
    ):
        header = str(spec.get(core) or "")
        value = fields.get(raw_name)
        if header and value not in (None, ""):
            identity_probes.append((header, value))
    camp_header = str(spec.get("camp_name") or "")
    camp_value = fields.get("camp_name")

    id_sets: List[set[int]] = []
    for header, value in identity_probes:
        col = state.headers.get(header)
        if not col or value in (None, ""):
            continue
        rows = set(state.find(col, value))
        if rows:
            id_sets.append(rows)

    if id_sets:
        intersection = set.intersection(*id_sets)
        base = intersection if intersection else set().union(*id_sets)
        if camp_header and camp_value not in (None, ""):
            col = state.headers.get(camp_header)
            if col:
                narrowed = base & set(state.find(col, camp_value))
                if narrowed:
                    base = narrowed
        return sorted(base)

    # 식별 신호가 없으면(새 기록 등) 예전 동작을 유지한다 — 캠프명 단독 매칭.
    found: set[int] = set()
    if camp_header and camp_value not in (None, ""):
        col = state.headers.get(camp_header)
        if col:
            found.update(state.find(col, camp_value))
    return sorted(found)


def _locate_records(
    records: Sequence[Mapping[str, Any]], states: Mapping[str, SheetState], template_sha: str
) -> Tuple[List[LocatedRecord], List[Dict[str, Any]]]:
    located: List[LocatedRecord] = []
    conflicts: List[Dict[str, Any]] = []
    # Located shadow rows are reserved before DB-only records consume empty rows.
    ordered = sorted(
        records,
        key=lambda rec: (
            1 if not rec.get("target") else 0,
            str((rec.get("target") or {}).get("sheet") or rec.get("kind") or ""),
            str(rec.get("business_key") or ""),
            str(rec.get("work_id") or ""),
        ),
    )
    available = {name: state.member for name, state in states.items()}
    for record in ordered:
        owner = str(record.get("work_id") or record.get("business_key") or "")
        sheet_name = _record_sheet(record, available)
        state = states[sheet_name]
        key_header = _business_key_column(record, state)
        key_col = state.headers[key_header]
        candidates = _candidate_rows(record, state, key_header)
        target = record.get("target") or {}
        source_row = int(target.get("source_row") or 0)
        source_sha = str((record.get("source") or {}).get("source_sha256") or "")
        anchored = bool(
            source_row
            and source_sha
            and source_sha == template_sha
            and state.row_fragment(source_row)
        )
        if len(candidates) > 1:
            # 원본 Excel 행 앵커가 있고 그 행이 후보에 들면 그것이 권위다.
            if anchored and source_row in candidates:
                candidates = [source_row]
            else:
                # 유일하게 못 붙는 기록(예: 밴드 출처 돌발AS — 프로젝트NO 하나에
                # Excel 건별 여러 행)은 기존 행을 건드리지 않고 충돌로 남기고 건너뛴다.
                # 행 하나를 골라 쓰면 엉뚱한 행에 값이 박힌다(빈칸보다 나쁨).
                # 한 건 때문에 보관본 회차 전체를 세우지 않는다 — 정본 DB엔 그대로 있다.
                conflicts.append(
                    {
                        "sheet": sheet_name,
                        "business_key": record.get("business_key"),
                        "work_id": record.get("work_id"),
                        "record_version": record.get("record_version"),
                        "rows": candidates[:20],
                        "reason": "ambiguous-no-anchor",
                    }
                )
                continue
        row: Optional[int] = candidates[0] if candidates else None
        inserted = False
        if row is None and anchored:
            row = source_row
        if row is None:
            spec = SHEET_SPECS.get(sheet_name, {})
            occupancy = [
                key_col,
                state.headers.get(str(spec.get("project_no") or ""), ""),
                state.headers.get(str(spec.get("camp_name") or ""), ""),
            ]
            try:
                row = state.allocate(owner, [c for c in occupancy if c])
            except ArchiveRenderError as exc:
                # 템플릿에 넣을 빈 행이 없다(마지막 행이 공유수식이라 복제 불가 포함).
                # 못 넣는 기록은 건너뛰고 충돌로 남긴다 — 보관본은 들어가는 것으로 완주한다.
                # _append_blank_row 는 실패 시 state.xml 을 안 바꾼다(부분 오염 없음).
                conflicts.append(
                    {
                        "sheet": sheet_name,
                        "business_key": record.get("business_key"),
                        "work_id": record.get("work_id"),
                        "record_version": record.get("record_version"),
                        "rows": [],
                        "reason": "template-full-cannot-insert",
                        "detail": str(exc),
                    }
                )
                continue
            inserted = True
        existing_owner = state.reserved.get(row)
        if existing_owner and existing_owner != owner:
            # 두 정본 기록이 같은 Excel 행을 가리킨다(예: 밴드 출처 여러 건이
            # 프로젝트NO 하나에 묶임). 먼저 잡은 쪽에 행을 두고 뒤 기록은 충돌로
            # 남기고 건너뛴다 — 덮으면 남의 값이 그 행에 박힌다(위 정책과 같다).
            conflicts.append(
                {
                    "sheet": sheet_name,
                    "business_key": record.get("business_key"),
                    "work_id": record.get("work_id"),
                    "record_version": record.get("record_version"),
                    "rows": [row],
                    "reason": "row-claimed-by-other",
                    "claimed_by": existing_owner,
                }
            )
            continue
        state.reserved[row] = owner
        state.note_index(key_col, record.get("business_key"), row)
        located.append(LocatedRecord(record, state, row, key_header, inserted))
    return located, conflicts


def _patch_cell(
    state: SheetState,
    row: int,
    header: str,
    value: Any,
    meta: Mapping[str, Any],
) -> Tuple[str, CellExpectation]:
    col = state.headers.get(header)
    if not col:
        raise ArchiveRenderError(f"{state.name}: workbook column missing: {header}")
    value_type = _value_type(header, value, meta)
    normalized = "" if value is None else value
    fragment = state.cell_fragment(row, col)
    if _formula(fragment):
        return "formula_preserved", CellExpectation(
            state.name, state.member, row, col, header, normalized, value_type, True
        )
    current = _cell_value(fragment, state.shared_strings)
    if _values_equal(current, normalized, value_type):
        return "unchanged", CellExpectation(
            state.name, state.member, row, col, header, normalized, value_type, False
        )
    row_fragment = state.row_fragment(row)
    if row_fragment.endswith("/>"):
        head = row_fragment[:-2] + ">"
        body, tail = "", "</row>"
    else:
        row_open = re.match(r"(<row\b[^>]*>)(.*)(</row>)$", row_fragment, re.S)
        if not row_open:
            raise ArchiveRenderError(f"{state.name}!row={row} cannot parse row XML")
        head, body, tail = row_open.groups()
    ref = f"{col}{row}"
    match = _cell_pattern(ref).search(body)
    style: Optional[str]
    if match:
        old_cell = match.group(0)
        style_match = re.search(r'\bs="([^"]+)"', old_cell)
        style = style_match.group(1) if style_match else state.style_for(col)
        new_cell = ledger_writer.cell_xml(col, row, style, normalized, value_type)
        body = body[: match.start()] + new_cell + body[match.end() :]
    else:
        style = state.style_for(col)
        new_cell = ledger_writer.cell_xml(col, row, style, normalized, value_type)
        target_number = ledger_writer.col_num(col)
        insert_at = len(body)
        for cell in _CELL_TAG.finditer(body):
            cell_ref = re.search(r'\br="([A-Z]{1,4})\d+"', cell.group(0))
            if cell_ref and ledger_writer.col_num(cell_ref.group(1)) > target_number:
                insert_at = cell.start()
                break
        body = body[:insert_at] + new_cell + body[insert_at:]
    state.set_row(row, head + body + tail)
    return "applied", CellExpectation(
        state.name, state.member, row, col, header, normalized, value_type, False
    )


def _force_full_recalc(name: str, data: bytes) -> Optional[bytes]:
    if name == "xl/calcChain.xml":
        return None
    if name == "[Content_Types].xml":
        text = data.decode("utf-8")
        text = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", text)
        return text.encode("utf-8")
    if name == "xl/_rels/workbook.xml.rels":
        text = data.decode("utf-8")
        text = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", text)
        return text.encode("utf-8")
    if name == "xl/workbook.xml":
        text = data.decode("utf-8")
        if "<calcPr" in text:
            match = re.search(r"<calcPr\b[^>]*/?>", text)
            if match:
                tag = match.group(0)
                for attr, value in (
                    ("calcMode", "auto"),
                    ("fullCalcOnLoad", "1"),
                    ("forceFullCalc", "1"),
                ):
                    if re.search(rf'\b{attr}="[^"]*"', tag):
                        tag = re.sub(rf'\b{attr}="[^"]*"', f'{attr}="{value}"', tag)
                    else:
                        tag = tag.replace("<calcPr", f'<calcPr {attr}="{value}"', 1)
                text = text[: match.start()] + tag + text[match.end() :]
        else:
            text = text.replace(
                "</workbook>",
                '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
                1,
            )
        return text.encode("utf-8")
    return data


def _write_patched_zip(
    template: Path,
    output: Path,
    patched_members: Mapping[str, str],
) -> None:
    output = _require_local_write(output, "archive adapter output")
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        raise FileExistsError(f"adapter output already exists: {output}")
    with zipfile.ZipFile(template, "r") as source, zipfile.ZipFile(
        output, "w", allowZip64=True
    ) as destination:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename in patched_members:
                data = patched_members[info.filename].encode("utf-8")
            data = _force_full_recalc(info.filename, data)
            if data is None:
                continue
            cloned = copy.copy(info)
            destination.writestr(cloned, data)
        for name, value in patched_members.items():
            if name not in source.namelist():
                destination.writestr(name, value.encode("utf-8"))


def _inline_cell(col: str, row: int, value: str) -> str:
    escaped = html.escape(value, quote=False)
    preserve = ' xml:space="preserve"' if value != value.strip() else ""
    return (
        f'<c r="{col}{row}" t="inlineStr"><is><t{preserve}>'
        f"{escaped}</t></is></c>"
    )


def _sidecar_payload(
    records_by_id: Mapping[str, Mapping[str, Any]],
    conflicts: Sequence[Mapping[str, Any]],
) -> Tuple[str, List[Dict[str, Any]], str]:
    """Return a deterministic audit worksheet for records unsafe for main sheets.

    The canonical record is stored in full, split across cells only to respect
    Excel's 32,767 character cell limit.  Main worksheet rows remain untouched.
    """

    rows: List[Dict[str, Any]] = []
    max_chunks = 1
    for conflict in sorted(
        conflicts,
        key=lambda item: (
            str(item.get("business_key") or ""),
            str(item.get("work_id") or ""),
        ),
    ):
        work_id = str(conflict.get("work_id") or "")
        record = records_by_id.get(work_id)
        if not record:
            raise ArchiveVerificationError(
                f"sidecar conflict references missing plan record: {work_id}"
            )
        record_json = canonical_json(record)
        chunks = [
            record_json[index : index + SIDECAR_CHUNK_SIZE]
            for index in range(0, len(record_json), SIDECAR_CHUNK_SIZE)
        ] or [""]
        max_chunks = max(max_chunks, len(chunks))
        record_sha = hashlib.sha256(record_json.encode("utf-8")).hexdigest()
        conflict_json = canonical_json(dict(conflict))
        rows.append(
            {
                "work_id": work_id,
                "record_version": int(record.get("record_version") or 0),
                "business_key": str(record.get("business_key") or ""),
                "kind": str(record.get("kind") or ""),
                "target_sheet": str(conflict.get("sheet") or ""),
                "reason": str(conflict.get("reason") or ""),
                "conflict_json": conflict_json,
                "record_sha256": record_sha,
                "chunks": chunks,
            }
        )
    headers = [
        "work_id",
        "record_version",
        "business_key",
        "DB업무종류",
        "주 시트 대상",
        "주 시트 미반영 사유",
        "충돌 근거(JSON)",
        "정본 레코드 SHA-256",
    ] + [f"정본 레코드 JSON {index:03d}" for index in range(1, max_chunks + 1)]
    xml_rows: List[str] = []
    for row_no, values in enumerate(
        [headers]
        + [
            [
                item["work_id"],
                str(item["record_version"]),
                item["business_key"],
                item["kind"],
                item["target_sheet"],
                item["reason"],
                item["conflict_json"],
                item["record_sha256"],
                *item["chunks"],
            ]
            for item in rows
        ],
        1,
    ):
        cells = "".join(
            _inline_cell(ledger_writer.col_letter(index), row_no, str(value))
            for index, value in enumerate(values, 1)
        )
        xml_rows.append(f'<row r="{row_no}">{cells}</row>')
    last_col = ledger_writer.col_letter(len(headers))
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_MAIN_NS}">'
        f'<dimension ref="A1:{last_col}{max(1, len(rows) + 1)}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" '
        'topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        f'<autoFilter ref="A1:{last_col}{max(1, len(rows) + 1)}"/>'
        '</worksheet>'
    )
    semantic_rows = [
        {
            "work_id": item["work_id"],
            "record_version": item["record_version"],
            "business_key": item["business_key"],
            "reason": item["reason"],
            "record_sha256": item["record_sha256"],
        }
        for item in rows
    ]
    return sheet_xml, rows, sha256_json(
        {"format": SIDECAR_FORMAT, "records": semantic_rows}
    )


def _install_sidecar(
    template: Path, sheet_xml: str
) -> Tuple[Dict[str, str], str]:
    """Return OOXML patches that add or replace the deterministic sidecar."""

    with zipfile.ZipFile(template, "r") as archive:
        names = set(archive.namelist())
        sheet_map = ledger_writer.sheet_file_map(archive)
        if SIDECAR_SHEET in sheet_map:
            return {sheet_map[SIDECAR_SHEET]: sheet_xml}, sheet_map[SIDECAR_SHEET]
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
        rel_name = "xl/_rels/workbook.xml.rels"
        rels = archive.read(rel_name).decode("utf-8")
        types = archive.read("[Content_Types].xml").decode("utf-8")
    sheet_numbers = [
        int(match.group(1))
        for name in names
        if (match := re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name))
    ]
    sheet_no = max(sheet_numbers or [0]) + 1
    member = f"xl/worksheets/sheet{sheet_no}.xml"
    rel_numbers = [int(value) for value in re.findall(r'\bId="rId(\d+)"', rels)]
    rel_id = f"rId{max(rel_numbers or [0]) + 1}"
    sheet_ids = [int(value) for value in re.findall(r'\bsheetId="(\d+)"', workbook)]
    sheet_id = max(sheet_ids or [0]) + 1
    # ★ `r:id` 를 쓰는 요소에는 그 접두사 **선언이 같이 있어야** 한다.
    #   2026-08-20 실사고: 오늘 깔린 lxml 6.1.2 로 openpyxl 이 제 직렬화기 대신
    #   lxml 을 쓰게 되면서 `xmlns:r` 을 **루트가 아니라 <sheet> 요소마다** 적는다.
    #   그래서 여기서 덧붙인 <sheet> 에만 선언이 없어 **워크북이 통째로** 안 읽혔다
    #   (`XMLSyntaxError: Namespace prefix r for id on sheet is not defined`).
    #   같은 뿌리로 findings_sheet · reorder_rows 도 함께 깨졌다 — 코드는 한 줄도
    #   안 바뀌었는데 관문이 세 곳에서 죽었다. **같은 URI 재선언은 XML 상 유효**하므로
    #   엑셀이 만든 정본(루트에 선언이 있는 파일)에서도 안전하다.
    _R_NS = ("http://schemas.openxmlformats.org/officeDocument"
             "/2006/relationships")
    sheet_tag = (
        f'<sheet xmlns:r="{_R_NS}" name="{html.escape(SIDECAR_SHEET, quote=True)}" '
        f'sheetId="{sheet_id}" r:id="{rel_id}"/>'
    )
    if "</sheets>" not in workbook:
        raise ArchiveRenderError("workbook.xml has no sheets collection")
    workbook = workbook.replace("</sheets>", f"{sheet_tag}</sheets>", 1)
    relation = (
        f'<Relationship Id="{rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{sheet_no}.xml"/>'
    )
    if "</Relationships>" not in rels:
        raise ArchiveRenderError("workbook relationships are malformed")
    rels = rels.replace("</Relationships>", f"{relation}</Relationships>", 1)
    override = (
        f'<Override PartName="/{member}" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    )
    if "</Types>" not in types:
        raise ArchiveRenderError("content types are malformed")
    types = types.replace("</Types>", f"{override}</Types>", 1)
    return {
        "xl/workbook.xml": workbook,
        rel_name: rels,
        "[Content_Types].xml": types,
        member: sheet_xml,
    }, member


def _validate_sidecar(
    output: Path, rows: Sequence[Mapping[str, Any]], semantic_sha256: str
) -> Dict[str, Any]:
    with zipfile.ZipFile(output, "r") as archive:
        sheet_map = ledger_writer.sheet_file_map(archive)
        member = sheet_map.get(SIDECAR_SHEET)
        if not member:
            raise ArchiveVerificationError("AppDB sidecar worksheet is missing")
        xml = archive.read(member).decode("utf-8")
        shared = _shared_strings(archive)
        fragments = {
            number: fragment for number, _a, _b, fragment in _row_fragments(xml)
        }
        observed: List[Dict[str, Any]] = []
        for row_no, expected in enumerate(rows, 2):
            fragment = fragments.get(row_no, "")
            fixed = [
                str(_cell_value(_cell_fragment(fragment, f"{col}{row_no}"), shared))
                for col in ("A", "B", "C", "D", "E", "F", "G", "H")
            ]
            chunks: List[str] = []
            col_no = 9
            while True:
                value = str(
                    _cell_value(
                        _cell_fragment(
                            fragment, f"{ledger_writer.col_letter(col_no)}{row_no}"
                        ),
                        shared,
                    )
                )
                if not value:
                    break
                chunks.append(value)
                col_no += 1
            record_json = "".join(chunks)
            record_sha = hashlib.sha256(record_json.encode("utf-8")).hexdigest()
            if fixed != [
                str(expected["work_id"]),
                str(expected["record_version"]),
                str(expected["business_key"]),
                str(expected["kind"]),
                str(expected["target_sheet"]),
                str(expected["reason"]),
                str(expected["conflict_json"]),
                str(expected["record_sha256"]),
            ] or record_sha != str(expected["record_sha256"]):
                raise ArchiveVerificationError(
                    f"sidecar semantic mismatch at row {row_no}: {expected['work_id']}"
                )
            observed.append(
                {
                    "work_id": fixed[0],
                    "record_version": int(fixed[1]),
                    "business_key": fixed[2],
                    "reason": fixed[5],
                    "record_sha256": record_sha,
                }
            )
    actual = sha256_json({"format": SIDECAR_FORMAT, "records": observed})
    if actual != semantic_sha256:
        raise ArchiveVerificationError("sidecar semantic hash mismatch")
    return {
        "sheet": SIDECAR_SHEET,
        "records": len(observed),
        "semantic_sha256": actual,
    }


def _validate_output(
    output: Path,
    expectations: Sequence[CellExpectation],
    expected_sheets: Iterable[str],
) -> Dict[str, Any]:
    errors: List[str] = []
    checked = 0
    formulas = 0
    with zipfile.ZipFile(output, "r") as archive:
        bad = archive.testzip()
        if bad:
            errors.append(f"ZIP CRC failed: {bad}")
        names = set(archive.namelist())
        if "xl/calcChain.xml" in names:
            errors.append("calcChain was not removed")
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        if not re.search(r'<calcPr\b[^>]*\bfullCalcOnLoad="1"', workbook_xml):
            errors.append("fullCalcOnLoad=1 missing")
        shared = _shared_strings(archive)
        sheet_map = ledger_writer.sheet_file_map(archive)
        for sheet in expected_sheets:
            if sheet not in sheet_map:
                errors.append(f"worksheet missing after render: {sheet}")
        xml_cache = {
            member: archive.read(member).decode("utf-8")
            for member in {item.member for item in expectations}
            if member in names
        }
        row_cache = {
            member: {
                number: fragment
                for number, _start, _end, fragment in _row_fragments(xml)
            }
            for member, xml in xml_cache.items()
        }
        for item in expectations:
            xml = xml_cache.get(item.member)
            if xml is None:
                errors.append(f"worksheet XML missing: {item.member}")
                continue
            row_fragment = row_cache.get(item.member, {}).get(item.row, "")
            fragment = _cell_fragment(row_fragment, f"{item.col}{item.row}")
            if item.formula_preserved:
                formulas += 1
                if not _formula(fragment):
                    errors.append(
                        f"formula lost: {item.sheet}!{item.col}{item.row} ({item.header})"
                    )
                continue
            current = _cell_value(fragment, shared)
            checked += 1
            if not _values_equal(current, item.value, item.value_type):
                errors.append(
                    f"semantic mismatch: {item.sheet}!{item.col}{item.row} "
                    f"expected={item.value!r} actual={current!r}"
                )
            if len(errors) >= 100:
                break
    if errors:
        raise ArchiveVerificationError("; ".join(errors[:100]))
    return {
        "zip_crc": True,
        "semantic_cells_checked": checked,
        "formula_cells_preserved": formulas,
        "full_recalc": True,
    }


class ArchiveWorker:
    """Production adapter and orchestration boundary for local Excel archives."""

    def __init__(
        self,
        store: AppStore,
        spool_dir: Optional[os.PathLike[str] | str] = None,
    ) -> None:
        self.exporter = ArchiveExporter(store, spool_dir)
        self.store = self.exporter.store
        self.spool_dir = self.exporter.spool_dir
        self.status_path = self.spool_dir / "worker-status.json"

    def adapter(
        self,
        plan_path: Path,
        template_copy_path: Path,
        output_path: Path,
    ) -> Mapping[str, Any]:
        plan_path = _require_local_write(plan_path, "command plan")
        template_copy_path = _require_local_write(template_copy_path, "template copy")
        output_path = _require_local_write(output_path, "adapter output")
        template_before = sha256_file(template_copy_path)
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        if plan.get("format") != PLAN_FORMAT:
            raise ArchiveRenderError(f"unsupported command plan: {plan.get('format')}")
        if template_before != str(plan.get("template", {}).get("sha256") or ""):
            raise ArchiveVerificationError("template copy hash differs from command plan")
        records = list(plan.get("records") or [])
        with zipfile.ZipFile(template_copy_path, "r") as archive:
            if archive.testzip() is not None:
                raise ArchiveVerificationError("template ZIP CRC failed")
            shared = _shared_strings(archive)
            sheet_map = ledger_writer.sheet_file_map(archive)
            required_sheets = {
                _record_sheet(record, sheet_map) for record in records
            }
            states = {
                name: SheetState(
                    name,
                    sheet_map[name],
                    archive.read(sheet_map[name]).decode("utf-8"),
                    shared,
                )
                for name in sorted(required_sheets)
            }
        located, locate_conflicts = _locate_records(records, states, template_before)
        counts = {
            "records_inserted": 0,
            "records_updated": 0,
            "commands_applied": 0,
            "commands_unchanged": 0,
            "formula_cells_preserved": 0,
        }
        warnings: List[str] = []
        record_coverage: List[Dict[str, Any]] = []
        for conflict in locate_conflicts:
            reason = conflict.get("reason", "")
            if reason == "template-full-cannot-insert":
                detail = "template full — no empty row to insert (expand template)"
            elif reason == "row-claimed-by-other":
                detail = f"row {conflict.get('rows')} already claimed by another record"
            else:
                detail = (
                    f"matches multiple rows {conflict.get('rows')} "
                    f"(no unique anchor; existing rows untouched)"
                )
            warnings.append(
                f"{conflict['sheet']}:{conflict.get('business_key')} "
                f"main sheet untouched; archived in {SIDECAR_SHEET} — {detail}"
            )
        expectations: List[CellExpectation] = []
        for item in located:
            fields, meta, field_warnings = _translated_fields(item.record, item.sheet.name)
            warnings.extend(field_warnings)
            record_counts = {
                "commands_applied": 0,
                "commands_unchanged": 0,
                "formula_cells_preserved": 0,
            }
            # A DB-only record must carry its canonical key even if the caller did
            # not duplicate that key inside fields.
            fields.setdefault(item.key_col, item.record.get("business_key"))
            meta.setdefault(item.key_col, {})
            for header in sorted(fields):
                if header not in item.sheet.headers:
                    if str(header).startswith("__") or header in DB_ONLY_ARCHIVE_FIELDS:
                        if header in DB_ONLY_ARCHIVE_FIELDS:
                            warnings.append(
                                f"{item.record.get('business_key')}:{header} retained in DB audit only"
                            )
                        continue
                    raise ArchiveRenderError(
                        f"{item.sheet.name}:{item.record.get('business_key')} "
                        f"cannot archive unknown field {header!r}"
                    )
                result, expectation = _patch_cell(
                    item.sheet, item.row, header, fields[header], meta.get(header) or {}
                )
                expectations.append(expectation)
                if result == "applied":
                    counts["commands_applied"] += 1
                    record_counts["commands_applied"] += 1
                elif result == "unchanged":
                    counts["commands_unchanged"] += 1
                    record_counts["commands_unchanged"] += 1
                else:
                    counts["formula_cells_preserved"] += 1
                    record_counts["formula_cells_preserved"] += 1
            if item.inserted:
                counts["records_inserted"] += 1
            else:
                counts["records_updated"] += 1
            # This is the semantic acknowledgement boundary.  Merely seeing a
            # record in the command plan is not coverage: it must have been
            # located/allocated and every writable cell must have passed the
            # output verifier below.  Conflicted records never enter this list.
            record_coverage.append(
                {
                    "work_id": str(item.record.get("work_id") or ""),
                    "business_key": str(item.record.get("business_key") or ""),
                    "record_version": int(item.record.get("record_version") or 0),
                    "sheet": item.sheet.name,
                    "row": int(item.row),
                    "outcome": (
                        "applied"
                        if record_counts["commands_applied"] or item.inserted
                        else "unchanged"
                    ),
                    **record_counts,
                }
            )
        records_by_id = {
            str(record.get("work_id") or ""): record for record in records
        }
        sidecar_xml, sidecar_rows, sidecar_sha = _sidecar_payload(
            records_by_id, locate_conflicts
        )
        sidecar_patches, _sidecar_member = _install_sidecar(
            template_copy_path, sidecar_xml
        )
        for row_no, item in enumerate(sidecar_rows, 2):
            record_coverage.append(
                {
                    "work_id": str(item["work_id"]),
                    "business_key": str(item["business_key"]),
                    "record_version": int(item["record_version"]),
                    "sheet": SIDECAR_SHEET,
                    "row": row_no,
                    "outcome": "archived_sidecar",
                    "reason": str(item["reason"]),
                    "record_sha256": str(item["record_sha256"]),
                    "sidecar_semantic_sha256": sidecar_sha,
                }
            )
        for state in states.values():
            state.materialize()
        patched_members = {state.member: state.xml for state in states.values()}
        patched_members.update(sidecar_patches)
        _write_patched_zip(
            template_copy_path,
            output_path,
            patched_members,
        )
        validation = _validate_output(output_path, expectations, states)
        validation["sidecar"] = _validate_sidecar(
            output_path, sidecar_rows, sidecar_sha
        )
        if sha256_file(template_copy_path) != template_before:
            raise ArchiveVerificationError("adapter mutated the template copy")
        return {
            "format": WORKER_PROOF_FORMAT,
            "status": "success",
            "snapshot_sha256": str(plan["snapshot"]["sha256"]),
            "command_plan_sha256": sha256_json(plan),
            "rows_considered": len(records),
            "records_skipped_ambiguous": len(locate_conflicts),
            "records_archived_sidecar": len(sidecar_rows),
            "records_covered": len(record_coverage),
            "record_coverage": record_coverage,
            # Unsafe main-sheet matches are not hidden or overwritten.  Their
            # full canonical records and conflict evidence live in the verified
            # sidecar, which is a positive coverage outcome rather than an
            # unresolved conflict.
            "conflicts": [],
            "sidecar": {
                "format": SIDECAR_FORMAT,
                "sheet": SIDECAR_SHEET,
                "records": len(sidecar_rows),
                "semantic_sha256": sidecar_sha,
                "entries": [
                    {
                        "work_id": str(item["work_id"]),
                        "record_version": int(item["record_version"]),
                        "business_key": str(item["business_key"]),
                        "reason": str(item["reason"]),
                        "record_sha256": str(item["record_sha256"]),
                    }
                    for item in sidecar_rows
                ],
            },
            "errors": [],
            "warnings": warnings[:500],
            "output_sha256": sha256_file(output_path),
            "template_sha256_before": template_before,
            "template_sha256_after": sha256_file(template_copy_path),
            "external_write_performed": False,
            **counts,
            "validation": validation,
        }

    def _write_state(self, state: Mapping[str, Any]) -> Dict[str, Any]:
        payload = {"format": WORKER_STATUS_FORMAT, **dict(state)}
        _atomic_json(self.status_path, payload)
        return payload

    def run(
        self,
        template_path: os.PathLike[str] | str,
        *,
        wait_seconds: float = 0.0,
    ) -> Dict[str, Any]:
        template_source = str(template_path)
        started = _utcnow()
        source_before = ""
        source_proof: Mapping[str, Any] = {}
        with _worker_lock(self.spool_dir, wait_seconds=wait_seconds):
            self._write_state(
                {
                    "state": "running",
                    "phase": "staging_template",
                    "ok": None,
                    "started_at": started,
                    "pid": os.getpid(),
                    "template_source": template_source,
                    "external_write_performed": False,
                }
            )
            try:
                with _staged_template(template_source, self.spool_dir) as (
                    template,
                    source_proof,
                ):
                    source_before = sha256_file(template)
                    source_proof = {**dict(source_proof), "sha256": source_before}

                    def bound_adapter(
                        plan_path: Path,
                        template_copy_path: Path,
                        output_path: Path,
                    ) -> Mapping[str, Any]:
                        proof = dict(
                            self.adapter(plan_path, template_copy_path, output_path)
                        )
                        proof["template_source"] = template_source
                        proof["source_copy_proof"] = dict(source_proof)
                        return proof

                    self._write_state(
                        {
                            "state": "running",
                            "phase": "rendering_archive",
                            "ok": None,
                            "started_at": started,
                            "pid": os.getpid(),
                            "template_source": template_source,
                            "template_local": str(template),
                            "template_sha256_before": source_before,
                            "source_copy_proof": source_proof,
                            "external_write_performed": False,
                        }
                    )
                    result = self.exporter.run_local(
                        template_path=template,
                        adapter=bound_adapter,
                        dry_run=False,
                    )
                    export_status = str(result.get("status") or "")
                    verification = self.exporter.verify_export(
                        result["artifact_dir"],
                        require_verified=export_status == "verified",
                    )
                    if not verification.get("ok"):
                        raise ArchiveVerificationError(
                            f"archive artifact failed recheck: {verification.get('errors')}"
                        )
                    if export_status not in {"verified", "partial"}:
                        raise ArchiveVerificationError(
                            f"archive adapter ended in unsupported status: {export_status}"
                        )
                    # A partial workbook is useful forensic output, but it must
                    # never replace the last fully covered archive.  Keeping the
                    # previous pointer also makes the dashboard state truthful.
                    last_good = self.exporter.last_good(verify=True)
                    source_after = sha256_file(template)
                    if source_after != source_before:
                        raise ArchiveVerificationError(
                            "stable template changed during archive run; retry from source"
                        )
                    payload = self._write_state(
                        {
                            "state": export_status,
                            "phase": "complete",
                            "ok": True,
                            "started_at": started,
                            "finished_at": _utcnow(),
                            "pid": os.getpid(),
                            "template_source": template_source,
                            "template_local": str(template),
                            "template_sha256_before": source_before,
                            "template_sha256_after": source_after,
                            "source_copy_proof": source_proof,
                            "export": result,
                            "verification": verification,
                            "last_good": last_good,
                            "external_write_performed": False,
                        }
                    )
                    return payload
            except Exception as exc:
                payload = self._write_state(
                    {
                        "state": "failed",
                        "phase": "failed",
                        "ok": False,
                        "started_at": started,
                        "finished_at": _utcnow(),
                        "pid": os.getpid(),
                        "template_source": template_source,
                        "template_sha256_before": source_before,
                        "source_copy_proof": source_proof,
                        "error_type": type(exc).__name__,
                        "error": str(exc)[:4_000],
                        "retryable": isinstance(
                            exc,
                            (OSError, TimeoutError, ArchiveWorkerBusy, ArchiveSourceError),
                        ),
                        "traceback": traceback.format_exc()[-8_000:],
                        "external_write_performed": False,
                    }
                )
                setattr(exc, "archive_worker_status", payload)
                raise

    def status(self) -> Dict[str, Any]:
        try:
            worker = json.loads(self.status_path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            worker = {
                "format": WORKER_STATUS_FORMAT,
                "state": "idle",
                "ok": None,
            }
        last_good: Optional[Mapping[str, Any]] = None
        last_good_error = ""
        try:
            last_good = self.exporter.last_good(verify=True)
        except Exception as exc:
            last_good_error = f"{type(exc).__name__}: {exc}"[:1_000]
        store_status = self.store.status()
        current = bool(
            last_good
            and int(last_good.get("snapshot_seq") or -1)
            == int(store_status.get("change_seq") or -2)
        )
        return {
            "ok": not bool(last_good_error),
            "worker": worker,
            "app_store": store_status,
            "last_good": last_good,
            "last_good_error": last_good_error,
            "archive_current": current,
            "spool_dir": str(self.spool_dir),
            "external_write_performed": False,
        }


def _resolve_template(
    store: AppStore,
    supplied: Optional[str],
    spool_dir: Optional[os.PathLike[str] | str] = None,
) -> Path:
    explicit: List[str] = []
    preferred: List[str] = []
    network_fallbacks: List[str] = []
    if supplied:
        explicit.append(supplied)
    env = str(os.environ.get("COUPANG_ARCHIVE_TEMPLATE") or "").strip()
    if env:
        explicit.append(env)
    try:
        source = store.get_setting("legacy_excel_source").get("value") or {}
        if isinstance(source, Mapping) and source.get("path"):
            raw_source = str(source["path"])
            if _is_network_or_z(raw_source):
                network_fallbacks.append(raw_source)
            else:
                preferred.append(raw_source)
    except Exception:
        pass
    # Once a verified local archive exists it is the safest next template: all
    # canonical fields are patched again, so chaining last-good is deterministic
    # and avoids touching SMB on every run.
    try:
        pointer = ArchiveExporter(store, spool_dir).last_good(verify=True)
        if pointer and pointer.get("archive_path"):
            preferred.append(str(pointer["archive_path"]))
    except Exception:
        pass
    try:
        from ecount_client import load_config

        configured = str(load_config()["reconcile"]["master_xlsx"])
        if _is_network_or_z(configured):
            network_fallbacks.append(configured)
        else:
            preferred.append(configured)
            parent = Path(configured).parent
            if parent.is_dir():
                versioned: List[Tuple[int, Path]] = []
                for path in parent.glob("쿠팡_통합업무_일일보고_관리대장_v*.xlsx"):
                    match = re.search(r"_v(\d+)\.xlsx$", path.name)
                    if match and not path.name.startswith("~$"):
                        versioned.append((int(match.group(1)), path))
                if versioned:
                    preferred.append(str(max(versioned, key=lambda item: item[0])[1]))
    except Exception:
        pass
    seen: set[str] = set()
    for raw in explicit:
        key = os.path.normcase(str(raw))
        if key in seen:
            continue
        seen.add(key)
        if _is_network_or_z(raw):
            # Existence is checked by the bounded source-stage child.  Calling
            # is_file()/resolve() here can itself hang forever on SMB.
            return Path(raw)
        path = Path(raw)
        if path.is_file():
            return path.resolve()
    for raw in preferred:
        key = os.path.normcase(os.path.abspath(str(raw)))
        if key in seen:
            continue
        seen.add(key)
        path = Path(raw)
        if path.is_file():
            return path.resolve()
    for raw in network_fallbacks:
        key = os.path.normcase(str(raw))
        if key not in seen:
            return Path(raw)
    raise FileNotFoundError(
        "no archive template found; set --template or COUPANG_ARCHIVE_TEMPLATE"
    )


def status(
    *,
    root: Optional[os.PathLike[str] | str] = None,
    store: Optional[AppStore] = None,
    db_path: Optional[os.PathLike[str] | str] = None,
    spool_dir: Optional[os.PathLike[str] | str] = None,
) -> Dict[str, Any]:
    """Return the compact archive status contract consumed by the app.

    ``ArchiveWorker.status`` deliberately exposes the complete verification
    proof.  The dashboard needs a stable, shallow shape instead.  Keeping the
    adapter here prevents the pipeline from guessing at nested exporter fields
    and, importantly, lets synthetic callers inject their temporary store.
    """

    if store is None:
        if db_path is not None:
            store = AppStore(db_path).initialize()
        elif root is not None:
            store = AppStore(Path(root) / "db" / "app_store.db").initialize()
        else:
            store = default_store().initialize()
    worker = ArchiveWorker(store, spool_dir)
    raw = worker.status()
    last_good = dict(raw.get("last_good") or {})
    worker_state = dict(raw.get("worker") or {})
    error = str(raw.get("last_good_error") or worker_state.get("error") or "")
    if error:
        state = "error"
    elif not last_good:
        state = str(worker_state.get("state") or "missing")
        if state == "idle":
            state = "missing"
    elif raw.get("archive_current"):
        state = "verified"
    else:
        state = "stale"
    return {
        "status": state,
        "last_good_at": (
            last_good.get("verified_at")
            or last_good.get("finished_at")
            or worker_state.get("finished_at")
        ),
        "snapshot_seq": last_good.get("snapshot_seq"),
        "path": last_good.get("archive_path") or "",
        "error": error,
        "archive_current": bool(raw.get("archive_current")),
        "worker_state": worker_state.get("state") or "idle",
        "external_write_performed": False,
    }


def _make_selftest_xlsx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheets><sheet name="02_돌발AS접수" sheetId="1" r:id="rId1"/></sheets>
 <calcPr calcId="191029"/>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""

    def text_cell(ref: str, value: str, style: str = "") -> str:
        s = f' s="{style}"' if style else ""
        return f'<c r="{ref}"{s} t="inlineStr"><is><t>{html.escape(value)}</t></is></c>'

    headers = ["접수ID", "프로젝트NO", "캠프명", "진행상태", "신청내용", "작업완료일"]
    header_xml = "".join(
        text_cell(f"{ledger_writer.col_letter(i)}4", value) for i, value in enumerate(headers, 1)
    )
    row5 = "".join(
        (
            text_cell("A5", "AS-2608-001"),
            text_cell("B5", "UJ2609001"),
            text_cell("C5", "기존캠프"),
            '<c r="D5" t="str"><f>IF(B5="","","접수")</f><v>접수</v></c>',
            text_cell("E5", "이전 내용"),
            '<c r="F5"/>',
        )
    )
    row6 = "".join(
        (
            '<c r="A6" t="str"><f>IF(B6="","","AS-2608-002")</f><v></v></c>',
            '<c r="B6"/>',
            '<c r="C6"/>',
            '<c r="D6" t="str"><f>IF(B6="","","접수")</f><v></v></c>',
            '<c r="E6"/>',
            '<c r="F6"/>',
        )
    )
    sheet = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <dimension ref="A1:F6"/><sheetData>
  <row r="4">{header_xml}</row>
  <row r="5">{row5}</row>
  <row r="6">{row6}</row>
 </sheetData>
</worksheet>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)


def self_test() -> Dict[str, Any]:
    """Exercise existing-row update, DB-only allocation, proof and idempotency."""

    with tempfile.TemporaryDirectory(prefix="csos-archive-worker-") as folder:
        base = Path(folder)
        template = base / "template.xlsx"
        _make_selftest_xlsx(template)
        template_hash = sha256_file(template)
        staged_path: Optional[Path] = None
        with _staged_template(template, base / "source-spool", force_child=True) as (
            staged_path,
            stage_proof,
        ):
            assert staged_path != template and staged_path.is_file()
            assert stage_proof.get("staged") is True
            assert sha256_file(staged_path) == template_hash
        assert staged_path is not None and not staged_path.exists()
        assert sha256_file(template) == template_hash
        store = AppStore(base / "app.db").initialize()
        imported = store.shadow_import(
            import_id="selftest-import",
            sheet="02_돌발AS접수",
            business_key="AS-2608-001",
            business_key_col="접수ID",
            row_number=5,
            kind="돌발AS",
            public_id="AS-2608-001",
            project_no="UJ2609001",
            camp_name="기존캠프",
            status="접수",
            fields={
                "접수ID": "AS-2608-001",
                "프로젝트NO": "UJ2609001",
                "캠프명": "기존캠프",
                "진행상태": "접수",
                "신청내용": "이전 내용",
                "작업완료일": "",
            },
            source_file=str(template),
            source_sha256=template_hash,
            apply_if_missing=True,
            idempotency_key="selftest-shadow",
        )
        existing = store.get_work(work_id=str(imported["work_id"]))
        store.update_work(
            existing["id"],
            expected_version=int(existing["record_version"]),
            patch={
                "fields": {
                    "신청내용": "DB에서 변경",
                    "객관완료여부": True,
                    "객관완료일": "2026-08-10",
                    "객관완료상태": "완료",
                    "객관완료근거": "합성 객관근거",
                    "접수취소여부": "예",
                    "접수취소사유": "유선전화 원격해결",
                    "접수취소확인일": "2026-08-10",
                    "처리구분": "원격해결",
                    "접수취소근거": "밴드·카톡 합성 근거",
                    "공급가액": "1000000",
                    # 담당자가 대표 캡처에서 적는 미처리 사유([321]). 아래 assert 가
                    # **DB_ONLY_ARCHIVE_FIELDS 전부**를 여기서 쓰라고 요구한다 —
                    # 이름만 늘리고 시험을 안 늘리면 그 필드는 한 번도 안 지나간다.
                    "미처리사유(담당자)": "부품 입고 8/25 예정 · 합성",
                    # 청구 제외(다녀옴 · 이 건으로는 청구 안 함). 위 assert 가 표 전부를
                    # 여기서 쓰라고 요구하므로 표에 올린 칸은 반드시 한 번 지나가야 한다.
                    "청구제외": "예",
                    "청구제외사유": "캠프 부담 — 합성",
                }
            },
            actor="selftest",
            source="selftest",
            evidence="archive worker existing-row test",
            idempotency_key="selftest-update",
        )
        store.create_work(
            kind="돌발AS",
            business_key="AS-2608-002",
            public_id="AS-2608-002",
            project_no="UJ2609002",
            camp_name="신규캠프",
            status="접수",
            fields={
                "접수ID": "AS-2608-002",
                "프로젝트NO": "UJ2609002",
                "캠프명": "신규캠프",
                "진행상태": "접수",
                "신청내용": "신규 내용",
                "작업완료일": "",
            },
            actor="selftest",
            source="selftest",
            evidence="archive worker DB-only row test",
            idempotency_key="selftest-create",
        )
        store.create_work(
            kind="돌발AS",
            business_key="AS-2608-003",
            public_id="AS-2608-003",
            project_no="UJ2609003",
            camp_name="확장캠프",
            status="접수",
            fields={
                "접수ID": "AS-2608-003",
                "프로젝트NO": "UJ2609003",
                "캠프명": "확장캠프",
                "진행상태": "접수",
                "신청내용": "행 확장 내용",
                "작업완료일": "",
            },
            actor="selftest",
            source="selftest",
            evidence="archive worker appended-row test",
            idempotency_key="selftest-create-append",
        )
        worker = ArchiveWorker(store, base / "spool")
        first = worker.run(template)
        assert first["ok"] and first["state"] == "verified"
        adapter_proof = json.loads(
            (Path(first["export"]["artifact_dir"]) / "adapter-result.json").read_text(
                encoding="utf-8"
            )
        )
        assert adapter_proof["source_copy_proof"]["sha256"] == template_hash
        assert sum(
            "retained in DB audit only" in warning
            for warning in adapter_proof.get("warnings") or []
        ) == len(DB_ONLY_ARCHIVE_FIELDS)
        archive_path = Path(first["last_good"]["archive_path"])
        first_hash = sha256_file(archive_path)
        with zipfile.ZipFile(archive_path, "r") as archive:
            shared = _shared_strings(archive)
            xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            state = SheetState("02_돌발AS접수", "xl/worksheets/sheet1.xml", xml, shared)
            assert state.value(5, state.headers["신청내용"]) == "DB에서 변경"
            assert state.value(6, state.headers["프로젝트NO"]) == "UJ2609002"
            assert state.value(6, state.headers["캠프명"]) == "신규캠프"
            assert _formula(state.cell_fragment(6, state.headers["진행상태"]))
            assert state.value(7, state.headers["프로젝트NO"]) == "UJ2609003"
            assert state.value(7, state.headers["캠프명"]) == "확장캠프"
            appended_formula = state.cell_fragment(7, state.headers["진행상태"])
            assert _formula(appended_formula) and "B7" in appended_formula
        second = worker.run(template)
        assert second["ok"] and second["state"] == "verified"
        assert second["export"]["export_id"] == first["export"]["export_id"]
        assert sha256_file(Path(second["last_good"]["archive_path"])) == first_hash
        assert sha256_file(template) == template_hash
        current = worker.status()
        assert current["archive_current"] is True
        try:
            _require_local_write(Path(r"Z:\forbidden\archive.xlsx"), "selftest")
        except ArchiveSafetyError:
            pass
        else:
            raise AssertionError("Z: write target was not rejected")
        lock_spool = base / "lock-selftest"
        with _worker_lock(lock_spool):
            assert (lock_spool / "archive-worker.lock").is_file()
            try:
                with _worker_lock(lock_spool):
                    raise AssertionError("nested live lock was acquired")
            except ArchiveWorkerBusy:
                pass
        assert not (lock_spool / "archive-worker.lock").exists()

        conflict_store = AppStore(base / "conflict.db").initialize()
        for index, key in enumerate(("AS-CONFLICT-A", "AS-CONFLICT-B"), 1):
            imported_conflict = conflict_store.shadow_import(
                import_id=f"conflict-import-{index}",
                sheet="02_돌발AS접수",
                business_key=key,
                business_key_col="접수ID",
                row_number=5,
                kind="돌발AS",
                public_id=key,
                project_no=f"UJ-CONFLICT-{index}",
                camp_name=f"충돌캠프{index}",
                status="접수",
                fields={
                    "접수ID": key,
                    "프로젝트NO": f"UJ-CONFLICT-{index}",
                    "캠프명": f"충돌캠프{index}",
                    "진행상태": "접수",
                    "신청내용": f"동일 원본 행 충돌 {index}",
                    "작업완료일": "",
                },
                source_file=str(template),
                source_sha256=template_hash,
                apply_if_missing=True,
                idempotency_key=f"conflict-shadow-{index}",
            )
            assert imported_conflict["status"] == "created"
        conflict_worker = ArchiveWorker(conflict_store, base / "conflict-spool")
        sidecar_result = conflict_worker.run(template)
        assert sidecar_result["ok"] and sidecar_result["state"] == "verified"
        assert sidecar_result["last_good"] is not None
        partial_proof = json.loads(
            (
                Path(sidecar_result["export"]["artifact_dir"]) / "adapter-result.json"
            ).read_text(encoding="utf-8")
        )
        assert partial_proof["coverage"]["planned_records"] == 2
        assert partial_proof["coverage"]["covered_records"] == 2
        assert partial_proof["coverage"]["conflict_records"] == 0
        assert partial_proof["coverage"]["sidecar_records"] == 1
        assert partial_proof["coverage"]["complete"] is True
        assert not partial_proof["conflicts"]
        assert partial_proof["sidecar"]["records"] == 1
        assert partial_proof["sidecar"]["entries"][0]["reason"] == "row-claimed-by-other"
        assert conflict_store.export_run(sidecar_result["export"]["export_id"])[
            "status"
        ] == "verified"
        return {
            "ok": True,
            "existing_row_updated": True,
            "db_only_row_inserted": True,
            "formula_preserved": True,
            "appended_row_formula_translated": True,
            "template_unchanged": True,
            "idempotent_export": True,
            "last_good_verified": True,
            "network_write_rejected": True,
            "live_lock_respected": True,
            "bounded_source_stage": True,
            "source_proof_manifested": True,
            "db_only_audit_fields_retained": True,
            "conflict_sidecar_verified": True,
        }


def _json_print(value: Mapping[str, Any]) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True, default=str))


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--run", action="store_true", help="render and verify the current DB snapshot")
    mode.add_argument("--status", action="store_true", help="show worker/store/last-good state")
    mode.add_argument("--self-test", action="store_true", help="run an isolated synthetic verification")
    mode.add_argument("--_copy-source", help=argparse.SUPPRESS)
    parser.add_argument("--_copy-dest", help=argparse.SUPPRESS)
    parser.add_argument("--template", help="read-only XLSX template; UNC/Z input is allowed")
    parser.add_argument("--spool", help="local spool directory (Z:/UNC is rejected)")
    parser.add_argument("--db", help="canonical app_store SQLite path")
    parser.add_argument("--wait-seconds", type=float, default=0.0, help="local worker-lock wait")
    args = parser.parse_args(argv)
    try:
        if args._copy_source:
            if not args._copy_dest:
                raise ValueError("--_copy-dest is required for the internal copy worker")
            _json_print(_copy_source_to_local(args._copy_source, args._copy_dest))
            return 0
        if args.self_test:
            _json_print(self_test())
            return 0
        store = default_store(args.db)
        worker = ArchiveWorker(store, args.spool)
        if args.run:
            template = _resolve_template(store, args.template, worker.spool_dir)
            _json_print(worker.run(template, wait_seconds=args.wait_seconds))
            return 0
        _json_print(worker.status())
        return 0
    except Exception as exc:
        status = getattr(exc, "archive_worker_status", None)
        _json_print(
            status
            or {
                "ok": False,
                "state": "failed",
                "error_type": type(exc).__name__,
                "error": str(exc),
                "external_write_performed": False,
            }
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
