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
        "pid": os.getpid(),
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
                age = max(0.0, time.time() - lock.stat().st_mtime)
            except OSError:
                continue
            owner_pid = owner.get("pid")
            live = pid_alive.alive(owner_pid)
            # Never reclaim a fresh, half-written or otherwise unidentifiable
            # claim.  Only a definitely dead PID, or an unidentifiable claim
            # older than the stale horizon, is safe to move aside.
            occupied = live is True or (live is None and age < LOCK_STALE_SECONDS)
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

    def __post_init__(self) -> None:
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

    def row_fragment(self, row_number: int) -> str:
        for number, _start, _end, fragment in _row_fragments(self.xml):
            if number == int(row_number):
                return fragment
        return ""

    def cell_fragment(self, row_number: int, col: str) -> str:
        return _cell_fragment(self.row_fragment(row_number), f"{col}{row_number}")

    def value(self, row_number: int, col: str) -> Any:
        return _cell_value(self.cell_fragment(row_number, col), self.shared_strings)

    def row_numbers(self) -> List[int]:
        return [number for number, _a, _b, _c in _row_fragments(self.xml)]

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
        self.xml, row = _append_blank_row(self.xml)
        self.indexes.clear()
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
    probes: List[Tuple[str, Any]] = [(key_col, record.get("business_key"))]
    spec = SHEET_SPECS.get(state.name, {})
    fields = dict(record.get("fields") or {})
    for core, raw_name in (
        ("public_id", "public_id"),
        ("project_no", "project_no"),
        ("camp_name", "camp_name"),
    ):
        header = str(spec.get(core) or "")
        value = fields.get(raw_name)
        if header and value not in (None, ""):
            probes.append((header, value))
    found: set[int] = set()
    strong: List[set[int]] = []
    for header, value in probes:
        col = state.headers.get(header)
        if not col or value in (None, ""):
            continue
        rows = set(state.find(col, value))
        if rows:
            strong.append(rows)
            found.update(rows)
    if strong:
        intersection = set.intersection(*strong)
        if intersection:
            return sorted(intersection)
    return sorted(found)


def _locate_records(
    records: Sequence[Mapping[str, Any]], states: Mapping[str, SheetState], template_sha: str
) -> List[LocatedRecord]:
    located: List[LocatedRecord] = []
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
        if len(candidates) > 1:
            raise ArchiveRenderError(
                f"{sheet_name}:{record.get('business_key')} matches multiple rows {candidates[:20]}"
            )
        target = record.get("target") or {}
        row: Optional[int] = candidates[0] if candidates else None
        inserted = False
        source_row = int(target.get("source_row") or 0)
        source_sha = str((record.get("source") or {}).get("source_sha256") or "")
        if row is None and source_row and source_sha and source_sha == template_sha:
            if state.row_fragment(source_row):
                row = source_row
        if row is None:
            spec = SHEET_SPECS.get(sheet_name, {})
            occupancy = [
                key_col,
                state.headers.get(str(spec.get("project_no") or ""), ""),
                state.headers.get(str(spec.get("camp_name") or ""), ""),
            ]
            row = state.allocate(owner, [c for c in occupancy if c])
            inserted = True
        existing_owner = state.reserved.get(row)
        if existing_owner and existing_owner != owner:
            raise ArchiveRenderError(
                f"{sheet_name}!row={row} is claimed by {existing_owner} and {owner}"
            )
        state.reserved[row] = owner
        state.note_index(key_col, record.get("business_key"), row)
        located.append(LocatedRecord(record, state, row, key_header, inserted))
    return located


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
    command = {
        "row": int(row),
        "colL": col,
        "value": normalized,
        "vtype": value_type,
        "only_if_empty": False,
    }
    changed, xml, reason = ledger_writer.apply_to_xml(state.xml, command)
    if not changed:
        raise ArchiveRenderError(
            f"{state.name}!{col}{row} could not be patched: {reason or 'unknown'}"
        )
    state.xml = xml
    state.indexes.clear()
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
        for item in expectations:
            xml = xml_cache.get(item.member)
            if xml is None:
                errors.append(f"worksheet XML missing: {item.member}")
                continue
            row_fragment = ""
            for number, _a, _b, fragment in _row_fragments(xml):
                if number == item.row:
                    row_fragment = fragment
                    break
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
        located = _locate_records(records, states, template_before)
        counts = {
            "records_inserted": 0,
            "records_updated": 0,
            "commands_applied": 0,
            "commands_unchanged": 0,
            "formula_cells_preserved": 0,
        }
        warnings: List[str] = []
        expectations: List[CellExpectation] = []
        for item in located:
            fields, meta, field_warnings = _translated_fields(item.record, item.sheet.name)
            warnings.extend(field_warnings)
            # A DB-only record must carry its canonical key even if the caller did
            # not duplicate that key inside fields.
            fields.setdefault(item.key_col, item.record.get("business_key"))
            meta.setdefault(item.key_col, {})
            for header in sorted(fields):
                if header not in item.sheet.headers:
                    if str(header).startswith("__"):
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
                elif result == "unchanged":
                    counts["commands_unchanged"] += 1
                else:
                    counts["formula_cells_preserved"] += 1
            if item.inserted:
                counts["records_inserted"] += 1
            else:
                counts["records_updated"] += 1
        _write_patched_zip(
            template_copy_path,
            output_path,
            {state.member: state.xml for state in states.values()},
        )
        validation = _validate_output(output_path, expectations, states)
        if sha256_file(template_copy_path) != template_before:
            raise ArchiveVerificationError("adapter mutated the template copy")
        return {
            "format": WORKER_PROOF_FORMAT,
            "status": "success",
            "snapshot_sha256": str(plan["snapshot"]["sha256"]),
            "command_plan_sha256": sha256_json(plan),
            "rows_considered": len(records),
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
                    verification = self.exporter.verify_export(
                        result["artifact_dir"], require_verified=True
                    )
                    if not verification.get("ok"):
                        raise ArchiveVerificationError(
                            f"verified artifact failed recheck: {verification.get('errors')}"
                        )
                    last_good = self.exporter.last_good(verify=True)
                    source_after = sha256_file(template)
                    if source_after != source_before:
                        raise ArchiveVerificationError(
                            "stable template changed during archive run; retry from source"
                        )
                    payload = self._write_state(
                        {
                            "state": "verified",
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
            patch={"fields": {"신청내용": "DB에서 변경"}},
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
