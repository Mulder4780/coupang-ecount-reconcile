"""Local, verifiable archive export for the SQLite application store.

The exporter never opens and saves an existing workbook.  A template is copied
byte-for-byte into a local spool, and a deterministic command-plan is produced
for a writer adapter.  An adapter must write a *new* local output and return a
machine-checkable result before that output can become ``last-good``.

No network share or Z: destination is accepted for writes.  External publishing
is intentionally a separate, human-approved operation.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
import threading
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List, Mapping, Optional, Sequence, Tuple

from app_store import (
    AppStore,
    NotFoundError,
    StoreError,
    ValidationError,
    VersionConflict,
    canonical_json,
    default_store,
    sha256_json,
)


PLAN_FORMAT = "csos-archive-command-plan/v1"
MANIFEST_FORMAT = "csos-archive-manifest/v1"
VALIDATION_FORMAT = "csos-archive-validation/v1"
COVERAGE_CONTRACT_VERSION = 2
CHUNK_SIZE = 1024 * 1024
_READ_ONLY_MASTER_RESOLVE_LOCK = threading.Lock()


class ArchiveExportError(RuntimeError):
    """Base error for local archive generation."""


class ArchiveSafetyError(ArchiveExportError):
    """A caller attempted to write outside the permitted local spool."""


class ArchiveVerificationError(ArchiveExportError):
    """An artifact or adapter proof failed verification."""


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def sha256_file(path: os.PathLike[str] | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        while True:
            chunk = handle.read(CHUNK_SIZE)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _file_entry(path: Path) -> Dict[str, Any]:
    return {"sha256": sha256_file(path), "size": path.stat().st_size}


def _manifest_digest(manifest: Mapping[str, Any]) -> str:
    unsigned = dict(manifest)
    unsigned.pop("manifest_sha256", None)
    return sha256_json(unsigned)


def _same_drive_local(path: Path) -> bool:
    """Return True only for a local path on the module's Windows drive."""

    raw = str(path)
    if raw.startswith("\\\\") or raw.startswith("//"):
        return False
    module_drive = Path(__file__).resolve().drive.upper()
    drive = path.drive.upper()
    if drive == "Z:":
        return False
    if module_drive and drive and drive != module_drive:
        return False
    return True


def _require_local_write_path(path: os.PathLike[str] | str, label: str) -> Path:
    candidate = Path(path).resolve()
    if not _same_drive_local(candidate):
        raise ArchiveSafetyError(
            f"{label} must be a local path on {Path(__file__).resolve().drive or 'the local drive'}; "
            f"network/Z: writes are disabled: {candidate}"
        )
    return candidate


def _ensure_within(path: Path, root: Path, label: str) -> None:
    try:
        path.resolve().relative_to(root.resolve())
    except ValueError as exc:
        raise ArchiveSafetyError(f"{label} escapes the local spool: {path}") from exc


def _atomic_write_bytes(path: Path, data: bytes) -> Path:
    path = _require_local_write_path(path, "atomic output")
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent)
    )
    temp_path = Path(temp_name)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return path


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> Path:
    return _atomic_write_bytes(path, canonical_json(payload).encode("utf-8"))


def _byte_copy_verified(source: Path, destination: Path) -> Dict[str, Any]:
    """Copy without workbook parsing/saving and verify source/destination hashes."""

    destination = _require_local_write_path(destination, "archive copy")
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{destination.name}.", suffix=".tmp", dir=str(destination.parent)
    )
    temp_path = Path(temp_name)
    source_digest = hashlib.sha256()
    destination_digest = hashlib.sha256()
    size = 0
    try:
        with source.open("rb") as src, os.fdopen(fd, "wb") as dst:
            while True:
                chunk = src.read(CHUNK_SIZE)
                if not chunk:
                    break
                source_digest.update(chunk)
                dst.write(chunk)
                destination_digest.update(chunk)
                size += len(chunk)
            dst.flush()
            os.fsync(dst.fileno())
        if source_digest.digest() != destination_digest.digest():
            raise ArchiveVerificationError("byte-copy hash mismatch")
        os.replace(temp_path, destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    final_hash = sha256_file(destination)
    if final_hash != source_digest.hexdigest() or destination.stat().st_size != size:
        raise ArchiveVerificationError("copied file changed after atomic replace")
    return {"sha256": final_hash, "size": size}


def _value_type(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    return "json"


def _validate_xlsx_container(path: Path) -> Dict[str, Any]:
    """Validate the ZIP container without using openpyxl or resaving it."""

    if not path.is_file() or path.stat().st_size == 0:
        raise ArchiveVerificationError(f"workbook output is missing or empty: {path}")
    if not zipfile.is_zipfile(path):
        raise ArchiveVerificationError(f"workbook output is not an OOXML ZIP: {path}")
    required = {"[Content_Types].xml", "xl/workbook.xml"}
    with zipfile.ZipFile(path, "r") as archive:
        names = set(archive.namelist())
        missing = sorted(required - names)
        if missing:
            raise ArchiveVerificationError(
                f"workbook output lacks required OOXML entries: {', '.join(missing)}"
            )
        bad = archive.testzip()
        if bad:
            raise ArchiveVerificationError(f"workbook ZIP CRC failed: {bad}")
        return {"zip_entries": len(names), "required_entries": sorted(required)}


WriterAdapter = Callable[[Path, Path, Path], Mapping[str, Any]]


class ArchiveExporter:
    """Create immutable local export plans and verified archive artifacts."""

    def __init__(
        self,
        store: AppStore,
        spool_dir: Optional[os.PathLike[str] | str] = None,
    ) -> None:
        default = Path(__file__).resolve().parent / "tmp" / "archive_spool"
        self.store = store.initialize()
        self.spool_dir = _require_local_write_path(
            spool_dir or os.environ.get("COUPANG_ARCHIVE_SPOOL") or default,
            "archive spool",
        )
        self.exports_dir = self.spool_dir / "exports"
        _ensure_within(self.exports_dir, self.spool_dir, "exports directory")
        self.exports_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _latest_shadow(work: Mapping[str, Any]) -> Optional[Mapping[str, Any]]:
        locations = list(work.get("shadow_locations") or [])
        return locations[0] if locations else None

    def build_command_plan(
        self,
        snapshot: Mapping[str, Any],
        snapshot_sha256: str,
        *,
        export_id: str,
        template_name: str,
        template_sha256: str,
    ) -> Dict[str, Any]:
        """Translate canonical records to a deterministic writer-adapter contract."""

        records: List[Dict[str, Any]] = []
        legacy_queue: List[Dict[str, Any]] = []
        for work in snapshot.get("work_items", []):
            shadow = self._latest_shadow(work)
            target = None
            if shadow:
                target = {
                    "sheet": shadow["sheet"],
                    "business_key_col": shadow.get("business_key_col") or "프로젝트NO",
                    "business_key": work["business_key"],
                    "source_row": shadow["row_number"],
                    "source_import_id": shadow["import_id"],
                    "source_row_sha256": shadow["row_sha256"],
                }
            record_fields = {
                "public_id": work.get("public_id"),
                "project_no": work.get("project_no"),
                "camp_name": work.get("camp_name"),
                "status": work.get("status"),
                **dict(work.get("fields") or {}),
            }
            records.append(
                {
                    # Soft-deleted records are still part of the immutable
                    # canonical snapshot.  The Excel adapter keeps their full
                    # record in the verified audit sidecar instead of silently
                    # dropping the deletion event or erasing a recoverable row.
                    "op": (
                        "archive_tombstone"
                        if work.get("deleted_at")
                        else "upsert_record"
                    ),
                    "work_id": work["id"],
                    "record_version": work["record_version"],
                    "kind": work["kind"],
                    "business_key": work["business_key"],
                    "deleted_at": work.get("deleted_at"),
                    "deleted_by": work.get("deleted_by"),
                    "target": target,
                    "fields": record_fields,
                    "field_meta": work.get("field_meta") or {},
                    "source": {
                        "source": work.get("source") or "",
                        "evidence": work.get("evidence") or "",
                        "source_ref": work.get("source_ref") or "",
                        "source_sha256": work.get("source_sha256") or "",
                    },
                }
            )
            # The legacy queue can safely patch located rows.  Field keys are
            # column headers; creation/allocation remains the adapter's job via
            # the high-level records list above.
            if target and not work.get("deleted_at"):
                for field_key, value in sorted((work.get("fields") or {}).items()):
                    if field_key.startswith("__"):
                        continue
                    meta = (work.get("field_meta") or {}).get(field_key, {})
                    legacy_queue.append(
                        {
                            "sheet": target["sheet"],
                            "key_col": target["business_key_col"],
                            "key": target["business_key"],
                            "col": field_key,
                            "value": value,
                            "vtype": _value_type(value),
                            "evidence": meta.get("evidence")
                            or work.get("evidence")
                            or f"app_store snapshot {snapshot_sha256[:12]}",
                            "only_if_empty": False,
                            "expected_record_version": work["record_version"],
                            "idempotency_key": (
                                f"{export_id}:{work['id']}:{field_key}:"
                                f"v{work['record_version']}"
                            ),
                        }
                    )
        plan = {
            "format": PLAN_FORMAT,
            "export_id": export_id,
            "snapshot": {
                "change_seq": int(snapshot.get("change_seq", 0)),
                "sha256": snapshot_sha256,
                "schema_version": snapshot.get("schema_version"),
            },
            "template": {
                "local_member": "template-copy.xlsx" if template_sha256 else None,
                "source_name": template_name,
                "sha256": template_sha256,
                "copy_semantics": "byte-for-byte; source workbook must never be saved",
            },
            "adapter_contract": {
                "name": "ledger-writer-compatible/v1",
                "record_coverage_version": COVERAGE_CONTRACT_VERSION,
                "execution": "explicit-local-callable-only",
                "call_signature": "adapter(plan_path, template_copy_path, output_path) -> result",
                "high_level_member": "records",
                "legacy_queue_member": "legacy_writer_queue",
                "requirements": [
                    "write a new output_path; never mutate template_copy_path",
                    "enforce per-command idempotency_key",
                    "report snapshot_sha256 and command_plan_sha256",
                    "report rows_considered and an empty errors list before verification",
                    "report applied/unchanged/archived_sidecar work_id+record_version coverage for every record",
                    "do not publish to Z: or any network share",
                ],
            },
            "settings": list(snapshot.get("settings") or []),
            "records": records,
            "legacy_writer_queue": legacy_queue,
        }
        return plan

    @staticmethod
    def _validation_plan(
        snapshot: Mapping[str, Any],
        snapshot_sha256: str,
        plan: Mapping[str, Any],
        plan_sha256: str,
        template_sha256: str,
    ) -> Dict[str, Any]:
        work_items = list(snapshot.get("work_items") or [])
        field_count = sum(len(w.get("fields") or {}) for w in work_items)
        located = sum(1 for w in work_items if w.get("shadow_locations"))
        return {
            "format": VALIDATION_FORMAT,
            "snapshot_sha256": snapshot_sha256,
            "command_plan_sha256": plan_sha256,
            "template_sha256": template_sha256,
            "expected": {
                "change_seq": int(snapshot.get("change_seq", 0)),
                "work_item_count": len(work_items),
                "work_field_count": field_count,
                "setting_count": len(snapshot.get("settings") or []),
                "located_record_count": located,
                "command_record_count": len(plan.get("records") or []),
                "legacy_writer_queue_count": len(
                    plan.get("legacy_writer_queue") or []
                ),
            },
            "gates": [
                "snapshot JSON hash matches manifest",
                "SQLite backup passes PRAGMA quick_check",
                "template-copy hash equals source hash",
                "adapter leaves template-copy byte-identical",
                "adapter proof identifies this snapshot and command plan",
                "adapter reports no errors",
                "adapter partitions every planned work_id+record_version into covered or conflicted",
                "only applied/unchanged/verified-sidecar record coverage can be acknowledged",
                "output is a valid OOXML ZIP with workbook.xml",
                "only complete record coverage can atomically advance last-good.json",
            ],
        }

    def _artifact_path(self, export_id: str) -> Path:
        if not export_id or any(c not in "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_" for c in export_id):
            raise ValidationError("invalid export_id")
        path = self.exports_dir / export_id
        _ensure_within(path, self.exports_dir, "export artifact")
        return path

    def prepare(
        self,
        *,
        template_path: Optional[os.PathLike[str] | str] = None,
    ) -> Dict[str, Any]:
        """Create/reuse a local immutable snapshot and dry-run command plan."""

        template: Optional[Path] = None
        template_sha = ""
        template_name = ""
        if template_path is not None:
            template = Path(template_path).resolve()
            if not template.is_file():
                raise ArchiveExportError(f"template does not exist: {template}")
            template_sha = sha256_file(template)
            template_name = template.name
        # A verified archive must cover deletion tombstones as well as active
        # rows.  Otherwise one soft-delete outbox message can never be
        # acknowledged and every five-minute pipeline round ends red forever.
        snapshot, snapshot_sha = self.store.snapshot(include_deleted=True)
        seq = int(snapshot.get("change_seq", 0))
        template_token = template_sha[:12] if template_sha else "no-template"
        # The coverage suffix prevents a pre-coverage verified artifact with the
        # same snapshot/template pair from being silently reused.  Old last-good
        # pointers remain readable, but every new export must satisfy this
        # stronger semantic contract.
        export_id = (
            f"exp-{seq:012d}-{snapshot_sha[:16]}-{template_token}"
            f"-cov{COVERAGE_CONTRACT_VERSION}"
        )
        final_dir = self._artifact_path(export_id)
        if final_dir.exists():
            verification = self.verify_export(final_dir)
            if not verification["ok"]:
                raise ArchiveVerificationError(
                    f"existing export is corrupt: {verification['errors']}"
                )
            existing_manifest = json.loads(
                (final_dir / "manifest.json").read_text(encoding="utf-8")
            )
            self._sync_export_run(final_dir, existing_manifest)
            if existing_manifest.get("status") == "verified":
                self._advance_last_good(final_dir, existing_manifest)
            return self._artifact_summary(final_dir)

        stage = self.exports_dir / f".{export_id}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
        _ensure_within(stage, self.exports_dir, "export stage")
        stage.mkdir(parents=False, exist_ok=False)
        try:
            snapshot_path = stage / "snapshot.json"
            _atomic_write_bytes(snapshot_path, canonical_json(snapshot).encode("utf-8"))
            if sha256_file(snapshot_path) != snapshot_sha:
                raise ArchiveVerificationError("snapshot JSON digest changed on disk")

            db_snapshot = self.store.backup_to(stage / "db-snapshot.sqlite3")
            template_entry: Optional[Dict[str, Any]] = None
            if template is not None:
                template_entry = _byte_copy_verified(
                    template, stage / "template-copy.xlsx"
                )
                if template_entry["sha256"] != template_sha:
                    raise ArchiveVerificationError("template byte-copy changed content")

            plan = self.build_command_plan(
                snapshot,
                snapshot_sha,
                export_id=export_id,
                template_name=template_name,
                template_sha256=template_sha,
            )
            plan_sha = sha256_json(plan)
            plan_path = stage / "command-plan.json"
            _atomic_write_bytes(plan_path, canonical_json(plan).encode("utf-8"))
            if sha256_file(plan_path) != plan_sha:
                raise ArchiveVerificationError("command plan digest changed on disk")

            validation = self._validation_plan(
                snapshot, snapshot_sha, plan, plan_sha, template_sha
            )
            validation_path = stage / "validation-manifest.json"
            _atomic_write_json(validation_path, validation)

            files: Dict[str, Dict[str, Any]] = {
                "snapshot.json": _file_entry(snapshot_path),
                "db-snapshot.sqlite3": _file_entry(db_snapshot),
                "command-plan.json": _file_entry(plan_path),
                "validation-manifest.json": _file_entry(validation_path),
            }
            if template_entry is not None:
                files["template-copy.xlsx"] = template_entry
            manifest: Dict[str, Any] = {
                "format": MANIFEST_FORMAT,
                "export_id": export_id,
                "status": "planned",
                "mode": "local-only",
                "created_at": _utcnow(),
                "snapshot": {
                    "change_seq": seq,
                    "sha256": snapshot_sha,
                    "schema_version": snapshot.get("schema_version"),
                },
                "template": {
                    "source_name": template_name,
                    "source_sha256": template_sha,
                    "copy_member": "template-copy.xlsx" if template else None,
                    "copy_method": "binary-stream-copy; no workbook parser",
                },
                "command_plan_sha256": plan_sha,
                "coverage_contract": {
                    "version": COVERAGE_CONTRACT_VERSION,
                    "record_identity": "work_id+record_version",
                    "covered_outcomes": ["applied", "unchanged", "archived_sidecar"],
                    "partial_last_good_allowed": False,
                },
                "files": files,
                "output": None,
                "publish": {
                    "external_write_performed": False,
                    "network_and_z_drive_write_allowed": False,
                },
            }
            manifest["manifest_sha256"] = _manifest_digest(manifest)
            _atomic_write_json(stage / "manifest.json", manifest)

            staged_check = self.verify_export(stage)
            if not staged_check["ok"]:
                raise ArchiveVerificationError(
                    f"staged export verification failed: {staged_check['errors']}"
                )
            try:
                os.replace(stage, final_dir)
            except OSError:
                if not final_dir.exists():
                    raise
                existing = self.verify_export(final_dir)
                if not existing["ok"]:
                    raise ArchiveVerificationError(
                        "concurrent export exists but did not verify"
                    )
                shutil.rmtree(stage)

            self._sync_export_run(final_dir, manifest)
            return self._artifact_summary(final_dir)
        except Exception:
            if stage.exists():
                _ensure_within(stage, self.exports_dir, "failed export stage")
                shutil.rmtree(stage, ignore_errors=True)
            raise

    def _artifact_summary(self, artifact_dir: Path) -> Dict[str, Any]:
        manifest = json.loads((artifact_dir / "manifest.json").read_text(encoding="utf-8"))
        return {
            "export_id": manifest["export_id"],
            "status": manifest["status"],
            "artifact_dir": str(artifact_dir),
            "snapshot_sha256": manifest["snapshot"]["sha256"],
            "command_plan_sha256": manifest["command_plan_sha256"],
            "manifest_sha256": manifest["manifest_sha256"],
            "coverage": dict(manifest.get("coverage") or {}),
            "last_good_eligible": manifest.get("status") == "verified",
            "command_plan": str(artifact_dir / "command-plan.json"),
            "template_copy": (
                str(artifact_dir / "template-copy.xlsx")
                if (artifact_dir / "template-copy.xlsx").exists()
                else None
            ),
            "archive": (
                str(artifact_dir / "archive.xlsx")
                if (artifact_dir / "archive.xlsx").exists()
                else None
            ),
        }

    def _sync_export_run(self, artifact_dir: Path, manifest: Mapping[str, Any]) -> None:
        """Repair DB bookkeeping after a crash between file and DB commits."""

        run = self.store.start_export(
            export_id=str(manifest["export_id"]),
            snapshot_seq=int(manifest["snapshot"]["change_seq"]),
            snapshot_sha256=str(manifest["snapshot"]["sha256"]),
            template_sha256=str(manifest.get("template", {}).get("source_sha256") or ""),
            plan_sha256=str(manifest.get("command_plan_sha256") or ""),
            mode="local-only",
            local_path=str(artifact_dir),
        )
        manifest_status = str(manifest.get("status") or "planned")
        status = manifest_status if manifest_status in {"planned", "partial", "verified"} else "planned"
        if run.get("status") == "verified" and status == "planned":
            return
        self.store.finish_export(
            str(manifest["export_id"]),
            status=status,
            manifest_sha256=str(manifest.get("manifest_sha256") or ""),
            local_path=str(artifact_dir),
        )

    def _advance_last_good(
        self, artifact_dir: Path, manifest: Mapping[str, Any]
    ) -> Dict[str, Any]:
        if manifest.get("status") != "verified":
            raise ArchiveVerificationError("only a verified artifact can become last-good")
        archive_path = artifact_dir / "archive.xlsx"
        archive_sha = sha256_file(archive_path)
        pointer = {
            "format": "csos-archive-last-good/v1",
            "export_id": manifest["export_id"],
            "snapshot_seq": manifest["snapshot"]["change_seq"],
            "snapshot_sha256": manifest["snapshot"]["sha256"],
            "archive_path": str(archive_path),
            "archive_sha256": archive_sha,
            "manifest_path": str(artifact_dir / "manifest.json"),
            "manifest_sha256": manifest["manifest_sha256"],
            "verified_at": manifest.get("verified_at") or _utcnow(),
            "external_write_performed": False,
        }
        _atomic_write_json(self.spool_dir / "last-good.json", pointer)
        return pointer

    @staticmethod
    def _validate_record_coverage(
        proof: Dict[str, Any], records: Sequence[Mapping[str, Any]]
    ) -> Dict[str, Any]:
        """Validate exact work revision coverage reported by an adapter.

        A snapshot hash proves which DB state was planned, not that each work
        revision reached the archive.  Coverage is therefore keyed by the
        immutable ``work_id`` plus the snapshot's ``record_version``.  Unsafe
        main-sheet matches may be covered only by the semantically verified
        AppDB sidecar; anything else must remain named in ``conflicts``.
        """

        expected: Dict[str, Dict[str, Any]] = {}
        for record in records:
            work_id = str(record.get("work_id") or "").strip()
            try:
                record_version = int(record.get("record_version") or 0)
            except (TypeError, ValueError) as exc:
                raise ArchiveVerificationError(
                    f"command record has invalid revision: {work_id or '<missing>'}"
                ) from exc
            if not work_id or record_version < 1:
                raise ArchiveVerificationError(
                    "every command record requires work_id and positive record_version"
                )
            if work_id in expected:
                raise ArchiveVerificationError(f"duplicate work_id in command plan: {work_id}")
            expected[work_id] = {
                "record_version": record_version,
                "business_key": str(record.get("business_key") or ""),
                "record_sha256": hashlib.sha256(
                    canonical_json(record).encode("utf-8")
                ).hexdigest(),
            }

        raw_coverage = proof.get("record_coverage")
        if not isinstance(raw_coverage, list):
            raise ArchiveVerificationError("adapter record_coverage must be a list")
        covered: Dict[str, Dict[str, Any]] = {}
        normalized: List[Dict[str, Any]] = []
        for raw in raw_coverage:
            if not isinstance(raw, Mapping):
                raise ArchiveVerificationError("adapter record_coverage entry must be an object")
            entry = dict(raw)
            work_id = str(entry.get("work_id") or "").strip()
            outcome = str(entry.get("outcome") or "").strip()
            try:
                record_version = int(entry.get("record_version") or 0)
            except (TypeError, ValueError) as exc:
                raise ArchiveVerificationError(
                    f"coverage has invalid revision for {work_id or '<missing>'}"
                ) from exc
            if outcome not in {"applied", "unchanged", "archived_sidecar"}:
                raise ArchiveVerificationError(
                    f"coverage outcome is not ackable for {work_id or '<missing>'}: {outcome!r}"
                )
            if work_id not in expected:
                raise ArchiveVerificationError(f"coverage references unknown work_id: {work_id}")
            if work_id in covered:
                raise ArchiveVerificationError(f"duplicate coverage for work_id: {work_id}")
            if record_version != int(expected[work_id]["record_version"]):
                raise ArchiveVerificationError(
                    f"coverage revision mismatch for {work_id}: "
                    f"expected {expected[work_id]['record_version']}, got {record_version}"
                )
            business_key = str(entry.get("business_key") or "")
            if business_key and business_key != expected[work_id]["business_key"]:
                raise ArchiveVerificationError(
                    f"coverage business_key mismatch for {work_id}"
                )
            if outcome == "archived_sidecar":
                if str(entry.get("sheet") or "") != "99_AppDB_미매칭보관":
                    raise ArchiveVerificationError(
                        f"sidecar coverage has wrong worksheet for {work_id}"
                    )
                if str(entry.get("record_sha256") or "") != expected[work_id]["record_sha256"]:
                    raise ArchiveVerificationError(
                        f"sidecar canonical record hash mismatch for {work_id}"
                    )
            entry.update(
                {
                    "work_id": work_id,
                    "record_version": record_version,
                    "outcome": outcome,
                }
            )
            covered[work_id] = entry
            normalized.append(entry)

        raw_conflicts = proof.get("conflicts") or []
        if not isinstance(raw_conflicts, list):
            raise ArchiveVerificationError("adapter conflicts must be a list")
        conflict_ids: set[str] = set()
        for raw in raw_conflicts:
            if not isinstance(raw, Mapping):
                raise ArchiveVerificationError("adapter conflict entry must be an object")
            work_id = str(raw.get("work_id") or "").strip()
            if not work_id or work_id not in expected:
                raise ArchiveVerificationError(
                    f"adapter conflict references unknown work_id: {work_id or '<missing>'}"
                )
            if work_id in conflict_ids:
                raise ArchiveVerificationError(
                    f"duplicate conflict for work_id: {work_id}"
                )
            if work_id in covered:
                raise ArchiveVerificationError(
                    f"work_id is both covered and conflicted: {work_id}"
                )
            try:
                conflict_version = int(raw.get("record_version") or 0)
            except (TypeError, ValueError) as exc:
                raise ArchiveVerificationError(
                    f"conflict has invalid revision for {work_id}"
                ) from exc
            if conflict_version != int(expected[work_id]["record_version"]):
                raise ArchiveVerificationError(
                    f"conflict revision mismatch for {work_id}"
                )
            conflict_ids.add(work_id)

        sidecar_covered = {
            work_id: entry
            for work_id, entry in covered.items()
            if entry.get("outcome") == "archived_sidecar"
        }
        raw_sidecar = proof.get("sidecar") or {}
        if sidecar_covered:
            if not isinstance(raw_sidecar, Mapping):
                raise ArchiveVerificationError("sidecar coverage proof must be an object")
            if str(raw_sidecar.get("format") or "") != "csos-appdb-sidecar/v1":
                raise ArchiveVerificationError("sidecar coverage format is invalid")
            if str(raw_sidecar.get("sheet") or "") != "99_AppDB_미매칭보관":
                raise ArchiveVerificationError("sidecar coverage worksheet is invalid")
            raw_entries = raw_sidecar.get("entries")
            if not isinstance(raw_entries, list):
                raise ArchiveVerificationError("sidecar entries must be a list")
            normalized_sidecar = []
            seen_sidecar: set[str] = set()
            for raw in raw_entries:
                if not isinstance(raw, Mapping):
                    raise ArchiveVerificationError("sidecar entry must be an object")
                work_id = str(raw.get("work_id") or "")
                if work_id not in sidecar_covered or work_id in seen_sidecar:
                    raise ArchiveVerificationError(
                        f"sidecar entry has unknown/duplicate work_id: {work_id}"
                    )
                entry = sidecar_covered[work_id]
                normalized_entry = {
                    "work_id": work_id,
                    "record_version": int(raw.get("record_version") or 0),
                    "business_key": str(raw.get("business_key") or ""),
                    "reason": str(raw.get("reason") or ""),
                    "record_sha256": str(raw.get("record_sha256") or ""),
                }
                if normalized_entry != {
                    "work_id": work_id,
                    "record_version": int(entry["record_version"]),
                    "business_key": str(entry.get("business_key") or ""),
                    "reason": str(entry.get("reason") or ""),
                    "record_sha256": str(entry.get("record_sha256") or ""),
                }:
                    raise ArchiveVerificationError(
                        f"sidecar entry differs from coverage for {work_id}"
                    )
                normalized_sidecar.append(normalized_entry)
                seen_sidecar.add(work_id)
            if seen_sidecar != set(sidecar_covered):
                raise ArchiveVerificationError("sidecar entries do not cover every sidecar outcome")
            normalized_sidecar.sort(
                key=lambda item: (item["business_key"], item["work_id"])
            )
            semantic_sha = sha256_json(
                {"format": "csos-appdb-sidecar/v1", "records": normalized_sidecar}
            )
            if semantic_sha != str(raw_sidecar.get("semantic_sha256") or ""):
                raise ArchiveVerificationError("sidecar semantic hash mismatch")
            validation_sidecar = (proof.get("validation") or {}).get("sidecar") or {}
            if (
                str(validation_sidecar.get("sheet") or "") != "99_AppDB_미매칭보관"
                or int(validation_sidecar.get("records") or -1) != len(normalized_sidecar)
                or str(validation_sidecar.get("semantic_sha256") or "") != semantic_sha
            ):
                raise ArchiveVerificationError(
                    "adapter did not semantically verify the sidecar worksheet"
                )
        elif raw_sidecar and int(raw_sidecar.get("records") or 0) != 0:
            raise ArchiveVerificationError("sidecar proof exists without sidecar coverage")

        uncovered_ids = sorted(set(expected) - set(covered))
        if set(uncovered_ids) != conflict_ids:
            unnamed = sorted(set(uncovered_ids) - conflict_ids)
            extra = sorted(conflict_ids - set(uncovered_ids))
            raise ArchiveVerificationError(
                f"coverage/conflict partition mismatch: unnamed={unnamed[:20]}, extra={extra[:20]}"
            )
        complete = not uncovered_ids
        status = str(proof.get("status") or "")
        if complete and status != "success":
            raise ArchiveVerificationError("complete record coverage must report status=success")
        if not complete and status != "partial":
            raise ArchiveVerificationError("incomplete record coverage must report status=partial")
        normalized.sort(key=lambda row: (str(row.get("work_id") or ""), int(row.get("record_version") or 0)))
        proof["record_coverage"] = normalized
        proof["records_covered"] = len(normalized)
        proof["coverage"] = {
            "version": COVERAGE_CONTRACT_VERSION,
            "planned_records": len(expected),
            "covered_records": len(normalized),
            "uncovered_records": len(uncovered_ids),
            "conflict_records": len(conflict_ids),
            "sidecar_records": len(sidecar_covered),
            "complete": complete,
        }
        return proof

    @classmethod
    def _adapter_proof(
        cls,
        result: Mapping[str, Any],
        *,
        snapshot_sha256: str,
        command_plan_sha256: str,
        records: Sequence[Mapping[str, Any]],
    ) -> Dict[str, Any]:
        proof = dict(result)
        required = {
            "status",
            "snapshot_sha256",
            "command_plan_sha256",
            "rows_considered",
            "errors",
        }
        missing = sorted(required - set(proof))
        if missing:
            raise ArchiveVerificationError(
                f"adapter proof is incomplete: {', '.join(missing)}"
            )
        if proof["status"] not in {"success", "partial"}:
            raise ArchiveVerificationError(
                f"adapter did not report a supported status: {proof.get('status')}"
            )
        if proof["snapshot_sha256"] != snapshot_sha256:
            raise ArchiveVerificationError("adapter proof references another snapshot")
        if proof["command_plan_sha256"] != command_plan_sha256:
            raise ArchiveVerificationError("adapter proof references another command plan")
        if not isinstance(proof["rows_considered"], int) or proof["rows_considered"] < 0:
            raise ArchiveVerificationError("adapter rows_considered must be a non-negative int")
        if proof["rows_considered"] != len(records):
            raise ArchiveVerificationError(
                "adapter rows_considered differs from command-plan records"
            )
        if not isinstance(proof["errors"], list) or proof["errors"]:
            raise ArchiveVerificationError("adapter reported unresolved errors")
        return cls._validate_record_coverage(proof, records)

    def finalize_local(
        self,
        export_id: str,
        rendered_path: os.PathLike[str] | str,
        *,
        adapter_result: Mapping[str, Any],
    ) -> Dict[str, Any]:
        """Promote a local adapter output only after all verification gates pass."""

        artifact_dir = self._artifact_path(export_id)
        if not artifact_dir.is_dir():
            raise NotFoundError(f"prepared export not found: {export_id}")
        manifest_path = artifact_dir / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        if manifest["status"] == "verified":
            verification = self.verify_export(artifact_dir, require_verified=True)
            if not verification["ok"]:
                raise ArchiveVerificationError(
                    f"existing verified export is corrupt: {verification['errors']}"
                )
            self._advance_last_good(artifact_dir, manifest)
            self._sync_export_run(artifact_dir, manifest)
            return self._artifact_summary(artifact_dir)

        rendered = _require_local_write_path(rendered_path, "adapter output")
        _validate_xlsx_container(rendered)
        plan_path = artifact_dir / "command-plan.json"
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        proof = self._adapter_proof(
            adapter_result,
            snapshot_sha256=manifest["snapshot"]["sha256"],
            command_plan_sha256=manifest["command_plan_sha256"],
            records=list(plan.get("records") or []),
        )
        if "output_sha256" in proof and proof["output_sha256"] != sha256_file(rendered):
            raise ArchiveVerificationError("adapter output_sha256 does not match output")

        template_path = artifact_dir / "template-copy.xlsx"
        if template_path.exists():
            expected_template = manifest["template"]["source_sha256"]
            if sha256_file(template_path) != expected_template:
                raise ArchiveVerificationError("adapter mutated the byte-copy template")

        archive_path = artifact_dir / "archive.xlsx"
        archive_entry = _byte_copy_verified(rendered, archive_path)
        container = _validate_xlsx_container(archive_path)
        proof.setdefault("output_sha256", archive_entry["sha256"])
        coverage_complete = bool((proof.get("coverage") or {}).get("complete"))
        proof["finished_at"] = _utcnow()
        proof["verified_at"] = _utcnow() if coverage_complete else None
        proof_path = artifact_dir / "adapter-result.json"
        _atomic_write_json(proof_path, proof)

        final_status = "verified" if coverage_complete else "partial"
        validation_result = {
            "format": VALIDATION_FORMAT,
            "export_id": export_id,
            "status": final_status,
            "finished_at": _utcnow(),
            "verified_at": _utcnow() if coverage_complete else None,
            "checks": {
                "snapshot_sha256": True,
                "command_plan_sha256": True,
                "template_unchanged": True,
                "adapter_errors_empty": True,
                "record_coverage_complete": coverage_complete,
                "record_coverage": dict(proof.get("coverage") or {}),
                "output_sha256": archive_entry["sha256"],
                "ooxml_container": container,
                "external_write_performed": False,
            },
        }
        result_path = artifact_dir / "validation-result.json"
        _atomic_write_json(result_path, validation_result)

        manifest["status"] = final_status
        manifest["verified_at"] = _utcnow() if coverage_complete else None
        manifest["finished_at"] = _utcnow()
        manifest["coverage"] = dict(proof.get("coverage") or {})
        manifest["files"]["archive.xlsx"] = archive_entry
        manifest["files"]["adapter-result.json"] = _file_entry(proof_path)
        manifest["files"]["validation-result.json"] = _file_entry(result_path)
        manifest["output"] = {
            "member": "archive.xlsx",
            "sha256": archive_entry["sha256"],
            "size": archive_entry["size"],
        }
        manifest["manifest_sha256"] = _manifest_digest(manifest)
        _atomic_write_json(manifest_path, manifest)

        verification = self.verify_export(
            artifact_dir, require_verified=coverage_complete
        )
        if not verification["ok"]:
            self.store.finish_export(
                export_id,
                status="failed",
                manifest_sha256=manifest["manifest_sha256"],
                local_path=str(artifact_dir),
                error="; ".join(verification["errors"]),
            )
            raise ArchiveVerificationError(
                f"final export verification failed: {verification['errors']}"
            )

        if coverage_complete:
            self._advance_last_good(artifact_dir, manifest)
        self._sync_export_run(artifact_dir, manifest)
        return self._artifact_summary(artifact_dir)

    def run_local(
        self,
        *,
        template_path: os.PathLike[str] | str,
        adapter: Optional[WriterAdapter] = None,
        dry_run: bool = True,
    ) -> Dict[str, Any]:
        """Prepare a plan, and optionally invoke an explicit in-process adapter.

        ``dry_run=True`` (the default) never calls the adapter.  The output path
        passed to an adapter is always a temporary file inside the local spool.
        """

        prepared = self.prepare(template_path=template_path)
        if dry_run:
            prepared["dry_run"] = True
            return prepared
        if prepared.get("status") == "verified":
            # A prior adapter run already passed every gate.  Re-running it
            # would create duplicate workbook work without changing the snapshot.
            return prepared
        if adapter is None or not callable(adapter):
            raise ValidationError("a callable adapter is required when dry_run=False")
        artifact_dir = Path(prepared["artifact_dir"])
        template_copy = artifact_dir / "template-copy.xlsx"
        plan_path = artifact_dir / "command-plan.json"
        if not template_copy.is_file():
            raise ArchiveExportError("prepared export has no local template copy")
        template_before = sha256_file(template_copy)
        output_path = artifact_dir / f".adapter-output.{uuid.uuid4().hex}.xlsx"
        _ensure_within(output_path, artifact_dir, "adapter output")
        try:
            result = adapter(plan_path, template_copy, output_path)
            if not isinstance(result, Mapping):
                raise ArchiveVerificationError("adapter must return a mapping proof")
            if sha256_file(template_copy) != template_before:
                raise ArchiveVerificationError("adapter changed the template copy")
            return self.finalize_local(
                prepared["export_id"], output_path, adapter_result=result
            )
        except Exception as exc:
            try:
                self.store.finish_export(
                    prepared["export_id"],
                    status="failed",
                    local_path=str(artifact_dir),
                    error=str(exc),
                )
            except StoreError:
                pass
            raise
        finally:
            if output_path.exists():
                output_path.unlink()

    def capture_existing_workbook(
        self,
        workbook_path: os.PathLike[str] | str,
        *,
        ledger_result: Mapping[str, Any],
    ) -> Dict[str, Any]:
        """Capture a writer-verified workbook by read-only byte copy.

        ``workbook_path`` may be on Z: because it is opened only for reading.
        Every created file remains under the local spool.  The normalized ledger
        result is embedded in the adapter proof and cryptographically tied to
        the current app-store snapshot and command plan.
        """

        if not ledger_result.get("writer_ok") or ledger_result.get("overlap_pending_ids"):
            raise ArchiveVerificationError(
                "legacy writer result is not fully accounted; workbook cannot be promoted"
            )
        prepared = self.prepare(template_path=workbook_path)
        if prepared.get("status") == "verified":
            return prepared
        artifact_dir = Path(prepared["artifact_dir"])
        plan_path = artifact_dir / "command-plan.json"
        template_copy = artifact_dir / "template-copy.xlsx"
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        if sha256_file(template_copy) != plan["template"]["sha256"]:
            raise ArchiveVerificationError("captured workbook differs from source hash")
        records = list(plan.get("records") or [])
        record_coverage = ledger_result.get("record_coverage")
        conflicts = ledger_result.get("conflicts") or []
        if not isinstance(record_coverage, list):
            raise ArchiveVerificationError(
                "legacy workbook capture requires explicit app-store "
                "work_id+record_version semantic coverage"
            )
        if not isinstance(conflicts, list):
            raise ArchiveVerificationError("legacy workbook conflicts must be a list")
        proof = {
            "status": "partial" if conflicts else "success",
            "snapshot_sha256": plan["snapshot"]["sha256"],
            "command_plan_sha256": sha256_json(plan),
            "rows_considered": len(records),
            "commands_applied": int(ledger_result.get("applied_count") or 0),
            "commands_skipped": int(ledger_result.get("skipped_count") or 0),
            "records_covered": len(record_coverage),
            "record_coverage": record_coverage,
            "conflicts": conflicts,
            "errors": [],
            "output_sha256": sha256_file(template_copy),
            "capture_mode": "existing-writer-verified-workbook",
            "semantic_scope": "legacy-writer-verified-cells-plus-app-store-command-plan",
            "ledger_batch_id": ledger_result.get("batch_id"),
            "ledger_slot": ledger_result.get("slot"),
            "ledger_result_sha256": ledger_result.get("result_sha256"),
            "source_result_sha256": ledger_result.get("source_result_sha256"),
            "external_write_performed": False,
        }
        return self.finalize_local(
            prepared["export_id"], template_copy, adapter_result=proof
        )

    def verify_export(
        self,
        artifact_dir: os.PathLike[str] | str,
        *,
        require_verified: bool = False,
    ) -> Dict[str, Any]:
        artifact = Path(artifact_dir).resolve()
        errors: List[str] = []
        manifest_path = artifact / "manifest.json"
        if not manifest_path.is_file():
            return {"ok": False, "status": "missing", "errors": ["manifest missing"]}
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception as exc:
            return {
                "ok": False,
                "status": "invalid",
                "errors": [f"manifest unreadable: {exc}"],
            }
        if manifest.get("format") != MANIFEST_FORMAT:
            errors.append("manifest format mismatch")
        if manifest.get("manifest_sha256") != _manifest_digest(manifest):
            errors.append("manifest SHA256 mismatch")
        if require_verified and manifest.get("status") != "verified":
            errors.append("artifact is not verified")
        manifest_status = str(manifest.get("status") or "")
        if manifest_status not in {"planned", "partial", "verified", "failed"}:
            errors.append(f"unsupported manifest status: {manifest_status!r}")
        for name, expected in (manifest.get("files") or {}).items():
            member = artifact / name
            try:
                _ensure_within(member, artifact, f"manifest member {name}")
            except ArchiveSafetyError as exc:
                errors.append(str(exc))
                continue
            if not member.is_file():
                errors.append(f"file missing: {name}")
                continue
            if member.stat().st_size != expected.get("size"):
                errors.append(f"size mismatch: {name}")
            if sha256_file(member) != expected.get("sha256"):
                errors.append(f"SHA256 mismatch: {name}")
        snapshot_path = artifact / "snapshot.json"
        plan_path = artifact / "command-plan.json"
        validation_path = artifact / "validation-manifest.json"
        try:
            snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))
            if sha256_json(snapshot) != manifest["snapshot"]["sha256"]:
                errors.append("canonical snapshot SHA256 mismatch")
        except Exception as exc:
            errors.append(f"snapshot validation failed: {exc}")
            snapshot = {}
        try:
            plan = json.loads(plan_path.read_text(encoding="utf-8"))
            if sha256_json(plan) != manifest.get("command_plan_sha256"):
                errors.append("canonical command plan SHA256 mismatch")
            if plan.get("snapshot", {}).get("sha256") != manifest.get("snapshot", {}).get(
                "sha256"
            ):
                errors.append("command plan references another snapshot")
        except Exception as exc:
            errors.append(f"command plan validation failed: {exc}")
            plan = {}
        try:
            validation = json.loads(validation_path.read_text(encoding="utf-8"))
            expected = validation["expected"]
            if expected["work_item_count"] != len(snapshot.get("work_items") or []):
                errors.append("work item count differs from validation manifest")
            if expected["command_record_count"] != len(plan.get("records") or []):
                errors.append("command record count differs from validation manifest")
            if expected["legacy_writer_queue_count"] != len(
                plan.get("legacy_writer_queue") or []
            ):
                errors.append("legacy queue count differs from validation manifest")
        except Exception as exc:
            errors.append(f"validation manifest failed: {exc}")
        db_path = artifact / "db-snapshot.sqlite3"
        if db_path.is_file():
            conn = None
            try:
                conn = sqlite3.connect(f"file:{db_path.as_posix()}?mode=ro", uri=True)
                check = conn.execute("PRAGMA quick_check").fetchone()[0]
                if check != "ok":
                    errors.append(f"SQLite backup quick_check failed: {check}")
                seq = int(
                    conn.execute(
                        "SELECT COALESCE(MAX(id),0) FROM change_event"
                    ).fetchone()[0]
                )
                if seq != int(manifest.get("snapshot", {}).get("change_seq", -1)):
                    errors.append("SQLite backup change_seq differs from snapshot")
            except Exception as exc:
                errors.append(f"SQLite backup validation failed: {exc}")
            finally:
                if conn is not None:
                    conn.close()
        template_copy = artifact / "template-copy.xlsx"
        if template_copy.exists() and sha256_file(template_copy) != manifest.get("template", {}).get(
            "source_sha256"
        ):
            errors.append("template byte-copy differs from source hash")
        coverage: Dict[str, Any] = {}
        coverage_contract = manifest.get("coverage_contract") or {}
        try:
            coverage_version = int(coverage_contract.get("version") or 0)
        except (TypeError, ValueError):
            coverage_version = -1
            errors.append("coverage contract version is invalid")
        if manifest_status in {"verified", "partial"}:
            archive = artifact / "archive.xlsx"
            try:
                _validate_xlsx_container(archive)
                output = manifest.get("output") or {}
                if sha256_file(archive) != output.get("sha256"):
                    errors.append("finalized output hash mismatch")
            except Exception as exc:
                errors.append(f"finalized workbook invalid: {exc}")
            if coverage_version >= COVERAGE_CONTRACT_VERSION:
                try:
                    proof_path = artifact / "adapter-result.json"
                    proof_raw = json.loads(proof_path.read_text(encoding="utf-8"))
                    proof = self._adapter_proof(
                        proof_raw,
                        snapshot_sha256=str(
                            manifest.get("snapshot", {}).get("sha256") or ""
                        ),
                        command_plan_sha256=str(
                            manifest.get("command_plan_sha256") or ""
                        ),
                        records=list(plan.get("records") or []),
                    )
                    coverage = dict(proof.get("coverage") or {})
                    output_sha = (manifest.get("output") or {}).get("sha256")
                    if proof.get("output_sha256") != output_sha:
                        errors.append("adapter output hash differs from manifest output")
                    if coverage != dict(manifest.get("coverage") or {}):
                        errors.append("manifest record coverage differs from adapter proof")
                    complete = bool(coverage.get("complete"))
                    if manifest_status == "verified" and not complete:
                        errors.append("verified artifact has incomplete record coverage")
                    if manifest_status == "partial" and complete:
                        errors.append("partial artifact unexpectedly has complete coverage")
                    result = json.loads(
                        (artifact / "validation-result.json").read_text(encoding="utf-8")
                    )
                    if result.get("status") != manifest_status:
                        errors.append("validation result status differs from manifest")
                    checks = result.get("checks") or {}
                    if bool(checks.get("record_coverage_complete")) != complete:
                        errors.append("validation result coverage completion differs")
                    if dict(checks.get("record_coverage") or {}) != coverage:
                        errors.append("validation result record coverage differs")
                    if checks.get("output_sha256") != output_sha:
                        errors.append("validation result output hash differs")
                except Exception as exc:
                    errors.append(f"record coverage validation failed: {exc}")
            else:
                # v2 도입 전에 정상 승격된 last-good은 record별 ACK 근거로는
                # 절대 쓰지 않지만, 다음 v2 보관본을 만들 원본 템플릿으로는 계속
                # 복구할 수 있어야 한다. 파일/manifest/DB snapshot/최종 XLSX 해시는
                # 위의 동시대 검증 규칙으로 확인하고, ACK 쪽은 별도로 v2를 강제한다.
                # 여기서 구 보관본 자체를 오류로 만들면 첫 v2 렌더가 실패한 순간
                # 정상 last-good까지 잃고 네트워크 원본 fallback에 고착된다.
                coverage = dict(manifest.get("coverage") or {})
                coverage.setdefault("version", max(0, coverage_version))
                coverage["legacy_read_only"] = True
        return {
            "ok": not errors,
            "status": manifest_status or "unknown",
            "export_id": manifest.get("export_id"),
            "coverage": coverage,
            "errors": errors,
        }

    def last_good(self, *, verify: bool = True) -> Optional[Dict[str, Any]]:
        pointer_path = self.spool_dir / "last-good.json"
        if not pointer_path.is_file():
            return None
        pointer = json.loads(pointer_path.read_text(encoding="utf-8"))
        if verify:
            archive = Path(pointer["archive_path"])
            manifest = Path(pointer["manifest_path"])
            if sha256_file(archive) != pointer["archive_sha256"]:
                raise ArchiveVerificationError("last-good archive hash mismatch")
            loaded = json.loads(manifest.read_text(encoding="utf-8"))
            if _manifest_digest(loaded) != pointer["manifest_sha256"]:
                raise ArchiveVerificationError("last-good manifest hash mismatch")
            verified = self.verify_export(manifest.parent, require_verified=True)
            if not verified["ok"]:
                raise ArchiveVerificationError(
                    f"last-good artifact failed verification: {verified['errors']}"
                )
        return pointer


def capture_existing_workbook(
    workbook_path: os.PathLike[str] | str,
    *,
    ledger_result: Mapping[str, Any],
    spool_dir: Optional[os.PathLike[str] | str] = None,
) -> Dict[str, Any]:
    """Module-level convenience wrapper for the legacy scheduler."""

    return ArchiveExporter(default_store(), spool_dir).capture_existing_workbook(
        workbook_path, ledger_result=ledger_result
    )


def _resolve_latest_master_readonly() -> Path:
    """Use the project resolver while suppressing its optional OLD-folder move."""

    from ecount_client import load_config
    from ecount_reconcile import resolve_master
    import ledger_versions

    configured = load_config()["reconcile"]["master_xlsx"]
    # resolve_master normally invokes autoprune as a convenience.  Archive
    # capture has a stricter contract: Z: is read-only.  Holding this lock and
    # temporarily marking pruning done makes concurrent calls skip (not perform)
    # that optional move; the prior process state is restored immediately.
    with _READ_ONLY_MASTER_RESOLVE_LOCK:
        previous = getattr(ledger_versions, "_AUTODONE", False)
        ledger_versions._AUTODONE = True
        try:
            resolved = Path(resolve_master(configured)).resolve()
        finally:
            ledger_versions._AUTODONE = previous
    if not resolved.is_file():
        raise FileNotFoundError(f"latest master workbook not found: {resolved}")
    return resolved


def record_ledger_result(
    *,
    batch_id: Any,
    slot: str,
    result: Mapping[str, Any],
    ok: bool,
    source_result: Optional[os.PathLike[str] | str] = None,
    workbook_path: Optional[os.PathLike[str] | str] = None,
) -> Dict[str, Any]:
    """Record a legacy ``ledger_writer`` result without publishing externally.

    This is the bridge used by ``ledger_db`` after its existing writer adapter
    finishes.  It does not treat a zero exit code as proof: applied/skipped
    entries must be lists, pending IDs must not overlap, and the original result
    file hash is retained when available.  The evidence is atomically spooled
    locally and also audited in the canonical store.

    The function intentionally returns an error object rather than raising so a
    bookkeeping fault cannot turn an already verified legacy writer round into
    an endless locked round.
    """

    try:
        if not isinstance(result, Mapping):
            raise ArchiveVerificationError("ledger result must be a mapping")
        applied = result.get("applied")
        skipped = result.get("skipped")
        if not isinstance(applied, list) or not isinstance(skipped, list):
            raise ArchiveVerificationError(
                "ledger result must contain applied/skipped lists"
            )
        record_coverage = result.get("record_coverage")
        conflicts = result.get("conflicts")
        if record_coverage is not None and not isinstance(record_coverage, list):
            raise ArchiveVerificationError("ledger record_coverage must be a list")
        if conflicts is not None and not isinstance(conflicts, list):
            raise ArchiveVerificationError("ledger conflicts must be a list")

        def pending_ids(rows: List[Any]) -> set[int]:
            ids: set[int] = set()
            for row in rows:
                if not isinstance(row, Mapping):
                    continue
                value = row.get("db_pending_id")
                if str(value or "").isdigit():
                    ids.add(int(value))
            return ids

        applied_ids = pending_ids(applied)
        skipped_ids = pending_ids(skipped)
        overlap = sorted(applied_ids & skipped_ids)
        verified = bool(ok) and not overlap
        source_path = Path(source_result).resolve() if source_result else None
        source_hash = (
            sha256_file(source_path) if source_path is not None and source_path.is_file() else ""
        )
        source_size = (
            source_path.stat().st_size
            if source_path is not None and source_path.is_file()
            else 0
        )
        recorded_at = str(result.get("recorded_at") or result.get("finished_at") or "")
        if not recorded_at and source_path is not None and source_path.is_file():
            recorded_at = datetime.fromtimestamp(
                source_path.stat().st_mtime, tz=timezone.utc
            ).isoformat(timespec="microseconds")
        if not recorded_at:
            recorded_at = _utcnow()
        payload: Dict[str, Any] = {
            "format": "csos-ledger-adapter-result/v1",
            "batch_id": batch_id,
            "slot": str(slot or ""),
            "status": "accounted" if verified else "failed",
            "writer_ok": bool(ok),
            "version": str(result.get("version") or ""),
            "applied_count": len(applied),
            "skipped_count": len(skipped),
            "applied_pending_ids": sorted(applied_ids),
            "skipped_pending_ids": sorted(skipped_ids),
            "overlap_pending_ids": overlap,
            "result_sha256": sha256_json(result),
            "source_result": str(source_path) if source_path else "",
            "source_result_sha256": source_hash,
            "source_result_size": source_size,
            "external_write_performed": False,
            "recorded_at": recorded_at,
        }
        if record_coverage is not None:
            payload["record_coverage"] = json.loads(canonical_json(record_coverage))
            payload["conflicts"] = json.loads(canonical_json(conflicts or []))
        store = default_store()
        exporter = ArchiveExporter(store)
        ledger_dir = exporter.spool_dir / "ledger-results"
        _ensure_within(ledger_dir, exporter.spool_dir, "ledger result directory")
        ledger_dir.mkdir(parents=True, exist_ok=True)
        safe_batch = "".join(c for c in str(batch_id) if c.isalnum() or c in "-_") or "unknown"
        evidence_path = ledger_dir / f"batch-{safe_batch}.json"
        _atomic_write_json(evidence_path, payload)
        payload["local_evidence"] = str(evidence_path)
        payload["local_evidence_sha256"] = sha256_file(evidence_path)

        def set_retry(key: str, value: Mapping[str, Any], token: str) -> None:
            for attempt in range(5):
                current = store.get_setting(key)
                if current.get("value") == value:
                    return
                version = int(current.get("record_version") or 0)
                try:
                    store.set_setting(
                        key,
                        dict(value),
                        expected_version=version,
                        actor="ledger-adapter",
                        source="ledger_writer",
                        evidence=f"batch {batch_id} result SHA256 {payload['result_sha256']}",
                        source_ref=str(source_path or evidence_path),
                        source_sha256=source_hash or payload["local_evidence_sha256"],
                        idempotency_key=f"{token}:v{version}",
                    )
                    return
                except VersionConflict:
                    if attempt < 4:
                        continue
                    raise

        stable_value = dict(payload)
        set_retry(
            f"legacy_ledger.result.{safe_batch}",
            stable_value,
            f"ledger-result:{safe_batch}:{payload['result_sha256']}",
        )
        latest = store.get_setting("legacy_ledger.last_result")
        previous = latest.get("value") if isinstance(latest.get("value"), dict) else {}
        try:
            is_newer = int(batch_id) >= int(previous.get("batch_id", -1))
        except (TypeError, ValueError):
            is_newer = str(payload["recorded_at"]) >= str(previous.get("recorded_at") or "")
        if is_newer:
            set_retry(
                "legacy_ledger.last_result",
                stable_value,
                f"ledger-last:{safe_batch}:{payload['result_sha256']}",
            )
        if not verified:
            return {
                "ok": False,
                "result_recorded": True,
                **payload,
                "error": (
                    f"writer_ok={bool(ok)}, overlapping pending IDs={overlap[:20]}"
                ),
            }
        try:
            master = Path(workbook_path).resolve() if workbook_path else _resolve_latest_master_readonly()
            captured = exporter.capture_existing_workbook(
                master, ledger_result=payload
            )
        except Exception as capture_exc:
            return {
                "ok": False,
                "result_recorded": True,
                **payload,
                "status": "archive_capture_failed",
                "archive_error": f"{type(capture_exc).__name__}: {capture_exc}"[:1_000],
            }
        return {
            "ok": True,
            "result_recorded": True,
            **payload,
            "archive_capture": captured,
        }
    except Exception as exc:
        return {
            "ok": False,
            "status": "recording_failed",
            "batch_id": batch_id,
            "error": f"{type(exc).__name__}: {exc}"[:1_000],
            "external_write_performed": False,
        }


def _make_minimal_xlsx(path: Path) -> None:
    """Create a tiny OOXML container for self-test without workbook libraries."""

    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '</Types>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheets/></workbook>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("xl/workbook.xml", workbook)


def self_test() -> bool:
    """Verify local-only planning, adapter proof, atomic promotion and last-good."""

    with tempfile.TemporaryDirectory(prefix="archive_export_selftest_") as tmp:
        root = Path(tmp)
        store = AppStore(root / "store.db").initialize()
        imported = store.shadow_import(
            import_id="archive-self-import",
            sheet="06_업무관리",
            business_key="UJ-ARCHIVE-001",
            business_key_col="UJ프로젝트",
            row_number=88,
            kind="돌발AS",
            fields={"청구상태": "작업완료", "공급가액": 123_450},
            apply_if_missing=True,
            source_file="original.xlsx",
            source_sha256="c" * 64,
            idempotency_key="archive-shadow-1",
        )
        assert imported["status"] == "created"
        template = root / "template.xlsx"
        _make_minimal_xlsx(template)
        template_hash = sha256_file(template)
        exporter = ArchiveExporter(store, root / "spool")
        planned = exporter.prepare(template_path=template)
        assert planned["status"] == "planned"
        assert exporter.verify_export(planned["artifact_dir"])["ok"]
        assert sha256_file(Path(planned["template_copy"])) == template_hash
        planned_again = exporter.prepare(template_path=template)
        assert planned_again["export_id"] == planned["export_id"]
        plan_payload = json.loads(Path(planned["command_plan"]).read_text(encoding="utf-8"))
        assert plan_payload["legacy_writer_queue"]

        def adapter(plan_path: Path, template_copy: Path, output_path: Path) -> Mapping[str, Any]:
            plan = json.loads(plan_path.read_text(encoding="utf-8"))
            shutil.copyfile(template_copy, output_path)
            return {
                "status": "success",
                "snapshot_sha256": plan["snapshot"]["sha256"],
                "command_plan_sha256": sha256_json(plan),
                "rows_considered": len(plan["records"]),
                "commands_applied": len(plan["legacy_writer_queue"]),
                "record_coverage": [
                    {
                        "work_id": record["work_id"],
                        "business_key": record["business_key"],
                        "record_version": record["record_version"],
                        "outcome": "applied",
                    }
                    for record in plan["records"]
                ],
                "conflicts": [],
                "errors": [],
                "output_sha256": sha256_file(output_path),
            }

        verified = exporter.run_local(
            template_path=template, adapter=adapter, dry_run=False
        )
        assert verified["status"] == "verified"
        assert exporter.verify_export(
            verified["artifact_dir"], require_verified=True
        )["ok"]
        pointer = exporter.last_good()
        assert pointer and pointer["export_id"] == verified["export_id"]
        assert pointer["external_write_performed"] is False
        assert store.export_run(verified["export_id"])["status"] == "verified"

        # v2 이전 정상 보관본은 새 outbox ACK에는 못 쓰지만, v2 렌더가 실패했을 때
        # 다음 회차의 읽기 전용 템플릿/복구 원본으로 계속 검증 가능해야 한다.
        legacy_dir = root / "legacy-v1-last-good"
        shutil.copytree(Path(verified["artifact_dir"]), legacy_dir)
        legacy_manifest_path = legacy_dir / "manifest.json"
        legacy_manifest = json.loads(legacy_manifest_path.read_text(encoding="utf-8"))
        legacy_manifest.pop("coverage_contract", None)
        legacy_manifest.pop("coverage", None)
        legacy_manifest["manifest_sha256"] = _manifest_digest(legacy_manifest)
        _atomic_write_json(legacy_manifest_path, legacy_manifest)
        legacy_check = exporter.verify_export(legacy_dir, require_verified=True)
        assert legacy_check["ok"] and legacy_check["coverage"].get("legacy_read_only"), \
            legacy_check

        # A sidecar is ackable only when its full canonical record hash, exact
        # revision and worksheet semantic hash all agree with the command plan.
        sidecar_record = {
            "work_id": "wrk-sidecar-selftest",
            "business_key": "UJ-SIDECAR-001",
            "record_version": 3,
            "kind": "돌발AS",
            "fields": {"프로젝트NO": "UJ-SIDECAR-001", "비고": "모호 행"},
        }
        sidecar_record_sha = hashlib.sha256(
            canonical_json(sidecar_record).encode("utf-8")
        ).hexdigest()
        sidecar_entry = {
            "work_id": sidecar_record["work_id"],
            "record_version": sidecar_record["record_version"],
            "business_key": sidecar_record["business_key"],
            "reason": "ambiguous-no-anchor",
            "record_sha256": sidecar_record_sha,
        }
        sidecar_sha = sha256_json(
            {"format": "csos-appdb-sidecar/v1", "records": [sidecar_entry]}
        )
        sidecar_proof = {
            "status": "success",
            "record_coverage": [
                {
                    **sidecar_entry,
                    "outcome": "archived_sidecar",
                    "sheet": "99_AppDB_미매칭보관",
                    "sidecar_semantic_sha256": sidecar_sha,
                }
            ],
            "conflicts": [],
            "sidecar": {
                "format": "csos-appdb-sidecar/v1",
                "sheet": "99_AppDB_미매칭보관",
                "records": 1,
                "semantic_sha256": sidecar_sha,
                "entries": [sidecar_entry],
            },
            "validation": {
                "sidecar": {
                    "sheet": "99_AppDB_미매칭보관",
                    "records": 1,
                    "semantic_sha256": sidecar_sha,
                }
            },
        }
        checked_sidecar = ArchiveExporter._validate_record_coverage(
            sidecar_proof, [sidecar_record]
        )
        assert checked_sidecar["coverage"]["complete"]
        tampered = json.loads(canonical_json(sidecar_proof))
        tampered["record_coverage"][0]["record_sha256"] = "0" * 64
        try:
            ArchiveExporter._validate_record_coverage(tampered, [sidecar_record])
        except ArchiveVerificationError:
            pass
        else:
            raise AssertionError("tampered sidecar record hash was accepted")

        # A conflicted record is useful forensic output, but is neither fully
        # verified nor eligible to replace the previous last-good pointer.
        second_created = store.create_work(
            kind="돌발AS",
            business_key="UJ-ARCHIVE-002",
            fields={"청구상태": "작업완료"},
            actor="selftest",
            source="selftest",
            evidence="partial coverage self-test",
            idempotency_key="archive-partial-work",
        )["work"]

        def partial_adapter(
            plan_path: Path, template_copy: Path, output_path: Path
        ) -> Mapping[str, Any]:
            plan = json.loads(plan_path.read_text(encoding="utf-8"))
            shutil.copyfile(template_copy, output_path)
            covered = [
                record
                for record in plan["records"]
                if record["work_id"] != second_created["id"]
            ]
            conflicted = next(
                record
                for record in plan["records"]
                if record["work_id"] == second_created["id"]
            )
            return {
                "status": "partial",
                "snapshot_sha256": plan["snapshot"]["sha256"],
                "command_plan_sha256": sha256_json(plan),
                "rows_considered": len(plan["records"]),
                "record_coverage": [
                    {
                        "work_id": record["work_id"],
                        "business_key": record["business_key"],
                        "record_version": record["record_version"],
                        "outcome": "unchanged",
                    }
                    for record in covered
                ],
                "conflicts": [
                    {
                        "work_id": conflicted["work_id"],
                        "business_key": conflicted["business_key"],
                        "record_version": conflicted["record_version"],
                        "reason": "synthetic-conflict",
                    }
                ],
                "errors": [],
                "output_sha256": sha256_file(output_path),
            }

        partial = exporter.run_local(
            template_path=template, adapter=partial_adapter, dry_run=False
        )
        assert partial["status"] == "partial"
        assert exporter.verify_export(partial["artifact_dir"])["ok"]
        assert not exporter.verify_export(
            partial["artifact_dir"], require_verified=True
        )["ok"]
        assert exporter.last_good()["export_id"] == verified["export_id"]
        assert store.export_run(partial["export_id"])["status"] == "partial"

        result_file = root / "ledger-result.json"
        ledger_result = {
            "applied": [{"db_pending_id": 1, "sheet": "02_돌발AS접수"}],
            "skipped": [{"db_pending_id": 2, "사유": "이미 동일"}],
            "version": "v-self-test",
        }
        _atomic_write_json(result_file, ledger_result)
        old_db = os.environ.get("COUPANG_APP_DB_PATH")
        old_spool = os.environ.get("COUPANG_ARCHIVE_SPOOL")
        os.environ["COUPANG_APP_DB_PATH"] = str(store.db_path)
        os.environ["COUPANG_ARCHIVE_SPOOL"] = str(root / "spool")
        try:
            recorded = record_ledger_result(
                batch_id=77,
                slot="self-test",
                result=ledger_result,
                ok=True,
                source_result=result_file,
                workbook_path=template,
            )
            assert not recorded["ok"] and Path(recorded["local_evidence"]).is_file()
            assert recorded["status"] == "archive_capture_failed"
            assert "semantic coverage" in recorded["archive_error"]
            assert default_store().get_setting("legacy_ledger.last_result")["value"][
                "batch_id"
            ] == 77

            snapshot = default_store().snapshot_payload()
            covered_result = {
                **ledger_result,
                "record_coverage": [
                    {
                        "work_id": work["id"],
                        "business_key": work["business_key"],
                        "record_version": work["record_version"],
                        "outcome": "unchanged",
                    }
                    for work in snapshot["work_items"]
                ],
                "conflicts": [],
            }
            _atomic_write_json(result_file, covered_result)
            captured = record_ledger_result(
                batch_id=78,
                slot="self-test",
                result=covered_result,
                ok=True,
                source_result=result_file,
                workbook_path=template,
            )
            assert captured["ok"]
            assert captured["archive_capture"]["status"] == "verified"
            assert default_store().get_setting("legacy_ledger.last_result")["value"][
                "batch_id"
            ] == 78
            seq_before_replay = default_store().status()["change_seq"]
            recorded_again = record_ledger_result(
                batch_id=78,
                slot="self-test",
                result=covered_result,
                ok=True,
                source_result=result_file,
                workbook_path=template,
            )
            assert recorded_again["ok"]
            assert default_store().status()["change_seq"] == seq_before_replay
        finally:
            if old_db is None:
                os.environ.pop("COUPANG_APP_DB_PATH", None)
            else:
                os.environ["COUPANG_APP_DB_PATH"] = old_db
            if old_spool is None:
                os.environ.pop("COUPANG_ARCHIVE_SPOOL", None)
            else:
                os.environ["COUPANG_ARCHIVE_SPOOL"] = old_spool

        try:
            ArchiveExporter(store, r"Z:\forbidden_archive_spool")
            raise AssertionError("Z: spool was accepted")
        except ArchiveSafetyError:
            pass
    return True


def _main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Create a local, verifiable Excel archive command plan"
    )
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--db")
    parser.add_argument("--template")
    parser.add_argument("--spool")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="prepare only (the CLI never invokes an adapter)",
    )
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("archive_export self-test: OK")
        return 0
    if args.db and args.template:
        exporter = ArchiveExporter(AppStore(args.db), args.spool)
        result = exporter.run_local(template_path=args.template, dry_run=True)
        print(canonical_json(result))
        return 0
    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
