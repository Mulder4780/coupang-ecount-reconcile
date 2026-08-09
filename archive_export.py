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
                    "op": "upsert_record",
                    "work_id": work["id"],
                    "record_version": work["record_version"],
                    "kind": work["kind"],
                    "business_key": work["business_key"],
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
            if target:
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
                "execution": "explicit-local-callable-only",
                "call_signature": "adapter(plan_path, template_copy_path, output_path) -> result",
                "high_level_member": "records",
                "legacy_queue_member": "legacy_writer_queue",
                "requirements": [
                    "write a new output_path; never mutate template_copy_path",
                    "enforce per-command idempotency_key",
                    "report snapshot_sha256 and command_plan_sha256",
                    "report rows_considered and an empty errors list before verification",
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
                "output is a valid OOXML ZIP with workbook.xml",
                "only then atomically advance last-good.json",
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
        snapshot, snapshot_sha = self.store.snapshot()
        seq = int(snapshot.get("change_seq", 0))
        template_token = template_sha[:12] if template_sha else "no-template"
        export_id = f"exp-{seq:012d}-{snapshot_sha[:16]}-{template_token}"
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
        status = "verified" if manifest.get("status") == "verified" else "planned"
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
    def _adapter_proof(
        result: Mapping[str, Any],
        *,
        snapshot_sha256: str,
        command_plan_sha256: str,
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
        if proof["status"] != "success":
            raise ArchiveVerificationError(
                f"adapter did not report success: {proof.get('status')}"
            )
        if proof["snapshot_sha256"] != snapshot_sha256:
            raise ArchiveVerificationError("adapter proof references another snapshot")
        if proof["command_plan_sha256"] != command_plan_sha256:
            raise ArchiveVerificationError("adapter proof references another command plan")
        if not isinstance(proof["rows_considered"], int) or proof["rows_considered"] < 0:
            raise ArchiveVerificationError("adapter rows_considered must be a non-negative int")
        if not isinstance(proof["errors"], list) or proof["errors"]:
            raise ArchiveVerificationError("adapter reported unresolved errors")
        return proof

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
        proof = self._adapter_proof(
            adapter_result,
            snapshot_sha256=manifest["snapshot"]["sha256"],
            command_plan_sha256=manifest["command_plan_sha256"],
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
        proof["verified_at"] = _utcnow()
        proof_path = artifact_dir / "adapter-result.json"
        _atomic_write_json(proof_path, proof)

        validation_result = {
            "format": VALIDATION_FORMAT,
            "export_id": export_id,
            "status": "verified",
            "verified_at": _utcnow(),
            "checks": {
                "snapshot_sha256": True,
                "command_plan_sha256": True,
                "template_unchanged": True,
                "adapter_errors_empty": True,
                "output_sha256": archive_entry["sha256"],
                "ooxml_container": container,
                "external_write_performed": False,
            },
        }
        result_path = artifact_dir / "validation-result.json"
        _atomic_write_json(result_path, validation_result)

        manifest["status"] = "verified"
        manifest["verified_at"] = _utcnow()
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

        verification = self.verify_export(artifact_dir, require_verified=True)
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
        proof = {
            "status": "success",
            "snapshot_sha256": plan["snapshot"]["sha256"],
            "command_plan_sha256": sha256_json(plan),
            "rows_considered": len(plan.get("records") or []),
            "commands_applied": int(ledger_result.get("applied_count") or 0),
            "commands_skipped": int(ledger_result.get("skipped_count") or 0),
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
        if manifest.get("status") == "verified":
            archive = artifact / "archive.xlsx"
            try:
                _validate_xlsx_container(archive)
                output = manifest.get("output") or {}
                if sha256_file(archive) != output.get("sha256"):
                    errors.append("verified output hash mismatch")
            except Exception as exc:
                errors.append(f"verified workbook invalid: {exc}")
        return {
            "ok": not errors,
            "status": manifest.get("status", "unknown"),
            "export_id": manifest.get("export_id"),
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
            assert recorded["ok"] and Path(recorded["local_evidence"]).is_file()
            assert recorded["archive_capture"]["status"] == "verified"
            assert default_store().get_setting("legacy_ledger.last_result")["value"][
                "batch_id"
            ] == 77
            seq_before_replay = default_store().status()["change_seq"]
            recorded_again = record_ledger_result(
                batch_id=77,
                slot="self-test",
                result=ledger_result,
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
