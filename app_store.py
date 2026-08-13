"""Transactional SQLite system-of-record for the Coupang work application.

This module is deliberately independent from the legacy Excel queue.  It stores
application writes, provenance and audit history first; Excel is an archive
consumer of a deterministic snapshot (see :mod:`archive_export`).

Only Python's standard library is used so the store can run in the desktop app,
scheduled jobs, or a small service without an additional runtime dependency.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sqlite3
import tempfile
import threading
import uuid
from contextlib import contextmanager, nullcontext
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Mapping, Optional, Sequence, Tuple


SCHEMA_VERSION = 2
DEFAULT_BUSY_TIMEOUT_MS = 30_000

# Canonical kind/core-column routing used by the thin legacy compatibility
# layer.  Unknown sheets still work under a namespaced ``sheet:<name>`` kind.
SHEET_SPECS: Dict[str, Dict[str, str]] = {
    "02_돌발AS접수": {
        "kind": "돌발AS",
        "public_id": "접수ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "진행상태",
    },
    "03_현장작업실적": {
        "kind": "현장작업",
        "public_id": "작업ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "완료여부",
    },
    "04_정기점검": {
        "kind": "정기점검",
        "public_id": "점검ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "점검상태",
    },
    "05_신규납품설치": {
        "kind": "신규납품설치",
        "public_id": "업무ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "진행상태",
    },
    "06_거래서류청구수금": {
        "kind": "정산",
        "public_id": "정산ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "청구상태",
    },
    "15_세금계산서관리": {
        "kind": "세금계산서",
        "public_id": "계산서관리ID",
        "business_key": "계산서관리ID",
        "relation_id": "정산ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "발행상태(자동)",
        "issued_at": "실제발행일",
    },
    "16_입금수금관리": {
        "kind": "입금수금",
        "public_id": "입금관리ID",
        "business_key": "입금관리ID",
        "relation_id": "정산ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        # v586에는 상태 열이 없다. 입금일·입금액·미수금액으로 화면에서
        # 파생할 수는 있어도 존재하지 않는 Excel 열을 정본 열로 만들지 않는다.
        "status": "",
    },
    "13_PO발주관리": {
        "kind": "PO발주",
        "public_id": "PO관리ID",
        "project_no": "프로젝트NO",
        "camp_name": "캠프명",
        "status": "PO상태(자동)",
    },
}


def _management_identity_field(kind: str) -> str:
    """Return the v586 management-ID field used as a canonical row identity."""

    for spec in SHEET_SPECS.values():
        if str(spec.get("kind") or "") == str(kind or ""):
            return str(spec.get("business_key") or "")
    return ""


class StoreError(RuntimeError):
    """Base error for the application store."""


class ValidationError(StoreError):
    """The caller supplied an invalid or incomplete request."""


class NotFoundError(StoreError):
    """A requested record does not exist (or is soft-deleted)."""


class VersionConflict(StoreError):
    """Optimistic-lock failure."""

    def __init__(self, work_id: str, expected: int, actual: Optional[int]) -> None:
        self.work_id = work_id
        self.expected = expected
        self.actual = actual
        super().__init__(
            f"record version conflict for {work_id}: expected {expected}, actual {actual}"
        )


class IdempotencyConflict(StoreError):
    """An idempotency key was reused with a different request body."""


SCHEMA_STATEMENTS: Tuple[str, ...] = (
    """
    CREATE TABLE IF NOT EXISTS schema_version (
        version INTEGER PRIMARY KEY,
        applied_at TEXT NOT NULL,
        checksum TEXT NOT NULL
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS app_setting (
        key TEXT PRIMARY KEY,
        value_json TEXT NOT NULL,
        value_type TEXT NOT NULL,
        record_version INTEGER NOT NULL DEFAULT 1 CHECK (record_version > 0),
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL DEFAULT ''
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS work_item (
        id TEXT PRIMARY KEY,
        kind TEXT NOT NULL,
        public_id TEXT,
        business_key TEXT NOT NULL,
        project_no TEXT,
        camp_name TEXT,
        status TEXT NOT NULL DEFAULT '',
        source TEXT NOT NULL DEFAULT '',
        evidence TEXT NOT NULL DEFAULT '',
        source_ref TEXT NOT NULL DEFAULT '',
        source_sha256 TEXT NOT NULL DEFAULT '',
        source_observed_at TEXT,
        record_version INTEGER NOT NULL DEFAULT 1 CHECK (record_version > 0),
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL DEFAULT '',
        deleted_at TEXT,
        deleted_by TEXT,
        UNIQUE (kind, business_key)
    )
    """,
    """
    CREATE UNIQUE INDEX IF NOT EXISTS ux_work_public_id
    ON work_item(public_id)
    WHERE public_id IS NOT NULL AND public_id <> ''
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_work_kind_status
    ON work_item(kind, status, deleted_at)
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_work_project
    ON work_item(project_no, deleted_at)
    """,
    """
    CREATE TABLE IF NOT EXISTS work_field (
        work_id TEXT NOT NULL,
        field_key TEXT NOT NULL,
        value_json TEXT NOT NULL,
        value_type TEXT NOT NULL,
        source TEXT NOT NULL DEFAULT '',
        evidence TEXT NOT NULL DEFAULT '',
        source_ref TEXT NOT NULL DEFAULT '',
        source_sha256 TEXT NOT NULL DEFAULT '',
        source_observed_at TEXT,
        field_version INTEGER NOT NULL DEFAULT 1 CHECK (field_version > 0),
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL DEFAULT '',
        PRIMARY KEY (work_id, field_key),
        FOREIGN KEY (work_id) REFERENCES work_item(id) ON DELETE CASCADE
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS change_event (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_id TEXT NOT NULL UNIQUE,
        work_id TEXT,
        aggregate_type TEXT NOT NULL,
        aggregate_key TEXT NOT NULL,
        action TEXT NOT NULL,
        actor TEXT NOT NULL DEFAULT '',
        request_id TEXT,
        idempotency_key TEXT,
        expected_version INTEGER,
        before_json TEXT,
        after_json TEXT,
        source TEXT NOT NULL DEFAULT '',
        evidence TEXT NOT NULL DEFAULT '',
        source_ref TEXT NOT NULL DEFAULT '',
        source_sha256 TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL,
        reverted_event_id TEXT,
        FOREIGN KEY (work_id) REFERENCES work_item(id)
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_change_work
    ON change_event(work_id, id)
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_change_aggregate
    ON change_event(aggregate_type, aggregate_key, id)
    """,
    """
    CREATE TABLE IF NOT EXISTS idempotency_key (
        scope TEXT NOT NULL,
        key TEXT NOT NULL,
        request_hash TEXT NOT NULL,
        response_json TEXT NOT NULL,
        created_at TEXT NOT NULL,
        expires_at TEXT,
        PRIMARY KEY (scope, key)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS outbox (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_id TEXT NOT NULL,
        topic TEXT NOT NULL,
        aggregate_type TEXT NOT NULL,
        aggregate_key TEXT NOT NULL,
        payload_json TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'pending'
            CHECK (status IN ('pending', 'leased', 'done', 'failed')),
        attempts INTEGER NOT NULL DEFAULT 0,
        available_at TEXT NOT NULL,
        leased_at TEXT,
        lease_token TEXT,
        finished_at TEXT,
        last_error TEXT,
        created_at TEXT NOT NULL,
        UNIQUE (event_id, topic)
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_outbox_ready
    ON outbox(status, available_at, id)
    """,
    """
    CREATE TABLE IF NOT EXISTS export_run (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        export_id TEXT NOT NULL UNIQUE,
        snapshot_seq INTEGER NOT NULL,
        snapshot_sha256 TEXT NOT NULL,
        template_sha256 TEXT NOT NULL DEFAULT '',
        plan_sha256 TEXT NOT NULL DEFAULT '',
        manifest_sha256 TEXT NOT NULL DEFAULT '',
        mode TEXT NOT NULL,
        status TEXT NOT NULL,
        local_path TEXT NOT NULL DEFAULT '',
        error TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL,
        finished_at TEXT
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_export_snapshot
    ON export_run(snapshot_seq, snapshot_sha256, status)
    """,
    """
    CREATE TABLE IF NOT EXISTS import_conflict (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_id TEXT NOT NULL,
        sheet TEXT NOT NULL,
        business_key TEXT NOT NULL DEFAULT '',
        row_number INTEGER,
        field_key TEXT NOT NULL DEFAULT '',
        incoming_json TEXT,
        current_json TEXT,
        reason TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'open'
            CHECK (status IN ('open', 'resolved', 'ignored')),
        created_at TEXT NOT NULL,
        resolved_at TEXT,
        resolution TEXT NOT NULL DEFAULT '',
        UNIQUE (import_id, sheet, row_number, field_key, reason)
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_import_conflict_open
    ON import_conflict(status, import_id, id)
    """,
    """
    CREATE TABLE IF NOT EXISTS public_id_sequence (
        kind TEXT NOT NULL,
        day TEXT NOT NULL,
        prefix TEXT NOT NULL,
        last_number INTEGER NOT NULL CHECK (last_number >= 0),
        updated_at TEXT NOT NULL,
        PRIMARY KEY (kind, day, prefix)
    )
    """,
    # A shadow locator is intentionally separate from the canonical record.  It
    # retains the exact Excel sheet/key/row evidence without making row numbers
    # part of the application's identity.
    """
    CREATE TABLE IF NOT EXISTS shadow_import_row (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_id TEXT NOT NULL,
        sheet TEXT NOT NULL,
        business_key TEXT NOT NULL,
        business_key_col TEXT NOT NULL DEFAULT '프로젝트NO',
        row_number INTEGER NOT NULL CHECK (row_number > 0),
        kind TEXT NOT NULL,
        work_id TEXT,
        row_sha256 TEXT NOT NULL,
        status TEXT NOT NULL,
        source_ref TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL,
        UNIQUE (import_id, sheet, row_number),
        FOREIGN KEY (work_id) REFERENCES work_item(id)
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_shadow_business_key
    ON shadow_import_row(kind, business_key, created_at)
    """,
    # Objective completion proof used to live only in the legacy queue DB.
    # Keep every verified decision in the application system-of-record so a
    # source refresh, archive export, or UI restart cannot silently lose it.
    """
    CREATE TABLE IF NOT EXISTS completion_evidence (
        id TEXT PRIMARY KEY,
        owner TEXT NOT NULL DEFAULT '',
        task_kind TEXT NOT NULL,
        record_id TEXT NOT NULL,
        project_no TEXT NOT NULL DEFAULT '',
        status TEXT NOT NULL,
        completed_on TEXT NOT NULL,
        basis TEXT NOT NULL,
        source TEXT NOT NULL DEFAULT '',
        source_ref TEXT NOT NULL DEFAULT '',
        evidence_sha256 TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
        first_seen TEXT NOT NULL,
        last_seen TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        UNIQUE(owner, task_kind, record_id)
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS ix_completion_active
    ON completion_evidence(active, owner, task_kind, completed_on)
    """,
)


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def _json_default(value: Any) -> Any:
    if isinstance(value, datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return {"$decimal": format(value, "f")}
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"unsupported JSON value: {type(value).__name__}")


def canonical_json(value: Any) -> str:
    """Return a stable UTF-8 JSON representation suitable for hashing."""

    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=_json_default,
    )


def sha256_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def _value_type(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, Decimal):
        return "decimal"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    if isinstance(value, (datetime, date)):
        return "date"
    if isinstance(value, (list, dict)):
        return "json"
    return "json"


def _require_text(name: str, value: Any, max_len: int = 500) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValidationError(f"{name} is required")
    if len(text) > max_len:
        raise ValidationError(f"{name} is too long (max {max_len})")
    return text


def _dict_row(row: Optional[sqlite3.Row]) -> Optional[Dict[str, Any]]:
    return dict(row) if row is not None else None


class AppStore:
    """SQLite-backed canonical store with audit, outbox and shadow import APIs."""

    def __init__(self, db_path: Optional[os.PathLike[str] | str] = None) -> None:
        default = Path(__file__).resolve().parent / "db" / "app_store.db"
        self.db_path = Path(
            db_path or os.environ.get("COUPANG_APP_DB_PATH") or default
        ).resolve()
        self._initialized = False
        self._init_lock = threading.Lock()

    def _raw_connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(
            str(self.db_path),
            timeout=DEFAULT_BUSY_TIMEOUT_MS / 1000,
            isolation_level=None,
        )
        conn.row_factory = sqlite3.Row
        conn.execute(f"PRAGMA busy_timeout={DEFAULT_BUSY_TIMEOUT_MS}")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA synchronous=FULL")
        return conn

    def initialize(self) -> "AppStore":
        if self._initialized:
            return self
        with self._init_lock:
            if self._initialized:
                return self
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            conn = self._raw_connect()
            try:
                # journal_mode cannot be changed inside a transaction.
                mode = str(conn.execute("PRAGMA journal_mode=WAL").fetchone()[0]).lower()
                if mode != "wal":
                    raise StoreError(f"SQLite WAL mode unavailable: {mode}")
                conn.execute("BEGIN IMMEDIATE")
                try:
                    for statement in SCHEMA_STATEMENTS:
                        conn.execute(statement)
                    checksum = hashlib.sha256(
                        "\n".join(s.strip() for s in SCHEMA_STATEMENTS).encode("utf-8")
                    ).hexdigest()
                    row = conn.execute(
                        "SELECT checksum FROM schema_version WHERE version=?",
                        (SCHEMA_VERSION,),
                    ).fetchone()
                    if row is None:
                        conn.execute(
                            "INSERT INTO schema_version(version, applied_at, checksum) VALUES(?,?,?)",
                            (SCHEMA_VERSION, _utcnow(), checksum),
                        )
                    elif row["checksum"] != checksum:
                        raise StoreError(
                            "schema checksum differs for the installed schema version"
                        )
                    newer = conn.execute(
                        "SELECT MAX(version) AS version FROM schema_version"
                    ).fetchone()["version"]
                    if int(newer or 0) > SCHEMA_VERSION:
                        raise StoreError(
                            f"database schema {newer} is newer than supported {SCHEMA_VERSION}"
                        )
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise
            finally:
                conn.close()
            self._initialized = True
        return self

    def connect(self) -> sqlite3.Connection:
        self.initialize()
        return self._raw_connect()

    @contextmanager
    def transaction(self, *, immediate: bool = True) -> Iterator[sqlite3.Connection]:
        conn = self.connect()
        try:
            conn.execute("BEGIN IMMEDIATE" if immediate else "BEGIN")
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @contextmanager
    def reader(self) -> Iterator[sqlite3.Connection]:
        conn = self.connect()
        try:
            yield conn
        finally:
            conn.close()

    @staticmethod
    def _new_id(prefix: str) -> str:
        return f"{prefix}_{uuid.uuid4().hex}"

    @staticmethod
    def _request_hash(payload: Mapping[str, Any]) -> str:
        return sha256_json(payload)

    def _idempotency_replay(
        self,
        conn: sqlite3.Connection,
        scope: str,
        key: Optional[str],
        payload: Mapping[str, Any],
    ) -> Tuple[Optional[Dict[str, Any]], str]:
        request_hash = self._request_hash(payload)
        if not key:
            return None, request_hash
        key = _require_text("idempotency_key", key, 300)
        row = conn.execute(
            "SELECT request_hash, response_json, expires_at FROM idempotency_key "
            "WHERE scope=? AND key=?",
            (scope, key),
        ).fetchone()
        if row is None:
            return None, request_hash
        if row["expires_at"] and row["expires_at"] < _utcnow():
            conn.execute(
                "DELETE FROM idempotency_key WHERE scope=? AND key=?", (scope, key)
            )
            return None, request_hash
        if row["request_hash"] != request_hash:
            raise IdempotencyConflict(
                f"idempotency key {key!r} in scope {scope!r} was reused with different input"
            )
        response = json.loads(row["response_json"])
        if isinstance(response, dict):
            response["idempotent_replay"] = True
        return response, request_hash

    @staticmethod
    def _save_idempotency(
        conn: sqlite3.Connection,
        scope: str,
        key: Optional[str],
        request_hash: str,
        response: Mapping[str, Any],
        *,
        ttl_seconds: Optional[int] = None,
    ) -> None:
        if not key:
            return
        expires = None
        if ttl_seconds:
            expires = (
                datetime.now(timezone.utc) + timedelta(seconds=ttl_seconds)
            ).isoformat(timespec="microseconds")
        conn.execute(
            "INSERT INTO idempotency_key(scope,key,request_hash,response_json,created_at,expires_at) "
            "VALUES(?,?,?,?,?,?)",
            (scope, key, request_hash, canonical_json(response), _utcnow(), expires),
        )

    @staticmethod
    def _field_provenance(
        field_key: str,
        field_meta: Optional[Mapping[str, Mapping[str, Any]]],
        defaults: Mapping[str, Any],
    ) -> Dict[str, Any]:
        meta = dict((field_meta or {}).get(field_key, {}))
        return {
            "source": str(meta.get("source", defaults.get("source", "")) or ""),
            "evidence": str(meta.get("evidence", defaults.get("evidence", "")) or ""),
            "source_ref": str(meta.get("source_ref", defaults.get("source_ref", "")) or ""),
            "source_sha256": str(
                meta.get("source_sha256", defaults.get("source_sha256", "")) or ""
            ),
            "source_observed_at": meta.get(
                "source_observed_at", defaults.get("source_observed_at")
            ),
        }

    def _put_field(
        self,
        conn: sqlite3.Connection,
        work_id: str,
        field_key: str,
        value: Any,
        actor: str,
        provenance: Mapping[str, Any],
        now: str,
    ) -> None:
        field_key = _require_text("field_key", field_key, 300)
        conn.execute(
            """
            INSERT INTO work_field(
                work_id,field_key,value_json,value_type,source,evidence,source_ref,
                source_sha256,source_observed_at,field_version,updated_at,updated_by
            ) VALUES(?,?,?,?,?,?,?,?,?,1,?,?)
            ON CONFLICT(work_id,field_key) DO UPDATE SET
                value_json=excluded.value_json,
                value_type=excluded.value_type,
                source=excluded.source,
                evidence=excluded.evidence,
                source_ref=excluded.source_ref,
                source_sha256=excluded.source_sha256,
                source_observed_at=excluded.source_observed_at,
                field_version=work_field.field_version+1,
                updated_at=excluded.updated_at,
                updated_by=excluded.updated_by
            """,
            (
                work_id,
                field_key,
                canonical_json(value),
                _value_type(value),
                provenance.get("source", ""),
                provenance.get("evidence", ""),
                provenance.get("source_ref", ""),
                provenance.get("source_sha256", ""),
                provenance.get("source_observed_at"),
                now,
                actor,
            ),
        )

    def _work_from_conn(
        self,
        conn: sqlite3.Connection,
        work_id: str,
        *,
        include_deleted: bool = False,
    ) -> Dict[str, Any]:
        row = conn.execute("SELECT * FROM work_item WHERE id=?", (work_id,)).fetchone()
        if row is None or (row["deleted_at"] and not include_deleted):
            raise NotFoundError(f"work item not found: {work_id}")
        result = dict(row)
        fields: Dict[str, Any] = {}
        field_meta: Dict[str, Dict[str, Any]] = {}
        for frow in conn.execute(
            "SELECT * FROM work_field WHERE work_id=? ORDER BY field_key", (work_id,)
        ):
            fields[frow["field_key"]] = json.loads(frow["value_json"])
            field_meta[frow["field_key"]] = {
                "value_type": frow["value_type"],
                "source": frow["source"],
                "evidence": frow["evidence"],
                "source_ref": frow["source_ref"],
                "source_sha256": frow["source_sha256"],
                "source_observed_at": frow["source_observed_at"],
                "field_version": frow["field_version"],
                "updated_at": frow["updated_at"],
                "updated_by": frow["updated_by"],
            }
        result["fields"] = fields
        result["field_meta"] = field_meta
        result["shadow_locations"] = [
            dict(r)
            for r in conn.execute(
                """
                SELECT import_id,sheet,business_key,business_key_col,row_number,kind,
                       row_sha256,status,source_ref,created_at
                FROM shadow_import_row WHERE work_id=?
                ORDER BY created_at DESC,id DESC
                """,
                (work_id,),
            )
        ]
        return result

    def _append_event(
        self,
        conn: sqlite3.Connection,
        *,
        work_id: Optional[str],
        aggregate_type: str,
        aggregate_key: str,
        action: str,
        actor: str,
        before: Any,
        after: Any,
        idempotency_key: Optional[str] = None,
        expected_version: Optional[int] = None,
        source: str = "",
        evidence: str = "",
        source_ref: str = "",
        source_sha256: str = "",
        topic: str = "app.change",
    ) -> Tuple[str, int]:
        event_id = self._new_id("evt")
        now = _utcnow()
        cur = conn.execute(
            """
            INSERT INTO change_event(
                event_id,work_id,aggregate_type,aggregate_key,action,actor,request_id,
                idempotency_key,expected_version,before_json,after_json,source,evidence,
                source_ref,source_sha256,created_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                event_id,
                work_id,
                aggregate_type,
                aggregate_key,
                action,
                actor,
                idempotency_key,
                idempotency_key,
                expected_version,
                canonical_json(before) if before is not None else None,
                canonical_json(after) if after is not None else None,
                source,
                evidence,
                source_ref,
                source_sha256,
                now,
            ),
        )
        seq = int(cur.lastrowid)
        payload = {
            "event_id": event_id,
            "event_seq": seq,
            "aggregate_type": aggregate_type,
            "aggregate_key": aggregate_key,
            "action": action,
            "record": after,
            "created_at": now,
        }
        conn.execute(
            """
            INSERT INTO outbox(
                event_id,topic,aggregate_type,aggregate_key,payload_json,status,
                attempts,available_at,created_at
            ) VALUES(?,?,?,?,?,'pending',0,?,?)
            """,
            (
                event_id,
                topic,
                aggregate_type,
                aggregate_key,
                canonical_json(payload),
                now,
                now,
            ),
        )
        return event_id, seq

    def _create_work_tx(
        self,
        conn: sqlite3.Connection,
        *,
        kind: str,
        business_key: str,
        public_id: Optional[str],
        project_no: Optional[str],
        camp_name: Optional[str],
        status: str,
        fields: Mapping[str, Any],
        field_meta: Optional[Mapping[str, Mapping[str, Any]]],
        actor: str,
        source: str,
        evidence: str,
        source_ref: str,
        source_sha256: str,
        source_observed_at: Optional[str],
        action: str,
        idempotency_key: Optional[str],
    ) -> Dict[str, Any]:
        kind = _require_text("kind", kind, 200)
        business_key = _require_text("business_key", business_key, 500)
        identity_field = _management_identity_field(kind)
        identity_value = (
            fields.get(identity_field) or public_id or business_key
            if identity_field
            else None
        )
        if identity_field and identity_value not in (None, ""):
            existing = conn.execute(
                """
                SELECT w.id FROM work_item w
                JOIN work_field f ON f.work_id=w.id
                WHERE w.kind=? AND w.deleted_at IS NULL
                  AND f.field_key=? AND f.value_json=?
                ORDER BY w.updated_at DESC,w.id DESC LIMIT 1
                """,
                (kind, identity_field, canonical_json(identity_value)),
            ).fetchone()
            if existing:
                raise ValidationError(
                    f"duplicate management identity: {kind}/{identity_field}={identity_value}"
                )
        work_id = self._new_id("wrk")
        now = _utcnow()
        try:
            conn.execute(
                """
                INSERT INTO work_item(
                    id,kind,public_id,business_key,project_no,camp_name,status,source,
                    evidence,source_ref,source_sha256,source_observed_at,record_version,
                    created_at,updated_at,updated_by
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,?)
                """,
                (
                    work_id,
                    kind,
                    str(public_id).strip() if public_id else None,
                    business_key,
                    str(project_no).strip() if project_no else None,
                    str(camp_name).strip() if camp_name else None,
                    str(status or ""),
                    str(source or ""),
                    str(evidence or ""),
                    str(source_ref or ""),
                    str(source_sha256 or ""),
                    source_observed_at,
                    now,
                    now,
                    str(actor or ""),
                ),
            )
        except sqlite3.IntegrityError as exc:
            raise ValidationError(
                f"duplicate public_id or kind/business_key: {kind}/{business_key}"
            ) from exc
        defaults = {
            "source": source,
            "evidence": evidence,
            "source_ref": source_ref,
            "source_sha256": source_sha256,
            "source_observed_at": source_observed_at,
        }
        for key, value in sorted(fields.items()):
            self._put_field(
                conn,
                work_id,
                key,
                value,
                actor,
                self._field_provenance(key, field_meta, defaults),
                now,
            )
        created = self._work_from_conn(conn, work_id, include_deleted=True)
        event_id, seq = self._append_event(
            conn,
            work_id=work_id,
            aggregate_type="work_item",
            aggregate_key=work_id,
            action=action,
            actor=actor,
            before=None,
            after=created,
            idempotency_key=idempotency_key,
            source=source,
            evidence=evidence,
            source_ref=source_ref,
            source_sha256=source_sha256,
            topic="work.changed",
        )
        return {"work": created, "event_id": event_id, "event_seq": seq}

    def create_work(
        self,
        *,
        kind: str,
        business_key: str,
        fields: Optional[Mapping[str, Any]] = None,
        public_id: Optional[str] = None,
        project_no: Optional[str] = None,
        camp_name: Optional[str] = None,
        status: str = "",
        field_meta: Optional[Mapping[str, Mapping[str, Any]]] = None,
        actor: str = "app",
        source: str = "app",
        evidence: str = "",
        source_ref: str = "",
        source_sha256: str = "",
        source_observed_at: Optional[str] = None,
        idempotency_key: Optional[str] = None,
    ) -> Dict[str, Any]:
        fields = dict(fields or {})
        payload = {
            "kind": kind,
            "business_key": business_key,
            "public_id": public_id,
            "project_no": project_no,
            "camp_name": camp_name,
            "status": status,
            "fields": fields,
            "field_meta": field_meta or {},
            "actor": actor,
            "source": source,
            "evidence": evidence,
            "source_ref": source_ref,
            "source_sha256": source_sha256,
            "source_observed_at": source_observed_at,
        }
        with self.transaction() as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "work:create", idempotency_key, payload
            )
            if replay is not None:
                return replay
            response = self._create_work_tx(
                conn,
                kind=kind,
                business_key=business_key,
                public_id=public_id,
                project_no=project_no,
                camp_name=camp_name,
                status=status,
                fields=fields,
                field_meta=field_meta,
                actor=actor,
                source=source,
                evidence=evidence,
                source_ref=source_ref,
                source_sha256=source_sha256,
                source_observed_at=source_observed_at,
                action="create",
                idempotency_key=idempotency_key,
            )
            self._save_idempotency(
                conn, "work:create", idempotency_key, request_hash, response
            )
            return response

    def get_work(
        self,
        work_id: Optional[str] = None,
        *,
        kind: Optional[str] = None,
        business_key: Optional[str] = None,
        include_deleted: bool = False,
    ) -> Dict[str, Any]:
        with self.reader() as conn:
            if work_id is None:
                if not kind or not business_key:
                    raise ValidationError("work_id or kind+business_key is required")
                row = conn.execute(
                    "SELECT id FROM work_item WHERE kind=? AND business_key=?",
                    (kind, business_key),
                ).fetchone()
                if row is None:
                    identity_field = _management_identity_field(kind)
                    if identity_field:
                        deleted_clause = "" if include_deleted else " AND w.deleted_at IS NULL"
                        rows = conn.execute(
                            """
                            SELECT w.id FROM work_item w
                            JOIN work_field f ON f.work_id=w.id
                            WHERE w.kind=? AND f.field_key=? AND f.value_json=?
                            """
                            + deleted_clause
                            + " ORDER BY w.deleted_at IS NOT NULL,w.updated_at DESC,w.id DESC LIMIT 2",
                            (kind, identity_field, canonical_json(business_key)),
                        ).fetchall()
                        if len(rows) > 1:
                            raise ValidationError(
                                f"duplicate management identity: "
                                f"{kind}/{identity_field}={business_key}"
                            )
                        row = rows[0] if rows else None
                if row is None:
                    raise NotFoundError(f"work item not found: {kind}/{business_key}")
                work_id = row["id"]
            return self._work_from_conn(
                conn, work_id, include_deleted=include_deleted
            )

    def list_work(
        self,
        *,
        kind: Optional[str] = None,
        status: Optional[str] = None,
        include_deleted: bool = False,
        limit: int = 500,
        offset: int = 0,
    ) -> List[Dict[str, Any]]:
        if limit < 1 or limit > 10_000 or offset < 0:
            raise ValidationError("invalid pagination")
        clauses: List[str] = []
        params: List[Any] = []
        if kind is not None:
            clauses.append("kind=?")
            params.append(kind)
        if status is not None:
            clauses.append("status=?")
            params.append(status)
        if not include_deleted:
            clauses.append("deleted_at IS NULL")
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        with self.reader() as conn:
            rows = conn.execute(
                f"SELECT id FROM work_item{where} ORDER BY updated_at DESC,id LIMIT ? OFFSET ?",
                (*params, limit, offset),
            ).fetchall()
            return [
                self._work_from_conn(conn, row["id"], include_deleted=include_deleted)
                for row in rows
            ]

    def update_work(
        self,
        work_id: str,
        *,
        expected_version: int,
        patch: Mapping[str, Any],
        actor: str = "app",
        source: str = "app",
        evidence: str = "",
        source_ref: str = "",
        source_sha256: str = "",
        source_observed_at: Optional[str] = None,
        idempotency_key: Optional[str] = None,
        _conn: Optional[sqlite3.Connection] = None,
    ) -> Dict[str, Any]:
        if not isinstance(expected_version, int) or expected_version < 1:
            raise ValidationError("expected_version must be a positive integer")
        patch = dict(patch)
        if not patch:
            raise ValidationError("patch must not be empty")
        allowed = {
            "public_id",
            "project_no",
            "camp_name",
            "status",
            "fields",
            "field_meta",
            "delete_fields",
        }
        unknown = sorted(set(patch) - allowed)
        if unknown:
            raise ValidationError(f"unsupported patch keys: {', '.join(unknown)}")
        payload = {
            "work_id": work_id,
            "expected_version": expected_version,
            "patch": patch,
            "actor": actor,
            "source": source,
            "evidence": evidence,
            "source_ref": source_ref,
            "source_sha256": source_sha256,
            "source_observed_at": source_observed_at,
        }
        # 직원 업무센터처럼 "상위 명령 멱등 응답"까지 같은 원자 단위에 묶어야 하는
        # 내부 호출은 이미 BEGIN IMMEDIATE 된 연결을 넘긴다. 이때도 아래의 공개 CRUD
        # 검증·낙관잠금·감사이벤트·outbox·work:update 멱등 처리를 그대로 한 번만 쓴다.
        # 공개 호출은 종전과 같이 이 메서드가 독립 트랜잭션을 소유한다.
        transaction_scope = nullcontext(_conn) if _conn is not None else self.transaction()
        with transaction_scope as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "work:update", idempotency_key, payload
            )
            if replay is not None:
                return replay
            before = self._work_from_conn(conn, work_id)
            if int(before["record_version"]) != expected_version:
                raise VersionConflict(
                    work_id, expected_version, int(before["record_version"])
                )
            now = _utcnow()
            core = {
                key: patch[key]
                for key in ("public_id", "project_no", "camp_name", "status")
                if key in patch
            }
            assignments = [f"{key}=?" for key in core]
            values: List[Any] = [core[key] for key in core]
            assignments.extend(
                [
                    "source=?",
                    "evidence=?",
                    "source_ref=?",
                    "source_sha256=?",
                    "source_observed_at=?",
                    "record_version=record_version+1",
                    "updated_at=?",
                    "updated_by=?",
                ]
            )
            values.extend(
                [
                    source,
                    evidence,
                    source_ref,
                    source_sha256,
                    source_observed_at,
                    now,
                    actor,
                    work_id,
                    expected_version,
                ]
            )
            cur = conn.execute(
                f"UPDATE work_item SET {','.join(assignments)} "
                "WHERE id=? AND record_version=? AND deleted_at IS NULL",
                values,
            )
            if cur.rowcount != 1:
                actual_row = conn.execute(
                    "SELECT record_version FROM work_item WHERE id=?", (work_id,)
                ).fetchone()
                raise VersionConflict(
                    work_id,
                    expected_version,
                    int(actual_row["record_version"]) if actual_row else None,
                )
            defaults = {
                "source": source,
                "evidence": evidence,
                "source_ref": source_ref,
                "source_sha256": source_sha256,
                "source_observed_at": source_observed_at,
            }
            field_meta = patch.get("field_meta") or {}
            for key, value in sorted(dict(patch.get("fields") or {}).items()):
                self._put_field(
                    conn,
                    work_id,
                    key,
                    value,
                    actor,
                    self._field_provenance(key, field_meta, defaults),
                    now,
                )
            delete_fields = list(patch.get("delete_fields") or [])
            if delete_fields:
                conn.executemany(
                    "DELETE FROM work_field WHERE work_id=? AND field_key=?",
                    [(work_id, _require_text("field_key", k, 300)) for k in delete_fields],
                )
            after = self._work_from_conn(conn, work_id)
            event_id, seq = self._append_event(
                conn,
                work_id=work_id,
                aggregate_type="work_item",
                aggregate_key=work_id,
                action="update",
                actor=actor,
                before=before,
                after=after,
                idempotency_key=idempotency_key,
                expected_version=expected_version,
                source=source,
                evidence=evidence,
                source_ref=source_ref,
                source_sha256=source_sha256,
                topic="work.changed",
            )
            response = {"work": after, "event_id": event_id, "event_seq": seq}
            self._save_idempotency(
                conn, "work:update", idempotency_key, request_hash, response
            )
            return response

    def soft_delete_work(
        self,
        work_id: str,
        *,
        expected_version: int,
        actor: str = "app",
        reason: str = "",
        source: str = "app",
        evidence: str = "",
        idempotency_key: Optional[str] = None,
    ) -> Dict[str, Any]:
        if expected_version < 1:
            raise ValidationError("expected_version must be positive")
        payload = {
            "work_id": work_id,
            "expected_version": expected_version,
            "actor": actor,
            "reason": reason,
            "source": source,
            "evidence": evidence,
        }
        with self.transaction() as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "work:delete", idempotency_key, payload
            )
            if replay is not None:
                return replay
            before = self._work_from_conn(conn, work_id)
            if int(before["record_version"]) != expected_version:
                raise VersionConflict(
                    work_id, expected_version, int(before["record_version"])
                )
            now = _utcnow()
            cur = conn.execute(
                """
                UPDATE work_item
                SET deleted_at=?,deleted_by=?,record_version=record_version+1,
                    updated_at=?,updated_by=?,source=?,evidence=?
                WHERE id=? AND record_version=? AND deleted_at IS NULL
                """,
                (
                    now,
                    actor,
                    now,
                    actor,
                    source,
                    evidence or reason,
                    work_id,
                    expected_version,
                ),
            )
            if cur.rowcount != 1:
                actual = conn.execute(
                    "SELECT record_version FROM work_item WHERE id=?", (work_id,)
                ).fetchone()
                raise VersionConflict(
                    work_id,
                    expected_version,
                    int(actual["record_version"]) if actual else None,
                )
            after = self._work_from_conn(conn, work_id, include_deleted=True)
            event_id, seq = self._append_event(
                conn,
                work_id=work_id,
                aggregate_type="work_item",
                aggregate_key=work_id,
                action="soft_delete",
                actor=actor,
                before=before,
                after=after,
                idempotency_key=idempotency_key,
                expected_version=expected_version,
                source=source,
                evidence=evidence or reason,
                topic="work.changed",
            )
            response = {"work": after, "event_id": event_id, "event_seq": seq}
            self._save_idempotency(
                conn, "work:delete", idempotency_key, request_hash, response
            )
            return response

    def restore_work(
        self,
        work_id: str,
        *,
        expected_version: int,
        actor: str = "app",
        reason: str = "",
        source: str = "app",
        evidence: str = "",
        idempotency_key: Optional[str] = None,
    ) -> Dict[str, Any]:
        """소프트 삭제한 업무를 되살린다 — `soft_delete_work` 의 짝.

        ★ 되돌릴 수 없는 것은 만들지 않는다 (2026-08-13 지시). 삭제만 있고 되살리기가
        없으면 잘못 지운 순간 사람이 할 수 있는 일이 없어지고, 그러면 '지워도 되는 것'과
        '지우면 안 되는 것'을 사람이 매번 겁내며 갈라야 한다. 물리 DELETE 를 쓰지 않는
        이유가 바로 여기다 — 지운 것을 되살릴 수 있어야 삭제가 안전한 손잡이가 된다.

        판단을 새로 만들지 않았다: 멱등키·낙관잠금·감사로그 모두 `soft_delete_work` 와
        같은 함수를 그대로 쓴다([162] — 사본이 둘이면 한쪽만 고쳐진다).
        """
        if expected_version < 1:
            raise ValidationError("expected_version must be positive")
        payload = {
            "work_id": work_id,
            "expected_version": expected_version,
            "actor": actor,
            "reason": reason,
            "source": source,
            "evidence": evidence,
        }
        with self.transaction() as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "work:restore", idempotency_key, payload
            )
            if replay is not None:
                return replay
            # 삭제된 행이므로 include_deleted 로 읽는다 — 안 그러면 '없는 업무'로 죽는다.
            before = self._work_from_conn(conn, work_id, include_deleted=True)
            if not before.get("deleted_at"):
                raise ValidationError("삭제되지 않은 업무는 되살릴 수 없습니다")
            if int(before["record_version"]) != expected_version:
                raise VersionConflict(
                    work_id, expected_version, int(before["record_version"])
                )
            now = _utcnow()
            cur = conn.execute(
                """
                UPDATE work_item
                SET deleted_at=NULL,deleted_by=NULL,record_version=record_version+1,
                    updated_at=?,updated_by=?,source=?,evidence=?
                WHERE id=? AND record_version=? AND deleted_at IS NOT NULL
                """,
                (
                    now,
                    actor,
                    source,
                    evidence or reason,
                    work_id,
                    expected_version,
                ),
            )
            if cur.rowcount != 1:
                actual = conn.execute(
                    "SELECT record_version FROM work_item WHERE id=?", (work_id,)
                ).fetchone()
                raise VersionConflict(
                    work_id,
                    expected_version,
                    int(actual["record_version"]) if actual else None,
                )
            after = self._work_from_conn(conn, work_id)
            event_id, seq = self._append_event(
                conn,
                work_id=work_id,
                aggregate_type="work_item",
                aggregate_key=work_id,
                action="restore",
                actor=actor,
                before=before,
                after=after,
                idempotency_key=idempotency_key,
                expected_version=expected_version,
                source=source,
                evidence=evidence or reason,
                topic="work.changed",
            )
            response = {"work": after, "event_id": event_id, "event_seq": seq}
            self._save_idempotency(
                conn, "work:restore", idempotency_key, request_hash, response
            )
            return response

    def get_setting(self, key: str, default: Any = None) -> Dict[str, Any]:
        with self.reader() as conn:
            row = conn.execute("SELECT * FROM app_setting WHERE key=?", (key,)).fetchone()
            if row is None:
                return {"key": key, "value": default, "record_version": 0}
            return {
                "key": row["key"],
                "value": json.loads(row["value_json"]),
                "value_type": row["value_type"],
                "record_version": row["record_version"],
                "updated_at": row["updated_at"],
                "updated_by": row["updated_by"],
            }

    def set_setting(
        self,
        key: str,
        value: Any,
        *,
        expected_version: Optional[int] = None,
        actor: str = "app",
        source: str = "app",
        evidence: str = "",
        source_ref: str = "",
        source_sha256: str = "",
        idempotency_key: Optional[str] = None,
    ) -> Dict[str, Any]:
        key = _require_text("setting key", key, 300)
        payload = {
            "key": key,
            "value": value,
            "expected_version": expected_version,
            "actor": actor,
            "source": source,
            "evidence": evidence,
            "source_ref": source_ref,
            "source_sha256": source_sha256,
        }
        with self.transaction() as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "setting:set", idempotency_key, payload
            )
            if replay is not None:
                return replay
            before_row = conn.execute(
                "SELECT * FROM app_setting WHERE key=?", (key,)
            ).fetchone()
            before = None
            now = _utcnow()
            if before_row is None:
                if expected_version not in (None, 0):
                    raise VersionConflict(key, int(expected_version), None)
                conn.execute(
                    """
                    INSERT INTO app_setting(
                        key,value_json,value_type,record_version,updated_at,updated_by
                    ) VALUES(?,?,?,1,?,?)
                    """,
                    (key, canonical_json(value), _value_type(value), now, actor),
                )
            else:
                before = {
                    "key": key,
                    "value": json.loads(before_row["value_json"]),
                    "record_version": before_row["record_version"],
                }
                actual = int(before_row["record_version"])
                if expected_version is None:
                    raise ValidationError(
                        "expected_version is required when updating an existing setting"
                    )
                if expected_version != actual:
                    raise VersionConflict(key, int(expected_version), actual)
                conn.execute(
                    """
                    UPDATE app_setting
                    SET value_json=?,value_type=?,record_version=record_version+1,
                        updated_at=?,updated_by=?
                    WHERE key=? AND record_version=?
                    """,
                    (
                        canonical_json(value),
                        _value_type(value),
                        now,
                        actor,
                        key,
                        expected_version,
                    ),
                )
            after_row = conn.execute(
                "SELECT * FROM app_setting WHERE key=?", (key,)
            ).fetchone()
            after = {
                "key": key,
                "value": json.loads(after_row["value_json"]),
                "value_type": after_row["value_type"],
                "record_version": after_row["record_version"],
                "updated_at": after_row["updated_at"],
                "updated_by": after_row["updated_by"],
            }
            event_id, seq = self._append_event(
                conn,
                work_id=None,
                aggregate_type="app_setting",
                aggregate_key=key,
                action="create" if before is None else "update",
                actor=actor,
                before=before,
                after=after,
                idempotency_key=idempotency_key,
                expected_version=expected_version,
                source=source,
                evidence=evidence,
                source_ref=source_ref,
                source_sha256=source_sha256,
                topic="setting.changed",
            )
            response = {"setting": after, "event_id": event_id, "event_seq": seq}
            self._save_idempotency(
                conn, "setting:set", idempotency_key, request_hash, response
            )
            return response

    def _record_import_conflict(
        self,
        conn: sqlite3.Connection,
        *,
        import_id: str,
        sheet: str,
        business_key: str,
        row_number: int,
        field_key: str,
        incoming: Any,
        current: Any,
        reason: str,
    ) -> int:
        conn.execute(
            """
            INSERT OR IGNORE INTO import_conflict(
                import_id,sheet,business_key,row_number,field_key,incoming_json,
                current_json,reason,status,created_at
            ) VALUES(?,?,?,?,?,?,?,?,'open',?)
            """,
            (
                import_id,
                sheet,
                business_key,
                row_number,
                field_key,
                canonical_json(incoming),
                canonical_json(current),
                reason,
                _utcnow(),
            ),
        )
        row = conn.execute(
            """
            SELECT id FROM import_conflict
            WHERE import_id=? AND sheet=? AND row_number=? AND field_key=? AND reason=?
            """,
            (import_id, sheet, row_number, field_key, reason),
        ).fetchone()
        return int(row["id"])

    def shadow_import(
        self,
        *,
        import_id: str,
        sheet: str,
        business_key: str,
        row_number: int,
        kind: str,
        fields: Mapping[str, Any],
        business_key_col: str = "프로젝트NO",
        public_id: Optional[str] = None,
        project_no: Optional[str] = None,
        camp_name: Optional[str] = None,
        status: str = "",
        source_file: str = "",
        source_sha256: str = "",
        observed_at: Optional[str] = None,
        evidence: str = "Excel shadow import",
        apply_if_missing: bool = False,
        actor: str = "shadow-import",
        idempotency_key: Optional[str] = None,
        _conn: Optional[sqlite3.Connection] = None,
    ) -> Dict[str, Any]:
        """Inspect one Excel row without overwriting canonical application data.

        The exact ``sheet + business_key + row_number`` locator is retained as
        evidence.  Existing records are only compared; differences become
        ``import_conflict`` rows.  ``apply_if_missing`` may create a new record,
        but it never overwrites an existing record.
        """

        import_id = _require_text("import_id", import_id, 300)
        sheet = _require_text("sheet", sheet, 300)
        business_key = _require_text("business_key", business_key, 500)
        business_key_col = _require_text("business_key_col", business_key_col, 300)
        kind = _require_text("kind", kind, 200)
        if not isinstance(row_number, int) or row_number < 1:
            raise ValidationError("row_number must be a positive integer")
        fields = dict(fields)
        source_ref = f"{source_file or 'xlsx'}::{sheet}!row={row_number};key={business_key}"
        incoming = {
            "kind": kind,
            "business_key": business_key,
            "business_key_col": business_key_col,
            "public_id": public_id,
            "project_no": project_no,
            "camp_name": camp_name,
            "status": status,
            "fields": fields,
            "source_file": source_file,
            "source_sha256": source_sha256,
            "observed_at": observed_at,
        }
        row_hash = sha256_json(incoming)
        payload = {
            "import_id": import_id,
            "sheet": sheet,
            "business_key": business_key,
            "row_number": row_number,
            "incoming": incoming,
            "apply_if_missing": apply_if_missing,
            "actor": actor,
        }
        # 대량 초기 이관은 호출자가 연 단일 트랜잭션을 공유한다. 평상시 공개 호출은
        # 이전과 같이 행 하나가 독립 트랜잭션이다. `_conn`은 내부 컷오버 전용이라
        # 실시간 앱 입력의 낙관잠금·원자성 계약을 바꾸지 않는다.
        transaction_scope = nullcontext(_conn) if _conn is not None else self.transaction()
        with transaction_scope as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "shadow:import", idempotency_key, payload
            )
            if replay is not None:
                return replay
            prior = conn.execute(
                """
                SELECT * FROM shadow_import_row
                WHERE import_id=? AND sheet=? AND row_number=?
                """,
                (import_id, sheet, row_number),
            ).fetchone()
            if prior is not None:
                if prior["row_sha256"] == row_hash:
                    response = {
                        "status": prior["status"],
                        "work_id": prior["work_id"],
                        "row_sha256": row_hash,
                        "conflict_ids": [],
                        "shadow_replay": True,
                    }
                    self._save_idempotency(
                        conn, "shadow:import", idempotency_key, request_hash, response
                    )
                    return response
                conflict_id = self._record_import_conflict(
                    conn,
                    import_id=import_id,
                    sheet=sheet,
                    business_key=business_key,
                    row_number=row_number,
                    field_key="__row__",
                    incoming=incoming,
                    current={"row_sha256": prior["row_sha256"]},
                    reason="same_import_locator_changed",
                )
                response = {
                    "status": "conflict",
                    "work_id": prior["work_id"],
                    "row_sha256": row_hash,
                    "conflict_ids": [conflict_id],
                }
                self._save_idempotency(
                    conn, "shadow:import", idempotency_key, request_hash, response
                )
                return response

            existing_row = conn.execute(
                "SELECT id FROM work_item WHERE kind=? AND business_key=?",
                (kind, business_key),
            ).fetchone()
            work_id: Optional[str] = existing_row["id"] if existing_row else None
            conflict_ids: List[int] = []
            # v586 changed sheets 15/16 from the settlement relation ID to a
            # dedicated management ID as the row identity.  Records imported
            # before that change can therefore still have ``business_key`` set
            # to the settlement ID while already carrying the new management
            # ID in ``work_field``.  Resolve that legacy identity before the
            # create path; otherwise ``_create_work_tx`` correctly detects the
            # duplicate field value and rolls the whole shadow-import batch
            # back.  Never choose arbitrarily when old data contains duplicate
            # management IDs: retain a normal import conflict instead.
            identity_field = _management_identity_field(kind)
            if work_id is None and identity_field:
                identity_rows = conn.execute(
                    """
                    SELECT w.id FROM work_item w
                    JOIN work_field f ON f.work_id=w.id
                    WHERE w.kind=? AND w.deleted_at IS NULL
                      AND f.field_key=? AND f.value_json=?
                    ORDER BY w.updated_at DESC,w.id DESC
                    """,
                    (kind, identity_field, canonical_json(business_key)),
                ).fetchall()
                if len(identity_rows) == 1:
                    work_id = str(identity_rows[0]["id"])
                elif len(identity_rows) > 1:
                    conflict_ids.append(
                        self._record_import_conflict(
                            conn,
                            import_id=import_id,
                            sheet=sheet,
                            business_key=business_key,
                            row_number=row_number,
                            field_key=identity_field,
                            incoming=business_key,
                            current={
                                "work_ids": [str(row["id"]) for row in identity_rows],
                                "count": len(identity_rows),
                            },
                            reason="duplicate_management_identity",
                        )
                    )
            event_id: Optional[str] = None
            event_seq: Optional[int] = None
            if conflict_ids:
                shadow_status = "conflict"
            elif work_id is None:
                if apply_if_missing:
                    created = self._create_work_tx(
                        conn,
                        kind=kind,
                        business_key=business_key,
                        public_id=public_id,
                        project_no=project_no,
                        camp_name=camp_name,
                        status=status,
                        fields=fields,
                        field_meta=None,
                        actor=actor,
                        source="excel-shadow",
                        evidence=evidence,
                        source_ref=source_ref,
                        source_sha256=source_sha256,
                        source_observed_at=observed_at,
                        action="shadow_import_create",
                        idempotency_key=idempotency_key,
                    )
                    work_id = created["work"]["id"]
                    event_id = created["event_id"]
                    event_seq = created["event_seq"]
                    shadow_status = "created"
                else:
                    shadow_status = "candidate"
            else:
                current = self._work_from_conn(conn, work_id, include_deleted=True)
                if current.get("deleted_at"):
                    conflict_ids.append(
                        self._record_import_conflict(
                            conn,
                            import_id=import_id,
                            sheet=sheet,
                            business_key=business_key,
                            row_number=row_number,
                            field_key="__record__",
                            incoming=incoming,
                            current={"deleted_at": current["deleted_at"]},
                            reason="canonical_record_soft_deleted",
                        )
                    )
                for key, incoming_value in (
                    ("public_id", public_id),
                    ("project_no", project_no),
                    ("camp_name", camp_name),
                    ("status", status),
                ):
                    if incoming_value in (None, ""):
                        continue
                    if canonical_json(current.get(key)) != canonical_json(incoming_value):
                        conflict_ids.append(
                            self._record_import_conflict(
                                conn,
                                import_id=import_id,
                                sheet=sheet,
                                business_key=business_key,
                                row_number=row_number,
                                field_key=key,
                                incoming=incoming_value,
                                current=current.get(key),
                                reason="canonical_value_differs",
                            )
                        )
                for key, incoming_value in sorted(fields.items()):
                    current_value = current["fields"].get(key)
                    if key not in current["fields"] or canonical_json(current_value) != canonical_json(
                        incoming_value
                    ):
                        conflict_ids.append(
                            self._record_import_conflict(
                                conn,
                                import_id=import_id,
                                sheet=sheet,
                                business_key=business_key,
                                row_number=row_number,
                                field_key=key,
                                incoming=incoming_value,
                                current=current_value,
                                reason="canonical_value_differs",
                            )
                        )
                shadow_status = "conflict" if conflict_ids else "matched"

            conn.execute(
                """
                INSERT INTO shadow_import_row(
                    import_id,sheet,business_key,business_key_col,row_number,kind,
                    work_id,row_sha256,status,source_ref,created_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    import_id,
                    sheet,
                    business_key,
                    business_key_col,
                    row_number,
                    kind,
                    work_id,
                    row_hash,
                    shadow_status,
                    source_ref,
                    _utcnow(),
                ),
            )
            response = {
                "status": shadow_status,
                "work_id": work_id,
                "row_sha256": row_hash,
                "conflict_ids": conflict_ids,
                "event_id": event_id,
                "event_seq": event_seq,
            }
            self._save_idempotency(
                conn, "shadow:import", idempotency_key, request_hash, response
            )
            return response

    def list_import_conflicts(
        self, *, import_id: Optional[str] = None, status: str = "open"
    ) -> List[Dict[str, Any]]:
        clauses = ["status=?"]
        params: List[Any] = [status]
        if import_id:
            clauses.append("import_id=?")
            params.append(import_id)
        with self.reader() as conn:
            rows = conn.execute(
                "SELECT * FROM import_conflict WHERE "
                + " AND ".join(clauses)
                + " ORDER BY id",
                params,
            ).fetchall()
            result = []
            for row in rows:
                item = dict(row)
                item["incoming"] = (
                    json.loads(item.pop("incoming_json"))
                    if item.get("incoming_json") is not None
                    else None
                )
                item["current"] = (
                    json.loads(item.pop("current_json"))
                    if item.get("current_json") is not None
                    else None
                )
                result.append(item)
            return result

    def lease_outbox(
        self, *, limit: int = 100, lease_seconds: int = 300
    ) -> Tuple[str, List[Dict[str, Any]]]:
        if limit < 1 or limit > 1_000 or lease_seconds < 1:
            raise ValidationError("invalid outbox lease")
        now = datetime.now(timezone.utc)
        now_text = now.isoformat(timespec="microseconds")
        expired = (now - timedelta(seconds=lease_seconds)).isoformat(
            timespec="microseconds"
        )
        token = self._new_id("lease")
        with self.transaction() as conn:
            conn.execute(
                """
                UPDATE outbox SET status='pending',lease_token=NULL,leased_at=NULL,
                    last_error=CASE WHEN last_error IS NULL OR last_error=''
                                    THEN 'lease expired' ELSE last_error END
                WHERE status='leased' AND leased_at<?
                """,
                (expired,),
            )
            rows = conn.execute(
                """
                SELECT id FROM outbox
                WHERE status IN ('pending','failed') AND available_at<=?
                ORDER BY id LIMIT ?
                """,
                (now_text, limit),
            ).fetchall()
            ids = [int(row["id"]) for row in rows]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                conn.execute(
                    f"""
                    UPDATE outbox SET status='leased',lease_token=?,leased_at=?,
                        attempts=attempts+1
                    WHERE id IN ({placeholders})
                    """,
                    (token, now_text, *ids),
                )
                leased = conn.execute(
                    f"SELECT * FROM outbox WHERE id IN ({placeholders}) ORDER BY id",
                    ids,
                ).fetchall()
            else:
                leased = []
            items: List[Dict[str, Any]] = []
            for row in leased:
                item = dict(row)
                item["payload"] = json.loads(item.pop("payload_json"))
                items.append(item)
            return token, items

    def ack_outbox(self, lease_token: str, ids: Sequence[int]) -> int:
        if not ids:
            return 0
        now = _utcnow()
        with self.transaction() as conn:
            placeholders = ",".join("?" for _ in ids)
            cur = conn.execute(
                f"""
                UPDATE outbox SET status='done',finished_at=?,lease_token=NULL,
                    leased_at=NULL,last_error=NULL
                WHERE lease_token=? AND status='leased' AND id IN ({placeholders})
                """,
                (now, lease_token, *ids),
            )
            return int(cur.rowcount)

    def fail_outbox(
        self,
        lease_token: str,
        ids: Sequence[int],
        error: str,
        *,
        retry_after_seconds: int = 60,
    ) -> int:
        if not ids:
            return 0
        available = (
            datetime.now(timezone.utc) + timedelta(seconds=max(0, retry_after_seconds))
        ).isoformat(timespec="microseconds")
        with self.transaction() as conn:
            placeholders = ",".join("?" for _ in ids)
            cur = conn.execute(
                f"""
                UPDATE outbox SET status='failed',available_at=?,lease_token=NULL,
                    leased_at=NULL,last_error=?
                WHERE lease_token=? AND status='leased' AND id IN ({placeholders})
                """,
                (available, str(error)[:4_000], lease_token, *ids),
            )
            return int(cur.rowcount)

    @staticmethod
    def _sheet_spec(sheet: str) -> Dict[str, str]:
        sheet = _require_text("sheet", sheet, 300)
        return dict(
            SHEET_SPECS.get(
                sheet,
                {
                    "kind": f"sheet:{sheet}",
                    "public_id": "",
                    "project_no": "프로젝트NO",
                    "camp_name": "캠프명",
                    "status": "진행상태",
                },
            )
        )

    @staticmethod
    def _legacy_value(item: Mapping[str, Any]) -> Any:
        value = item.get("value")
        vtype = str(item.get("vtype") or "").lower()
        if value in (None, ""):
            return value
        if vtype == "number" and isinstance(value, str):
            cleaned = value.replace(",", "").strip()
            number = float(cleaned)
            return int(number) if number.is_integer() else number
        if vtype == "bool" and isinstance(value, str):
            return value.strip().lower() in {"1", "true", "yes", "y", "예"}
        if vtype == "json" and isinstance(value, str):
            return json.loads(value)
        return value

    @staticmethod
    def _legacy_empty(value: Any) -> bool:
        return value is None or (isinstance(value, str) and value.strip() in {"", "-"})

    def _find_legacy_work_id(
        self, *, sheet: str, key_col: str, key: str, kind: str
    ) -> Optional[str]:
        encoded_key = canonical_json(key)
        spec = self._sheet_spec(sheet)
        with self.reader() as conn:
            row = conn.execute(
                """
                SELECT work_id FROM shadow_import_row
                WHERE sheet=? AND business_key_col=? AND business_key=?
                  AND work_id IS NOT NULL
                ORDER BY id DESC LIMIT 1
                """,
                (sheet, key_col, key),
            ).fetchone()
            if row:
                return str(row["work_id"])
            row = conn.execute(
                "SELECT id FROM work_item WHERE kind=? AND public_id=? "
                "ORDER BY deleted_at IS NOT NULL,updated_at DESC LIMIT 1",
                (kind, key),
            ).fetchone()
            if row:
                return str(row["id"])
            identity_col = str(spec.get("business_key") or "")
            if identity_col and key_col == identity_col:
                # During the v586 transition existing 15/16 records can still
                # have the old 정산ID business_key while already carrying the
                # new 관리ID in work_field.  Match that field so an app edit
                # updates the existing record instead of creating a duplicate,
                # but reject duplicate management IDs rather than choosing one.
                rows = conn.execute(
                    """
                    SELECT w.id FROM work_item w
                    JOIN work_field f ON f.work_id=w.id
                    WHERE w.kind=? AND w.deleted_at IS NULL
                      AND f.field_key=? AND f.value_json=?
                    ORDER BY w.updated_at DESC,w.id DESC LIMIT 2
                    """,
                    (kind, identity_col, encoded_key),
                ).fetchall()
                if len(rows) > 1:
                    raise ValidationError(
                        f"{sheet}: 식별키 {identity_col}={key}가 여러 건입니다; "
                        "중복 관리ID를 먼저 정리하세요"
                    )
                if rows:
                    return str(rows[0]["id"])
            relation_col = str(spec.get("relation_id") or "")
            if relation_col and key_col == relation_col:
                # 정산ID는 15·16시트의 부모 관계키이지 행 식별자가 아니다.
                # 한 정산에 계산서/입금이 복수일 수 있으므로 관계키가 여러 행에
                # 걸리면 최신 한 건을 임의 선택하지 말고 식별 관리ID를 요구한다.
                rows = conn.execute(
                    """
                    SELECT w.id FROM work_item w
                    JOIN work_field f ON f.work_id=w.id
                    WHERE w.kind=? AND w.deleted_at IS NULL
                      AND f.field_key=? AND f.value_json=?
                    ORDER BY w.updated_at DESC,w.id DESC LIMIT 2
                    """,
                    (kind, relation_col, encoded_key),
                ).fetchall()
                if len(rows) > 1:
                    raise ValidationError(
                        f"{sheet}: 관계키 {relation_col}={key}가 여러 건입니다; "
                        f"{spec.get('public_id') or '관리ID'}로 지정하세요"
                    )
                return str(rows[0]["id"]) if rows else None
            row = conn.execute(
                "SELECT id FROM work_item WHERE kind=? AND business_key=? "
                "ORDER BY deleted_at IS NOT NULL,updated_at DESC LIMIT 1",
                (kind, key),
            ).fetchone()
            if row:
                return str(row["id"])
            if key_col == "프로젝트NO":
                row = conn.execute(
                    "SELECT id FROM work_item WHERE kind=? AND project_no=? "
                    "ORDER BY deleted_at IS NOT NULL,updated_at DESC LIMIT 1",
                    (kind, key),
                ).fetchone()
                if row:
                    return str(row["id"])
            row = conn.execute(
                """
                SELECT w.id FROM work_item w
                JOIN work_field f ON f.work_id=w.id
                WHERE w.kind=? AND f.field_key=? AND f.value_json=?
                ORDER BY w.deleted_at IS NOT NULL,w.updated_at DESC LIMIT 1
                """,
                (kind, key_col, encoded_key),
            ).fetchone()
            return str(row["id"]) if row else None

    @staticmethod
    def _legacy_current_value(
        work: Mapping[str, Any], spec: Mapping[str, str], column: str
    ) -> Any:
        for core_name in ("public_id", "project_no", "camp_name", "status"):
            if column == spec.get(core_name):
                return work.get(core_name)
        return (work.get("fields") or {}).get(column)

    def _apply_legacy_row_group(
        self,
        group: Sequence[Tuple[int, Mapping[str, Any]]],
        *,
        sheet: str,
        row_number: int,
        source: str,
        actor: str,
        token: str,
    ) -> Dict[str, Any]:
        """Turn coordinate-mode cells for one Excel row into one work record."""

        spec = self._sheet_spec(sheet)
        incoming: Dict[str, Any] = {}
        policies: Dict[str, bool] = {}
        evidences: List[str] = []
        for _index, raw in group:
            column = _require_text("col", raw.get("col"), 300)
            incoming[column] = self._legacy_value(raw)
            policies[column] = bool(raw.get("only_if_empty", True))
            evidence = str(raw.get("evidence") or "").strip()
            if evidence and evidence not in evidences:
                evidences.append(evidence)
        key_candidates = [
            spec.get("business_key") or "",
            spec.get("public_id") or "",
            "프로젝트NO",
            "접수ID",
            "점검ID",
            "정산ID",
            "작업ID",
            "업무ID",
            "PO관리ID",
        ]
        key_col = next(
            (
                column
                for column in key_candidates
                if column and not self._legacy_empty(incoming.get(column))
            ),
            "",
        )
        if not key_col:
            raise ValidationError("coordinate row group has no stable business key")
        key = str(incoming[key_col]).strip()
        kind = spec["kind"]
        evidence = " | ".join(evidences)[:4_000]
        source_ref = f"legacy-row:{sheet}!row={row_number};{key_col}={key}"
        work_id = self._find_legacy_work_id(
            sheet=sheet, key_col=key_col, key=key, kind=kind
        )
        if work_id is None:
            required_key_col = str(spec.get("business_key") or "")
            if required_key_col and key_col != required_key_col:
                raise ValidationError(
                    f"{sheet}: 새 {kind} 건은 {required_key_col}가 필요합니다; "
                    f"{key_col}는 관계/참조키로만 사용할 수 있습니다"
                )
            try:
                response = self.create_work(
                    kind=kind,
                    business_key=key,
                    fields=incoming,
                    public_id=(key if key_col == spec.get("public_id") else None),
                    project_no=(
                        str(incoming.get(spec.get("project_no") or "") or "") or None
                    ),
                    camp_name=(
                        str(incoming.get(spec.get("camp_name") or "") or "") or None
                    ),
                    status=str(incoming.get(spec.get("status") or "") or ""),
                    actor=actor,
                    source=source,
                    evidence=evidence,
                    source_ref=source_ref,
                    idempotency_key=f"{token}:create",
                )
                return {
                    "action": "created",
                    "work_id": response["work"]["id"],
                    "event_id": response["event_id"],
                    "row_number": row_number,
                    "business_key": key,
                    "fields": sorted(incoming),
                }
            except ValidationError:
                work_id = self._find_legacy_work_id(
                    sheet=sheet, key_col=key_col, key=key, kind=kind
                )
                if work_id is None:
                    raise

        for attempt in range(5):
            current = self.get_work(work_id)
            changed: Dict[str, Any] = {}
            blocked: List[str] = []
            for column, value in incoming.items():
                existing = self._legacy_current_value(current, spec, column)
                if canonical_json(existing) == canonical_json(value):
                    continue
                if policies.get(column, True) and not self._legacy_empty(existing):
                    blocked.append(column)
                    continue
                changed[column] = value
            if not changed:
                return {
                    "action": "skipped",
                    "work_id": work_id,
                    "reason": "already_equal_or_only_if_empty",
                    "row_number": row_number,
                    "business_key": key,
                    "blocked_fields": sorted(blocked),
                }
            patch: Dict[str, Any] = {"fields": changed}
            for core_name in ("public_id", "project_no", "camp_name", "status"):
                column = spec.get(core_name) or ""
                if column in changed:
                    patch[core_name] = changed[column]
            version = int(current["record_version"])
            try:
                response = self.update_work(
                    work_id,
                    expected_version=version,
                    patch=patch,
                    actor=actor,
                    source=source,
                    evidence=evidence,
                    source_ref=source_ref,
                    idempotency_key=f"{token}:v{version}",
                )
                return {
                    "action": "updated",
                    "work_id": work_id,
                    "event_id": response["event_id"],
                    "row_number": row_number,
                    "business_key": key,
                    "fields": sorted(changed),
                    "blocked_fields": sorted(blocked),
                }
            except VersionConflict:
                if attempt == 4:
                    raise
        raise StoreError("coordinate row optimistic retry exhausted")

    def apply_legacy_items(
        self,
        items: Iterable[Mapping[str, Any]],
        source: str,
        idempotency_key: Optional[str] = None,
        actor: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply legacy ledger queue items to the canonical DB, never to Excel.

        Keyed items resolve ``sheet + key_col + key`` to a work record.  Existing
        records use optimistic locking with bounded retry; missing records are
        created.  Coordinate (``cell``) items become audited settings under a
        stable ``excel-cell:<sheet>:<cell>`` key.  Repeating an identical item is
        harmless even when a batch-level idempotency key is omitted.
        """

        normalized = [dict(item) for item in items]
        source = _require_text("source", source, 300)
        actor = str(actor or source)
        batch_payload = {"items": normalized, "source": source, "actor": actor}
        if idempotency_key:
            with self.transaction() as conn:
                replay, _request_hash = self._idempotency_replay(
                    conn, "legacy:batch", idempotency_key, batch_payload
                )
                if replay is not None:
                    return replay

        result: Dict[str, Any] = {
            "ok": True,
            "accepted": len(normalized),
            "created": 0,
            "updated": 0,
            "settings": 0,
            "skipped": 0,
            "errors": [],
            "items": [],
        }
        batch_token = idempotency_key or sha256_json(
            {"source": source, "items": normalized}
        )
        coordinate_groups: Dict[Tuple[str, int], List[Tuple[int, Mapping[str, Any]]]] = {}
        for position, candidate in enumerate(normalized):
            cell = str(candidate.get("cell") or "").strip().upper()
            if not cell or not str(candidate.get("col") or "").strip():
                continue
            if str(candidate.get("key_col") or "").strip() not in {"", "-"}:
                continue
            match = re.fullmatch(r"\$?[A-Z]{1,3}\$?(\d+)", cell)
            sheet = str(candidate.get("sheet") or "").strip()
            if match and sheet:
                coordinate_groups.setdefault((sheet, int(match.group(1))), []).append(
                    (position, candidate)
                )
        group_leaders: Dict[int, Tuple[str, int, Sequence[Tuple[int, Mapping[str, Any]]]]] = {}
        grouped_members: set[int] = set()
        for (sheet, row_number), group in coordinate_groups.items():
            spec = self._sheet_spec(sheet)
            stable_columns = {
                spec.get("business_key") or "",
                spec.get("public_id") or "",
                "프로젝트NO",
                "접수ID",
                "점검ID",
                "정산ID",
                "작업ID",
                "업무ID",
                "PO관리ID",
            }
            if not any(
                str(raw.get("col") or "") in stable_columns
                and not self._legacy_empty(self._legacy_value(raw))
                for _position, raw in group
            ):
                continue
            leader = min(position for position, _raw in group)
            group_leaders[leader] = (sheet, row_number, group)
            grouped_members.update(position for position, _raw in group)
        for index, raw in enumerate(normalized):
            if index in grouped_members and index not in group_leaders:
                continue
            item_result: Dict[str, Any] = {"index": index}
            try:
                sheet = _require_text("sheet", raw.get("sheet"), 300)
                value = self._legacy_value(raw)
                evidence = str(raw.get("evidence") or "")
                only_if_empty = bool(raw.get("only_if_empty", True))
                item_hash = sha256_json(raw)
                item_token = f"legacy:{batch_token}:{index}:{item_hash[:20]}"
                if index in group_leaders:
                    grouped_sheet, row_number, group = group_leaders[index]
                    group_token = (
                        f"legacy:{batch_token}:row:{grouped_sheet}:{row_number}:"
                        f"{sha256_json([dict(x[1]) for x in group])[:20]}"
                    )
                    grouped = self._apply_legacy_row_group(
                        group,
                        sheet=grouped_sheet,
                        row_number=row_number,
                        source=source,
                        actor=actor,
                        token=group_token,
                    )
                    item_result.update(grouped)
                    item_result["indexes"] = [position for position, _item in group]
                    action = grouped["action"]
                    if action == "created":
                        result["created"] += 1
                    elif action == "updated":
                        result["updated"] += 1
                    else:
                        result["skipped"] += 1
                    result["items"].append(item_result)
                    continue
                cell = str(raw.get("cell") or "").strip().upper()
                key_col = str(raw.get("key_col") or "").strip()
                key = str(raw.get("key") or "").strip()
                col = str(raw.get("col") or "").strip()

                if cell or key_col == "-":
                    stable_cell = cell or key
                    stable_cell = _require_text("cell", stable_cell, 50).upper()
                    setting_key = f"excel-cell:{sheet}:{stable_cell}"
                    for attempt in range(5):
                        current = self.get_setting(setting_key)
                        current_value = current.get("value")
                        if canonical_json(current_value) == canonical_json(value):
                            item_result.update(
                                {"action": "skipped", "reason": "already_equal"}
                            )
                            result["skipped"] += 1
                            break
                        if int(current.get("record_version") or 0) and only_if_empty and not self._legacy_empty(
                            current_value
                        ):
                            item_result.update(
                                {"action": "skipped", "reason": "only_if_empty"}
                            )
                            result["skipped"] += 1
                            break
                        version = int(current.get("record_version") or 0)
                        try:
                            response = self.set_setting(
                                setting_key,
                                value,
                                expected_version=version,
                                actor=actor,
                                source=source,
                                evidence=evidence,
                                source_ref=f"{sheet}!{stable_cell}",
                                idempotency_key=f"{item_token}:v{version}",
                            )
                            item_result.update(
                                {
                                    "action": "setting",
                                    "setting_key": setting_key,
                                    "event_id": response["event_id"],
                                }
                            )
                            result["settings"] += 1
                            break
                        except VersionConflict:
                            if attempt == 4:
                                raise
                    result["items"].append(item_result)
                    continue

                key_col = _require_text("key_col", key_col, 300)
                key = _require_text("key", key, 500)
                col = _require_text("col", col, 300)
                spec = self._sheet_spec(sheet)
                kind = spec["kind"]
                work_id = self._find_legacy_work_id(
                    sheet=sheet, key_col=key_col, key=key, kind=kind
                )
                if work_id is None:
                    required_key_col = str(spec.get("business_key") or "")
                    if required_key_col and key_col != required_key_col:
                        raise ValidationError(
                            f"{sheet}: 새 {kind} 건은 {required_key_col}가 필요합니다; "
                            f"{key_col}는 관계/참조키로만 사용할 수 있습니다"
                        )
                    fields = {key_col: key, col: value}
                    public_id = key if key_col == spec.get("public_id") else None
                    project_no = key if key_col == spec.get("project_no") else None
                    camp_name = str(value) if col == spec.get("camp_name") else None
                    status_value = str(value) if col == spec.get("status") else ""
                    try:
                        response = self.create_work(
                            kind=kind,
                            business_key=key,
                            fields=fields,
                            public_id=public_id,
                            project_no=project_no,
                            camp_name=camp_name,
                            status=status_value,
                            actor=actor,
                            source=source,
                            evidence=evidence,
                            source_ref=f"legacy:{sheet}:{key_col}={key}",
                            idempotency_key=f"{item_token}:create",
                        )
                        work_id = response["work"]["id"]
                        item_result.update(
                            {
                                "action": "created",
                                "work_id": work_id,
                                "event_id": response["event_id"],
                            }
                        )
                        result["created"] += 1
                        result["items"].append(item_result)
                        continue
                    except ValidationError:
                        # A concurrent creator may have won after our lookup.
                        work_id = self._find_legacy_work_id(
                            sheet=sheet, key_col=key_col, key=key, kind=kind
                        )
                        if work_id is None:
                            raise

                for attempt in range(5):
                    current = self.get_work(work_id)
                    current_value = current.get("fields", {}).get(col)
                    if col == spec.get("status"):
                        current_value = current.get("status")
                    elif col == spec.get("project_no"):
                        current_value = current.get("project_no")
                    elif col == spec.get("camp_name"):
                        current_value = current.get("camp_name")
                    elif col == spec.get("public_id"):
                        current_value = current.get("public_id")
                    if canonical_json(current_value) == canonical_json(value):
                        item_result.update(
                            {
                                "action": "skipped",
                                "work_id": work_id,
                                "reason": "already_equal",
                            }
                        )
                        result["skipped"] += 1
                        break
                    if only_if_empty and not self._legacy_empty(current_value):
                        item_result.update(
                            {
                                "action": "skipped",
                                "work_id": work_id,
                                "reason": "only_if_empty",
                            }
                        )
                        result["skipped"] += 1
                        break
                    patch: Dict[str, Any] = {"fields": {col: value}}
                    for core_name in ("public_id", "project_no", "camp_name", "status"):
                        if col == spec.get(core_name):
                            patch[core_name] = value
                    version = int(current["record_version"])
                    try:
                        response = self.update_work(
                            work_id,
                            expected_version=version,
                            patch=patch,
                            actor=actor,
                            source=source,
                            evidence=evidence,
                            source_ref=f"legacy:{sheet}:{key_col}={key}",
                            idempotency_key=f"{item_token}:v{version}",
                        )
                        item_result.update(
                            {
                                "action": "updated",
                                "work_id": work_id,
                                "event_id": response["event_id"],
                            }
                        )
                        result["updated"] += 1
                        break
                    except VersionConflict:
                        if attempt == 4:
                            raise
                result["items"].append(item_result)
            except Exception as exc:
                result["ok"] = False
                error = {
                    "index": index,
                    "type": type(exc).__name__,
                    "message": str(exc)[:500],
                }
                result["errors"].append(error)
                result["items"].append({**item_result, "action": "error", **error})

        state = self.status()
        result["change_seq"] = state["change_seq"]
        result["outbox_pending"] = state["outbox_pending"]
        if idempotency_key:
            with self.transaction() as conn:
                replay, request_hash = self._idempotency_replay(
                    conn, "legacy:batch", idempotency_key, batch_payload
                )
                if replay is not None:
                    return replay
                self._save_idempotency(
                    conn,
                    "legacy:batch",
                    idempotency_key,
                    request_hash,
                    result,
                )
        return result

    def list_sheet_rows(
        self, sheet: str, *, include_deleted: bool = False
    ) -> List[Dict[str, Any]]:
        """Render canonical records as Excel-shaped dictionaries for the UI."""

        spec = self._sheet_spec(sheet)
        works = self.list_work(
            kind=spec["kind"], include_deleted=include_deleted, limit=10_000
        )
        rendered: List[Tuple[int, str, Dict[str, Any]]] = []
        for work in works:
            locations = [
                loc
                for loc in (work.get("shadow_locations") or [])
                if loc.get("sheet") == sheet
            ]
            location = locations[0] if locations else None
            row = dict(work.get("fields") or {})
            for core_name in ("public_id", "project_no", "camp_name", "status"):
                column = spec.get(core_name) or ""
                value = work.get(core_name)
                if column and value not in (None, ""):
                    row[column] = value
            if location:
                row[location.get("business_key_col") or "프로젝트NO"] = work[
                    "business_key"
                ]
            elif not any(
                str(row.get(col) or "")
                for col in (
                    spec.get("public_id") or "",
                    spec.get("project_no") or "",
                )
                if col
            ):
                fallback_col = spec.get("public_id") or spec.get("project_no") or "업무키"
                row[fallback_col] = work["business_key"]
            row.update(
                {
                    "_store_id": work["id"],
                    "_record_version": work["record_version"],
                    "_business_key": work["business_key"],
                    "_deleted": bool(work.get("deleted_at")),
                    "_source": work.get("source") or "",
                    "_evidence": work.get("evidence") or "",
                }
            )
            rendered.append(
                (
                    int(location["row_number"]) if location else 2_147_483_647,
                    str(work["business_key"]),
                    row,
                )
            )
        rendered.sort(key=lambda item: (item[0], item[1]))
        return [item[2] for item in rendered]

    def overlay_rows(
        self, sheet: str, rows: Iterable[Mapping[str, Any]], key_col: str
    ) -> List[Dict[str, Any]]:
        """Overlay canonical DB values on a legacy Excel read model.

        Soft-deleted canonical rows disappear, DB-only rows are appended, and
        legacy rows without a canonical match remain visible during migration.
        """

        key_col = _require_text("key_col", key_col, 300)
        canonical = self.list_sheet_rows(sheet, include_deleted=True)
        by_key: Dict[str, Dict[str, Any]] = {}
        for row in canonical:
            key = str(row.get(key_col) or row.get("_business_key") or "").strip()
            if key:
                by_key[key] = row
        seen: set[str] = set()
        output: List[Dict[str, Any]] = []
        for legacy in rows:
            legacy_row = dict(legacy)
            key = str(legacy_row.get(key_col) or "").strip()
            current = by_key.get(key)
            if current:
                seen.add(key)
                if current.get("_deleted"):
                    continue
                legacy_row.update(current)
            output.append(legacy_row)
        for key, current in sorted(by_key.items()):
            if key not in seen and not current.get("_deleted"):
                output.append(dict(current))
        return output

    def status(self) -> Dict[str, Any]:
        """Return a compact health/readiness status for the app UI."""

        with self.reader() as conn:
            change_tip = conn.execute(
                "SELECT COALESCE(MAX(id),0) AS change_seq, "
                "MAX(created_at) AS last_change_at FROM change_event"
            ).fetchone()
            counts = {
                "work_active": int(
                    conn.execute(
                        "SELECT COUNT(*) FROM work_item WHERE deleted_at IS NULL"
                    ).fetchone()[0]
                ),
                "work_deleted": int(
                    conn.execute(
                        "SELECT COUNT(*) FROM work_item WHERE deleted_at IS NOT NULL"
                    ).fetchone()[0]
                ),
                "outbox_pending": int(
                    conn.execute(
                        "SELECT COUNT(*) FROM outbox WHERE status IN ('pending','failed','leased')"
                    ).fetchone()[0]
                ),
                "import_conflicts_open": int(
                    conn.execute(
                        "SELECT COUNT(*) FROM import_conflict WHERE status='open'"
                    ).fetchone()[0]
                ),
                "change_seq": int(change_tip["change_seq"]),
                "completion_evidence_active": int(
                    conn.execute(
                        "SELECT COUNT(*) FROM completion_evidence WHERE active=1"
                    ).fetchone()[0]
                ),
            }
            mode_row = conn.execute(
                "SELECT value_json FROM app_setting WHERE key='source_of_truth_mode'"
            ).fetchone()
            export = conn.execute(
                "SELECT export_id,status,snapshot_seq,local_path,finished_at "
                "FROM export_run ORDER BY id DESC LIMIT 1"
            ).fetchone()
            completed_export = conn.execute(
                "SELECT export_id,status,snapshot_seq,local_path,finished_at "
                "FROM export_run "
                "WHERE status IN ('verified','partial') AND finished_at IS NOT NULL "
                "ORDER BY finished_at DESC,id DESC LIMIT 1"
            ).fetchone()
            return {
                "ok": True,
                "db_path": str(self.db_path),
                "schema_version": SCHEMA_VERSION,
                "journal_mode": str(
                    conn.execute("PRAGMA journal_mode").fetchone()[0]
                ).lower(),
                "foreign_keys": bool(
                    int(conn.execute("PRAGMA foreign_keys").fetchone()[0])
                ),
                "source_of_truth_mode": (
                    json.loads(mode_row["value_json"])
                    if mode_row
                    else "db_primary_pending_cutover"
                ),
                "last_export": dict(export) if export else None,
                "last_completed_export": (
                    dict(completed_export) if completed_export else None
                ),
                "last_change_at": str(change_tip["last_change_at"] or ""),
                **counts,
            }

    def upsert_completion_evidence(
        self,
        entries: Iterable[Mapping[str, Any]],
        *,
        source: str = "objective-sync",
        actor: str = "automation",
        idempotency_key: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Persist objective completion decisions in the canonical database.

        A repeated observation only advances ``last_seen`` and does not create
        another event/outbox message.  Retractions are intentionally explicit;
        a temporarily unavailable source must never erase previously verified
        proof.
        """

        normalized: List[Dict[str, Any]] = []
        for raw in entries or []:
            item = dict(raw or {})
            owner = str(item.get("owner") or "").strip()
            task_kind = _require_text("task_kind", item.get("task_kind"), 200)
            record_id = _require_text("record_id", item.get("record_id"), 500)
            completed_on = _require_text(
                "completed_on", str(item.get("completed_on") or "")[:10], 10
            )
            try:
                completed_date = date.fromisoformat(completed_on)
            except ValueError as exc:
                raise ValidationError(
                    f"invalid completion date for {task_kind}/{record_id}"
                ) from exc
            if completed_date > date.today():
                raise ValidationError(
                    f"future completion date for {task_kind}/{record_id}"
                )
            basis = _require_text("basis", item.get("basis"), 4_000)
            normalized.append(
                {
                    "owner": owner[:200],
                    "task_kind": task_kind,
                    "record_id": record_id,
                    "project_no": str(item.get("project_no") or item.get("project") or "").strip()[:500],
                    "status": _require_text("status", item.get("status") or "완료", 300),
                    "completed_on": completed_on,
                    "basis": basis,
                    "source_ref": str(item.get("source_ref") or "").strip()[:2_000],
                    "first_seen": str(item.get("first_seen") or "").strip(),
                    "last_seen": str(item.get("last_seen") or "").strip(),
                }
            )
        payload = {"entries": normalized, "source": source, "actor": actor}
        result: Dict[str, Any] = {
            "ok": True,
            "accepted": len(normalized),
            "created": 0,
            "updated": 0,
            "unchanged": 0,
            "event_seq": 0,
        }
        now = _utcnow()
        with self.transaction() as conn:
            replay, request_hash = self._idempotency_replay(
                conn, "completion:batch", idempotency_key, payload
            )
            if replay is not None:
                return replay
            for item in normalized:
                key = (item["owner"], item["task_kind"], item["record_id"])
                row = conn.execute(
                    "SELECT * FROM completion_evidence "
                    "WHERE owner=? AND task_kind=? AND record_id=?",
                    key,
                ).fetchone()
                proof_sha = sha256_json(
                    {
                        "project_no": item["project_no"],
                        "status": item["status"],
                        "completed_on": item["completed_on"],
                        "basis": item["basis"],
                        "source_ref": item["source_ref"],
                    }
                )
                first_seen = item["first_seen"] or now
                last_seen = item["last_seen"] or now
                before = dict(row) if row else None
                if row is None:
                    evidence_id = self._new_id("cmp")
                    conn.execute(
                        """
                        INSERT INTO completion_evidence(
                            id,owner,task_kind,record_id,project_no,status,completed_on,
                            basis,source,source_ref,evidence_sha256,active,first_seen,
                            last_seen,updated_at
                        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,1,?,?,?)
                        """,
                        (
                            evidence_id,
                            *key,
                            item["project_no"],
                            item["status"],
                            item["completed_on"],
                            item["basis"],
                            source,
                            item["source_ref"],
                            proof_sha,
                            first_seen,
                            last_seen,
                            now,
                        ),
                    )
                    action = "completion_create"
                    result["created"] += 1
                else:
                    evidence_id = str(row["id"])
                    same = (
                        str(row["project_no"]) == item["project_no"]
                        and str(row["status"]) == item["status"]
                        and str(row["completed_on"]) == item["completed_on"]
                        and str(row["basis"]) == item["basis"]
                        and str(row["source_ref"]) == item["source_ref"]
                        and int(row["active"]) == 1
                    )
                    if same:
                        conn.execute(
                            "UPDATE completion_evidence SET last_seen=? WHERE id=?",
                            (last_seen, evidence_id),
                        )
                        result["unchanged"] += 1
                        continue
                    conn.execute(
                        """
                        UPDATE completion_evidence SET
                            project_no=?,status=?,completed_on=?,basis=?,source=?,
                            source_ref=?,evidence_sha256=?,active=1,last_seen=?,updated_at=?
                        WHERE id=?
                        """,
                        (
                            item["project_no"],
                            item["status"],
                            item["completed_on"],
                            item["basis"],
                            source,
                            item["source_ref"],
                            proof_sha,
                            last_seen,
                            now,
                            evidence_id,
                        ),
                    )
                    action = "completion_update"
                    result["updated"] += 1
                after = dict(
                    conn.execute(
                        "SELECT * FROM completion_evidence WHERE id=?", (evidence_id,)
                    ).fetchone()
                )
                _event_id, seq = self._append_event(
                    conn,
                    work_id=None,
                    aggregate_type="completion_evidence",
                    aggregate_key="|".join(key),
                    action=action,
                    actor=actor,
                    before=before,
                    after=after,
                    source=source,
                    evidence=item["basis"],
                    source_ref=item["source_ref"],
                    source_sha256=proof_sha,
                    topic="completion.changed",
                )
                result["event_seq"] = max(result["event_seq"], seq)
            self._save_idempotency(
                conn, "completion:batch", idempotency_key, request_hash, result
            )
        return result

    def list_completion_evidence(
        self,
        *,
        owner: Optional[str] = None,
        task_kind: Optional[str] = None,
        active_only: bool = True,
        limit: int = 10_000,
    ) -> List[Dict[str, Any]]:
        if limit < 1 or limit > 50_000:
            raise ValidationError("invalid completion evidence limit")
        clauses: List[str] = []
        params: List[Any] = []
        if owner is not None:
            clauses.append("owner=?")
            params.append(owner)
        if task_kind is not None:
            clauses.append("task_kind=?")
            params.append(task_kind)
        if active_only:
            clauses.append("active=1")
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        with self.reader() as conn:
            return [
                dict(row)
                for row in conn.execute(
                    "SELECT * FROM completion_evidence"
                    + where
                    + " ORDER BY completed_on DESC,id LIMIT ?",
                    (*params, limit),
                )
            ]

    def reserve_public_id(
        self,
        kind: str,
        day: date | datetime | str,
        prefix: str,
    ) -> Dict[str, Any]:
        """Reserve a collision-free public ID under one SQLite write lock.

        Plain prefixes yield ``PREFIX-YYMM-NNN``.  A format prefix may use
        ``{yyyy}``, ``{yyyymm}``, ``{yymm}``, ``{mm}``, ``{dd}``, and
        ``{seq}``/``{seq03}`` for legacy ID shapes.
        """

        kind = _require_text("kind", kind, 200)
        prefix = _require_text("prefix", prefix, 100)
        if isinstance(day, datetime):
            day_value = day.date()
        elif isinstance(day, date):
            day_value = day
        else:
            text = _require_text("date", day, 30)[:10]
            try:
                day_value = date.fromisoformat(text)
            except ValueError as exc:
                raise ValidationError("date must be YYYY-MM-DD") from exc
        day_key = day_value.isoformat()
        with self.transaction() as conn:
            row = conn.execute(
                "SELECT last_number FROM public_id_sequence "
                "WHERE kind=? AND day=? AND prefix=?",
                (kind, day_key, prefix),
            ).fetchone()
            number = int(row["last_number"] if row else 0)
            while True:
                number += 1
                if "{" in prefix:
                    try:
                        candidate = prefix.format(
                            yyyy=f"{day_value:%Y}",
                            yyyymm=f"{day_value:%Y%m}",
                            yymm=f"{day_value:%y%m}",
                            mm=f"{day_value:%m}",
                            dd=f"{day_value:%d}",
                            seq=number,
                            seq03=f"{number:03d}",
                        )
                    except (KeyError, ValueError) as exc:
                        raise ValidationError(f"invalid public ID prefix format: {exc}") from exc
                else:
                    candidate = f"{prefix.rstrip('-')}-{day_value:%y%m}-{number:03d}"
                exists = conn.execute(
                    "SELECT 1 FROM work_item WHERE public_id=?", (candidate,)
                ).fetchone()
                if not exists:
                    break
            conn.execute(
                """
                INSERT INTO public_id_sequence(kind,day,prefix,last_number,updated_at)
                VALUES(?,?,?,?,?)
                ON CONFLICT(kind,day,prefix) DO UPDATE SET
                    last_number=excluded.last_number,updated_at=excluded.updated_at
                """,
                (kind, day_key, prefix, number, _utcnow()),
            )
            event_id, event_seq = self._append_event(
                conn,
                work_id=None,
                aggregate_type="public_id_sequence",
                aggregate_key=f"{kind}:{day_key}:{prefix}",
                action="reserve",
                actor="app",
                before=None,
                after={"public_id": candidate, "number": number, "day": day_key},
                source="app",
                topic="public_id.reserved",
            )
            return {
                "public_id": candidate,
                "number": number,
                "kind": kind,
                "date": day_key,
                "event_id": event_id,
                "event_seq": event_seq,
            }

    def snapshot_payload(self, *, include_deleted: bool = False) -> Dict[str, Any]:
        """Return a deterministic canonical snapshot (no volatile generated time)."""

        # One deferred read transaction pins a single WAL snapshot across all
        # tables.  Without it, a concurrent write between SELECTs could produce
        # a JSON snapshot that never existed as one database state.
        with self.transaction(immediate=False) as conn:
            schema_rows = [
                dict(r)
                for r in conn.execute(
                    "SELECT version,applied_at,checksum FROM schema_version ORDER BY version"
                )
            ]
            settings = []
            for row in conn.execute("SELECT * FROM app_setting ORDER BY key"):
                item = dict(row)
                item["value"] = json.loads(item.pop("value_json"))
                settings.append(item)
            clause = "" if include_deleted else " WHERE deleted_at IS NULL"
            ids = conn.execute(
                "SELECT id FROM work_item"
                + clause
                + " ORDER BY kind,business_key,id"
            ).fetchall()
            works = [
                self._work_from_conn(conn, row["id"], include_deleted=include_deleted)
                for row in ids
            ]
            completions = [
                dict(row)
                for row in conn.execute(
                    "SELECT * FROM completion_evidence "
                    "ORDER BY owner,task_kind,record_id,id"
                )
            ]
            public_id_sequences = [
                dict(row)
                for row in conn.execute(
                    "SELECT * FROM public_id_sequence ORDER BY kind,day,prefix"
                )
            ]
            seq = int(
                conn.execute("SELECT COALESCE(MAX(id),0) AS n FROM change_event").fetchone()[
                    "n"
                ]
            )
            return {
                "format": "csos-app-store-snapshot/v1",
                "schema_version": SCHEMA_VERSION,
                "schema_history": schema_rows,
                "change_seq": seq,
                "include_deleted": bool(include_deleted),
                "settings": settings,
                "work_items": works,
                "completion_evidence": completions,
                "public_id_sequences": public_id_sequences,
            }

    def snapshot(self, *, include_deleted: bool = False) -> Tuple[Dict[str, Any], str]:
        payload = self.snapshot_payload(include_deleted=include_deleted)
        return payload, sha256_json(payload)

    def backup_to(self, destination: os.PathLike[str] | str) -> Path:
        """Create a transactionally consistent SQLite backup and atomically replace it."""

        destination = Path(destination).resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=f".{destination.name}.", suffix=".tmp", dir=str(destination.parent)
        )
        os.close(fd)
        temp_path = Path(temp_name)
        source = self.connect()
        target = sqlite3.connect(str(temp_path))
        try:
            source.backup(target)
            target.commit()
            check = target.execute("PRAGMA quick_check").fetchone()[0]
            if check != "ok":
                raise StoreError(f"backup quick_check failed: {check}")
        finally:
            target.close()
            source.close()
        try:
            os.replace(temp_path, destination)
        finally:
            if temp_path.exists():
                temp_path.unlink()
        return destination

    def start_export(
        self,
        *,
        export_id: str,
        snapshot_seq: int,
        snapshot_sha256: str,
        template_sha256: str,
        plan_sha256: str,
        mode: str,
        local_path: str,
    ) -> Dict[str, Any]:
        with self.transaction() as conn:
            existing = conn.execute(
                "SELECT * FROM export_run WHERE export_id=?", (export_id,)
            ).fetchone()
            if existing:
                result = dict(existing)
                for key, value in (
                    ("snapshot_seq", snapshot_seq),
                    ("snapshot_sha256", snapshot_sha256),
                    ("template_sha256", template_sha256),
                    ("plan_sha256", plan_sha256),
                ):
                    if str(result[key]) != str(value):
                        raise ValidationError(
                            f"export_id {export_id} already exists with different {key}"
                        )
                return result
            conn.execute(
                """
                INSERT INTO export_run(
                    export_id,snapshot_seq,snapshot_sha256,template_sha256,
                    plan_sha256,mode,status,local_path,created_at
                ) VALUES(?,?,?,?,?,?,'planned',?,?)
                """,
                (
                    export_id,
                    snapshot_seq,
                    snapshot_sha256,
                    template_sha256,
                    plan_sha256,
                    mode,
                    local_path,
                    _utcnow(),
                ),
            )
            return dict(
                conn.execute(
                    "SELECT * FROM export_run WHERE export_id=?", (export_id,)
                ).fetchone()
            )

    def finish_export(
        self,
        export_id: str,
        *,
        status: str,
        manifest_sha256: str = "",
        local_path: str = "",
        error: str = "",
    ) -> Dict[str, Any]:
        if status not in {"planned", "partial", "verified", "failed"}:
            raise ValidationError("invalid export status")
        with self.transaction() as conn:
            cur = conn.execute(
                """
                UPDATE export_run SET status=?,manifest_sha256=?,
                    local_path=CASE WHEN ?<>'' THEN ? ELSE local_path END,
                    error=?,finished_at=?
                WHERE export_id=?
                """,
                (
                    status,
                    manifest_sha256,
                    local_path,
                    local_path,
                    str(error)[:8_000],
                    _utcnow(),
                    export_id,
                ),
            )
            if cur.rowcount != 1:
                raise NotFoundError(f"export run not found: {export_id}")
            return dict(
                conn.execute(
                    "SELECT * FROM export_run WHERE export_id=?", (export_id,)
                ).fetchone()
            )

    def export_run(self, export_id: str) -> Dict[str, Any]:
        with self.reader() as conn:
            row = conn.execute(
                "SELECT * FROM export_run WHERE export_id=?", (export_id,)
            ).fetchone()
            if row is None:
                raise NotFoundError(f"export run not found: {export_id}")
            return dict(row)


_DEFAULT_STORE_LOCK = threading.Lock()
_DEFAULT_STORES: Dict[str, AppStore] = {}
_SYNTHETIC_DB_PATH = Path(tempfile.gettempdir()) / (
    f"coupang_app_store_synthetic_{os.getpid()}_{uuid.uuid4().hex}.db"
)


def default_store(db_path: Optional[os.PathLike[str] | str] = None) -> AppStore:
    """Return the process-local store for the configured canonical DB path.

    ``COUPANG_APP_DB_PATH`` always wins over the repository default.  Synthetic
    runners can therefore point at their temporary DB before importing this
    module and never touch the real application database.
    """

    configured = db_path or os.environ.get("COUPANG_APP_DB_PATH")
    if configured is None and (
        os.environ.get("CSOS_SYNTHETIC") == "1"
        or os.environ.get("COUPANG_SYNTHETIC_MODE") == "1"
        or os.environ.get("COUPANG_TEST_MODE") == "1"
    ):
        configured = _SYNTHETIC_DB_PATH
    probe = AppStore(configured)
    cache_key = str(probe.db_path)
    with _DEFAULT_STORE_LOCK:
        store = _DEFAULT_STORES.get(cache_key)
        if store is None:
            store = probe.initialize()
            _DEFAULT_STORES[cache_key] = store
        return store


def apply_legacy_items(
    items: Iterable[Mapping[str, Any]],
    source: str,
    idempotency_key: Optional[str] = None,
    actor: Optional[str] = None,
) -> Dict[str, Any]:
    return default_store().apply_legacy_items(
        items,
        source,
        idempotency_key=idempotency_key,
        actor=actor,
    )


def overlay_rows(
    sheet: str, rows: Iterable[Mapping[str, Any]], key_col: str
) -> List[Dict[str, Any]]:
    return default_store().overlay_rows(sheet, rows, key_col)


def list_sheet_rows(sheet: str) -> List[Dict[str, Any]]:
    return default_store().list_sheet_rows(sheet)


def status() -> Dict[str, Any]:
    return default_store().status()


def reserve_public_id(
    kind: str, day: date | datetime | str, prefix: str
) -> Dict[str, Any]:
    return default_store().reserve_public_id(kind, day, prefix)


def self_test() -> bool:
    """Exercise transactional, locking, import, outbox and snapshot invariants."""

    assert SHEET_SPECS["03_현장작업실적"]["status"] == "완료여부"
    assert SHEET_SPECS["13_PO발주관리"]["status"] == "PO상태(자동)"
    invoice_spec = SHEET_SPECS["15_세금계산서관리"]
    assert invoice_spec["public_id"] == invoice_spec["business_key"] == "계산서관리ID"
    assert invoice_spec["relation_id"] == "정산ID"
    assert invoice_spec["status"] == "발행상태(자동)"
    assert invoice_spec["issued_at"] == "실제발행일"
    receipt_spec = SHEET_SPECS["16_입금수금관리"]
    assert receipt_spec["public_id"] == receipt_spec["business_key"] == "입금관리ID"
    assert receipt_spec["relation_id"] == "정산ID" and receipt_spec["status"] == ""

    with tempfile.TemporaryDirectory(prefix="app_store_selftest_") as tmp:
        db = Path(tmp) / "store.db"
        store = AppStore(db).initialize()
        with store.reader() as conn:
            assert str(conn.execute("PRAGMA journal_mode").fetchone()[0]).lower() == "wal"
            assert int(conn.execute("PRAGMA foreign_keys").fetchone()[0]) == 1

        created = store.create_work(
            kind="돌발AS",
            business_key="UJ-SELF-001",
            project_no="JS-SELF-001",
            camp_name="테스트캠프",
            status="접수",
            fields={"공급가액": 100_000, "완료": False},
            evidence="self-test",
            idempotency_key="create-1",
        )
        replay = store.create_work(
            kind="돌발AS",
            business_key="UJ-SELF-001",
            project_no="JS-SELF-001",
            camp_name="테스트캠프",
            status="접수",
            fields={"공급가액": 100_000, "완료": False},
            evidence="self-test",
            idempotency_key="create-1",
        )
        assert replay["work"]["id"] == created["work"]["id"]
        assert replay["idempotent_replay"] is True
        try:
            store.create_work(
                kind="돌발AS",
                business_key="DIFFERENT",
                idempotency_key="create-1",
            )
            raise AssertionError("idempotency conflict was not raised")
        except IdempotencyConflict:
            pass

        work_id = created["work"]["id"]
        updated = store.update_work(
            work_id,
            expected_version=1,
            patch={"status": "작업완료", "fields": {"완료": True}},
            evidence="objective DB proof",
            idempotency_key="update-1",
        )
        assert updated["work"]["record_version"] == 2
        try:
            store.update_work(
                work_id, expected_version=1, patch={"status": "stale"}
            )
            raise AssertionError("stale optimistic update was accepted")
        except VersionConflict:
            pass

        setting = store.set_setting(
            "dashboard.layout", {"cards": ["today"]}, idempotency_key="setting-1"
        )
        assert setting["setting"]["record_version"] == 1
        setting2 = store.set_setting(
            "dashboard.layout",
            {"cards": ["today", "inspection"]},
            expected_version=1,
            idempotency_key="setting-2",
        )
        assert setting2["setting"]["record_version"] == 2

        completion = {
            "owner": "류지영",
            "task_kind": "field_as",
            "record_id": "AS-SELF-001",
            "project": "UJ-SELF-001",
            "status": "류지영 완료",
            "completed_on": "2026-08-09",
            "basis": "합성 객관근거",
        }
        completion_result = store.upsert_completion_evidence(
            [completion], idempotency_key="completion-1"
        )
        assert completion_result["created"] == 1
        completion_repeat = store.upsert_completion_evidence([completion])
        assert completion_repeat["unchanged"] == 1
        completion_rows = store.list_completion_evidence(owner="류지영")
        assert len(completion_rows) == 1
        assert completion_rows[0]["record_id"] == "AS-SELF-001"

        legacy_items = [
            {
                "sheet": "02_돌발AS접수",
                "key_col": "접수ID",
                "key": "AS-LEGACY-1",
                "col": "진행상태",
                "value": "접수",
                "vtype": "text",
                "only_if_empty": True,
                "evidence": "legacy self-test",
            },
            {
                "sheet": "02_돌발AS접수",
                "key_col": "접수ID",
                "key": "AS-LEGACY-1",
                "col": "담당기사",
                "value": "김필우",
                "vtype": "text",
                "only_if_empty": True,
                "evidence": "legacy self-test",
            },
            {
                "sheet": "00_대시보드",
                "key_col": "-",
                "key": "B3",
                "cell": "B3",
                "value": "2026-08-10",
                "vtype": "date",
                "only_if_empty": False,
                "evidence": "legacy cell self-test",
            },
        ]
        legacy = store.apply_legacy_items(
            legacy_items,
            "legacy-self-test",
            idempotency_key="legacy-batch-1",
        )
        assert legacy["ok"] and legacy["created"] == 1 and legacy["updated"] == 1
        assert legacy["settings"] == 1
        legacy_replay = store.apply_legacy_items(
            legacy_items,
            "legacy-self-test",
            idempotency_key="legacy-batch-1",
        )
        assert legacy_replay["idempotent_replay"] is True
        seq_before_noop = store.status()["change_seq"]
        legacy_noop = store.apply_legacy_items(
            legacy_items,
            "different-ingest-prefix",
            idempotency_key="legacy-batch-2",
        )
        assert legacy_noop["ok"] and legacy_noop["skipped"] == len(legacy_items)
        assert store.status()["change_seq"] == seq_before_noop
        row_cells = [
            {
                "sheet": "02_돌발AS접수",
                "cell": "B120",
                "key_col": "-",
                "key": "B120",
                "col": "프로젝트NO",
                "value": "UJ-ROW-1",
                "only_if_empty": True,
                "evidence": "project_resolve row self-test",
            },
            {
                "sheet": "02_돌발AS접수",
                "cell": "C120",
                "key_col": "-",
                "key": "C120",
                "col": "캠프명",
                "value": "행그룹캠프",
                "only_if_empty": True,
                "evidence": "project_resolve row self-test",
            },
            {
                "sheet": "02_돌발AS접수",
                "cell": "D120",
                "key_col": "-",
                "key": "D120",
                "col": "진행상태",
                "value": "접수",
                "only_if_empty": True,
                "evidence": "project_resolve row self-test",
            },
        ]
        grouped = store.apply_legacy_items(
            row_cells, "project-resolve-A", idempotency_key="row-group-A"
        )
        assert grouped["ok"] and grouped["created"] == 1 and grouped["settings"] == 0
        grouped_work = store.get_work(kind="돌발AS", business_key="UJ-ROW-1")
        assert grouped_work["camp_name"] == "행그룹캠프"
        assert grouped_work["fields"]["진행상태"] == "접수"
        grouped_overlay = store.overlay_rows("02_돌발AS접수", [], "프로젝트NO")
        assert any(row.get("프로젝트NO") == "UJ-ROW-1" for row in grouped_overlay)
        grouped_seq = store.status()["change_seq"]
        grouped_noop = store.apply_legacy_items(
            row_cells, "project-resolve-B", idempotency_key="row-group-B"
        )
        assert grouped_noop["ok"] and grouped_noop["skipped"] == 1
        assert store.status()["change_seq"] == grouped_seq

        invoice_cells = [
            {
                "sheet": "15_세금계산서관리", "cell": "A210", "key_col": "-",
                "key": "A210", "col": "계산서관리ID", "value": "TX-SELF-001",
            },
            {
                "sheet": "15_세금계산서관리", "cell": "B210", "key_col": "-",
                "key": "B210", "col": "정산ID", "value": "JS-SELF-210",
            },
            {
                "sheet": "15_세금계산서관리", "cell": "L210", "key_col": "-",
                "key": "L210", "col": "실제발행일", "value": "2026-08-10",
            },
            {
                "sheet": "15_세금계산서관리", "cell": "AF210", "key_col": "-",
                "key": "AF210", "col": "발행상태(자동)", "value": "발행완료",
            },
        ]
        invoice_result = store.apply_legacy_items(
            invoice_cells, "invoice-header-self-test", idempotency_key="invoice-row"
        )
        assert invoice_result["ok"] and invoice_result["created"] == 1
        invoice = store.get_work(kind="세금계산서", business_key="TX-SELF-001")
        assert invoice["public_id"] == "TX-SELF-001"
        assert invoice["fields"]["정산ID"] == "JS-SELF-210"
        assert invoice["fields"]["실제발행일"] == "2026-08-10"
        assert invoice["status"] == "발행완료"

        receipt_cells = [
            {
                "sheet": "16_입금수금관리", "cell": "A220", "key_col": "-",
                "key": "A220", "col": "입금관리ID", "value": "RC-SELF-001",
            },
            {
                "sheet": "16_입금수금관리", "cell": "B220", "key_col": "-",
                "key": "B220", "col": "정산ID", "value": "JS-SELF-210",
            },
            {
                "sheet": "16_입금수금관리", "cell": "N220", "key_col": "-",
                "key": "N220", "col": "입금일", "value": "2026-08-11",
            },
        ]
        receipt_result = store.apply_legacy_items(
            receipt_cells, "receipt-header-self-test", idempotency_key="receipt-row"
        )
        assert receipt_result["ok"] and receipt_result["created"] == 1
        receipt = store.get_work(kind="입금수금", business_key="RC-SELF-001")
        assert receipt["public_id"] == "RC-SELF-001"
        assert receipt["fields"]["정산ID"] == "JS-SELF-210"
        assert receipt["status"] == "" and "수금상태" not in receipt["fields"]

        legacy_invoice = store.create_work(
            kind="세금계산서",
            business_key="JS-LEGACY-230",
            fields={
                "계산서관리ID": "TX-LEGACY-230",
                "정산ID": "JS-LEGACY-230",
            },
            idempotency_key="legacy-invoice-before-v586",
        )["work"]
        legacy_identity_update = store.apply_legacy_items(
            [
                {
                    "sheet": "15_세금계산서관리",
                    "key_col": "계산서관리ID",
                    "key": "TX-LEGACY-230",
                    "col": "비고",
                    "value": "v586 관리ID로 기존 행 갱신",
                    "only_if_empty": False,
                }
            ],
            "legacy-management-id-self-test",
        )
        assert legacy_identity_update["ok"] and legacy_identity_update["updated"] == 1
        assert legacy_identity_update["created"] == 0
        assert store.get_work(work_id=legacy_invoice["id"])["fields"]["비고"] == (
            "v586 관리ID로 기존 행 갱신"
        )
        assert store.get_work(
            kind="세금계산서", business_key="TX-LEGACY-230"
        )["id"] == legacy_invoice["id"]
        legacy_shadow = store.shadow_import(
            import_id="imp-legacy-management-id",
            sheet="15_세금계산서관리",
            business_key="TX-LEGACY-230",
            business_key_col="계산서관리ID",
            row_number=230,
            kind="세금계산서",
            fields={
                "계산서관리ID": "TX-LEGACY-230",
                "정산ID": "JS-LEGACY-230",
            },
            apply_if_missing=True,
            idempotency_key="shadow-legacy-management-id",
        )
        assert legacy_shadow["status"] == "matched"
        assert legacy_shadow["work_id"] == legacy_invoice["id"]
        with store.reader() as conn:
            assert conn.execute(
                "SELECT COUNT(*) FROM work_item WHERE kind=? AND business_key=?",
                ("세금계산서", "TX-LEGACY-230"),
            ).fetchone()[0] == 0
        try:
            store.create_work(
                kind="세금계산서",
                business_key="TX-LEGACY-230",
                public_id="TX-LEGACY-230",
                fields={"계산서관리ID": "TX-LEGACY-230"},
                idempotency_key="legacy-management-id-duplicate",
            )
            raise AssertionError("legacy management identity duplicate was accepted")
        except ValidationError:
            pass

        duplicate_legacy = store.create_work(
            kind="세금계산서",
            business_key="JS-LEGACY-231",
            fields={"정산ID": "JS-LEGACY-231"},
            idempotency_key="legacy-invoice-duplicate-fixture",
        )["work"]
        with store.transaction() as conn:
            store._put_field(
                conn,
                duplicate_legacy["id"],
                "계산서관리ID",
                "TX-LEGACY-230",
                "self-test",
                {"source": "legacy-fixture"},
                _utcnow(),
            )
        duplicate_shadow = store.shadow_import(
            import_id="imp-duplicate-management-id",
            sheet="15_세금계산서관리",
            business_key="TX-LEGACY-230",
            business_key_col="계산서관리ID",
            row_number=231,
            kind="세금계산서",
            fields={"계산서관리ID": "TX-LEGACY-230"},
            apply_if_missing=True,
            idempotency_key="shadow-duplicate-management-id",
        )
        assert duplicate_shadow["status"] == "conflict"
        assert duplicate_shadow["work_id"] is None
        duplicate_conflicts = store.list_import_conflicts(
            import_id="imp-duplicate-management-id"
        )
        assert len(duplicate_conflicts) == 1
        assert duplicate_conflicts[0]["reason"] == "duplicate_management_identity"

        legacy_receipt = store.create_work(
            kind="입금수금",
            business_key="JS-LEGACY-240",
            fields={
                "입금관리ID": "RC-LEGACY-240",
                "정산ID": "JS-LEGACY-240",
            },
            idempotency_key="legacy-receipt-before-v586",
        )["work"]
        legacy_receipt_shadow = store.shadow_import(
            import_id="imp-legacy-receipt-management-id",
            sheet="16_입금수금관리",
            business_key="RC-LEGACY-240",
            business_key_col="입금관리ID",
            row_number=240,
            kind="입금수금",
            fields={
                "입금관리ID": "RC-LEGACY-240",
                "정산ID": "JS-LEGACY-240",
            },
            apply_if_missing=True,
            idempotency_key="shadow-legacy-receipt-management-id",
        )
        assert legacy_receipt_shadow["status"] == "matched"
        assert legacy_receipt_shadow["work_id"] == legacy_receipt["id"]

        duplicate_relation = store.create_work(
            kind="세금계산서", business_key="TX-SELF-002", public_id="TX-SELF-002",
            fields={"정산ID": "JS-SELF-210"}, idempotency_key="invoice-duplicate-relation",
        )
        assert duplicate_relation["work"]["public_id"] == "TX-SELF-002"
        ambiguous_relation = store.apply_legacy_items(
            [{
                "sheet": "15_세금계산서관리", "key_col": "정산ID",
                "key": "JS-SELF-210", "col": "비고", "value": "관계키 모호성",
            }],
            "relation-ambiguity-self-test",
        )
        assert not ambiguous_relation["ok"]
        assert ambiguous_relation["errors"][0]["type"] == "ValidationError"
        sheet_rows = store.list_sheet_rows("02_돌발AS접수")
        legacy_row = next(row for row in sheet_rows if row.get("접수ID") == "AS-LEGACY-1")
        assert legacy_row["담당기사"] == "김필우"
        overlaid = store.overlay_rows(
            "02_돌발AS접수",
            [{"접수ID": "AS-LEGACY-1", "담당기사": "옛값", "legacy_only": 1}],
            "접수ID",
        )
        assert overlaid[0]["담당기사"] == "김필우" and overlaid[0]["legacy_only"] == 1
        reserved1 = store.reserve_public_id("돌발AS", "2026-08-10", "AS")
        reserved2 = store.reserve_public_id("돌발AS", "2026-08-10", "AS")
        assert reserved1["public_id"] == "AS-2608-001"
        assert reserved2["number"] == reserved1["number"] + 1
        health = store.status()
        assert health["journal_mode"] == "wal" and health["foreign_keys"]

        imported = store.shadow_import(
            import_id="imp-self-1",
            sheet="06_업무관리",
            business_key="UJ-SELF-002",
            row_number=44,
            kind="돌발AS",
            fields={"청구상태": "완료"},
            source_file="self.xlsx",
            source_sha256="a" * 64,
            apply_if_missing=True,
            idempotency_key="shadow-1",
        )
        assert imported["status"] == "created"
        conflict = store.shadow_import(
            import_id="imp-self-2",
            sheet="06_업무관리",
            business_key="UJ-SELF-002",
            row_number=44,
            kind="돌발AS",
            fields={"청구상태": "미완료"},
            source_file="self.xlsx",
            source_sha256="b" * 64,
            apply_if_missing=False,
            idempotency_key="shadow-2",
        )
        assert conflict["status"] == "conflict"
        assert store.list_import_conflicts(import_id="imp-self-2")

        token, messages = store.lease_outbox(limit=100)
        assert messages
        assert store.ack_outbox(token, [m["id"] for m in messages]) == len(messages)

        snap1, digest1 = store.snapshot()
        snap2, digest2 = store.snapshot()
        assert snap1 == snap2 and digest1 == digest2
        backup = store.backup_to(Path(tmp) / "backup.db")
        backup_conn = sqlite3.connect(str(backup))
        try:
            assert backup_conn.execute("PRAGMA quick_check").fetchone()[0] == "ok"
        finally:
            backup_conn.close()

        deleted = store.soft_delete_work(
            work_id,
            expected_version=2,
            reason="self-test",
            idempotency_key="delete-1",
        )
        assert deleted["work"]["deleted_at"]
        try:
            store.get_work(work_id)
            raise AssertionError("soft-deleted work was visible by default")
        except NotFoundError:
            pass
        assert store.get_work(work_id, include_deleted=True)["record_version"] == 3

        old_configured_db = os.environ.pop("COUPANG_APP_DB_PATH", None)
        old_synthetic = os.environ.get("CSOS_SYNTHETIC")
        os.environ["CSOS_SYNTHETIC"] = "1"
        try:
            synthetic_store = default_store()
            assert synthetic_store.db_path == _SYNTHETIC_DB_PATH.resolve()
            assert synthetic_store.db_path != AppStore().db_path
        finally:
            if old_configured_db is not None:
                os.environ["COUPANG_APP_DB_PATH"] = old_configured_db
            if old_synthetic is None:
                os.environ.pop("CSOS_SYNTHETIC", None)
            else:
                os.environ["CSOS_SYNTHETIC"] = old_synthetic
            with _DEFAULT_STORE_LOCK:
                _DEFAULT_STORES.pop(str(_SYNTHETIC_DB_PATH.resolve()), None)
            for candidate in (
                _SYNTHETIC_DB_PATH,
                Path(str(_SYNTHETIC_DB_PATH) + "-wal"),
                Path(str(_SYNTHETIC_DB_PATH) + "-shm"),
            ):
                try:
                    candidate.unlink()
                except FileNotFoundError:
                    pass
    return True


def _main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="SQLite application system-of-record")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--init", metavar="DB_PATH")
    parser.add_argument("--snapshot", metavar="DB_PATH")
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("app_store self-test: OK")
        return 0
    if args.init:
        store = AppStore(args.init).initialize()
        print(store.db_path)
        return 0
    if args.snapshot:
        payload, digest = AppStore(args.snapshot).snapshot()
        print(canonical_json({"sha256": digest, "snapshot": payload}))
        return 0
    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
