# -*- coding: utf-8 -*-
"""Bridge verified legacy facts into the canonical application database.

The legacy SQLite file remains an evidence/archive compatibility source.  This
module makes the application store current immediately after a collector or
reconciler has queued facts; it never writes an Excel workbook.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sqlite3
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from app_store import AppStore, VersionConflict, canonical_json, default_store, sha256_json


ROOT = Path(__file__).resolve().parent
DEFAULT_LEDGER_DB = ROOT / "db" / "ledger_queue.db"
DEFAULT_REPORT = ROOT / "reports" / "canonical_sync.json"

KIND_ROUTE = {
    "field_as": "돌발AS",
    "field_pm": "정기점검",
    "field_install": "신규납품설치",
    "settlement": "정산",
}
STATUS_ROUTE = {
    "field_as": "작업완료",
    "field_pm": "완료",
    "field_install": "완료",
}
TERMINAL_CONFLICTS = ("취소", "철회", "점검불가", "AS전환")
SETTLEMENT_RANK = {
    "": 0,
    "금액확정": 10,
    "거래명세서발행": 20,
    "세금계산서발행": 30,
    "청구진행": 40,
    "입금완료": 50,
}


def _utc_now() -> str:
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _atomic_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=1, default=str)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def _read_rows(conn: sqlite3.Connection, sql: str) -> List[sqlite3.Row]:
    try:
        return list(conn.execute(sql))
    except sqlite3.OperationalError as exc:
        if "no such table" in str(exc).lower():
            return []
        raise


def _ledger_facts(path: Path) -> Dict[str, Any]:
    if not path.is_file():
        return {"pending": [], "completion": [], "error": f"legacy DB 없음: {path}"}
    uri = f"file:{path.as_posix()}?mode=ro"
    conn = sqlite3.connect(uri, uri=True, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        pending = [
            dict(row)
            for row in _read_rows(
                conn,
                "SELECT id,sheet,key_col,key,cell,col,value,vtype,evidence,only_if_empty "
                "FROM pending WHERE status='pending' ORDER BY id",
            )
        ]
        completion: List[Dict[str, Any]] = []
        for row in _read_rows(
            conn,
            "SELECT kind,record_id,project,status,completed_on,basis,first_seen,last_seen "
            "FROM work_resolution ORDER BY kind,record_id",
        ):
            completion.append(
                {
                    "owner": "",
                    "task_kind": f"field_{row['kind']}",
                    "record_id": row["record_id"],
                    "project": row["project"],
                    "status": row["status"] or "완료",
                    "completed_on": row["completed_on"],
                    "basis": row["basis"],
                    "first_seen": row["first_seen"],
                    "last_seen": row["last_seen"],
                    "source_ref": "ledger_queue.db/work_resolution",
                }
            )
        for row in _read_rows(
            conn,
            "SELECT settle_id,project,status,basis,first_seen,last_seen "
            "FROM resolution ORDER BY settle_id",
        ):
            completion.append(
                {
                    "owner": "",
                    "task_kind": "settlement",
                    "record_id": row["settle_id"],
                    "project": row["project"],
                    "status": row["status"] or "완료",
                    "completed_on": str(row["first_seen"] or "")[:10],
                    "basis": row["basis"],
                    "first_seen": row["first_seen"],
                    "last_seen": row["last_seen"],
                    "source_ref": "ledger_queue.db/resolution",
                }
            )
        for row in _read_rows(
            conn,
            "SELECT owner,task_kind,record_id,project,status,completed_on,basis,"
            "first_seen,last_seen FROM staff_resolution "
            "ORDER BY owner,task_kind,record_id",
        ):
            item = dict(row)
            item["source_ref"] = "ledger_queue.db/staff_resolution"
            completion.append(item)
        return {"pending": pending, "completion": completion, "error": ""}
    finally:
        conn.close()


def _pending_token(rows: Sequence[Mapping[str, Any]]) -> str:
    return "canonical-pending:" + sha256_json(
        [
            {
                key: row.get(key)
                for key in (
                    "id",
                    "sheet",
                    "key_col",
                    "key",
                    "cell",
                    "col",
                    "value",
                    "vtype",
                    "only_if_empty",
                )
            }
            for row in rows
        ]
    )


def _work_index(store: AppStore) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    index: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    for kind in sorted(set(KIND_ROUTE.values())):
        by_key: Dict[str, List[Dict[str, Any]]] = {}
        try:
            rows = store.list_work(kind=kind, limit=10_000)
        except Exception:
            rows = []
        for work in rows:
            keys = {
                str(work.get("public_id") or "").strip(),
                str(work.get("business_key") or "").strip(),
                str(work.get("project_no") or "").strip(),
            }
            for key in keys - {""}:
                by_key.setdefault(key, []).append(work)
        index[kind] = by_key
    return index


def _match_work(
    index: Mapping[str, Mapping[str, List[Dict[str, Any]]]],
    entry: Mapping[str, Any],
) -> Tuple[Optional[Dict[str, Any]], str]:
    kind = KIND_ROUTE.get(str(entry.get("task_kind") or ""))
    if not kind:
        return None, "업무 레코드 직접 연결 대상 아님"
    by_key = index.get(kind) or {}
    record_id = str(entry.get("record_id") or "").strip()
    project = str(entry.get("project") or entry.get("project_no") or "").strip()
    if record_id:
        exact = by_key.get(record_id) or []
        if len(exact) == 1:
            return exact[0], "ID 정확 일치"
        if len(exact) > 1:
            return None, f"ID 다중 일치 {len(exact)}건"
    if project:
        candidates = by_key.get(project) or []
        unique = {str(row["id"]): row for row in candidates}
        if len(unique) == 1:
            return next(iter(unique.values())), "프로젝트NO 유일 일치"
        if len(unique) > 1:
            return None, f"프로젝트NO 다중 일치 {len(unique)}건"
    return None, "연결 업무 없음"


def _settlement_status(proof_status: str) -> str:
    text = str(proof_status or "")
    if "수금" in text or "입금" in text:
        return "입금완료"
    if "계산서" in text or "발행상태" in text:
        return "세금계산서발행"
    if any(word in text for word in ("금액", "견적", "명세서")):
        return "금액확정"
    return ""


def _promoted_status(work: Mapping[str, Any], entry: Mapping[str, Any]) -> str:
    current = str(work.get("status") or "").strip()
    if any(word in current for word in TERMINAL_CONFLICTS):
        return current
    task_kind = str(entry.get("task_kind") or "")
    target = STATUS_ROUTE.get(task_kind, "")
    if task_kind == "settlement":
        target = _settlement_status(str(entry.get("status") or ""))
        if SETTLEMENT_RANK.get(current, 0) >= SETTLEMENT_RANK.get(target, 0):
            return current
    return target or current


def _apply_completion_to_work(
    store: AppStore,
    index: Mapping[str, Mapping[str, List[Dict[str, Any]]]],
    entries: Iterable[Mapping[str, Any]],
) -> Dict[str, Any]:
    result: Dict[str, Any] = {"updated": 0, "unchanged": 0, "unmatched": []}
    for entry in entries:
        # The owner-specific row is a responsibility audit.  The ownerless row
        # is the single objective business decision that may promote status.
        if str(entry.get("owner") or "").strip():
            continue
        work, reason = _match_work(index, entry)
        if work is None:
            result["unmatched"].append(
                {
                    "task_kind": entry.get("task_kind"),
                    "record_id": entry.get("record_id"),
                    "project": entry.get("project") or entry.get("project_no"),
                    "reason": reason,
                }
            )
            continue
        fields = {
            "객관완료여부": True,
            "객관완료일": str(entry.get("completed_on") or "")[:10],
            "객관완료상태": str(entry.get("status") or "완료"),
            "객관완료근거": str(entry.get("basis") or ""),
        }
        target_status = _promoted_status(work, entry)
        current_fields = dict(work.get("fields") or {})
        status_changed = target_status != str(work.get("status") or "")
        changed_fields = {
            key: value
            for key, value in fields.items()
            if canonical_json(current_fields.get(key)) != canonical_json(value)
        }
        if not status_changed and not changed_fields:
            result["unchanged"] += 1
            continue
        proof = sha256_json(
            {
                "work_id": work["id"],
                "status": target_status,
                "fields": fields,
                "basis": entry.get("basis"),
            }
        )
        for attempt in range(5):
            current = store.get_work(str(work["id"]))
            patch: Dict[str, Any] = {"fields": changed_fields}
            promoted = _promoted_status(current, entry)
            if promoted != str(current.get("status") or ""):
                patch["status"] = promoted
            if not patch.get("fields") and "status" not in patch:
                result["unchanged"] += 1
                break
            version = int(current["record_version"])
            try:
                store.update_work(
                    str(current["id"]),
                    expected_version=version,
                    patch=patch,
                    actor="automation",
                    source="objective-sync",
                    evidence=str(entry.get("basis") or "")[:4_000],
                    source_ref=str(entry.get("source_ref") or ""),
                    idempotency_key=f"objective:{current['id']}:{proof}:v{version}",
                )
                result["updated"] += 1
                break
            except VersionConflict:
                if attempt == 4:
                    raise
    result["unmatched_count"] = len(result["unmatched"])
    result["unmatched"] = result["unmatched"][:200]
    return result


def sync(
    *,
    store: Optional[AppStore] = None,
    ledger_path: Optional[os.PathLike[str] | str] = None,
    report_path: Optional[os.PathLike[str] | str] = None,
) -> Dict[str, Any]:
    store = (store or default_store()).initialize()
    ledger = Path(ledger_path or os.environ.get("COUPANG_LEDGER_DB_PATH") or DEFAULT_LEDGER_DB)
    facts = _ledger_facts(ledger)
    started = _utc_now()
    result: Dict[str, Any] = {
        "ok": not bool(facts.get("error")),
        "started_at": started,
        "ledger_db": str(ledger),
        "pending_seen": len(facts["pending"]),
        "completion_seen": len(facts["completion"]),
        "error": facts.get("error") or "",
    }
    if facts["pending"]:
        result["pending_to_app"] = store.apply_legacy_items(
            facts["pending"],
            source="legacy-backlog-sync",
            actor="automation",
            idempotency_key=_pending_token(facts["pending"]),
        )
    else:
        result["pending_to_app"] = {
            "ok": True,
            "accepted": 0,
            "created": 0,
            "updated": 0,
            "settings": 0,
            "skipped": 0,
            "errors": [],
        }
    completion_token = "canonical-completion:" + sha256_json(facts["completion"])
    result["completion_to_app"] = store.upsert_completion_evidence(
        facts["completion"],
        source="ledger-objective-proof",
        actor="automation",
        idempotency_key=completion_token,
    )
    index = _work_index(store)
    result["work_promotions"] = _apply_completion_to_work(
        store, index, facts["completion"]
    )
    result["app_db"] = store.status()
    result["finished_at"] = _utc_now()
    result["ok"] = bool(
        result["ok"]
        and result["pending_to_app"].get("ok")
        and result["completion_to_app"].get("ok")
    )
    if report_path is not False:
        _atomic_json(Path(report_path or DEFAULT_REPORT), result)
    return result


def self_test() -> bool:
    with tempfile.TemporaryDirectory(prefix="canonical-sync-") as temp:
        root = Path(temp)
        ledger = root / "legacy.db"
        conn = sqlite3.connect(ledger)
        conn.executescript(
            """
            CREATE TABLE pending(
              id INTEGER PRIMARY KEY, sheet TEXT,key_col TEXT,key TEXT,cell TEXT,
              col TEXT,value TEXT,vtype TEXT,evidence TEXT,only_if_empty INTEGER,status TEXT
            );
            CREATE TABLE work_resolution(
              kind TEXT,record_id TEXT,project TEXT,status TEXT,completed_on TEXT,
              basis TEXT,first_seen TEXT,last_seen TEXT
            );
            CREATE TABLE resolution(
              settle_id TEXT,project TEXT,status TEXT,basis TEXT,first_seen TEXT,last_seen TEXT
            );
            CREATE TABLE staff_resolution(
              owner TEXT,task_kind TEXT,record_id TEXT,project TEXT,status TEXT,
              completed_on TEXT,basis TEXT,first_seen TEXT,last_seen TEXT
            );
            """
        )
        yesterday = (date.today() - timedelta(days=1)).isoformat()
        conn.execute(
            "INSERT INTO pending VALUES(1,?,?,?,?,?,?,?,?,?,?)",
            (
                "02_돌발AS접수",
                "접수ID",
                "AS-SYNC-1",
                "",
                "담당기사",
                "김필우",
                "text",
                "카톡 객관근거",
                1,
                "pending",
            ),
        )
        conn.execute(
            "INSERT INTO work_resolution VALUES(?,?,?,?,?,?,?,?)",
            (
                "as",
                "AS-SYNC-1",
                "UJ-SYNC-1",
                "완료",
                yesterday,
                "완료 사진+일자",
                yesterday,
                yesterday,
            ),
        )
        conn.execute(
            "INSERT INTO staff_resolution VALUES(?,?,?,?,?,?,?,?,?)",
            (
                "류지영",
                "field_as",
                "AS-SYNC-1",
                "UJ-SYNC-1",
                "류지영 완료",
                yesterday,
                "완료 사진+일자",
                yesterday,
                yesterday,
            ),
        )
        conn.commit()
        conn.close()
        store = AppStore(root / "app.db").initialize()
        store.create_work(
            kind="돌발AS",
            business_key="AS-SYNC-1",
            public_id="AS-SYNC-1",
            project_no="UJ-SYNC-1",
            status="접수",
            fields={"접수ID": "AS-SYNC-1"},
        )
        first = sync(store=store, ledger_path=ledger, report_path=root / "report.json")
        assert first["ok"]
        work = store.get_work(kind="돌발AS", business_key="AS-SYNC-1")
        assert work["status"] == "작업완료"
        assert work["fields"]["담당기사"] == "김필우"
        assert work["fields"]["객관완료여부"] is True
        assert len(store.list_completion_evidence()) == 2
        before_seq = store.status()["change_seq"]
        second = sync(store=store, ledger_path=ledger, report_path=root / "report2.json")
        assert second["ok"]
        assert store.status()["change_seq"] == before_seq
    return True


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="legacy evidence -> canonical app DB")
    parser.add_argument("--ledger-db")
    parser.add_argument("--report")
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("canonical_sync self-test: OK")
        return 0
    if args.status:
        report = Path(args.report or DEFAULT_REPORT)
        if report.is_file():
            print(report.read_text(encoding="utf-8"))
            return 0
        print(json.dumps({"ok": False, "error": "아직 동기화 기록 없음"}, ensure_ascii=False))
        return 1
    result = sync(ledger_path=args.ledger_db, report_path=args.report)
    print(json.dumps(result, ensure_ascii=False, indent=1, default=str))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
