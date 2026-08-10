# -*- coding: utf-8 -*-
"""Incremental, event-driven source -> app DB -> Excel archive automation.

KakaoTalk is the only intentionally human-fed source: the user or Ryu Ji-young
exports a text file and drops/uploads it.  Band and ERP collectors still obey
their login boundary, but every step after authenticated collection is fully
automatic and idempotent.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from app_store import AppStore, default_store


ROOT = Path(__file__).resolve().parent
REPORT_DIR = ROOT / "reports"
STATE_PATH = REPORT_DIR / "automation_pipeline_state.json"
LOCK_PATH = REPORT_DIR / ".automation_pipeline.lock"
KAKAO_DROP = ROOT / "kakao" / "dropbox"
MAX_KAKAO_BYTES = 30_000_000
PIPELINE_VERSION = 1

StageRunner = Callable[[str, Sequence[str], int], Mapping[str, Any]]


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _atomic_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=1, default=str)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def _read_json(path: Path, default: Any) -> Any:
    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(handle)
        return value
    except (OSError, ValueError, TypeError):
        return default


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_name(name: str) -> str:
    base = os.path.basename(str(name or "").replace("\x00", ""))
    base = re.sub(r"[^0-9A-Za-z._()\-가-힣 ]+", "_", base).strip(" .")
    return (base or "KakaoTalk_upload.txt")[:180]


def submit_kakao_file(
    filename: str,
    data: bytes,
    *,
    drop_dir: Optional[os.PathLike[str] | str] = None,
) -> Dict[str, Any]:
    """Atomically enqueue one Kakao export; identical content is a no-op."""

    if not isinstance(data, (bytes, bytearray)) or not data:
        raise ValueError("카카오톡 파일 내용이 비었습니다")
    if len(data) > MAX_KAKAO_BYTES:
        raise ValueError("카카오톡 파일은 30MB 이하여야 합니다")
    safe = _safe_name(filename)
    if Path(safe).suffix.lower() != ".txt":
        raise ValueError("카카오톡 내보내기 .txt 파일만 등록할 수 있습니다")
    raw = bytes(data)
    # Reject obvious binary uploads early, without requiring one exact Kakao
    # version/locale string that could change later.
    if b"\x00" in raw[:4096]:
        raise ValueError("텍스트 파일이 아닙니다")
    digest = _sha256_bytes(raw)
    target_dir = Path(drop_dir or KAKAO_DROP)
    target_dir.mkdir(parents=True, exist_ok=True)
    for existing in target_dir.glob("*.txt"):
        try:
            if _sha256_file(existing) == digest:
                return {
                    "ok": True,
                    "duplicate": True,
                    "path": str(existing),
                    "sha256": digest,
                    "size": len(raw),
                }
        except OSError:
            continue
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = target_dir / f"{stamp}_{safe}"
    temp = target.with_name(f".{target.name}.{os.getpid()}.tmp")
    with temp.open("wb") as handle:
        handle.write(raw)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, target)
    return {
        "ok": True,
        "duplicate": False,
        "path": str(target),
        "sha256": digest,
        "size": len(raw),
    }


def _pid_alive(pid: int) -> bool:
    try:
        os.kill(int(pid), 0)
        return True
    except (OSError, ValueError, TypeError):
        return False


class PipelineLock:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.owned = False

    def acquire(self) -> bool:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        for _attempt in range(2):
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w", encoding="ascii") as handle:
                    handle.write(f"{os.getpid()} {time.time():.3f}\n")
                    handle.flush()
                    os.fsync(handle.fileno())
                self.owned = True
                return True
            except FileExistsError:
                try:
                    words = self.path.read_text(encoding="ascii").split()
                    pid = int(words[0]) if words else 0
                    age = time.time() - self.path.stat().st_mtime
                except (OSError, ValueError):
                    pid, age = 0, 999999
                if pid and _pid_alive(pid):
                    return False
                if age < 60:
                    return False
                try:
                    self.path.unlink()
                except OSError:
                    return False
        return False

    def heartbeat(self) -> None:
        if not self.owned:
            return
        try:
            os.utime(self.path, None)
        except OSError:
            pass

    def release(self) -> None:
        if not self.owned:
            return
        try:
            words = self.path.read_text(encoding="ascii").split()
            if words and int(words[0]) == os.getpid():
                self.path.unlink()
        except (OSError, ValueError):
            pass
        self.owned = False


def _metadata_signature(paths: Iterable[Path]) -> Tuple[str, Optional[float], int]:
    rows: List[Tuple[str, int, int]] = []
    latest: Optional[float] = None
    for path in paths:
        try:
            stat = path.stat()
        except OSError:
            continue
        rows.append((str(path), int(stat.st_mtime_ns), int(stat.st_size)))
        latest = stat.st_mtime if latest is None else max(latest, stat.st_mtime)
    rows.sort()
    digest = hashlib.sha256(
        json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
    return digest, latest, len(rows)


def _local_files(folder: Path, pattern: str) -> List[Path]:
    try:
        return [path for path in folder.glob(pattern) if path.is_file()]
    except OSError:
        return []


def _desktop_download_files(pattern: str) -> List[Path]:
    home = Path.home()
    folders = (
        home / "Desktop",
        home / "Downloads",
        home / "OneDrive" / "Desktop",
        home / "OneDrive" / "바탕 화면",
    )
    found: Dict[str, Path] = {}
    cutoff = time.time() - 7 * 86400
    for folder in folders:
        for path in _local_files(folder, pattern):
            try:
                if path.stat().st_mtime >= cutoff:
                    found[str(path.resolve()).lower()] = path
            except OSError:
                continue
    return list(found.values())


def source_signals(root: Path = ROOT) -> Dict[str, Dict[str, Any]]:
    include_user_drop = (
        Path(root).resolve() == ROOT.resolve()
        and os.environ.get("CSOS_SYNTHETIC") != "1"
        and os.environ.get("COUPANG_SYNTHETIC_MODE") != "1"
    )
    kakao_files = _local_files(root / "kakao" / "dropbox", "*.txt")
    if include_user_drop:
        kakao_files.extend(_desktop_download_files("KakaoTalk*.txt"))
    kakao_sig, kakao_latest, kakao_count = _metadata_signature(kakao_files)
    band_files = _local_files(root / "band" / "cache", "*.json")
    if include_user_drop:
        band_files.extend(_desktop_download_files("dump_*.json"))
    band_sig, band_latest, band_count = _metadata_signature(band_files)
    erp_markers = [
        root / "reports" / "download_intake.json",
        root / "reports" / "upload_intake.json",
        root / "reports" / "ERP_판매프로젝트색인.json",
        root / "reports" / "ERP원장_대조.csv",
    ]
    erp_name = re.compile(
        r"(?i)^(?:[A-Za-z0-9]{12,20}|E[A-Z]*\d{3,6}[A-Z]?|ECTAX\d+[A-Z]?)\.xlsx$"
    )
    if include_user_drop:
        erp_markers.extend(
            path for path in _desktop_download_files("*.xlsx") if erp_name.fullmatch(path.name)
        )
    erp_sig, erp_latest, erp_count = _metadata_signature(erp_markers)

    def row(signature: str, latest: Optional[float], count: int) -> Dict[str, Any]:
        return {
            "fingerprint": signature,
            "latest_mtime": (
                datetime.fromtimestamp(latest, timezone.utc).isoformat(timespec="seconds")
                if latest
                else None
            ),
            "count": count,
        }

    return {
        "kakao": {**row(kakao_sig, kakao_latest, kakao_count), "files": [str(p) for p in kakao_files]},
        "band": row(band_sig, band_latest, band_count),
        "erp": row(erp_sig, erp_latest, erp_count),
    }


def _real_stage_runner(name: str, args: Sequence[str], timeout: int) -> Dict[str, Any]:
    from proc_guard import run_tree

    env = dict(os.environ)
    env.update(
        {
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
            "COUPANG_UNATTENDED": "1",
        }
    )
    result = run_tree(
        [sys.executable, *map(str, args)],
        cwd=str(ROOT),
        env=env,
        timeout=timeout,
        drain_timeout=30,
        output_limit=160_000,
    )
    output = "\n".join(part for part in (result.stdout or "", result.stderr or "") if part)
    return {
        "name": name,
        "ok": result.returncode == 0 and not result.timed_out,
        "returncode": int(result.returncode),
        "timed_out": bool(result.timed_out),
        "stuck_pid": int(result.stuck_pid or 0),
        "summary": "\n".join([line for line in output.splitlines() if line.strip()][-12:])[-3000:],
    }


def _default_state() -> Dict[str, Any]:
    return {
        "version": PIPELINE_VERSION,
        "running": False,
        "sources": {
            name: {
                "status": "never",
                "fingerprint": "",
                "last_success_at": None,
                "latest_record_at": None,
                "attempts": 0,
                "error": "",
            }
            for name in ("kakao", "band", "erp")
        },
        "last_run": None,
        "history": [],
    }


class AutomationPipeline:
    def __init__(
        self,
        *,
        root: Path = ROOT,
        state_path: Optional[Path] = None,
        lock_path: Optional[Path] = None,
        store: Optional[AppStore] = None,
        stage_runner: Optional[StageRunner] = None,
    ) -> None:
        self.root = Path(root)
        self.state_path = Path(state_path or (self.root / "reports" / STATE_PATH.name))
        self.lock = PipelineLock(Path(lock_path or (self.root / "reports" / LOCK_PATH.name)))
        self.store = (store or default_store()).initialize()
        self.stage_runner = stage_runner or _real_stage_runner
        self.state = _read_json(self.state_path, _default_state())
        if not isinstance(self.state, dict) or self.state.get("version") != PIPELINE_VERSION:
            self.state = _default_state()
        self.run_record: Dict[str, Any] = {}

    def _save(self) -> None:
        _atomic_json(self.state_path, self.state)

    def _stage(self, name: str, args: Sequence[str], timeout: int) -> Dict[str, Any]:
        self.lock.heartbeat()
        self.run_record["current_stage"] = name
        self.run_record["updated_at"] = _now()
        self.state["last_run"] = self.run_record
        self._save()
        started = time.monotonic()
        try:
            result = dict(self.stage_runner(name, args, timeout))
        except Exception as exc:
            result = {
                "name": name,
                "ok": False,
                "returncode": -1,
                "timed_out": False,
                "summary": f"{type(exc).__name__}: {exc}",
            }
        result["elapsed_ms"] = round((time.monotonic() - started) * 1000)
        result["finished_at"] = _now()
        self.run_record.setdefault("stages", []).append(result)
        self._save()
        return result

    def _run_source(
        self,
        source: str,
        signal: Mapping[str, Any],
        commands: Sequence[Tuple[str, Sequence[str], int]],
    ) -> bool:
        source_state = self.state["sources"].setdefault(source, {})
        source_state.update(
            {
                "status": "running",
                "latest_record_at": signal.get("latest_mtime"),
                "attempts": int(source_state.get("attempts") or 0) + 1,
                "error": "",
            }
        )
        self._save()
        for name, args, timeout in commands:
            stage = self._stage(name, args, timeout)
            if not stage.get("ok"):
                source_state.update(
                    {
                        "status": "error",
                        "error": stage.get("summary") or f"rc={stage.get('returncode')}",
                        "last_attempt_at": _now(),
                    }
                )
                self._save()
                return False
        source_state.update(
            {
                "status": "ok",
                "fingerprint": signal.get("fingerprint") or "",
                "last_success_at": _now(),
                "last_attempt_at": _now(),
                "error": "",
            }
        )
        self._save()
        return True

    def _commands_for(
        self, source: str, signal: Mapping[str, Any]
    ) -> List[Tuple[str, Sequence[str], int]]:
        root = self.root
        if source == "kakao":
            files = [str(path) for path in signal.get("files") or []]
            if not files:
                return []
            return [
                ("카카오톡 원본 흡수·신규등록", [str(root / "kakao_apply.py"), *files], 1500),
                ("카카오톡 대조", [str(root / "kakao" / "kakao_reconcile.py")], 900),
                ("카톡·밴드 교차신호", [str(root / "cross_signal.py")], 900),
            ]
        if source == "band":
            commands: List[Tuple[str, Sequence[str], int]] = []
            if (root / "band" / ".band_token.json").is_file():
                commands.append(("밴드 인증수집", [str(root / "band" / "band_sync.py")], 1200))
            commands.extend(
                [
                    ("밴드 덤프 흡수", [str(root / "band" / "convert_dump.py")], 900),
                    ("밴드 앱 DB 신규·변경등록", [str(root / "band_canonical.py")], 900),
                    ("밴드 대조", [str(root / "band" / "band_reconcile.py")], 1200),
                    ("카톡·밴드 교차신호", [str(root / "cross_signal.py")], 900),
                ]
            )
            return commands
        if source == "erp":
            return [
                ("다운로드 원본 흡수", [str(root / "download_intake.py"), "--apply"], 900),
                ("업로드함 원본 분류", [str(root / "upload_intake.py"), "--apply"], 900),
                ("ERP 판매·세금 대조", [str(root / "ecount_reconcile.py")], 1200),
                ("ERP 판매 프로젝트 색인", [str(root / "erp_sales_index.py")], 900),
                ("ERP 매출서류 대조", [str(root / "erp_docs_check.py")], 1200),
                ("쿠팡 PO 대조", [str(root / "po_reconcile.py")], 900),
                ("입금 대조", [str(root / "receipt_fill.py"), "--queue"], 900),
                ("청구상태 대조", [str(root / "billing_status.py"), "--queue"], 900),
            ]
        return []

    def _common_commands(self) -> List[Tuple[str, Sequence[str], int]]:
        root = self.root
        return [
            ("입력 큐 DB 흡수", [str(root / "ledger_db.py"), "--intake"], 600),
            ("객관완료 판정", [str(root / "complete_verified.py"), "--queue"], 900),
            ("정산 객관완료", [str(root / "settlement_completion.py"), "--sync"], 1200),
            ("담당자 객관완료", [str(root / "staff_completion.py"), "--sync"], 600),
            ("앱 DB 정본 동기화", [str(root / "canonical_sync.py")], 900),
        ]

    def _needs_archive(self) -> bool:
        status = self.store.status()
        last_export = status.get("last_export") or {}
        return bool(
            int(status.get("outbox_pending") or 0) > 0
            or int(last_export.get("snapshot_seq") or 0) < int(status.get("change_seq") or 0)
            or last_export.get("status") != "verified"
        )

    def _ack_archived_outbox(self) -> Dict[str, Any]:
        status = self.store.status()
        export = status.get("last_export") or {}
        if export.get("status") != "verified":
            return {"acked": 0, "reason": "검증된 보관본 없음"}
        snapshot_seq = int(export.get("snapshot_seq") or 0)
        total = 0
        for _batch in range(20):
            token, messages = self.store.lease_outbox(limit=1000, lease_seconds=300)
            if not messages:
                break
            ready = [
                int(message["id"])
                for message in messages
                if int((message.get("payload") or {}).get("event_seq") or 0) <= snapshot_seq
            ]
            wait = [int(message["id"]) for message in messages if int(message["id"]) not in ready]
            total += self.store.ack_outbox(token, ready)
            if wait:
                self.store.fail_outbox(
                    token,
                    wait,
                    "아직 검증된 Excel 보관본 스냅샷에 포함되지 않음",
                    retry_after_seconds=30,
                )
                break
        return {"acked": total, "snapshot_seq": snapshot_seq}

    def prime(self, *, trigger: str = "deployment") -> Dict[str, Any]:
        """Record the current source fingerprints without claiming ingestion.

        This is an explicit one-time deployment operation.  It prevents the
        first five-minute scheduler tick from replaying every historical raw
        file after the canonical DB has already been cut over and archived.
        ``last_success_at`` is deliberately left untouched: a baseline is not
        presented to operators as a successful collection run.
        """

        if not self.lock.acquire():
            return {
                "ok": True,
                "status": "already_running",
                "message": "자동화가 이미 실행 중입니다",
            }
        try:
            signals = source_signals(self.root)
            primed: List[str] = []
            primed_at = _now()
            for source in ("kakao", "band", "erp"):
                signal = signals[source]
                source_state = self.state["sources"].setdefault(source, {})
                source_state.update(
                    {
                        "status": "baseline",
                        "fingerprint": signal.get("fingerprint") or "",
                        "latest_record_at": signal.get("latest_mtime"),
                        "baseline_at": primed_at,
                        "error": "",
                    }
                )
                primed.append(source)
            record = {
                "status": "primed",
                "trigger": trigger,
                "started_at": primed_at,
                "finished_at": primed_at,
                "summary": "현재 원본을 증분 자동화 기준선으로 등록했습니다",
                "current_stage": "기준선 등록 완료",
                "changed_sources": [],
                "failures": [],
                "primed_sources": primed,
                "stages": [],
            }
            self.state["running"] = False
            self.state["last_run"] = record
            history = list(self.state.get("history") or [])
            history.insert(0, dict(record))
            self.state["history"] = history[:30]
            self._save()
            return {"ok": True, **record}
        except Exception as exc:
            return {
                "ok": False,
                "status": "failed",
                "summary": f"{type(exc).__name__}: {exc}",
            }
        finally:
            self.lock.release()

    def run_once(self, *, trigger: str = "scheduler", force: bool = False) -> Dict[str, Any]:
        if not self.lock.acquire():
            return {"ok": True, "status": "already_running", "message": "자동화가 이미 실행 중입니다"}
        started = _now()
        self.run_record = {
            "status": "running",
            "trigger": trigger,
            "started_at": started,
            "finished_at": None,
            "summary": "",
            "current_stage": "변경 감지",
            "stages": [],
        }
        self.state["running"] = True
        self.state["last_run"] = self.run_record
        self._save()
        failures: List[str] = []
        changed: List[str] = []
        try:
            signals = source_signals(self.root)
            for source in ("kakao", "band", "erp"):
                signal = signals[source]
                previous = str(self.state["sources"].get(source, {}).get("fingerprint") or "")
                has_input = int(signal.get("count") or 0) > 0
                if not force and (not has_input or signal.get("fingerprint") == previous):
                    continue
                commands = self._commands_for(source, signal)
                if not commands:
                    continue
                changed.append(source)
                if not self._run_source(source, signal, commands):
                    failures.append(source)

            app_status = self.store.status()
            must_finalize = bool(changed or int(app_status.get("outbox_pending") or 0))
            if must_finalize:
                for name, args, timeout in self._common_commands():
                    stage = self._stage(name, args, timeout)
                    if not stage.get("ok"):
                        failures.append(name)
                        break
                # Refresh because canonical_sync and objective decisions create
                # events that must be present in the same archive snapshot.
                if not failures and self._needs_archive():
                    worker = self.root / "archive_worker.py"
                    if worker.is_file():
                        archive = self._stage(
                            "Excel 보관본 생성·검증", [str(worker), "--run"], 1800
                        )
                        if archive.get("ok"):
                            self.run_record["outbox"] = self._ack_archived_outbox()
                        else:
                            failures.append("Excel 보관본 생성·검증")
                    else:
                        failures.append("archive_worker.py 없음")

            status = "success" if not failures else ("partial" if changed else "failed")
            self.run_record.update(
                {
                    "status": status,
                    "finished_at": _now(),
                    "current_stage": "완료" if not failures else "실패 확인",
                    "summary": (
                        f"변경원천 {', '.join(changed) if changed else '없음'} · "
                        f"실패 {len(failures)}건"
                    ),
                    "changed_sources": changed,
                    "failures": failures,
                }
            )
            self.state["running"] = False
            self.state["last_run"] = self.run_record
            history = list(self.state.get("history") or [])
            history.insert(0, dict(self.run_record))
            self.state["history"] = history[:30]
            self._save()
            return {"ok": not failures, **self.run_record}
        except Exception as exc:
            self.run_record.update(
                {
                    "status": "failed",
                    "finished_at": _now(),
                    "current_stage": "예외 종료",
                    "summary": f"{type(exc).__name__}: {exc}",
                    "failures": [type(exc).__name__],
                }
            )
            self.state["running"] = False
            self.state["last_run"] = self.run_record
            self._save()
            return {"ok": False, **self.run_record}
        finally:
            self.lock.release()


def _age_days(value: Optional[str]) -> Optional[float]:
    if not value:
        return None
    try:
        stamp = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if stamp.tzinfo is None:
            stamp = stamp.replace(tzinfo=timezone.utc)
        return max(0.0, (datetime.now(timezone.utc) - stamp).total_seconds() / 86400)
    except (TypeError, ValueError):
        return None


def status(
    *,
    root: Path = ROOT,
    state_path: Optional[Path] = None,
    store: Optional[AppStore] = None,
) -> Dict[str, Any]:
    state = _read_json(Path(state_path or (root / "reports" / STATE_PATH.name)), _default_state())
    store_obj = (store or default_store()).initialize()
    app = store_obj.status()
    sources: Dict[str, Any] = {}
    gates: List[Dict[str, str]] = []
    limits = {"kakao": 2, "band": 1, "erp": 7}
    labels = {"kakao": "카카오톡", "band": "밴드", "erp": "ERP"}
    for name in ("kakao", "band", "erp"):
        raw = dict((state.get("sources") or {}).get(name) or {})
        latest = raw.get("latest_record_at") or raw.get("last_success_at")
        age = _age_days(latest)
        stale = age is None or age > limits[name]
        login_required = bool(name in {"band", "erp"} and stale)
        source_status = raw.get("status") or "never"
        if raw.get("error"):
            source_status = "error"
        elif stale:
            source_status = "stale"
        detail = raw.get("error") or (
            f"최신 자료 {age:.1f}일 전" if age is not None else "처리된 자료 기록 없음"
        )
        sources[name] = {
            "status": source_status,
            "last_success_at": raw.get("last_success_at"),
            "latest_record_at": latest,
            "detail": detail,
            "login_required": login_required,
            "stale": stale,
            "error": raw.get("error") or "",
        }
        if name == "kakao" and stale:
            gates.append(
                {
                    "source": "kakao",
                    "kind": "upload",
                    "message": "형님 또는 류지영 매니저가 오늘 카카오톡 내보내기 파일을 올려 주세요.",
                }
            )
        elif login_required:
            gates.append(
                {
                    "source": name,
                    "kind": "login",
                    "message": (
                        "로그인된 밴드 탭과 자동수집 유저스크립트를 확인해 주세요."
                        if name == "band"
                        else "ERP 로그인 세션을 확인해 주세요. 로그인 뒤 내보내기부터는 자동 처리됩니다."
                    ),
                }
            )
    excel: Dict[str, Any]
    try:
        import archive_worker

        excel = archive_worker.status(root=root, store=store_obj)
    except Exception as exc:
        last = app.get("last_export") or {}
        excel = {
            "status": last.get("status") or "missing",
            "last_good_at": last.get("finished_at"),
            "snapshot_seq": last.get("snapshot_seq"),
            "path": last.get("local_path") or "",
            "error": str(exc)[:300] if not last else "",
        }
    last_run = state.get("last_run") or {}
    return {
        "ok": True,
        "generated_at": _now(),
        "running": bool(state.get("running")),
        "last_run": {
            "status": last_run.get("status") or "never",
            "started_at": last_run.get("started_at"),
            "finished_at": last_run.get("finished_at"),
            "summary": last_run.get("summary") or "실행 기록 없음",
            "current_stage": last_run.get("current_stage") or "",
        },
        "sources": sources,
        "app_db": {
            "status": "ok" if app.get("ok") else "error",
            "last_sync_at": last_run.get("finished_at"),
            "pending": int(app.get("import_conflicts_open") or 0),
            "outbox_pending": int(app.get("outbox_pending") or 0),
            "work_active": int(app.get("work_active") or 0),
            "completion_evidence": int(app.get("completion_evidence_active") or 0),
            "detail": f"업무 {int(app.get('work_active') or 0):,}건 · 객관근거 {int(app.get('completion_evidence_active') or 0):,}건",
        },
        "excel": excel,
        "human_gates": gates,
    }


def self_test() -> bool:
    with tempfile.TemporaryDirectory(prefix="automation-pipeline-") as temp:
        root = Path(temp)
        for folder in (
            root / "reports",
            root / "kakao" / "dropbox",
            root / "band" / "cache",
        ):
            folder.mkdir(parents=True, exist_ok=True)
        (root / "archive_worker.py").write_text("# synthetic marker\n", encoding="utf-8")
        upload = submit_kakao_file(
            "KakaoTalk_test.txt", "테스트 카카오톡 대화\n2026-08-09\n".encode("utf-8"),
            drop_dir=root / "kakao" / "dropbox",
        )
        duplicate = submit_kakao_file(
            "another.txt", "테스트 카카오톡 대화\n2026-08-09\n".encode("utf-8"),
            drop_dir=root / "kakao" / "dropbox",
        )
        assert upload["ok"] and not upload["duplicate"] and duplicate["duplicate"]
        store = AppStore(root / "app.db").initialize()
        calls: List[str] = []

        def fake(name: str, args: Sequence[str], timeout: int) -> Mapping[str, Any]:
            calls.append(name)
            return {"name": name, "ok": True, "returncode": 0, "timed_out": False, "summary": "ok"}

        prime_pipe = AutomationPipeline(
            root=root,
            store=store,
            stage_runner=fake,
            state_path=root / "reports" / "prime-state.json",
            lock_path=root / "reports" / ".prime-lock",
        )
        primed = prime_pipe.prime(trigger="self-test")
        assert primed["ok"] and primed["status"] == "primed"
        assert all(
            item.get("status") == "baseline"
            for item in prime_pipe.state["sources"].values()
        )
        after_prime = prime_pipe.run_once(trigger="self-test")
        assert after_prime["ok"] and not calls

        pipe = AutomationPipeline(
            root=root,
            store=store,
            stage_runner=fake,
            state_path=root / "reports" / "state.json",
            lock_path=root / "reports" / ".lock",
        )
        result = pipe.run_once(trigger="self-test")
        assert result["ok"] and "kakao" in result["changed_sources"]
        assert calls[:3] == ["카카오톡 원본 흡수·신규등록", "카카오톡 대조", "카톡·밴드 교차신호"]
        before = len(calls)
        again = pipe.run_once(trigger="self-test")
        assert again["ok"] and len(calls) == before
        api = status(root=root, state_path=root / "reports" / "state.json", store=store)
        assert set(api["sources"]) == {"kakao", "band", "erp"}
        assert "app_db" in api and "excel" in api and "human_gates" in api
    return True


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="CSOS incremental automation pipeline")
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--watch", action="store_true")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--prime", action="store_true")
    parser.add_argument("--trigger", default="command")
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--interval", type=int, default=30)
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("automation_pipeline self-test: OK")
        return 0
    if args.status:
        print(json.dumps(status(), ensure_ascii=False, indent=1, default=str))
        return 0
    pipeline = AutomationPipeline()
    if args.prime:
        result = pipeline.prime(trigger=args.trigger or "deployment")
        print(json.dumps(result, ensure_ascii=False, indent=1, default=str))
        return 0 if result.get("ok") else 1
    if args.watch:
        interval = max(10, min(60, int(args.interval)))
        while True:
            result = pipeline.run_once(trigger=args.trigger or "watch", force=args.force)
            print(json.dumps(result, ensure_ascii=False, default=str))
            time.sleep(interval)
    result = pipeline.run_once(trigger=args.trigger, force=args.force)
    print(json.dumps(result, ensure_ascii=False, indent=1, default=str))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
