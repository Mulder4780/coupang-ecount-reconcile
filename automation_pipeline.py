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
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from app_store import AppStore, default_store
from archive_export import ArchiveExporter, COVERAGE_CONTRACT_VERSION
import pid_alive


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


def _pid_alive(
    pid: int,
    born_before: Optional[float] = None,
    pid_started_at: Optional[float] = None,
) -> bool:
    """Conservative shared PID verdict: only a certain death is reclaimable."""

    # Windows ``os.kill(pid, 0)`` is not the process-liveness contract used by
    # the rest of this project.  ``pid_alive`` checks the actual exit code and
    # deliberately returns None when access/API state is inconclusive.  A live
    # owner's lock is more valuable than an eager retry, so unknown means live.
    # ``born_before`` (잠금이 쓰인 시각): 그보다 뒤에 태어난 프로세스는 pid 가
    # 재사용된 남이다 — 번호만 보고 살아 있다고 오판하지 않는다(검증 [210]).
    return pid_alive.owner_alive(
        pid, pid_started_at=pid_started_at, born_before=born_before
    ) is not False


def pipeline_lock_status(path: Path = LOCK_PATH) -> Dict[str, Any]:
    """Read and verify the current pipeline lock without changing it.

    New lock lines are ``pid claim_time token run_id pid_started_at``.  The
    fifth field is optional so four-field locks already on disk remain valid.
    """
    lock_path = Path(path)
    try:
        words = lock_path.read_text(encoding="ascii").split()
        stat = lock_path.stat()
    except FileNotFoundError:
        return {"exists": False, "alive": False, "path": str(lock_path)}
    except OSError:
        return {"exists": True, "alive": None, "path": str(lock_path)}
    try:
        owner_pid = int(words[0]) if words else 0
    except (TypeError, ValueError):
        owner_pid = 0
    try:
        claimed_at = float(words[1]) if len(words) >= 2 else stat.st_mtime
    except (TypeError, ValueError):
        claimed_at = stat.st_mtime
    try:
        owner_started_at = float(words[4]) if len(words) >= 5 else None
    except (TypeError, ValueError):
        owner_started_at = None
    verdict = (
        pid_alive.owner_alive(
            owner_pid,
            pid_started_at=owner_started_at,
            born_before=claimed_at,
        )
        if owner_pid > 0
        else None
    )
    return {
        "exists": True,
        "alive": verdict,
        "pid": owner_pid,
        "claimed_at": claimed_at,
        "pid_started_at": owner_started_at,
        "token": words[2] if len(words) >= 3 else "",
        "run_id": words[3] if len(words) >= 4 else "",
        "age_seconds": max(0.0, time.time() - stat.st_mtime),
        "path": str(lock_path),
    }


class LockOwnershipLost(RuntimeError):
    """The current round no longer owns the lock capability."""


class PipelineLock:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.owned = False
        self.token: Optional[str] = None
        self.run_id: Optional[str] = None
        self.pid_started_at: Optional[float] = None

    def _owner_words(self) -> Tuple[int, str, str, Optional[float]]:
        status = pipeline_lock_status(self.path)
        return (
            int(status.get("pid") or 0),
            str(status.get("token") or ""),
            str(status.get("run_id") or ""),
            status.get("pid_started_at"),
        )

    def acquire(self, run_id: str) -> Optional[str]:
        run_id = str(run_id or "").strip()
        if not run_id:
            raise ValueError("pipeline run_id is required")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        owner_identity = pid_alive.identity()
        owner_started_at = owner_identity.get("pid_started_at")
        for _attempt in range(2):
            token = uuid.uuid4().hex
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w", encoding="ascii") as handle:
                    fields = [str(os.getpid()), f"{time.time():.3f}", token, run_id]
                    if owner_started_at is not None:
                        fields.append(f"{float(owner_started_at):.7f}")
                    handle.write(" ".join(fields) + "\n")
                    handle.flush()
                    os.fsync(handle.fileno())
                self.owned = True
                self.token = token
                self.run_id = run_id
                self.pid_started_at = owner_started_at
                return token
            except FileExistsError:
                status = pipeline_lock_status(self.path)
                pid = int(status.get("pid") or 0)
                age = float(status.get("age_seconds") or 999999)
                if pid and status.get("alive") is not False:
                    return None
                if age < 60:
                    return None
                try:
                    self.path.unlink()
                except OSError:
                    return None
        return None

    def is_owner(self, token: Optional[str], run_id: Optional[str]) -> bool:
        if not self.owned or not token or not run_id:
            return False
        pid, disk_token, disk_run_id, disk_started_at = self._owner_words()
        return (
            pid == os.getpid()
            and pid_alive.owner_alive(
                pid,
                pid_started_at=disk_started_at,
            ) is True
            and (
                self.pid_started_at is None
                or disk_started_at is None
                or abs(float(self.pid_started_at) - float(disk_started_at))
                <= pid_alive.FINGERPRINT_TOLERANCE_S
            )
            and token == self.token == disk_token
            and run_id == self.run_id == disk_run_id
        )

    def heartbeat(self, token: Optional[str], run_id: Optional[str]) -> bool:
        if not self.is_owner(token, run_id):
            self.owned = False
            return False
        try:
            os.utime(self.path, None)
            return True
        except OSError:
            return False

    def release(self, token: Optional[str], run_id: Optional[str]) -> bool:
        if not self.is_owner(token, run_id):
            self.owned = False
            self.token = None
            self.run_id = None
            self.pid_started_at = None
            return False
        released = False
        try:
            self.path.unlink()
            released = True
        except OSError:
            released = False
        self.owned = False
        self.token = None
        self.run_id = None
        self.pid_started_at = None
        return released


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
        # download_intake가 Desktop 파일을 Z: 정본으로 먼저 옮긴 회차라도 다음
        # 증분 회차가 놓치지 않는다. 예전에는 신호가 Desktop/dropbox만 보아서
        # 신규 접수는 수동 반영됐는데 완료 보고는 대표화면에 안 들어왔다.
        try:
            import band_extract as _kakao_source
            kakao_files.extend(Path(path) for path in
                               _kakao_source.kakao_source_paths(dedupe_content=False))
        except Exception:
            pass
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


_STAGE_REASON_MAX = 400


def _stage_reason(stage: Mapping[str, Any], limit: int = _STAGE_REASON_MAX) -> str:
    """실패한 단계 → **비지 않는** 한 줄. 만드는 자리는 여기 하나다([162]).

    ★ 무엇이 있었나 (2026-08-20 실측) — `_run_source` 가 `stage["summary"]` 하나만
      담았다. 그 summary 는 자식의 **출력 꼬리**라, `ERP 판매·세금 대조` 가
      **20분 시간초과(rc=-9)** 로 끊긴 회차의 사유가 파일 이름 3,000자로 나갔다:
      `… ESD007E (36).xlsx→sale, … 판매 8781건 / 세금계산서 0건` + openpyxl 경고 둘.
      **겉은 경보인데 왜인지는 못 읽는다**([169]) — '시간초과' 라는 말이 한 글자도
      안 나오므로 사람은 멀쩡한 코드를 뒤지러 간다([172]).

    ★ 그래서 **비지 않는 것을 먼저 세운다**([292] 와 같은 규칙): 단계 이름과
      시간초과/종료코드는 절대 비지 않는다. 읽는 쪽이 앞부분만 자르므로([325])
      사유가 **맨 앞**에 와야 살아남는다.

    ★ 출력 꼬리는 버리지 않는다 — 진짜 트레이스백은 그 **끝**에 있다. 그래서
      뒤에서부터 싣고 자른 만큼을 숫자로 말한다([169]·[273]).
    """
    name = str(stage.get("name") or "이름 없는 단계")
    if stage.get("timed_out"):
        secs = round((stage.get("elapsed_ms") or 0) / 1000)
        head = "%s 실패 — 시간초과(%d초)" % (name, secs)
    else:
        head = "%s 실패 — 종료코드 %s" % (name, stage.get("returncode"))
    tail = " ".join(str(stage.get("summary") or "").split())
    room = limit - len(head) - 3
    if not tail or room <= 0:
        return head[:limit]
    if len(tail) > room:
        keep = max(0, room - 24)
        if not keep:
            return head[:limit]
        tail = "…(출력 %d자 중 뒤만) %s" % (len(tail), tail[-keep:])
    return (head + " · " + tail)[:limit]

def _default_state() -> Dict[str, Any]:
    return {
        "version": PIPELINE_VERSION,
        "running": False,
        "active_run_id": None,
        "lock_token": None,
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
        self._lock_token: Optional[str] = None
        self._run_id: Optional[str] = None

    def _reload_state(self) -> None:
        fresh = _read_json(self.state_path, _default_state())
        self.state = (
            fresh
            if isinstance(fresh, dict) and fresh.get("version") == PIPELINE_VERSION
            else _default_state()
        )

    def _acquire_run(self) -> bool:
        run_id = uuid.uuid4().hex
        token = self.lock.acquire(run_id)
        if not token:
            return False
        self._run_id = run_id
        self._lock_token = token
        # Another AutomationPipeline object may have been constructed before a
        # prior round finished.  Reload only after this round owns the lock so a
        # stale in-memory history/fingerprint cannot overwrite the newer state.
        self._reload_state()
        return True

    def _owns_run(self) -> bool:
        return self.lock.is_owner(self._lock_token, self._run_id)

    def _release_run(self) -> bool:
        released = self.lock.release(self._lock_token, self._run_id)
        self._lock_token = None
        self._run_id = None
        return released

    def _save(self) -> None:
        if not self._owns_run():
            raise LockOwnershipLost(
                f"automation round {self._run_id or '-'} no longer owns its lock"
            )
        _atomic_json(self.state_path, self.state)

    def _stage(self, name: str, args: Sequence[str], timeout: int) -> Dict[str, Any]:
        if not self.lock.heartbeat(self._lock_token, self._run_id):
            raise LockOwnershipLost(
                f"automation round {self._run_id or '-'} lost its lock before {name}"
            )
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
                        "error": _stage_reason(stage),
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
                # 카톡 완료 글도 band_extract가 같은 업무 레코드로 읽는다. 신규 접수만
                # 원장에 넣고 이 단계를 빼면 대표보고만 완료로 바뀌고 담당자 업무센터의
                # AppStore 상태는 계속 접수로 남는다 — 화면별 판정이 갈리는 사고다.
                ("카톡 앱 DB 신규·변경등록", [str(root / "band_canonical.py")], 900),
                # 카톡 접수 글도 `band_extract.load_records()` 가 같은 양식으로 읽는다 —
                # 캠프 담당자가 카톡으로만 바뀌는 날이 있어 여기서도 다시 뽑는다.
                ("캠프 담당자", [str(root / "camp_contacts.py"), "--write"], 900),
            ]
        if source == "band":
            commands: List[Tuple[str, Sequence[str], int]] = []
            if (root / "band" / ".band_token.json").is_file():
                commands.append(("밴드 인증수집", [str(root / "band" / "band_sync.py")], 1200))
            commands.extend(
                [
                    ("밴드 덤프 흡수", [str(root / "band" / "convert_dump.py")], 900),
                    # ★ 흡수 **바로 뒤**에 대기열을 다시 만든다 (2026-08-20 지시:
                    #   "자료 변경이나 업로드 및 추가가 반영되면 실시간으로 백그라운드에서
                    #   대기하고 있다가 긁어오는 알고리즘"). 방금 들어온 글은 대기열에서
                    #   빠지고 새로 생긴 구멍은 들어간다 — 그러면 로그인된 밴드 탭의
                    #   유저스크립트가 다음 폴링에서 그것을 그대로 받아 간다([182]).
                    #   09:50 에만 만들면 하루를 기다린다(바로 아래 캠프 담당자와 같은 이유).
                    #   실측 0.7초 — 캐시만 읽고 Z: 는 안 훑는다([168]).
                    ("밴드 수집대기열", [str(root / "band" / "collect_queue.py"), "--write"], 300),
                    ("밴드 앱 DB 신규·변경등록", [str(root / "band_canonical.py")], 900),
                    ("밴드 대조", [str(root / "band" / "band_reconcile.py")], 1200),
                    ("카톡·밴드 교차신호", [str(root / "cross_signal.py")], 900),
                    # 캠프 담당자는 **밴드 접수 글 본문**에서 나온다([295]). 09:50 에만
                    # 뽑으면 새 접수로 번호가 바뀌어도 하루를 기다린다 — 새 자료가 들어온
                    # 이 회차에서 같이 다시 뽑는다(2026-08-18 "추가·변경시 자동으로 반영").
                    # 사람이 앱에서 고친 값은 앱 DB 에 따로 있어 이 덮어쓰기에 안 지워진다.
                    ("캠프 담당자", [str(root / "camp_contacts.py"), "--write"], 900),
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
            # 밴드·카톡의 명시 접수취소를 원천업무 AppStore에 즉시 반영한다.
            # Excel은 건드리지 않고 아래 archive 회차가 검증된 보관본을 만든다.
            ("접수취소·원격해결 DB 동기화", [str(root / "cancel_watch.py"), "--sync"], 900),
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
        export_status = str(export.get("status") or "")
        if export_status not in {"verified", "partial"}:
            return {"acked": 0, "complete": False, "reason": "검증된 보관본 없음"}
        local_path = str(export.get("local_path") or "").strip()
        if not local_path:
            return {
                "acked": 0,
                "complete": False,
                "reason": "보관본 경로 없음",
            }
        artifact = Path(local_path).resolve()
        if not artifact.is_dir() or artifact.parent.name != "exports":
            return {
                "acked": 0,
                "complete": False,
                "reason": "보관본 경로 없음 또는 spool 구조 불일치",
            }
        try:
            exporter = ArchiveExporter(self.store, artifact.parent.parent)
            verification = exporter.verify_export(
                artifact, require_verified=export_status == "verified"
            )
            if not verification.get("ok"):
                raise ValueError(f"artifact verification failed: {verification.get('errors')}")
            manifest = _read_json(artifact / "manifest.json", {})
            if str(manifest.get("status") or "") != export_status:
                raise ValueError("DB export status differs from manifest")
            contract = manifest.get("coverage_contract") or {}
            if int(contract.get("version") or 0) < COVERAGE_CONTRACT_VERSION:
                raise ValueError("record coverage contract missing or obsolete")
            proof = _read_json(artifact / "adapter-result.json", {})
        except Exception as exc:
            return {
                "acked": 0,
                "complete": False,
                "reason": f"보관본 coverage 검증 실패: {type(exc).__name__}: {exc}"[:1000],
            }

        snapshot_seq = int(manifest.get("snapshot", {}).get("change_seq") or 0)
        if snapshot_seq != int(export.get("snapshot_seq") or 0):
            return {
                "acked": 0,
                "complete": False,
                "reason": "DB export snapshot_seq와 manifest 불일치",
            }
        covered: Dict[str, int] = {}
        covered_outcomes: Dict[str, str] = {}
        for entry in proof.get("record_coverage") or []:
            if not isinstance(entry, Mapping):
                continue
            work_id = str(entry.get("work_id") or "")
            outcome = str(entry.get("outcome") or "")
            try:
                revision = int(entry.get("record_version") or 0)
            except (TypeError, ValueError):
                continue
            if work_id and revision > 0 and outcome in {
                "applied",
                "unchanged",
                "archived_sidecar",
            }:
                if revision >= covered.get(work_id, 0):
                    covered[work_id] = revision
                    covered_outcomes[work_id] = outcome

        total = 0
        work_acked = 0
        snapshot_acked = 0
        deferred_newer = 0
        deferred_uncovered = 0
        for _batch in range(20):
            token, messages = self.store.lease_outbox(limit=1000, lease_seconds=300)
            if not messages:
                break
            ready: List[int] = []
            future: List[int] = []
            uncovered: List[int] = []
            batch_work = 0
            batch_snapshot = 0
            for message in messages:
                message_id = int(message["id"])
                payload = message.get("payload") or {}
                event_seq = int(payload.get("event_seq") or 0)
                if event_seq > snapshot_seq:
                    future.append(message_id)
                    continue
                aggregate_type = str(message.get("aggregate_type") or "")
                topic = str(message.get("topic") or "")
                if aggregate_type != "work_item" or topic != "work.changed":
                    # Settings, completion evidence and public-id reservations
                    # are semantically preserved by the independently verified
                    # canonical JSON/SQLite snapshot.  They do not claim Excel
                    # row coverage and therefore cannot mask a work conflict.
                    ready.append(message_id)
                    batch_snapshot += 1
                    continue
                record = payload.get("record") or {}
                work_id = str(record.get("id") or payload.get("aggregate_key") or "")
                try:
                    revision = int(record.get("record_version") or 0)
                except (TypeError, ValueError):
                    revision = 0
                if revision > 0 and covered.get(work_id, 0) >= revision:
                    ready.append(message_id)
                    batch_work += 1
                else:
                    uncovered.append(message_id)
            total += self.store.ack_outbox(token, ready)
            work_acked += batch_work
            snapshot_acked += batch_snapshot
            if future:
                self.store.fail_outbox(
                    token,
                    future,
                    "아직 검증된 보관본 스냅샷보다 새로운 이벤트",
                    retry_after_seconds=30,
                )
                deferred_newer += len(future)
            if uncovered:
                self.store.fail_outbox(
                    token,
                    uncovered,
                    "Excel 보관본의 주 시트/검증 sidecar 레코드 coverage에 포함되지 않음",
                    retry_after_seconds=30,
                )
                deferred_uncovered += len(uncovered)
        return {
            "acked": total,
            "work_acked": work_acked,
            "snapshot_acked": snapshot_acked,
            "snapshot_seq": snapshot_seq,
            "export_status": export_status,
            "covered_work_records": len(covered),
            "covered_outcomes": {
                "applied": sum(1 for value in covered_outcomes.values() if value == "applied"),
                "unchanged": sum(
                    1 for value in covered_outcomes.values() if value == "unchanged"
                ),
                "archived_sidecar": sum(
                    1
                    for value in covered_outcomes.values()
                    if value == "archived_sidecar"
                ),
            },
            "deferred_newer": deferred_newer,
            "deferred_uncovered": deferred_uncovered,
            "complete": export_status == "verified" and deferred_uncovered == 0,
        }

    def prime(self, *, trigger: str = "deployment") -> Dict[str, Any]:
        """Record the current source fingerprints without claiming ingestion.

        This is an explicit one-time deployment operation.  It prevents the
        first five-minute scheduler tick from replaying every historical raw
        file after the canonical DB has already been cut over and archived.
        ``last_success_at`` is deliberately left untouched: a baseline is not
        presented to operators as a successful collection run.
        """

        if not self._acquire_run():
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
                "run_id": self._run_id,
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
            self.state["active_run_id"] = None
            self.state["lock_token"] = None
            self.state["last_run"] = record
            history = list(self.state.get("history") or [])
            history.insert(0, dict(record))
            self.state["history"] = history[:30]
            self._save()
            return {"ok": True, **record}
        except LockOwnershipLost as exc:
            return {
                "ok": False,
                "status": "lost_lock",
                "run_id": self._run_id,
                "summary": str(exc),
            }
        except Exception as exc:
            return {
                "ok": False,
                "status": "failed",
                "run_id": self._run_id,
                "summary": f"{type(exc).__name__}: {exc}",
            }
        finally:
            self._release_run()

    def run_once(self, *, trigger: str = "scheduler", force: bool = False) -> Dict[str, Any]:
        if not self._acquire_run():
            return {"ok": True, "status": "already_running", "message": "자동화가 이미 실행 중입니다"}
        started = _now()
        self.run_record = {
            "status": "running",
            "run_id": self._run_id,
            "trigger": trigger,
            "started_at": started,
            "finished_at": None,
            "summary": "",
            "current_stage": "변경 감지",
            "stages": [],
        }
        failures: List[str] = []
        changed: List[str] = []
        try:
            self.state["running"] = True
            self.state["active_run_id"] = self._run_id
            self.state["lock_token"] = self._lock_token
            self.state["last_run"] = self.run_record
            self._save()
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
                            outbox_result = self._ack_archived_outbox()
                            self.run_record["outbox"] = outbox_result
                            if not outbox_result.get("complete"):
                                failures.append("Excel 보관본 일부충돌·coverage 미완료")
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
            self.state["active_run_id"] = None
            self.state["lock_token"] = None
            self.state["last_run"] = self.run_record
            history = list(self.state.get("history") or [])
            history.insert(0, dict(self.run_record))
            self.state["history"] = history[:30]
            self._save()
            return {"ok": not failures, **self.run_record}
        except LockOwnershipLost as exc:
            # A successor owns the shared state now.  Keep this result local;
            # writing even a truthful failure here would overwrite that round.
            self.run_record.update(
                {
                    "status": "lost_lock",
                    "finished_at": _now(),
                    "current_stage": "잠금 소유권 상실",
                    "summary": str(exc),
                    "failures": ["LockOwnershipLost"],
                }
            )
            return {"ok": False, **self.run_record}
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
            if self._owns_run():
                self.state["running"] = False
                self.state["active_run_id"] = None
                self.state["lock_token"] = None
                self.state["last_run"] = self.run_record
                self._save()
            return {"ok": False, **self.run_record}
        finally:
            self._release_run()


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
            "change_seq": int(app.get("change_seq") or 0),
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
        assert not result["ok"] and "kakao" in result["changed_sources"]
        assert not result["outbox"]["complete"]
        assert calls[:3] == ["카카오톡 원본 흡수·신규등록", "카카오톡 대조", "카톡·밴드 교차신호"]
        before = len(calls)
        again = pipe.run_once(trigger="self-test")
        assert again["ok"] and len(calls) == before
        api = status(root=root, state_path=root / "reports" / "state.json", store=store)
        assert set(api["sources"]) == {"kakao", "band", "erp"}
        assert "app_db" in api and "excel" in api and "human_gates" in api

        # Only work revisions explicitly covered in a main worksheet or the
        # semantically verified AppDB sidecar may be acknowledged.  A partial
        # artifact can still ACK its covered records and snapshot-only metadata
        # while leaving unresolved conflicts and newer events.
        import shutil

        from archive_export import _make_minimal_xlsx, sha256_file, sha256_json

        ack_store = AppStore(root / "ack.db").initialize()
        covered_work = ack_store.create_work(
            kind="돌발AS",
            business_key="ACK-COVERED",
            fields={"진행상태": "접수"},
            actor="selftest",
            source="selftest",
            evidence="covered outbox self-test",
            idempotency_key="ack-covered",
        )["work"]
        conflicted_work = ack_store.create_work(
            kind="돌발AS",
            business_key="ACK-CONFLICT",
            fields={"진행상태": "접수"},
            actor="selftest",
            source="selftest",
            evidence="conflicted outbox self-test",
            idempotency_key="ack-conflict",
        )["work"]
        ack_store.set_setting(
            "ack.selftest",
            {"value": 1},
            expected_version=0,
            actor="selftest",
            source="selftest",
            evidence="snapshot-only event self-test",
            idempotency_key="ack-setting",
        )
        ack_template = root / "ack-template.xlsx"
        _make_minimal_xlsx(ack_template)
        ack_exporter = ArchiveExporter(ack_store, root / "ack-spool")

        def partial_adapter(
            plan_path: Path, template_copy: Path, output_path: Path
        ) -> Mapping[str, Any]:
            plan = _read_json(plan_path, {})
            shutil.copyfile(template_copy, output_path)
            covered_record = next(
                item for item in plan["records"] if item["work_id"] == covered_work["id"]
            )
            conflict_record = next(
                item
                for item in plan["records"]
                if item["work_id"] == conflicted_work["id"]
            )
            return {
                "status": "partial",
                "snapshot_sha256": plan["snapshot"]["sha256"],
                "command_plan_sha256": sha256_json(plan),
                "rows_considered": len(plan["records"]),
                "record_coverage": [
                    {
                        "work_id": covered_record["work_id"],
                        "business_key": covered_record["business_key"],
                        "record_version": covered_record["record_version"],
                        "outcome": "applied",
                    }
                ],
                "conflicts": [
                    {
                        "work_id": conflict_record["work_id"],
                        "business_key": conflict_record["business_key"],
                        "record_version": conflict_record["record_version"],
                        "reason": "synthetic-conflict",
                    }
                ],
                "errors": [],
                "output_sha256": sha256_file(output_path),
            }

        partial_export = ack_exporter.run_local(
            template_path=ack_template,
            adapter=partial_adapter,
            dry_run=False,
        )
        assert partial_export["status"] == "partial"
        ack_store.update_work(
            covered_work["id"],
            expected_version=int(covered_work["record_version"]),
            patch={"status": "작업중"},
            actor="selftest",
            source="selftest",
            evidence="future event self-test",
            idempotency_key="ack-future",
        )
        ack_pipe = AutomationPipeline(
            root=root,
            store=ack_store,
            state_path=root / "reports" / "ack-state.json",
            lock_path=root / "reports" / ".ack-lock",
        )
        ack_result = ack_pipe._ack_archived_outbox()
        assert ack_result["acked"] == 2
        assert ack_result["work_acked"] == 1
        assert ack_result["snapshot_acked"] == 1
        assert ack_result["deferred_uncovered"] == 1
        assert ack_result["deferred_newer"] == 1
        assert ack_result["complete"] is False
        with ack_store.reader() as conn:
            outbox_status = {
                row["aggregate_key"]: row["status"]
                for row in conn.execute("SELECT aggregate_key,status FROM outbox ORDER BY id")
            }
        assert outbox_status[conflicted_work["id"]] == "failed"
        assert outbox_status["ack.selftest"] == "done"
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
