# -*- coding: utf-8 -*-
"""Claude Code 우선·Codex 폴백 작업 인계 큐.

실행 버튼이 실제 업무 스크립트를 한 번만 실행하는 동안, AI 검토·후속조치 요청은
별도 큐에 남긴다. AI CLI를 임의로 실행해 관리대장을 동시에 건드리지 않는다.
Claude Code가 사용 불가(크레딧/인증/명령 없음)이면 Codex 요청으로 자동 전환한다.

큐는 ``reports/agent_dispatch`` 아래에만 두며 비밀값·대화 원문은 저장하지 않는다.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
REPORT_DIR = ROOT / "reports" / "agent_dispatch"
STATUS_PATH = ROOT / "reports" / "agent_dispatch_status.json"
PROBE_TIMEOUT_SECONDS = 4
ROUTE_CACHE_SECONDS = 60

_UNAVAILABLE_RE = re.compile(
    r"credit|quota|usage|rate.?limit|billing|insufficient|not.?authenticated|"
    r"access is denied|permission denied|not recognized|not found|인증|크레딧|할당량|권한",
    re.I,
)


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(value, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def _clean_message(value: str, limit: int = 180) -> str:
    return " ".join((value or "").split())[:limit]


def probe_agent(name: str) -> dict[str, str]:
    """Return a small, non-secret availability record without launching any task."""
    command = {"claude": "claude", "codex": "codex"}.get(name)
    if not command:
        raise ValueError(f"unknown agent: {name}")
    executable = shutil.which(command)
    if not executable:
        return {"agent": name, "state": "unavailable", "reason": "실행 명령을 찾지 못함"}
    try:
        result = subprocess.run(
            [executable, "--version"], capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=PROBE_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return {"agent": name, "state": "unavailable", "reason": "버전 확인 시간 초과"}
    except OSError as exc:
        return {"agent": name, "state": "unavailable", "reason": _clean_message(str(exc))}

    message = _clean_message((result.stdout or "") + " " + (result.stderr or ""))
    if result.returncode == 0:
        return {"agent": name, "state": "ready", "reason": message or "사용 가능"}
    if _UNAVAILABLE_RE.search(message):
        reason = "크레딧·인증·권한 또는 실행 환경을 확인해야 함"
    else:
        reason = message or f"버전 확인 실패(코드 {result.returncode})"
    return {"agent": name, "state": "unavailable", "reason": reason}


def _cached_route() -> dict[str, Any] | None:
    """Reuse a recent probe result so normal status refreshes stay instant."""
    try:
        value = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
        checked_at = datetime.fromisoformat(str(value.get("checked_at", "")))
        if datetime.now() - checked_at > timedelta(seconds=ROUTE_CACHE_SECONDS):
            return None
        if value.get("selected") and isinstance(value.get("agents"), dict):
            return {key: value[key] for key in ("primary", "selected", "note", "agents", "checked_at")}
    except (OSError, ValueError, TypeError, json.JSONDecodeError, KeyError):
        pass
    return None


def route_status(*, force: bool = False) -> dict[str, Any]:
    """Claude Code first, then Codex.  Never executes either AI by itself."""
    if not force:
        cached = _cached_route()
        if cached:
            return cached
    claude = probe_agent("claude")
    if claude["state"] == "ready":
        selected, note = "claude", "Claude Code 우선"
        codex = {"agent": "codex", "state": "standby", "reason": "Claude Code 사용 가능 시 대기"}
    else:
        codex = probe_agent("codex")
    if claude["state"] != "ready" and codex["state"] == "ready":
        selected, note = "codex", "Claude Code 사용 불가 → Codex 폴백"
    elif claude["state"] != "ready":
        selected, note = "codex_pending", "Claude Code 사용 불가 → Codex 작업 요청 대기"
    route = {
        "primary": "claude",
        "selected": selected,
        "note": note,
        "agents": {"claude": claude, "codex": codex},
        "checked_at": datetime.now().isoformat(timespec="seconds"),
    }
    _atomic_json(STATUS_PATH, route)
    return route


def enqueue(task_key: str, title: str, args: list[str]) -> dict[str, Any]:
    """Persist an AI review/follow-up request and return the selected route.

    The caller still runs the deterministic local script itself. This keeps one
    writer at a time and prevents an AI retry from duplicating ledger actions.
    """
    route = route_status(force=True)
    now = datetime.now()
    request_id = f"{now:%Y%m%d_%H%M%S}_{task_key}"
    record = {
        "id": request_id,
        "created_at": now.isoformat(timespec="seconds"),
        "task_key": task_key,
        "title": title,
        "local_command": [os.path.basename(os.environ.get("PYTHON", "python")), *args],
        "primary": route["primary"],
        "selected": route["selected"],
        "route_note": route["note"],
        "status": "queued",
        "safety": "AI 요청은 검토·실패 후속조치용이며, 실행 버튼의 업무 스크립트는 로컬에서 1회만 실행합니다.",
    }
    _atomic_json(REPORT_DIR / f"{request_id}.json", record)
    _atomic_json(STATUS_PATH, {**route, "last_request": record})
    return record


def status() -> dict[str, Any]:
    """Latest availability and pending request summary for desktop/web status."""
    route = route_status()
    try:
        jobs = sorted(REPORT_DIR.glob("*.json"), reverse=True)
        latest = json.loads(jobs[0].read_text(encoding="utf-8")) if jobs else None
    except Exception:
        latest = None
    return {**route, "last_request": latest}


def route_label(info: dict[str, Any]) -> str:
    selected = (info or {}).get("selected", "codex_pending")
    if selected == "claude":
        return "Claude Code 우선"
    if selected == "codex":
        return "Claude Code 사용 불가 → Codex 폴백"
    return "Claude Code 사용 불가 → Codex 요청 대기"
