# -*- coding: utf-8 -*-
"""Windows에서도 제한 시간이 실제로 끝나는 자식 프로세스 실행기.

``subprocess.run(timeout=...)``은 Windows에서 자식이 SMB 대기에 걸리면 kill 뒤의
무제한 ``communicate()``에서 다시 멈출 수 있다. 자동화 회차와 AI CLI가 같은 함정에
빠지지 않도록 생성·나무 종료·제한된 드레인을 이 한 곳에서 수행한다.
"""
from __future__ import annotations

import os
import subprocess
from dataclasses import dataclass
from typing import Mapping, Sequence


@dataclass(frozen=True)
class ProcessResult:
    returncode: int
    stdout: str = ""
    stderr: str = ""
    timed_out: bool = False
    stuck_pid: int = 0


def _kill_tree(process: subprocess.Popen) -> None:
    if os.name == "nt":
        try:
            # taskkill 자체도 무한정 기다리지 않는다. 이 호출은 로컬 프로세스 표만 본다.
            killer = subprocess.Popen(
                ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            try:
                killer.wait(timeout=20)
            except subprocess.TimeoutExpired:
                killer.kill()
        except OSError:
            pass
    try:
        process.kill()
    except OSError:
        pass


def run_tree(
    command: Sequence[str],
    *,
    cwd: str | os.PathLike | None = None,
    env: Mapping[str, str] | None = None,
    timeout: int | float = 600,
    drain_timeout: int | float = 30,
    output_limit: int = 200_000,
) -> ProcessResult:
    """명령을 실행하고 시간 초과 시 자식 나무를 끊은 뒤 반드시 반환한다."""
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        list(command),
        cwd=cwd,
        env=dict(env) if env is not None else None,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=flags,
    )
    try:
        stdout, stderr = process.communicate(timeout=timeout)
        return ProcessResult(
            returncode=int(process.returncode or 0),
            stdout=(stdout or "")[-output_limit:],
            stderr=(stderr or "")[-output_limit:],
        )
    except subprocess.TimeoutExpired:
        _kill_tree(process)
        try:
            stdout, stderr = process.communicate(timeout=drain_timeout)
            stuck = 0
        except subprocess.TimeoutExpired:
            stdout, stderr, stuck = "", "", process.pid
        return ProcessResult(
            returncode=-9,
            stdout=(stdout or "")[-output_limit:],
            stderr=(stderr or "")[-output_limit:],
            timed_out=True,
            stuck_pid=stuck,
        )
