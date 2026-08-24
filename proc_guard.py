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


def background_popen_kwargs() -> dict:
    """Windows child-process settings that can never create or activate a console.

    ``CREATE_NO_WINDOW`` is the primary boundary for console programs.  ``SW_HIDE``
    is intentionally applied as a second boundary because launchers and packaged
    executables do not all honour console creation flags in the same way.  Keeping
    this in one function prevents Claude/Codex and ordinary scheduled workers from
    slowly growing different popup behaviour.
    """
    if os.name != "nt":
        return {}
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 0)
    startup.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startup,
    }


CHILD_IO_ENCODING = "utf-8"


def _child_env(env: Mapping[str, str] | None) -> dict:
    """자식 파이썬이 **UTF-8 로 쓰게** 한다 — 읽는 쪽이 UTF-8 이기 때문이다.

    2026-08-24 실사고: `run_tree` 는 `encoding="utf-8"` 로 **읽는데** 자식 env 에
    아무것도 안 줘서, 윈도우 자식 파이썬이 `locale.getpreferredencoding()`
    (실측 **cp949**)로 파이프에 **썼다**. 그래서 한글 출력이 통째로 깨졌다 —
    자율복구 대기열의 조치 문구가 인계 문서에 이렇게 박혔다:
    `?????? . Self-Customizing . ... . [IP???] . ? ??? IP ??`.
    **실패는 적혔는데 사람이 못 읽는다** — [289]·[292]·[365] 가 세운
    '왜인지 말한다'가 마지막 한 걸음에서 무너지는 자리다.

    ★ **읽는 쪽과 쓰는 쪽을 한 곳에서 정한다**([162]). 실측 2026-08-24:
      `run_tree` 호출 52곳 중 env 를 챙긴 곳은 14곳뿐이라 **38곳이 새고 있었다** —
      부르는 쪽이 각자 기억해 붙이는 구조는 새 호출자가 생길 때마다 또 샌다.
    ★ **부르는 쪽이 이미 정했으면 안 덮는다**([172]) — `setdefault` 다.
    ★ **파이썬이 아닌 자식**(git·node·taskkill)은 이 값을 그냥 무시한다. 해가 없다.
    ★ `errors` 는 안 붙인다 — UTF-8 은 못 쓰는 글자가 없어 cp949 보다 언제나
      덜 죽고, 이미 이 값을 쓰던 14곳이 전부 `"utf-8"` 이라 사본을 만들지 않는다.
    """
    base = dict(env) if env is not None else dict(os.environ)
    base.setdefault("PYTHONIOENCODING", CHILD_IO_ENCODING)
    return base


def _kill_tree(process: subprocess.Popen) -> None:
    if os.name == "nt":
        try:
            # taskkill 자체도 무한정 기다리지 않는다. 이 호출은 로컬 프로세스 표만 본다.
            killer = subprocess.Popen(
                ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                **background_popen_kwargs(),
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
    process = subprocess.Popen(
        list(command),
        cwd=cwd,
        env=_child_env(env),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding=CHILD_IO_ENCODING,
        errors="replace",
        **background_popen_kwargs(),
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
