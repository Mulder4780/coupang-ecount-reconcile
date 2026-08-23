"""autopilot 상태 조회가 윈도우 CP949 콘솔에서도 죽지 않는지 확인한다."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "cp949"
    result = subprocess.run(
        [sys.executable, str(ROOT / "autopilot.py"), "--status"],
        cwd=ROOT,
        env=env,
        capture_output=True,
        encoding="utf-8",
        errors="strict",
        timeout=30,
        check=False,
    )
    assert result.returncode == 0, result.stderr
    value = json.loads(result.stdout)
    assert isinstance(value, dict), value

    # pythonw처럼 출력 통로가 없는 백그라운드 실행도 정상 종료해야 한다.
    background = subprocess.run(
        [
            sys.executable,
            "-c",
            "import autopilot,sys;sys.stdout=None;sys.stderr=None;"
            "raise SystemExit(autopilot.main(['--status']))",
        ],
        cwd=ROOT,
        timeout=30,
        check=False,
    )
    assert background.returncode == 0, background.returncode
    print("autopilot CP949 · pythonw 상태 조회: OK")


if __name__ == "__main__":
    main()
