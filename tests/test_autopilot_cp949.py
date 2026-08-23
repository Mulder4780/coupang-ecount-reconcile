"""autopilot 상태 조회가 윈도우 CP949 콘솔에서도 죽지 않는지 확인한다."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import proc_guard  # noqa: E402  (ROOT 를 경로에 넣은 뒤라야 보인다)


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
        # 창을 띄우지 않는다(`[272]`) — 이 파일은 09:50 회차가 pythonw 로 돌리므로
        # 깃발이 없으면 그때마다 검은 창이 뜬다. 잡는 출력은 그대로다(capture_output).
        **proc_guard.background_popen_kwargs(),
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
        # 출력을 잡아 둔다 — 안 잡으면 창을 없애는 순간 **아무 데도 안 남고**(`[248]`)
        # 실패했을 때 무엇이 나왔는지 말할 수 없다.
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        timeout=30,
        check=False,
        **proc_guard.background_popen_kwargs(),
    )
    assert background.returncode == 0, (background.returncode, background.stderr)
    print("autopilot CP949 · pythonw 상태 조회: OK")


if __name__ == "__main__":
    main()
