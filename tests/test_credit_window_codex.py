# -*- coding: utf-8 -*-
"""Codex 크래딧 분리·재개 검증 — 실측 reports 를 건드리지 않는다."""
import json
import io
import os
import tempfile
import time
import sys
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

ROOT = str(Path(__file__).resolve().parents[1])
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import agent_dispatch as dispatch
import credit_window as credit


def _limit_text(when):
    day = when.day
    suffix = "th" if 10 <= day % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    stamp = when.strftime("%b {day}{suffix}, %Y %I:%M %p").format(day=day, suffix=suffix)
    return "You've hit your usage limit; try again at %s." % stamp.lstrip("0")


def main():
    now = time.time()
    future = datetime.fromtimestamp(now + 3600)
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        log = root / "one.log"
        log.write_text(_limit_text(future), encoding="utf-8")
        st = credit.codex_state(now=now, report_dir=td)
        assert st["갈래"] == "소진" and 59 <= st["남은분"] <= 60, st

        combined = credit.combined_state(now=now, dirs=[], report_dir=td)
        assert combined["agents"]["codex"]["갈래"] == "소진", combined
        assert "Codex 소진" in credit.line(combined), credit.line(combined)

        # CLI 사람 출력도 Claude 전용 state()가 아니라 두 에이전트를 함께 보여야 한다.
        shown = io.StringIO()
        with patch.object(credit, "combined_state", return_value=combined), \
             redirect_stdout(shown):
            assert credit.main([]) == 0
        assert "Codex 소진" in shown.getvalue(), shown.getvalue()

        old_dir, old_status = dispatch.REPORT_DIR, dispatch.STATUS_PATH
        dispatch.REPORT_DIR = root
        dispatch.STATUS_PATH = root / "status.json"
        try:
            def fake_probe(name):
                return {"agent": name, "state": "ready", "reason": "test"}
            with patch.object(dispatch, "probe_agent", side_effect=fake_probe):
                route = dispatch.route_status(force=True)
            assert route["selected"] == "claude", route
            assert route["agents"]["codex"]["state"] == "unavailable", route

            ticket = root / "waiting.json"
            ticket.write_text(json.dumps({"id": "waiting", "status": "waiting",
                                          "local_returncode": 0}, ensure_ascii=False),
                              encoding="utf-8")
            assert dispatch.resume_pending() == [], "소진 중인데 대기표를 보냈다"

            # 충전 시각이 지나면 같은 표를 새로 만들지 않고 기존 것을 다시 보낸다.
            past = datetime.fromtimestamp(now - 60)
            log.write_text(_limit_text(past), encoding="utf-8")
            with patch.object(dispatch, "probe_agent", side_effect=fake_probe), \
                 patch.object(dispatch, "dispatch_async", return_value=True):
                sent = dispatch.resume_pending()
            assert sent == ["waiting"], sent
        finally:
            dispatch.REPORT_DIR, dispatch.STATUS_PATH = old_dir, old_status
    print("Codex 크래딧 분리 · 실패표 비증식 · 충전 뒤 기존표 재개 OK")


if __name__ == "__main__":
    main()
