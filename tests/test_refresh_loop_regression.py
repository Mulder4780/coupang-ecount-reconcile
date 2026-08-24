import tempfile
import unittest
import json
import os
from pathlib import Path

import automation_pipeline as pipeline
from app_store import AppStore


ROOT = Path(__file__).resolve().parents[1]


class RefreshLoopRegressionTests(unittest.TestCase):
    def test_kakao_history_ignores_repeat_time_and_never_scans_canonical_share(self):
        with tempfile.TemporaryDirectory(prefix="csos-kakao-signal-") as td:
            root = Path(td)
            reports = root / "reports"
            reports.mkdir(parents=True)
            history = reports / "카톡_반영회차.json"
            history.write_text(
                json.dumps([{"시각": "one", "받은파일": ["KakaoTalk_A.txt"]}]),
                encoding="utf-8",
            )
            first = pipeline.source_signals(root)["kakao"]["fingerprint"]
            history.write_text(
                json.dumps([
                    {"시각": "two", "받은파일": ["KakaoTalk_A.txt"]},
                    {"시각": "one", "받은파일": ["KakaoTalk_A.txt"]},
                ]),
                encoding="utf-8",
            )
            second = pipeline.source_signals(root)["kakao"]["fingerprint"]
            self.assertEqual(first, second)
            history.write_text(
                json.dumps([{"시각": "three", "받은파일": ["KakaoTalk_B.txt"]}]),
                encoding="utf-8",
            )
            self.assertNotEqual(second, pipeline.source_signals(root)["kakao"]["fingerprint"])

    def test_pipeline_outputs_do_not_wake_their_own_sources(self):
        with tempfile.TemporaryDirectory(prefix="csos-refresh-signal-") as td:
            root = Path(td)
            cache = root / "band" / "cache"
            reports = root / "reports"
            inbox = root / "inbox"
            cache.mkdir(parents=True)
            reports.mkdir(parents=True)
            inbox.mkdir(parents=True)

            raw = cache / "raw_90610953.json"
            canonical = cache / "90610953.json"
            raw.write_text('{"raw":1}', encoding="utf-8")
            canonical.write_text('{"derived":1}', encoding="utf-8")
            marker = reports / "ERP판매_프로젝트색인.json"
            marker.write_text('{"src":[],"count":0,"index":{}}', encoding="utf-8")
            (reports / "download_intake.json").write_text('{"time":"one"}', encoding="utf-8")
            (reports / "upload_intake.json").write_text('{"time":"one"}', encoding="utf-8")
            (reports / "ERP원장_대조.csv").write_text("one", encoding="utf-8")

            first = pipeline.source_signals(root)
            canonical.write_text('{"derived":2}', encoding="utf-8")
            (reports / "download_intake.json").write_text('{"time":"two"}', encoding="utf-8")
            (reports / "upload_intake.json").write_text('{"time":"two"}', encoding="utf-8")
            (reports / "ERP원장_대조.csv").write_text("two", encoding="utf-8")
            second = pipeline.source_signals(root)
            self.assertEqual(first["band"]["fingerprint"], second["band"]["fingerprint"])
            self.assertEqual(first["erp"]["fingerprint"], second["erp"]["fingerprint"])

            raw.write_text('{"raw":2}', encoding="utf-8")
            self.assertNotEqual(
                second["band"]["fingerprint"],
                pipeline.source_signals(root)["band"]["fingerprint"],
            )
            erp_input = inbox / "new.xlsx"
            erp_input.write_bytes(b"new-erp-input")
            self.assertNotEqual(
                second["erp"]["fingerprint"],
                pipeline.source_signals(root)["erp"]["fingerprint"],
            )

    def test_success_records_the_post_command_signal(self):
        with tempfile.TemporaryDirectory(prefix="csos-refresh-settle-") as td:
            root = Path(td)
            raw = root / "band" / "cache" / "raw_90610953.json"
            raw.parent.mkdir(parents=True)
            (root / "reports").mkdir(parents=True)
            raw.write_text('{"raw":1}', encoding="utf-8")
            before = pipeline.source_signals(root)["band"]

            def runner(name, _args, _timeout):
                raw.write_text('{"raw":2,"written_by":"collector"}', encoding="utf-8")
                return {"name": name, "ok": True, "returncode": 0, "timed_out": False}

            worker = pipeline.AutomationPipeline(
                root=root,
                state_path=root / "reports" / "state.json",
                lock_path=root / "reports" / ".lock",
                store=AppStore(root / "db" / "app.db").initialize(),
                stage_runner=runner,
            )
            self.assertTrue(worker._acquire_run())
            worker.run_record = {"status": "running", "run_id": worker._run_id, "stages": []}
            worker.state["running"] = True
            worker.state["active_run_id"] = worker._run_id
            worker.state["lock_token"] = worker._lock_token
            try:
                self.assertTrue(worker._run_source("band", before, [("collector", [], 10)]))
                after = pipeline.source_signals(root)["band"]
                self.assertEqual(
                    worker.state["sources"]["band"]["fingerprint"], after["fingerprint"]
                )
                self.assertNotEqual(before["fingerprint"], after["fingerprint"])
            finally:
                worker._release_run()

    def test_daily_reconciliation_has_priority_over_incremental_pipeline(self):
        with tempfile.TemporaryDirectory(prefix="csos-daily-priority-") as td:
            root = Path(td)
            reports = root / "reports"
            reports.mkdir(parents=True)
            identity = pipeline.pid_alive.identity()
            (reports / ".daily_run.lock").write_text(
                json.dumps(
                    {
                        "pid": os.getpid(),
                        "pid_started_at": identity.get("pid_started_at"),
                        "token": "test",
                        "started_at": "2026-08-24T12:00:00+09:00",
                    }
                ),
                encoding="utf-8",
            )
            calls = []
            worker = pipeline.AutomationPipeline(
                root=root,
                state_path=reports / "state.json",
                lock_path=reports / ".pipeline.lock",
                store=AppStore(root / "db" / "app.db").initialize(),
                stage_runner=lambda *args: calls.append(args),
            )
            result = worker.run_once(trigger="test-daily-priority", force=True)
            self.assertEqual(result["status"], "deferred_daily")
            self.assertEqual(calls, [])
            self.assertFalse((reports / ".pipeline.lock").exists())

    def test_run_view_does_not_show_or_fetch_dashboard_heavy_sections(self):
        html = (ROOT / "webapp" / "index.html").read_text(encoding="utf-8")
        mapping = html.split("function liveDataKeysForView(view){", 1)[1].split(
            "function liveViewIsCurrent", 1
        )[0]
        refresh = html.split("async function refreshLiveViewData", 1)[1].split(
            "async function retryDataSection", 1
        )[0]
        health = html.split("function renderDataHealth(){", 1)[1].split(
            "function sectionOfMode", 1
        )[0]

        self.assertIn("dash:['settlements','works','issues','representative']", mapping)
        self.assertIn("run:[]", mapping)
        self.assertNotIn("dash:Object.keys(DATA_SECTION_DEFS)", mapping)
        self.assertIn("liveDataKeysForView(view).filter", refresh)
        self.assertNotIn("view==='dash'?Object.keys(DATA_SECTION_DEFS)", refresh)
        self.assertIn("relevant.has(s.key)", health)
        self.assertIn("!inRun&&LIVE_SYNC.refreshFlight", health)


if __name__ == "__main__":
    unittest.main()
