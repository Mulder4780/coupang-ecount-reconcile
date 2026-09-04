"""[361] Intentional Band stop suppresses requests, never evidence or ERP faults.

Run from ecount: python -m unittest discover -s tests -p test_band_collection_stop.py -v
All files and input states are isolated; no browser, workbook, or real DB writes.
"""
import copy
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from contextlib import ExitStack, redirect_stdout
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from band import collect_switch as CS
from band import browser_chain as BC
from band import recollect as RC
from band import userscript_watch as UW
import schedule_watch as SW
import system_audit as SA


class BandCollectionStopTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.root = Path(self.stack.enter_context(tempfile.TemporaryDirectory(prefix="band-stop-")))
        self.reports = self.root / "reports"
        self.reports.mkdir()
        self.stack.enter_context(patch.dict(os.environ))
        os.environ.pop(CS.ENV, None)
        self.mark = self.reports / "stop.json"
        self.stack.enter_context(patch.object(CS, "MARK", str(self.mark)))
        self.stack.enter_context(patch.object(SW, "ROOT", str(self.root)))
        self.stack.enter_context(patch.object(BC, "STATE", str(self.reports / "chain.json")))
        self.stack.enter_context(patch.object(RC, "STATE", str(self.reports / "recollect.json")))
        self.stack.enter_context(patch.object(SA, "REPORTS", self.reports))
        self.stack.enter_context(patch.object(UW, "OUT", str(self.reports / "userscript.md")))
        self.mark.write_text(json.dumps({"중단": True, "왜": "사용자 지시"}), encoding="utf-8")

    def trace(self, late):
        return {"작업": "CSOS_BrowserChain", "갈래": "기회놓침", "시각": "2026-09-03",
                "늦은것": late, "무엇": "old summary", "어떻게": "manual erp"}

    def test_stopped_watch_never_probes_chrome_or_reports(self):
        with patch.object(UW, "chrome_side", side_effect=AssertionError("Chrome touched")), \
             patch.object(UW, "load_reports", side_effect=AssertionError("old report used")):
            state = UW.check()
            self.assertEqual(state["갈래"], "수집중단")
            self.assertEqual(UW.lines(), [])
        text = Path(UW.OUT).read_text(encoding="utf-8")
        self.assertIn("사용자 지시", text)
        self.assertIn("collect_switch.py --resume", text)
        self.assertIn("흡수·대조", text)
        self.assertNotIn("F12", text)

    def test_resume_immediately_restores_alarm(self):
        with patch.object(UW, "load_reports", return_value=(None, "안 왔다")), \
             patch.object(UW, "plan_state", return_value={"있음": False}):
            self.assertEqual(UW.lines(), [])
            CS.resume()
            self.assertTrue(UW.lines())
            self.assertEqual(UW.current_state()["갈래"], "안옴")

    def test_missing_invalid_false_marker_never_suppresses(self):
        for raw in (None, "{broken", "[]", '{"중단":false}'):
            with self.subTest(raw=raw):
                if raw is None:
                    self.mark.unlink(missing_ok=True)
                else:
                    self.mark.write_text(raw, encoding="utf-8")
                self.assertFalse(CS.warning_status()["수집중단"])
                with patch.object(UW, "load_reports", return_value=(None, "안 왔다")), \
                     patch.object(UW, "plan_state", return_value={}):
                    self.assertTrue(UW.lines())

    def test_environment_override_and_policy_exception(self):
        os.environ[CS.ENV] = "1"
        self.assertFalse(CS.warning_status()["수집중단"])
        os.environ[CS.ENV] = "0"
        self.assertTrue(CS.warning_status()["수집중단"])
        with patch.object(CS, "stopped", side_effect=OSError("unreadable")):
            self.assertFalse(CS.warning_status()["수집중단"])

    def test_old_band_trace_hidden_without_deleting_evidence(self):
        trace = self.trace(["밴드 댓글 84789192 22일째"])
        path = self.reports / "브라우저수집_오류.json"
        path.write_text(json.dumps(trace, ensure_ascii=False), encoding="utf-8")
        before = path.read_bytes()
        self.assertEqual(SW.traces(), [])
        self.assertEqual(path.read_bytes(), before)
        CS.resume()
        self.assertEqual(len(SW.traces()), 1)
        self.assertEqual(path.read_bytes(), before)

    def test_mixed_trace_preserves_erp_fault(self):
        trace = self.trace(["밴드 댓글 84789192 22일째", "ERP 전화면 몰이 12일째"])
        before = copy.deepcopy(trace)
        filtered = BC.visible_missed_trace(trace)
        self.assertEqual(filtered["늦은것"], ["ERP 전화면 몰이 12일째"])
        self.assertIn("ERP", filtered["무엇"])
        self.assertNotIn("밴드", filtered["무엇"])
        self.assertEqual(trace, before)

    def test_unknown_shape_or_other_failure_is_not_hidden(self):
        for late in (None, [], [None], ["알 수 없는 작업 12일째"], "밴드 12일째"):
            trace = self.trace(late)
            self.assertEqual(BC.visible_missed_trace(trace), trace)
        trace = self.trace(["밴드 댓글 84789192 22일째"])
        trace["갈래"] = "흡수실패"
        self.assertEqual(BC.visible_missed_trace(trace), trace)
        other = self.reports / "ERP_오류.json"
        other.write_text(json.dumps({"무엇": "ERP 요청 오류"}), encoding="utf-8")
        self.assertEqual(len(SW.traces()), 1)

    def audit(self):
        recollect = {"손볼것": ["밴드 붙여넣기 요구"],
                     "최근변경": {"바뀐글": ["84789192/1"]}, "확인함": False}
        def read(path):
            return recollect if Path(path).name == "밴드_재수집.json" else None
        with patch.object(SA, "_read_json", side_effect=read):
            return SA.build()

    def test_system_audit_retains_changed_evidence(self):
        report = self.audit()
        ids = {x["id"] for x in report["findings"]}
        self.assertNotIn("band-recollect-pending", ids)
        self.assertIn("band-changes-unacknowledged", ids)
        policy = report["sources"]["band_recollect"]["collection"]
        self.assertTrue(policy["수집중단"])
        self.assertIn("--resume", policy["재개"])
        CS.resume()
        self.assertIn("band-recollect-pending", {x["id"] for x in self.audit()["findings"]})

    def test_recollect_still_absorbs_and_retains_unacknowledged_changes(self):
        old = {"최근변경": {"바뀐글": ["84789192/1"], "되돌아감": []}, "확인함": False}
        got = {"신규": 0, "변경": 0, "그대로": 1, "버림": 0,
               "새글": [], "바뀐글": [], "변경상세": []}
        with patch.object(RC, "absorb", return_value=[]) as absorb, \
             patch.object(RC, "ingest", return_value=got) as ingest, \
             patch.object(RC, "plan", return_value=({"84789192": {"번호": [1], "수": 1, "넘침": 0}}, "2026-08-05")), \
             patch.object(RC, "_not_recollected", return_value=[1]), \
             patch.object(RC, "load_state", return_value=old), \
             patch.object(RC, "save_state"), patch.object(RC, "_log"), \
             patch.object(RC, "make_paste") as paste:
            state = RC.run()
            absorb.assert_called_once()
            ingest.assert_called_once()
            paste.assert_not_called()
            self.assertEqual(state["손볼것"], [])
            self.assertEqual(state["대상"]["84789192"]["다시받을것"], 1)
            self.assertEqual(state["최근변경"], old["최근변경"])
            self.assertFalse(state["확인함"])
            CS.resume()
            paste.return_value = (str(self.root / "paste.js"), "안내")
            self.assertTrue(RC.run()["손볼것"])
            paste.assert_called_once()

    def test_old_recollect_print_no_longer_requests_paste(self):
        state = {"손볼것": ["F12 붙여넣기 요구"], "대상": {
            "84789192": {"창안글": 1, "다시받을것": 1, "붙여넣기": "old.js"}}}
        output = io.StringIO()
        with redirect_stdout(output):
            RC.show(state)
        text = output.getvalue()
        self.assertIn("중단", text)
        self.assertIn("--resume", text)
        self.assertNotIn("F12", text)
        self.assertNotIn("old.js", text)
        self.assertEqual(state["손볼것"], ["F12 붙여넣기 요구"])


if __name__ == "__main__":
    unittest.main()
