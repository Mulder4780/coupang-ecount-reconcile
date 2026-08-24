# -*- coding: utf-8 -*-
import json
import os
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from webapp import app_server as A


class CalendarLiveMasterMetaTest(unittest.TestCase):
    def test_live_ledger_label_replaces_old_gcal_label(self):
        with tempfile.TemporaryDirectory() as td:
            os.makedirs(os.path.join(td, "reports"), exist_ok=True)
            with open(os.path.join(td, "reports", "gcal_events.json"), "w",
                      encoding="utf-8") as fh:
                json.dump({"갱신": "2026-08-23T14:30:06", "관리대장": 614,
                           "일정": [], "원천": []}, fh, ensure_ascii=False)
            old_root, old_mt = A.ROOT, dict(A._MT)
            old_events, old_cut = A._calendar_work_events, A.band_collect_cutoff
            try:
                A.ROOT = td
                A._MT.update({"at": time.time(), "v": 1787536800.0,
                              "path": os.path.join(td,
                                  "쿠팡_통합업무_일일보고_관리대장_v616.xlsx")})
                A._calendar_work_events = lambda: []
                A.band_collect_cutoff = lambda: {"읽음": True, "밀림": False}
                got = A.get_calendar()
            finally:
                A.ROOT = old_root
                A._MT.clear()
                A._MT.update(old_mt)
                A._calendar_work_events, A.band_collect_cutoff = old_events, old_cut
            self.assertEqual(got["관리대장"], 616)
            self.assertEqual(got["달력원천갱신"], "2026-08-23T14:30:06")
            self.assertTrue(got["업무원장갱신"])


if __name__ == "__main__":
    unittest.main()
