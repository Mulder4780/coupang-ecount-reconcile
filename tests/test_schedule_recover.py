# -*- coding: utf-8 -*-
import os
import sys
import tempfile
import unittest
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import schedule_recover as S


class ScheduleRecoverTest(unittest.TestCase):
    def setUp(self):
        self.now = datetime(2026, 8, 24, 11, 30)
        self.rows = [
            {"작업": "쿠팡업무_일일자동대조", "갈래": "안돎", "상태": "Ready",
             "예정": "2026-08-24T09:50:00"},
            {"작업": "쿠팡업무_원본자료자동정리", "갈래": "안돎", "상태": "Ready",
             "예정": "2026-08-24T09:35:00"},
            {"작업": "남의작업", "갈래": "안돎", "상태": "Ready",
             "예정": "2026-08-24T08:00:00"},
        ]

    def test_one_at_a_time_in_source_then_reconcile_order(self):
        self.assertEqual(S.pick(self.rows, {}, self.now)["작업"],
                         "쿠팡업무_원본자료자동정리")
        state = {"시도": {"쿠팡업무_원본자료자동정리": {
            "예정": "2026-08-24T09:35:00", "요청성공": True,
            "시도시각": self.now.isoformat()}}}
        self.assertIsNone(S.pick(self.rows, state, self.now))
        # 다음 schedule_watch가 원본정리 완료를 확인해 목록에서 내린 뒤에만
        # 일일대조를 시작한다.
        self.assertEqual(S.pick(self.rows[:1], state, self.now)["작업"],
                         "쿠팡업무_일일자동대조")

    def test_running_future_and_recent_failure_are_not_started(self):
        rows = [dict(self.rows[0], 상태="Running"),
                dict(self.rows[1], 예정="2026-08-24T12:35:00")]
        self.assertIsNone(S.pick(rows, {}, self.now))
        state = {"시도": {"쿠팡업무_원본자료자동정리": {
            "예정": "2026-08-24T09:35:00", "요청성공": False,
            "시도시각": (self.now - timedelta(minutes=5)).isoformat()}}}
        self.assertIsNone(S.pick(self.rows, state, self.now))

    def test_run_records_request_but_not_completion(self):
        with tempfile.TemporaryDirectory() as td:
            old = S.STATE
            S.STATE = os.path.join(td, "state.json")
            try:
                calls = []
                msg = S.run(rows=self.rows, now=self.now,
                            trigger=lambda name: (calls.append(name) is None, "queued"),
                            pipeline_busy=lambda: False)
                self.assertEqual(calls, ["쿠팡업무_원본자료자동정리"])
                self.assertIn("시작 요청", msg)
                self.assertIn("실제 완료는 다음 감시", msg)
            finally:
                S.STATE = old

    def test_live_pipeline_blocks_heavy_recovery(self):
        with tempfile.TemporaryDirectory() as td:
            old = S.STATE
            S.STATE = os.path.join(td, "state.json")
            try:
                calls = []
                msg = S.run(
                    rows=self.rows,
                    now=self.now,
                    trigger=lambda name: (calls.append(name) is None, "queued"),
                    pipeline_busy=lambda: True,
                )
                self.assertEqual(calls, [])
                self.assertIn("자료 갱신이 끝난 뒤", msg)
            finally:
                S.STATE = old


if __name__ == "__main__":
    unittest.main()
