# -*- coding: utf-8 -*-
"""접수취소 동기화의 반복 실패 회귀검증.

운영 DB·밴드·관리대장을 열지 않고 메모리 가짜 저장소와 임시 캐시만 쓴다.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import cancel_resolution as resolution
import cancel_watch


class FakeStore:
    def __init__(self):
        self.work = {
            "id": "work-1",
            "kind": "돌발AS",
            "project_no": "UJ2601078",
            "status": "접수",
            "record_version": 1,
            "fields": {"진행상태": "접수", "작업완료일": ""},
        }
        self.requests = {}

    def list_work(self, *, kind=None, limit=10_000):
        return [dict(self.work)] if kind == "돌발AS" else []

    def update_work(self, work_id, *, expected_version, patch, idempotency_key, **kwargs):
        payload = (work_id, expected_version, json.dumps(patch, ensure_ascii=False, sort_keys=True))
        previous = self.requests.get(idempotency_key)
        if previous is not None and previous != payload:
            raise RuntimeError("idempotency key reused with different input")
        self.requests[idempotency_key] = payload
        if expected_version != self.work["record_version"]:
            raise RuntimeError("version conflict")
        self.work["status"] = patch["status"]
        self.work["fields"] = {**self.work["fields"], **patch["fields"]}
        self.work["record_version"] += 1
        return {"event_id": "event-%d" % self.work["record_version"]}


def hit(band="90610953", evidence="✅접수취소"):
    return {
        "프로젝트NO": "UJ2601078",
        "업무종류": "돌발AS",
        "밴드": band,
        "게시글": "5179",
        "게시일": "2026-06-22",
        "자리": "본문",
        "근거": evidence,
        "원문": evidence,
        "처리구분": "접수취소",
        "근거URL": "https://band.us/band/%s/post/5179" % band,
    }


class CancelResolutionRegressionTest(unittest.TestCase):
    def test_missing_cross_report_does_not_erase_existing_proof(self):
        store = FakeStore()
        cross = {"UJ2601078": {"카톡": "대화.txt 담당자", "카톡글": "접수취소 확인"}}
        first = resolution.sync_hits({"UJ2601078": hit()}, store=store, corroborations=cross)
        self.assertEqual(first["updated"], 1)
        second = resolution.sync_hits({"UJ2601078": hit()}, store=store, corroborations={})
        self.assertTrue(second["ok"], second)
        self.assertEqual(second["unchanged"], 1)
        self.assertIn("카톡 교차", store.work["fields"]["접수취소근거"])

    def test_evidence_a_b_a_uses_version_scoped_idempotency(self):
        store = FakeStore()
        for band in ("90610953", "밴드 홈", "90610953"):
            result = resolution.sync_hits(
                {"UJ2601078": hit(band=band)}, store=store, corroborations={}
            )
            self.assertTrue(result["ok"], result)
            self.assertEqual(result["updated"], 1)
        self.assertEqual(store.work["record_version"], 4)
        self.assertEqual(len(store.requests), 3)

    def test_numeric_cache_filename_wins_over_unstable_display_name(self):
        old_dir = cancel_watch.CACHE_DIR
        try:
            with tempfile.TemporaryDirectory() as tmp:
                cancel_watch.CACHE_DIR = tmp
                payload = {
                    "band_name": "밴드 홈",
                    "posts": {
                        "5179": {
                            "created_at": 1782085680000,
                            "content": "프로젝트NO : UJ2601078\n✅접수취소",
                            "comments": [],
                        }
                    },
                }
                with open(os.path.join(tmp, "90610953.json"), "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False)
                hits, _blind, _total = cancel_watch.scan_band(quiet=True)
                self.assertEqual(hits["UJ2601078"]["밴드"], "90610953")
                self.assertIn("/band/90610953/post/5179", hits["UJ2601078"]["근거URL"])
        finally:
            cancel_watch.CACHE_DIR = old_dir

    def test_sync_error_reason_is_written_to_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "report.md")
            with open(path, "w", encoding="utf-8") as f:
                f.write("# report\n")
            cancel_watch.append_sync_result(path, {
                "updated": 0, "unchanged": 0, "conflicts": 1,
                "ambiguous": 0, "missing": 0,
                "records": [{
                    "project": "UJ2601078", "action": "conflict",
                    "current_status": "작업완료",
                    "reason": "완료일 또는 종료상태가 있어 자동 취소하지 않음",
                }],
                "errors": [{
                    "project": "UJ2601078", "type": "IdempotencyConflict",
                    "message": "idempotency key reused with different input",
                }],
            })
            text = open(path, encoding="utf-8").read()
            self.assertIn("안전보류 1건", text)
            self.assertIn("IdempotencyConflict", text)
            self.assertIn("UJ2601078", text)


if __name__ == "__main__":
    unittest.main()
