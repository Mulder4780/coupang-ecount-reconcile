# -*- coding: utf-8 -*-
"""convert_dump 파일별 증분 상태 전용 회귀검증.

대형 synthetic_check에 끼워 넣지 않는다. 변환기 상태·잠금·redirect 안전조건만 작은
임시 폴더에서 재현해, 실데이터와 서버를 건드리지 않고 빠르게 돌린다.
"""
import json
import os
import sys
import tempfile
import unittest
from contextlib import contextmanager
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import band.convert_dump as cd
import pid_alive
import source_dirs


class ConvertDumpIncrementalTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = self.tmp.name
        self.cache = os.path.join(self.root, "cache")
        self.reports = os.path.join(self.root, "reports")
        self.band_root = os.path.join(self.root, "band_source")
        self.z_dumps = os.path.join(self.band_root, "브라우저덤프", "2026-08-24")
        os.makedirs(self.cache)
        os.makedirs(self.reports)
        os.makedirs(self.z_dumps)
        os.makedirs(os.path.join(self.band_root, "수집본"))

    def tearDown(self):
        self.tmp.cleanup()

    @contextmanager
    def isolated(self, version="v1"):
        attrs = {
            "CACHE": self.cache,
            "STATE": os.path.join(self.reports, "밴드덤프_변환상태.json"),
            "LOCK": os.path.join(self.cache, ".convert_dump.lock"),
            "CHANGED_LOG": os.path.join(self.reports, "밴드_수정글.json"),
            "PROBE_LOG": os.path.join(self.reports, "밴드_확인시각.json"),
        }
        with mock.patch.multiple(cd, **attrs), \
             mock.patch.object(cd, "converter_version", return_value=version), \
             mock.patch.object(source_dirs, "BAND_DIR", self.band_root):
            yield attrs

    def dump(self, name, cap, post_no="1", content="정상 본문", notime=None):
        path = os.path.join(self.z_dumps, name)
        doc = {
            "band": "90610953", "name": "테스트 밴드", "capturedAt": cap,
            "posts": {str(post_no): {
                "created_at": cap - 1000, "author": "담당자", "content": content,
                "comments": [], "comments_full": True,
            }},
            "notime": notime or {}, "missing": [],
        }
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False)
        return path

    def state(self):
        return json.load(open(os.path.join(self.reports, "밴드덤프_변환상태.json"),
                              encoding="utf-8"))

    def cache_doc(self):
        return json.load(open(os.path.join(self.cache, "90610953.json"), encoding="utf-8"))

    def test_unchanged_files_are_skipped_and_new_file_is_merged_first(self):
        first = self.dump("90610953_first.json", 1_800_000_000_000, "1", "첫 글")
        with self.isolated("v1"):
            self.assertEqual(cd.main(budget_sec=30, lock_wait_sec=0), 0)
            before = open(os.path.join(self.cache, "90610953.json"), "rb").read()
            self.assertEqual(cd.main(budget_sec=30, lock_wait_sec=0), 0)
            self.assertEqual(open(os.path.join(self.cache, "90610953.json"), "rb").read(), before)

            second = self.dump("90610953_second.json", 1_800_000_010_000, "2", "둘째 글")
            self.assertEqual(cd.main(budget_sec=0, lock_wait_sec=0), 0)
            self.assertEqual(set(self.cache_doc()["posts"]), {"1", "2"})
            state = self.state()
            self.assertEqual(state["completed_version"], "v1")
            self.assertEqual(state["files"][cd._norm_path(first)]["converter_version"], "v1")
            self.assertEqual(state["files"][cd._norm_path(second)]["converter_version"], "v1")

    def test_code_version_replay_obeys_budget_but_fresh_dump_does_not_wait(self):
        old1 = self.dump("90610953_old1.json", 1_800_000_000_000, "1", "옛 글 1")
        old2 = self.dump("90610953_old2.json", 1_800_000_001_000, "2", "옛 글 2")
        with self.isolated("v1"):
            cd.main(budget_sec=30, lock_wait_sec=0)

        fresh = self.dump("90610953_fresh.json", 1_800_000_020_000, "3", "새 글")
        with self.isolated("v2"):
            cd.main(budget_sec=0, lock_wait_sec=0)
            state = self.state()
            self.assertEqual(state["files"][cd._norm_path(fresh)]["converter_version"], "v2")
            self.assertEqual(state["files"][cd._norm_path(old1)]["converter_version"], "v1")
            self.assertEqual(state["files"][cd._norm_path(old2)]["converter_version"], "v1")
            self.assertEqual(state["completed_version"], "v1")
            self.assertIn("3", self.cache_doc()["posts"])

            cd.main(budget_sec=30, lock_wait_sec=0)
            state = self.state()
            self.assertEqual(state["completed_version"], "v2")
            self.assertTrue(all(e["converter_version"] == "v2"
                                for e in state["files"].values()))

    def test_redirect_tombstone_waits_for_two_complete_rounds(self):
        sigs = {"99": "same-feed", "100": "same-feed"}
        self.dump("90610953_round1.json", 1_800_000_000_000, "1", "정상", sigs)
        with self.isolated("v1"):
            cd.main(budget_sec=30, lock_wait_sec=0)
            self.assertNotIn("99", self.cache_doc()["posts"])

            self.dump("90610953_round2.json", 1_800_000_010_000, "2", "두 번째 정상", sigs)
            cd.main(budget_sec=30, lock_wait_sec=0)
            posts = self.cache_doc()["posts"]
            self.assertTrue(posts["99"]["deleted"])
            self.assertEqual(posts["99"]["deleted_by"], "redirect")

    def test_state_does_not_advance_when_cache_swap_fails(self):
        path = self.dump("90610953_fail.json", 1_800_000_000_000)
        with self.isolated("v1"):
            real_swap = cd.swap_in

            def fail_cache(tmp, dst, *args, **kwargs):
                if cd._norm_path(dst) == cd._norm_path(os.path.join(self.cache, "90610953.json")):
                    raise RuntimeError("simulated cache failure")
                return real_swap(tmp, dst, *args, **kwargs)

            with mock.patch.object(cd, "swap_in", side_effect=fail_cache):
                with self.assertRaisesRegex(RuntimeError, "simulated cache failure"):
                    cd.main(budget_sec=30, lock_wait_sec=0)
            entry = self.state()["files"][cd._norm_path(path)]
            self.assertEqual(entry["status"], "pending")
            self.assertEqual(entry["converter_version"], "")

    def test_unavailable_root_is_not_pruned_and_live_lock_is_not_reclaimed(self):
        missing_root = cd._norm_path(os.path.join(self.root, "disconnected-z"))
        missing_file = cd._norm_path(os.path.join(missing_root, "old.json"))
        state = {"files": {missing_file: {
            "path": missing_file, "root": missing_root, "source_kind": "z_dump"}}}
        cd._prune_state_files(state, {"files": {}, "complete_roots": set(),
                                      "unavailable_roots": {missing_root}})
        self.assertIn(missing_file, state["files"])
        cd._prune_state_files(state, {"files": {}, "complete_roots": {missing_root},
                                      "unavailable_roots": set()})
        self.assertNotIn(missing_file, state["files"])

        with self.isolated("v1") as attrs:
            with open(attrs["LOCK"], "w", encoding="utf-8") as fh:
                fh.write(f"{os.getpid()} {pid_alive.stamp()} 2026-08-24T12:00:00\n")
            self.assertFalse(cd._lock_acquire(wait_sec=0))
            self.assertTrue(os.path.isfile(attrs["LOCK"]))


if __name__ == "__main__":
    unittest.main()
