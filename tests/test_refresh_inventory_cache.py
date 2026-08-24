# -*- coding: utf-8 -*-
import json
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import billing_fill as B
import band_extract as Band
import erp_sales_index as S
import fill_erp_status as F
import inbox_scan as I
import kakao_extract as K


class RefreshInventoryCacheTests(unittest.TestCase):
    def test_recent_inventory_is_reused_without_statting_each_source(self):
        with tempfile.TemporaryDirectory(prefix="csos-inventory-") as td:
            root = Path(td) / "source"
            root.mkdir()
            book = root / "판매조회.xlsx"
            book.write_bytes(b"not-opened-by-inventory")
            cache = Path(td) / "inbox_classify.json"
            cache.write_text(json.dumps({
                str(book): [book.stat().st_size, int(book.stat().st_mtime),
                            I.RULES_VERSION, "sales"]
            }), encoding="utf-8")
            old_file, old_disk = I._CLS_FILE, I._CLS_DISK
            try:
                I._CLS_FILE, I._CLS_DISK = str(cache), None
                real_stat = os.stat
                def guarded_stat(path, *args, **kwargs):
                    if os.path.abspath(path) == os.path.abspath(book):
                        raise AssertionError("source stat")
                    return real_stat(path, *args, **kwargs)
                with mock.patch.object(I.os, "stat", side_effect=guarded_stat):
                    rows = I.cached_inventory([str(root)], max_age_s=60)
                self.assertEqual("sales", rows[0]["kind"])
                self.assertEqual(str(book), rows[0]["path"])
            finally:
                I._CLS_FILE, I._CLS_DISK = old_file, old_disk

    def test_old_or_wrong_rule_inventory_forces_safe_full_scan(self):
        with tempfile.TemporaryDirectory(prefix="csos-inventory-stale-") as td:
            root = Path(td) / "source"
            root.mkdir()
            book = root / "판매조회.xlsx"
            book.write_bytes(b"x")
            cache = Path(td) / "inbox_classify.json"
            cache.write_text(json.dumps({
                str(book): [1, int(book.stat().st_mtime), I.RULES_VERSION - 1, "sales"]
            }), encoding="utf-8")
            old_file, old_disk = I._CLS_FILE, I._CLS_DISK
            try:
                I._CLS_FILE, I._CLS_DISK = str(cache), None
                self.assertIsNone(I.cached_inventory([str(root)], max_age_s=60))
                cache.write_text(json.dumps({
                    str(book): [1, int(book.stat().st_mtime), I.RULES_VERSION, "sales"]
                }), encoding="utf-8")
                I._CLS_DISK = None
                os.utime(cache, (time.time() - 120, time.time() - 120))
                self.assertIsNone(I.cached_inventory([str(root)], max_age_s=60))
            finally:
                I._CLS_FILE, I._CLS_DISK = old_file, old_disk

    def test_content_hash_is_reused_but_duplicate_protection_remains(self):
        with tempfile.TemporaryDirectory(prefix="csos-hash-cache-") as td:
            root = Path(td)
            a, b, c = root / "a.xlsx", root / "b.xlsx", root / "c.xlsx"
            a.write_bytes(b"same")
            b.write_bytes(b"same")
            c.write_bytes(b"other")
            cache = root / "hashes.json"
            first = B.dedupe_files([str(a), str(b), str(c)], cache_path=str(cache))
            self.assertEqual([str(a), str(c)], first)

            with mock.patch.object(B.hashlib, "sha256",
                                   side_effect=AssertionError("content rehashed")):
                second = B.dedupe_files([str(a), str(b), str(c)], cache_path=str(cache))
            self.assertEqual(first, second)

    def test_erp_project_index_reuses_only_the_exact_input_fingerprint(self):
        with tempfile.TemporaryDirectory(prefix="csos-erp-index-") as td:
            out = Path(td) / "ERP판매_프로젝트색인.json"
            out.write_text(json.dumps({
                "fingerprint": "exact", "src": ["판매.xlsx"],
                "index": {"UJ2600001": {"supply": 10}}
            }), encoding="utf-8")
            with mock.patch.object(S, "OUT", str(out)), \
                    mock.patch.object(S, "sales_candidate_paths", return_value=["판매.xlsx"]), \
                    mock.patch.object(S, "_candidate_stamp", return_value="exact"), \
                    mock.patch.object(S, "sales_exports",
                                      side_effect=AssertionError("workbook reopened")):
                index, sources = S.build()
            self.assertEqual(10, index["UJ2600001"]["supply"])
            self.assertEqual(["판매.xlsx"], sources)

    def test_erp_status_reuses_the_shared_project_index(self):
        shared = {"UJ2600001": {"state": "7.수금완료"}}
        with mock.patch("erp_sales_index.build", return_value=(shared, ["판매.xlsx"])), \
                mock.patch("glob.glob", side_effect=AssertionError("ERP folder rescanned")), \
                mock.patch("billing_fill.dedupe_files",
                           side_effect=AssertionError("sales workbooks rehashed")):
            projects, sources = F.erp_projects()
        self.assertEqual("7.수금완료", projects["UJ2600001"])
        self.assertEqual([("판매.xlsx", "공용색인", 0)], sources)

    def test_kakao_selector_accepts_early_history_instead_of_falling_back_to_all_files(self):
        selected = ["latest-as.txt", "latest-pm.txt", "early-as.txt", "early-pm.txt"]
        with mock.patch.object(Band, "kakao_source_paths", return_value=selected), \
                mock.patch.object(K.glob, "glob",
                                  side_effect=AssertionError("full recursive fallback")):
            self.assertEqual(selected, K.source_paths())

    def test_kakao_structured_rows_cache_rehydrates_dates_without_reparsing(self):
        with tempfile.TemporaryDirectory(prefix="csos-kakao-rows-") as td:
            old_cache = K.EXTRACT_CACHE
            K.EXTRACT_CACHE = str(Path(td) / "rows.json")
            try:
                row = {"프로젝트NO": "UJ2600001", "예정일": K.date(2026, 8, 24),
                       "신청일자": None, "완료일": K.date(2026, 8, 25)}
                K._save_extract_cache("exact", [row])
                with mock.patch.object(K, "_extract_fingerprint", return_value="exact"), \
                        mock.patch.object(K, "_load_reconcile",
                                          side_effect=AssertionError("source reparsed")):
                    got = K.extract(["immutable-source.txt"])
                self.assertEqual(K.date(2026, 8, 24), got[0]["예정일"])
                self.assertEqual(K.date(2026, 8, 25), got[0]["완료일"])
            finally:
                K.EXTRACT_CACHE = old_cache

    def test_exec_report_restart_uses_exact_disk_last_good_before_rebuild(self):
        from webapp import app_server as A
        last_good = {"meta": {"보고일": "2026-08-24"}, "sections": []}
        old_cache = dict(A._cache)
        try:
            A._cache.clear()
            with mock.patch.object(A, "DEMO", False), \
                    mock.patch.object(A, "_fresh", return_value=None), \
                    mock.patch.object(A, "_disk_cache_load", return_value=last_good), \
                    mock.patch.object(A, "_spawn_refresh") as refresh, \
                    mock.patch("ecount_reconcile.resolve_master",
                               side_effect=AssertionError("master reopened")):
                got = A.get_exec_report()
            self.assertIs(last_good, got)
            self.assertIs(last_good, A._cache["exec_stale"])
            refresh.assert_called_once()
        finally:
            A._cache.clear()
            A._cache.update(old_cache)


if __name__ == "__main__":
    unittest.main()
