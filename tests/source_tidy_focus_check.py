# -*- coding: utf-8 -*-
"""원본자료자동정리의 대용량·잠금 회귀검증.

실제 Z:에 10만 파일을 만들지 않는다. 공용 워커가 연도 폴더로 내려가려 할 때만
10만 정본을 가상으로 내놓게 해서, 하강 전 가지치기가 빠지면 즉시 드러나게 한다.
"""
from __future__ import annotations

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock


HERE = Path(__file__).resolve().parent
ECOUNT = HERE.parent
if str(ECOUNT) not in sys.path:
    sys.path.insert(0, str(ECOUNT))

import collect_sources  # noqa: E402
import pid_alive  # noqa: E402
import source_index  # noqa: E402
import source_organizer as organizer  # noqa: E402


class SourceTidyFocusCheck(unittest.TestCase):
    def setUp(self):
        organizer._MTIME.clear()
        organizer.start_clock(60)

    def test_completed_100k_tree_is_pruned_and_new_one_goes_one_to_zero(self):
        """정리완료 10만 건은 안 훑고, 신규 1건만 계획·적용한 뒤 계획 0건이다."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            erp = root / "1. ERP 내보내기"
            completed = erp / "2026" / "08" / "2026-08-23"
            completed.mkdir(parents=True)
            incoming = erp / "new.xlsx"
            incoming.write_bytes(b"new")

            real_walk = source_index.walk_stat
            observed = {"virtual_completed": 0, "erp_skip": set()}

            def virtual_walk(folder, skip_dirs=None):
                skip = set(skip_dirs or ())
                if os.path.normcase(os.path.abspath(folder)) == os.path.normcase(str(erp)):
                    observed["erp_skip"] = skip
                    # 가지치기가 퇴행하면 10만 정본이 실제 계획으로 흘러든다.
                    if "2026" not in skip:
                        stamp = type("Stat", (), {"st_mtime": time.time()})()
                        for number in range(100_000):
                            observed["virtual_completed"] += 1
                            yield str(completed), f"done-{number:06d}.xlsx", stamp
                yield from real_walk(folder, skip_dirs=skip)

            with mock.patch.object(source_index, "walk_stat", virtual_walk):
                first = organizer.planned_moves(str(root))
                self.assertEqual(1, len(first))
                self.assertEqual(str(incoming), first[0].src)
                self.assertIn("2026", observed["erp_skip"])
                self.assertEqual(0, observed["virtual_completed"])

                done, errors = organizer.apply_moves(first, root=str(root))
                self.assertEqual(1, done)
                self.assertEqual([], errors)

                organizer._MTIME.clear()
                second = organizer.planned_moves(str(root))
                self.assertEqual([], second)

    def test_guide_prunes_years_and_band_post_archive(self):
        """12MB 안내문을 만들던 연도 정본·밴드 게시글 9만 건을 나열하지 않는다."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            band = root / "4. 밴드 원본"
            band.mkdir()
            with mock.patch.object(collect_sources, "BAND_DIR", str(band)):
                skip = collect_sources._guide_skip_dirs(str(band))
            self.assertIn("2026", skip)
            self.assertIn("게시글보관", skip)
            self.assertNotIn("브라우저덤프", skip)

    def test_live_lock_is_never_reclaimed_by_age(self):
        """오래됐어도 살아 있는 정확한 PID 지문이면 잠금을 빼앗지 않는다."""
        with tempfile.TemporaryDirectory() as tmp:
            lock = Path(tmp) / ".source_organizer.lock"
            fingerprint = pid_alive.stamp()
            self.assertTrue(fingerprint, "현재 프로세스 지문을 읽어야 잠금 검증이 유효합니다")
            content = f"{os.getpid()} {fingerprint} 2000-01-01T00:00:00\n"
            lock.write_text(content, encoding="utf-8")

            with mock.patch.object(organizer, "ORIGIN_ROOT", tmp), \
                    mock.patch.object(organizer, "LOCK", str(lock)):
                self.assertFalse(organizer._lock_acquire())

            self.assertTrue(lock.exists())
            self.assertEqual(content, lock.read_text(encoding="utf-8"))

    def test_browser_dump_is_a_protected_band_input(self):
        self.assertIn("브라우저덤프", organizer.PROTECTED_BAND_DIRS)
        with tempfile.TemporaryDirectory() as tmp:
            browser_dump = Path(tmp) / "4. 밴드 원본" / "브라우저덤프"
            browser_dump.mkdir(parents=True)
            dump = browser_dump / "84789192_20260824.json"
            dump.write_text("{}", encoding="utf-8")
            moves = organizer.planned_moves(tmp)
            self.assertFalse(any(os.path.normcase(m.src) == os.path.normcase(str(dump))
                                 for m in moves))


if __name__ == "__main__":
    unittest.main(verbosity=2)
