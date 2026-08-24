# -*- coding: utf-8 -*-
"""ESA001M이 많은 최신 ERP 파일 뒤로 밀리지 않는지 보는 집중 회귀 시험."""
import os
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import customer_index as customer


class CustomerSourcePriorityTest(unittest.TestCase):
    def test_esa001m_survives_one_hundred_newer_noise_files(self):
        with tempfile.TemporaryDirectory(prefix="csos-customer-source-") as td:
            esa = os.path.join(td, "ESA001M.xlsx")
            open(esa, "wb").close()
            os.utime(esa, (1, 1))

            pairs = [(1, 0, esa)]
            now = time.time()
            for n in range(100):
                path = os.path.join(td, f"noise_{n:03d}.xlsx")
                open(path, "wb").close()
                stamp = now + n
                os.utime(path, (stamp, stamp))
                pairs.append((stamp, 0, path))

            selected = customer.prioritize_customer_sources(pairs, limit=80)
            selected_paths = [row[2] for row in selected]

            self.assertEqual(selected_paths[0], esa)
            self.assertIn(esa, selected_paths)
            self.assertEqual(len(selected), 81)
            self.assertEqual(len(set(map(os.path.normcase, selected_paths))), 81)


if __name__ == "__main__":
    unittest.main()
