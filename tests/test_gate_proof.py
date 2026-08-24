# -*- coding: utf-8 -*-
import json
import tempfile
import unittest
from pathlib import Path

import daily_run as D


class GateProofTests(unittest.TestCase):
    def _root(self, td):
        root = Path(td)
        (root / "reports").mkdir()
        (root / "tests").mkdir()
        (root / "app.py").write_text("VALUE = 1\n", encoding="utf-8")
        (root / "tests" / "synthetic_check.py").write_text(
            "print('ALL GREEN')\n", encoding="utf-8"
        )
        return root

    def test_unchanged_code_reuses_green_proof_without_running_suite(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-proof-") as td:
            root = self._root(td)
            stamp = D._gate_fingerprint(root)
            D._save_gate_proof(stamp, 12.3, root)
            calls = []

            result = D._run_gate(
                root,
                runner=lambda *a, **k: calls.append((a, k)),
            )

            self.assertTrue(result["ok"])
            self.assertTrue(result["cached"])
            self.assertIn("ALL GREEN", result["out"])
            self.assertEqual([], calls)

    def test_changed_code_forces_suite_and_records_new_proof_only_on_green(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-change-") as td:
            root = self._root(td)
            old = D._gate_fingerprint(root)
            D._save_gate_proof(old, 10, root)
            (root / "app.py").write_text("VALUE = 2\n", encoding="utf-8")
            calls = []

            def green(*args, **kwargs):
                calls.append((args, kwargs))
                return {"name": "합성검증", "ok": True, "returncode": 0,
                        "out": "ALL GREEN"}

            result = D._run_gate(root, runner=green)
            proof = json.loads(D._gate_proof_path(root) and Path(
                D._gate_proof_path(root)
            ).read_text(encoding="utf-8"))
            self.assertTrue(result["ok"])
            self.assertEqual(1, len(calls))
            self.assertNotEqual(old["fingerprint"], proof["fingerprint"])
            self.assertEqual(D._gate_fingerprint(root)["fingerprint"], proof["fingerprint"])

    def test_failure_preserves_last_good_proof(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-fail-") as td:
            root = self._root(td)
            old = D._gate_fingerprint(root)
            D._save_gate_proof(old, 10, root)
            proof_path = Path(D._gate_proof_path(root))
            before = proof_path.read_bytes()
            (root / "app.py").write_text("VALUE = 3\n", encoding="utf-8")

            result = D._run_gate(
                root,
                runner=lambda *a, **k: {
                    "name": "합성검증", "ok": False, "returncode": 1,
                    "out": "시간초과(1500s)",
                },
            )
            self.assertFalse(result["ok"])
            self.assertEqual(before, proof_path.read_bytes())

    def test_corrupt_proof_never_skips_suite(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-corrupt-") as td:
            root = self._root(td)
            Path(D._gate_proof_path(root)).write_text("{broken", encoding="utf-8")
            calls = []

            result = D._run_gate(
                root,
                runner=lambda *a, **k: calls.append(1) or {
                    "name": "합성검증", "ok": True, "returncode": 0,
                    "out": "ALL GREEN",
                },
            )
            self.assertTrue(result["ok"])
            self.assertEqual([1], calls)


if __name__ == "__main__":
    unittest.main()
