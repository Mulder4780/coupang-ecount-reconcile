# -*- coding: utf-8 -*-
import tempfile
import unittest
from pathlib import Path

import daily_run as D


class GateSplitTests(unittest.TestCase):
    def _root(self, td):
        root = Path(td)
        (root / "reports").mkdir()
        (root / "tests").mkdir()
        (root / "app.py").write_text("VALUE = 1\n", encoding="utf-8")
        (root / "tests" / "synthetic_check.py").write_text(
            "# --window-audit-only --skip-window-audit\n",
            encoding="utf-8",
        )
        return root

    def test_gate_runs_full_window_audit_then_core_with_separate_limits(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-split-") as td:
            root = self._root(td)
            calls = []

            def green(name, args, **kwargs):
                calls.append((name, list(args), dict(kwargs)))
                marker = "WINDOW AUDIT GREEN" if "--window-audit-only" in args else "ALL GREEN"
                return {"name": name, "ok": True, "returncode": 0, "out": marker}

            result = D._run_gate(root, runner=green)
            self.assertTrue(result["ok"])
            self.assertTrue(result["분리관문"])
            self.assertEqual(2, len(calls))
            self.assertIn("--window-audit-only", calls[0][1])
            self.assertEqual(D.GATE_WINDOW_AUDIT_TIMEOUT_S, calls[0][2]["timeout"])
            self.assertIn("--skip-window-audit", calls[1][1])
            self.assertEqual(D.GATE_TIMEOUT_S, calls[1][2]["timeout"])
            self.assertIn("WINDOW AUDIT GREEN", result["out"])
            self.assertIn("ALL GREEN", result["out"])

    def test_failed_window_audit_never_runs_or_certifies_core(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-split-fail-") as td:
            root = self._root(td)
            calls = []

            def fail(name, args, **kwargs):
                calls.append(list(args))
                return {"name": name, "ok": False, "returncode": 1, "out": "audit failed"}

            result = D._run_gate(root, runner=fail)
            self.assertFalse(result["ok"])
            self.assertEqual(1, len(calls))
            self.assertIn("창 팝업 감사 분리 관문 실패", result["out"])
            self.assertFalse(Path(D._gate_proof_path(root)).exists())

    def test_old_unsplit_checker_keeps_one_call_compatibility(self):
        with tempfile.TemporaryDirectory(prefix="csos-gate-old-") as td:
            root = Path(td)
            (root / "reports").mkdir()
            (root / "tests").mkdir()
            (root / "tests" / "synthetic_check.py").write_text(
                "print('ALL GREEN')\n", encoding="utf-8"
            )
            calls = []
            result = D._run_gate(
                root,
                runner=lambda *a, **k: calls.append((a, k)) or {
                    "name": "합성검증", "ok": True, "returncode": 0, "out": "ALL GREEN"
                },
            )
            self.assertTrue(result["ok"])
            self.assertEqual(1, len(calls))


if __name__ == "__main__":
    unittest.main()
