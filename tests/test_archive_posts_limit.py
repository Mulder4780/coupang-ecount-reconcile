import json
import os
import sys
import tempfile
import unittest
from unittest import mock


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from band import archive_posts


class ArchivePostsLimitTests(unittest.TestCase):
    def test_limit_is_shared_across_all_bands(self):
        with tempfile.TemporaryDirectory() as td:
            for band in ("84789192", "90610953"):
                with open(os.path.join(td, f"{band}.json"), "w", encoding="utf-8") as f:
                    json.dump({"posts": {"1": {}, "2": {}}}, f)

            seen = []

            def fake_archive(_band, _posts, limit, _force, _stat):
                seen.append(limit)
                return min(2, limit)

            with mock.patch.object(archive_posts, "CACHE", td), \
                    mock.patch.object(archive_posts, "archive_band", side_effect=fake_archive), \
                    mock.patch.object(sys, "argv", ["archive_posts.py", "--limit", "3"]):
                self.assertEqual(archive_posts.main(), 0)

            self.assertEqual(seen, [3, 1])

    def test_daily_run_does_not_repeat_archive_after_collect_all(self):
        with open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8") as f:
            source = f.read()
        self.assertIn('os.path.join(ROOT, "collect_all.py")', source)
        self.assertNotIn('[os.path.join(ROOT, "band", "archive_posts.py")', source)
        self.assertIn('[os.path.join(ROOT, "stmt_archive.py"), "--limit", "40"]', source)
        self.assertIn('[os.path.join(ROOT, "tax_archive.py"), "--limit", "40"]', source)

    def test_document_archives_publish_only_completed_pdf(self):
        import stmt_archive
        import tax_archive

        with tempfile.TemporaryDirectory() as td:
            for module, args in (
                    (stmt_archive, ({}, "", "", os.path.join(td, "stmt.pdf"))),
                    (tax_archive, ({}, os.path.join(td, "tax.pdf")))):
                destination = args[-1]

                def interrupted(*render_args):
                    with open(render_args[-1], "wb") as f:
                        f.write(b"partial")
                    return None

                with mock.patch.object(module, "render", side_effect=interrupted):
                    self.assertIsNone(module.render_atomic(*args))
                self.assertFalse(os.path.exists(destination))
                self.assertFalse(any(name.startswith(os.path.basename(destination) + ".part-")
                                     for name in os.listdir(td)))


if __name__ == "__main__":
    unittest.main()
