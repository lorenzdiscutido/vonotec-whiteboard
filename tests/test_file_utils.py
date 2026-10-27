"""Tests for file_utils module."""
import os
import json
import tempfile
import unittest

from file_utils import sanitize_batch_filename, load_processed_log, save_processed_log, load_prefs, save_prefs
import settings


class TestFileUtils(unittest.TestCase):
    def test_sanitize_batch_filename_normal(self):
        self.assertEqual(sanitize_batch_filename("my_batch"), "my_batch.pptx")
        self.assertEqual(sanitize_batch_filename("my_batch.pptx"), "my_batch.pptx")
        self.assertEqual(sanitize_batch_filename("my_batch.PPTX"), "my_batch.PPTX")

    def test_sanitize_batch_filename_invalid_chars(self):
        # Characters like \ / : * ? " < > |
        sanitized = sanitize_batch_filename('test/batch:name?*')
        self.assertEqual(sanitized, "test_batch_name__.pptx")

    def test_sanitize_batch_filename_empty_or_whitespace(self):
        self.assertIsNone(sanitize_batch_filename(""))
        self.assertIsNone(sanitize_batch_filename("   "))
        self.assertIsNone(sanitize_batch_filename(None))
        self.assertIsNone(sanitize_batch_filename("..."))

    def test_load_and_save_prefs(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            orig_prefs = settings.PREFS_FILE
            test_prefs_file = os.path.join(tmp_dir, "test_prefs.json")
            settings.PREFS_FILE = test_prefs_file
            try:
                # Initially empty
                self.assertEqual(load_prefs(), {})

                # Save updates
                save_prefs({"test_key": "test_val", "number": 42})
                loaded = load_prefs()
                self.assertEqual(loaded.get("test_key"), "test_val")
                self.assertEqual(loaded.get("number"), 42)

                # Incremental update
                save_prefs({"another_key": True})
                loaded = load_prefs()
                self.assertEqual(loaded.get("test_key"), "test_val")
                self.assertEqual(loaded.get("another_key"), True)
            finally:
                settings.PREFS_FILE = orig_prefs


if __name__ == "__main__":
    unittest.main()
