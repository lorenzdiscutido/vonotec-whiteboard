"""Tests for extraction module helpers."""
import io
import unittest
import json
from PIL import Image
from gemini_client import parse_and_clean_json, find_problems
from image_utils import load_image


class TestExtractionLogic(unittest.TestCase):
    def test_load_image_from_raw_bytes(self):
        src = Image.new("RGB", (2000, 1000), color=(100, 150, 200))
        buf = io.BytesIO()
        src.save(buf, format="JPEG")
        raw_bytes = buf.getvalue()

        img = load_image(raw_bytes, max_dimension=1600)
        self.assertEqual(img.size, (1600, 800))

    def test_parse_and_clean_json_valid(self):
        json_sample = json.dumps({
            "Submitter": "John Doe",
            "Date": "August 1, 2026",
            "Tower": "1",
            "Floor": "12",
            "Drop": "3",
            "Sealant Damage": ["DS CC"],
            "Sealant Dimension": ["150 CM"],
            "Concrete Damage": [],
            "Concrete Dimension": []
        })
        res = parse_and_clean_json(json_sample)
        self.assertEqual(res.get("Submitter"), "John Doe")
        self.assertEqual(res.get("Sealant Damage"), ["DS CC"])

    def test_find_problems_mismatched_arrays(self):
        bad_data = {
            "Sealant Damage": ["DS CC", "MS FG"],
            "Sealant Dimension": ["150 CM"]  # Only 1 dimension for 2 damages!
        }
        problems = find_problems(bad_data)
        # Should detect a structural problem for array length mismatch
        self.assertTrue(any(kind == "structural" for _, _, kind in problems))


if __name__ == "__main__":
    unittest.main()
