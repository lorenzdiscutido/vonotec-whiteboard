"""Tests for pptx_writer module."""
import io
import unittest
from PIL import Image

from pptx_writer import (
    format_defect_text,
    _fit_inside,
    _slot_geometry,
    _extract_metadata,
    optimize_image_in_memory,
)


class TestPptxWriter(unittest.TestCase):
    def test_format_defect_text_no_defects(self):
        parsed = {
            "Sealant Damage": [],
            "Paint Damage": [],
            "Concrete Damage": [],
        }
        self.assertEqual(format_defect_text(parsed), "No defect recorded")

    def test_fit_inside_maintains_aspect_ratio(self):
        frame = (100, 200, 400, 600)  # left, top, width, height
        img_w, img_h = 800, 1600  # 1:2 aspect ratio
        pic_left, pic_top, pic_w, pic_h = _fit_inside(frame, img_w, img_h)
        self.assertEqual(pic_w, 300)
        self.assertEqual(pic_h, 600)
        self.assertEqual(pic_left, 100 + (400 - 300) // 2)
        self.assertEqual(pic_top, 200)

    def test_slot_geometry(self):
        frames, caption_top = _slot_geometry(3)
        self.assertEqual(len(frames), 3)
        self.assertGreater(caption_top, 0)

    def test_optimize_image_in_memory(self):
        img = Image.new("RGB", (200, 400), color=(255, 0, 0))
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        raw_bytes = buf.getvalue()

        opt_stream, opt_w, opt_h = optimize_image_in_memory(raw_bytes)
        self.assertIsInstance(opt_stream, io.BytesIO)
        self.assertGreater(opt_w, 0)
        self.assertGreater(opt_h, 0)


if __name__ == "__main__":
    unittest.main()
