"""飞书主表报关附件下载脚本的离线测试。"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from scripts.fetch_ai_customs_attachments import (
    attachment_items,
    date_text,
    is_pdf_file,
    normalize_month,
    plain_text,
    safe_filename,
)


class FeishuCustomsAttachmentTests(unittest.TestCase):
    def test_month_and_date_normalization(self):
        self.assertEqual(normalize_month("2026-10"), "2026-10")
        self.assertEqual(normalize_month("2026/1"), "2026-01")
        self.assertEqual(date_text(1790784000000), "2026-10-01")
        self.assertEqual(date_text("2026-10-20 12:00:00"), "2026-10-20")
        with self.assertRaises(ValueError):
            normalize_month("2026-13")

    def test_feishu_value_flattening_and_pdf_filter(self):
        self.assertEqual(plain_text([{"text": "出口"}, {"text": "预录"}]), "出口 预录")
        attachments = [
            {"file_token": "a", "type": "application/pdf"},
            {"file_token": "b", "type": "image/png"},
            {"type": "application/pdf"},
        ]
        self.assertEqual(
            [item["file_token"] for item in attachment_items(attachments)], ["a"]
        )

    def test_safe_filename_and_pdf_validation(self):
        self.assertEqual(safe_filename(r"a:b*c?.pdf"), "a_b_c_.pdf")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ok.pdf"
            path.write_bytes(b"%PDF-1.7\nfixture")
            self.assertTrue(is_pdf_file(path))
            path.write_bytes(b"not pdf")
            self.assertFalse(is_pdf_file(path))


if __name__ == "__main__":
    unittest.main()
