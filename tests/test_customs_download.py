"""单一窗口下载器的离线回归测试。"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from app.services.customs_download import (
    DownloadManifest,
    PageRecord,
    date_range,
    declaration_index,
    default_run_id,
    extract_declaration_numbers,
    load_manifest,
    normalize_date,
    save_manifest,
    valid_pdf,
    write_pdf_atomic,
)


class CustomsDownloadTests(unittest.TestCase):
    def test_date_normalization_and_default_range(self):
        self.assertEqual(normalize_date("2026-10-01"), "20261001")
        self.assertEqual(normalize_date("2026/10/01"), "20261001")
        self.assertEqual(date_range("2026-10-01", "2026-10-03"), ("20261001", "20261003"))
        self.assertEqual(date_range("", "2026-10-03", days=2), ("20261001", "20261003"))
        with self.assertRaises(ValueError):
            date_range("2026-10-04", "2026-10-03")

    def test_declaration_numbers_are_deduplicated_in_page_order(self):
        rows = [
            "1 310120260518420169 出口\n",
            "2 310120260518420170 出口",
            "副本 310120260518420169",
            "不是报关单号 12345678901234567",
        ]
        self.assertEqual(
            extract_declaration_numbers(rows),
            ["310120260518420169", "310120260518420170"],
        )

    def test_atomic_pdf_write_rejects_non_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "x.pdf"
            with self.assertRaises(ValueError):
                write_pdf_atomic(target, b"<html></html>")
            self.assertFalse(target.exists())
            write_pdf_atomic(target, b"%PDF-1.7\nfixture")
            self.assertTrue(valid_pdf(target))

    def test_manifest_resume_and_declaration_index(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            raw = root / "raw" / "page.pdf"
            raw.parent.mkdir(parents=True)
            raw.write_bytes(b"%PDF-1.7\nfixture")
            manifest = DownloadManifest(
                run_id="export_20261001_20261002",
                flow="export",
                start_date="20261001",
                end_date="20261002",
                output_dir=str(root),
                pages=[
                    PageRecord(
                        page=1,
                        declaration_numbers=[
                            "310120260518420169",
                            "310120260518420170",
                        ],
                        pdf="raw/page.pdf",
                        bytes=raw.stat().st_size,
                        sha256="abc",
                        status="downloaded",
                    )
                ],
            )
            path = root / "manifest.json"
            save_manifest(path, manifest)
            loaded = load_manifest(path)
            self.assertIsNotNone(loaded)
            self.assertEqual(loaded.pages[0].declaration_numbers, manifest.pages[0].declaration_numbers)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["status"], "running")
            index = declaration_index(loaded, root)
            self.assertEqual(index["310120260518420169"]["references"][0]["page"], 1)
            self.assertEqual(
                index["310120260518420170"]["references"][0]["pdf"],
                "raw/page.pdf",
            )

    def test_default_run_id_is_stable_for_resume(self):
        self.assertEqual(
            default_run_id("20261001", "20261002"),
            "export_20261001_20261002",
        )


if __name__ == "__main__":
    unittest.main()
