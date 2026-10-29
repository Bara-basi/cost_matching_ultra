from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from app.services import pipeline
from app.services.parser.models import DeclarationHeader, ParsedDeclaration


def _parsed(path: Path, sheet_type: str, declaration: str) -> ParsedDeclaration:
    return ParsedDeclaration(
        source_file=path.name,
        header=DeclarationHeader(
            declaration_no=declaration,
            contract_raw="26MT-01A100",
            sheet_type=sheet_type,
        ),
        items=[],
    )


class PipelineCustomsTests(unittest.TestCase):
    def test_pre_record_rows_are_used_and_formal_rows_win_duplicates(self):
        paths = [Path("formal.pdf"), Path("pre-duplicate.pdf"), Path("pre-only.pdf")]
        parsed_by_name = {
            "formal.pdf": _parsed(paths[0], "出口退税联", "D1"),
            "pre-duplicate.pdf": _parsed(paths[1], "预录单（仅供核对用）", "D1"),
            "pre-only.pdf": _parsed(paths[2], "预录单（仅供核对用）", "D2"),
        }

        def build_rows(parsed):
            return [
                {
                    "来源文件": parsed.source_file,
                    "单据类型": parsed.header.sheet_type,
                    "合同号_1": parsed.header.contract_raw,
                    "报关单号": parsed.header.declaration_no,
                }
            ]

        captured = []
        with patch.object(
            pipeline,
            "parse_declaration",
            side_effect=lambda path: parsed_by_name[path.name],
        ), patch(
            "app.services.parser.pre_record_parser.parse_pre_record",
            side_effect=lambda path: parsed_by_name[path.name],
        ), patch.object(pipeline, "build_parse_rows", side_effect=build_rows), patch.object(
            pipeline,
            "write_workbook",
            side_effect=lambda rows, _path: captured.extend(rows),
        ):
            stats = pipeline.parse_pdfs(paths, Path("unused.xlsx"))

        self.assertEqual(
            [(row["报关单号"], row["单据类型"]) for row in captured],
            [("D1", "出口退税联"), ("D2", "预录单")],
        )
        self.assertEqual(stats["formalRows"], 1)
        self.assertEqual(stats["preRecordRows"], 1)


if __name__ == "__main__":
    unittest.main()
