"""从报关单 PDF 抽取文本。

默认模式在该版式下会「一行一个字段」，最适合按行取值；
layout 模式仅作为默认模式失败的兜底。
"""
from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader


def extract_text(path: Path) -> str:
    """抽取 PDF 全文。默认模式失败时回退 layout 模式。"""
    reader = PdfReader(str(path))
    chunks: list[str] = []
    for page in reader.pages:
        text = ""
        try:
            text = page.extract_text() or ""
        except Exception:  # noqa: BLE001
            text = ""
        if not text.strip():
            try:
                text = page.extract_text(extraction_mode="layout") or ""
            except Exception:  # noqa: BLE001
                text = ""
        chunks.append(text)
    return "\n".join(chunks)


def text_lines(text: str) -> list[str]:
    return [line.rstrip() for line in text.splitlines()]
