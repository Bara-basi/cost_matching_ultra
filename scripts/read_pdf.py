"""读取报关单 PDF 文本，输出到 tmp 目录便于查看。"""
from __future__ import annotations

import sys
from pathlib import Path

from pypdf import PdfReader

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def read_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    chunks: list[str] = []
    for idx, page in enumerate(reader.pages, 1):
        chunks.append(f"===== page {idx} =====")
        chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def main() -> None:
    names = sys.argv[1:]
    out: list[str] = []
    for name in names:
        path = Path(name)
        if not path.is_absolute():
            path = PROJECT_ROOT / name
        out.append(f"########## {path.name} ##########")
        try:
            out.append(read_pdf(path))
        except Exception as exc:  # noqa: BLE001
            out.append(f"[ERROR] {exc}")
        out.append("")
    target = PROJECT_ROOT / "tmp" / "pdf_text.txt"
    target.write_text("\n".join(out), encoding="utf-8")
    print(f"written {target}")


if __name__ == "__main__":
    main()
