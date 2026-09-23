"""把报关单 PDF 文本导出到 tmp，便于分析版式与字段位置。"""
from __future__ import annotations

import sys
from pathlib import Path

from pypdf import PdfReader

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT_ROOT / "data" / "raw" / "customs_declaration"
OUT_DIR = PROJECT_ROOT / "tmp" / "pdf_probe"


def main() -> None:
    names = sys.argv[1:]
    files = [RAW_DIR / n for n in names] if names else sorted(RAW_DIR.glob("*.pdf"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for path in files:
        if not path.exists():
            print(f"missing {path.name}")
            continue
        reader = PdfReader(str(path))
        chunks = []
        for idx, page in enumerate(reader.pages, 1):
            chunks.append(f"===== page {idx} =====")
            chunks.append(page.extract_text() or "")
        (OUT_DIR / f"{path.stem}.txt").write_text("\n".join(chunks), encoding="utf-8")
        print(f"{path.name}: pages={len(reader.pages)}")


if __name__ == "__main__":
    main()
