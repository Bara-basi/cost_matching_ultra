"""从已下载的附件 ZIP 里解出含「入库单」的文件（供后续解析结算金额）。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_attachments import ATTACH_DIR, extract_grn_files, load_manifest  # noqa: E402


def main() -> None:
    manifest = load_manifest()
    out: list[str] = []
    total = 0
    for code, item in manifest.items():
        if not item.get("hasGrn"):
            continue
        files = extract_grn_files(code)
        total += len(files)
        out.append(f"{code}: {[f.name for f in files]}")
    reports = ATTACH_DIR.parent / "reports"
    reports.mkdir(parents=True, exist_ok=True)
    (reports / "grn_extracted.txt").write_text("\n".join(out), encoding="utf-8")
    print(f"含入库单的采购单={len(out)} 解出文件={total}")


if __name__ == "__main__":
    main()
