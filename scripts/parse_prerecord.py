"""批量解析“仅供核对用”预录报关单 PDF。

默认读取 ``tmp/customs_ai_attachments/<当月>/pre_record``，结果仍写在 tmp：

    python -X utf8 scripts/parse_prerecord.py
    python -X utf8 scripts/parse_prerecord.py \
        --pdf-dir tmp/customs_ai_attachments/2026-10/pre_record \
        --output tmp/customs_ai_attachments/2026-10/prerecord_parse
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.parser.excel_writer import build_rows, write_workbook  # noqa: E402
from app.services.parser.pre_record_parser import parse_pre_record  # noqa: E402


def parse_args() -> argparse.Namespace:
    month = datetime.now().strftime("%Y-%m")
    parser = argparse.ArgumentParser(description="批量解析预录报关单 PDF")
    parser.add_argument(
        "--pdf-dir",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_ai_attachments" / month / "pre_record",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_ai_attachments" / month / "prerecord_parse",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    pdf_dir = args.pdf_dir.resolve()
    output_dir = args.output.resolve()
    paths = sorted(pdf_dir.glob("*.pdf")) if pdf_dir.is_dir() else []
    if not paths:
        raise SystemExit(f"没有找到预录 PDF：{pdf_dir}")

    rows: list[dict] = []
    failed: list[dict] = []
    warnings: Counter[str] = Counter()
    item_counts: Counter[int] = Counter()
    seen: set[tuple] = set()
    declarations: dict[str, dict] = {}
    for path in paths:
        try:
            parsed = parse_pre_record(path)
        except Exception as exc:  # noqa: BLE001
            failed.append({"文件": path.name, "原因": str(exc)})
            continue
        warnings.update(parsed.warnings)
        item_counts[len(parsed.items)] += 1
        declarations[parsed.header.declaration_no or path.stem] = {
            "file": path.name,
            "declarationNo": parsed.header.declaration_no,
            "contract": parsed.header.contract_raw,
            "exportDate": parsed.header.export_date,
            "items": len(parsed.items),
            "warnings": parsed.warnings,
        }
        for row in build_rows(parsed):
            marker = (
                str(row.get("报关单号") or ""),
                str(row.get("商品序号") or ""),
                str(row.get("报关品名") or ""),
                str(row.get("海关编码") or ""),
                str(row.get("报关重量") or ""),
                str(row.get("总价") or ""),
            )
            if marker in seen:
                continue
            seen.add(marker)
            row["单据类型"] = "预录单（仅供核对用）"
            # 预录单上的商品总价就是本行报关金额；成本匹配仍以报关单原件事实为准。
            row["报关金额"] = row.get("总价") or ""
            rows.append(row)

    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = output_dir / "预录单解析结果.xlsx"
    if rows:
        write_workbook(rows, workbook)
    stats = {
        "pdf总数": len(paths),
        "解析成功": len(paths) - len(failed),
        "解析失败": len(failed),
        "报关单数": len(declarations),
        "商品行数": len(rows),
        "每单商品数分布": {str(key): value for key, value in sorted(item_counts.items())},
        "警告统计": dict(warnings),
        "失败清单": failed,
        "报关单": declarations,
        "生成时间": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    (output_dir / "解析统计.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    print(
        f"预录 PDF={len(paths)} 成功={stats['解析成功']} 失败={len(failed)} "
        f"报关单={len(declarations)} 商品行={len(rows)}"
    )
    print(f"Excel：{workbook}")
    print(f"统计：{output_dir / '解析统计.json'}")
    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
