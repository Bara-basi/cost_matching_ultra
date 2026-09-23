"""批量解析 data/raw/customs_declaration 下的报关单 PDF，输出到 outputs/customs_parse。

产出：
- 报关单解析结果_出口退税联.xlsx   正式报关单（成本匹配用）
- 报关单解析结果_预录单.xlsx       「仅供核对用」版式，对应飞书预录单，不参与成本匹配
- 解析统计.json                   覆盖率、版式分布、重复与失败清单
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.parser import parse_declaration  # noqa: E402
from app.services.parser.excel_writer import build_rows, write_workbook  # noqa: E402
from app.services.scope import out_of_scope_reason  # noqa: E402

RAW_DIR = PROJECT_ROOT / "data" / "raw" / "customs_declaration"
OUT_DIR = PROJECT_ROOT / "outputs" / "customs_parse"


def main() -> None:
    paths = sorted(RAW_DIR.glob("*.pdf"))
    rows_by_type: dict[str, list[dict]] = {"出口退税联": [], "预录单（仅供核对用）": []}
    intercepted: list[dict] = []
    failed: list[dict] = []
    warnings = Counter()
    item_counts: Counter[int] = Counter()
    seen: dict[tuple, list[str]] = {}
    for path in paths:
        try:
            parsed = parse_declaration(path)
        except Exception as exc:  # noqa: BLE001
            failed.append({"文件": path.name, "原因": str(exc)})
            continue
        for warning in parsed.warnings:
            warnings[warning] += 1
        item_counts[len(parsed.items)] += 1
        built = build_rows(parsed)
        # 系统输入即拦截：历史遗留合同号（CY / xxSM / SP）不进主结果
        kept: list[dict] = []
        for row in built:
            reason = out_of_scope_reason(str(row.get("合同号_1") or ""))
            if reason:
                intercepted.append({**row, "拦截原因": reason})
            else:
                kept.append(row)
        built = kept
        if not built:
            continue
        rows_by_type.setdefault(parsed.header.sheet_type, []).extend(built)
        # 同一报关单号的重复 PDF：保留全部文件，但在结果里标注重复来源
        key = (
            built[0]["报关单号"] if built else "",
            built[0]["合同号_1"] if built else "",
            len(built),
        )
        if key[0]:
            seen.setdefault(key, []).append(path.name)

    dup_names: dict[str, list[str]] = {}
    for key, names in seen.items():
        if len(names) > 1:
            dup_names.setdefault(key[0], []).extend(names)
    duplicate_files = [{"报关单号": decl, "文件": names} for decl, names in sorted(dup_names.items())]
    for rows in rows_by_type.values():
        for row in rows:
            names = dup_names.get(row.get("报关单号"))
            row["重复来源文件"] = " / ".join(names) if names else ""

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    written: dict[str, int] = {}
    for sheet_type, rows in rows_by_type.items():
        if not rows:
            continue
        filename = (
            "报关单解析结果_出口退税联.xlsx"
            if sheet_type == "出口退税联"
            else "报关单解析结果_预录单.xlsx"
        )
        write_workbook(rows, OUT_DIR / filename)
        written[sheet_type] = len(rows)

    total_rows = sum(written.values())
    stats = {
        "PDF总数": len(paths),
        "解析成功": len(paths) - len(failed),
        "解析失败": len(failed),
        "商品行数": total_rows,
        "拦截_历史遗留": len(intercepted),
        "按单据类型": written,
        "同一报关单存在多个PDF的组数": len(duplicate_files),
        "重复报关单明细": duplicate_files,
        "每条报关单商品数分布": {str(k): v for k, v in sorted(item_counts.items())},
        "警告统计": dict(warnings),
        "失败清单": failed,
    }
    (OUT_DIR / "解析统计.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if intercepted:
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "拦截_历史遗留"
        columns = list(intercepted[0].keys())
        ws.append(columns)
        for row in intercepted:
            ws.append([row.get(c, "") for c in columns])
        wb.save(OUT_DIR / "拦截_历史遗留合同号.xlsx")
    print(
        f"PDF={len(paths)} 成功={stats['解析成功']} 失败={len(failed)} 商品行={total_rows} "
        f"出口退税联={written.get('出口退税联', 0)} 预录单={written.get('预录单（仅供核对用）', 0)}"
    )


if __name__ == "__main__":
    main()
