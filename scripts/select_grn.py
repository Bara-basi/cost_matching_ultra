"""按「分批 + 去旧版」挑出有效入库单，并导出三张表。

用法：
    python scripts/select_grn.py                 # 全量
    python scripts/select_grn.py --code 25MT-03R625Y-HD

输出：
    outputs/grn_attachments/有效入库单清单.xlsx
    outputs/grn_attachments/旧版丢弃清单.xlsx
    outputs/grn_attachments/人工确认清单.xlsx
    .cache/erp/reports/grn_selection.json
"""
from __future__ import annotations

import argparse
import collections
import json
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.grn_extract import PARSED_ROOT  # noqa: E402
from app.services.grn_select import select  # noqa: E402
from app.services.erp_cache import read_jsonl  # noqa: E402

OUT_DIR = PROJECT_ROOT / "outputs" / "grn_attachments"
REPORT = CACHE_ROOT / "reports" / "grn_selection.json"


def write_sheet(path: Path, title: str, headers: list[str], rows: list[list]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append(row)
    for index, header in enumerate(headers, 1):
        width = max(10, min(60, max([len(str(header))] + [len(str(r[index - 1])) for r in rows] or [10]) + 2))
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--code", default="")
    args = parser.parse_args()

    codes = sorted(path.name for path in PARSED_ROOT.iterdir() if path.is_dir())
    if args.code:
        codes = [code for code in codes if code == args.code]
    print(f"采购单 {len(codes)} 个", flush=True)

    kept_rows: list[list] = []
    dropped_rows: list[list] = []
    issue_rows: list[list] = []
    check_rows: list[list] = []
    details: list[dict] = []
    stats: collections.Counter = collections.Counter()
    purchases = {
        row.get("purchase_code"): row
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
    }

    for code in codes:
        result = select(code)
        details.append(result)
        stats["采购单"] += 1
        stats["附件"] += result["files"]
        stats["保留"] += len(result["kept"])
        stats["丢弃旧版/重复"] += len(result["dropped"])
        fees = sum(1 for item in result["kept"] if item["is_fee"])
        stats["其中费用补充件"] += fees
        stats["批次组"] += len(result["groups"])
        if len(result["groups"]) > 1:
            stats["多批次采购单"] += 1
        for item in result["kept"]:
            kept_rows.append(
                [code, item["batch"], item["original"], item["amount"], "费用件" if item["is_fee"] else "正单"]
            )
        for item in result["dropped"]:
            dropped_rows.append([code, item["batch"], item["original"]])
        for note in result["issues"]:
            issue_rows.append([code, note])

        # 质量校验：有效入库单合计 vs 睿贝采购单金额（ERP 金额只是量级参考）
        erp = purchases.get(code) or {}
        erp_amount = erp.get("amount")
        kept_total = sum(
            (Decimal(item["amount"]) for item in result["kept"] if item["amount"]),
            Decimal(0),
        )
        if erp_amount is None or not result["kept"]:
            verdict = "缺参照" if erp_amount is None else "无有效附件"
        else:
            diff = Decimal(str(erp_amount)) - kept_total
            if abs(diff) <= Decimal("0.05"):
                verdict = "完全一致"
                stats["金额一致"] += 1
            elif abs(diff) <= Decimal("1"):
                verdict = f"差 {diff:+}"
                stats["差1元内"] += 1
            else:
                verdict = f"差 {diff:+}"
                stats["金额不一致"] += 1
        check_rows.append(
            [code, erp_amount or "", str(kept_total) if result["kept"] else "", verdict,
             len(result["kept"]), len(result["dropped"])]
        )

    write_sheet(
        OUT_DIR / "有效入库单清单.xlsx",
        "有效入库单",
        ["采购单号", "批次", "附件名", "金额", "类型"],
        kept_rows,
    )
    write_sheet(
        OUT_DIR / "旧版丢弃清单.xlsx",
        "已丢弃",
        ["采购单号", "批次", "附件名"],
        dropped_rows,
    )
    write_sheet(
        OUT_DIR / "人工确认清单.xlsx",
        "待确认",
        ["采购单号", "说明"],
        issue_rows,
    )
    write_sheet(
        OUT_DIR / "与采购单金额对照.xlsx",
        "金额对照",
        ["采购单号", "ERP采购金额", "有效入库单合计", "差异", "保留份数", "丢弃份数"],
        check_rows,
    )
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(details, ensure_ascii=False, indent=1), encoding="utf-8")

    print(json.dumps(dict(stats), ensure_ascii=False))
    print("输出目录：", OUT_DIR)
    print("明细：", REPORT)


if __name__ == "__main__":
    main()
