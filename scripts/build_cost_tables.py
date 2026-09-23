"""把入库单拆成两张「面向成本」的中间表（用户建议的方案）。

输出（`outputs/cost_match/`）：

    采购单产品明细.xlsx   一行 = 采购单 × 附件 × 区块批次 × 一条产品明细（材质/品名/规格/数量/单价/金额/重量）
    采购单入库单汇总.xlsx 一行 = 采购单 × 一份有效入库单（订单金额 / 实发金额 / 各区块 / 费用行）

用途：

* 一个采购单只对应一张报关单、或被报关单包含时，直接从汇总表取「实发金额」；
* 需要按批次/按货物拆分时，用产品明细表按 SKU（材质+规格+数量）去重与求和。
"""
from __future__ import annotations

import collections
import json
import sys
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

from app.services.cost_match import block_batch, dec  # noqa: E402
from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.grn_extract import PARSED_ROOT, dec_str  # noqa: E402
from app.services.grn_select import select  # noqa: E402

OUT_DIR = PROJECT_ROOT / "outputs" / "cost_match"


def line_value(line: dict, *keys: str) -> str:
    for key in keys:
        value = line.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def detail_rows(purchase_code: str) -> list[list]:
    result = select(purchase_code)
    rows: list[list] = []
    for item in result["kept"]:
        payload = None
        folder = PARSED_ROOT / purchase_code
        if folder.exists():
            from app.services.grn_extract import safe_name

            path = folder / (safe_name(item["file"], "file") + ".json")
            if path.exists():
                payload = json.loads(path.read_text(encoding="utf-8"))
        if not payload:
            continue
        for line in payload.get("lines") or []:
            if not isinstance(line, dict):
                continue
            rows.append(
                [
                    purchase_code,
                    item["original"],
                    item.get("batch") or "",
                    line_value(line, "material", "材质"),
                    line_value(line, "name", "品名"),
                    line_value(line, "spec", "规格"),
                    line_value(line, "qty", "数量", "结算数量")
                    or line_value(line, "数量 米", "数量 支", "数量 只"),
                    line_value(line, "unit", "单位"),
                    line_value(line, "unit_price", "单价", "单价 元/kg", "单价 元/只"),
                    line_value(line, "amount", "金额", "金额 元", "含税金额（元）"),
                    line_value(line, "weight", "重量", "净重", "实发重量"),
                ]
            )
    return rows


def summary_rows(purchase_code: str) -> list[list]:
    result = select(purchase_code)
    rows: list[list] = []
    folder = PARSED_ROOT / purchase_code
    for item in result["kept"]:
        blocks = "；".join(
            f"{block.get('label')}={block.get('amount')}" for block in (item.get("blocks") or [])
        )
        payload = {}
        if folder.exists():
            from app.services.grn_extract import safe_name

            path = folder / (safe_name(item["file"], "file") + ".json")
            if path.exists():
                payload = json.loads(path.read_text(encoding="utf-8"))
        # 费用行（木箱费/包装费/渗透费…）单独列出来：飞书的「采购金额」只记货值，
        # 我方实发合计含费用行，两个口径都要能给财务看（2026-09-24 用户建议）。
        fees = payload.get("extra_fees") or []
        fee_text = "；".join(
            f"{fee.get('label')}={fee.get('amount')}"
            for fee in fees
            if isinstance(fee, dict) and fee.get("amount")
        )
        fee_sum = sum(
            (dec(fee.get("amount")) for fee in fees if isinstance(fee, dict)),
            Decimal(0),
        )
        tail_notes = payload.get("tail_notes") or []
        added_labels = {
            str(fee.get("label") or "")
            for fee in fees
            if isinstance(fee, dict) and str(fee.get("label") or "").startswith("表尾补款")
        }
        tail_text = "；".join(
            f"{note.get('label')}={note.get('amount')}"
            f"（{'已计入成本' if f'表尾补款：{str(note.get('label') or '')[:60]}' in added_labels else '未计入'}）"
            for note in tail_notes
            if isinstance(note, dict)
        )
        settled = dec(item.get("amount"))
        rows.append(
            [
                purchase_code,
                item["original"],
                item.get("batch") or "",
                item.get("amount") or "",
                blocks,
                payload.get("order_amount") or "",
                fee_text,
                dec_str(fee_sum) if fee_sum else "",
                dec_str(settled - fee_sum) if settled else "",
                tail_text,
                "费用件" if item["is_fee"] else "正单",
            ]
        )
    for item in result["dropped"]:
        rows.append(
            [
                purchase_code,
                item["original"],
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "已丢弃（旧版/重复）",
            ]
        )
    return rows


def write(path: Path, title: str, headers: list[str], rows: list[list]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append(row)
    for index, header in enumerate(headers, 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            10, min(48, len(header) * 2 + 4)
        )
    sheet.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    print(f"written {path} ({len(rows)} rows)")


def main() -> None:
    codes = sorted(path.name for path in PARSED_ROOT.iterdir() if path.is_dir())
    detail: list[list] = []
    summary: list[list] = []
    for index, code in enumerate(codes, 1):
        detail.extend(detail_rows(code))
        summary.extend(summary_rows(code))
        if index % 200 == 0:
            print(f"  [{index}/{len(codes)}] 明细 {len(detail)} 行", flush=True)
    write(
        OUT_DIR / "采购单产品明细.xlsx",
        "产品明细",
        ["采购单号", "来源附件", "批次", "材质", "品名", "规格", "数量", "单位", "单价", "金额", "重量"],
        detail,
    )
    write(
        OUT_DIR / "采购单入库单汇总.xlsx",
        "入库单汇总",
        [
            "采购单号", "附件名", "批次", "实发金额", "各区块金额", "订单金额",
            "费用行明细", "费用行合计", "纯货值(实发-费用行)", "表尾补款批注(含是否计入成本)", "类型",
        ],
        summary,
    )


if __name__ == "__main__":
    main()
