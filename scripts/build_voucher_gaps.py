"""凭证缺口清单：睿贝采购单表里的单，哪些拿不到「入库单」这种成本凭证，影响多少钱。

分类原因：
  1. 睿贝无入库记录（grnQuantity=0 且无入库日期）——收货还没发生；
  2. 有附件但没有入库单（只有合同/发票/结算单）——供应链代采常见；
  3. 未抓到附件列表——抓取缺口，需要补抓。

输出：outputs/grn_attachments/凭证缺口清单.xlsx
"""
from __future__ import annotations

import json
import re
import sys
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from app.services.cost_match import unit_core  # noqa: E402
from app.services.erp_cache import CACHE_ROOT, read_jsonl  # noqa: E402
from app.services.grn_extract import PARSED_ROOT  # noqa: E402
from app.services.grn_select import ORDER_KEY_RE, erp_supplier  # noqa: E402
from app.services.supplier_names import SupplierNames  # noqa: E402

LIST_DIR = CACHE_ROOT / "attachments" / "lists"
OUT = PROJECT_ROOT / "outputs" / "grn_attachments" / "凭证缺口清单.xlsx"
COST_WORDS = ("结算", "对账", "送货", "收货", "磅单", "过磅", "发货清单", "结算单")
FILLS = {
    "睿贝无入库记录（未收货）": PatternFill("solid", fgColor="EDEDED"),
    "附件里没有入库单（只有合同/发票/结算单）": PatternFill("solid", fgColor="FFF2CC"),
    "只差结算单（可补抓）": PatternFill("solid", fgColor="DDEBF7"),
    "未抓到附件列表（需补抓）": PatternFill("solid", fgColor="FFC7CE"),
}


def dec(value) -> Decimal:
    try:
        return Decimal(str(value or 0))
    except Exception:
        return Decimal(0)


def norm(text: str) -> str:
    return str(text or "").replace(" ", "").replace("_", "").upper()


def loose_key(text: str, supplier: str = "") -> tuple:
    """宽松键：单号（去年份）+ ADD + 供应商简称。

    睿贝的采购单号会带「原材料合同 / （补流程）/（已付款）」这类尾巴，
    和附件目录名对不上，只按整串比会把有附件的单误报成缺口。
    """
    match = ORDER_KEY_RE.search(str(text or ""))
    if not match:
        return ("", "", supplier)
    body = str(match.group(2) or "").upper()          # 部门+业务员+流水
    add = str(match.group(4) or "").upper()
    if add == "ADD":
        add = "ADD1"
    return (body, add, supplier)


def main() -> None:
    names = SupplierNames()
    rows = read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
    lists = {norm(p.stem): p for p in LIST_DIR.glob("*.json")}
    parsed = {norm(p.name) for p in PARSED_ROOT.iterdir() if p.is_dir()}
    parsed_loose = {
        loose_key(p.name, names.short(erp_supplier(p.name)))
        for p in PARSED_ROOT.iterdir()
        if p.is_dir()
    }

    records = json.loads(
        (PROJECT_ROOT / "data" / "cache" / "ref_2026_customs_full.json").read_text(encoding="utf-8")
    ).get("records") or []
    ref_by_core: dict[tuple[str, str], Decimal] = {}
    for record in records:
        core = unit_core(record.get("合同号（应收表格）") or record.get("合同号_1"))
        if core and record.get("采购金额"):
            key = (core, str(record.get("供应商简称") or "").strip())
            ref_by_core[key] = ref_by_core.get(key, Decimal(0)) + dec(record.get("采购金额"))

    output: list[list] = []
    stats: dict[str, list] = {}
    for row in rows:
        code = str(row.get("purchase_code") or "")
        key = norm(code)
        supplier = names.short(str(row.get("supplierName") or ""))
        if key in parsed or loose_key(code, supplier) in parsed_loose:
            continue
        amount = dec(row.get("amount"))
        grn_qty = dec(row.get("grnQuantity"))
        last_grn = str(row.get("lastGrnDate") or "").strip()
        list_path = lists.get(key)
        attachment_names: list[str] = []
        if list_path is not None:
            try:
                payload = json.loads(list_path.read_text(encoding="utf-8"))
                attachment_names = [
                    str(item.get("attachmentName") or "")
                    for item in payload.get("attachmentList") or []
                ]
            except ValueError:
                attachment_names = []
        has_settlement = [n for n in attachment_names if any(w in n for w in COST_WORDS)]
        if grn_qty == 0 and not last_grn:
            kind = "睿贝无入库记录（未收货）"
        elif list_path is None:
            kind = "未抓到附件列表（需补抓）"
        elif has_settlement and not any("入库单" in n for n in attachment_names):
            kind = "只差结算单（可补抓）"
        else:
            kind = "附件里没有入库单（只有合同/发票/结算单）"
        ref_money = ref_by_core.get((unit_core(code), supplier), Decimal(0))
        bucket = stats.setdefault(kind, [0, Decimal(0), Decimal(0)])
        bucket[0] += 1
        bucket[1] += amount
        bucket[2] += ref_money
        output.append(
            [
                code,
                supplier,
                str(amount),
                str(ref_money) if ref_money else "",
                str(grn_qty),
                last_grn,
                kind,
                "；".join(has_settlement) if has_settlement else "",
            ]
        )

    order = [
        "只差结算单（可补抓）",
        "未抓到附件列表（需补抓）",
        "附件里没有入库单（只有合同/发票/结算单）",
        "睿贝无入库记录（未收货）",
    ]
    output.sort(key=lambda row: (order.index(row[6]) if row[6] in order else 9, -float(row[2] or 0)))
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "凭证缺口"
    sheet.append(
        ["采购单号", "供应商", "睿贝采购金额", "飞书同单采购金额", "入库数量", "最后入库日", "缺口原因", "可补抓的文件"]
    )
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in output:
        sheet.append(row)
        fill = FILLS.get(str(row[6]))
        if fill:
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    for index, name in enumerate(sheet[1], 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            14, min(60, len(str(name)) * 2 + 8)
        )
    sheet.freeze_panes = "A2"

    summary = workbook.create_sheet("汇总")
    summary.append(["缺口原因", "采购单数", "睿贝金额合计", "飞书金额合计"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    for kind in order:
        if kind not in stats:
            continue
        count, amount, ref_money = stats[kind]
        summary.append([kind, count, str(amount), str(ref_money)])
    for index, name in enumerate(summary[1], 1):
        summary.column_dimensions[summary.cell(row=1, column=index).column_letter].width = max(
            18, len(str(name)) * 2 + 8
        )
    OUT.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUT)

    print(f"拿不到入库单的采购单 {len(output)} 张")
    for kind in order:
        if kind in stats:
            count, amount, ref_money = stats[kind]
            print(f"  {kind:<32} {count:>4} 张，睿贝 {amount}，飞书 {ref_money}")
    print("输出：", OUT)


if __name__ == "__main__":
    main()
