"""把解析结果写成固定格式的 .xlsx。"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from app.services.parser.models import ParsedDeclaration
from app.services.parser.output_columns import OUTPUT_COLUMNS, derive_contract

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)


def _to_excel_date(value: str):
    """yyyy-MM-dd -> datetime（Excel 日期列）。"""
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return value


def _split_quantity(quantity: str) -> tuple[str, str]:
    """`106千克` -> ('106', '千克')。"""
    match = re.match(r"^([\d,.]+)\s*(.*)$", (quantity or "").strip())
    if not match:
        return (quantity or "", "")
    return (match.group(1), match.group(2).strip())


WEIGHT_UNITS = ("千克", "公斤", "KG", "kg", "KGS", "公斤/千克")


def _is_weight_unit(unit: str) -> bool:
    """这个单位是不是重量（报关单上「申报数量/申报单位」可能是套/件/个）。"""
    text = (unit or "").strip()
    if not text:
        return False
    return any(key in text for key in WEIGHT_UNITS)


def build_rows(parsed: ParsedDeclaration) -> list[dict]:
    """把一份报关单展开成「一条商品一行」的字典列表。"""
    header = parsed.header
    contract_raw = header.contract_raw
    contract = derive_contract(contract_raw)
    rows: list[dict] = []
    warning = "；".join(parsed.warnings)
    for item in parsed.items or [None]:
        base = {
            "来源文件": parsed.source_file,
            "单据类型": header.sheet_type,
            "合同号_1": contract_raw,
            "合同号（应收表格）": contract_raw,
            "合同号": contract,
            "报关单号": header.declaration_no,
            "出口日期": _to_excel_date(header.export_date),
            "申报日期": _to_excel_date(header.declare_date),
            "运输方式": header.transport_mode,
            "成交方式": header.deal_mode,
            "运费": header.freight,
            "保费": header.insurance,
            "杂费": header.misc_fee,
            "贸易方式": header.trade_mode,
            "出口口岸": header.export_port,
            "指运港": header.destination_port,
            "境内货源地": header.domestic_source,
            "境内收发货人": header.consignor,
            "生产销售单位": header.producer,
            "件数": header.pieces,
            "包装种类": header.package_kind,
            "毛重": header.gross_weight,
            "净重": header.net_weight,
            "解析警告": warning,
        }
        if header.contract_remark and header.contract_remark != contract_raw:
            base["解析警告"] = "；".join(filter(None, [warning, "合同号取自备注/续行"]))
        if item is not None:
            # 报关重量优先取「申报数量」；它是套装/件这类非重量单位时，
            # 再退回「第二数量」（报关单常把重量放在这一栏），最后才用法定数量。
            weight, weight_unit = _split_quantity(item.declare_quantity or item.quantity)
            if not _is_weight_unit(weight_unit):
                alt_weight, alt_unit = _split_quantity(item.second_quantity)
                if alt_weight and _is_weight_unit(alt_unit):
                    weight, weight_unit = alt_weight, alt_unit
            base.update(
                {
                    "商品序号": item.serial,
                    "报关品名": item.product_name,
                    "海关编码": item.hs_code,
                    "报关重量": weight,
                    "报关重量单位": weight_unit or item.unit,
                    "申报数量": item.declare_quantity,
                    "申报单位": item.declare_unit,
                    "目的国": item.destination_country,
                    "单价": item.unit_price,
                    "总价": item.total_price,
                    "币种": item.currency,
                }
            )
        rows.append(base)
    return rows


def write_workbook(rows: list[dict], path: Path, sheet_name: str = "报关解析结果") -> Path:
    """写出固定列顺序的 Excel。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(list(OUTPUT_COLUMNS))
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(column, "") for column in OUTPUT_COLUMNS])
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), OUTPUT_COLUMNS):
        width = max(10, min(28, len(name) * 2))
        sheet.column_dimensions[column[0].column_letter].width = width
    sheet.freeze_panes = "A2"
    workbook.save(path)
    return path
