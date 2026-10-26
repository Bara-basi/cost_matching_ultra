"""“仅供核对用”预录报关单解析。

预录单复用正式报关单的文本版式，但 pypdf 对一票多商品的数量排序并不可靠。
例如“不锈钢螺栓 + 镍合金丝”的 PDF 文本会按 47.5、47.5、2.325、2.325
列出两个商品的重量，顺序却与商品行相反。

这里保留现有版式识别和表头解析，再用 ``总价 / 单价`` 反推每个商品数量，
从文本候选值中选择最匹配的一项；只有每个商品在数量与单位上同时匹配时才
覆盖原结果，匹配不到则保留原值并写入警告。
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.services.parser.declaration_parser import (
    ITEM_V2_RE,
    parse_declaration,
)
from app.services.parser.models import DeclarationItem, ParsedDeclaration
from app.services.parser.pdf_text import extract_text


QTY_TOKEN_RE = re.compile(r"^([\d,.]+)\s*([^\d\s]+)$")
NUMBER_RE = re.compile(r"^-?[\d,]+(?:\.\d+)?$")


def _number(value: str) -> Decimal | None:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        result = Decimal(text)
    except InvalidOperation:
        return None
    return result if result.is_finite() else None


def quantity_candidates(text: str) -> list[tuple[Decimal, str]]:
    """抽取商品区上方由文本层列出的“数量+单位”候选值。"""
    out: list[tuple[Decimal, str]] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if ITEM_V2_RE.match(line):
            break
        match = QTY_TOKEN_RE.match(line)
        if not match:
            continue
        number = _number(match.group(1))
        unit = match.group(2).strip()
        if number is None or not unit or unit in ("/", "页"):
            continue
        pair = (number, unit)
        if pair not in out:
            out.append(pair)
    return out


def _price_quantity(item: DeclarationItem) -> Decimal | None:
    unit_price = _number(item.unit_price)
    total_price = _number(item.total_price)
    if unit_price is None or total_price is None or unit_price <= 0:
        return None
    return total_price / unit_price


def correct_item_quantities(
    text: str, items: list[DeclarationItem]
) -> tuple[list[DeclarationItem], list[str]]:
    """用价格关系纠正预录单的文本数量顺序。

    仅当反推值和一个候选数量在绝对 0.02 或千分之一以内一致时才覆盖。
    这样可以同时支持千克、米、件等不同单位，又不会把税费或单价误当数量。
    """
    candidates = quantity_candidates(text)
    warnings: list[str] = []
    if not candidates or not items:
        return items, warnings
    corrected: list[DeclarationItem] = []
    for position, item in enumerate(items, 1):
        derived = _price_quantity(item)
        if derived is None:
            corrected.append(item)
            continue
        matches = sorted(
            candidates,
            key=lambda pair: abs(pair[0] - derived),
        )
        value, unit = matches[0]
        tolerance = max(Decimal("0.02"), abs(derived) * Decimal("0.001"))
        if abs(value - derived) > tolerance:
            warnings.append(
                f"第 {position} 个商品无法用总价/单价匹配文本数量，保留原值"
            )
            corrected.append(item)
            continue
        quantity = f"{value:f}{unit}"
        corrected.append(
            DeclarationItem(
                serial=item.serial,
                hs_code=item.hs_code,
                product_name=item.product_name,
                spec=item.spec,
                quantity=quantity,
                unit=unit,
                destination_country=item.destination_country,
                unit_price=item.unit_price,
                total_price=item.total_price,
                currency=item.currency,
                declare_quantity=quantity,
                declare_unit=unit,
                second_quantity=item.second_quantity,
                second_unit=item.second_unit,
            )
        )
    return corrected, warnings


def parse_pre_record(path: Path) -> ParsedDeclaration:
    """解析一份预录单，并修正一票多商品时的数量顺序。"""
    parsed = parse_declaration(path)
    if parsed.header.sheet_type != "预录单（仅供核对用）":
        raise ValueError(f"{path.name} 不是“仅供核对用”预录单")
    text = extract_text(path)
    parsed.items, quantity_warnings = correct_item_quantities(text, parsed.items)
    parsed.warnings.extend(quantity_warnings)
    # 预录单的“杂费保费运费”合并标签可能把上一行合同号误当杂费。
    if parsed.header.misc_fee == parsed.header.contract_raw:
        parsed.header.misc_fee = ""
    return parsed
