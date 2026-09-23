"""报关单解析结果的数据结构。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class DeclarationHeader:
    """报关单表头字段（单证级）。"""

    declaration_no: str = ""
    pre_entry_no: str = ""
    contract_raw: str = ""
    contract_remark: str = ""
    sheet_type: str = "出口退税联"
    export_date: str = ""
    declare_date: str = ""
    export_port: str = ""
    transport_mode: str = ""
    trade_mode: str = ""
    deal_mode: str = ""
    # 报关单上登记的运费/保费/杂费（原文，如 `USD/2800/总价`）
    freight: str = ""
    insurance: str = ""
    misc_fee: str = ""
    destination_country: str = ""
    destination_port: str = ""
    domestic_source: str = ""
    consignor: str = ""
    producer: str = ""
    package_kind: str = ""
    pieces: str = ""
    gross_weight: str = ""
    net_weight: str = ""
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class DeclarationItem:
    """报关单商品项。"""

    serial: int = 0
    hs_code: str = ""
    product_name: str = ""
    spec: str = ""
    quantity: str = ""
    unit: str = ""
    destination_country: str = ""
    unit_price: str = ""
    total_price: str = ""
    currency: str = ""
    declare_quantity: str = ""
    declare_unit: str = ""


@dataclass
class ParsedDeclaration:
    """一份报关单的完整解析结果。"""

    source_file: str
    header: DeclarationHeader
    items: list[DeclarationItem]
    warnings: list[str] = field(default_factory=list)
