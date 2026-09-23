"""睿贝 ERP 各列表接口的调用定义与分页抓取。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterator

from app.services.erp_replay import ErpSession, ListResult


@dataclass
class ListSpec:
    """一个列表接口的定义。"""

    name: str
    path: str
    base_query: dict[str, Any] = field(default_factory=dict)
    page_param: str = "p"


# 外销订单（业务中心 → 外销订单）
SALE_ORDERS = ListSpec(
    name="sale_orders",
    path="/saleOrder_list",
    base_query={
        "condition": "all",
        "contractDate_start": "",
        "contractDate_end": "",
        "onCreate_start": "",
        "onCreate_end": "",
        "userDefaultTableName": "biz_sale_orders",
        "searchValue": "",
    },
)

# 出运明细单（业务中心 → 出运明细单）
SHIPMENTS = ListSpec(
    name="shipments",
    path="/shipment_select",
    base_query={
        "condition": "all",
        "userDefaultTableName": "biz_shipment",
    },
)

# 采购记录（工作台 → 采购单，含入库单标记）
PURCHASES = ListSpec(
    name="purchases",
    path="/purchase_selectPur",
    base_query={
        "purchaseIsOutSale": "all",
        "condition": "all",
        "userDefaultTableName": "biz_purchases",
        "comePage": "grn",
    },
)

# 采购记录（仅含入库单附件的那批，用于附件下载）
PURCHASES_WITH_GRN = ListSpec(
    name="purchases",
    path="/purchase_selectPur",
    base_query={
        "purchaseIsOutSale": "all",
        "grnChoose": "Y",
        "condition": "ordered",
        "userDefaultTableName": "biz_purchases",
        "comePage": "grn",
    },
)

# 某个外销订单下的出运明细
SHIPMENT_BY_ORDER = "/shipmentItem_selectShipFollow"
# 某个外销订单下的采购记录
PURCHASE_BY_ORDER = "/purchaseItems_listPlaceOrder"


def fetch_all(session: ErpSession, spec: ListSpec, *, max_pages: int = 500) -> tuple[list[dict], int]:
    """按分页抓完一个列表接口，返回 (rows, total)。"""
    rows: list[dict] = []
    total = 0
    seen_pages = 0
    for result in session.iter_list(spec.path, spec.base_query, page_param=spec.page_param, max_pages=max_pages):
        total = result.total or total
        if not result.rows:
            break
        rows.extend(result.rows)
        seen_pages += 1
        if total and len(rows) >= total:
            break
    return rows, total


def fetch_by_order(
    session: ErpSession,
    path: str,
    order_id: str,
    table_name: str,
    *,
    max_pages: int = 50,
    extra: dict[str, Any] | None = None,
) -> list[dict]:
    """抓取挂在某个外销订单下的明细（出运 / 采购）。"""
    query: dict[str, Any] = {"order_id": order_id, "userDefaultTableName": table_name}
    if extra:
        query.update(extra)
    rows: list[dict] = []
    for result in session.iter_list(path, query, max_pages=max_pages):
        if not result.rows:
            break
        rows.extend(result.rows)
        if result.total and len(rows) >= result.total:
            break
    return rows
