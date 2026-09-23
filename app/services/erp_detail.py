"""产品行明细抓取：按订单/出运/采购逐单取明细并落盘。

明细接口（实测）：
- `/shipmentItem_selectShipFollow`  form: order_id, userDefaultTableName=salefollow_shipment
- `/purchaseItems_listPlaceOrder`   form: order_id
返回的 `itemList.*` 是**带前缀的列名**，一行 = 一张出运/采购单里的一条产品行。
"""
from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT, ensure_dir
from app.services.erp_replay import ErpSession, columns_to_rows

DETAIL_DIR = CACHE_ROOT / "details"

SHIPMENT_ITEMS = "/shipmentItem_selectShipFollow"
PURCHASE_ITEMS = "/purchaseItems_listPlaceOrder"


def fetch_order_details(
    session: ErpSession,
    order_id: str,
    order_code: str,
) -> dict[str, list[dict[str, Any]]]:
    """抓一个订单对应的出运明细与采购明细。"""
    shipments: list[dict[str, Any]] = []
    purchases: list[dict[str, Any]] = []
    for path, table, bucket in (
        (SHIPMENT_ITEMS, "salefollow_shipment", shipments),
        (PURCHASE_ITEMS, "purchase_order_waiting_purchase", purchases),
    ):
        page = 1
        while page <= 50:
            query = {"order_id": order_id, "p": page}
            if table:
                query["userDefaultTableName"] = table
            payload = session.fetch_json(path, query)
            rows = columns_to_rows(payload.get("root"))
            if not rows:
                break
            for row in rows:
                row["_orderCode"] = order_code
                row["_orderId"] = order_id
            bucket.extend(rows)
            total = int(payload.get("total") or 0)
            if total and len(bucket) >= total:
                break
            page += 1
            time.sleep(0.05)
    return {"shipment_items": shipments, "purchase_items": purchases}


def detail_path(kind: str, code: str) -> Path:
    safe = str(code).replace("/", "_").replace("\\", "_").strip() or "unknown"
    return DETAIL_DIR / kind / f"{safe}.json"


def load_detail(kind: str, code: str) -> dict[str, Any] | None:
    path = detail_path(kind, code)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None


def save_detail(kind: str, code: str, payload: dict[str, Any]) -> Path:
    path = detail_path(kind, code)
    ensure_dir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return path
