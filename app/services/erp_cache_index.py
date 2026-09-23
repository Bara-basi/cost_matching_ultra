"""睿贝缓存索引：按单号直查三类单据（供成本匹配环节使用）。

索引放在 .cache/erp/index/，条目形如：
    {"orderCode": "...", "orderId": "...", "shipmentIds": [...], "purchaseIds": [...]}
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT, read_jsonl, write_json

INDEX_DIR = CACHE_ROOT / "index"
NOISE_KEYS = {"合计", "小计", "总计"}


def _tokens(value: Any) -> list[str]:
    """把 `A,B` / `A&B` 等写法拆成单号列表。"""
    text = str(value or "").strip()
    if not text:
        return []
    parts = re.split(r"[,&;，、]", text)
    return [p.strip() for p in parts if p.strip() and p.strip() not in NOISE_KEYS]


def build_indexes() -> dict[str, int]:
    orders = read_jsonl(CACHE_ROOT / "sale_orders" / "orders.jsonl")
    shipments = read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl")
    purchases = read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")

    nodes: dict[str, dict[str, Any]] = {}

    def node(key: str) -> dict[str, Any]:
        return nodes.setdefault(key, {"orderCode": key})

    for row in orders:
        code = str(row.get("orderCode") or "").strip()
        if not code:
            continue
        node(code)["order"] = {
            "orderId": row.get("orderId"),
            "comName": row.get("comName"),
            "contractDate": row.get("contractDate"),
            "amount": row.get("amount"),
            "currencyName": row.get("currencyName"),
        }

    for row in shipments:
        payload = {
            "shipmentId": row.get("shipmentId"),
            "invoiceCode": row.get("invoiceCode"),
            "purchaseCode": row.get("purchaseCode"),
            "shipDate": row.get("shipDate"),
            "amount": row.get("amount"),
        }
        keys = _tokens(row.get("invoiceCode")) + _tokens(row.get("purchaseCode"))
        for key in keys:
            node(key).setdefault("shipments", []).append(payload)

    for row in purchases:
        payload = {
            "purchase_id": row.get("purchase_id"),
            "purchase_code": row.get("purchase_code"),
            "supplierName": row.get("supplierName"),
            "purchase_date": row.get("purchase_date"),
            "amount": row.get("amount"),
        }
        keys = _tokens(row.get("orderCode")) + _tokens(row.get("purchase_code"))
        for key in keys:
            node(key).setdefault("purchases", []).append(payload)

    # 去重同一 shipment/purchase 被多次登记
    for value in nodes.values():
        for field in ("shipments", "purchases"):
            if field in value:
                seen = set()
                unique = []
                for item in value[field]:
                    marker = json.dumps(item, ensure_ascii=False, sort_keys=True)
                    if marker in seen:
                        continue
                    seen.add(marker)
                    unique.append(item)
                value[field] = unique

    write_json(INDEX_DIR / "order_index.json", nodes)
    return {
        "nodes": len(nodes),
        "with_order": sum(1 for v in nodes.values() if v.get("order")),
        "with_purchases": sum(1 for v in nodes.values() if v.get("purchases")),
        "with_shipments": sum(1 for v in nodes.values() if v.get("shipments")),
        "complete": sum(
            1
            for v in nodes.values()
            if v.get("order") and v.get("purchases") and v.get("shipments")
        ),
    }


def load_index() -> dict[str, dict[str, Any]]:
    path = INDEX_DIR / "order_index.json"
    if not path.exists():
        build_indexes()
    return json.loads(path.read_text(encoding="utf-8"))


def lookup(order_code: str) -> dict[str, Any]:
    """按任意单号（外销订单号 / 采购单号 / 外销发票号）查关联。"""
    index = load_index()
    key = (order_code or "").strip()
    return index.get(key, {})


if __name__ == "__main__":
    print(build_indexes())
