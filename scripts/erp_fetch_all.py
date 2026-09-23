"""抓取睿贝 ERP 三类单据列表并落盘为 .cache/erp（贪婪、幂等）。

用法：
    python scripts/erp_fetch_all.py list          # 只抓三类列表（快）
    python scripts/erp_fetch_all.py list --year 2026
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import (  # noqa: E402
    CACHE_ROOT,
    now_iso,
    read_jsonl,
    save_meta,
    write_json,
    write_jsonl,
    write_samples,
)
from app.services.erp_lists import PURCHASES, SALE_ORDERS, SHIPMENTS  # noqa: E402
from app.services.erp_replay import ErpSession  # noqa: E402


ID_FIELDS = {
    "sale_orders": "orderId",
    "shipments": "shipmentId",
    "purchases": "purchase_id",
}


def fetch_list(session: ErpSession, spec, *, max_pages: int = 200, pause: float = 0.15) -> list[dict]:
    rows: list[dict] = []
    seen: set[str] = set()
    id_field = ID_FIELDS.get(spec.name, "id")
    page = 1
    while page <= max_pages:
        query = dict(spec.base_query)
        query[spec.page_param] = page
        payload = session.fetch_json(spec.path, query)
        from app.services.erp_replay import columns_to_rows

        page_rows = columns_to_rows(payload.get("root"))
        if not page_rows:
            break
        added = 0
        for row in page_rows:
            # 主键去重；分页边界重复时才会命中
            key = str(row.get(id_field) or json.dumps(row, ensure_ascii=False, sort_keys=True))
            if key in seen:
                continue
            seen.add(key)
            rows.append(row)
            added += 1
        total = int(payload.get("total") or 0)
        print(
            f"  {spec.name} page {page}: +{added}/{len(page_rows)} total={total} got={len(rows)}",
            flush=True,
        )
        if total and len(rows) >= total:
            break
        if len(page_rows) < 64 and not total:
            break
        page += 1
        time.sleep(pause)
    return rows


def year_of(row: dict, code_fields: list[str], date_fields: list[str]) -> str:
    for field in code_fields:
        value = str(row.get(field) or "")
        match = re.search(r"(20\d{2})", value)
        if match:
            return match.group(1)
        match = re.search(r"(?:^|[-_/])PI-?(\d{2})MT", value.upper())
        if match:
            return "20" + match.group(1)
        match = re.search(r"^(\d{2})MT", value.upper())
        if match:
            return "20" + match.group(1)
    for field in date_fields:
        value = str(row.get(field) or "")
        if len(value) >= 4 and value[:4].isdigit():
            return value[:4]
    return ""


def build_indexes(
    orders: list[dict], shipments: list[dict], purchases: list[dict]
) -> dict[str, dict]:
    order_index: dict[str, dict] = {}
    for row in orders:
        code = str(row.get("orderCode") or "").strip()
        if code:
            order_index[code] = {
                "orderId": row.get("orderId"),
                "comName": row.get("comName"),
                "contractDate": row.get("contractDate"),
                "amount": row.get("amount"),
                "currencyName": row.get("currencyName"),
            }

    shipment_index: dict[str, dict] = {}
    for row in shipments:
        entry = {
            "shipmentId": row.get("shipmentId"),
            "invoiceCode": row.get("invoiceCode"),
            "purchaseCode": row.get("purchaseCode"),
            "shipDate": row.get("shipDate"),
            "amount": row.get("amount"),
        }
        for key in (row.get("invoiceCode"), row.get("purchaseCode")):
            for token in str(key or "").replace(",", "&").split("&"):
                token = token.strip()
                if token:
                    shipment_index.setdefault(token, entry)

    purchase_index: dict[str, dict] = {}
    for row in purchases:
        code = str(row.get("purchase_code") or "").strip()
        if code:
            purchase_index[code] = {
                "purchase_id": row.get("purchase_id"),
                "orderCode": row.get("orderCode"),
                "supplierName": row.get("supplierName"),
                "purchase_date": row.get("purchase_date"),
                "amount": row.get("amount"),
            }
    return {
        "orders": order_index,
        "shipments": shipment_index,
        "purchases": purchase_index,
    }


def build_graph(orders: list[dict], shipments: list[dict], purchases: list[dict]) -> dict[str, dict]:
    """建立「一张单子查另外两张」的关联图。"""
    graph: dict[str, dict] = {}
    for row in orders:
        code = str(row.get("orderCode") or "").strip()
        if code:
            graph.setdefault(code, {"orderCode": code})["order"] = {
                "orderId": row.get("orderId"),
                "comName": row.get("comName"),
                "amount": row.get("amount"),
                "contractDate": row.get("contractDate"),
            }
    for row in shipments:
        entry_keys = set()
        for key in (row.get("invoiceCode"), row.get("purchaseCode")):
            for token in str(key or "").replace(",", "&").split("&"):
                token = token.strip()
                if token:
                    entry_keys.add(token)
        payload = {
            "shipmentId": row.get("shipmentId"),
            "invoiceCode": row.get("invoiceCode"),
            "purchaseCode": row.get("purchaseCode"),
            "shipDate": row.get("shipDate"),
            "amount": row.get("amount"),
        }
        for token in entry_keys:
            node = graph.setdefault(token, {"orderCode": token})
            node.setdefault("shipments", []).append(payload)
    for row in purchases:
        order_code = str(row.get("orderCode") or "").strip()
        payload = {
            "purchase_id": row.get("purchase_id"),
            "purchase_code": row.get("purchase_code"),
            "supplierName": row.get("supplierName"),
            "amount": row.get("amount"),
        }
        for token in {order_code, str(row.get("purchase_code") or "").strip()}:
            if not token:
                continue
            node = graph.setdefault(token, {"orderCode": token})
            node.setdefault("purchases", []).append(payload)
    return graph


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("action", nargs="?", default="list")
    parser.add_argument("--year", default="2026")
    parser.add_argument("--max-pages", type=int, default=200)
    parser.add_argument("--force", action="store_true", help="已有列表也重新抓取")
    args = parser.parse_args()

    started = now_iso()
    targets = {
        "sale_orders": CACHE_ROOT / "sale_orders" / "orders.jsonl",
        "shipments": CACHE_ROOT / "shipments" / "shipments.jsonl",
        "purchases": CACHE_ROOT / "purchases" / "purchases.jsonl",
    }

    def cached(spec) -> list[dict] | None:
        if args.force:
            return None
        rows = read_jsonl(targets[spec.name])
        return rows or None

    with ErpSession() as session:
        print("[1/3] 外销订单列表", flush=True)
        orders = cached(SALE_ORDERS) or fetch_list(session, SALE_ORDERS, max_pages=args.max_pages)
        print(f"  sale_orders: {len(orders)} 条", flush=True)
        print("[2/3] 出运单列表", flush=True)
        shipments = cached(SHIPMENTS) or fetch_list(session, SHIPMENTS, max_pages=args.max_pages)
        print(f"  shipments: {len(shipments)} 条", flush=True)
        print("[3/3] 采购订单列表", flush=True)
        purchases = cached(PURCHASES) or fetch_list(session, PURCHASES, max_pages=args.max_pages)
        print(f"  purchases: {len(purchases)} 条", flush=True)

    for row in orders:
        row["_year"] = year_of(row, ["orderCode"], ["contractDate"])
    for row in shipments:
        row["_year"] = year_of(row, ["invoiceCode", "purchaseCode"], ["shipDate"])
    for row in purchases:
        row["_year"] = year_of(row, ["purchase_code", "orderCode"], ["purchase_date"])

    write_jsonl(CACHE_ROOT / "sale_orders" / "orders.jsonl", orders)
    write_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl", shipments)
    write_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl", purchases)
    write_samples(orders, CACHE_ROOT / "sale_orders" / "samples")
    write_samples(shipments, CACHE_ROOT / "shipments" / "samples")
    write_samples(purchases, CACHE_ROOT / "purchases" / "samples")

    indexes = build_indexes(orders, shipments, purchases)
    write_json(CACHE_ROOT / "sale_orders" / "index.json", indexes["orders"])
    write_json(CACHE_ROOT / "shipments" / "index.json", indexes["shipments"])
    write_json(CACHE_ROOT / "purchases" / "index.json", indexes["purchases"])
    write_json(CACHE_ROOT / "links" / "order_graph.json", build_graph(orders, shipments, purchases))

    counts = {
        "sale_orders": len(orders),
        "shipments": len(shipments),
        "purchases": len(purchases),
        "sale_orders_2026": sum(1 for r in orders if r["_year"] == args.year),
        "shipments_2026": sum(1 for r in shipments if r["_year"] == args.year),
        "purchases_2026": sum(1 for r in purchases if r["_year"] == args.year),
    }
    save_meta(
        {
            "started_at": started,
            "finished_at": now_iso(),
            "counts": counts,
            "params": vars(args),
        }
    )
    print(json.dumps(counts, ensure_ascii=False))


if __name__ == "__main__":
    main()
