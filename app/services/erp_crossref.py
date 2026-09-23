"""三单互查（外销订单 ↔ 出运单 ↔ 采购单）——**只使用睿贝原生关联字段**。

证据来源（全部是睿贝自己存的外键，不做任何单号前缀猜测）：

1. 外销订单列表 `orderId`（销售订单主键）。
2. 出运明细接口 `POST /shipmentItem_selectShipFollow`
   形式参数 `order_id=<orderId>`，返回行含 `shipmentId` 与 `invoiceCode`。
   —— 这是「订单 ↔ 出运单」的原生关系。
3. 采购明细接口 `POST /purchaseItems_listPlaceOrder`
   形式参数 `order_id=<orderId>`，返回行含 `purchaseId` 与 `purchaseCode`。
   —— 这是「订单 ↔ 采购单」的原生关系。
4. 出运单头 `purchaseCode`（存的是它包含的采购单号）——「出运单 ↔ 采购单」的原生关系。

明细接口把一个单据的产品行按 `单据ID + 若干产品行` 返回：
首行带 `shipmentId`/`purchaseId`，同单据的后续产品行该字段为空，因此按空值向下继承。
"""
from __future__ import annotations

import json
from collections import defaultdict
from typing import Any, Iterator

from app.services.erp_cache import CACHE_ROOT, read_jsonl
from app.services.erp_detail import DETAIL_DIR

# 出运/采购明细里承载「所属单据 ID」的原生字段
SHIPMENT_ID_FIELD = "shipmentId"
PURCHASE_ID_FIELD = "purchaseId"
INVOICE_CODE_FIELD = "invoiceCode"
PURCHASE_CODE_FIELD = "purchaseCode"


def norm(value: Any) -> str:
    """仅做大小写/空白归一，用于比较同一个单据号的不同写法（不做任何推断）。"""
    return "".join(str(value or "").split()).upper()


def strip_pi(value: Any) -> str:
    """去掉睿贝部分字段带的 `PI-` 前缀（同一订单号的两种存法）。"""
    text = norm(value)
    return text[3:] if text.startswith("PI-") else text


def _split_codes(value: Any) -> list[str]:
    """出运单头 `purchaseCode` 原生存储的是多个采购单号，用 `,`/`&` 连接。"""
    raw = str(value or "").strip()
    if not raw:
        return []
    out: list[str] = []
    for chunk in raw.replace("&", ",").replace(";", ",").replace("；", ",").split(","):
        token = chunk.strip()
        if token and token not in out:
            out.append(token)
    return out


def iter_shipment_groups(payload: dict[str, Any]) -> Iterator[dict[str, Any]]:
    """把一个订单的出运明细按 `shipmentId` 分组（首行带 ID，后续行继承）。"""
    current: dict[str, Any] | None = None
    for row in payload.get("shipment_items") or []:
        sid = str(row.get(SHIPMENT_ID_FIELD) or "").strip()
        code = str(row.get(INVOICE_CODE_FIELD) or "").strip()
        if sid:
            if current:
                yield current
            current = {
                "shipmentId": sid,
                "invoiceCode": code,
                "shipDate": row.get("shipDate"),
                "shipPortName": row.get("shipPortName"),
                "transportTypeName": row.get("transportTypeName"),
                "items": [],
            }
        if current is None:
            continue
        current["items"].append(row)
    if current:
        yield current


def iter_purchase_groups(payload: dict[str, Any]) -> Iterator[dict[str, Any]]:
    """把一个订单的采购明细按 `purchaseId` 分组。"""
    current: dict[str, Any] | None = None
    for row in payload.get("purchase_items") or []:
        pid = str(row.get(PURCHASE_ID_FIELD) or "").strip()
        code = str(row.get(PURCHASE_CODE_FIELD) or "").strip()
        if pid:
            if current:
                yield current
            current = {
                "purchaseId": pid,
                "purchaseCode": code,
                "purchaseDate": row.get("purchaseDate"),
                "comName": row.get("comName"),
                "items": [],
            }
        if current is None:
            continue
        current["items"].append(row)
    if current:
        yield current


class CrossRef:
    """基于原生外键的三单互查。"""

    def __init__(self) -> None:
        self.orders: dict[str, dict[str, Any]] = {}
        self.shipments: dict[str, dict[str, Any]] = {}
        self.purchases: dict[str, dict[str, Any]] = {}
        # 关联表：orderId -> 出运/采购单据ID
        self.order_shipments: dict[str, list[str]] = defaultdict(list)
        self.order_purchases: dict[str, list[str]] = defaultdict(list)
        # 反向索引：单据ID -> orderId
        self.shipment_orders: dict[str, list[str]] = defaultdict(list)
        self.purchase_orders: dict[str, list[str]] = defaultdict(list)
        # 出运单头 purchaseCode 建立的 出运→采购 原生关系
        self.shipment_purchase_codes: dict[str, list[str]] = defaultdict(list)

    # ---------- 组装 ----------

    def load(self) -> "CrossRef":
        self._load_orders()
        self._load_shipment_headers()
        self._load_details()
        self._link_shipment_order_codes()
        return self

    def _load_orders(self) -> None:
        for row in read_jsonl(CACHE_ROOT / "sale_orders" / "orders.jsonl"):
            oid = str(row.get("orderId") or "").strip()
            if not oid:
                continue
            self.orders[oid] = row

    def _load_shipment_headers(self) -> None:
        """出运单头：`orderCode` 是它包含的订单号，`purchaseCode` 是它包含的采购单号。"""
        for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
            sid = str(row.get("shipmentId") or "").strip()
            if not sid:
                continue
            self.shipments[sid] = row
            self.shipment_purchase_codes[sid] = _split_codes(row.get("purchaseCode"))

    def _load_details(self) -> None:
        """明细接口是「订单 ↔ 出运/采购」的原生关系来源。"""
        # 采购列表里有供应商，用 purchase_id 建索引补齐（明细的 comName 是客户名）
        supplier_by_purchase: dict[str, dict[str, Any]] = {}
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            pid = str(row.get("purchase_id") or "").strip()
            if pid:
                supplier_by_purchase[pid] = row

        if not DETAIL_DIR.exists():
            return
        for path in sorted((DETAIL_DIR / "orders").glob("*.json")):
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except ValueError:
                continue
            order_id = str(payload.get("_orderId") or "").strip()
            if not order_id:
                continue

            for group in iter_shipment_groups(payload):
                sid = group["shipmentId"]
                entry = self.shipments.setdefault(
                    sid,
                    {"shipmentId": sid, "invoiceCode": group.get("invoiceCode")},
                )
                entry.setdefault("itemsByOrder", {})[order_id] = group["items"]
                if sid not in self.order_shipments[order_id]:
                    self.order_shipments[order_id].append(sid)
                if order_id not in self.shipment_orders[sid]:
                    self.shipment_orders[sid].append(order_id)

            for group in iter_purchase_groups(payload):
                pid = group["purchaseId"]
                entry = self.purchases.setdefault(
                    pid,
                    {"purchaseId": pid, "purchaseCode": group.get("purchaseCode")},
                )
                listed = supplier_by_purchase.get(pid, {})
                entry["supplierName"] = entry.get("supplierName") or listed.get("supplierName")
                entry["amount"] = entry.get("amount") or listed.get("amount")
                entry["purchase_date"] = entry.get("purchase_date") or listed.get("purchase_date")
                entry.setdefault("itemsByOrder", {})[order_id] = group["items"]
                if pid not in self.order_purchases[order_id]:
                    self.order_purchases[order_id].append(pid)
                if order_id not in self.purchase_orders[pid]:
                    self.purchase_orders[pid].append(order_id)

    def _link_shipment_order_codes(self) -> None:
        """出运单头 `orderCode` 是睿贝原生存的订单号，用它补齐「出运单 → 订单」。"""
        by_code: dict[str, str] = {}
        for oid, row in self.orders.items():
            code = strip_pi(row.get("orderCode"))
            if code:
                by_code.setdefault(code, oid)
        for sid, row in self.shipments.items():
            for code in _split_codes(row.get("orderCode")):
                order_id = by_code.get(strip_pi(code))
                if not order_id:
                    continue
                if order_id not in self.shipment_orders[sid]:
                    self.shipment_orders[sid].append(order_id)
                if sid not in self.order_shipments[order_id]:
                    self.order_shipments[order_id].append(sid)

    # ---------- 查询 ----------

    def order(self, order_code: str) -> dict[str, Any] | None:
        target = strip_pi(order_code)
        for row in self.orders.values():
            if strip_pi(row.get("orderCode")) == target:
                return row
        return None

    def lookup(self, code: str) -> dict[str, Any]:
        """按订单号（或 `PI-订单号`）查三单；其余单号请用 shipments_of / purchases_of。"""
        order = self.order(code)
        if not order:
            return {"order": None, "shipments": [], "purchases": []}
        oid = str(order.get("orderId"))
        shipments = [self.shipments[s] for s in self.order_shipments.get(oid, [])]
        purchases = [self.purchases[p] for p in self.order_purchases.get(oid, [])]
        return {"order": order, "shipments": shipments, "purchases": purchases}

    def shipments_of_order(self, order_code: str) -> list[dict[str, Any]]:
        return self.lookup(order_code)["shipments"]

    def purchases_of_order(self, order_code: str) -> list[dict[str, Any]]:
        return self.lookup(order_code)["purchases"]

    def orders_of_shipment(self, shipment_id: str) -> list[dict[str, Any]]:
        return [
            self.orders[oid]
            for oid in self.shipment_orders.get(str(shipment_id), [])
            if oid in self.orders
        ]

    def purchases_of_shipment(self, shipment_id: str) -> list[dict[str, Any]]:
        """出运单 → 采购单：先用睿贝原生存的 `purchaseCode`，再用订单关联补全。"""
        sid = str(shipment_id)
        found: list[dict[str, Any]] = []
        wanted = {strip_pi(c) for c in self.shipment_purchase_codes.get(sid, [])}
        for pid, entry in self.purchases.items():
            if strip_pi(entry.get("purchaseCode")) in wanted:
                found.append(entry)
        if found:
            return found
        for oid in self.shipment_orders.get(sid, []):
            for pid in self.order_purchases.get(oid, []):
                if pid in self.purchases and self.purchases[pid] not in found:
                    found.append(self.purchases[pid])
        return found

    def shipments_of_purchase(self, purchase_id: str) -> list[dict[str, Any]]:
        """采购单 → 出运单：经订单这一公共维度（同样是原生外键，不是字符串推断）。"""
        pid = str(purchase_id)
        result: list[dict[str, Any]] = []
        for oid in self.purchase_orders.get(pid, []):
            for sid in self.order_shipments.get(oid, []):
                if sid in self.shipments and self.shipments[sid] not in result:
                    result.append(self.shipments[sid])
        return result

    def stats(self) -> dict[str, int]:
        return {
            "orders": len(self.orders),
            "shipments": len(self.shipments),
            "purchases": len(self.purchases),
            "orders_with_shipment": sum(1 for v in self.order_shipments.values() if v),
            "orders_with_purchase": sum(1 for v in self.order_purchases.values() if v),
            "shipments_with_order": sum(1 for v in self.shipment_orders.values() if v),
            "purchases_with_order": sum(1 for v in self.purchase_orders.values() if v),
        }


def build() -> CrossRef:
    return CrossRef().load()
