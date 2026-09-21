"""从睿贝 ERP 取成本数据（出运单 / 采购单 / 外销订单），带本地缓存与节流。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app.services.erp_client import ErpClient, ErpError
from app.services.order_numbers import expand_combined, purchase_candidates, strip_pi

CACHE_DIR = Path(__file__).resolve().parents[2] / "data" / "cache" / "erp"


def _load(name: str) -> dict[str, Any]:
    path = CACHE_DIR / f"{name}.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def _save(name: str, data: dict[str, Any]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    (CACHE_DIR / f"{name}.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


class ErpDataStore:
    """ERP 数据访问层：缓存优先，缺失才联网；结果写回缓存。"""

    def __init__(self, client: ErpClient | None = None, offline: bool = False) -> None:
        self.client = client or ErpClient()
        self.offline = offline
        self._shipments: dict[str, Any] = _load("shipment_cache")
        self._purchases: dict[str, Any] = _load("purchase_cache")
        self._sales: dict[str, Any] = _load("sale_order_cache")

    # ---------- 出运单 ----------

    def shipment(self, contract: str) -> dict[str, Any] | None:
        """按外销发票号（合同号）取出运单详情。"""
        key = strip_pi(contract)
        if key in self._shipments:
            return self._shipments[key]
        if self.offline:
            return None
        result = self._try_shipment(key)
        self._shipments[key] = result
        _save("shipment_cache", self._shipments)
        return result

    def _try_shipment(self, key: str) -> dict[str, Any] | None:
        for candidate in (f"PI-{key}", key):
            try:
                result = self.client.call_tool(
                    "shipment.find",
                    {
                        "invoiceCode": candidate,
                        "includeProduct": "true",
                        "includeExpense": "true",
                        "includePurchaseExpense": "true",
                    },
                )
            except ErpError as exc:
                return {"success": False, "msg": str(exc)}
            if isinstance(result, dict) and result.get("success"):
                return result
        return {"success": False, "msg": "未查询到此出运单"}

    # ---------- 采购单 ----------

    def purchase(self, code: str) -> dict[str, Any] | None:
        key = strip_pi(code)
        if key in self._purchases:
            return self._purchases[key]
        if self.offline:
            return None
        result = self._try_purchase(key)
        self._purchases[key] = result
        _save("purchase_cache", self._purchases)
        return result

    def _try_purchase(self, key: str) -> dict[str, Any] | None:
        for candidate in purchase_candidates(key):
            try:
                result = self.client.call_tool(
                    "purchase.find",
                    {
                        "purchaseCode": candidate,
                        "includeProduct": "true",
                        "includeExpense": "true",
                    },
                )
            except ErpError as exc:
                return {"success": False, "msg": str(exc)}
            if isinstance(result, dict) and result.get("success"):
                return result
        return {"success": False, "msg": "未查询到此采购单"}

    # ---------- 外销订单 ----------

    def sale_order(self, contract: str) -> dict[str, Any] | None:
        key = strip_pi(contract)
        if key in self._sales:
            return self._sales[key]
        if self.offline:
            return None
        try:
            result = self.client.call_tool(
                "sale_order.find",
                {
                    "orderCode": f"PI-{key}",
                    "includeProduct": "true",
                    "includeExpense": "true",
                },
            )
        except ErpError as exc:
            result = {"success": False, "msg": str(exc)}
        self._sales[key] = result
        _save("sale_order_cache", self._sales)
        return result

    # ---------- 便捷视图 ----------

    @staticmethod
    def shipment_rows(shipment: dict[str, Any] | None) -> list[dict[str, Any]]:
        """取出运单产品行，兼容 value 为 dict 或 JSON 字符串。"""
        if not shipment or not shipment.get("success"):
            return []
        value = shipment.get("value")
        if isinstance(value, str):
            try:
                value = json.loads(value)
            except ValueError:
                return []
        if not isinstance(value, dict):
            return []
        return value.get("productList") or []

    @staticmethod
    def shipment_base(shipment: dict[str, Any] | None) -> dict[str, Any]:
        if not shipment or not shipment.get("success"):
            return {}
        value = shipment.get("value")
        if isinstance(value, str):
            try:
                value = json.loads(value)
            except ValueError:
                return {}
        return (value or {}).get("shipmentBaseInfo") or {}

    def shipment_orders(self, contract: str) -> list[str]:
        """合同号可能是联合出运号，展开成涉及的订单号列表。"""
        return expand_combined(contract)
