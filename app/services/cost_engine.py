"""拆单与成本分摊引擎。

口径：以「合同号(外销订单) + 报关单号」为一组，把 ERP 出运单的采购金额
先按产品类型归集（同一产品类型下再按供应商/采购订单细分），
再按报关重量比例把每个「产品类型 × 供应商」的金额分摊到对应的报关商品行。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

from app.services.erp_data import ErpDataStore
from app.services.order_numbers import expand_combined, is_excluded, strip_pi
from app.services.reference_data import ReferenceData

# 产品类型兜底判定：ERP「海关商品（中文）」里带英文或未收录时的关键字规则
TYPE_KEYWORDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("法兰", ("flange", "法兰")),
    ("管件", ("fitting", "elbow", "elbow", "tee", "reducer", "stub end", "coupling", "管件")),
    ("无缝管", ("seamless", "smls", "无缝")),
    ("焊管", ("welded pipe", "welded tube", "焊管")),
    ("板棒", ("plate", "round bar", "flat bar", "板", "棒")),
    ("焊材", ("welding wire", "焊丝", "焊材")),
    ("三角丝", ("triangle wire", "三角丝")),
    ("盘管", ("coiled tube", "coil tube", "盘管")),
    ("螺栓", ("bolt", "螺栓")),
    ("螺母", ("nut", "螺母")),
    ("垫圈", ("washer", "垫圈")),
)


def _money(value: Any) -> Decimal:
    try:
        return Decimal(str(value or 0))
    except Exception:  # noqa: BLE001
        return Decimal(0)


def supplier_key(name: str, short: str = "") -> str:
    """把供应商全称/简称归一到同一比较键。"""
    text = (name or short or "").strip()
    if not text:
        return ""
    text = re.sub(r"(有限责任公司|有限公司|公司|厂|供应链)$", "", text)
    return text


def to_decimal(value: Any) -> Decimal:
    """宽松解析数值（兼容千分位、空值、列表）。"""
    if isinstance(value, list):
        value = value[0] if value else None
    if isinstance(value, dict):
        value = value.get("text") or value.get("value")
    text = str(value or "").strip().replace(",", "")
    if not text:
        return Decimal(0)
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def _round2(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


@dataclass
class DeclaredRow:
    """一条报关商品行（来自飞书多维表）。"""

    record_id: str
    contract: str
    declaration_no: str = ""
    product_name: str = ""
    weight: Decimal = Decimal(0)
    amount: Decimal = Decimal(0)
    product_type: str = ""
    supplier: str = ""


@dataclass
class AllocationResult:
    """一条报关行的分摊结果。"""

    record_id: str
    purchase_amount: Decimal = Decimal(0)
    supplier: str = ""
    product_type: str = ""
    purchase_orders: list[str] = field(default_factory=list)
    status: str = "自动匹配"
    message: str = ""


class CostEngine:
    """拆单 + 成本分摊。"""

    def __init__(
        self,
        store: ErpDataStore | None = None,
        reference: ReferenceData | None = None,
    ) -> None:
        self.store = store or ErpDataStore()
        self.reference = reference or ReferenceData()

    # ---------- ERP 产品行 ----------

    def _shipment_products(self, contract: str) -> tuple[list[dict[str, Any]], list[str]]:
        """取出运单产品行。合同号可能是联合出运号，逐个订单号尝试。"""
        orders = expand_combined(contract) or [strip_pi(contract)]
        products: list[dict[str, Any]] = []
        used: list[str] = []
        for order in orders:
            rows = self.store.shipment_rows(self.store.shipment(order))
            if rows:
                products.extend(rows)
                used.append(order)
        return products, used

    def product_type(self, name: str) -> str:
        """解析产品类型：先查字典表，再按中英文关键字兜底。"""
        if isinstance(name, (list, tuple)):
            name = " ".join(str(x) for x in name)
        text = str(name or "").strip()
        if not text:
            return ""
        kind = self.reference.product_type(text)
        if kind:
            return kind
        low = text.lower()
        for kind, keywords in TYPE_KEYWORDS:
            for keyword in keywords:
                if re.search(r"[a-z]", keyword):
                    if re.search(r"(?<![a-z])" + re.escape(keyword) + r"s?(?![a-z])", low):
                        return kind
                elif keyword in text:
                    return kind
        if "管" in text:
            return "钢管"
        if "丝" in text or "线" in text:
            return "线材"
        if "manifold" in low or "盘管" in text:
            return "盘管"
        if text:
            return "其他"
        return ""

    def _purchase_groups(self, products: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """把 ERP 出运产品行按采购订单号归集，得到「采购单 → 金额/供应商/产品类型/重量」。"""
        groups: dict[str, dict[str, Any]] = {}
        for row in products:
            purchase = str(row.get("采购订单号") or "").strip()
            if not purchase:
                continue
            entry = groups.setdefault(
                purchase,
                {
                    "purchase": purchase,
                    "amount": Decimal(0),
                    "weight": Decimal(0),
                    "supplier_code": "",
                    "supplier": "",
                    "types": [],
                },
            )
            entry["amount"] += _money(row.get("出运采购金额(RMB)"))
            entry["weight"] += to_decimal(row.get("每个净重")) * to_decimal(row.get("出运数量"))
            code = str(row.get("供应商编码") or "").strip()
            if code and not entry["supplier_code"]:
                entry["supplier_code"] = code
                entry["supplier"] = self.reference.supplier_name(code)
            raw_name = str(row.get("海关商品（中文）") or row.get("商品名称及规格型号") or "")
            kind = self.product_type(raw_name)
            if kind and kind not in entry["types"]:
                entry["types"].append(kind)
        return list(groups.values())

    # ---------- 主流程 ----------

    def allocate(self, rows: list[DeclaredRow]) -> list[AllocationResult]:
        """按合同号分组，把 ERP 采购金额按产品类型/供应商分摊到该合同下的报关行。

        同一合同可能对应多张报关单，先做合同级归集，再按报关重量分摊到各行。
        """
        results: dict[str, AllocationResult] = {}
        groups: dict[str, list[DeclaredRow]] = {}
        skipped: dict[str, str] = {}

        for row in rows:
            reason = is_excluded(row.contract)
            if reason:
                skipped[row.record_id] = reason
                continue
            if not row.declaration_no:
                skipped[row.record_id] = "内销/预录订单（无报关单号）"
                continue
            groups.setdefault(strip_pi(row.contract), []).append(row)

        for contract, members in groups.items():
            products, used_orders = self._shipment_products(contract)
            if not products:
                for row in members:
                    results[row.record_id] = AllocationResult(
                        record_id=row.record_id,
                        product_type=row.product_type or self.reference.product_type(row.product_name),
                        status="失败",
                        message="ERP 未找到对应出运单/产品行",
                    )
                continue

            purchase_groups = self._purchase_groups(products)
            assigned: set[str] = set()
            for entry in purchase_groups:
                bucket = self._bucket_for(entry, members)
                self._distribute(bucket, entry["amount"], entry["purchase"], results, assigned)

            # 未分到金额的报关行（ERP 无对应采购金额）
            for row in members:
                if row.record_id in assigned or row.record_id in results:
                    continue
                results[row.record_id] = AllocationResult(
                    record_id=row.record_id,
                    product_type=row.product_type or self.reference.product_type(row.product_name),
                    status="失败",
                    message="该报关行没有对应的 ERP 采购金额",
                    purchase_orders=used_orders,
                )

        for row in rows:
            if row.record_id in results:
                continue
            results[row.record_id] = AllocationResult(
                record_id=row.record_id, status="跳过", message=skipped.get(row.record_id, "未处理")
            )
        return [results[row.record_id] for row in rows]

    def _bucket_for(self, entry: dict[str, Any], members: list[DeclaredRow]) -> list[DeclaredRow]:
        """为一份采购金额挑选对应的报关行：先按产品类型，再按供应商。

        产品类型来自 ERP 与报关记录两端的解析结果，命中即收窄；供应商用归一化后的
        全称/简称比较。两者都命中不了时退回全部报关行，保证金额不丢失。
        """
        types = [t for t in entry.get("types") or [] if t]
        supplier = (entry.get("supplier") or "").strip()
        target_key = supplier_key(supplier)

        candidates = members
        if types:
            typed = [
                row
                for row in members
                if (row.product_type or self.product_type(row.product_name)) in types
            ]
            if typed:
                candidates = typed
        if target_key:
            by_supplier = [
                row
                for row in candidates
                if supplier_key(row.supplier) and (
                    supplier_key(row.supplier) == target_key
                    or supplier_key(row.supplier) in target_key
                    or target_key in supplier_key(row.supplier)
                )
            ]
            if by_supplier:
                candidates = by_supplier
        return candidates or members

    @staticmethod
    def _distribute(
        bucket: list[DeclaredRow],
        total: Decimal,
        purchase: str,
        results: dict[str, AllocationResult],
        assigned: set[str],
    ) -> None:
        """把一份采购金额按报关重量分摊到一组报关行，末行吸收分位尾差。"""
        weights = [row.weight if row.weight > 0 else Decimal(0) for row in bucket]
        total_weight = sum(weights, Decimal(0))
        left = total
        for index, row in enumerate(bucket):
            if index < len(bucket) - 1 and total_weight > 0:
                part = _round2(total * weights[index] / total_weight)
            elif index < len(bucket) - 1:
                part = _round2(total / len(bucket))
            else:
                part = left
            left -= part
            current = results.get(row.record_id)
            if current and current.status == "自动匹配":
                current.purchase_amount += part
                current.purchase_orders.append(purchase)
                continue
            results[row.record_id] = AllocationResult(
                record_id=row.record_id,
                purchase_amount=part,
                supplier=row.supplier,
                product_type=row.product_type or "",
                purchase_orders=[purchase],
            )
        assigned.update(row.record_id for row in bucket)
