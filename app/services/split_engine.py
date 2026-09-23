"""拆单：把一条纯净的报关商品行按「外销订单 → 采购订单（供应商）→ 产品类型」拆开。

数据源（全部是睿贝原生数据）：
- 出运单产品行：MCP `shipment.find` 的 `productList`
  （含「销售订单号 / 采购订单号 / 供应商编码 / 出运采购金额(RMB) / 海关商品（中文）」）
- 采购单：`purchases.jsonl`，用「采购单号」取供应商名与采购金额

拆单口径：
1. 报关行的 `合同号_1` 是外销订单号（可能是 `&` 联合号），逐个订单找它的出运单；
2. 在该出运单的产品行里筛出属于这些订单的行；
3. 按「采购订单号」归组——同一采购单 = 同一供应商；
4. 报关行的重量按各采购组的外销金额占比分摊（末组吸收分位尾差）；
5. 找不到对应出运行的报关行进异常表，并保留原因。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

from app.services.erp_cache import CACHE_ROOT, read_jsonl
from app.services.shipment_index import (
    base_code,
    by_base_index,
    by_invoice_index,
    by_order_index,
    by_purchase_prefix_index,
    strip_pi,
)


def _dec(value: Any) -> Decimal:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return Decimal(0)
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def _round2(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


SUPPLIER_SUFFIX = re.compile(r"[-－]?\s*(供应链|总部|本部|一部|二部)\s*$")


def normalize_supplier(name: str) -> str:
    """去掉 `-供应链`、`总部` 这类内部后缀，便于与飞书比对。"""
    text = str(name or "").strip()
    cleaned = SUPPLIER_SUFFIX.sub("", text).strip()
    return cleaned or text


TOKEN_SPLIT = re.compile(r"[&;,，、]")

# 产品类型归一：ERP「海关商品（中文）」与飞书「产品类型」写法不完全一致
TYPE_ALIASES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("法兰", ("法兰", "flange")),
    ("管件", ("管件", "fitting", "elbow", "cap", "tee", "reducer", "弯头", "三通", "管帽")),
    ("焊管", ("焊管", "welded")),
    ("无缝管", ("无缝管", "seamless", "smls", "钢管")),
    ("板棒", ("板", "棒", "plate", "bar")),
    ("焊材", ("焊丝", "焊材", "welding wire")),
    ("三角丝", ("三角丝", "线材", "wire")),
    ("盘管", ("盘管", "线圈", "coil")),
)


def _ptype_match(customs_name: Any, want_type: str) -> bool:
    """判断出运行的产品名是否属于目标产品类型。"""
    text = str(customs_name or "").strip()
    want = str(want_type or "").strip()
    if not text or not want:
        return False
    if want in text:
        return True
    low = text.lower()
    # 找出「目标类型」所属的同义组，再用该组的关键词去匹配出运产品名
    for kind, aliases in TYPE_ALIASES:
        group = (kind, *aliases)
        if not any(word in want or want in word for word in group):
            continue
        for alias in group:
            if re.search(r"[\u4e00-\u9fff]", alias):
                if alias in text:
                    return True
            elif re.search(r"(?<![a-z])" + re.escape(alias) + r"s?(?![a-z])", low):
                return True
    return False


def contract_orders(contract: str) -> list[str]:
    """`26MT-01S180&190&241` -> [26MT-01S180, 26MT-01S190, 26MT-01S241]。"""
    raw = (contract or "").strip()
    if not raw:
        return []
    if not TOKEN_SPLIT.search(raw):
        return [raw]
    parts = [p.strip() for p in TOKEN_SPLIT.split(raw) if p.strip()]
    base = parts[0]
    head = re.match(r"^(\d{2}MT-\d{2}[A-Z])", base, re.IGNORECASE)
    prefix = head.group(1) if head else ""
    out: list[str] = []
    for part in parts:
        if re.match(r"^\d{2}MT-", part, re.IGNORECASE):
            out.append(part)
        elif re.fullmatch(r"\d{1,3}", part) and prefix:
            out.append(f"{prefix}{int(part):03d}")
        elif re.match(r"^\d{2}[A-Z]", part, re.IGNORECASE) and prefix:
            out.append(f"{prefix}{part}")
        else:
            out.append(f"{base}-{part}" if part.upper().startswith("ADD") else f"{base}{part}")
    return list(dict.fromkeys(out))


@dataclass
class DeclaredLine:
    """一条待拆的报关商品行。"""

    source_file: str
    contract: str
    declaration_no: str
    product_name: str
    hs_code: str
    weight: Decimal
    currency: str
    product_type: str = ""
    amount: Decimal = Decimal(0)


@dataclass
class SplitRow:
    """拆单结果一行。"""

    contract: str
    declaration_no: str
    product_name: str
    hs_code: str
    product_type: str
    supplier: str
    purchase_code: str
    shipment_id: str
    ship_date: str
    weight: Decimal
    purchase_amount: Decimal
    currency: str
    source_file: str


@dataclass
class SplitOutcome:
    rows: list[SplitRow] = field(default_factory=list)
    errors: list[dict[str, Any]] = field(default_factory=list)


class SplitEngine:
    """基于出运单产品行的拆单器。"""

    def __init__(self, product_type_map: dict[tuple[str, str], str] | None = None) -> None:
        """`product_type_map`：(报关单号, 报关品名) -> 产品类型。

        同一张出运单可能混装多个产品类型（焊管/管件/法兰…），
        只有该报关行对应产品类型的那部分才能分摊给它；
        没有产品类型信息时退化为按整张出运单分摊（兼容旧口径）。
        """
        self._product_type_map = product_type_map or {}
        self._purchase_supplier: dict[str, str] = {}
        self._purchase_amount: dict[str, Decimal] = {}
        self._by_order = by_order_index()
        self._by_base = by_base_index()
        self._by_prefix = by_purchase_prefix_index()
        self._by_invoice = by_invoice_index()
        self._load_purchases()

    def _load_purchases(self) -> None:
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            code = str(row.get("purchase_code") or "").strip()
            if not code:
                continue
            name = str(row.get("supplierName") or "").strip()
            if name:
                self._purchase_supplier.setdefault(code, name)
            self._purchase_amount.setdefault(code, _dec(row.get("amount")))

    # ---------- 出运产品行 ----------

    def _lines_for(self, order_codes: list[str]) -> list[dict[str, Any]]:
        """找出运产品行。报关单上的合同号通常就是**出运发票号**，所以优先按发票号匹配。

        匹配顺序：
        1. 出运发票号（精确，含 `PI-`/空格归一）
        2. 出运发票号（加/去 `-A`、`-B` 等尾缀）
        3. 订单号（产品行里的「销售订单号」）
        4. 基号唯一时用基号兜底
        5. 采购单号前缀兜底
        """
        wanted = [strip_pi(c) for c in order_codes if c]
        if not wanted:
            return []
        out: list[dict[str, Any]] = []
        seen: set[int] = set()

        def add(rows: list[dict[str, Any]]) -> None:
            for row in rows:
                if id(row) not in seen:
                    seen.add(id(row))
                    out.append(row)

        # 1/2. 出运发票号（先精确，再加/去批次尾缀）
        for code in wanted:
            add(self._by_invoice.get(code, []))
        if not out:
            for code in wanted:
                for variant in (f"{code}-A", f"{code}-B", code[:-2] if code.endswith(("-A", "-B")) else ""):
                    if variant:
                        add(self._by_invoice.get(variant, []))

        # 3. 订单号
        if not out:
            for code in wanted:
                add(self._by_order.get(code, []))

        # 4. 基号兜底（唯一出运单时才用）
        if not out:
            candidates: list[dict[str, Any]] = []
            for code in wanted:
                base = base_code(code)
                if not base:
                    continue
                candidates.extend(self._by_base.get(base, []))
            invoices = {str(c.get("invoice_code") or "") for c in candidates}
            if len(invoices) == 1:
                add(candidates)

        # 5. 采购单号前缀兜底
        if not out:
            for code in wanted:
                add(self._by_prefix.get(code, []))
        return out

    def split_line(self, line: DeclaredLine) -> SplitOutcome:
        outcome = SplitOutcome()
        orders = contract_orders(line.contract)
        lines = self._lines_for(orders)
        # 按产品类型筛选：一张出运单常混装多种产品，只取与该报关行同类型的产品行
        if not lines:
            outcome.errors.append(
                {
                    "来源文件": line.source_file,
                    "合同号_1": line.contract,
                    "报关单号": line.declaration_no,
                    "报关品名": line.product_name,
                    "海关编码": line.hs_code,
                    "报关重量": str(line.weight),
                    "异常原因": "未找到对应出运产品行",
                }
            )
            return outcome

        want_type = self._product_type_map.get((line.declaration_no, line.product_name))
        if want_type:
            typed = [r for r in lines if _ptype_match(r.get("customs_name"), want_type)]
            if typed:
                lines = typed
            # 类型没命中时不报错：飞书的产品类型是粗分类，措辞与 ERP 的
            # 「海关商品（中文）」并不一一对应，硬拦会把能拆的也判失败。

        # 按采购订单号归组
        groups: dict[str, dict[str, Any]] = {}
        for row in lines:
            purchase = str(row.get("purchase_code") or "").strip() or "(无采购单号)"
            entry = groups.setdefault(
                purchase,
                {
                    "purchase": purchase,
                    "amount": Decimal(0),
                    "purchase_amount": self._purchase_amount.get(purchase, Decimal(0)),
                    "supplier": self._purchase_supplier.get(purchase, ""),
                    "product_types": [],
                    "shipments": [],
                    "dates": [],
                },
            )
            entry["amount"] += _dec(row.get("amount_usd"))
            kind = str(row.get("customs_name") or "").strip()
            if kind and kind not in entry["product_types"]:
                entry["product_types"].append(kind)
            if not entry["supplier"]:
                entry["supplier"] = str(row.get("supplier") or "").strip()
            if not entry["purchase_amount"]:
                entry["purchase_amount"] = _dec(row.get("amount_rmb"))
            invoice = str(row.get("invoice_code") or "")
            if invoice and invoice not in entry["shipments"]:
                entry["shipments"].append(invoice)

        # 采购单里没有的供应商，直接用出运行自带的

        keys = list(groups.keys())
        total = sum((groups[k]["amount"] for k in keys), Decimal(0))
        if total > 0:
            weights: list[Decimal] = []
            left = line.weight
            for key in keys[:-1]:
                part = _round2(line.weight * groups[key]["amount"] / total)
                weights.append(part)
                left -= part
            weights.append(left)
        else:
            weights = [line.weight / len(keys) for _ in keys]

        for key, weight in zip(keys, weights):
            entry = groups[key]
            outcome.rows.append(
                SplitRow(
                    contract=line.contract,
                    declaration_no=line.declaration_no,
                    product_name=line.product_name,
                    hs_code=line.hs_code,
                    product_type="、" .join(entry["product_types"]),
                    supplier=normalize_supplier(entry["supplier"]),
                    purchase_code=entry["purchase"],
                    shipment_id="、".join(entry["shipments"]),
                    ship_date="",
                    weight=weight,
                    purchase_amount=entry["purchase_amount"],
                    currency=line.currency,
                    source_file=line.source_file,
                )
            )
        return outcome

    def split_many(self, lines: list[DeclaredLine]) -> SplitOutcome:
        total = SplitOutcome()
        for line in lines:
            outcome = self.split_line(line)
            total.rows.extend(outcome.rows)
            total.errors.extend(outcome.errors)
        return total
