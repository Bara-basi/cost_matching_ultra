"""出运单产品行明细（来自 MCP shipment.find 的 productList）。

每行关键字段：
- 销售订单号 / SKU            -> 该行属于哪个外销订单
- 采购订单号                   -> 该行对应哪个采购单（拆单的分组依据）
- 供应商编码                   -> 配合采购单得到供应商名称
- 出运采购金额(RMB)            -> 采购成本（成本匹配用）
- 海关商品（中文）             -> 产品类型/品名的判定依据
- 出运数量 / 订单数量          -> 半包含订单的进度判断
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT
from app.services.shipment_index import base_code, canonical

SHIPMENT_DETAIL_DIR = CACHE_ROOT / "details" / "shipments"

TOKEN_SPLIT = re.compile(r"[;&,，、]")
ORDER_RE = re.compile(r"(\d{2}MT-\d{2}[A-Z]\d{3}[A-Za-z0-9\-]*)", re.IGNORECASE)
# 订单核心（含佣金 Y 与 -ADDn 尾缀）；核心之外的字母就是报关单点明的**批次**
CORE_RE = re.compile(r"^(\d{2}MT-\d{2}[A-Z]\d{3}Y?(?:-?ADD\d*)?)", re.IGNORECASE)


def _safe(text: str) -> str:
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in str(text or "unknown"))


def detail_path(invoice_code: str) -> Path:
    return SHIPMENT_DETAIL_DIR / f"{_safe(invoice_code)}.json"


def load_shipment_detail(invoice_code: str) -> dict[str, Any] | None:
    path = detail_path(invoice_code)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None


def strip_pi(value: Any) -> str:
    text = "".join(str(value or "").split()).upper()
    return text[3:] if text.startswith("PI-") else text


def row_orders(row: dict[str, Any]) -> list[str]:
    """该产品行属于哪个外销订单：优先「销售订单号」，其次从 SKU/采购订单号里取。"""
    out: list[str] = []
    for field in ("销售订单号", "SKU", "采购订单号"):
        raw = str(row.get(field) or "")
        for token in TOKEN_SPLIT.split(raw):
            token = token.strip()
            if not token:
                continue
            match = ORDER_RE.search(token)
            if match:
                value = match.group(1).upper()
                if value not in out:
                    out.append(value)
    return out


def rows_for_order(invoice_code: str, order_code: str) -> list[dict[str, Any]]:
    """取某出运单里属于指定订单的产品行。"""
    detail = load_shipment_detail(invoice_code)
    if not detail:
        return []
    target = strip_pi(order_code)
    return [
        row
        for row in detail.get("productList") or []
        if target in {strip_pi(x) for x in row_orders(row)}
    ]


def names_batch(contract: Any) -> bool:
    """合同号里是否点明了批次 / 尾缀（`26MT-03R036F`、`25MT-03P495Y-ADD1-A`）。

    核心（年份+MT+部门+业务员+流水[Y][-ADDn]）之外的字母就是批次记号。
    `26MT-03R036F` → True；`26MT-03R302`（没写批次）→ False。
    """
    from app.services.contract_shipments import literal_parts

    for part in literal_parts(contract):
        match = CORE_RE.match(strip_pi(part))
        if match and canonical(part) != canonical(match.group(1)):
            return True
    return False


def iter_shipment_details() -> list[dict[str, Any]]:
    """遍历已缓存的全部出运单明细。"""
    if not SHIPMENT_DETAIL_DIR.exists():
        return []
    out: list[dict[str, Any]] = []
    for path in sorted(SHIPMENT_DETAIL_DIR.glob("*.json")):
        try:
            out.append(json.loads(path.read_text(encoding="utf-8")))
        except ValueError:
            continue
    return out


def supplier_name_map() -> dict[str, str]:
    """供应商编码 -> 名称：用采购单（原生存了供应商名）与出运行编码交叉建立。"""
    mapping: dict[str, str] = {}
    # 1) 采购明细行里有供应商编码吗？没有，则用「采购单号 -> 供应商」+ 出运行的采购单号
    purchase_supplier: dict[str, str] = {}
    for row in _read_purchase_cache():
        code = str(row.get("purchase_code") or "").strip()
        name = str(row.get("supplierName") or "").strip()
        if code and name:
            purchase_supplier[code] = name
    for detail in iter_shipment_details():
        for row in detail.get("productList") or []:
            encoded = str(row.get("供应商编码") or "").strip()
            purchase = str(row.get("采购订单号") or "").strip()
            name = purchase_supplier.get(purchase)
            if encoded and name:
                mapping.setdefault(encoded, name)
    return mapping


def _read_purchase_cache() -> list[dict[str, Any]]:
    from app.services.erp_cache import read_jsonl

    return read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")


def shipment_invoices_for_contract(contract: str, *, use_map: bool = True) -> list[str]:
    """通过出运单列表（原生 `orderCode`/`purchaseCode`）找出该合同对应的出运发票号。

    这是「三单互查」的正向用法：合同号 → 出运单，不依赖字符串猜采购单号。

    合同号是联合写法（`26MT-06H124&08C122`）时，先按订单核心覆盖定位
    （一张报关单可能对应多张出运单）；覆盖不唯一时退回下面的精确/基号匹配。

    `use_map=False` 时跳过「拆分结果 → 原值」映射表，只用覆盖/基号逻辑，
    供调用方把两种解析结果都当作候选池比较。
    """
    from app.services.erp_cache import read_jsonl
    from app.services.contract_shipments import (
        cover_shipments,
        resolve_by_split_map,
        resolve_each_part,
    )

    if use_map:
        resolved = resolve_by_split_map(contract)
        if resolved:
            return resolved
    # 「报关单把出运单拆开写」：逐段各自定位自己的出运单（use_map=False 时也适用，
    # 这比「最少出运单覆盖」更精确——覆盖法只按订单核心匹配，会选到别的批次的合并出运单）
    per_part = resolve_each_part(contract)
    if per_part:
        return per_part
    chosen, ambiguous = cover_shipments(contract)
    if chosen and not ambiguous:
        return [row.invoice for row in chosen if row.invoice]

    target = canonical(contract)
    base = base_code(contract)
    exact: list[str] = []
    fuzzy: list[str] = []
    for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
        invoice = str(row.get("invoiceCode") or "").strip()
        if not invoice:
            continue
        codes = [row.get("orderCode"), row.get("invoiceCode")]
        codes += str(row.get("purchaseCode") or "").replace("&", ",").split(",")
        keys = {canonical(c) for c in codes if c}
        if target in keys:
            if invoice not in exact:
                exact.append(invoice)
        elif base and base in {base_code(c) for c in codes if c}:
            if invoice not in fuzzy:
                fuzzy.append(invoice)
    # 精确命中优先；没有精确命中才用基号兜底。
    # 但合同号里**已经点明批次**（`26MT-03R036F`）时不能兜底：同一订单的其它批次是另一批货，
    # 拿它们顶替等于把成本算成别的批次的（实测 26MT-03R036F 被 A–E 五批共 279 行顶替，
    # 56 美元的报关行被凑出 325.27 元采购金额，并凭空多出 3 条「睿贝出运明细补充」行）。
    # 只有合同号根本没写批次（`26MT-03R302`、`26MT-08C020`）时，取该订单的全部批次才是合理的。
    if names_batch(contract):
        return exact
    return exact or fuzzy


def ensure_shipment_detail(invoice_code: str, *, online: bool = True) -> dict[str, Any] | None:
    """确保出运单明细在本地；缺失时按需从睿贝 MCP 抓取（仅当 online）。"""
    cached = load_shipment_detail(invoice_code)
    if cached:
        return cached
    if not online:
        return None
    try:
        from scripts.mcp_erp import McpClient  # 延迟导入，避免离线场景依赖
        from app.services.feishu_client import get_config
    except Exception:  # noqa: BLE001
        return None
    token = get_config("ERP_API_KEY")
    client = McpClient(token)
    client.initialize()
    for _attempt in range(3):
        text = client.call_text(
            "shipment.find",
            {
                "invoiceCode": invoice_code,
                "includeProduct": "true",
                "includeExpense": "true",
                "includePurchaseExpense": "true",
            },
        )
        if not text.strip().startswith("{"):
            continue
        payload = json.loads(text)
        value = payload.get("value") or {}
        if not value:
            return None
        base = value.get("shipmentBaseInfo") or {}
        record = {
            "invoiceCode": invoice_code,
            "shipmentId": base.get("shipmentId"),
            "baseInfo": base,
            "productList": value.get("productList") or [],
            "expenseList": value.get("expenseList") or [],
            "purchaseExpenseList": value.get("purchaseExpenseList") or [],
        }
        path = detail_path(invoice_code)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(record, ensure_ascii=False, indent=1), encoding="utf-8")
        return record
    return None
