"""出运单产品行统一索引。

合并两个来源：
1. 历史项目缓存 `MT_cost_match_final/cache/erp/web_shipment_detail_cache.json`
   （按 shipmentId 存，产品行字段与 MCP 一致）
2. 本项目通过 MCP `shipment.find` 抓取的 `details/shipments/<发票号>.json`

输出 `.cache/erp/index/shipment_lines.jsonl`，每行一条产品行，含
`invoice_code` / `shipment_id` / `order_code` / `purchase_code` / `supplier` /
`amount_rmb` / `customs_name` / `quantity`，供拆单直接消费。
也可按需重算：`python -c "from app.services.shipment_index import build; print(build())"`
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Iterator

from app.services.erp_cache import CACHE_ROOT, ensure_dir, read_jsonl, write_jsonl

LEGACY_CACHE = (
    Path("E:/财务/cost matching/MT_cost_match_final/cache/erp/web_shipment_detail_cache.json")
)
INDEX_PATH = CACHE_ROOT / "index" / "shipment_lines.jsonl"

TOKEN_SPLIT = re.compile(r"[;&,，、;]")
ORDER_RE = re.compile(r"(\d{2}MT-\d{2}[A-Z]\d{3}[A-Za-z0-9\-]*)", re.IGNORECASE)
# 基号 = 年份 + MT + 部门 + 业务员 + 流水；尾部的 Y（佣金单）也是订单身份的一部分，必须保留
BASE_RE = re.compile(r"^(\d{2}MT-\d{2}[A-Z]\d{3}Y?)", re.IGNORECASE)


def strip_pi(value: Any) -> str:
    text = "".join(str(value or "").split()).upper()
    return text[3:] if text.startswith("PI-") else text


def base_code(value: Any) -> str:
    match = BASE_RE.match(strip_pi(value))
    return match.group(1).upper() if match else ""


def canonical(value: Any) -> str:
    """规范化单号：去掉 `PI-` 前缀与所有分隔符（空格/连字符/下划线）。

    睿贝里同一张出运单可能写成 `26MT-10E101-Y-A`、`26MT-10E101-Y A`、
    `26MT-10E101-YA` 等，去掉分隔符后可以统一比较。
    """
    text = strip_pi(value)
    return re.sub(r"[-_\s]", "", text)


def row_orders(row: dict[str, Any]) -> list[str]:
    """该产品行属于哪个外销订单。"""
    out: list[str] = []
    for field in ("销售订单号", "SKU", "采购订单号"):
        for token in TOKEN_SPLIT.split(str(row.get(field) or "")):
            match = ORDER_RE.search(token)
            if match:
                code = match.group(1).upper()
                if code not in out:
                    out.append(code)
    return out


DOUBLE_PREFIX = re.compile(r"^(\d{2}MT-)\1", re.IGNORECASE)


def normalise_purchase_code(code: str) -> str:
    """采购单号归一：ERP 出运行里偶尔把年份前缀写两遍
    （`26MT-26MT-03T094Y-YH`，真实采购单是 `26MT-03T094Y-YH`），不归一就找不到入库单。
    """
    text = str(code or "").strip()
    while True:
        fixed = DOUBLE_PREFIX.sub(r"\1", text)
        if fixed == text:
            return text
        text = fixed


def _normalise(row: dict[str, Any], invoice: str, shipment_id: str) -> dict[str, Any]:
    def text_field(name: str) -> str:
        """字段值归一：把内部空白（含换行）压成单空格。

        MCP 明细的 SKU 常带换行（`1smls tube\\n0.250.065`），而历史缓存里是空格，
        不归一就会让同一产品行在索引里存两份。
        """
        raw = str(row.get(name) or "")
        return re.sub(r"\s+", " ", raw).strip()

    purchase = normalise_purchase_code(str(row.get("采购订单号") or ""))
    product_code = text_field("产品编码")
    from app.services.product_master import category_of

    # 产品类型 = 商品资料「类别名称」去掉部门括号后的粗分类（2026-09-28 用户口径）；
    # 商品资料没抓到/没有类别时留空，交给拆单阶段用报关品名兜底。
    product_type = category_of(product_code) if product_code else ""
    return {
        "invoice_code": invoice,
        "shipment_id": shipment_id,
        "order_codes": row_orders(row),
        "purchase_code": purchase,
        "supplier": str(row.get("供应商名称") or "").strip(),
        "supplier_code": str(row.get("供应商编码") or "").strip(),
        "amount_rmb": str(row.get("出运采购金额(RMB)") or "").strip(),
        "amount_usd": str(row.get("出运金额") or "").strip(),
        "hs_code": str(row.get("海关编码") or "").strip(),
        "unit_price_usd": str(row.get("外销单价") or "").strip(),
        "customs_name": text_field("海关商品（中文）"),
        "product_code": product_code,
        "product_type": product_type,
        "quantity": str(row.get("出运数量") or "").strip(),
        "order_quantity": str(row.get("订单数量") or "").strip(),
        "unit": str(row.get("计量单位") or "").strip(),
        "sku": text_field("SKU"),
    }


_PURCHASE_QUANTITY: dict[str, float] | None = None
_PURCHASE_AMOUNT: dict[str, float] | None = None


def purchase_quantity_map() -> dict[str, float]:
    """采购单号 → 采购数量（拆「一格两个订单」的合并行时要按采购数量分摊）。"""
    global _PURCHASE_QUANTITY
    if _PURCHASE_QUANTITY is None:
        table: dict[str, float] = {}
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            code = str(row.get("purchase_code") or "").strip()
            if not code:
                continue
            quantity = _plain_number(row.get("quantity"))
            if quantity > 0:
                table.setdefault(code.upper(), quantity)
        _PURCHASE_QUANTITY = table
    return _PURCHASE_QUANTITY


def purchase_amount_map() -> dict[str, float]:
    """采购单号 → 采购金额（RMB，不含采购费用）。"""
    global _PURCHASE_AMOUNT
    if _PURCHASE_AMOUNT is None:
        table: dict[str, float] = {}
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            code = str(row.get("purchase_code") or "").strip()
            if not code:
                continue
            amount = _plain_number(row.get("productAmount")) or _plain_number(row.get("amount"))
            if amount > 0:
                table.setdefault(code.upper(), amount)
        _PURCHASE_AMOUNT = table
    return _PURCHASE_AMOUNT


def _supplier_full_name(code: str) -> str:
    if not code:
        return ""
    from app.services.shipment_detail import supplier_name_map

    return supplier_name_map().get(code, "")


def _company_key(name: str) -> str:
    """公司名归一：去掉「（麦金老账号）」这类括号备注，便于判断是否同一家。"""
    return re.sub(r"[（(][^）)]*[）)]", "", str(name or "")).strip()


def split_merged_parts(item: dict[str, Any]) -> list[dict[str, Any]]:
    """把「一格两个订单、两个公司」的合并产品行拆成多条。

    出运产品行偶尔会把多个采购订单塞进一格：`采购订单号 = A,B`、
    `供应商 = 甲公司,乙公司`。这时一行货其实来自两家工厂，必须拆开才能拆单。
    只拆**能确认是两家不同公司**的情况：供应商编码全部能在名录里解析出名称，
    且归一后的公司名不止一个。解析不出、或只是同一家的两个账号
    （如「麦金老账号 / 迈拓新账号」）时保持原样，避免把一家公司拆成两家。
    金额按**采购数量**分摊（采购数量之和等于该行出运数量，是 ERP 自身口径）。
    """
    codes = [c.strip() for c in re.split(r"[,，、;；]", str(item.get("purchase_code") or "")) if c.strip()]
    if len(codes) < 2:
        return [item]
    supplier_codes = [
        c.strip() for c in re.split(r"[,，、;；]", str(item.get("supplier_code") or "")) if c.strip()
    ]
    suppliers = [
        s.strip() for s in re.split(r"[,，、;；]", str(item.get("supplier") or "")) if s.strip()
    ]
    if len(supplier_codes) != len(codes):
        return [item]
    full_names = [_supplier_full_name(code) for code in supplier_codes]
    if not all(full_names):
        return [item]
    if len({_company_key(name) for name in full_names}) < 2:
        return [item]

    quantities = purchase_quantity_map()
    weights = [quantities.get(code.upper(), 0.0) for code in codes]
    if min(weights) <= 0:
        if len(suppliers) != len(codes):
            return [item]
        weights = [1.0] * len(codes)
    total_weight = sum(weights)
    total_amount = _plain_number(item.get("amount_usd"))
    total_quantity = _plain_number(item.get("quantity"))
    # 人民币采购金额不能按数量摊（两家单价不同），要用各采购单自己的金额
    rmb_amounts = purchase_amount_map()
    rmb_parts = [rmb_amounts.get(code.upper(), 0.0) for code in codes]
    total_rmb = _plain_number(item.get("amount_rmb"))
    split_rmb = total_rmb > 0 and min(rmb_parts) > 0
    parts: list[dict[str, Any]] = []
    allocated_amount = 0.0
    allocated_quantity = 0.0
    for position, code in enumerate(codes):
        last = position == len(codes) - 1
        share = weights[position] / total_weight
        amount = (
            total_amount - allocated_amount
            if last
            else round(total_amount * share, 2)
        )
        quantity = (
            total_quantity - allocated_quantity
            if last
            else round(total_quantity * share, 6)
        )
        allocated_amount += amount
        allocated_quantity += quantity
        part = dict(item)
        part["purchase_code"] = code
        part["supplier_code"] = supplier_codes[position]
        part["supplier"] = full_names[position]
        part["amount_usd"] = f"{amount:.5f}"
        part["quantity"] = f"{quantity:.6f}"
        if split_rmb:
            part["amount_rmb"] = f"{rmb_parts[position]:.5f}"
        part["merged_from"] = item.get("purchase_code") or ""
        part["merged_parts"] = len(codes)
        parts.append(part)
    return parts


def _iter_legacy() -> Iterator[tuple[str, str, dict[str, Any]]]:
    if not LEGACY_CACHE.exists():
        return
    payload = json.loads(LEGACY_CACHE.read_text(encoding="utf-8"))
    for key, item in (payload.get("details") or {}).items():
        if not isinstance(item, dict) or not item.get("success"):
            continue
        value = item.get("value") or {}
        base = value.get("shipmentBaseInfo") or {}
        invoice = str(base.get("外销发票号") or "")
        sid = str(base.get("shipmentId") or key)
        for row in value.get("productList") or []:
            yield invoice, sid, row


def _iter_mcp() -> Iterator[tuple[str, str, dict[str, Any]]]:
    detail_dir = CACHE_ROOT / "details" / "shipments"
    if not detail_dir.exists():
        return
    for path in sorted(detail_dir.glob("*.json")):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        base = payload.get("baseInfo") or {}
        invoice = str(payload.get("invoiceCode") or base.get("外销发票号") or "")
        sid = str(payload.get("shipmentId") or base.get("shipmentId") or "")
        for row in payload.get("productList") or []:
            yield invoice, sid, row


def build() -> dict[str, int]:
    """重建统一索引，返回统计。"""
    invoice_to_id: dict[str, str] = {}
    for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
        code = str(row.get("invoiceCode") or "").strip()
        sid = str(row.get("shipmentId") or "").strip()
        if code and sid:
            invoice_to_id.setdefault(code, sid)

    seen: set[tuple] = set()
    lines: list[dict[str, Any]] = []
    for invoice, sid, row in list(_iter_legacy()) + list(_iter_mcp()):
        invoice = invoice or ""
        sid = sid or invoice_to_id.get(invoice, "")
        item = _normalise(row, invoice, sid)
        # 去重标记：数量/金额按数值归一（`3.4` 与 `3.400000` 视为同一条）
        marker = (
            item["invoice_code"].strip(),
            item["purchase_code"].strip(),
            item["sku"].strip(),
            _num(item["amount_rmb"]),
            _num(item["quantity"]),
            item["customs_name"].strip(),
            tuple(item["order_codes"]),
        )
        if marker in seen:
            continue
        seen.add(marker)
        # 一格两个订单、两个公司的合并行：拆成多条，各自带自己的采购单号与供应商
        lines.extend(split_merged_parts(item))

    ensure_dir(INDEX_PATH.parent)
    write_jsonl(INDEX_PATH, lines)
    stats = {
        "lines": len(lines),
        "invoices": len({x["invoice_code"] for x in lines if x["invoice_code"]}),
        "with_supplier": sum(1 for x in lines if x["supplier"]),
        "with_amount": sum(1 for x in lines if x["amount_rmb"]),
    }
    return stats


def _num(text: Any) -> str:
    """把数量/金额文本归一成数值字符串（去尾零）。"""
    raw = str(text or "").replace(",", "").strip()
    if not raw:
        return ""
    try:
        value = float(raw)
    except ValueError:
        return raw
    return f"{value:.6f}".rstrip("0").rstrip(".")


def _plain_number(text: Any, default: float = 0.0) -> float:
    """`1,234.50` 这类带千分位的数值文本 → float。"""
    raw = str(text or "").replace(",", "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def load_lines() -> list[dict[str, Any]]:
    return read_jsonl(INDEX_PATH)


def by_order_index() -> dict[str, list[dict[str, Any]]]:
    """订单号 -> 产品行。"""
    return _group_by(lambda line: [strip_pi(c) for c in (line.get("order_codes") or [])])


def by_invoice_index() -> dict[str, list[dict[str, Any]]]:
    """出运发票号 -> 产品行（报关合同号通常就是它）。"""
    return _group_by(lambda line: [strip_pi(line.get("invoice_code"))])


def _group_by(key_fn) -> dict[str, list[dict[str, Any]]]:
    index: dict[str, list[dict[str, Any]]] = {}
    for line in load_lines():
        for code in key_fn(line):
            if code:
                index.setdefault(code, []).append(line)
    return index


def by_base_index() -> dict[str, list[dict[str, Any]]]:
    """基号 -> 产品行（忽略 `PI-` 前缀与 `-A/-B/ADD/Y` 尾缀）。"""
    index: dict[str, list[dict[str, Any]]] = {}
    for line in load_lines():
        keys = set()
        for code in line.get("order_codes") or []:
            keys.add(base_code(code))
        keys.add(base_code(line.get("invoice_code")))
        keys.add(base_code(line.get("purchase_code")))
        for key in keys:
            if key:
                index.setdefault(key, []).append(line)
    return index


def by_purchase_prefix_index(min_len: int = 12) -> dict[str, list[dict[str, Any]]]:
    """采购单号前缀 → 产品行。用于「订单带批次字母、采购单号写法不同」的情况。

    例如订单 `26MT-03P201Y-E` 的采购单号可能是 `26MT-03P201Y-E-HD`，
    通过前缀 `26MT-03P201Y-E` 就能命中。
    """
    index: dict[str, list[dict[str, Any]]] = {}
    for line in load_lines():
        code = strip_pi(line.get("purchase_code"))
        if len(code) < min_len:
            continue
        for length in range(min_len, len(code) + 1):
            index.setdefault(code[:length], []).append(line)
    return index
