"""把一轮拆单结果整理成「财务看得懂」的逐条记录。

每行只给**一个异常类型**和一个**异常明细**（短句、不含技术术语），
并带上睿贝三类单据入口与飞书记录入口，方便财务点开核对。

金额口径（2026-09-28 用户确认）：
- 采购金额 = 入库单实发金额按各记录出运采购金额(RMB)占比摊到本行（同采购单内分摊）；
- 报关金额 = 拆单结果里的「报关金额」（已含客户费用分摊）；
- 只有「金额算漏」才算异常；同一采购组内分摊位置不同不算异常。
"""
from __future__ import annotations

import collections
import json
import math
import re
from decimal import Decimal
from functools import lru_cache
from pathlib import Path

from openpyxl import load_workbook

from app.services.cost_match import (
    dec,
    shipment_money,
)
from app.services.erp_cache import CACHE_ROOT
from app.services.shipment_index import normalise_purchase_code, strip_pi
from app.services.supplier_names import SupplierNames

PROJECT_ROOT = Path(__file__).resolve().parents[2]
# 金额比对容差（元/本币）：一张报关单的金额基本不可能低于 50，±5 内的差异视为分位、
# 汇率舍入等正常误差（2026-09-29 用户口径：由 ±1 放宽到 ±5）。
EQUAL = Decimal("5")

# 飞书那套粗分类（与拆单保持一致）
KNOWN_CATEGORIES = {
    "法兰", "管件", "无缝管", "镍基无缝管", "焊管", "焊材", "板棒", "盘管", "三角丝", "其他",
}


def read_split_rows(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    if not rows:
        return []
    header = [str(cell) for cell in rows[0]]
    return [dict(zip(header, row)) for row in rows[1:]]


def purchase_key(code: str) -> str:
    """合同号（应收表格）= 采购订单去掉工厂简称（`26MT-03T094Y-YH` → `26MT-03T094Y`）。"""
    text = str(code or "").strip().lstrip("PI-")
    if not text:
        return ""
    head, _sep, tail = text.partition("-")
    parts = text.split("-")
    while parts:
        token = parts[-1].strip()
        if token and token.isascii() and token.isalpha() and len(token) <= 5:
            parts.pop()
            continue
        if token and token.isdigit() and len(token) <= 3:
            parts.pop()
            continue
        break
    return "-".join(parts)


def allocate_costs(rows: list[dict]) -> list[dict]:
    """界面直接复用离线核算的正式金额分摊逻辑（amount），不读取历史结果。"""
    from scripts.build_cost_match import allocate_all

    return allocate_all(rows, policy="amount")


def _short_basis(basis: str) -> str:
    """把分摊依据压成一句财务话。"""
    text = str(basis or "").strip()
    if not text:
        return ""
    head = text.split("（")[0].strip()
    return head[:40]


LOCAL_EXCEPTION_FILE = "拆单_本地异常.xlsx"
# 与 `scripts/run_split_knapsack.ERP_MISSING_CATEGORY` 同一口径：
# 睿贝没有这批数据 = 算不出来的异常，不是「已核实」。
ERP_MISSING_TYPE = "睿贝缺数据"


def _friendly_local_note(note: str, row_type: str = "") -> tuple[str, str]:
    """拆单引擎的本地异常说明 → 一句给财务看的话（异常类型, 异常明细）。"""
    text = str(note or "").strip()
    if str(row_type or "").strip() == ERP_MISSING_TYPE:
        return ERP_MISSING_TYPE, text[:80] or "睿贝里没有这批货的出运单或产品行，成本算不出来。"
    if "无出运产品行" in text:
        return "找不到出运单", "睿贝里没有这张出运单（可能还没生成），这条成本算不出来，请先核对合同号。"
    if "没有这张出运单" in text:
        return "找不到出运单", text[:60]
    if "子集和无精确解" in text or "凑不出" in text:
        return "拆单失败", "报关金额与睿贝出运金额凑不出精确对应，请核对这一单的报关单与出运单。"
    if "多解" in text:
        return "拆单失败", "有不止一种拆法都能凑出同额，需要人工确认。"
    if "费用口径不唯一" in text:
        return "拆单失败", "这一单的费用有多种分摊口径都能对上，需要人工确认。"
    if "ERP数据异常" in text or "反推失败" in text:
        return "拆单失败", "睿贝里这家供应商不生产这个产品，采购单号也对不上，需要人工确认。"
    return "拆单失败", text[:80]


def _local_exception_rows(split_xlsx: Path) -> list[dict]:
    """拆单失败的报关行也要出现在结果里。

    拆单引擎把「无出运产品行 / ERP 数据矛盾」写进 `拆单_本地异常.xlsx`，但那张表
    以前不进结果，财务在界面上根本看不到这条记录（选中的未核算行会凭空少一行）。
    这里把它补成正常的结果行，只给一个异常类型和一句异常明细。
    """
    path = split_xlsx.parent / LOCAL_EXCEPTION_FILE
    if not path.exists():
        return []
    out: list[dict] = []
    for item in read_split_rows(path):
        contract = str(item.get("合同号_1") or "").strip()
        note = str(item.get("说明") or "").strip()
        kind, detail = _friendly_local_note(note, str(item.get("类型") or ""))
        purchase = str(item.get("采购单号") or "").strip()
        purchase_link, shipment_link = erp_links(purchase, contract)
        out.append(
            {
                "报关单号": str(item.get("报关单号") or ""),
                "合同号_1": contract,
                "合同号（应收表格）": "",
                "外销订单号": contract,
                "产品类型": "",
                "报关品名": str(item.get("报关品名") or ""),
                "供应商简称": "",
                "供应商名称": "",
                "采购订单号": purchase,
                "报关金额": "",
                "采购金额": "",
                "报关重量": "",
                "出运金额合计": "",
                "出运采购金额合计": "",
                "异常类型": kind,
                "异常明细": detail,
                "分摊依据": "",
                "_睿贝采购单": purchase,
                "_睿贝出运单": contract,
                "_飞书链接": "",
                "_睿贝采购单链接": purchase_link,
                "_睿贝出运单链接": shipment_link,
            }
        )
    return out


def _batch_purchase_totals(contracts: set[str]) -> dict[str, Decimal]:
    """本批出运单里，各采购单的货值（RMB）。

    判「缺报关单」要拿**本批**该覆盖的货值做比较：同一张采购单常常分 A/B/C 多批出运，
    用它在本地缓存里所有出运单的累计货值去比，任何单批都会被误判成"还有报关单没拿到"。
    """
    from app.services.shipment_detail import load_shipment_detail

    totals: dict[str, Decimal] = collections.defaultdict(Decimal)
    visited: set[str] = set()
    for raw in contracts:
        for key in (raw, strip_pi(raw), f"PI-{strip_pi(raw)}"):
            if not key or key in visited:
                continue
            payload = load_shipment_detail(key)
            if not payload:
                continue
            visited.add(key)
            for line in payload.get("productList") or []:
                code = normalise_purchase_code(str(line.get("采购订单号") or "").strip())
                if code:
                    totals[code] += dec(line.get("出运采购金额(RMB)"))
            break
    return dict(totals)


def _shipment_fee_total(contract: str) -> Decimal | None:
    """出运单的客户费用合计；取不到出运明细时返回 None（不敢下结论）。"""
    from app.services.shipment_detail import load_shipment_detail

    for key in (contract, strip_pi(contract), f"PI-{strip_pi(contract)}"):
        if not key:
            continue
        payload = load_shipment_detail(key)
        if not payload:
            continue
        total = Decimal(0)
        for item in payload.get("expenseList") or []:
            total += dec(item.get("金额"))
        return total
    return None


def erp_sources(purchase_code: str, contract: str) -> dict:
    """与核算共用合同→出运单解析；跳转 ID 只取网页列表的数字主键。"""
    from app.services.erp_cache import CACHE_ROOT
    paths = (CACHE_ROOT / "purchases/purchases.jsonl", CACHE_ROOT / "shipments/shipments.jsonl")
    signature = tuple(path.stat().st_mtime_ns if path.exists() else 0 for path in paths)
    return _erp_sources_cached(str(purchase_code or "").strip(), str(contract or "").strip(), signature)


@lru_cache(maxsize=1)
def _erp_source_lists(signature: tuple) -> tuple[dict, dict]:
    from app.services.erp_cache import CACHE_ROOT, read_jsonl
    from app.services.shipment_detail import strip_pi
    purchases, shipments = {}, {}
    for kind, path, code_key, id_key in [
        (purchases, "purchases/purchases.jsonl", "purchase_code", "purchase_id"),
        (shipments, "shipments/shipments.jsonl", "invoiceCode", "shipmentId"),
    ]:
        for row in read_jsonl(CACHE_ROOT / path):
            if row.get("_remotePresent") is False:
                continue
            if re.fullmatch(r"[0-9]+", str(row.get(id_key) or "")):
                kind[strip_pi(row.get(code_key))] = row
    return purchases, shipments


@lru_cache(maxsize=2048)
def _erp_sources_cached(purchase_code: str, contract: str, signature: tuple) -> dict:
    from urllib.parse import urlencode
    from app.services.erp_web import ERP_ROOT
    from app.services.shipment_detail import shipment_invoices_for_contract, strip_pi
    purchases, shipments = _erp_source_lists(signature)
    purchase = purchases.get(strip_pi(purchase_code)) if purchase_code else None
    invoices = shipment_invoices_for_contract(contract) if contract else []
    # 精确同名也可直接定位，不需要合同拆分映射猜测。
    if not invoices and contract and strip_pi(contract) in shipments:
        invoices = [shipments[strip_pi(contract)]["invoiceCode"]]
    selected = []
    seen = set()
    for invoice in invoices:
        item = shipments.get(strip_pi(invoice))
        if item and item["shipmentId"] not in seen:
            selected.append(item)
            seen.add(item["shipmentId"])
    def link(kind, ident):
        return f"{ERP_ROOT}/{kind}_toUpdate?" + urlencode({"openWindow": "Y", "type": "view", "id": ident, "openhash": "costReview"})
    return {"purchase": {"purchases": [purchase] if purchase else []},
            "shipment": {"shipments": selected},
            "purchaseUrl": link("purchase", purchase["purchase_id"]) if purchase else "",
            "shipmentLinks": [{"invoiceCode": item["invoiceCode"], "url": link("shipment", item["shipmentId"])} for item in selected]}


def erp_links(purchase_code: str, contract: str) -> tuple[str, str]:
    sources = erp_sources(purchase_code, contract)
    links = sources["shipmentLinks"]
    return sources["purchaseUrl"], links[0]["url"] if len(links) == 1 else ""


def refresh_source_links(row: dict) -> dict:
    sources = erp_sources(row.get("采购订单号", ""), row.get("合同号_1", ""))
    links = sources["shipmentLinks"]
    return {**row, "_睿贝采购单链接": sources["purchaseUrl"],
            "_睿贝出运单链接": links[0]["url"] if len(links) == 1 else "",
            "_睿贝出运单来源": links}


def build_rows(split_xlsx: Path, amount_source: dict[str, str] | None = None) -> list[dict]:
    """拆单明细 → 财务核对行。`amount_source` 可传 {报关单号: 飞书记录链接}。"""
    rows = read_split_rows(split_xlsx)
    allocated = allocate_costs(rows)
    names = SupplierNames()
    amount_source = amount_source or {}

    by_po_goods: dict[str, Decimal] = collections.defaultdict(Decimal)
    for row in rows:
        code = str(row.get("采购单号") or "").strip()
        by_po_goods[code] += dec(row.get("出运采购金额合计"))
    # 本批（本次涉及的出运单）各采购单应有的货值：缺报关单要按这个口径判
    batch_totals = _batch_purchase_totals(
        {str(row.get("合同号_1") or "").strip() for row in rows if row.get("合同号_1")}
    )

    out: list[dict] = []
    for row, cost in zip(rows, allocated):
        po = str(row.get("采购单号") or "").strip()
        goods = dec(row.get("出运采购金额合计"))
        declared = dec(row.get("报关金额"))
        weight = dec(row.get("报关重量"))
        supplier_full = str(row.get("供应商") or "").strip()
        amount = cost.get("amount")
        total = dec(cost.get("total"))
        issues: list[str] = []

        if str(cost.get("basis") or "").startswith("批次成本缺失"):
            issues.append(("批次成本缺失", cost["basis"]))
        if goods <= 0:
            issues.append(("睿贝未定价", "睿贝出运单这一行没有采购金额，需要先补价。"))
        if total <= 0:
            issues.append(("缺入库单", "该采购单还没有可用的入库单，采购金额暂时算不出来。"))
        # 客户费用口径（2026-09-30 用户口径：没有明确找到额外费用就不能谎报费用）：
        # 「客户费用分摊」只有在睿贝出运单**确实登记了客户费用**时才算数；出运单没有任何费用
        # 却摊出了费用，说明是拿差额硬凑金额——必须报异常，绝不能当正常金额写回。
        fee_share = dec(row.get("客户费用分摊"))
        if abs(fee_share) > Decimal("1"):
            fee_total = _shipment_fee_total(str(row.get("合同号_1") or ""))
            if fee_total is not None and fee_total == 0:
                issues.append(
                    (
                        "金额对不上",
                        f"报关金额比睿贝出运金额多 {abs(fee_share):.2f}，但睿贝出运单里没有任何客户费用，"
                        "差额来源不明，请核对这一单的报关金额。",
                    )
                )
        # 本批该覆盖多少：优先用「本批出运单上这张采购单的货值」；
        # 出运明细取不到时退回整单累计值（宁可少报，也不要把单批误判成缺报关单）
        expected = batch_totals.get(po, Decimal(0))
        if expected > 0 and by_po_goods.get(po, Decimal(0)) + EQUAL < expected:
            issues.append(
                ("缺报关单", "这批还有报关单没拿到，采购金额可能分不完整。")
            )
        # 报关单号为空 = 这条来自「睿贝出运明细补充」（该批报关单未到齐），不按缺失算异常
        supplemented = not str(row.get("报关单号") or "").strip()
        if not str(row.get("产品类型") or "").strip():
            issues.append(("产品类型缺失", "这一行没有产品类型，请补上再核对。"))
        if weight <= 0 and not supplemented:
            issues.append(("缺重量", "报关单上没有这一行的重量，无法核对。"))

        if not issues:
            kind = "正常"
            detail = "本行金额来自睿贝出运明细（该批报关单未到齐）。" if supplemented else ""
        else:
            kind = issues[0][0]
            detail = issues[0][1]

        links = erp_links(po, str(row.get("合同号_1") or ""))
        out.append(
            {
                "_成本未核算": str(cost.get("basis") or "").startswith("批次成本缺失"),
                "报关单号": str(row.get("报关单号") or ""),
                "合同号_1": str(row.get("合同号_1") or ""),
                "合同号（应收表格）": purchase_key(po),
                "外销订单号": str(row.get("合同号_1") or ""),
                "产品类型": str(row.get("产品类型") or ""),
                "报关品名": str(row.get("报关品名") or ""),
                "供应商简称": str(row.get("供应商简称") or ""),
                "供应商名称": supplier_full,
                "采购订单号": po,
                "报关金额": f"{declared:.2f}" if declared else "",
                "采购金额": f"{amount:.2f}" if isinstance(amount, Decimal) and amount else "",
                "报关重量": f"{weight:.2f}".rstrip("0").rstrip(".") if weight else "",
                "出运金额合计": f"{dec(row.get('出运金额合计')):.2f}",
                "出运采购金额合计": f"{goods:.2f}",
                "异常类型": kind,
                "异常明细": detail,
                "分摊依据": _short_basis(cost.get("basis") or ""),
                "_睿贝采购单": po,
                "_睿贝出运单": str(row.get("合同号_1") or ""),
                "_飞书链接": amount_source.get(str(row.get("报关单号") or ""), ""),
                "_睿贝采购单链接": links[0],
                "_睿贝出运单链接": links[1],
            }
        )
    # 拆单失败的报关行也要露面（否则选中的行会凭空少一条）
    out.extend(_local_exception_rows(split_xlsx))
    return out


def summary(rows: list[dict]) -> dict:
    total = len(rows)
    bad = sum(1 for row in rows if row.get("异常类型") != "正常")
    amount = sum(
        (dec(row.get("采购金额")) for row in rows),
        Decimal(0),
    )
    return {
        "resultRows": total,
        "successRows": total - bad,
        "exceptionRows": bad,
        "purchaseTotal": float(amount),
    }
