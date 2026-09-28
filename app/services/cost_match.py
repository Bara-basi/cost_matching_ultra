"""采购金额匹配：用入库单「实发金额」给拆单记录定成本。

口径（2026-09-24 与用户确认）
============================

1. **金额来源 = 入库单「实发合计」**，不是 ERP 采购订单金额、也不是入库单里的
   订单数据区（那是下单时的预期金额）。一份采购单有多份有效入库单（分批 / 补款）时按份累加。
2. **金额单位 = 采购单**（一张采购单 = 一个工厂 = 一份采购合同）。
   同一采购单拆成多条记录时，把该采购单的实发总额摊到各条记录上。
3. 摊法**按下面顺序自上而下、命中即返回**；每条结果都带一句「分摊依据」写进 Excel，
   财务照着这一句就能复算：

  ① **批次区块直取**：报关行合同号里的批次记号（`26MT-01P242Y-C` → `C`）与入库单
      区块名（`实发数据-A`）全部对得上，且该区块金额与这一批报关行的金额能对上
      （±15%）→ 取该区块金额；
   ①′ **无归属费用整笔归位**（2026-09-24 新增，可选策略）：入库单里那些没有批次
      归属的费用行（绳子费/管帽费/木箱费…）不再按货值摊，而是整笔归到某一张报关单
      的记录上——前提是**这张采购单的报关单全齐**（我方记录的人民币货值合计 =
      睿贝整单出运人民币合计）。归到哪一张由调用方的策略决定（货值最大 / 最后一批），
      没指定就维持原来"按货值占比摊"；
   ② 按各记录自己的**出运采购金额(RMB)** 占比摊实发总额；
   ③ 本批报关单只装了整张采购单的一小部分（出运金额 < 50%）→ **不放大**，
      直接取出运采购金额本身，剩下的钱留给对应批次；
   ④ 入库单里的钱是**整单**的（实发总额更贴近整单 ERP 货值而不是手上这几张报关单）
      → 按整单 ERP 占比取；
   ⑤ 睿贝没给某行定价（该行 RMB 金额为 0）→ 该行留空交人工，其余行按占比摊满；
      整单覆盖且未定价行本身有分量时，按出运金额(USD) 占比摊。

输出：`scripts/build_cost_match.py` 调本模块，产出 `outputs/cost_match/` 下的
`成本匹配_全部 / _正常 / _异常 / _已核实` 四个文件。

**采购订单界面「费用信息」**（`purchases.jsonl` 的 `purchaseExpenseAmount`，= amount −
productAmount）默认**不加**到成本里：197 张有费用信息的采购单中，有入库单的 179 张里
107 张「实发 = 采购订单金额」（费用工厂已经开进票里了），加一遍就重复计费（异常 60 → 104）。
开关见 `COST_PO_FEE`，默认 `off`；`gap`/`cap` 是给单个异常单（`26MT-02K067`）做对照用的。
"""
from __future__ import annotations

import collections
import json
import os
import re
from dataclasses import dataclass
from decimal import Decimal
from typing import Any

from app.services.contract_shipments import cores_of
from app.services.erp_cache import CACHE_ROOT
from app.services.grn_extract import (
    PARSED_ROOT,
    blocks_to_amounts,
    line_amounts,
    line_quantities,
    safe_name,
)
from app.services.grn_select import erp_amount, order_key, select

# erp_cache.CACHE_ROOT 已经指向 `.cache/erp`
DETAIL_DIR = CACHE_ROOT / "details" / "shipments"

# 采购订单界面「费用信息」怎么参与成本（2026-09-24 用户确认：额外费用来自采购订单的费用信息）。
# 金额来源 = `purchases.jsonl` 的 `purchaseExpenseAmount`（= amount − productAmount），
# 一行一个采购单，正数=加、负数=赔款/折让；不需要再去出运单里拼 `purchaseExpenseList`。
#   off（默认）  只以入库单实发为准，不看采购订单费用（老口径，用来对照）
#   add         入库单实发 + 全部采购订单费用
#   gap         只补"入库单没覆盖的那段"：
#               应补 = 费用 − max(0, 入库单实发 − 采购单产品金额)，取正数才补
#   cap         目标 = max(入库单实发, 采购订单金额)：应补 = 费用 − (实发 − 产品金额)，取正
PO_FEE_MODE = os.environ.get("COST_PO_FEE", "off").strip().lower()
PO_FEE_ENABLED = PO_FEE_MODE in ("add", "gap", "cap", "1", "2", "true", "yes")
_PO_FEES: dict[str, Decimal] | None = None


def purchase_fees() -> dict[str, Decimal]:
    """采购单号 → 采购订单界面「费用信息」合计（元）。

    来源：`purchases/purchases.jsonl` 的 `purchaseExpenseAmount`
    （采购订单界面上「金额」= 产品金额 + 费用信息，所以它 = amount − productAmount）。
    只有 197 张采购单有这个字段；其余为空 = 没有费用信息。
    """
    global _PO_FEES
    if _PO_FEES is not None:
        return _PO_FEES
    table: dict[str, Decimal] = collections.defaultdict(Decimal)
    path = CACHE_ROOT / "purchases" / "purchases.jsonl"
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except ValueError:
                continue
            code = str(row.get("purchase_code") or "").strip()
            amount = dec(row.get("purchaseExpenseAmount"))
            if code and amount:
                table[code] += amount
    _PO_FEES = dict(table)
    return _PO_FEES


_PO_PRODUCT_AMOUNT: dict[str, Decimal] | None = None


def po_product_amounts() -> dict[str, Decimal]:
    """采购单号 → 采购订单界面的**产品金额**（purchases.jsonl 的 productAmount）。

    采购订单界面：`amount = productAmount（产品）+ 费用信息`，
    用它来判断"这笔费用是不是已经在入库单实发里了"。
    """
    global _PO_PRODUCT_AMOUNT
    if _PO_PRODUCT_AMOUNT is not None:
        return _PO_PRODUCT_AMOUNT
    table: dict[str, Decimal] = {}
    path = CACHE_ROOT / "purchases" / "purchases.jsonl"
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except ValueError:
                continue
            code = str(row.get("purchase_code") or "").strip()
            value = dec(row.get("productAmount"))
            if code and value:
                table[code] = value
    _PO_PRODUCT_AMOUNT = table
    return table

# 门槛（都来自实测对照，改动前请先跑一遍 scripts/build_cost_match.py 看差异条数）
BLOCK_COVERAGE_TOL = Decimal("0.15")   # 批次区块金额 vs 该批报关行金额，±15% 内算对得上
PARTIAL_COVER_RATIO = Decimal("0.5")   # 出运金额占比低于此值 = 只装了整单的一小部分
UNPRICED_MIN_SHARE = Decimal("0.05")   # 未定价行要占 ≥5% 出运金额才允许整单按 USD 摊
# 「小于 1 个货币单位的补充差额忽略不计」（2026-09-28 用户口径）：
# 业务为了把报关金额与出运金额抹平，会手工补一笔 ≤1 个货币单位的客户费用；
# 这种零头不参与判定——批次区块覆盖、整单覆盖、费用归位等"齐不齐"的比较都按它放宽。
IGNORABLE_DIFF = Decimal("1")


def dec(value: Any) -> Decimal:
    """宽松转 Decimal（空值/千分位/脏文本都吞掉，取不到算 0）。"""
    if value in (None, ""):
        return Decimal(0)
    try:
        return Decimal(str(value).replace(",", "").strip())
    except Exception:  # noqa: BLE001
        return Decimal(0)


def unit_core(code: str) -> str:
    """比对用的采购单核心：去掉年份前缀与佣金标记 Y。

    飞书里同一张采购单常被写成不同年份（`25MT-03P495Y-ADD1` vs 我们的
    `26MT-03P495Y-ADD1-HD`）、或带/不带佣金 `Y`（`26MT-02N126Y` vs
    `26MT-02N126-YH`）；不归一就会把「同一张单」判成「飞书无对应行」。
    """
    cores = cores_of(str(code or ""))
    core = sorted(cores)[0] if cores else str(code or "").strip()
    text = re.sub(r"^\d{2}(?=[A-Za-z]{2})", "", core)
    return re.sub(r"Y(?=-|$)", "", text, flags=re.IGNORECASE)


# --------------------------------------------------------------------------- ERP 出运产品行金额


@dataclass(frozen=True)
class ShipmentMoney:
    """ERP 出运产品行的金额索引（一张采购单一行金额，一次扫盘建好）。"""

    rmb_by_po: dict[str, Decimal]           # 采购单 → 整单出运采购金额(RMB)
    rmb_by_kind: dict[tuple[str, str], Decimal]  # (采购单, 海关商品) → 出运采购金额(RMB)
    usd_by_po: dict[str, Decimal]           # 采购单 → 整单出运金额(USD)


_MONEY: ShipmentMoney | None = None


def shipment_money() -> ShipmentMoney:
    """扫一遍 `.cache/erp/details/shipments/`，同时建三个金额索引。

    以前是三个函数各扫一遍同样的 1000+ 个 JSON（3 倍 I/O），现在只扫一次。
    """
    global _MONEY
    if _MONEY is not None:
        return _MONEY
    rmb_by_po: dict[str, Decimal] = collections.defaultdict(Decimal)
    rmb_by_kind: dict[tuple[str, str], Decimal] = collections.defaultdict(Decimal)
    usd_by_po: dict[str, Decimal] = collections.defaultdict(Decimal)
    if DETAIL_DIR.exists():
        for path in DETAIL_DIR.glob("*.json"):
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except ValueError:
                continue
            for row in payload.get("productList") or []:
                code = str(row.get("采购订单号") or "").strip()
                if not code:
                    continue
                rmb = dec(row.get("出运采购金额(RMB)"))
                rmb_by_po[code] += rmb
                usd_by_po[code] += dec(row.get("出运金额"))
                kind = str(row.get("海关商品（中文）") or "").strip()
                if kind:
                    rmb_by_kind[(code, kind)] += rmb
    _MONEY = ShipmentMoney(
        rmb_by_po=dict(rmb_by_po), rmb_by_kind=dict(rmb_by_kind), usd_by_po=dict(usd_by_po)
    )
    return _MONEY


# --------------------------------------------------------------------------- 报关行批次记号

ORDER_TOKEN_RE = re.compile(
    r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:[\s\-_]?Y)?(?:[\s\-_]*ADD\d*)?"
)


def _token_of(text: str) -> str:
    """从一段单号里取批次记号：`26MT-10E101-Y-A` → `A`；`477G` → `G`。

    先把「单号核心（含佣金 Y、ADD）」整段剥掉，剩下的第一个独立字母就是批次。
    """
    stripped = ORDER_TOKEN_RE.sub("", str(text or ""), count=1)
    match = re.search(r"(?<![A-Za-z])([A-Za-z])(?![A-Za-z0-9])", stripped)
    if match:
        return match.group(1).upper()
    chinese = re.search(r"第([一二三四五六七八九十\d]+)批", str(text or ""))
    return "第" + chinese.group(1) + "批" if chinese else ""


def record_batch_token(contract: str, purchase_code: str = "") -> str:
    """报关行合同号里的批次记号：`26MT-01P242Y-C` → `C`。

    联合号（`25MT-07F379Y-F&477G`）要挑出属于本采购单的那一段：
    按流水号匹配（本单 `25MT-07F477` → 取 `477G` → `G`）。
    """
    text = str(contract or "").strip()
    if not text:
        return ""
    parts = [part for part in re.split(r"[&,，;；]", text) if part.strip()]
    if len(parts) > 1 and purchase_code:
        serial = re.search(r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z](\d{2,3})", str(purchase_code))
        if serial:
            for part in parts:
                if serial.group(1) not in part:
                    continue
                token = _token_of(part.strip())
                if token:
                    return token
                # 联合号里这一段常常只剩「流水 + 批次」（`477G`），直接取末尾字母
                letters = re.findall(r"([A-Za-z])(?!\w)", part.strip())
                if letters:
                    return letters[-1].upper()
    for part in parts:
        token = _token_of(part.strip())
        if token:
            return token
    return ""


# --------------------------------------------------------------------------- 入库单实发总额


def batch_amounts_of(kept_items: list[dict]) -> dict[str, Decimal]:
    """一个采购单全部有效入库单的「批次记号 → 金额」（跨附件合并）。"""
    blocks = [
        block
        for item in kept_items
        for block in (item.get("blocks") or [])
    ]
    return blocks_to_amounts(blocks)


def costs_for(codes: list[str]) -> dict[str, tuple[Decimal, dict[str, Any]]]:
    """批量算每个采购单的入库单实发合计；返回 {采购单: (金额, 明细)}。

    同一份入库单被多个采购单引用时（实测有 3 份附件同时挂在主单与附加单下，
    例如 `26MT-06N145-GYL ADD1 溪流供应链 入库单`），两边都算就会重复计成本，
    所以按 ERP 采购金额当权重拆分——它本身就是这两张采购单的合同金额。
    """
    selections = {code: select(code) for code in codes}

    # 附加单目录里放的是主单的合并入库单（文件自己的 ADD 标记与采购单不符）时，
    # 这张表里的金额无法拆到本单，退回该采购单的 ERP 采购金额并注明。
    overrides: dict[str, Decimal] = {}
    for code, result in selections.items():
        kept = result["kept"]
        if not kept:
            continue
        mine = _add_marker(code)
        if mine and all(_add_marker(item["original"]) != mine for item in kept):
            reference = erp_amount(code)
            if reference:
                overrides[code] = reference
                result["issues"].append(
                    "附件是主单的合并入库单（不含本附加单标记），"
                    f"本单按 ERP 采购金额 {reference} 取值，待人工确认"
                )

    # 按内容指纹分堆：同一份附件（同 digest）被多张采购单引用时只能算一次
    by_digest: dict[str, list[tuple[str, dict[str, Any]]]] = collections.defaultdict(list)
    fees_by_digest: dict[str, Decimal] = {}
    fee_items_by_digest: dict[str, list[dict[str, str]]] = {}
    # 采购单 → 入库单里的费用行明细（费用名 + 金额），只在「待核实」表里展示用
    fee_items: dict[str, list[dict[str, str]]] = collections.defaultdict(list)
    # 采购单 → 入库数量（按单位族），给出运数量核验用
    qty_by_family: dict[str, dict[str, Decimal]] = collections.defaultdict(
        lambda: collections.defaultdict(Decimal)
    )
    block_qty: dict[str, Decimal] = collections.defaultdict(Decimal)
    own_totals: dict[str, Decimal] = {code: Decimal(0) for code in codes}
    own_fees: dict[str, Decimal] = {code: Decimal(0) for code in codes}
    for code, result in selections.items():
        for item in result["kept"]:
            if code in overrides:
                continue  # 已用 ERP 金额兜底，不再参与附件拆分
            digest = item.get("digest") or ""
            if digest:
                by_digest[digest].append((code, item))
                fees_by_digest[digest] = max(
                    fees_by_digest.get(digest, Decimal(0)),
                    _unattributed_fee(item, code),
                )
                fee_items_by_digest.setdefault(digest, _fee_items(item, code))
                for family, value in _quantities(item, code).items():
                    qty_by_family[code][family] += value
                block_qty[code] += _block_qty(item, code)
            else:
                own_totals[code] += dec(item["amount"])
                own_fees[code] += _unattributed_fee(item, code)
                fee_items[code].extend(_fee_items(item, code))
                for family, value in _quantities(item, code).items():
                    qty_by_family[code][family] += value
                block_qty[code] += _block_qty(item, code)

    out: dict[str, tuple[Decimal, dict[str, Any]]] = {}
    for code, result in selections.items():
        out[code] = (
            own_totals[code],
            {
                "files": result["files"],
                "kept": len(result["kept"]),
                "dropped": len(result["dropped"]),
                "batches": len(result["groups"]),
                "kept_names": [item["original"] for item in result["kept"]],
                "issues": list(result["issues"]),
                "fee_items": list(fee_items[code]),
                "batch_amounts": {
                    token: str(amount)
                    for token, amount in batch_amounts_of(result["kept"]).items()
                },
                "block_count": sum(
                    len(item.get("blocks") or []) for item in result["kept"]
                ),
            },
        )
    for code, amount in overrides.items():
        out[code] = (amount, out[code][1])

    for refs in by_digest.values():
        owners = sorted({code for code, _item in refs})
        amount = max((dec(item["amount"]) for _code, item in refs), default=Decimal(0))
        if len(owners) == 1:
            code = owners[0]
            out[code][1]["issues"].append(
                f"附件 {refs[0][1]['original']} 计入本单（唯一引用）"
            )
            out[code] = (out[code][0] + amount, out[code][1])
            own_fees[code] += fees_by_digest.get(refs[0][1].get("digest") or "", Decimal(0))
            fee_items[code].extend(
                fee_items_by_digest.get(refs[0][1].get("digest") or "", [])
            )
            continue
        weights = {code: (erp_amount(code) or Decimal(0)) for code in owners}
        total_weight = sum(weights.values(), Decimal(0))
        if total_weight > 0 and all(value > 0 for value in weights.values()):
            basis = "按 ERP 采购金额拆分"
        else:
            weights = {code: Decimal(1) for code in owners}
            total_weight = Decimal(len(owners))
            basis = "无 ERP 金额，按份数均分"
        for code in owners:
            share = (amount * weights[code] / total_weight).quantize(Decimal("0.01"))
            fee = fees_by_digest.get(refs[0][1].get("digest") or "", Decimal(0))
            own_fees[code] += (fee * weights[code] / total_weight).quantize(Decimal("0.01"))
            fee_items[code].extend(fee_items_by_digest.get(refs[0][1].get("digest") or "", []))
            name = next(item["original"] for owner, item in refs if owner == code)
            out[code][1]["issues"].append(
                f"附件 {name} 被 {len(owners)} 个采购单共用，{basis}，本单分得 {share}"
            )
            out[code] = (out[code][0] + share, out[code][1])
    # 无归属费用（实发 − 产品明细求和）挂在明细里，供成本匹配决定费用落位
    for code in codes:
        out[code][1]["unattributed_fee"] = str(own_fees[code]) if own_fees[code] > 0 else ""
        # 费用行明细要等「按内容指纹分堆」跑完才是全的（同 digest 的附件在上面才并进来）
        out[code][1]["fee_items"] = list(fee_items[code])
        out[code][1]["qty_by_family"] = {
            family: str(value) for family, value in qty_by_family[code].items()
        }
        out[code][1]["block_qty"] = str(block_qty[code])
    # 采购订单界面「费用信息」（试样费/检验费/赔款…）是否计入成本（默认 off，只以入库单实发为准）
    if PO_FEE_ENABLED:
        fees = purchase_fees()
        products = po_product_amounts()
        for code in codes:
            fee = fees.get(code, Decimal(0))
            if not fee:
                continue
            if PO_FEE_MODE in ("gap", "cap"):
                # 入库单实发比「采购单产品金额」多出来的那部分，就是入库单自己已经算进去的钱。
                # gap：只把「多出来的」算作已含费用（实发低于产品金额时不认，避免被单价差异带跑）；
                # cap：把差额整个算作已含费用，效果 = 目标成本 max(入库单实发, 采购订单金额)。
                inside = out[code][0] - products.get(code, Decimal(0))
                if PO_FEE_MODE == "gap" and inside < 0:
                    inside = Decimal(0)
                fee = fee - inside
            if fee <= 0:
                continue
            out[code] = (out[code][0] + fee, out[code][1])
            out[code][1]["purchase_fee"] = str(fee)
            out[code][1]["issues"].append(f"另加采购订单界面费用信息 {fee}")
    return out


def _unattributed_fee(item: dict[str, Any], code: str) -> Decimal:
    """一份入库单里「没有产品行归属的费用」= 实发合计 − 产品明细金额之和。

    绳子费/管帽费/木箱费这类费用行在单据里是独立行、不属于任何产品行；
    不用解析结果的 `extra_fees`，因为它可能含没进实发合计的补款/批注行（会算大）。
    """
    path = PARSED_ROOT / safe_name(code) / (safe_name(item["file"], "file") + ".json")
    if not path.exists():
        return Decimal(0)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return Decimal(0)
    detail_sum = line_amounts(payload)
    if not detail_sum:
        return Decimal(0)
    diff = dec(item["amount"]) - detail_sum
    return diff if diff > 0 else Decimal(0)


def _fee_items(item: dict[str, Any], code: str) -> list[dict[str, str]]:
    """这份入库单里解析出来的**费用行明细**（费用名 + 金额），给「待核实」表用。

    只做展示与留痕：金额口径仍然只用 `_unattributed_fee`（实发 − 产品明细求和）。
    """
    path = PARSED_ROOT / safe_name(code) / (safe_name(item["file"], "file") + ".json")
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return []
    out: list[dict[str, str]] = []
    for fee in payload.get("extra_fees") or []:
        if not isinstance(fee, dict):
            continue
        label = str(fee.get("label") or "").strip()
        amount = dec(fee.get("amount"))
        if label or amount:
            out.append({"file": str(item.get("original") or ""), "label": label, "amount": str(amount)})
    return out


def _quantities(item: dict[str, Any], code: str) -> dict[str, Decimal]:
    """这份入库单的数量（按单位族：count/length/weight），给出运数量核验用。"""
    path = PARSED_ROOT / safe_name(code) / (safe_name(item["file"], "file") + ".json")
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return {}
    return line_quantities(payload)


def _block_qty(item: dict[str, Any], code: str) -> Decimal:
    """这份入库单**实发区块**的数量合计（单据自己写的"实发数量"，按批次累加）。

    `26MT-06H132` 这种"一批一张单、行里只登记本批数量"的表，只有区块数量是完整的，
    明细列加起来会少（实测少 9 支），所以数量核验把区块数量也作为一路口径。
    """
    path = PARSED_ROOT / safe_name(code) / (safe_name(item["file"], "file") + ".json")
    if not path.exists():
        return Decimal(0)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return Decimal(0)
    return sum(
        (dec(block.get("qty")) for block in (payload.get("settled_blocks") or [])),
        Decimal(0),
    )


def _add_marker(text: str) -> str:
    """这段文字指的是哪张附加单（`ADD1`/`ADD2`…），没有则空串。

    不能只按 `order_key()` 的 ADD 段判断：`26MT-06N145-GYL ADD1 溪流供应链 入库单`
    这种「ADD 在工厂后缀后面」的写法 `order_key` 认不出，会被误判成「目录里没有本
    附加单标记」，于是主单把整份入库单全吃下（重复计成本）。这里补一层宽松匹配。
    """
    key = order_key(str(text or ""))
    for part in key.split("#")[1:]:
        if part.startswith("ADD"):
            return "ADD1" if part == "ADD" else part
    match = re.search(r"ADD\d*", str(text or ""), re.IGNORECASE)
    if match:
        token = match.group(0).upper()
        return "ADD1" if token == "ADD" else token
    return ""


# --------------------------------------------------------------------------- 分摊


@dataclass(frozen=True)
class RecordCost:
    """一条拆单记录参与分摊所需的最小信息。"""

    amount_rmb: Decimal = Decimal(0)  # 本记录出运采购金额合计（ERP 人民币货值）
    amount_usd: Decimal = Decimal(0)  # 本记录出运金额合计（USD）
    batch: str = ""                   # 报关行合同号里的批次记号
    product_type: str = ""
    purchase_code: str = ""


def _split(weights: list[Decimal], total: Decimal, source: str) -> list[tuple[Decimal, str]]:
    """按权重把 total 摊到各行（权重全为 0 时均分）。"""
    weight_sum = sum(weights, Decimal(0))
    if weight_sum <= 0:
        weights = [Decimal(1)] * len(weights)
        weight_sum = Decimal(len(weights))
        source = "均分"
    return [
        (
            (total * weight / weight_sum).quantize(Decimal("0.01")),
            f"{source} {weight / weight_sum * 100:.4f}%",
        )
        for weight in weights
    ]


def _by_batch_blocks(
    records: list[RecordCost], total: Decimal, batch_amounts: dict[str, Decimal]
) -> list[tuple[Decimal, str]] | None:
    """规则①：报关批次记号与入库单区块名全部对上时，按区块直取。

    区块金额要和各报关行的人民币金额「对得上」才敢直取（±15%）：实测有采购单的
    区块字母与报关批次并不同名（B 区块对应的是另一批货），只看单边阈值会放过错配。
    """
    tokens = [record.batch for record in records]
    if not batch_amounts or not all(token and token in batch_amounts for token in tokens):
        return None
    # 各批次金额之和不可能超过这张采购单自己的成本（超过说明区块属于别的采购单，
    # 典型：附加单目录里挂着主单的合并入库单）；按去重后的批次比较。
    if sum((batch_amounts[token] for token in set(tokens)), Decimal(0)) > total * Decimal("1.05"):
        return None

    grouped: dict[str, list[int]] = collections.defaultdict(list)
    for position, token in enumerate(tokens):
        grouped[token].append(position)
    results: list[tuple[Decimal, str] | None] = [None] * len(tokens)
    for token, positions in grouped.items():
        amount = batch_amounts[token]
        weights = [records[position].amount_rmb for position in positions]
        weight_sum = sum(weights, Decimal(0))
        alone = len(positions) == 1
        if alone and weight_sum <= 0:
            # 批次记号已唯一确定，睿贝又没给这行定价 → 直接取区块金额
            results[positions[0]] = (
                amount.quantize(Decimal("0.01")),
                f"入库单区块-{token}（睿贝该行无金额，按批次直取）",
            )
            continue
        covered = (
            weight_sum > 0
            and amount > 0
            and (
                abs(weight_sum - amount) / amount <= BLOCK_COVERAGE_TOL
                # 「小于 1 个货币单位的补充差额忽略不计」（2026-09-28 用户口径）：
                # 区块金额比我方这几条记录多/少不到 1 元，就当这一批的报关单是齐的
                or abs(weight_sum - amount) <= IGNORABLE_DIFF
            )
        )
        if not covered:
            # 这一批还出了别的报关单（那张单不在这批数据里）→ 不能把整批金额摊给眼前的记录
            for position, weight in zip(positions, weights):
                results[position] = (
                    weight.quantize(Decimal("0.01")),
                    f"入库单区块-{token} 只覆盖一部分（{weight_sum} / {amount}），按本行金额取",
                )
            continue
        if alone:
            results[positions[0]] = (amount.quantize(Decimal("0.01")), f"入库单区块-{token}")
            continue
        # 同一批次落在多张报关单上 → 按各报关行的出运采购金额占比拆这一个区块
        weight_sum = sum(weights, Decimal(0))
        if weight_sum <= 0:
            weights = [Decimal(1)] * len(positions)
            weight_sum = Decimal(len(positions))
            basis = f"入库单区块-{token}（无权重，均分）"
        else:
            basis = f"入库单区块-{token}（按出运采购金额占比）"
        for position, weight in zip(positions, weights):
            results[position] = ((amount * weight / weight_sum).quantize(Decimal("0.01")), basis)
    if any(item is None for item in results):
        return None
    return [item for item in results if item is not None]


def _split_unpriced(
    records: list[RecordCost],
    weights: list[Decimal],
    total: Decimal,
    coverage: Decimal | None,
    source: str,
) -> list[tuple[Decimal, str]] | None:
    """规则⑤：睿贝没给部分行定价时，怎么把这些行的金额处理掉。返回 None = 不走这条路。"""
    missing = [position for position, weight in enumerate(weights) if weight <= 0]
    if not missing or sum(weights, Decimal(0)) <= 0:
        return None
    usd = [record.amount_usd for record in records]
    usd_total = sum(usd, Decimal(0))
    if coverage is not None and coverage >= PARTIAL_COVER_RATIO and usd_total > 0:
        # 这几张报关单已覆盖整单，缺的只是睿贝那一行没定价 → 用出运金额(USD) 全组分摊。
        # 但未定价行本身要有分量（≥5%）；只是一条几十 USD 的零头时留空让人工看更稳妥。
        shares = [weight / usd_total for weight in usd]
        if all(shares[position] >= UNPRICED_MIN_SHARE for position in missing):
            return [
                (
                    (total * weight / usd_total).quantize(Decimal("0.01")),
                    f"按出运金额(USD)占比分摊 {weight / usd_total * 100:.4f}%（睿贝部分行未定价）",
                )
                for weight in usd
            ]
    kept = [position for position in range(len(weights)) if position not in missing]
    share_total = sum((weights[position] for position in kept), Decimal(0))
    results: list[tuple[Decimal, str]] = [
        (Decimal(0), "睿贝该行未定价（采购金额留空，待人工）")
        if position in missing
        else (
            (total * weights[position] / share_total).quantize(Decimal("0.01")),
            f"{source} {weights[position] / share_total * 100:.4f}%（其余行未定价）",
        )
        for position in range(len(weights))
    ]
    return results


def allocate_amounts(
    records: list[RecordCost],
    total: Decimal,
    batch_amounts: dict[str, Decimal] | None = None,
    rmb_by_kind: dict[tuple[str, str], Decimal] | None = None,
    po_usd_total: Decimal | None = None,
    po_rmb_total: Decimal | None = None,
    *,
    unattributed_fee: Decimal = Decimal(0),
    fee_target: int | None = None,
    fee_note: str = "",
) -> list[tuple[Decimal, str]]:
    """把一张采购单的入库单实发总额摊到它的各条拆单记录上。

    `total` = 该采购单的实发合计；返回 `[(金额, 分摊依据)]`，与 `records` 一一对应。
    规则见模块开头；这里只是把它们按顺序摆开。

    `unattributed_fee` / `fee_target` / `fee_note`：规则①′ 用——把入库单里**没有批次
    归属的费用行**整笔记到第 `fee_target` 条记录上（`fee_note` 说明为什么是它）。
    调用方只在"报关单全齐"时才传这两个参数，否则保持按占比摊。
    """
    if not records:
        return []
    direct = _by_batch_blocks(records, total, batch_amounts or {})
    if direct is not None:
        # 区块自带批次归属，费用已经在区块金额里，不再单独归位
        return direct
    if (
        unattributed_fee > 0
        and fee_target is not None
        and 0 <= fee_target < len(records)
        and unattributed_fee < total
    ):
        # 规则①′：先把「货值部分」按正常规则摊，再把无归属费用整笔加到目标记录上
        base = allocate_amounts(
            records,
            total - unattributed_fee,
            None,
            rmb_by_kind,
            po_usd_total,
            po_rmb_total,
        )
        amount, basis = base[fee_target]
        base[fee_target] = (
            amount + unattributed_fee,
            f"{basis}；含无归属费用 {unattributed_fee}（{fee_note}）",
        )
        return base

    weights = [record.amount_rmb for record in records]
    source = "本记录出运采购金额(RMB)"
    usd_total = sum((record.amount_usd for record in records), Decimal(0))
    coverage = (usd_total / po_usd_total) if po_usd_total else None
    if sum(weights, Decimal(0)) > 0:
        # 有金额、但个别行没定价 → 先决定这些行怎么办
        unpriced = _split_unpriced(records, weights, total, coverage, source)
        if unpriced is not None:
            return unpriced
    else:
        # 规则②的兜底：睿贝没给出运采购金额(RMB) 时，改用 (采购单, 产品类型) 的 ERP 金额
        weights = [
            (rmb_by_kind or {}).get((record.purchase_code, record.product_type), Decimal(0))
            for record in records
        ]
        source = "ERP出运采购金额(RMB)"
        if sum(weights, Decimal(0)) <= 0:
            weights = [record.amount_usd for record in records]
            source = "出运金额(USD)"

    # 手上这几张报关单只装了整张采购单的一小部分 → 不放大，直接取出运采购金额本身，
    # 未覆盖的部分留给对应批次（实测：只装了整单 9.8% / 15.5% 的单子都是这样）。
    weight_sum = sum(weights, Decimal(0))
    partial = (
        coverage < PARTIAL_COVER_RATIO
        if coverage is not None
        else total > 0 and weight_sum < total * PARTIAL_COVER_RATIO
    )
    if total > 0 and weight_sum > 0 and partial:
        return [
            (
                weight.quantize(Decimal("0.01")),
                f"{source} 直接取值（本单只覆盖采购单一部分：{weight_sum} / {total}）",
            )
            for weight in weights
        ]

    # 入库单里的钱是整单的：实发总额更贴近「整单 ERP 货值」而不是手上这几张报关单的货值
    # → 按整单占比取，否则会把别人那部分的钱摊到眼前这几条记录上。
    if (
        total > 0
        and po_rmb_total
        and po_rmb_total > weight_sum * Decimal("1.005")
        and abs(total - po_rmb_total) < abs(total - weight_sum)
    ):
        return [
            (
                (total * weight / po_rmb_total).quantize(Decimal("0.01")),
                f"{source} 按整单ERP占比 {weight / po_rmb_total * 100:.4f}%"
                f"（入库单覆盖整单 {po_rmb_total}，本单只装其中 {weight_sum}）",
            )
            for weight in weights
        ]
    return _split(weights, total, source)
