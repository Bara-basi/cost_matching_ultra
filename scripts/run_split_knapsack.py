"""背包拆单：按报关金额把出运产品行分配到报关单，再落回报关行。

容量口径（实测确认）：
    报关行容量 = 报关金额（报关单总价）
    报关单容量 = 该报关单各行报关金额之和 − 本次出运单运费均摊（按报关单张数平分）

流程：
1. 按合同取「我解析的报关行」与「出运产品行」；
2. 以报关单为单位跑 meet-in-the-middle 子集和（含启发式排序）；
3. 命中后，在报关单内部按供应商/海关编码/品名把产品行落到具体报关行；
4. 与飞书「2026年报关数据」的供应商简称集合比对。

输出：outputs/shipments_split/ 下 拆单明细_全部 / 拆单结果_正常 / 拆单结果_异常。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, replace
from typing import Any
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from app.services.freight import (  # noqa: E402
    FEE_DIRECTION_DEFAULTS,
    expense_cents,
    expense_kinds,
    fee_directions,
)
from app.services.knapsack_split import (  # noqa: E402
    FREIGHT_MODES,
    allocate_fee,
    assign_to_rows,
    bundle_lines,
    decl_category,
    iter_lines,
    line_product_type,
    name_matches,
    solve,
    solve_residual,
    to_cents,
)
from app.services.shipment_index import (  # noqa: E402
    canonical,
    load_lines,
    split_merged_parts,
    strip_pi,
)
from app.services.contract_shipments import (  # noqa: E402
    container_shipments,
    cores_of,
    expand_contract,
    line_cores,
    literal_parts,
    lookup_original_by_parts,
    order_core,
    shipments as shipment_refs,
)
from app.services.shipment_detail import (  # noqa: E402
    ensure_shipment_detail,
    names_batch,
    shipment_invoices_for_contract,
)
from app.services.shipment_index import normalise_purchase_code  # noqa: E402
from app.services.supplier_names import SupplierNames  # noqa: E402

CACHE = PROJECT_ROOT / "data" / "cache"
REF_PATH = CACHE / "records_full_match_ref.json"
OUT_DIR = PROJECT_ROOT / "outputs" / "shipments_split"
# 拆单输入 = 报关单解析结果（出口退税联）
PARSE_XLSX = PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_出口退税联.xlsx"
# 已知 PDF 解析不全导致单号错误的两条，直接过滤
BAD_CONTRACTS = {
    "26mt-03p200y-a&229y-a&245y-a&",
    "26mt-03p200y-a&229y-a&245y-a&262",
}


@dataclass(frozen=True)
class DeclaredLine:
    """一条报关商品行（拆单的输入）。"""

    source_file: str
    contract: str
    declaration_no: str
    product_name: str
    hs_code: str
    weight: Decimal
    amount: Decimal
    currency: str = ""
    product_type: str = ""


def flatten(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        return " | ".join(filter(None, (flatten(item) for item in value)))
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return flatten(value[key])
    return ""


def read_declared(parse_xlsx: Path | None = None) -> list[DeclaredLine]:
    """读报关单解析结果 → 报关商品行（含去重与「超范围合同号」二次拦截）。"""
    from openpyxl import load_workbook

    from app.services.scope import out_of_scope_reason

    workbook = load_workbook(parse_xlsx or PARSE_XLSX, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    header = [str(cell) for cell in rows[0]]
    index = {name: position for position, name in enumerate(header)}
    lines: list[DeclaredLine] = []
    seen: set[tuple] = set()
    for row in rows[1:]:
        def cell(name: str) -> Any:
            position = index.get(name)
            return row[position] if position is not None and row[position] is not None else ""

        contract = str(cell("合同号_1")).strip()
        if not contract or contract.lower() in BAD_CONTRACTS:
            continue
        if out_of_scope_reason(contract):
            continue
        weight = str(cell("报关重量") or 0).replace(",", "") or "0"
        declared = str(cell("总价") or 0).replace(",", "") or "0"
        try:
            weight_value = Decimal(weight)
        except Exception:  # noqa: BLE001
            weight_value = Decimal(0)
        try:
            amount_value = Decimal(declared)
        except Exception:  # noqa: BLE001
            amount_value = Decimal(0)
        # 同一张报关单在 data/raw 下有重复 PDF 时会产生重复行：按
        # 「报关单号 + 合同号 + 品名 + 重量 + 金额」判重，只保留第一条
        marker = (str(cell("报关单号")), contract, str(cell("报关品名")), weight, declared)
        if marker in seen:
            continue
        seen.add(marker)
        lines.append(
            DeclaredLine(
                source_file=str(cell("来源文件")),
                contract=contract,
                declaration_no=str(cell("报关单号")),
                product_name=str(cell("报关品名")),
                hs_code=str(cell("海关编码")),
                weight=weight_value,
                amount=amount_value,
                currency=str(cell("币种")),
                product_type=str(cell("产品类型") or ""),
            )
        )
    return lines
# 报关金额与 ERP 出运金额之间允许的舍入差（0.5 元）
TOLERANCE_CENTS = 50
# 「小于 1 个货币单位的补充差额忽略不计」（2026-09-28 用户口径）：业务为了把
# 报关金额与出运金额抹平，会手工补一笔 ≤1 个货币单位的客户费用（如 0.01）。
# 这种零头不参与"齐不齐"的比较——定向兜底里「一组货 ↔ 多张报关行」的金额核对也按它放宽。
IGNORABLE_DIFF_CENTS = 100
# 「极小客户费用」阈值（元）：业务上为了把报关金额与出运金额抹平，会手工补一笔
# 很小的客户费用（如 0.01，业务确认**一定在 1 个货币单位以内**）。
# 这类单据的金额关系是**精确**的，出现它就禁用容差。
TINY_FEE_YUAN = 1.0
# 本地缺出运明细时，是否按需从睿贝 MCP 临时抓取
ONLINE_TOPKUP = True
# 粒度由粗到细：整厂货物 → 工厂+产品类型 → 逐条产品行
LEVELS: list[tuple[str, object]] = [
    ("按供应商打包", lambda line, n: n.short(line.supplier) or "(未知供应商)"),
    ("按供应商+产品类型打包", lambda line, n: (
        n.short(line.supplier) or "(未知供应商)",
        # 打包仍以 ERP 出运品名为主（求解器一直在用的口径），缺名时才用商品资料粗分类兜底
        line.customs_name or line_product_type(line) or line.hs_code,
    )),
    ("逐条产品行", None),
]


def flat(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        return " | ".join(filter(None, (flat(v) for v in value)))
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return flat(value[key])
    return ""


_PURCHASE_SUPPLIER: dict[str, str] | None = None
_SUPPLIER_BY_CODE: dict[str, str] | None = None


def supplier_by_code() -> dict[str, str]:
    """供应商编码 → 名称（出运产品行常只有编码、没有名称，缺了会把同一家工厂拆成两组）。"""
    global _SUPPLIER_BY_CODE
    if _SUPPLIER_BY_CODE is None:
        from app.services.shipment_detail import supplier_name_map

        _SUPPLIER_BY_CODE = supplier_name_map()
    return _SUPPLIER_BY_CODE


def purchase_supplier_map() -> dict[str, str]:
    """采购单号 → 供应商名称（原生 `supplierName`，取自采购列表缓存）。"""
    global _PURCHASE_SUPPLIER
    if _PURCHASE_SUPPLIER is None:
        from app.services.erp_cache import CACHE_ROOT, read_jsonl

        mapping: dict[str, str] = {}
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            code = str(row.get("purchase_code") or "").strip()
            name = str(row.get("supplierName") or "").strip()
            if not code or not name:
                continue
            mapping.setdefault(code, name)
            mapping.setdefault(canonical(code), name)
        _PURCHASE_SUPPLIER = mapping
    return _PURCHASE_SUPPLIER


def purchase_supplier_name(row: dict) -> str:
    """出运产品行不带供应商名称时，用采购订单号回查原生的供应商名称。"""
    code = str(row.get("采购订单号") or "").strip()
    if not code:
        return ""
    table = purchase_supplier_map()
    for key in (code, canonical(code)):
        if key in table:
            return table[key]
    # 采购单号写法差异（带 `-HD` / `-供应链` 等后缀）时用前缀唯一命中兜底
    hits = {name for key, name in table.items() if key.startswith(code)}
    return hits.pop() if len(hits) == 1 else ""


_SHIPMENT_TOTALS: dict[str, int] | None = None

_PRODUCT_CATEGORY: dict[str, str] | None = None


def product_category(code: Any) -> str:
    """产品编码 → 商品资料里的粗分类（类别名称去部门括号）；没有缓存返回空串。

    数据由 `scripts/fetch_product_categories.py` 抓取到
    `.cache/erp/details/products/`，这里是进程内一次性的内存索引。
    """
    global _PRODUCT_CATEGORY
    if _PRODUCT_CATEGORY is None:
        from app.services.product_master import load_index

        _PRODUCT_CATEGORY = {
            key: str(value.get("类别") or "") for key, value in load_index().items()
        }
    return _PRODUCT_CATEGORY.get(str(code or "").strip(), "")


def row_product_type(row: dict) -> str:
    """出运产品行的品类：索引里已算好的 product_type 优先，否则按产品编码现算。"""
    return str(row.get("product_type") or "").strip() or product_category(
        row.get("产品编码") or row.get("product_code")
    )


def _dec(value: Any) -> Decimal:
    """宽松转 Decimal（空/千分位/脏文本都吞掉，取不到算 0）。"""
    text = str(value or "").replace(",", "").strip()
    if not text:
        return Decimal(0)
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def fmt_weight(value: Decimal) -> str:
    """重量输出：保留 2 位小数，去掉无意义的尾零。"""
    text = f"{value:.2f}"
    return text.rstrip("0").rstrip(".") if "." in text else text


_GRN_WEIGHT: dict[str, Decimal] = {}


def grn_weight_for(purchase_code: str) -> Decimal:
    """采购单 → 入库单附件里能核到的重量（kg）；核不到返回 0。

    取 `grn_select.select()` 保留下来的入库单，优先加明细行 `weight`，
    没有行级重量时退回 `settled_totals.weight`。
    """
    code = str(purchase_code or "").strip()
    if not code:
        return Decimal(0)
    if code in _GRN_WEIGHT:
        return _GRN_WEIGHT[code]
    total = Decimal(0)
    try:
        from app.services.grn_extract import cache_path, to_decimal
        from app.services.grn_select import select

        for entry in select(code).get("kept") or []:
            path = cache_path(code, entry.get("file") or "")
            if not path.exists():
                continue
            payload = json.loads(path.read_text(encoding="utf-8"))
            found = Decimal(0)
            for line in payload.get("lines") or []:
                value = to_decimal(line.get("weight"))
                if value:
                    found += value
            if not found:
                value = to_decimal((payload.get("settled_totals") or {}).get("weight"))
                if value:
                    found = value
            total += found
    except Exception:  # noqa: BLE001
        total = Decimal(0)
    _GRN_WEIGHT[code] = total
    return total


# 飞书那套粗分类（data/reference/feishu_product_map.json 的取值集合）
KNOWN_CATEGORIES = {
    "法兰", "管件", "无缝管", "镍基无缝管", "焊管", "焊材", "板棒", "盘管", "三角丝", "其他",
}


def coarse_label(*candidates: Any) -> str:
    """把若干候选名归到同一套粗分类：已是粗分类 → 直接用；否则按映射表折算；
    都折算不出（如「不锈钢线」）时再用报关品名折算；最后才保留原值。
    """
    for value in candidates:
        text = str(value or "").strip()
        if not text:
            continue
        if text in KNOWN_CATEGORIES:
            return text
        mapped = decl_category(text)
        if mapped:
            return mapped
    for value in candidates:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def allocate_weight(
    total: Decimal,
    purchase_codes: list[str],
    amount_cents: list[int],
) -> tuple[list[Decimal], str]:
    """把一行报关行重量摊到它拆出的各条记录上（2026-09-28 用户口径）。

    1. 单条记录 → 整笔给它（重量没被拆）；
    2. 多条记录、且各采购单的入库单附件都能核到重量 → 按附件重量占比摊；
    3. 否则 → 按已算好的报关金额占比摊。
    最后一组吸收分位尾差，保证合计等于该报关行的报关重量（不漏）。
    """
    count = len(purchase_codes)
    if count == 0 or total <= 0:
        return [Decimal(0)] * count, ""
    if count == 1:
        return [total], "单条记录，报关重量直接回填"

    grn = [grn_weight_for(code) for code in purchase_codes]
    if all(value > 0 for value in grn):
        basis = grn
        source = f"按入库单附件重量占比分摊（附件重量合计 {fmt_weight(sum(grn, Decimal(0)))}kg）"
    else:
        basis = [Decimal(cents) for cents in amount_cents]
        source = "按报关金额占比分摊（未取得入库单附件重量）"
    base = sum(basis, Decimal(0))
    if base <= 0:
        basis = [Decimal(1)] * count
        base = Decimal(count)

    parts: list[Decimal] = []
    used = Decimal(0)
    for position, value in enumerate(basis):
        if position == count - 1:
            part = total - used
        else:
            part = (total * value / base).quantize(Decimal("0.01"))
            used += part
        parts.append(part)
    return parts, source
# 报关金额 ≈ 出运单出运总金额 的容差（分）：实测只差 1 分（26MT-03P074Y-B
# 报关 48840.75 vs 出运单 48840.74），留一点点余量但不至于认错单
AMOUNT_MATCH_CENTS = 5

_ADD_SUFFIX_RE = re.compile(r"-?ADD\d*", re.IGNORECASE)


def order_family(code: Any) -> str:
    """订单家族键：订单核心去掉**年份前缀**与 **-ADDn 尾缀**。

    「按报关金额认领出运单」只在这个家族内认：
    实测需要认的情况是同一张订单的不同写法（合同 `25MT-03P495Y-ADD1-A` 对应出运单
    `25MT-03P495Y-B`：年份 25MT/26MT、批次字母、ADD 都可能写得不一致）；
    绝不允许跨订单认——`26MT-03P315B`（400 美元）曾认到毫不相关的
    `25MT-03T614`（出运总金额恰巧也是 400 美元），成本被算成 ¥751.20。
    """
    core = order_core(code) or strip_pi(code)
    if core[:2].isdigit():
        core = core[2:]
    return _ADD_SUFFIX_RE.sub("", core)


def shipment_totals() -> dict[str, int]:
    """出运发票号 → 出运总金额（分）。

    用于「按报关金额认领出运单」：报关单金额通常就等于某一张出运单的出运总金额，
    这比按单号猜更硬——这批单子的报关合同号与出运单号在批次字母上对不上
    （合同 `25MT-03P495Y-ADD1-A` 实际对应出运单 `25MT-03P495Y-B`）。
    明细文件缺失时退回出运单列表里的 `totalAmount`。
    """
    global _SHIPMENT_TOTALS
    if _SHIPMENT_TOTALS is None:
        from app.services.erp_cache import CACHE_ROOT, read_jsonl

        table: dict[str, int] = {}
        detail_dir = CACHE_ROOT / "details" / "shipments"
        if detail_dir.exists():
            for path in detail_dir.glob("*.json"):
                try:
                    payload = json.loads(path.read_text(encoding="utf-8"))
                except ValueError:
                    continue
                base = payload.get("baseInfo") or {}
                invoice = str(
                    payload.get("invoiceCode") or base.get("外销发票号") or ""
                ).strip()
                amount = to_cents(base.get("出运总金额"))
                if invoice and amount:
                    table.setdefault(invoice, amount)
                    table.setdefault(strip_pi(invoice), amount)
        for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
            invoice = str(row.get("invoiceCode") or "").strip()
            amount = to_cents(row.get("totalAmount"))
            if invoice and amount:
                table.setdefault(invoice, amount)
                table.setdefault(strip_pi(invoice), amount)
        _SHIPMENT_TOTALS = table
    return _SHIPMENT_TOTALS


def shipment_lines_by_contract() -> dict[str, list[dict]]:
    """合同基号 -> 去重后的出运产品行。"""
    out: dict[str, list[dict]] = defaultdict(list)
    by_invoice: dict[str, list[dict]] = defaultdict(list)
    seen: set[tuple] = set()
    for line in load_lines():
        # 出运产品行本身不带供应商名称时，用采购订单号回查原生供应商名
        # （否则「按供应商打包」会把所有缺名行挤进同一个筐里）
        if not str(line.get("supplier") or "").strip():
            filled = purchase_supplier_name(
                {"采购订单号": line.get("purchase_code")}
            ) or supplier_by_code().get(str(line.get("supplier_code") or "").strip(), "")
            if filled:
                line = {**line, "supplier": filled}
        if not str(line.get("product_type") or "").strip():
            kind = row_product_type(line)
            if kind:
                line = {**line, "product_type": kind}
        keys = {strip_pi(c) for c in (line.get("order_codes") or [])}
        keys.add(strip_pi(line.get("invoice_code")))
        keys.add(strip_pi(line.get("purchase_code")))
        invoice = strip_pi(line.get("invoice_code"))
        if invoice:
            by_invoice[invoice].append(line)
        marker = (
            strip_pi(line.get("invoice_code")),
            line.get("purchase_code"),
            line.get("sku"),
            line.get("amount_usd"),
            line.get("quantity"),
        )
        if marker in seen:
            continue
        seen.add(marker)
        for key in keys:
            if key:
                out[key].append(line)
                canonical_key = canonical(key)
                if canonical_key != key:
                    out[canonical_key].append(line)
    return out, by_invoice


def attempt_level(
    decls: dict[str, list[dict]],
    items: list,
    fees: dict[str, int],
    freight: int,
    tolerance: int,
    directions: dict[str, str] | None = None,
) -> dict:
    """在给定粒度下试遍费用口径，返回最优结果与诊断信息。

    口径写法 `sub:equal` 表示「报关金额里已经含了这笔费用，要扣掉」；
    `add:equal` 表示「出运金额 = 报关金额 + 这笔费用」（如客户赔货金额）。
    """
    import itertools

    SUBSET_MAX_DECLS = 5

    def options_for(name: str, dirs: tuple[str, ...]) -> list[str]:
        opts = [f"{direction}:{mode}" for direction in dirs for mode in FREIGHT_MODES]
        # 单笔费用还额外试「只在某个报关单子集里分摊」——业务上常见「运费只摊给其中几张
        # 报关单、剩下的为 0」（如 26MT-05G019A：2308 只在两张焊管单里平分）。
        if len(fees) == 1 and len(decls) <= SUBSET_MAX_DECLS:
            names = list(decls)
            for size in range(1, len(names)):
                for combo in itertools.combinations(names, size):
                    for direction in dirs:
                        for mode in FREIGHT_MODES:
                            opts.append(f"{direction}:{mode}@{','.join(combo)}")
        return opts

    def combos_for(
        directions: tuple[str, ...],
        per_fee: dict[str, tuple[str, ...]] | None = None,
    ) -> list[dict[str, str]]:
        if not fees:
            return [{}]

        def dirs_of(name: str) -> tuple[str, ...]:
            if per_fee and name in per_fee:
                return per_fee[name]
            return directions

        if len(fees) > 3:
            return [
                {name: value} for name in fees for value in options_for(name, dirs_of(name))
            ]
        return [
            dict(zip(fees, values))
            for values in itertools.product(*(options_for(name, dirs_of(name)) for name in fees))
        ]

    def fee_shares(combo: dict[str, str]) -> dict[str, int]:
        """费用名 -> 每张报关单要扣掉的金额；`add` 口径记成负数（等于加回去）。"""
        shares = {decl: 0 for decl in decls}
        for name, value in combo.items():
            direction, _, mode = str(value).partition(":")
            if not mode:
                direction, mode = "sub", direction
            if "@" in mode:
                mode, _, subset_text = mode.partition("@")
                members = [token for token in subset_text.split(",") if token in decls]
                scope = {decl: decls[decl] for decl in members} or decls
            else:
                scope = decls
            part = allocate_fee(scope, fees.get(name, 0), mode)
            for decl, value_part in part.items():
                shares[decl] += value_part if direction != "add" else -value_part
        return shares

    def evaluate(combos: list[dict[str, str]], allow_fraction: bool = False):
        best = None
        combo_results: list[tuple[tuple[int, int], dict, object]] = []
        combo_scores: dict[str, int] = {}
        seen_shares: dict[tuple, str] = {}
        for combo in combos:
            shares = fee_shares(combo) if fees else {d: 0 for d in decls}
            label = "、".join(f"{k}:{v}" for k, v in combo.items()) or "(无费用)"
            signature = tuple(sorted(shares.items()))
            if signature in seen_shares:
                # 单张报关单时 equal / by_weight / by_amount / none 的分摊额可能完全相同，
                # 等价口径不重复求解
                combo_scores[label] = combo_scores.get(seen_shares[signature], 0)
                continue
            seen_shares[signature] = label
            attempt = solve(
                decls,
                # 每个口径都用**产品行的副本**求解：部分数量分摊会改写行对象
                # （数量/金额），共用同一批对象会让后面的口径拿到被改坏的数据。
                [replace(line) for line in items],
                freight,
                use_heuristics=True,
                freight_shares=shares,
                tolerance=tolerance,
                allow_fraction=allow_fraction,
            )
            exacts: dict[str, bool] = {}
            for decl, rows in decls.items():
                used = attempt.assignments.get(decl) or []
                capacity = sum(to_cents(r.get("amount")) for r in rows) - shares.get(decl, 0)
                exacts[decl] = bool(used) and abs(
                    sum(x.amount for x in used) - capacity
                ) <= tolerance
            exact = sum(1 for v in exacts.values() if v)
            covered = sum(1 for decl in decls if attempt.assignments.get(decl))
            combo_scores[label] = exact
            combo_results.append(((exact, covered), combo, attempt))
            if best is None or (exact, covered) > best[0]:
                best = ((exact, covered), exact, combo, attempt, shares, exacts)
        return best, combo_results, combo_scores

    # 第一轮：按已知的「加减方向」定死（ERP 的 `加减` 字段 MCP 取不到值，
    # 方向是按已闭合单据统计出来的：检验费/手续费/木箱费…→sub，罚金/FOB退费→add）。
    # 方向定死能砍掉一半组合，显著减少"费用口径不唯一"的假多解。
    # 方向只认 ERP 恒等式（出运总金额 − 产品合计 = Σ±费用）解出的唯一解。
    # 不用「按费用名称的经验表」定方向：费用命名不规范时它可能误判，
    # 而恒等式是纯金额算术、与名称无关；解不出来就两个方向都枚举、按多解处理。
    resolved_dirs = directions or {}
    known_dirs = {name: (resolved_dirs[name],) for name in fees if name in resolved_dirs}
    best, combo_results, combo_scores = evaluate(combos_for(("sub",), known_dirs or None))
    if best[1] != len(decls):
        # 定死方向解不齐 → 放开两个方向重来（保留两种结果一起比较）
        best_full, results_full, scores_full = evaluate(combos_for(("sub",)))
        combo_results = combo_results + results_full
        combo_scores.update(scores_full)
        if best_full[0] > best[0]:
            best = best_full
    # 「费用加在报关金额之外」是少数情况（如客户赔货金额），只在常规口径解不齐时启用，
    # 而且一次只把一笔费用设成 add，避免口径组合爆炸
    if best[1] != len(decls) and fees:
        names = list(fees)
        add_combos: list[dict[str, str]] = []
        for target_name in names:
            others = [name for name in names if name != target_name]
            if len(others) <= 2:
                other_options = [f"sub:{mode}" for mode in FREIGHT_MODES]
                combinations: list[tuple] = list(
                    itertools.product(other_options, repeat=len(others))
                )
            else:
                combinations = [tuple("sub:equal" for _ in others)]
            for values in combinations:
                for mode in ("equal", "by_weight", "by_amount"):
                    combo = dict(zip(others, values))
                    combo[target_name] = f"add:{mode}"
                    add_combos.append(combo)
        best2, results2, scores2 = evaluate(add_combos)
        combo_results = combo_results + results2
        combo_scores.update(scores2)
        if best2[0] > best[0]:
            best = best2

    assert best is not None
    if best[1] != len(decls):
        # 整数口径怎么都解不齐 → 才放开「可分数单位」（米/千克这类）按小数配平。
        # 必须放在最后：分数解天然能吸收任意余量，与整数解并列时会把本来
        # 拆得出来的单子判成「多解」（26MT-05G019A 就这么被误判过）。
        retry_combos = [combo for _score, combo, _attempt in combo_results]
        best_frac, results_frac, scores_frac = evaluate(
            retry_combos, allow_fraction=True
        )
        combo_results = combo_results + results_frac
        combo_scores.update({f"{k}（分数单位）": v for k, v in scores_frac.items()})
        if best_frac[0] > best[0]:
            best = best_frac
    _, exact_best, combo_best, result, shares_best, exacts_best = best
    # 多口径：另一种费用口径也能得到同样精确且指派不同的解 → 不唯一
    signatures = set()
    for score, _combo, attempt in combo_results:
        if score[0] != exact_best or exact_best <= 0:
            continue
        # 用「供应商 + 采购单 + 金额」判等价：不同口径挑到不同产品行、
        # 但只要最终落给每张报关单的工厂/采购单/金额一样，拆单结果就是一样的
        signatures.add(
            tuple(
                sorted(
                    (
                        decl,
                        tuple(
                            sorted(
                                (line.supplier, line.purchase_code, line.amount)
                                for line in expand(items, lines or [])
                            )
                        ),
                    )
                    for decl, lines in attempt.assignments.items()
                )
            )
        )
    return {
        "result": result,
        "shares": shares_best,
        "exacts": exacts_best,
        "exact_count": exact_best,
        "combo": combo_best,
        "combo_scores": combo_scores,
        "mode_ambiguous": len(signatures) > 1,
        "all_exact": exact_best == len(decls),
    }


def expand(items: list, assignment: list) -> list:
    """把「组合记录」还原成真实产品行。"""
    return flatten(assignment)


def flatten(records: list) -> list:
    """把「组合记录」摊平成真实产品行（组合记录自带 members）。"""
    out: list = []
    for record in records:
        out.extend(record.members or [record])
    return out


def _norm_cell(value: Any) -> str:
    """单元格归一：文本压空白、数字按数值（`10` 与 `10.000000` 同一条）。"""
    text = re.sub(r"\s+", " ", str(value or "").strip()).replace(",", "")
    if not text:
        return ""
    try:
        return f"{float(text):.6f}"
    except ValueError:
        return text.upper()


def row_key(raw: dict) -> tuple:
    """一条出运产品行的稳定身份（内容键）。

    不能用 `id(raw)`：`load_lines()` 每读一次 JSON 就新建一批 dict，
    **同一行在不同合同 / 不同候选池里是两个不同的对象**，按 id 比会把
    「真实同一行」判成两条。文本去空白、数字按数值归一，和去重标记同一套口径。
    """
    return (
        _norm_cell(raw.get("invoice_code")),
        _norm_cell(raw.get("purchase_code")),
        _norm_cell(raw.get("sku")),
        _norm_cell(raw.get("amount_usd")),
        _norm_cell(raw.get("quantity")),
        _norm_cell(raw.get("amount_rmb")),
        _norm_cell(raw.get("customs_name")),
        tuple(_norm_cell(code) for code in (raw.get("order_codes") or [])),
    )


def twin_key(line) -> tuple:
    """「同单孪生行」的判据：同一张出运单 + 同一采购单 + 同品名 + 同金额 + 同数量。

    26MT-05X246A 的两条产品行就是这样：`26MT-05X246-1` / `26MT-05X246-2`，
    金额、规格、数量全一样，**只有 SKU 不同**，报关单的金额只闭合得了其中一条。
    """
    raw = line.raw or {}
    return (
        _norm_cell(raw.get("invoice_code")),
        str(line.purchase_code or ""),
        str(line.customs_name or ""),
        line.amount,
        str(line.quantity),
    )


def hs_compatible(left: str, right: str) -> bool:
    """海关编码比较：睿贝侧偶尔少一位（730441000 vs 7304419000），按前 6 位比对。"""
    a = "".join(ch for ch in str(left or "") if ch.isdigit())
    b = "".join(ch for ch in str(right or "") if ch.isdigit())
    if not a or not b:
        return False
    head = min(len(a), len(b), 6)
    return a[:head] == b[:head]


def belongs_to_orders(cores: set[str], orders: set[str]) -> bool:
    """该产品行是否属于这些订单（含同订单的 ADD 单）。"""
    for core in cores:
        if core in orders:
            return True
        if any(core.startswith(order + "-ADD") for order in orders):
            return True
    return False


def directed_split_rows(
    contract: str,
    decls: dict[str, list[dict]],
    lines_raw: list[dict],
    names: SupplierNames,
    *,
    deviance: float = 0.2,
    only_decls: set[str] | None = None,
    debug: bool = False,
) -> tuple[list[dict], str]:
    """最后兜底：按「合同订单 + 海关编码/品名」直接落组。

    针对「一张出运单合并了多个外销单、报关单又按外销单拆开」的情况
    （如出运单 26MT-10E185B&228A ↔ 报关合同 26MT-10E185B / 26MT-10E228A /
    26MT-10E185B&228A）。这类单据金额常有几位数的人工调整，凑不出精确解，
    但**订单 + 品名**的对应关系是唯一的：把产品行按订单过滤后，
    每一报关行只认自己品名/海关编码相符的那一家工厂。

    只有「每行都能认到货、且没有一家工厂同时匹配两行」时才采纳。
    """
    orders = cores_of(contract)
    if not orders:
        return [], "合同订单解析失败"

    def belongs_to_contract(cores: set[str]) -> bool:
        """该产品行是否属于本合同的订单。

        除了合同自身的订单，还要认**同订单的 ADD 单**：例如合同 `26MT-01P242Y-C`
        对应的出运单里还挂着 `26MT-01P242Y-ADD` 的货（金额不够时业务会追加），
        只按订单核心做交集会把这些行丢掉。
        """
        for core in cores:
            if core in orders:
                return True
            if any(core.startswith(order + "-ADD") for order in orders):
                return True
        return False

    pool = [
        line for line in iter_lines(lines_raw) if belongs_to_contract(line_cores(line.raw))
    ]
    if not pool:
        return [], "按合同订单过滤后没有产品行"
    # **同一张出运单下的其它订单产品行也要纳入**：一张出运单常常合并了多个外销单
    # （如出运单 `PI-26MT-05G062&26MT-05X137`），而报关单的合同号只写了其中一个
    # （`26MT-05G062`）。只按合同号过滤会把 `26MT-05X137` 的货整行丢掉，
    # 那份采购单的成本就永远落不到任何一条记录上。
    invoices = {
        str(line.raw.get("invoice_code") or "").strip()
        for line in pool
        if str(line.raw.get("invoice_code") or "").strip()
    }
    if invoices:
        def _marker(item) -> tuple:
            raw = getattr(item, "raw", None) or {}
            return (
                str(getattr(item, "purchase_code", "") or "").strip().upper(),
                re.sub(r"\s+", " ", str(raw.get("sku") or "").strip()).upper(),
                str(raw.get("amount_usd") or "").strip(),
                str(raw.get("quantity") or "").strip(),
            )

        seen_markers = {_marker(line) for line in pool}
        for line in iter_lines(lines_raw):
            marker = _marker(line)
            if marker in seen_markers:
                continue
            if str(line.raw.get("invoice_code") or "").strip() in invoices:
                pool.append(line)
                seen_markers.add(marker)
    groups: dict[str, list] = {}
    for line in pool:
        key = names.short(line.supplier) or line.purchase_code or "(未知供应商)"
        groups.setdefault(key, []).append(line)

    def emit(decl: str, row: dict, key: str, members: list, total: int) -> dict:
        declared_cents = to_cents(row.get("amount"))
        declared_weight = _dec(row.get("weight"))
        parts, basis = allocate_weight(declared_weight, [members[0].purchase_code], [declared_cents])
        return {
            "报关单号": decl,
            "合同号_1": contract,
            "报关品名": str(row.get("name") or ""),
            "海关编码": str(row.get("hs") or ""),
            "产品类型": coarse_label(line_product_type(members[0]), row.get("name")),
            "供应商简称": key,
            "供应商": members[0].supplier,
            "采购单号": "、".join(
                sorted({line.purchase_code for line in members if line.purchase_code})
            ),
            "出运金额合计": f"{total / 100:.2f}",
            # 出运采购金额(RMB) 是成本匹配的分摊权重，必须和出运金额一起落下来，
            # 否则成本匹配拿不到人民币口径，只能退回按出运金额(USD) 占比摊整单
            # （实测 26MT-03T203Y-HX 就是这样把两张报关单摊差 5,111.09 的）。
            "出运采购金额合计": f"{sum(line_rmb_cents(line) for line in members) / 100:.2f}",
            "报关金额": f"{declared_cents / 100:.2f}",
            "客户费用分摊": f"{(declared_cents - total) / 100:.2f}",
            "报关重量": fmt_weight(parts[0]),
            "重量分摊依据": basis,
            "产品行数": len(members),
        }

    def emit_grouped(decl: str, row: dict, key: str, members: list, total: int) -> list[dict]:
        """按「采购订单 + 产品类型」拆成多条记录（同一工厂两种产品类型要分开）。

        多条时报关金额按各组货值比例分（末组吸收尾差）；
        只有一条时保持原样——报关金额取整张单，差额全部作为客户费用分摊，
        这样「整包 + 费用 = 报关金额」的关系仍一眼可见。
        """
        parts = split_by_item(members)
        if len(parts) <= 1:
            return [emit(decl, row, key, members, total)]
        declared_cents = to_cents(row.get("amount"))
        declared_weight = _dec(row.get("weight"))
        out: list[dict] = []
        allocated_all = 0
        allocated_list: list[int] = []
        for position, (_item_key, items) in enumerate(parts):
            share = sum(line.amount for line in items)
            if position == len(parts) - 1:
                allocated_cents = declared_cents - allocated_all
            else:
                allocated_cents = int(
                    round(declared_cents * share / total) if total else share
                )
                allocated_all += allocated_cents
            allocated_list.append(allocated_cents)
        weight_parts, weight_basis = allocate_weight(
            declared_weight, [items[0].purchase_code for _k, items in parts], allocated_list
        )
        for position, (_item_key, items) in enumerate(parts):
            share = sum(line.amount for line in items)
            allocated_cents = allocated_list[position]
            out.append(
                {
                    "报关单号": decl,
                    "合同号_1": contract,
                    "报关品名": str(row.get("name") or ""),
                    "海关编码": str(row.get("hs") or ""),
                    "产品类型": coarse_label(line_product_type(items[0]), row.get("name")),
                    "供应商简称": key,
                    "供应商": items[0].supplier,
                    "采购单号": items[0].purchase_code,
                    "出运金额合计": f"{share / 100:.2f}",
                    "出运采购金额合计": (
                        f"{sum(line_rmb_cents(line) for line in items) / 100:.2f}"
                    ),
                    "报关金额": f"{allocated_cents / 100:.2f}",
                    "客户费用分摊": f"{(allocated_cents - share) / 100:.2f}",
                    "报关重量": fmt_weight(weight_parts[position]),
                    "重量分摊依据": weight_basis,
                    "产品行数": len(items),
                }
            )
        return out

    # 展平所有报关行；只处理「金额法没落上货」的那些报关单
    flat = [
        (decl, row)
        for decl, decl_rows in decls.items()
        if only_decls is None or decl in only_decls
        for row in decl_rows
    ]
    if not flat:
        return [], "没有待落组的报关行"
    # 每个报关行认一家工厂：命中多家时，只允许「金额贴合」把它唯一区分开
    candidates: dict[int, list[str]] = {}
    for key, members in groups.items():
        for position, (_decl, row) in enumerate(flat):
            if _group_matches(members, row):
                candidates.setdefault(position, []).append(key)
    resolved: dict[int, str] = {}
    for position in range(len(flat)):
        decl, row = flat[position]
        keys = candidates.get(position) or []
        if not keys:
            return [], f"{decl} {row.get('name')} 没有品名/海关编码相符的产品行"
        if len(keys) == 1:
            resolved[position] = keys[0]
            continue
        target = to_cents(row.get("amount"))
        within = [
            key
            for key in keys
            if target
            and abs(sum(line.amount for line in groups[key]) - target) / target <= deviance
        ]
        if len(within) != 1:
            return [], (
                f"{decl} {row.get('name')} 命中 {len(keys)} 家工厂"
                f"（{'、'.join(sorted(keys))}）且金额无法唯一区分，不硬猜"
            )
        resolved[position] = within[0]

    rows: list[dict] = []
    if debug:
        print(
            "[directed-groups] "
            + " | ".join(
                f"{key}={sum(line.amount for line in members) / 100:.2f}({len(members)}行)"
                for key, members in groups.items()
            ),
            flush=True,
        )
    for key, members in groups.items():
        total = sum(line.amount for line in members)
        hit_rows = [position for position, chosen in resolved.items() if chosen == key]
        if not hit_rows:
            continue
        if len(hit_rows) == 1:
            decl, row = flat[hit_rows[0]]
            target = to_cents(row.get("amount"))
            # 这组货明显装不下这一张报关单（>2 倍）→ 说明该组还对应别的报关单，不硬塞。
            # 「小于 1 个货币单位」的零头不算"装不下"。
            if target and total > target + IGNORABLE_DIFF_CENTS and total > target * 2:
                continue
            rows.extend(emit_grouped(decl, row, key, members, total))
            continue
        # 一家工厂的货对应多张报关单的同一品名（如 26MT-01P242Y-C 的两张无缝管）：
        # 只有合计对得上才拆开，按金额贴合度分配
        targets = [to_cents(flat[position][1].get("amount")) for position in hit_rows]
        if not all(targets):
            return [], f"{key} 命中多行但报关金额缺失"
        summary = sum(targets)
        # 「小于 1 个货币单位的补充差额忽略不计」（2026-09-28 用户口径）：
        # 实测 26MT-01P242Y-C 的两张无缝管报关行合计 71,293.68、出运单这一组的货
        # 71,293.67，只差 0.01 USD —— 这种零头不该挡住拆分。
        if summary <= 0 or (
            abs(total - summary) > IGNORABLE_DIFF_CENTS
            and abs(total - summary) / summary > 0.02
        ):
            return [], (
                f"{key} 命中多行（{len(hit_rows)}）但金额合计对不上"
                f"（{total / 100:.2f} vs {summary / 100:.2f}）"
            )
        filled = [0] * len(hit_rows)
        picked: list[list] = [[] for _ in hit_rows]
        for line in sorted(members, key=lambda item: -item.amount):
            slot = min(
                range(len(hit_rows)),
                key=lambda i: abs(targets[i] - (filled[i] + line.amount)),
            )
            picked[slot].append(line)
            filled[slot] += line.amount
        for slot, position in enumerate(hit_rows):
            decl, row = flat[position]
            if not picked[slot]:
                return [], f"{key} 拆给 {decl} {row.get('name')} 时没有产品行"
            rows.extend(emit_grouped(decl, row, key, picked[slot], filled[slot]))
    if not rows:
        return [], "没有被品名/海关编码唯一认领的产品行"
    return rows, ""


def split_by_item(members: list) -> list[tuple[tuple[str, str], list]]:
    """把一组产品行按「采购订单号 + 产品类型」再拆细。

    拆单口径：**采购订单 / 供应商 / 产品类型** 任一不同，都要拆成独立记录
    （供应商在上一层已经分好，这里再按采购单与产品类型分组）。
    同一家工厂既有法兰又有管件时（勇恒就有这两种），必须出两条记录，
    否则一条记录里混着两种产品类型，成本也就没法按产品类型对账。
    返回 [(采购单号, 产品类型), 该组产品行]，按首次出现顺序。
    """
    buckets: dict[tuple[str, str], list] = {}
    for line in members:
        key = (
            str(getattr(line, "purchase_code", "") or ""),
            # 品类键：先用 ERP 出运品名（求解器一直在用的细分口径，保住已拆对的结果）；
            # 出运品名缺失时用商品资料粗分类补位（本任务要修的缺口）
            str(
                getattr(line, "customs_name", "")
                or line_product_type(line)
                or ""
            ).strip(),
        )
        buckets.setdefault(key, []).append(line)
    return list(buckets.items())


def _group_matches(members: list, row: dict) -> bool:
    """一组产品行是否与某个报关行相符（海关编码前 6 位或品名一致）。"""
    want_hs = str(row.get("hs") or "").strip()
    want_name = str(row.get("name") or "").strip()
    if want_hs and any(hs_compatible(line.hs_code, want_hs) for line in members):
        return True
    return bool(want_name) and any(name_matches(line, want_name) for line in members)


def line_rmb_cents(line) -> int:
    """出运产品行的人民币采购金额（分）。

    组合行（按供应商/产品类型打包）取成员之和；被部分出运占用的行按数量折算。
    这个数值是「采购金额分摊」的权重来源：同一采购单拆到多张报关单时，
    按各行实际装了多少货（人民币口径）分摊，比按产品类型汇总更准。
    """
    members = getattr(line, "members", None)
    if members:
        return sum(line_rmb_cents(member) for member in members)
    cents = to_cents((getattr(line, "raw", None) or {}).get("amount_rmb"))
    total_qty = getattr(line, "total_qty", None) or Decimal(0)
    qty = getattr(line, "quantity", None) or Decimal(0)
    if cents and total_qty > 0 and 0 < qty < total_qty:
        cents = int((Decimal(cents) * qty / total_qty).to_integral_value())
    return cents


def run(
    parse_xlsx: Path | None = None,
    out_dir: Path | None = None,
    only: str = "",
    debug: bool = False,
) -> dict:
    """跑一轮拆单。`parse_xlsx` / `out_dir` 可指定；`only` 是逗号分隔的合同过滤。"""
    parse_xlsx = Path(parse_xlsx or PARSE_XLSX)
    out_dir = Path(out_dir or out_dir)
    declared = read_declared(parse_xlsx)
    if only:
        want = [x.strip() for x in only.split(",") if x.strip()]
        declared = [d for d in declared if any(w in d.contract for w in want)]
    names = SupplierNames()

    # 合同 -> 报关单号 -> 报关行（solve 需要 dict 形式：金额/品名/海关编码/供应商）
    by_contract: dict[str, dict[str, list[dict]]] = defaultdict(lambda: defaultdict(list))
    for line in declared:
        by_contract[line.contract][line.declaration_no].append(
            {
                "amount": str(line.amount),
                "name": line.product_name,
                "hs": line.hs_code,
                "supplier": "",
                "weight": str(line.weight),
            }
        )

    pool, by_invoice = shipment_lines_by_contract()

    def collect_lines(
        contract: str,
        *,
        via_pool: bool = False,
        invoices: list[str] | None = None,
    ) -> list[dict]:
        """取该合同的出运产品行，按 (采购单,SKU,金额,数量) 去重。

        首选「三单互查」：合同号 → 出运单（原生 orderCode/invoiceCode）→ 出运单明细，
        只纳入这些出运单自己的产品行；`via_pool=True` 时才放宽到「整张订单下所有
        出运单的产品行」，作为一张报关单跨多张出运单时的兜底。
        """
        merged: list[dict] = []
        seen: set[tuple] = set()

        def marker_key(value) -> str:
            """去重用的归一化。

            数值按小数比较（`10` 与 `10.000000` 视为同一条）；
            文本把**内部空白**统一（MCP 明细的 SKU 带换行 `1smls tube\\n0.25`，
            历史缓存里是空格 `1smls tube 0.25`，不归一就会被当成两条不同产品行）。
            """
            text = re.sub(r"\s+", " ", str(value or "").strip()).replace(",", "")
            if not text:
                return ""
            try:
                return f"{float(text):.6f}"
            except ValueError:
                return text.upper()

        def add(row: dict) -> None:
            marker = tuple(
                marker_key(x) for x in (
                row.get("purchase_code"),
                row.get("sku"),
                row.get("amount_usd"),
                row.get("quantity"),
                )
            )
            if marker in seen:
                return
            seen.add(marker)
            merged.append(row)

        def index_unit(invoice: str, purchase_code: Any, sku: Any) -> str:
            """MCP 明细的「计量单位」常是空的，退回出运产品行索引里同一行的单位。

            单位决定了能不能按小数（米/千克这类）配平，缺了会让这类单子算不出来。
            """
            for candidate in by_invoice.get(strip_pi(invoice), []):
                if (
                    str(candidate.get("purchase_code") or "").strip()
                    == str(purchase_code or "").strip()
                    and marker_key(candidate.get("sku")) == marker_key(sku)
                ):
                    return str(candidate.get("unit") or "").strip()
            return ""

        # 路线一（首选）：用出运单列表按合同号直接定位出运单，缺明细就按需抓
        target_invoices = (
            invoices if invoices is not None else shipment_invoices_for_contract(contract)
        )
        for invoice in target_invoices:
            detail = ensure_shipment_detail(invoice, online=ONLINE_TOPKUP)
            product_rows = list((detail or {}).get("productList") or [])
            if not product_rows:
                # 睿贝 MCP 偶发取不到明细（返回空单据）：退回出运产品行索引
                product_rows = list(by_invoice.get(strip_pi(invoice), []))
                for row in product_rows:
                    add(row)
                continue
            for row in product_rows:
                # 「一格两个订单、两个公司」的合并行要拆开，否则整行只能当一家工厂
                for part in split_merged_parts(
                    {
                        "invoice_code": invoice,
                        "order_codes": [row.get("销售订单号")],
                        "purchase_code": normalise_purchase_code(
                            str(row.get("采购订单号") or "")
                        ),
                        "supplier": row.get("供应商名称")
                        or purchase_supplier_name(row)
                        or supplier_by_code().get(str(row.get("供应商编码") or "").strip(), ""),
                        "supplier_code": row.get("供应商编码"),
                        "amount_usd": row.get("出运金额"),
                        "amount_rmb": row.get("出运采购金额(RMB)"),
                        "hs_code": row.get("海关编码"),
                        "unit_price_usd": row.get("外销单价"),
                        "customs_name": row.get("海关商品（中文）"),
                        "product_code": row.get("产品编码"),
                        "product_type": row_product_type(row),
                        "quantity": row.get("出运数量"),
                        "sku": row.get("SKU"),
                        "unit": str(row.get("计量单位") or "").strip()
                        or index_unit(invoice, row.get("采购订单号"), row.get("SKU")),
                    }
                ):
                    add(part)

        if via_pool:
            keys: list[str] = []
            for order in expand_contract(contract) or [contract]:
                key = strip_pi(order)
                if key:
                    keys.append(key)
                    keys.append(canonical(key))
            for key in keys:
                for row in pool.get(key, []):
                    add(row)
            if not merged:
                for order in expand_contract(contract) or [contract]:
                    base = strip_pi(order)[:11]
                    for row in pool.get(base, []):
                        add(row)
        # 同一张出运单下的其它订单产品行也要纳入（一张报关单可能跨多个订单）
        invoices = {strip_pi(x.get("invoice_code")) for x in merged if x.get("invoice_code")}
        for invoice in invoices:
            for row in by_invoice.get(invoice, []):
                add(row)
        return merged
    split_rows: list[dict] = []
    failures: list[dict] = []
    # 「出运单里有、但没有任何报关单认领」的产品行（以前是静默丢弃）
    pool_rows_all: dict[int, tuple[str, Any]] = {}
    claimed_rows_all: set[int] = set()
    claimed_twins_all: set[tuple] = set()
    extra_rows_all: dict[tuple, str] = {}   # 已按「孪生行」补记成本的产品行 → 合同号
    strategy_log: list[dict] = []
    # 报关单号 -> 该单实际装到的 ERP 产品类别（用于「混装报关」的比对放行）
    decl_products: dict[str, set[str]] = defaultdict(set)

    def collect_by_declared_amount(
        contract: str, decls: dict[str, list[dict]]
    ) -> list[dict]:
        """按「报关金额 = 出运单出运总金额」认领出运单，返回这些出运单的产品行。

        每张报关单都要能**唯一**认领到一张出运单才采纳（有一张认不到就整体放弃，
        交回原有候选池），避免金额巧合把不相干的出运单拉进来。
        """
        totals = shipment_totals()
        if not totals:
            return []
        # 认领范围限定在「同一订单家族」的出运单上（年份/ADD/批次写法不一致时才靠金额认单），
        # 否则金额巧合会把毫不相关的订单拉进来
        families: dict[str, set[str]] = {}
        for ref in shipment_refs():
            if not ref.invoice:
                continue
            families[strip_pi(ref.invoice)] = {order_family(core) for core in ref.cores}
        wanted = {order_family(core) for core in cores_of(contract)} - {""}
        invoices: list[str] = []
        for _decl, rows in decls.items():
            amount = sum(to_cents(row.get("amount")) for row in rows)
            if not amount:
                return []
            hits = [
                invoice
                for invoice, total in totals.items()
                if abs(total - amount) <= AMOUNT_MATCH_CENTS
                and (not wanted or (families.get(strip_pi(invoice), set()) & wanted))
            ]
            if len(hits) != 1:
                return []
            if hits[0] not in invoices:
                invoices.append(hits[0])
        return collect_lines(contract, invoices=invoices) if invoices else []

    for contract, decls in sorted(by_contract.items()):
        import time as _time

        _started = _time.perf_counter()
        # 候选产品行池：先用「三单互查」直取的出运单（窄），
        # 窄集解不出来时再放宽到整张订单的全部出运产品行（宽）兜底
        # 两种解析并列做候选：映射表（拆分结果→原值，更精确）+ 覆盖/基号兜底（更宽），
        # 哪个能拆得更全就用哪个
        map_invoices = shipment_invoices_for_contract(contract)
        legacy_invoices = shipment_invoices_for_contract(contract, use_map=False)
        map_lines = collect_lines(contract, invoices=map_invoices)
        legacy_lines = collect_lines(contract, invoices=legacy_invoices)
        wide = collect_lines(contract, via_pool=True)
        # 「一张报关单 ↔ 一张出运单，金额相等」：报关合同号与出运单号对不上
        # （批次字母不同、年份写错、或出运单没写 ADD）时，按金额认领最硬。
        # 作为**并列候选**参与择优；两边都没其它数据时它就是唯一的池子。
        amount_lines = collect_by_declared_amount(contract, decls)
        # 报关单点明了批次（`26MT-03R036F`），但睿贝里没有这批出运单、金额也认领不到出运单时，
        # 不能再拿同订单的其它批次顶替：那是另一批货，成本会算成别人的
        # （实测 26MT-03R036F 被 A–E 五批 279 行顶替，56 美元的报关行凑出 325.27 元成本）。
        # 直接报异常，让财务去核对合同号或等出运单生成。
        if names_batch(contract) and not (map_invoices or legacy_invoices or amount_lines):
            siblings = sorted(
                {
                    row.invoice
                    for row in container_shipments(contract)
                    if row.invoice
                }
            )
            core = order_core(contract) or strip_pi(contract)
            batches = []
            for invoice in siblings:
                text = strip_pi(invoice)
                suffix = text[len(core) :].strip("-_ ") if text.startswith(core) else ""
                batches.append(suffix or invoice)
            note = f"睿贝里没有这张出运单（{contract}）"
            if batches:
                note += "，该订单只有 " + "/".join(batches) + " 批"
            for decl, rows in decls.items():
                for row in rows:
                    failures.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            "报关品名": row.get("name"),
                            "比对结果": note,
                        }
                    )
            continue
        narrow = map_lines or legacy_lines or amount_lines or wide
        candidates: list[tuple[str, list[dict]]] = []

        def invoice_signature(pool: list[dict]) -> tuple:
            return tuple(
                sorted({str(x.get("invoice_code") or "") for x in pool if x.get("invoice_code")})
            )
        # 按合同订单过滤的候选池：只保留确实属于本合同订单（含 ADD 单）的产品行。
        # 一个订单常常分批出运（26MT-06M087 就分散在 5 张出运单里），
        # 混在一起解会被无关货物和费用口径搅乱，过滤后能精确对上的部分先拆出来。
        contract_orders = cores_of(contract)
        filtered_lines = [
            row
            for row in (map_lines or legacy_lines or wide)
            if belongs_to_orders(line_cores(row), contract_orders)
        ]
        # 注意：过滤池只作**最后兜底**（见下面 best_choice 为空的处理），
        # 不参与正常择优——否则会改变其它合同已经拆对的结果。
        if map_lines:
            candidates.append(("映射表直取", map_lines))
        if amount_lines and invoice_signature(amount_lines) not in {
            invoice_signature(map_lines),
            invoice_signature(legacy_lines),
        }:
            candidates.insert(0, ("按报关金额认领出运单", amount_lines))
        if legacy_lines and legacy_invoices != map_invoices:
            candidates.append(("出运单直取（覆盖/基号兜底）", legacy_lines))
        if not candidates or len(wide) > len(narrow):
            candidates.append(("订单全量兜底", wide))
        if debug:
            print(
                f"[debug] {contract} 出运产品行 窄={len(narrow)} 宽={len(wide)} "
                f"出运单={sorted({str(x.get('invoice_code') or '') for x in narrow})}",
                flush=True,
            )
        if not narrow and not wide:
            for decl, rows in decls.items():
                for row in rows:
                    failures.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            "报关品名": row.get("name"),
                            "比对结果": "无出运产品行",
                        }
                    )
            continue

        best_choice = None
        for scope_label, lines_raw in candidates:
            invoices = sorted(
                {str(x.get("invoice_code") or "") for x in lines_raw if x.get("invoice_code")}
            )
            kinds = expense_kinds(invoices)
            fees = {name: int(round(amount * 100)) for name, amount in kinds.items() if amount}
            freight = expense_cents(invoices)
            # 出现「极小客户费用」（业务手工补差）时，报关金额与出运金额的关系是精确的，
            # 禁用容差、只认完全一致的解，避免容差把一堆近似子集都算成候选造成假多解
            tiny_fee = any(
                0 < abs(amount) <= TINY_FEE_YUAN + 1e-9 for amount in kinds.values()
            )
            tolerance = 0 if tiny_fee else TOLERANCE_CENTS
            item_lines = iter_lines(lines_raw)

            # —— 分层求解：先用「整厂货物」等粗粒度，唯一命中即采纳；否则退到更细粒度 ——
            chosen = None
            attempts: dict[str, dict] = {}
            for level_name, key_fn in LEVELS:
                if key_fn is None:
                    items = item_lines
                else:
                    items = bundle_lines(
                        item_lines, lambda line, k=key_fn: k(line, names), label=level_name
                    )
                info = attempt_level(
                    decls, items, fees, freight, tolerance, fee_directions(invoices)
                )
                attempts[level_name] = info
                if debug:
                    print(
                        f"[level] {contract} {level_name} 精确={info['exact_count']}/{len(decls)} "
                        f"多解={info['mode_ambiguous']} 口径={info['combo']}",
                        flush=True,
                    )
                # 粗粒度要求：全部报关单精确闭合，且费用口径不冲突
                if info["all_exact"] and not info["mode_ambiguous"]:
                    chosen = (level_name, info, items)
                    break
            if chosen is None:
                # 最后一招：不猜费用口径，改用「容量=报关金额、允许残差」，
                # 以整张采购单为单位把产品行分给各张报关单（残差即未分摊的客户费用）
                residual_items = bundle_lines(
                    item_lines,
                    lambda line: line.purchase_code
                    or names.short(line.supplier)
                    or "(未知供应商)",
                    label="按采购单分组",
                )
                residual_result, residual_exact = solve_residual(
                    decls, residual_items, tolerance
                )
                if residual_result.assignments and not any(
                    "多解" in note for note in residual_result.notes
                ):
                    info_res = {
                        "result": residual_result,
                        "shares": {decl: 0 for decl in decls},
                        "exacts": {},
                        "exact_count": residual_exact,
                        "combo": {"(按报关金额残余)": "residual"},
                        "combo_scores": {"residual": residual_exact},
                        "mode_ambiguous": False,
                        "all_exact": residual_exact == len(decls),
                    }
                    chosen = ("按采购单分组（金额残余）", info_res, residual_items)
            if chosen is None:
                level_name, info = LEVELS[-1][0], attempts[LEVELS[-1][0]]
                chosen = (level_name, info, item_lines)

            chosen_level, info, items = chosen
            score = (
                1 if (info["all_exact"] and not info["mode_ambiguous"]) else 0,
                0 if info["mode_ambiguous"] else 1,
                int(info["exact_count"]),
                sum(1 for decl in decls if info["result"].assignments.get(decl)),
            )
            if debug:
                print(
                    f"[cand] {contract} {scope_label} 产品行={len(lines_raw)} "
                    f"出运单={len(invoices)} 粒度={chosen_level} 精确={info['exact_count']}/{len(decls)} "
                    f"多解={info['mode_ambiguous']} 有货报关单="
                    f"{sum(1 for d in decls if info['result'].assignments.get(d))}/{len(decls)}",
                    flush=True,
                )
            if best_choice is None or score > best_choice[0]:
                best_choice = (score, scope_label, invoices, chosen_level, info, items)
            if score[0]:
                break

        assert best_choice is not None
        _best_score, scope_label, invoices, chosen_level, info, items = best_choice
        # 最后兜底：所有常规候选池一行都没产出时，改用「按合同订单过滤」的池子
        # （一个订单分批出运、且被无关货物/费用口径搅乱时的场景，如 26MT-06M087）
        if (
            (
                not any(info["result"].assignments.get(decl) for decl in decls)
                or info["mode_ambiguous"]
            )
            and filtered_lines
            and len(filtered_lines) < len(narrow)
        ):
            filtered_invoices = sorted(
                {str(row.get("invoice_code") or "") for row in filtered_lines if row.get("invoice_code")}
            )
            filtered_kinds = expense_kinds(filtered_invoices)
            filtered_items = iter_lines(filtered_lines)
            fallback_info = attempt_level(
                decls,
                filtered_items,
                {
                    name: int(round(amount * 100))
                    for name, amount in filtered_kinds.items()
                    if amount
                },
                expense_cents(filtered_invoices),
                (
                    0
                    if any(0 < abs(a) <= TINY_FEE_YUAN + 1e-9 for a in filtered_kinds.values())
                    else TOLERANCE_CENTS
                ),
                fee_directions(filtered_invoices),
            )
            if any(fallback_info["result"].assignments.get(decl) for decl in decls):
                scope_label = "按合同订单过滤"
                invoices = filtered_invoices
                chosen_level = "按合同订单过滤"
                info = fallback_info
                items = filtered_items
        _elapsed = _time.perf_counter() - _started
        if debug and _elapsed > 1.0:
            print(f"[slow] {contract} {_elapsed:.1f}s 池={len(narrow)}/{len(wide)}", flush=True)
        if debug:
            print(
                f"[debug] {contract} 采用={scope_label} 粒度={chosen_level} "
                f"精确={info['exact_count']}/{len(decls)} "
                f"多解={info['mode_ambiguous']} 口径={info['combo']} notes={info['result'].notes}",
                flush=True,
            )
        result = info["result"]
        exact_best = info["exact_count"]
        combo_best = info["combo"]
        combo_scores = info["combo_scores"]
        ambiguous = info["mode_ambiguous"]
        directed_by_decl: dict[str, list[dict]] = {}
        extra_directed_reason = ""
        if any(not result.assignments.get(d) for d in decls):
            d_rows, d_reason = directed_split_rows(
                contract,
                decls,
                # 窄池按映射表直取（含 ADD 等后缀单号），宽池按订单全量兜底；
                # 两者取并集，避免「精确出运单金额不足时 ADD 订单的行被漏掉」
                narrow + [row for row in wide if row not in narrow],
                names,
                only_decls={d for d in decls if not result.assignments.get(d)},
                debug=debug,
            )
            for d_row in d_rows:
                directed_by_decl.setdefault(d_row["报关单号"], []).append(d_row)
            if directed_by_decl:
                strategy_log.append(
                    {
                        "合同号": contract,
                        "出运单": "、".join(invoices),
                        "费用明细": "、".join(f"{k}={v/100:.2f}" for k, v in fees.items()) or "(无)",
                        "费用种类数": len(fees),
                        "报关单张数": len(decls),
                        "命中粒度": "按合同订单+品名定向落组",
                        "选中口径": "directed",
                        "精确闭合报关单数": len(directed_by_decl),
                        "有区分度": False,
                        "各组合精确数": "directed",
                    }
                )
            elif d_reason:
                extra_directed_reason = d_reason
        if debug:
            if directed_by_decl:
                print(
                    f"[directed] {contract} 定向兜底报关单="
                    + "、".join(sorted(directed_by_decl)),
                    flush=True,
                )
            elif extra_directed_reason:
                print(f"[directed] {contract} 定向兜底未生效：{extra_directed_reason}", flush=True)
        ambiguity_note = (
            "费用口径不唯一：" + "、".join(sorted(combo_scores)) if ambiguous else ""
        )
        # 命中类型是否与其它组合有区分度
        distinct = len({v for v in combo_scores.values()}) > 1
        strategy_log.append(
            {
                "合同号": contract,
                "出运单": "、".join(invoices),
                "费用明细": "、".join(f"{k}={v/100:.2f}" for k, v in fees.items()) or "(无)",
                "费用种类数": len(fees),
                "报关单张数": len(decls),
                "命中粒度": chosen_level,
                "选中口径": "、".join(f"{k}:{v}" for k, v in combo_best.items()) or "(无费用)",
                "精确闭合报关单数": exact_best,
                "有区分度": distinct,
                **{f"费用[{name}]口径": mode for name, mode in combo_best.items()},
                "各组合精确数": "、".join(f"{k}:{v}" for k, v in combo_scores.items()),
            }
        )

        # 该合同最终采用的出运产品行池。下面这个循环里 `items` 会被按「采购单+产品类型」
        # 重新赋值，要留住池子本身，做完之后才能查「有没有行没被认领」。
        split_any = any(result.assignments.get(decl) for decl in decls)
        pool_items = list(items)
        # —— 同一出运单里「只差 SKU」的孪生产品行 ——
        # 报关单金额只能闭合其中一条时，另一条既没有报关单认领、也不会报错，
        # 于是整张采购单只算了一半成本（26MT-05X246A：PI-26MT-05X246A 有两条 24,300 的
        # 产品行，报关单 223320260001297365 只闭合了一条，飞书按整单 48,600 给钱）。
        # 处理：只在「已经有一条同出运单 / 同采购单 / 同品名 / 同金额 / 同数量的行被这张
        # 合同的记录认领」时，把剩下那条作为**成本补记行**并进同一条记录：
        # `出运金额`、`报关金额` 一点不动（那是报关单原件上的事实），只把它的
        # `出运采购金额(RMB)` 加进成本，并在「补记成本」列写明。
        extra_cost: dict[tuple, list] = {}
        if split_any and not ambiguous and not directed_by_decl:
            claimed_lines = [
                line
                for _solution in (result.assignments or {}).values()
                for line in flatten(_solution)
            ]
            claimed_rows = {row_key(line.raw) for line in claimed_lines}
            twins = {twin_key(line) for line in claimed_lines}
            for line in flatten(pool_items):
                if row_key(line.raw) in claimed_rows:
                    continue
                if twin_key(line) in twins:
                    extra_cost.setdefault(twin_key(line), []).append(line)
                    extra_rows_all[row_key(line.raw)] = contract
        for decl, rows in decls.items():
            used = result.assignments.get(decl)
            if not used:
                # 最后兜底：按「合同订单 + 海关编码/品名」定向落组
                # （出运单合并了多个外销单、报关单又按外销单拆开的场景）
                if directed_by_decl.get(decl):
                    split_rows.extend(directed_by_decl[decl])
                    continue
                note = next(
                    (n for n in result.notes if n.startswith(f"{decl}:")), ""
                )
                for row in rows:
                    failures.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            "报关品名": row.get("name"),
                            "比对结果": (
                                ambiguity_note
                                or note
                                or "子集和无解（容量 "
                                + f"{sum(to_cents(r.get('amount')) for r in rows) / 100:.2f}）"
                            ),
                        }
                    )
                continue
            if ambiguous:
                for row in rows:
                    failures.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            "报关品名": row.get("name"),
                            "比对结果": ambiguity_note,
                        }
                    )
                continue
            used = expand(items, used)
            row_map = assign_to_rows(rows, used, names.short)
            if debug:
                print(
                    f"[assign] {decl} 报关行="
                    + str([(r.get("name"), r.get("hs"), r.get("amount")) for r in rows])
                    + " 落行="
                    + str(
                        {
                            index: [(line.supplier, line.amount / 100) for line in picked]
                            for index, picked in row_map.items()
                        }
                    ),
                    flush=True,
                )
            for index, row in enumerate(rows):
                picked = row_map.get(index, [])
                for line in picked:
                    # 混装放行要比对报关单上的品名，出运品名（细分）与粗分类都登记一份
                    for kind in (line_product_type(line), str(line.customs_name or "").strip()):
                        if kind:
                            decl_products[decl].add(kind)
                by_supplier: dict[str, list] = defaultdict(list)
                for item in picked:
                    by_supplier[names.short(item.supplier)].append(item)
                if not by_supplier:
                    # 该报关单整体已拆出，但这一行没分到产品行：不能静默丢弃
                    failures.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            "报关品名": row.get("name"),
                            "比对结果": "该报关行未分到产品行（供应商/海关编码/品名/金额均未匹配）",
                        }
                    )
                    continue
                # 拆到「采购订单 + 供应商 + 产品类型」：同一家工厂两种产品类型要两条记录
                item_groups = [
                    (
                        names.short(items[0].supplier),
                        purchase_code,
                        product_type,
                        items,
                    )
                    for short, supplier_items in by_supplier.items()
                    for (purchase_code, product_type), items in split_by_item(supplier_items)
                ]
                declared_cents = to_cents(row.get("amount"))
                assigned_total = sum(x.amount for x in picked)
                # —— 报关金额按各组货值比例分摊，最后一组吸收分位尾差 ——
                # 出运金额合计 + 客户费用分摊 = 报关金额（每行都能对上，不会被误读成漏拆）
                allocated_cents_list: list[int] = []
                allocated_all = 0
                for position, (_s, _p, _t, items) in enumerate(item_groups):
                    share = sum(x.amount for x in items)
                    if position == len(item_groups) - 1:
                        allocated_cents = declared_cents - allocated_all
                    else:
                        allocated = (
                            declared_cents * share / assigned_total
                            if assigned_total
                            else share
                        )
                        allocated_cents = int(round(allocated))
                        allocated_all += allocated_cents
                    allocated_cents_list.append(allocated_cents)

                # —— 报关重量分摊（2026-09-28 用户口径）——
                # 单条记录直接回填；多条记录优先按入库单附件重量，其次按报关金额占比，最后一组吸收尾差。
                declared_weight = _dec(row.get("weight"))
                weight_parts, weight_basis = allocate_weight(
                    declared_weight,
                    [group[1] for group in item_groups],
                    allocated_cents_list,
                )

                for position, (short, purchase_code, product_type, items) in enumerate(
                    item_groups
                ):
                    share = sum(x.amount for x in items)
                    allocated_cents = allocated_cents_list[position]
                    # 成本补记行：同出运单里「只差 SKU」、报关单闭合不了也没人认领的那条，
                    # 只进成本（出运采购金额(RMB)），不进出运金额/报关金额
                    extra_lines: list = []
                    extra_seen: set[tuple] = set()
                    for line in items:
                        for extra in extra_cost.get(twin_key(line), []):
                            key = row_key(extra.raw)
                            if key in extra_seen:
                                continue
                            extra_seen.add(key)
                            extra_lines.append(extra)
                    # 产品类型：商品资料粗分类优先；没有类别时按出运品名折算，
                    # 仍取不到才用报关品名兜底（都归到同一套粗分类，避免粒度混用）
                    final_product_type = coarse_label(product_type, row.get("name"))
                    fee_cents = allocated_cents - share
                    split_rows.append(
                        {
                            "报关单号": decl,
                            "合同号_1": contract,
                            # 报关品名保持「报关单自己的品名」（飞书 AI 表的 406 行全按这个口径），
                            # 实际产品类型另立一列，拆单记录才既对得上主表、又能按产品类型对成本
                            "报关品名": str(row.get("name") or ""),
                            "海关编码": str(row.get("hs") or ""),
                            "产品类型": final_product_type,
                            "供应商简称": short,
                            "供应商": items[0].supplier,
                            "采购单号": purchase_code,
                            "出运金额合计": f"{share / 100:.2f}",
                            "出运采购金额合计": f"{sum(line_rmb_cents(x) for x in items + extra_lines) / 100:.2f}",
                            "报关金额": f"{allocated_cents / 100:.2f}",
                            "客户费用分摊": f"{fee_cents / 100:.2f}",
                            "报关重量": fmt_weight(weight_parts[position]),
                            "重量分摊依据": weight_basis,
                            "产品行数": len(items),
                            "补记成本": "、".join(
                                f"{extra.purchase_code} {str((extra.raw or {}).get('sku') or '')}"
                                f" {line_rmb_cents(extra) / 100:.2f}"
                                for extra in extra_lines
                            ),
                        }
                    )

        # —— 出运产品行有没有被「静默落下」——
        # 报关单只认领了整张出运单的一部分时，剩下的产品行既不进记录、也不报错。
        # 26MT-05X246A 就是这样：PI-26MT-05X246A 有两条 24,300 的产品行（SKU 只差尾号），
        # 报关单 223320260001297365 的金额只能闭合其中一条，另一条被无声丢掉，
        # 于是我方只有半个采购单（24,300），飞书按整单（48,600）给钱。
        # 只在「这张合同真的拆出了记录」时才谈"落下"：多解被整单挂起或走了定向兜底的，
        # 池子里的行本来就还没分配，不算漏（会在异常表里单独说明）。
        if split_any and not ambiguous and not directed_by_decl:
            for _solution in (result.assignments or {}).values():
                for _line in flatten(_solution):
                    claimed_rows_all.add(row_key(_line.raw))
                    claimed_twins_all.add(
                        (
                            _line.purchase_code,
                            _line.customs_name,
                            _line.amount,
                            str(_line.quantity),
                        )
                    )
            for line in flatten(pool_items):
                # 同一张出运单可能被两张报关单（两个合同号）共用，行会被两边都算进池子，
                # 所以按「原始行」去重，最后统一看谁没被任何报关单认领。
                pool_rows_all.setdefault(row_key(line.raw), (contract, line))

    # 与飞书比对
    # 先做一次 ERP 数据自检：供应商（编码/名称）与产品类别矛盾时，
    # 按「工厂→产品类别」知识库反推实际工厂，并把这条数据**同时记成异常**——
    # 睿贝里出现这种矛盾本身就是异常数据，不能因为最后算出了金额就当正常数据用。
    from app.services.supplier_products import SupplierProducts  # noqa: E402

    supplier_products = SupplierProducts()
    data_anomalies: list[dict] = []
    for row in split_rows:
        short = str(row.get("供应商简称") or "")
        product = str(row.get("报关品名") or "")
        if not short:
            continue
        ok, verdict_reason = supplier_products.verdict(short, product)
        if ok:
            # 名录不全但 ERP 采购历史反复出现 → 直接把产品补进映射表，下次不用再算
            learned = supplier_products.learn(short, product)
            if learned:
                print(
                    f"[learn] 映射表新增：{short} += {learned}"
                    f"（{verdict_reason}；{row.get('报关单号')} {product}）",
                    flush=True,
                )
            continue
        inferred, reason = supplier_products.infer_factory(
            product, str(row.get("采购单号") or "")
        )
        if inferred:
            row["供应商简称"] = inferred
            row["供应商"] = names.full(inferred) or str(row.get("供应商") or "")
            note = f"ERP数据异常：{short} 不生产「{product}」，按产品反推为 {inferred}（{reason}）"
        else:
            note = f"ERP数据异常：{short} 不生产「{product}」，反推失败（{reason}）"
        data_anomalies.append(
            {
                "报关单号": row.get("报关单号"),
                "合同号_1": row.get("合同号_1"),
                "报关品名": product,
                "采购单号": row.get("采购单号"),
                "系统供应商简称": short,
                "飞书供应商简称": "",
                "比对结果": note,
            }
        )

    # 「混装报关」的品名口径（业务规则，纯结构性、对全量一视同仁）：
    # 同一 `报关单号` + 同一 `采购单号` + 同一 `供应商` 下拆出了两种及以上的**实际**
    # 产品类型时，谁是谁无法区分，此时以**实际品名**为准（飞书手工拆行时也是这么写的）；
    # 其余情况一律保留报关单原件上的品名。
    mixed_groups: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    for row in split_rows:
        product = str(row.get("产品类型") or "").strip()
        if product:
            mixed_groups[
                (
                    str(row.get("报关单号") or ""),
                    str(row.get("采购单号") or ""),
                    str(row.get("供应商简称") or ""),
                )
            ].add(product)
    ambiguous_groups = {key for key, kinds in mixed_groups.items() if len(kinds) >= 2}
    mixed_named = 0
    for row in split_rows:
        key = (
            str(row.get("报关单号") or ""),
            str(row.get("采购单号") or ""),
            str(row.get("供应商简称") or ""),
        )
        product = str(row.get("产品类型") or "").strip()
        # 只有「实际品名是细分名」时才改写报关品名；产品类型已经统一成粗分类后，
        # 再拿它去覆盖报关品名会把「报关单品名」这一列写坏（文档口径：该列必须跟报关单一致）
        if (
            key in ambiguous_groups
            and product
            and product not in KNOWN_CATEGORIES
            and str(row.get("报关品名") or "") != product
        ):
            row["报关品名"] = product
            mixed_named += 1
    if mixed_named:
        print(
            f"[混装品名] {len(ambiguous_groups)} 组「同报关单+同采购单+同供应商」含多种实际品名，"
            f"其中 {mixed_named} 行按实际品名修正报关品名",
            flush=True,
        )

    ref: dict[tuple[str, str], set[str]] = defaultdict(set)
    ref_contract: dict[tuple[str, str], str] = {}
    for record in json.loads(REF_PATH.read_text(encoding="utf-8")):
        fields = record.get("fields", {})
        decl = flat(fields.get("报关单号"))
        name = flat(fields.get("报关品名"))
        short = flat(fields.get("供应商简称"))
        contract = flat(fields.get("合同号_1"))
        if decl and name and short:
            ref[(decl, name)].add(short)
        if decl and name and contract:
            ref_contract.setdefault((decl, name), contract)

    system: dict[tuple[str, str], set[str]] = defaultdict(set)
    system_contract: dict[tuple[str, str], str] = {}
    for row in split_rows:
        if row["供应商简称"]:
            # 一条拆单结果可能横跨多家供应商（供应商简称里带顿号），拆开再入集合
            for part in str(row["供应商简称"]).replace("、", "|").split("|"):
                part = part.strip()
                if part:
                    system[(row["报关单号"], row["报关品名"])].add(part)
        system_contract.setdefault(
            (row["报关单号"], row["报关品名"]), row.get("合同号_1") or ""
        )

    same: list[dict] = []
    bad: list[dict] = []
    missing: list[dict] = []
    # 同一报关单内供应商的**整体**集合：用来识别「飞书漏拆」——
    # 我方按实际品名多出一行、飞书没写这一行，但整张单的供应商其实都对得上。
    decl_system: dict[str, set[str]] = defaultdict(set)
    decl_ref: dict[str, set[str]] = defaultdict(set)
    for (decl, _name), shorts in system.items():
        decl_system[decl] |= shorts
    for (decl, _name), shorts in ref.items():
        decl_ref[decl] |= shorts
    for (decl, name), shorts in sorted(system.items()):
        expected = ref.get((decl, name), set())
        record = {
            "报关单号": decl,
            "合同号_1": system_contract.get((decl, name)) or ref_contract.get((decl, name), ""),
            "报关品名": name,
            "系统供应商简称": "、".join(sorted(shorts)),
            "飞书供应商简称": "、".join(sorted(expected)) or "(未填)",
        }
        if shorts == expected:
            record["比对结果"] = "一致"
            same.append(record)
        elif not expected and decl_system.get(decl, set()) == decl_ref.get(decl, set()):
            # 飞书这一行没写：整张单的供应商都对得上，只是飞书少拆了一行
            record["比对结果"] = (
                f"飞书漏拆（我方多一行「{name}」，飞书该单没有这一行）"
            )
            missing.append(record)
        else:
            record["比对结果"] = (
                f"不一致 多={sorted(shorts - expected)} 缺={sorted(expected - shorts)}"
            )
            bad.append(record)
    # 飞书有、系统没覆盖到的
    decl_suppliers: dict[str, set[str]] = defaultdict(set)
    for row in split_rows:
        if row.get("供应商简称"):
            decl_suppliers[row["报关单号"]].add(row["供应商简称"])
    for (decl, name), shorts in sorted(ref.items()):
        if (decl, name) in system:
            continue
        my_decls = {d.declaration_no for d in declared}
        if decl not in my_decls:
            continue
        # 混装报关：报关单只写了一个主品名，实际装了多种产品（业务员有意为之）。
        # 只要这个品名确实是我们分给该报关单的产品、且供应商也一致，就不算异常。
        if name in decl_products.get(decl, set()) and shorts and shorts <= decl_suppliers.get(
            decl, set()
        ):
            same.append(
                {
                    "报关单号": decl,
                    "合同号_1": ref_contract.get((decl, name), ""),
                    "报关品名": name,
                    "系统供应商简称": "、".join(sorted(decl_suppliers.get(decl, set()))),
                    "飞书供应商简称": "、".join(sorted(shorts)),
                    "比对结果": "一致（混装报关：报关品名与实际商品名不同）",
                }
            )
            continue
        bad.append(
            {
                "报关单号": decl,
                "合同号_1": ref_contract.get((decl, name), ""),
                "报关品名": name,
                "系统供应商简称": "(未拆出)",
                "飞书供应商简称": "、".join(sorted(shorts)),
                "比对结果": "未拆出",
            }
        )
    for row in failures:
        bad.append(
            {
                "报关单号": row.get("报关单号"),
                "合同号_1": row.get("合同号_1") or ref_contract.get(
                    (row.get("报关单号"), row.get("报关品名")), ""
                ),
                "报关品名": row.get("报关品名"),
                "系统供应商简称": "",
                "飞书供应商简称": "",
                "比对结果": row.get("比对结果"),
            }
        )
    # ERP 自身的矛盾数据（供应商与产品对不上）也要进异常表
    bad.extend(data_anomalies)

    # 已人工核实的飞书/睿贝冲突：从「正常/待解决异常」里摘出来单列，
    # 避免和真·未解决问题混在一起。注意**正常行也要摘**——有些单子比对当时
    # 只是供应商集合一致，事后核实整张单的数据都是错的（如 25MT-03R625Y）。
    verified_path = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
    verified: list[dict] = []
    if verified_path.exists():
        cases = json.loads(verified_path.read_text(encoding="utf-8")).get("cases") or []
        # 核实是按**整张报关单**核的（人工核实结论本来就是针对这张单的），
        # 所以以报关单号为键：混装单的报关品名会按实际品名改写，用「报关单号+品名」
        # 当键会在改写后失配。
        index = {str(case.get("报关单号") or ""): case for case in cases}
        for source in ("bad", "same", "missing"):
            rows = {"bad": bad, "same": same, "missing": missing}[source]
            remaining: list[dict] = []
            for row in rows:
                case = index.get(str(row.get("报关单号") or ""))
                if case is None:
                    remaining.append(row)
                    continue
                verified.append(
                    {
                        **row,
                        "比对结果": (
                            f"已核实（{case.get('类别') or '飞书错'}）：{case.get('结论')}"
                            f"｜{case.get('证据')}"
                        ),
                    }
                )
            if source == "bad":
                bad = remaining
            elif source == "missing":
                missing = remaining
            else:
                same = remaining
    else:
        index = {}
        print("（未找到 data/reference/verified_conflicts.json，跳过已核实冲突分列）")

    detail_columns = [
        "报关单号", "合同号_1", "报关品名", "海关编码", "产品类型",
        "供应商简称", "供应商", "采购单号", "出运金额合计", "出运采购金额合计",
        "报关金额", "客户费用分摊", "报关重量", "重量分摊依据", "产品行数", "补记成本",
    ]
    group_columns = [
        "报关单号", "合同号_1", "报关品名",
        "系统供应商简称", "飞书供应商简称", "比对结果",
    ]
    write_rows(split_rows, out_dir / "拆单明细_全部.xlsx", detail_columns)
    write_rows(same, out_dir / "拆单结果_正常.xlsx", group_columns)
    write_rows(bad, out_dir / "拆单结果_异常.xlsx", group_columns)
    write_rows(verified, out_dir / "拆单结果_已核实.xlsx", group_columns)
    write_rows(missing, out_dir / "拆单结果_飞书漏拆.xlsx", group_columns)
    # 出运单里没有报关单认领的产品行（以前是静默丢弃，现在单独列表）
    unclaimed: list[dict] = []
    for key, (owner, line) in pool_rows_all.items():
        if key in claimed_rows_all:
            continue
        raw = dict(line.raw or {})
        unclaimed.append(
            {
                "合同号_1": owner,
                "出运单": str(raw.get("invoice_code") or ""),
                "采购单号": line.purchase_code,
                "SKU": str(raw.get("sku") or ""),
                "供应商": line.supplier,
                "出运金额": f"{line.amount / 100:.2f}",
                "出运采购金额(RMB)": raw.get("amount_rmb") or "",
                "报关品名": line.customs_name,
                "同单已有孪生行": (
                    "是"
                    if (
                        line.purchase_code,
                        line.customs_name,
                        line.amount,
                        str(line.quantity),
                    )
                    in claimed_twins_all
                    else ""
                ),
                "处理": (
                    f"已按孪生行规则补记成本到 {extra_rows_all[key]} 的记录"
                    if key in extra_rows_all
                    else ""
                ),
            }
        )
    write_rows(
        unclaimed,
        out_dir / "拆单_未认领产品行.xlsx",
        ["合同号_1", "出运单", "采购单号", "SKU", "供应商", "出运金额",
         "出运采购金额(RMB)", "报关品名", "同单已有孪生行", "处理"],
    )
    # 本地异常（不依赖飞书比对）：拆不出/多解/未落到报关行 + ERP 自身矛盾。
    # 已人工核实的报关单不再算「未解决异常」（它们在拆单结果_已核实.xlsx 里）。
    local_bad: list[dict] = []
    for row in failures:
        if index.get(str(row.get("报关单号") or "")):
            continue
        local_bad.append(
            {
                "类型": "拆单失败",
                "报关单号": row.get("报关单号"),
                "合同号_1": row.get("合同号_1"),
                "报关品名": row.get("报关品名"),
                "说明": row.get("比对结果"),
            }
        )
    for row in data_anomalies:
        if index.get(str(row.get("报关单号") or "")):
            continue
        local_bad.append(
            {
                "类型": "ERP数据矛盾",
                "报关单号": row.get("报关单号"),
                "合同号_1": row.get("合同号_1"),
                "报关品名": row.get("报关品名"),
                "采购单号": row.get("采购单号"),
                "说明": row.get("比对结果"),
            }
        )
    write_rows(
        local_bad,
        out_dir / "拆单_本地异常.xlsx",
        ["类型", "报关单号", "合同号_1", "报关品名", "采购单号", "说明"],
    )
    # 费用分摊台账：逐种费用记录命中的分摊口径
    report_dir = PROJECT_ROOT / ".cache" / "erp" / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    write_rows(
        strategy_log,
        report_dir / "fee_strategy.xlsx",
        ["合同号", "出运单", "费用明细", "费用种类数", "报关单张数",
         "命中粒度", "选中口径", "精确闭合报关单数", "有区分度", "各组合精确数"],
    )
    from collections import Counter

    # 统计口径：
    # 1) 报关单张数 > 1（只有多张报关单时，equal/by_weight/by_amount 才可能不同）
    # 2) 各组合结果有区分度
    # 3) 至少一张报关单精确闭合
    disc = [
        r
        for r in strategy_log
        if r.get("有区分度")
        and r["精确闭合报关单数"] > 0
        and int(r.get("报关单张数") or 0) > 1
    ]
    ties = [
        r
        for r in strategy_log
        if r["精确闭合报关单数"] > 0 and int(r.get("报关单张数") or 0) <= 1
    ]
    per_fee: dict[str, Counter] = {}
    for row in disc:
        for key, value in row.items():
            if key.startswith("费用[") and value:
                per_fee.setdefault(key[3:-3], Counter())[value] += 1
    summary_lines = [
        "各费用分摊口径台账（只统计有区分度、且至少一张报关单精确闭合的出运单）",
        f"有区分度的出运单数={len(disc)}",
    ]
    for fee_name, counter in sorted(per_fee.items()):
        top = counter.most_common()
        verdict = "一致，可认定为该口径" if len(top) == 1 else "不一致，需更多样本"
        summary_lines.append(
            f"  费用「{fee_name}」命中分布：" + "、".join(f"{k}={v}" for k, v in top) + f" → {verdict}"
        )
    if not per_fee:
        summary_lines.append("  暂无有区分度的样本")
    summary_lines.append(f"单张报关单（口径无法区分，不计入推断）={len(ties)}")
    summary_lines.append(
        "无附加费用的出运单数="
        + str(sum(1 for r in strategy_log if r.get("费用种类数") == 0))
    )
    (report_dir / "fee_strategy.txt").write_text("\n".join(summary_lines), encoding="utf-8")
    print("\n".join(summary_lines))
    summary = {
        "输入报关商品行": len(declared),
        "拆出记录": len(split_rows),
        "比对单元(报关单+品名)": len(system),
        "一致": len(same),
        "已核实": len(verified),
        "飞书漏拆": len(missing),
        "异常": len(bad),
    }
    (out_dir / "拆单统计.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False))
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--only", default="", help="只跑这些合同（逗号分隔）")
    parser.add_argument("--parse", default="", help="报关单解析结果 xlsx（默认用固定的出口退税联）")
    parser.add_argument("--out", default="", help="输出目录（默认 outputs/shipments_split）")
    parser.add_argument("--debug", action="store_true", help="打印逐合同的求解诊断")
    args = parser.parse_args()
    run(
        parse_xlsx=Path(args.parse) if args.parse else None,
        out_dir=Path(args.out) if args.out else None,
        only=args.only,
        debug=args.debug,
    )


def write_rows(rows: list[dict], path: Path, columns: list[str]) -> None:
    """写出 Excel；目标被 Excel 占用时改写到带时间戳的文件，避免整批失败。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "结果"
    sheet.append(columns)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(c, "") for c in columns])
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), columns):
        sheet.column_dimensions[column[0].column_letter].width = max(12, min(38, len(name) * 2))
    sheet.freeze_panes = "A2"
    try:
        workbook.save(path)
    except PermissionError:
        from datetime import datetime

        alt = path.with_name(f"{path.stem}_{datetime.now():%H%M%S}{path.suffix}")
        workbook.save(alt)
        print(f"（{path.name} 被占用，改写到 {alt.name}）", flush=True)


if __name__ == "__main__":
    main()
