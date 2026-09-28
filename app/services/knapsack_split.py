"""用「子集和（0-1 背包）」把出运产品行分配给报关单。

问题：一张出运单可能对应多张报关单，单看数据无法直接知道哪些产品属于哪张报关单。
办法：报关单上有报关金额，出运产品行有出运金额，于是变成**子集和**问题——
找一组产品行，其出运金额之和正好等于该报关单的容量。

已确认的口径（实测 26MT-07T042 / 26MT-05X167）：
- 报关金额 = 产品出运金额 + 分摊到该报关单的运费（CIF 成交方式下）；
- 运费按**报关单张数**平分（07T042 海运费 2000，两张报关单各 1000，精确闭合）；
- 因此容量 = 该报关单所有行的报关金额之和 − 运费均摊份额。

算法：
1. 用 meet-in-the-middle 精确求子集和（22 条产品行只需 2^11 规模）；
2. 按容量从小到大逐个报关单求解，已占用的产品行不再参与；
3. 启发式：优先同一供应商 + 同海关编码的产品行（实测能显著缩小搜索范围）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from dataclasses import replace
from decimal import Decimal
import bisect
import json
import os
from pathlib import Path

from itertools import combinations
from functools import lru_cache
from typing import Any, Iterable


# 报关品名 → 粗分类（飞书「产品类型」表，data/reference/feishu_product_map.json）。
# 用于「出运行的粗分类」与「报关行的品名」互相印证。
_DECL_CATEGORY: dict[str, str] | None = None


def decl_category(name: str) -> str:
    global _DECL_CATEGORY
    if _DECL_CATEGORY is None:
        path = (
            Path(__file__).resolve().parents[2]
            / "data"
            / "reference"
            / "feishu_product_map.json"
        )
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            _DECL_CATEGORY = {
                str(key).strip(): str(value).strip()
                for key, value in payload.items()
                if str(key).strip()
            }
        except Exception:  # noqa: BLE001
            _DECL_CATEGORY = {}
    return _DECL_CATEGORY.get(str(name or "").strip(), "")


def name_matches(line: "Line", want_name: str) -> bool:
    """报关品名与产品行是否相符。

    优先比 ERP 出运品名（原口径）；出运品名缺失时的兜底：把报关品名折算成粗分类，
    与产品行的粗分类比较（产品类型现在是商品资料的粗分类）。
    """
    if not want_name:
        return False
    if line.customs_name and line.customs_name == want_name:
        return True
    want = decl_category(want_name)
    return bool(want) and line_product_type(line) == want


def line_product_type(line: "Line") -> str:
    """产品行的品类（粗分类）。

    **出运品名折算优先**：`海关商品（中文）` 是报关/对账一直在用的口径，折算成粗分类
    后仍能区分「法兰 / 管件」这类实际差异；商品资料里同一个产品编码偶尔写错类别
    （实测 三通的编码被标成法兰/管件混用），用它当主口径反而会把两种货并成一条。
    出运品名缺失时才退回商品资料「类别名称」（这正是本任务要补的缺口）。
    """
    return decl_category(line.customs_name) or line.product_type


def to_cents(value: Any) -> int:
    """金额转「分」，避免浮点误差。"""
    text = str(value or "").replace(",", "").strip()
    if not text:
        return 0
    try:
        return int((Decimal(text) * 100).to_integral_value())
    except Exception:  # noqa: BLE001
        return 0


def _half_sums(items: list[tuple[int, int]]) -> dict[int, tuple[int, ...]]:
    """枚举一半元素的所有子集和。返回 {和: 元素下标组合}。"""
    return _half_sums_cached(tuple(items))


@lru_cache(maxsize=4)
def _half_sums_cached(items: tuple) -> dict[int, tuple[int, ...]]:
    """带缓存的半集枚举。

    同一批候选行会被「多个口径组合 × 多张报关单 × 多个粒度」反复求解，
    缓存后可以把最贵的 2^(n/2) 枚举从上百次降到个位数。
    返回值只读，调用方不要就地修改。
    """
    out: dict[int, tuple[int, ...]] = {}
    n = len(items)
    for size in range(n + 1):
        for combo in combinations(range(n), size):
            total = sum(items[i][1] for i in combo)
            # 同一个和只保留一种组合（价格相同的行在逻辑上等效）
            out.setdefault(total, tuple(items[i][0] for i in combo))
    return out


def subset_sum_solutions(
    items: list[tuple[int, int]],
    target: int,
    tolerance: int = 0,
    limit: int = 2,
) -> list[list[int]]:
    if target <= 0:
        return [[]]
    # 单条金额已经超过「容量 + 容差」的记录不可能出现在解里，先剪掉：
    # 既缩小搜索空间，也让半集枚举小好几个数量级
    usable = tuple(
        (index, amount)
        for index, amount in items
        if 0 < amount <= target + tolerance
    )
    return _subset_sum_solutions_cached(usable, target, tolerance, limit)


@lru_cache(maxsize=256)
def _subset_sum_solutions_cached(
    items: tuple,
    target: int,
    tolerance: int,
    limit: int,
) -> list[list[int]]:
    """一次 meet-in-the-middle，同时拿到最多 `limit` 组**不同**的子集和。

    既做「找解」，也做「多解校验」：只枚举一遍就能判断是否存在第二组记录
    也能凑出同样的金额，比「逐条剔除再重解」快一个数量级。

    同一批候选行会被「多个费用口径 × 多张报关单 × 多个粒度」反复求解，
    因此按 (候选行, 目标金额) 做缓存；返回值只读，调用方不要就地修改。
    """
    if target <= 0:
        return [[]] if target <= 0 else []
    if not items:
        return []
    total_sum = sum(amount for _, amount in items)
    if total_sum < target - tolerance:
        return []
    if abs(total_sum - target) <= tolerance:
        return [[index for index, _ in items]]
    mid = len(items) // 2
    left_sums = _half_sums(items[:mid])
    right_sums = _half_sums(items[mid:])
    ordered = sorted(right_sums.items())
    keys = [key for key, _ in ordered]

    found: list[tuple[int, list[int]]] = []
    signatures: set[frozenset] = set()
    for left_sum, left_pick in left_sums.items():
        need = target - left_sum
        if tolerance <= 0:
            right_pick = right_sums.get(need)
            if right_pick is None:
                continue
            combo = list(left_pick) + list(right_pick)
            signature = frozenset(combo)
            if signature in signatures:
                continue
            signatures.add(signature)
            found.append((0, combo))
        else:
            position = bisect.bisect_left(keys, need)
            for candidate_index in (position - 1, position, position + 1):
                if not 0 <= candidate_index < len(ordered):
                    continue
                right_sum, right_pick = ordered[candidate_index]
                diff = abs(left_sum + right_sum - target)
                if diff > tolerance:
                    continue
                combo = list(left_pick) + list(right_pick)
                signature = frozenset(combo)
                if signature in signatures:
                    continue
                signatures.add(signature)
                found.append((diff, combo))
        if len(found) >= limit:
            # 已经凑够 `limit` 组「不同」的解：多解结论已成立，不必再扫完
            break
    found.sort(key=lambda item: item[0])
    return [combo for _diff, combo in found[:limit]]


@dataclass
class Line:
    """一条出运产品行。"""

    index: int
    amount: int            # 出运金额（分）
    supplier: str = ""
    hs_code: str = ""
    customs_name: str = ""
    product_type: str = ""  # 粗分类（商品资料「类别名称」去部门括号），拆分/落单的品类键
    unit_price: int = 0    # 外销单价（分）
    quantity: Decimal = Decimal(0)
    purchase_code: str = ""
    raw: dict[str, Any] = field(default_factory=dict)
    # 打包出来的「组合记录」不允许再按数量拆解
    splittable: bool = True
    members: list["Line"] = field(default_factory=list)
    # 原始数量/金额：部分出运时按比例折算剩余，一条产品行可被多张报关单分摊
    total_qty: Decimal = Decimal(0)
    total_amount: int = 0


@dataclass
class Target:
    """一张报关单（以报关单号为单位）。"""

    declaration_no: str
    rows: list[dict[str, Any]]
    declared_cents: int
    capacity: int          # 扣除均摊运费后的容量（分）
    lines: list[Line] = field(default_factory=list)


@dataclass
class SolveResult:
    assignments: dict[str, list[Line]] = field(default_factory=dict)
    unassigned: list[Line] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _line_priority(line: Line, row: dict[str, Any], row_hs: str) -> tuple[int, int]:
    """启发式排序键：越小越优先。

    1. 同海关编码 + 同供应商；
    2. 同海关编码；
    3. 海关编码不一致但金额匹配可能性高。
    """
    same_hs = 0 if (row_hs and line.hs_code == row_hs) else 1
    same_supplier = 0 if (row.get("supplier") and line.supplier == row["supplier"]) else 1
    return (same_hs, same_supplier)


def solve(
    declarations: dict[str, list[dict[str, Any]]],
    lines: list[Line],
    freight_cents: int = 0,
    *,
    use_heuristics: bool = True,
    freight_shares: dict[str, int] | None = None,
    tolerance: int = 0,
    allow_fraction: bool = False,
) -> SolveResult:
    """把产品行分配给各报关单。

    `declarations`: 报关单号 -> 该报关单的报关行列表（每行含 报关金额/报关品名/海关编码/供应商）。
    `lines`: 该合同下出运单的全部产品行。
    `freight_cents`: 本次出运单的运费总额（按报关单张数均摊）。
    """
    result = SolveResult()
    if not declarations or not lines:
        result.unassigned = list(lines)
        return result

    shares = dict(freight_shares or {}) if freight_shares else {}
    if not shares:
        decl_count = len(declarations)
        share = (freight_cents // decl_count) if decl_count else 0
        shares = {decl: share for decl in declarations}
    targets: list[Target] = []
    for decl, rows in declarations.items():
        declared = sum(to_cents(r.get("amount")) for r in rows)
        capacity = declared - shares.get(decl, 0)
        targets.append(
            Target(
                declaration_no=decl,
                rows=rows,
                declared_cents=declared,
                capacity=capacity,
            )
        )
    # 容量小的先解，约束更强
    targets.sort(key=lambda t: t.capacity)

    pool = list(lines)
    for target in targets:
        if not pool:
            result.notes.append(f"{target.declaration_no}: 无可用产品行")
            continue
        pool = _assign(
            target,
            pool,
            result,
            use_heuristics=use_heuristics,
            tolerance=tolerance,
            allow_fraction=allow_fraction,
        )
    result.unassigned = pool
    return result


def _assign(
    target: Target,
    pool: list[Line],
    result: SolveResult,
    *,
    use_heuristics: bool,
    tolerance: int = 0,
    allow_fraction: bool = False,
) -> list[Line]:
    """给一张报关单挑一组产品行，返回剩余的池子。"""
    original_pool = pool
    # 先用「海关编码」把候选收窄。同一张出运单常常混装多家工厂、多个产品类型
    # （如一张 25MT-07F521Y 里既有管件又有螺栓），整池求解会出现大量等价多解；
    # 收窄后若能唯一精确闭合，就直接采用，不必再动整池。
    narrowed = _narrow(target, pool)
    if 0 < len(narrowed) < len(pool) and len(narrowed) <= MAX_ITEMS:
        ordered_narrow = sorted(
            narrowed,
            key=lambda line: _line_priority(
                line, target.rows[0] if target.rows else {}, str((target.rows[0] if target.rows else {}).get("hs") or "")
            ),
        )
        narrow_items = [(line.index, line.amount) for line in ordered_narrow]
        narrow_solutions = subset_sum_solutions(
            narrow_items, target.capacity, tolerance, limit=2
        )
        if narrow_solutions and (
            len(narrow_solutions) == 1
            or solutions_equivalent(narrow_solutions, ordered_narrow)
        ):
            picked_index = set(narrow_solutions[0])
            result.assignments[target.declaration_no] = [
                line for line in ordered_narrow if line.index in picked_index
            ]
            return [line for line in pool if line.index not in picked_index]
    # 规模保护：候选过多时先用「海关编码 / 供应商」收窄，避免 2^n 组合爆炸
    if len(pool) > MAX_ITEMS:
        candidates = _narrow(target, pool)
        if len(candidates) > MAX_ITEMS:
            # 仍然很大：改用「按元粒度」的容量受限 DP
            items_big = [(line.index, line.amount) for line in candidates]
            picked_big = subset_sum_bounded(items_big, target.capacity, tolerance)
            if picked_big is None:
                # DP 只认「整行」组合。报关时少报/多报几支（部分出运数量配平）
                # 这类解它看不到，于是退一步：按「供应商」把池子切小，逐组走
                # 常规路径（整行子集和 + 部分数量配平），只有唯一命中才采纳。
                grouped = _large_pool_fallback(
                    target,
                    candidates,
                    original_pool,
                    tolerance,
                    allow_fraction=allow_fraction,
                )
                if grouped is not None:
                    used_group, remaining, why = grouped
                    result.assignments[target.declaration_no] = used_group
                    result.notes.append(f"{target.declaration_no}: {why}")
                    return remaining
                result.notes.append(
                    f"{target.declaration_no}: 候选产品行过多（{len(candidates)} 条），DP 未找到解"
                )
                result.assignments[target.declaration_no] = []
                return original_pool
            chosen_big = set(picked_big)
            used_big = [line for line in candidates if line.index in chosen_big]
            result.assignments[target.declaration_no] = used_big
            return [line for line in original_pool if line.index not in chosen_big]
        pool = candidates
    if len(target.rows) == 1 and not use_heuristics:
        ordered = pool
    else:
        row = target.rows[0] if target.rows else {}
        row_hs = str(row.get("海关编码") or "")
        ordered = sorted(pool, key=lambda line: _line_priority(line, row, row_hs))

    items = [(line.index, line.amount) for line in ordered]
    solutions = subset_sum_solutions(items, target.capacity, tolerance, limit=2)
    if len(solutions) > 1 and not solutions_equivalent(solutions, ordered):
        result.notes.append(
            f"{target.declaration_no}: 子集和多解（另有一组记录也能凑出同额），候选不唯一"
        )
        result.assignments[target.declaration_no] = []
        return original_pool
    solutions = solutions[:1]
    picked = solutions[0] if solutions else None
    if picked is None:
        partials = _try_partial(
            target, ordered, tolerance, limit=2, allow_fraction=allow_fraction
        )
        if len(partials) > 1:
            result.notes.append(
                f"{target.declaration_no}: 部分数量配平多解（至少 {len(partials)} 组组合），候选不唯一"
            )
            result.assignments[target.declaration_no] = []
            return original_pool
        if partials:
            used_lines, consumed, partial_line = partials[0]
            used_units = Decimal(str(consumed))
            consumed_amount = int(partial_line.unit_price * used_units)
            # 记录用快照：这条行稍后会被扣减剩余数量，避免串改已出结果的金额
            snapshot = replace(
                partial_line,
                amount=consumed_amount,
                quantity=used_units,
                members=[],
                splittable=False,
            )
            assignment = [snapshot if line is partial_line else line for line in used_lines]
            # 部分匹配属于「猜」，必须确认解唯一（背包内外无同价记录）
            if equivalent_swap_exists(assignment, original_pool):
                result.notes.append(
                    f"{target.declaration_no}: 部分匹配存在同价可替代记录，解不唯一"
                )
                result.assignments[target.declaration_no] = []
                return original_pool
            result.assignments[target.declaration_no] = assignment
            units_text = (
                f"{consumed:.4f}"
                if abs(consumed - round(consumed)) > 1e-9
                else f"{consumed:.2f}"
            )
            result.notes.append(
                f"{target.declaration_no}: 由部分出运数量配平（占用 {units_text} 个单位）"
            )
            # 只扣掉「整行用掉」的行；被部分占用的行按剩余数量折算后留在池子里，
            # 供同一订单的其它报关单继续分摊
            removed = {line.index for line in used_lines if line is not partial_line}
            remaining_qty = partial_line.quantity - used_units
            if remaining_qty <= 0:
                removed.add(partial_line.index)
                return [line for line in original_pool if line.index not in removed]
            total_qty = partial_line.total_qty or partial_line.quantity
            total_amount = partial_line.total_amount or partial_line.amount
            partial_line.quantity = remaining_qty
            partial_line.amount = int(
                (Decimal(total_amount) * remaining_qty / total_qty).to_integral_value()
            )
            return [line for line in original_pool if line.index not in removed]
    if picked is None:
        # 精确解不存在：先记录，保留整个池子给后面的报关单
        result.notes.append(
            f"{target.declaration_no}: 子集和无精确解（容量 {target.capacity / 100:.2f}）"
        )
        result.assignments[target.declaration_no] = []
        return original_pool

    chosen = {index for index in picked}
    used = [line for line in ordered if line.index in chosen]
    result.assignments[target.declaration_no] = used
    return [line for line in original_pool if line.index not in chosen]


def _large_pool_fallback(
    target: Target,
    candidates: list[Line],
    original_pool: list[Line],
    tolerance: int,
    *,
    allow_fraction: bool = False,
) -> tuple[list[Line], list[Line], str] | None:
    """大池子（候选行数超过 MAX_ITEMS）DP 未命中时的兜底。

    两条路依次试，都要「唯一命中」才采纳：
    1. 整包少装：除一条产品行按数量少装以外，其余全部属于这张报关单
       （业务上是报关时少报/多报几支，DP 只认整行组合看不到）；
    2. 按供应商整包：把池子切到单个供应商再走完整求解。
    """
    whole = _try_whole_pool_partial(target, candidates, original_pool, tolerance)
    if whole is not None:
        return whole
    return _solve_within_suppliers(
        target, candidates, original_pool, tolerance, allow_fraction=allow_fraction
    )


def _try_whole_pool_partial(
    target: Target,
    lines: list[Line],
    original_pool: list[Line],
    tolerance: int,
) -> tuple[list[Line], list[Line], str] | None:
    """「整包 + 其中一条按数量少装」：只枚举哪一条少装、少装多少。

    这是**猜**出来的解，所以要确认唯一：池子里若存在单价与总量都一致的
    另一条记录，换哪一条少装在业务上无法区分，按多解处理。

    先按「分」求精确闭合（报关金额通常就是出运金额减去整数个单位），
    精确解唯一才采纳；一条都没有时才退到容差内找，且同样要求唯一
    ——容差是 0.5 元、单位价常常才一两元，容差内会冒出十几个假解。
    """
    if target.capacity <= 0 or not lines:
        return None
    total = sum(line.amount for line in lines)
    for step_tolerance in (0, tolerance):
        hits = _whole_pool_partial_hits(
            target, lines, original_pool, total, step_tolerance
        )
        if len(hits) == 1:
            used, remaining, why = hits[0]
            return (*_trim_whole_pool(lines, used, remaining), why)
        if len(hits) > 1:
            return None
    return None


def _whole_pool_partial_hits(
    target: Target,
    lines: list[Line],
    original_pool: list[Line],
    total: int,
    tolerance: int,
) -> list[tuple[list[Line], list[Line], str]]:
    hits: list[tuple[list[Line], list[Line], str]] = []
    for line in lines:
        if not line.splittable or line.unit_price <= 0 or line.quantity <= 0:
            continue
        need = target.capacity - (total - line.amount)
        if need <= 0:
            continue
        exact_units = Decimal(need) / Decimal(line.unit_price)
        max_units = int(line.quantity)
        for units in {int(exact_units), int(exact_units) + 1}:
            if not 1 <= units < max_units:      # == max_units 就是「整包全吃」
                continue
            amount = units * line.unit_price
            if abs(amount - need) > tolerance:
                continue
            # 同单价 + 同总量的另一条记录：换谁少装都一样，解不唯一
            rival = any(
                other.index != line.index
                and other.unit_price == line.unit_price
                and other.quantity == line.quantity
                for other in lines
            )
            if rival:
                return []
            snapshot = replace(
                line,
                amount=amount,
                quantity=Decimal(units),
                members=[],
                splittable=False,
            )
            used = [snapshot if other is line else other for other in lines]
            removed = {other.index for other in lines if other is not line}
            remaining = [item for item in original_pool if item.index not in removed]
            label = "精确" if amount == need else "容差内"
            hits.append(
                (
                    used,
                    remaining,
                    f"整包少装（{label}闭合：{line.supplier or '未知供应商'} 有一条少装 "
                    f"{max_units - units} 个单位）",
                )
            )
    return hits


def _trim_whole_pool(
    lines: list[Line],
    used: list[Line],
    remaining: list[Line],
) -> tuple[list[Line], list[Line]]:
    """采纳「整包少装」前，把少装那条行在池子里扣减剩余数量供其它报关单分摊。"""
    for snapshot, source in zip(used, lines):
        if snapshot is source:
            continue
        remaining_units = source.quantity - snapshot.quantity
        if remaining_units <= 0:
            remaining = [item for item in remaining if item.index != source.index]
        else:
            total_qty = source.total_qty or source.quantity
            total_amount = source.total_amount or source.amount
            source.quantity = remaining_units
            source.amount = int(
                (Decimal(total_amount) * remaining_units / total_qty).to_integral_value()
            )
        break
    return (used, remaining)


def _solve_within_suppliers(
    target: Target,
    candidates: list[Line],
    original_pool: list[Line],
    tolerance: int,
    *,
    allow_fraction: bool = False,
) -> tuple[list[Line], list[Line], str] | None:
    """大池子 DP 未命中时的兜底：按「供应商」切小，再走常规路径。

    大池子只能按「元」粒度枚举整行组合，看不到两类真实存在的解：
    一是报关时少报/多报几支（部分出运数量配平），二是同供应商整包。
    按供应商分组后每组规模可控，就能用上完整的求解 + 部分配平逻辑。
    只有恰好一个供应商组能闭合时才采纳，命中多组按多解处理。
    """
    if target.capacity <= 0:
        return None
    groups: dict[str, list[Line]] = {}
    for line in candidates:
        groups.setdefault(line.supplier or "", []).append(line)
    hits: list[tuple[list[Line], list[Line], str]] = []
    for group in groups.values():
        if not group:
            continue
        hit = _try_group_solution(
            target, group, original_pool, tolerance, allow_fraction=allow_fraction
        )
        if hit is not None:
            hits.append(hit)
            if len(hits) > 1:
                return None
    if len(hits) != 1:
        return None
    return hits[0]


def _try_group_solution(
    target: Target,
    group: list[Line],
    original_pool: list[Line],
    tolerance: int,
    *,
    allow_fraction: bool = False,
) -> tuple[list[Line], list[Line], str] | None:
    """在一个供应商组内求解：整行子集和优先，其次部分出运数量配平。"""
    row = target.rows[0] if target.rows else {}
    row_hs = str(row.get("海关编码") or row.get("hs") or "")
    ordered = sorted(group, key=lambda line: _line_priority(line, row, row_hs))
    single_supplier = len({line.supplier for line in group}) == 1
    supplier_name = ordered[0].supplier if ordered else ""

    if len(ordered) > MAX_ITEMS:
        # 单个供应商的池子仍然太大：用元粒度 DP（命中后仍按「分」精确校验）
        picked = subset_sum_bounded(
            [(line.index, line.amount) for line in ordered],
            target.capacity,
            tolerance,
        )
        if not picked:
            return None
        chosen = set(picked)
        used = [line for line in ordered if line.index in chosen]
        if not used or equivalent_swap_exists(used, original_pool):
            return None
        remaining = [line for line in original_pool if line.index not in chosen]
        return (used, remaining, f"按供应商（{supplier_name}）整包闭合")

    items = [(line.index, line.amount) for line in ordered]
    solutions = subset_sum_solutions(items, target.capacity, tolerance, limit=2)
    if solutions:
        if len(solutions) > 1 and not solutions_equivalent(solutions, ordered):
            return None
        chosen = set(solutions[0])
        used = [line for line in ordered if line.index in chosen]
        if equivalent_swap_exists(used, original_pool):
            return None
        remaining = [line for line in original_pool if line.index not in chosen]
        label = f"按供应商（{supplier_name}）" if single_supplier else "按供应商分组"
        return (used, remaining, f"{label}整包闭合")

    partials = _try_partial(
        target, ordered, tolerance, limit=2, allow_fraction=allow_fraction
    )
    if len(partials) != 1:
        return None
    used_lines, consumed, partial_line = partials[0]
    used_units = Decimal(str(consumed))
    consumed_amount = int(partial_line.unit_price * used_units)
    snapshot = replace(
        partial_line,
        amount=consumed_amount,
        quantity=used_units,
        members=[],
        splittable=False,
    )
    assignment = [snapshot if line is partial_line else line for line in used_lines]
    if equivalent_swap_exists(assignment, original_pool):
        return None
    removed = {line.index for line in used_lines if line is not partial_line}
    remaining_qty = partial_line.quantity - used_units
    if remaining_qty <= 0:
        removed.add(partial_line.index)
    else:
        total_qty = partial_line.total_qty or partial_line.quantity
        total_amount = partial_line.total_amount or partial_line.amount
        partial_line.quantity = remaining_qty
        partial_line.amount = int(
            (Decimal(total_amount) * remaining_qty / total_qty).to_integral_value()
        )
    remaining = [line for line in original_pool if line.index not in removed]
    return (
        assignment,
        remaining,
        f"按供应商（{supplier_name}）分组、由部分出运数量配平（占用 {consumed:.2f} 个单位）",
    )


def _try_partial(
    target: Target,
    ordered: list[Line],
    tolerance: int,
    *,
    max_scan: int = 24,
    limit: int = 1,
    probes: int = 60_000,
    fraction_index: int | None = None,
    allow_fraction: bool = False,
) -> list[tuple[list[Line], float, Line]]:
    """兜底：允许多用/少用一条产品行的一部分出运数量来配平。

    做法是「反着解」：先枚举其它整行能凑出的所有金额（meet-in-the-middle），
    再按「外销单价必须整除剩余金额」筛出合法的整数单位数，
    避免逐个单位数去试子集和（数量上万时会非常慢）。

    返回最多 `limit` 组不同解：`limit=2` 时能发现「同一张报关单有多种
    部分配平方式」的多解情况。

    `probes` 是整次搜索的探测预算，避免在「证明确实没有解」上把整轮跑挂住。
    """
    if target.capacity <= 0:
        return []
    # 「可分数单位」（米 / 千克 / 吨…）允许按小数配平——管材按米、线材按千克出运，
    # 一张出运单拆到两张报关单时本来就不是整数。但条件极严：
    #   ① **当前候选池里只能有一条这样的记录**（多条就无从判断余量归谁，
    #      用户口径：多个以米为单位的候选项时禁止分数匹配）；
    #   ② **只在整数配平找不到解时才算**——分数解能和整数解并存，
    #      若并列返回会被当成「多解」，把本来拆得出来的单子判死
    #      （26MT-05G019A 就这样被误判过）；
    #   ③ 其余候选不超过 FRACTION_MAX_OTHERS 条。
    integer_hits = _partial_scan(target, ordered, tolerance, max_scan, limit, probes, None)
    if integer_hits:
        return integer_hits
    if not allow_fraction:
        return []
    fraction_lines = [line for line in ordered if is_fraction_unit(line)]
    fraction_index = fraction_lines[0].index if len(fraction_lines) == 1 else None
    if fraction_index is None:
        return []
    return _partial_scan(
        target, ordered, tolerance, max_scan, limit, probes, fraction_index
    )


def _partial_scan(
    target: Target,
    ordered: list[Line],
    tolerance: int,
    max_scan: int,
    limit: int,
    probes: int,
    fraction_index: int | None,
) -> list[tuple[list[Line], float, Line]]:
    """`_try_partial` 的实际搜索：`fraction_index` 非空时才允许那一条按小数配平。"""
    found: list[tuple[list[Line], float, Line]] = []
    seen: set[tuple] = set()
    capacity = target.capacity
    for line in ordered[:max_scan]:
        if not line.splittable:
            continue
        if line.unit_price <= 0 or line.quantity <= 0:
            continue
        unit_price = line.unit_price
        others = [other for other in ordered if other.index != line.index]
        # 可分数解只在候选很少时才可信：候选一多，能凑出余量的组合就会成群，
        # 唯一性无从谈起，索性退回整数口径。
        fractional = (
            fraction_index is not None
            and line.index == fraction_index
            and len(others) <= FRACTION_MAX_OTHERS
        )
        half = len(others) // 2
        left = _half_sums([(other.index, other.amount) for other in others[:half]])
        right = _half_sums([(other.index, other.amount) for other in others[half:]])
        right_items = sorted(right.items())
        buckets: dict[int, list[int]] = {}
        for position, (right_sum, _pick) in enumerate(right_items):
            buckets.setdefault(right_sum % unit_price, []).append(position)
        max_units = int(line.quantity)
        for left_sum, left_pick in left.items():
            need = capacity - left_sum
            if need < 0:
                continue
            # 分数模式下余量可以是任意小数，不能按「整数余数」分桶筛
            positions = (
                range(len(right_items)) if fractional else buckets.get(need % unit_price, ())
            )
            for position in positions:
                if probes <= 0:
                    return found
                probes -= 1
                right_sum, right_pick = right_items[position]
                delta = need - right_sum
                if delta < 0:
                    continue
                if fractional:
                    exact = Decimal(delta) / Decimal(unit_price)
                    if not 0 < exact <= line.quantity:
                        continue
                    if abs(Decimal(delta) - exact * unit_price) > tolerance:
                        continue
                    options = (exact,)
                else:
                    base = int((Decimal(delta) / Decimal(unit_price)).to_integral_value())
                    options = (base, base + 1)
                for units in options:
                    if not fractional:
                        if not 1 <= units <= max_units:
                            continue
                        if abs(delta - units * unit_price) > tolerance:
                            continue
                    keep = tuple(sorted(left_pick + right_pick))
                    signature = (line.index, units, keep)
                    if signature in seen:
                        continue
                    seen.add(signature)
                    keep_set = set(keep)
                    used = [other for other in others if other.index in keep_set]
                    used.append(line)
                    found.append((used, float(units), line))
                    if len(found) >= limit:
                        return found
    return found


def equivalent_swap_exists(solution: list[Line], pool: list[Line]) -> bool:
    """解的等价性检查：背包内某条记录，在背包外是否存在**价格一致**的记录。

    两条记录只要出运金额一致（同价），互换后仍是同一组金额，
    业务上无法区分是哪一个，属于「候选不唯一」。
    """
    inside = {line.index for line in solution}
    amounts_inside = {line.amount for line in solution}
    for line in pool:
        if line.index in inside:
            continue
        if line.amount in amounts_inside:
            return True
    return False


def solutions_equivalent(
    solutions: list[list[int]],
    lines: list[Line],
) -> bool:
    """多组子集和是否**业务等价**。

    两组记录只要供应商、采购单、出运金额完全一致，最终拆单结果就一样
    （例如同厂家同金额的重复行），不该判成「候选不唯一」。
    """
    if len(solutions) <= 1:
        return True
    by_index = {line.index: line for line in lines}
    signatures = set()
    for solution in solutions:
        picked = [by_index[i] for i in solution if i in by_index]
        signatures.add(
            tuple(
                sorted(
                    (line.supplier, line.purchase_code, line.amount) for line in picked
                )
            )
        )
    return len(signatures) == 1


# 单次精确求解的最大候选数（2^18 规模，实测很快；再大先用海关编码收窄）
MAX_ITEMS = 36

# 可分数单位：长度 / 重量 / 面积这类按小数计量的单位。
# 其它单位（个 / 支 / 件 / 套 / EA / PC / PCS…）必须整数——半个法兰没有意义。
FRACTION_UNITS = (
    "米", "M(", "MTR", "METER", "METRE", "FT(", "INCH", "英寸",
    "KG", "千克", "公斤", "TON", "吨", "LB(",
)
# 启用分数配平时允许的「其余候选」上限：候选越多越容易凑出假解
FRACTION_MAX_OTHERS = 12


def is_fraction_unit(line: Line) -> bool:
    """这条产品行的单位能不能按小数配平（米 / 千克 / 吨 …）。

    单位取自出运产品行的「计量单位」；取不到（历史缓存里常是空的）一律当整数处理，
    不给分数匹配开口子。
    """
    unit = str((line.raw or {}).get("unit") or "").strip()
    if not unit:
        return False
    text = unit.upper()
    return any(token in text for token in FRACTION_UNITS)


def _narrow(target: Target, pool: list[Line]) -> list[Line]:
    """用目标报关行的海关编码/供应商把候选行收窄。"""
    hs_codes = {str(r.get("hs") or "").strip() for r in target.rows}
    hs_codes.discard("")
    narrowed = [line for line in pool if line.hs_code in hs_codes] if hs_codes else []
    if narrowed:
        return narrowed
    names = {str(r.get("name") or "").strip() for r in target.rows}
    names.discard("")
    narrowed = (
        [line for line in pool if any(name_matches(line, name) for name in names)]
        if names
        else []
    )
    return narrowed or pool


def subset_sum_bounded(
    items: list[tuple[int, int]],
    target: int,
    tolerance: int = 0,
) -> list[int] | None:
    """容量受限的子集和（按「元」粒度的位图 DP），用于产品行很多的大池子。

    以元为单位把状态空间压到「目标金额」量级（几十万位），
    位图运算在 Python 大整数上是按机器字并行做的，比逐状态字典 DP 快两个数量级。
    命中的和再用「分」精确校验，校验不过视为该粒度无解。

    粒度由细到粗自动选取：目标金额不大时直接按「分」做，保证精确；
    金额很大时才退到「元」。**每条记录取整都会漂移最多半个粒度**，
    因此窗口必须按行数放宽，否则「整包正好等于容量」这类解会被漏掉
    （实测 66 条记录按元取整累计漂移 1.99 元，正好掉在 ±0.5 元窗口外）。
    """
    if target <= 0 or not items:
        return [] if target <= 0 else None
    item_count = max(len(items), 1)
    scale = 100
    for candidate_scale in (1, 10, 100, 1000, 10_000):
        if (((target + tolerance) // candidate_scale) + 1) * item_count <= 4_000_000_000:
            scale = candidate_scale
            break
    # 取整漂移：每条记录最多 scale/2，行数条累计后仍要让真解留在窗口里
    slack = scale * (item_count // 2 + 2)
    low = max(0, target - tolerance - slack) // scale
    high = (target + tolerance + slack) // scale
    if high <= 0:
        return None
    limit_mask = (1 << (high + 1)) - 1
    bits = 1
    snapshots: list[int] = []
    picked_items: list[tuple[int, int, int]] = []
    for index, amount in items:
        yuan = int(round(amount / scale))
        if yuan <= 0:
            continue
        snapshots.append(bits)
        picked_items.append((index, yuan, amount))
        bits |= (bits << yuan) & limit_mask

    window = bits & (((1 << (high - low + 1)) - 1) << low) if high >= low else 0
    if not window:
        return None
    totals: list[int] = []
    while window:
        lowest = window & -window
        totals.append(lowest.bit_length() - 1)
        window ^= lowest
    totals.sort(key=lambda total: abs(total * scale - target))
    amount_of = {index: amount for index, _yuan, amount in picked_items}
    for total in totals[:24]:
        chosen: list[int] = []
        remaining = total
        for position in range(len(picked_items) - 1, -1, -1):
            index, yuan, _amount = picked_items[position]
            if remaining >= yuan and (snapshots[position] >> (remaining - yuan)) & 1:
                chosen.append(index)
                remaining -= yuan
        if remaining != 0:
            continue
        current = sum(amount_of[i] for i in chosen)
        if abs(current - target) <= tolerance:
            return chosen
        # 元粒度命中、分精度差一点点：做一次局部修补（加一条 / 去一条 / 换一条），
        # 把差额拉回容差内。大池子里这种"差几角几分"是最常见的假无解。
        repaired = _repair_subset(chosen, items, target, tolerance)
        if repaired is not None:
            return repaired
    return None


def _repair_subset(
    chosen: list[int],
    items: list[tuple[int, int]],
    target: int,
    tolerance: int,
) -> list[int] | None:
    """对子集和做一次局部修补：加一条、去一条、或换一条。"""
    if tolerance <= 0:
        return None
    amount_of = {index: amount for index, amount in items}
    inside = set(chosen)
    current = sum(amount_of[i] for i in inside)
    diff = target - current
    outside = sorted(
        ((amount, index) for index, amount in items if index not in inside),
        key=lambda pair: pair[0],
    )
    if not outside:
        return None
    amounts_out = [amount for amount, _index in outside]

    def nearest(want: int) -> tuple[int, int] | None:
        position = bisect.bisect_left(amounts_out, want)
        best = None
        for candidate in (position - 1, position):
            if 0 <= candidate < len(outside):
                amount, index = outside[candidate]
                if best is None or abs(amount - want) < abs(best[0] - want):
                    best = (amount, index)
        return best

    # 1) 直接补一条
    hit = nearest(diff)
    if hit and abs(diff - hit[0]) <= tolerance:
        return chosen + [hit[1]]
    # 2) 去掉一条
    inside_sorted = sorted(((amount_of[i], i) for i in inside), key=lambda pair: pair[0])
    amounts_in = [amount for amount, _index in inside_sorted]
    position = bisect.bisect_left(amounts_in, -diff)
    for candidate in (position - 1, position):
        if 0 <= candidate < len(inside_sorted):
            amount, index = inside_sorted[candidate]
            if abs(diff + amount) <= tolerance:
                return [i for i in chosen if i != index]
    # 3) 换一条：inside 的 a 换成 outside 的 b，要求 b - a ≈ diff
    for amount_in, index_in in inside_sorted:
        hit = nearest(amount_in + diff)
        if hit and abs(diff - (hit[0] - amount_in)) <= tolerance:
            return [i for i in chosen if i != index_in] + [hit[1]]
    return None


def iter_lines(rows: Iterable[dict[str, Any]]) -> list[Line]:
    """把出运产品行（索引里的 dict）转成 Line。"""
    out: list[Line] = []
    for index, row in enumerate(rows):
        price = to_cents(row.get("amount_usd"))
        quantity = Decimal(0)
        raw_qty = str(row.get("quantity") or "0").replace(",", "").strip()
        try:
            quantity = Decimal(raw_qty) if raw_qty else Decimal(0)
        except Exception:  # noqa: BLE001
            quantity = Decimal(0)
        unit = 0
        if quantity > 0:
            unit = int((Decimal(price) / quantity).to_integral_value())
        out.append(
            Line(
                index=index,
                amount=price,
                supplier=str(row.get("supplier") or ""),
                hs_code=str(row.get("hs_code") or ""),
                customs_name=str(row.get("customs_name") or ""),
                product_type=str(row.get("product_type") or ""),
                unit_price=unit,
                quantity=quantity,
                purchase_code=str(row.get("purchase_code") or ""),
                raw=dict(row),
                total_qty=quantity,
                total_amount=price,
            )
        )
    return out


def bundle_lines(
    lines: list[Line],
    key_fn,
    *,
    label: str = "",
) -> list[Line]:
    """按 `key_fn(line)` 把产品行打包成「组合记录」。

    先验：不同工厂的交期与所在地不同，所以一张报关单通常对应
    **某个工厂的全部货物**，或**若干工厂的全部货物**。
    把同工厂（可选再叠加同产品类型）的行合并成一条大记录，
    能大幅缩小搜索空间并消除大量"等价多解"。
    """
    groups: dict[Any, list[Line]] = {}
    for line in lines:
        groups.setdefault(key_fn(line), []).append(line)
    out: list[Line] = []
    for index, (key, members) in enumerate(sorted(groups.items(), key=lambda x: str(x[0]))):
        out.append(
            Line(
                index=index,
                amount=sum(m.amount for m in members),
                supplier=members[0].supplier,
                hs_code=members[0].hs_code,
                customs_name=members[0].customs_name,
                product_type=line_product_type(members[0]),
                purchase_code="、".join(sorted({m.purchase_code for m in members if m.purchase_code})),
                unit_price=0,
                quantity=Decimal(0),
                raw={"_bundle": str(key), "_label": label},
                splittable=False,
                members=members,
            )
        )
    return out


# 费用分摊口径。`none` = 这笔费用**不分配给任何报关单**（即不参与报关金额闭合）。
# 临时诊断开关：`SPLIT_NO_FEE_NONE=1` 时禁用 `none`，用来审计有多少数据依赖它。
FREIGHT_MODES = ("equal", "by_weight", "by_amount", "none")
if os.environ.get("SPLIT_NO_FEE_NONE") == "1":  # noqa: SIM108
    FREIGHT_MODES = ("equal", "by_weight", "by_amount")


def allocate_fee(
    declarations: dict[str, list[dict[str, Any]]],
    amount: int,
    mode: str,
) -> dict[str, int]:
    """把**一笔**费用按指定口径分到各张报关单。"""
    decls = list(declarations)
    if not decls or amount <= 0 or mode == "none":
        return {decl: 0 for decl in decls}

    if mode == "equal":
        share = amount // len(decls)
        shares = {decl: share for decl in decls}
        rest = amount - share * len(decls)
        if rest:
            biggest = max(
                decls,
                key=lambda d: sum(to_cents(r.get("amount")) for r in declarations[d]),
            )
            shares[biggest] += rest
        return shares

    def weight_sum(rows: list[dict[str, Any]]) -> int:
        total = 0
        for row in rows:
            text = str(row.get("weight") or "0").replace(",", "").strip()
            try:
                total += int(round(float(text) * 100))
            except ValueError:
                continue
        return total

    if mode == "by_weight":
        keys = {decl: max(weight_sum(rows), 0) for decl, rows in declarations.items()}
    else:
        keys = {
            decl: sum(to_cents(r.get("amount")) for r in rows)
            for decl, rows in declarations.items()
        }
    total = sum(keys.values())
    if total <= 0:
        return {decl: 0 for decl in decls}
    shares: dict[str, int] = {}
    left = amount
    for index, decl in enumerate(decls):
        if index == len(decls) - 1:
            shares[decl] = left
        else:
            part = int(amount * keys[decl] / total)
            shares[decl] = part
            left -= part
    return shares


def solve_residual(
    declarations: dict[str, list[dict[str, Any]]],
    items: list[Line],
    tolerance: int = 0,
) -> tuple[SolveResult, int]:
    """容量 = 报关金额、**允许残差**的分配：把整组产品行分给各张报关单。

    适用场景：一张出运单对应多张报关单，而客户费用（检验费/运费等）在飞书侧
    并没有按某个固定口径摊到每张报关单上——例如 25MT-07F521Y：

        报关金额合计 318220.31 = 产品行 299384.00 + 客户费用 18836.31
        4 张报关单里 2 张与产品行精确闭合，另 2 张分别带 3836.31 / 15000.00 的残差，
        残差之和恰好是客户费用总额。

    因此这里不再预先猜费用分摊口径，而是要求：

    - 每组产品行整体分给某一张报关单；
    - 每张报关单分到的产品金额 ≤ 它的报关金额；
    - 所有报关单都要分到东西。

    在这些约束下取「残差平方和最小」的解；若最优解不唯一则判多解。
    返回 (SolveResult, 精确闭合的报关单数)。
    """
    result = SolveResult()
    decls = list(declarations)
    if not decls or not items:
        result.unassigned = list(items)
        return result, 0
    caps = {d: sum(to_cents(r.get("amount")) for r in declarations[d]) for d in decls}
    amounts = [line.amount for line in items]
    if sum(amounts) > sum(caps.values()) + tolerance:
        result.unassigned = list(items)
        return result, 0

    order = sorted(range(len(items)), key=lambda index: -amounts[index])
    used: dict[str, int] = {d: 0 for d in decls}
    assign = [-1] * len(items)
    solutions: list[list[int]] = []
    budget = [2_000_000]

    if len(items) > 18:
        result.unassigned = list(items)
        return result, 0

    def dfs(position: int) -> None:
        budget[0] -= 1
        if budget[0] <= 0:
            return
        if len(solutions) >= 2:
            return
        if position == len(order):
            solutions.append(list(assign))
            return
        index = order[position]
        for decl_index, decl in enumerate(decls):
            if used[decl] + amounts[index] <= caps[decl] + tolerance:
                assign[index] = decl_index
                used[decl] += amounts[index]
                dfs(position + 1)
                used[decl] -= amounts[index]
                assign[index] = -1
                if len(solutions) >= 2:
                    return

    dfs(0)
    if not solutions:
        result.unassigned = list(items)
        return result, 0

    def score(solution: list[int]) -> float:
        residual = {d: caps[d] for d in decls}
        for index, decl_index in enumerate(solution):
            if decl_index >= 0:
                residual[decls[decl_index]] -= amounts[index]
        return sum(value * value / max(caps[d], 1) for d, value in residual.items())

    scored = sorted(((score(sol), sol) for sol in solutions), key=lambda pair: pair[0])
    best = scored[0]
    if len(scored) > 1 and abs(scored[1][0] - best[0]) <= 1e-9:
        result.notes.append("残余匹配多解（至少两组分配方式的残差一致），候选不唯一")
        result.assignments = {}
        result.unassigned = list(items)
        return result, 0

    grouped: dict[str, list[Line]] = {d: [] for d in decls}
    for index, decl_index in enumerate(best[1]):
        if decl_index >= 0:
            grouped[decls[decl_index]].append(items[index])
    if any(not grouped[d] for d in decls):
        result.unassigned = list(items)
        return result, 0

    exact = 0
    for decl in decls:
        assigned_sum = sum(line.amount for line in grouped[decl])
        residual = caps[decl] - assigned_sum
        if abs(residual) <= tolerance:
            exact += 1
        else:
            result.notes.append(
                f"{decl}: 按报关金额分配，扣产品行后残差 {residual / 100:.2f}（视为客户费用）"
            )
    result.assignments = grouped
    return result, exact


def assign_to_rows(
    rows: list[dict[str, Any]],
    lines: list[Line],
    supplier_short: "callable | None" = None,
) -> dict[int, list[Line]]:
    """把某张报关单已装入的产品行，分配到它的各个报关行。

    报关行的供应商来自 PDF 解析（通常为空），所以真正能用的证据是
    **海关编码 / 报关品名 / 金额**。这里按「金额从大到小」逐行落位，
    每次挑得分最高的报关行：供应商（若有）> 海关编码 > 品名，
    再用「该行已落金额 + 当前产品行金额」与报关行报关金额的偏差做惩罚。

    关键点：同一张报关单的两行常常共用一个海关编码（如法兰/管件都是 7507200000），
    早先的实现会让第一行把同编码的产品行全部吃掉，第二行颗粒无收；
    改成金额逼近后，两行会各拿到与自身报关金额贴合的那批产品行。
    """
    if not rows:
        return {}
    if len(rows) == 1:
        return {0: list(lines)}

    short_of = supplier_short or (lambda name: str(name or ""))
    by_bundle = _assign_to_rows_by_bundle(rows, lines, short_of)
    if by_bundle is not None:
        return by_bundle
    legacy = _assign_to_rows_by_attribute(rows, lines, short_of)
    if all(legacy[index] for index in range(len(rows))):
        # 老口径已经每行都有产品行，保持原有结果（避免动到已拆对的单）
        return legacy
    return _repair_empty_rows(rows, legacy)


def _assign_to_rows_by_bundle(
    rows: list[dict[str, Any]],
    lines: list[Line],
    short_of,
) -> dict[int, list[Line]] | None:
    """按「采购单（供应商）」整组对号入座；只有整组放不下时才拆开按金额分摊。

    一张报关单里的几行常常正好对应出运单里的几家工厂：

        26MT-05Q171：溪流 72854.16 / 钟兴 4108.80 / 勇恒 31179.50 / 鸿迪 32679.50
        报关行金额：72854.16 / 4108.80 / 31179.50 / 32679.50   —— 一一对应

    逐条产品行去凑会把这些组打散，出现「同一家的货串到别的报关行」。
    返回 None 表示这条路走不通（有报关行分不到货），交给后面的兜底逻辑。
    """
    targets = [to_cents(row.get("amount")) for row in rows]
    if not lines or any(target <= 0 for target in targets):
        return None
    groups: dict[str, list[Line]] = {}
    for line in lines:
        # 同一家工厂可能有多张采购单（如 26MT-06N145 与 26MT-06N145-ADD-HD 都是鸿迪），
        # 报关行认的是「工厂」，所以按供应商简称归组，采购单号只作兜底键。
        key = short_of(line.supplier) or line.purchase_code or "(未知)"
        groups.setdefault(key, []).append(line)
    row_hs = [str(row.get("hs") or "").strip() for row in rows]
    row_name = [str(row.get("name") or "").strip() for row in rows]
    pending = [(key, members) for key, members in groups.items()]
    open_rows = set(range(len(rows)))
    result: dict[int, list[Line]] = {index: [] for index in range(len(rows))}

    def compatible(members: list[Line], index: int) -> bool:
        """整组货是否与这一报关行的海关编码/品名相符。"""
        if row_hs[index] and all(line.hs_code == row_hs[index] for line in members):
            return True
        if row_name[index] and any(name_matches(line, row_name[index]) for line in members):
            return True
        return False

    # 1) 整组配对：品名/海关编码相符的优先，其次要求金额几乎完全吻合
    while pending and open_rows:
        best: tuple[tuple[int, float], int, int] | None = None
        for position, (_key, members) in enumerate(pending):
            total = sum(line.amount for line in members)
            for index in open_rows:
                error = abs(total - targets[index]) / targets[index]
                if error <= 0.02:
                    rank = (0, error)
                elif compatible(members, index) and error <= 0.5:
                    rank = (1, error)
                else:
                    continue
                if best is None or rank < best[0]:
                    best = (rank, position, index)
        if best is None:
            break
        _rank, position, index = best
        _key, members = pending.pop(position)
        result[index] = members
        open_rows.discard(index)

    # 2) 剩下的组只能拆开：按金额贴合度逐条分给仍空着的报关行
    leftover_lines = [line for _key, members in pending for line in members]
    if leftover_lines and not open_rows:
        return None
    if leftover_lines:
        filled = {index: sum(line.amount for line in result[index]) for index in open_rows}
        for line in sorted(leftover_lines, key=lambda item: -item.amount):
            index = min(
                open_rows,
                key=lambda i: abs(filled[i] + line.amount - targets[i]),
            )
            result[index].append(line)
            filled[index] += line.amount
    if any(not result[index] for index in range(len(rows))):
        return None
    return result


def _repair_empty_rows(
    rows: list[dict[str, Any]],
    assignment: dict[int, list[Line]],
) -> dict[int, list[Line]]:
    """有报关行没分到产品行时，从别的行「借」最贴合的产品行。

    只在该行与出借行的**报关金额偏差合计变小**时才搬，保证不会把已拆对的单搞坏。
    """
    targets = [to_cents(row.get("amount")) for row in rows]
    result = {index: list(assignment.get(index, [])) for index in range(len(rows))}

    def deficit(index: int) -> int:
        return abs(targets[index] - sum(line.amount for line in result[index]))

    def relative(index: int) -> float:
        """相对偏差：报关金额大的行容忍的绝对差也大，避免"总差守恒"时无法比较。"""
        target = targets[index]
        return deficit(index) / target if target else float(deficit(index))

    for _ in range(len(rows) * 3):
        empty = [index for index in range(len(rows)) if not result[index]]
        if not empty:
            break
        best_move: tuple[int, int, int, Line] | None = None
        for target_index in empty:
            for donor in range(len(rows)):
                if donor == target_index or len(result[donor]) <= 1:
                    continue
                for line in result[donor]:
                    donor_sum = sum(x.amount for x in result[donor])
                    donor_after = abs(targets[donor] - (donor_sum - line.amount))
                    target_after = abs(targets[target_index] - line.amount)
                    donor_rel = donor_after / targets[donor] if targets[donor] else donor_after
                    target_rel = (
                        target_after / targets[target_index] if targets[target_index] else target_after
                    )
                    gain = relative(target_index) + relative(donor) - target_rel - donor_rel
                    if gain > 0 and (best_move is None or gain > best_move[0]):
                        best_move = (gain, target_index, donor, line)
        if best_move is None:
            break
        _gain, target_index, donor, line = best_move
        result[donor] = [x for x in result[donor] if x is not line]
        result[target_index].append(line)
    return result


def _assign_to_rows_by_attribute(
    rows: list[dict[str, Any]],
    lines: list[Line],
    short_of,
) -> dict[int, list[Line]]:
    """旧口径：供应商 → 海关编码 → 报关品名 → 金额最近兜底。"""
    result: dict[int, list[Line]] = {i: [] for i in range(len(rows))}
    remaining = list(lines)

    for index, row in enumerate(rows):
        want = short_of(row.get("supplier") or "")
        if not want:
            continue
        picked = [line for line in remaining if short_of(line.supplier) == want]
        if picked:
            result[index].extend(picked)
            remaining = [line for line in remaining if line not in picked]

    for index, row in enumerate(rows):
        if result[index] or not remaining:
            continue
        want_hs = str(row.get("hs") or "").strip()
        if not want_hs:
            continue
        picked = [line for line in remaining if line.hs_code == want_hs]
        if picked:
            result[index].extend(picked)
            remaining = [line for line in remaining if line not in picked]

    for index, row in enumerate(rows):
        if result[index] or not remaining:
            continue
        want_name = str(row.get("name") or "").strip()
        if not want_name:
            continue
        picked = [line for line in remaining if name_matches(line, want_name)]
        if picked:
            result[index].extend(picked)
            remaining = [line for line in remaining if line not in picked]

    # 兜底：按报关金额贴合度落行，并对「同一采购单」加亲和度——
    # 同一家工厂的产品行应当落在同一报关行，避免把螺丝行的尾差丢给阀门行造成供应商串味
    dominant = [""] * len(rows)
    for line in remaining:
        best_index = 0
        best_score: float | None = None
        for index in range(len(rows)):
            target = to_cents(rows[index].get("amount"))
            total = sum(x.amount for x in result[index]) + line.amount
            ratio = abs(total - target) / target if target else 1.0
            score = -min(ratio, 10.0) * 1000
            if dominant[index] and line.purchase_code and dominant[index] == line.purchase_code:
                score += 800
            if best_score is None or score > best_score:
                best_index, best_score = index, score
        result[best_index].append(line)
        if line.purchase_code and not dominant[best_index]:
            dominant[best_index] = line.purchase_code
    return result


def _assign_to_rows_by_amount(
    rows: list[dict[str, Any]],
    lines: list[Line],
    short_of,
) -> dict[int, list[Line]]:
    """金额逼近口径：逐条落位，用报关金额偏差打分（同编码/同品名加分）。"""
    targets = [to_cents(row.get("amount")) for row in rows]
    row_hs = [str(row.get("hs") or "").strip() for row in rows]
    row_name = [str(row.get("name") or "").strip() for row in rows]
    row_supplier = [short_of(row.get("supplier") or "") for row in rows]
    result: dict[int, list[Line]] = {index: [] for index in range(len(rows))}
    filled = [0] * len(rows)

    for line in sorted(lines, key=lambda item: item.amount, reverse=True):
        best_index = 0
        best_score: float | None = None
        for index in range(len(rows)):
            score = 0.0
            if row_supplier[index] and line.supplier:
                if short_of(line.supplier) == row_supplier[index]:
                    score += 10_000
            if row_hs[index] and line.hs_code and row_hs[index] == line.hs_code:
                score += 1_000
            if row_name[index] and name_matches(line, row_name[index]):
                score += 500
            target = targets[index]
            deficit = abs(target - (filled[index] + line.amount))
            ratio = deficit / target if target else 1.0
            score -= min(ratio, 10.0) * 100
            if best_score is None or score > best_score:
                best_index, best_score = index, score
        result[best_index].append(line)
        filled[best_index] += line.amount
    return result
