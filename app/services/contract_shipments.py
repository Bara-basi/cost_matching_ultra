"""合同号 → 出运单定位（拆单取数的第一步）。

报关单上的「合同协议号」是出运单号的联合写法，遵循**最小重复原则**：

    26MT-06H124&08C122              -> 26MT-06H124 / 26MT-08C122
    26MT-01S180&190&241             -> 26MT-01S180 / 26MT-01S190 / 26MT-01S241
    26MT-06H251&ADD1                -> 26MT-06H251 / 26MT-06H251-ADD1
    26MT-01T050D&ADD1&080B&ADD1-B   -> 26MT-01T050D / 26MT-01T050ADD1
                                       / 26MT-01T080B / 26MT-01T080ADD1-B

一张报关单可以对应**多张出运单**（如 `26MT-06H124&08C122` 对应两张），
所以不能把合同号逐段当成订单去查，而是：

1. **整串精确匹配**：合同号（去掉 `PI-` 与分隔符）与出运单的
   invoiceCode / orderCode / 采购单号一致（含 `&` 联合出运单）；
2. **订单核心覆盖**：合同号与出运单都归一到「订单核心」——年份+MT+部门+业务员+三位流水
   （可含 `Y`）加可选的 `-ADDn`，批次/工厂等尾缀一律丢弃；
   再用**最少的出运单**覆盖合同号涉及的订单集合；
3. 覆盖出现多解或覆盖不全时不猜，退回旧的按基号模糊匹配。
"""
from __future__ import annotations

import re
from typing import Any

from app.services.erp_cache import CACHE_ROOT, read_jsonl
from app.services.shipment_index import canonical, strip_pi

CORE_RE = re.compile(r"^(\d{2})MT-(\d{2})([A-Z])(\d{3})(Y?)", re.IGNORECASE)
HEAD_RE = re.compile(r"^(\d{2}MT-\d{2}[A-Z])", re.IGNORECASE)
ADD_RE = re.compile(r"ADD\d*", re.IGNORECASE)
TOKEN_SPLIT = re.compile(r"[&,;、；]")


def expand_contract(text: Any) -> list[str]:
    """按最小重复原则把联合合同号展开成单个订单号。

    `26MT-06H124&08C122`   -> [26MT-06H124, 26MT-08C122]
    `26MT-06M052&087A&112` -> [26MT-06M052, 26MT-06M087A, 26MT-06M112]
    `26MT-01S180&190&241`  -> [26MT-01S180, 26MT-01S190, 26MT-01S241]
    `26MT-06H251&ADD1`     -> [26MT-06H251, 26MT-06H251-ADD1]
    """
    text = strip_pi(text)
    if not text:
        return []
    parts = [p for p in TOKEN_SPLIT.split(text) if p]
    year = text[:2]
    head = ""
    previous = ""
    out: list[str] = []
    for part in parts:
        if re.match(r"^\d{2}MT-", part):
            full = part
        elif re.match(r"^\d{2}[A-Z]\d{3}", part):
            # 换了一个订单：`&08C122` -> 26MT-08C122
            full = f"{year}MT-{part}"
        elif re.match(r"^\d{3}", part) and head:
            # 同订单的流水续写：`&190` -> 26MT-01S190
            full = f"{head}{part}"
        elif previous:
            # ADD 等尾缀：挂在上一段后面
            full = f"{previous}-{part}"
        else:
            full = part
        match = HEAD_RE.match(full)
        if match:
            head = match.group(1).upper()
        out.append(full)
        previous = full
    return list(dict.fromkeys(out))


def order_core(code: Any) -> str:
    """订单核心：去掉批次/工厂等尾缀，只留 `年份+MT+部门+业务员+三位流水[Y][-ADDn]`。"""
    text = strip_pi(code)
    match = CORE_RE.match(text)
    if not match:
        return ""
    base = match.group(0).upper()
    rest = text[match.end() :].upper()
    add = ADD_RE.search(rest)
    return f"{base}-{add.group(0)}" if add else base


def cores_of(value: Any) -> set[str]:
    """一段单号（可能是联合写法）包含的订单核心集合。"""
    out: set[str] = set()
    for code in expand_contract(value):
        core = order_core(code)
        if core:
            out.add(core)
    return out


class ShipmentRef:
    """出运单列表里的一行（只保留定位用的字段）。"""

    __slots__ = ("invoice", "shipment_id", "cores", "keys")

    def __init__(self, invoice: str, shipment_id: str, cores: set[str], keys: set[str]) -> None:
        self.invoice = invoice
        self.shipment_id = shipment_id
        self.cores = cores
        self.keys = keys


_SHIPMENTS: list[ShipmentRef] | None = None


def _load_shipments() -> list[ShipmentRef]:
    out: list[ShipmentRef] = []
    for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
        invoice = str(row.get("invoiceCode") or "").strip()
        shipment_id = str(row.get("shipmentId") or "").strip()
        keys: set[str] = set()
        for field in ("invoiceCode", "orderCode"):
            value = str(row.get(field) or "")
            for token in re.split(r"[,;、]", value):
                token = token.strip()
                if token:
                    keys.add(canonical(token))
        purchase = str(row.get("purchaseCode") or "")
        for token in re.split(r"[,;、]", purchase):
            token = token.strip()
            if token:
                keys.add(canonical(token))
        # 覆盖关系以睿贝原生关联字段（orderCode / purchaseCode）为准。
        # invoiceCode 只当「整串命中」的键：同一张出运单被拆成两张记录时，
        # 旧记录的发票名可能仍带着另一个订单号（如 `26MT-06H278&08C282`
        # 实际只有 06H278 的行，08C282 已另立 `S260814001`）。
        cores = cores_of(str(row.get("orderCode") or "")) | cores_of(purchase)
        if not cores:
            cores = cores_of(invoice)
        if invoice or cores:
            out.append(ShipmentRef(invoice, shipment_id, cores, keys))
    return out


def shipments() -> list[ShipmentRef]:
    global _SHIPMENTS
    if _SHIPMENTS is None:
        _SHIPMENTS = _load_shipments()
    return _SHIPMENTS


def _explicit_hit(row: ShipmentRef, contract_key: str, tokens: set[str]) -> int:
    """整串/整段命中合同号的加分：整串同名最优先，其次命中任一展开段。"""
    score = 0
    if contract_key and contract_key in row.keys:
        score += 1000
    score += 100 * len(tokens & row.keys)
    return score


# 组合搜索上限：候选出运单数与组合规模都做保护，避免极复杂合同拖慢全量
MAX_COVER_CANDIDATES = 60
MAX_COVER_SIZE = 4
MAX_COVER_COMBOS = 400_000


def _target_mask(contract: Any) -> tuple[dict[str, int], int]:
    """订单核心 -> 位；返回 (位映射, 目标掩码)。"""
    target = sorted(cores_of(contract))
    bits = {core: 1 << i for i, core in enumerate(target)}
    mask = (1 << len(target)) - 1
    return bits, mask


def cover_shipments(contract: Any) -> tuple[list[ShipmentRef], bool]:
    """用**最少**的出运单覆盖合同号涉及的订单集合。

    返回 (出运单列表, 是否多解)：
    - 唯一最小覆盖 -> 返回该覆盖、多解=False；
    - 没有覆盖 / 存在多个等价的覆盖 -> 返回空列表，由调用方走旧的模糊兜底。

    等价覆盖的判定用「显式命中」（合同号整串或整段与出运单的发票号/订单号一致）优先，
    例如 `26MT-01T050D&...` 中 `26MT-01T050D` 明确命中 D 批次出运单，
    就不会被 B/C 批次的同名订单抢走。
    """
    text = str(contract or "")
    bits, target_mask = _target_mask(text)
    if not bits:
        return [], False
    contract_key = canonical(text)
    tokens = {canonical(t) for t in expand_contract(text)}
    tokens.discard("")

    pool: list[tuple[ShipmentRef, int, int]] = []
    for row in shipments():
        mask = 0
        for core in row.cores:
            if core in bits:
                mask |= bits[core]
        if mask:
            pool.append((row, mask, _explicit_hit(row, contract_key, tokens)))
    if not pool or len(pool) > MAX_COVER_CANDIDATES:
        return [], False

    from itertools import combinations

    combos = 0
    for size in range(1, min(MAX_COVER_SIZE, len(pool)) + 1):
        covers: list[tuple[tuple[ShipmentRef, ...], int]] = []
        for combo in combinations(pool, size):
            combos += 1
            if combos > MAX_COVER_COMBOS:
                return [], False
            mask = 0
            for _row, part, _bonus in combo:
                mask |= part
            if mask == target_mask:
                covers.append((tuple(row for row, _p, _b in combo), sum(b for _r, _p, b in combo)))
        if not covers:
            continue
        best_bonus = max(bonus for _combo, bonus in covers)
        best = [combo for combo, bonus in covers if bonus == best_bonus]
        signatures = {frozenset(row.invoice for row in combo) for combo in best}
        if len(signatures) == 1:
            return list(best[0]), False
        return [], True
    return [], False


def resolve_invoice_codes(contract: Any) -> list[str]:
    """合同号 → 出运发票号；定位不到或多解时返回空列表（由调用方走旧兜底）。"""
    by_map = resolve_by_split_map(contract)
    if by_map:
        return by_map
    chosen, ambiguous = cover_shipments(contract)
    if not chosen or ambiguous:
        return []
    return [row.invoice for row in chosen if row.invoice]


def resolve_each_part(contract: Any) -> list[str]:
    """多段合同 → 每一段各自定位自己的出运单（全部段都能各自定位才采用）。

    报关单经常把出运单**拆开写**（`26MT-01T050C&080A&080-ADD1` 对应三张出运单），
    而出运单列表里也确实各有一张同名出运单。逐段定位比「凑一张包含所有段的
    合并出运单」更安全：合并出运单里往往还挂着别的批次的货，会把无关工厂的
    产品行拉进来——实测合同 `26MT-01T050C&080A&080-ADD1` 被
    `PI-26MT-01T050ADD1&080B&080ADD1-B` 拉进了勇恒 26MT-01T080-YH /
    26MT-01T080-ADD1-YH 的货，而这三张报关单实际全是鸿迪生产的
    （飞书三行 47375.03 + 9089.78 + 2106.60 = 报关金额 58571.41，正好是三张
    同名出运单的出运总金额）。

    只有**每一段都能定位到恰好一张同名出运单**时才采纳；否则返回空，
    交给调用方走原来的映射表/覆盖逻辑。
    """
    parts = literal_parts(contract)
    if len(parts) < 2:
        return []
    by_invoice: dict[str, set[str]] = {}
    for row in shipments():
        if not row.invoice:
            continue
        by_invoice.setdefault(canonical(row.invoice), set()).add(row.invoice)
    out: list[str] = []
    for part in parts:
        hits = by_invoice.get(canonical(part)) or set()
        if len(hits) != 1:
            return []
        invoice = next(iter(hits))
        if invoice not in out:
            out.append(invoice)
    return out


def resolve_by_split_map(contract: Any) -> list[str]:
    """用「拆分结果 → 原值」映射表解析合同号。

    1. **整段命中**：合同号拆出来的订单集合与某张出运单完全一致，直接取那张出运单；
    2. **被包含命中**：合同号只是某张（联合）出运单的一段——这是「出运单合并了外销单、
       报关单又把它拆开」的情形，取**最小的**包含它的出运单；若有多张同样小，判多解不猜。

    例：报关合同 `26MT-10E228A` 被出运单 `26MT-10E185B&228A` 包含，
    而 `26MT-10E228B&281A&285A&348A` 的字面段是 `26MT-10E228B/281A/285A/348A`，
    与 `26MT-10E228A` 并不相同，所以不会被误匹配。
    """
    parts = set(literal_parts(contract))
    if not parts:
        return []
    mapping = split_map()
    exact = mapping.get(tuple(sorted(parts)))
    if exact:
        return [exact]
    # 「报关单把出运单拆开写」：每一段各自有自己的出运单时，按段取单，
    # 避免去凑合并出运单而把别的批次的货拉进来（见 resolve_each_part 注释）
    per_part = resolve_each_part(contract)
    if per_part:
        return per_part
    hits: list[tuple[int, str]] = []
    for key, original in mapping.items():
        if len(key) > len(parts) and parts <= set(key):
            hits.append((len(key) - len(parts), original))
    if not hits:
        return []
    smallest = min(extra for extra, _ in hits)
    chosen = list(dict.fromkeys(original for extra, original in hits if extra == smallest))
    return chosen if len(chosen) == 1 else []


# ---------------------------------------------------------------------------
# 出运单号 ↔ 外销单号 的对应表
# ---------------------------------------------------------------------------

_SPLIT_MAP: dict[tuple[str, ...], str] | None = None
_SPLIT_CONFLICTS: list[tuple[tuple[str, ...], str, str]] = []


def literal_parts(value: Any) -> tuple[str, ...]:
    """按最小重复原则字面拆开单号，**保留批次等尾缀**（`26MT-03R360A` ≠ `26MT-03R360B`）。"""
    return tuple(sorted({strip_pi(part) for part in expand_contract(value) if strip_pi(part)}))


def split_map() -> dict[tuple[str, ...], str]:
    """「拆分结果 → 原值」映射表（按你的口径：键是拆出来的订单集合，值是原始写法）。

    建表来源：睿贝出运单列表里**每张出运单的 invoiceCode**。
    键用字面拆分（保留批次尾缀），所以 `26MT-03R360A/B/C` 各占一个键，不会互相覆盖；
    同一张出运单被写了两遍（写法不同）时才会撞键。
    例如：

        ("26MT-10E185B", "26MT-10E228A")        -> "26MT-10E185B&228A"
        ("26MT-05G069", "26MT-05G070", "26MT-05G071") -> "PI-26MT-05G069&070&071"
        ("26MT-06M052", "26MT-06M087", "26MT-06M112") -> "PI-26MT-06M052&26MT-06M112&26MT-06M087A"

    同一个键理论上只应来自一个原值；若出现两个不同写法，记进 `split_map_conflicts()`。
    """
    global _SPLIT_MAP, _SPLIT_CONFLICTS
    if _SPLIT_MAP is not None:
        return _SPLIT_MAP
    mapping: dict[tuple[str, ...], str] = {}
    conflicts: list[tuple[tuple[str, ...], str, str]] = []
    seen_invoice: set[str] = set()
    for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
        invoice = str(row.get("invoiceCode") or "").strip()
        if not invoice or invoice in seen_invoice:
            continue
        seen_invoice.add(invoice)
        parts = literal_parts(invoice)
        if not parts:
            continue
        if parts in mapping and mapping[parts] != invoice:
            conflicts.append((parts, mapping[parts], invoice))
            continue
        mapping.setdefault(parts, invoice)
    _SPLIT_MAP = mapping
    _SPLIT_CONFLICTS = conflicts
    return mapping


def split_map_conflicts() -> list[tuple[tuple[str, ...], str, str]]:
    split_map()
    return list(_SPLIT_CONFLICTS)


def dump_split_map(path: str | "Any" = None) -> int:
    """把映射表写成 CSV 供人工核对（拆分结果 | 原值 | 段数）。"""
    import csv
    from pathlib import Path as _Path

    target = _Path(path) if path else (CACHE_ROOT / "reports" / "contract_split_map.csv")
    target.parent.mkdir(parents=True, exist_ok=True)
    rows = sorted(split_map().items(), key=lambda item: (-len(item[0]), item[1]))
    with target.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["拆分结果", "原值", "段数"])
        for parts, original in rows:
            writer.writerow([" + ".join(parts), original, len(parts)])
    return len(rows)


def lookup_original_by_parts(parts: "Iterable[str]") -> str:
    """把「报关单上拆出来的订单集合」反查成出运单号原值。"""
    key = tuple(sorted(part for part in parts if part))
    return split_map().get(key, "")


def order_cores_of(value: Any) -> set[str]:
    """出运产品行/任意单号归属的订单核心集合。"""
    if isinstance(value, (list, tuple)):
        value = ",".join(str(item) for item in value)
    return cores_of(value) if value else set()


def line_cores(row: dict[str, Any]) -> set[str]:
    """一条出运产品行属于哪些外销订单。

    优先用行自带的销售订单号（`order_codes` 是索引里的归一字段），
    其次采购订单号，最后才用发票号——发票号可能是联合写法，
    直接把整张出运单的订单都算进来，会破坏「逐个对应」。
    """
    for field in ("order_codes", "销售订单号"):
        cores = order_cores_of(row.get(field))
        if cores:
            return cores
    cores = order_cores_of(row.get("采购订单号") or row.get("purchase_code"))
    if cores:
        return cores
    return order_cores_of(row.get("invoice_code"))


def container_shipments(contract: Any) -> list[ShipmentRef]:
    """包含合同订单的出运单（允许出运单比合同大：一张出运单装多张报关单的货）。"""
    target = cores_of(contract)
    if not target:
        return []
    return [row for row in shipments() if row.cores & target]


def contract_part_labels(contract: Any) -> tuple[str, ...]:
    """合同号拆出来的订单集合（映射表的键口径）。"""
    return tuple(sorted(cores_of(contract)))
