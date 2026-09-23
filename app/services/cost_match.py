"""采购金额匹配：用入库单「实发金额」给拆单记录定成本。

口径（2026-09-24 与用户确认）：

1. **金额来源 = 入库单实发合计**（不是订单数据区、不是 ERP 采购订单界面数据）。
   实发区才是「产品做出来以后」的落地金额；订单区是当初下单的预期金额。
   预付款/尾款属于付款信息，与本任务无关，不进成本。
2. **金额单位 = 采购单**（一个采购单对应一个工厂、一个采购合同）。
   同一采购单拆成多条拆单记录时（不同产品类型 / 跨报关单），
   按 ERP 出运产品行的**出运采购金额(RMB)** 占比分摊；
   出运行缺失时退回该记录的**出运金额(USD)** 占比。
3. 一个采购单的入库单可能有多份（分批 / 补款），实发金额按份累加。

输出：`outputs/cost_match/采购金额_匹配.xlsx`（记录级）+ 单元级差异表。
"""
from __future__ import annotations

import collections
import json
import re
from decimal import Decimal
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT
from app.services.contract_shipments import cores_of
from app.services.grn_select import erp_amount, order_key, select

# erp_cache.CACHE_ROOT 已经指向 `.cache/erp`
DETAIL_DIR = CACHE_ROOT / "details" / "shipments"


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


def dec(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal(0)
    text = str(value).replace(",", "").strip()
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def block_batch(label: str) -> str:
    """从「实发数据-X」这类区块名里认出批次记号（A/B/C/D、第N批）。"""
    text = str(label or "").strip().upper()
    if not text or "+" in text or "&" in text:
        return ""  # 合并区块（A+B）不拆
    chinese = re.search(r"第([一二三四五六七八九十\d]+)批", text)
    if chinese:
        return "第" + chinese.group(1) + "批"
    cleaned = re.sub(r"[\d\./\s\-_]+", " ", text)
    letters = re.findall(r"(?<![A-Z])([A-Z])(?![A-Z])", cleaned)
    if letters:
        return letters[-1]
    return ""


def batch_map_of(items: list[dict]) -> dict[str, Decimal]:
    """入库单的批次区块 → 金额；同一批次以后出现的为准（覆盖更正值）。

    同一份表里同一个批次标签出现两次且金额不同时（供应商把两个区块都写成
    `实发数据-B`），这个批次记号含义不唯一，**不放进映射**，避免批次直取取错。
    """
    out: dict[str, Decimal] = {}
    ambiguous: set[str] = set()
    for item in items:
        seen_here: dict[str, Decimal] = {}
        for block in item.get("blocks") or []:
            token = block_batch(block.get("label"))
            amount = dec(block.get("amount"))
            if not token or not amount:
                continue
            if token in seen_here and seen_here[token] != amount:
                ambiguous.add(token)
            seen_here[token] = amount
            out[token] = amount
    for token in ambiguous:
        out.pop(token, None)
    return out


def block_tokens(payload_items: list[dict]) -> dict[str, Decimal]:
    """单个附件的批次区块映射（不做跨附件合并）。"""
    out: dict[str, Decimal] = {}
    for item in payload_items:
        for block in item.get("blocks") or []:
            token = block_batch(block.get("label"))
            amount = dec(block.get("amount"))
            if token and amount:
                out[token] = amount
    return out


def _token_of(text: str) -> str:
    # 先把「单号核心（含佣金 Y、ADD）」整段剥掉，剩下的第一个独立字母就是批次：
    # `26MT-10E101-Y-A` → `-A` → `A`；`477G` → `G`
    stripped = re.sub(
        r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:[\s\-_]?Y)?(?:[\s\-_]*ADD\d*)?",
        "",
        str(text or ""),
        count=1,
    )
    match = re.search(r"(?<![A-Za-z])([A-Za-z])(?![A-Za-z0-9])", stripped)
    if match:
        return match.group(1).upper()
    chinese = re.search(r"第([一二三四五六七八九十\d]+)批", text)
    if chinese:
        return "第" + chinese.group(1) + "批"
    return ""


def record_batch_token(contract: str, purchase_code: str = "") -> str:
    """报关行合同号里的批次记号：`26MT-01P242Y-C` → `C`。

    联合号（`25MT-07F379Y-F&477G`）要挑出属于本采购单的那一段——
    按流水号匹配（本单 25MT-07F477 → 取 `477G` → `G`）。
    """
    text = str(contract or "").strip()
    if not text:
        return ""
    parts = [part for part in re.split(r"[&,，;；]", text) if part.strip()]
    if len(parts) > 1 and purchase_code:
        serial = re.search(r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z](\d{2,3})", str(purchase_code))
        if serial:
            for part in parts:
                if serial.group(1) in part:
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


def shipment_rmb_by_kind() -> dict[tuple[str, str], Decimal]:
    """(采购单号, 产品类型) → ERP 出运采购金额(RMB) 合计。"""
    table: dict[tuple[str, str], Decimal] = collections.defaultdict(Decimal)
    if not DETAIL_DIR.exists():
        return table
    for path in DETAIL_DIR.glob("*.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for row in payload.get("productList") or []:
            code = str(row.get("采购订单号") or "").strip()
            kind = str(row.get("海关商品（中文）") or "").strip()
            if not code:
                continue
            table[(code, kind)] += dec(row.get("出运采购金额(RMB)"))
    return table


def shipment_rmb_by_po() -> dict[str, Decimal]:
    """采购单号 → **整张采购单**全部出运产品行的 出运采购金额(RMB) 合计。

    用来判断"入库单里的钱覆盖的是整单，还是只覆盖我们手里这几张报关单的货"：

    * `26MT-02K065Y-GYL`：入库单 480,793.11 ≈ 整单货值 477,919.56（+样管试样费）
      → 覆盖整单；而我们的报关单只装了其中 426,407.34 的货（另一部分在出运单 B 上、
      对应的报关单不在数据里），所以这一条只能拿它的份额，飞书给的 429,280.88
      正是"该报关单的货值 + 全部样管试样费"。
    * `25MT-05G595-GYL`：入库单 295,039.10 ≈ 手里报关单的货值 295,039.10（整单 ERP 是 442,447.10）
      → 入库单只覆盖我们这一部分 → 该摊满就摊满（飞书也是 295,039.10）。
    """
    table: dict[str, Decimal] = collections.defaultdict(Decimal)
    if not DETAIL_DIR.exists():
        return {}
    for path in DETAIL_DIR.glob("*.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for row in payload.get("productList") or []:
            code = str(row.get("采购订单号") or "").strip()
            if code:
                table[code] += dec(row.get("出运采购金额(RMB)"))
    return dict(table)


def shipment_usd_by_po() -> dict[str, Decimal]:
    """采购单 → 它全部出运产品行的出运金额(USD) 合计。

    用来判断「当前这几张报关单是否覆盖了整张采购单的货」：
    记录 USD 合计 ≈ 采购单出运 USD 合计 → 全覆盖，可以把入库单实发总额摊满；
    明显小于 → 只覆盖一部分（例如只出了某个批次），只能按各行自己的金额取。
    """
    table: dict[str, Decimal] = collections.defaultdict(Decimal)
    if not DETAIL_DIR.exists():
        return {}
    for path in DETAIL_DIR.glob("*.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for row in payload.get("productList") or []:
            code = str(row.get("采购订单号") or "").strip()
            if code:
                table[code] += dec(row.get("出运金额"))
    return dict(table)


def po_cost(purchase_code: str) -> tuple[Decimal, dict[str, Any]]:
    """采购单的入库单实发合计 + 细节。"""
    result = select(purchase_code)
    total = sum((dec(item["amount"]) for item in result["kept"]), Decimal(0))
    detail = {
        "files": result["files"],
        "kept": len(result["kept"]),
        "dropped": len(result["dropped"]),
        "batches": len(result["groups"]),
        "kept_names": [item["original"] for item in result["kept"]],
        "issues": result["issues"],
    }
    return total, detail


def costs_for(codes: list[str]) -> dict[str, tuple[Decimal, dict[str, Any]]]:
    """批量算成本；同一份入库单被多个采购单引用时按 ERP 采购金额拆分。

    实测有 3 份附件同时挂在主单与附加单下（例如
    `26MT-06N145-GYL ADD1 溪流供应链 入库单`），两边都算就会重复计成本。
    拆法用 ERP 采购金额当权重——它本身就是这两张采购单的合同金额。
    """
    selections = {code: select(code) for code in codes}

    def add_marker(text: str) -> str:
        """这张单/这个采购单号是不是附加单（ADD n）。

        2026-09-24 修：只按 `order_key()` 里的 ADD 段判断会漏——`26MT-06N145-GYL ADD1 …`
        这种"ADD 在工厂后缀后面"的写法，`order_key` 认不出，于是被误判成"目录里没有本附加单
        标记"，退回睿贝金额并退出共表拆分，导致主单把整份入库单全吃下（29,916.67 重复计）。
        这里补一层宽松匹配：文本里任何位置出现 ADDn 都算。
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

    # 附加单目录里放的是主单的合并入库单（文件自己的 ADD 标记与采购单不符）时，
    # 这张表里的金额无法拆到本单，退回该采购单的 ERP 采购金额并注明。
    overrides: dict[str, Decimal] = {}
    for code, result in selections.items():
        kept = result["kept"]
        if not kept:
            continue
        mine = add_marker(code)
        if mine and all(add_marker(item["original"]) != mine for item in kept):
            reference = erp_amount(code)
            if reference:
                overrides[code] = reference
                result["issues"].append(
                    "附件是主单的合并入库单（不含本附加单标记），"
                    f"本单按 ERP 采购金额 {reference} 取值，待人工确认"
                )
    by_digest: dict[str, list[tuple[str, dict[str, Any]]]] = collections.defaultdict(list)
    own_totals: dict[str, Decimal] = {code: Decimal(0) for code in codes}
    for code, result in selections.items():
        for item in result["kept"]:
            if code in overrides:
                continue  # 已用 ERP 金额兜底，不再参与附件拆分
            digest = item.get("digest") or ""
            if digest:
                by_digest[digest].append((code, item))
            else:
                own_totals[code] += dec(item["amount"])

    out: dict[str, tuple[Decimal, dict[str, Any]]] = {}
    for code, result in selections.items():
        detail = {
            "files": result["files"],
            "kept": len(result["kept"]),
            "dropped": len(result["dropped"]),
            "batches": len(result["groups"]),
            "kept_names": [item["original"] for item in result["kept"]],
            "issues": list(result["issues"]),
            "batch_amounts": {
                token: str(amount) for token, amount in batch_map_of(result["kept"]).items()
            },
            "block_count": sum(
                1
                for item in result["kept"]
                for _block in (item.get("blocks") or [])
            ),
        }
        out[code] = (own_totals[code], detail)

    for code, amount in overrides.items():
        out[code] = (amount, out[code][1])

    for digest, refs in by_digest.items():
        owners = sorted({code for code, _item in refs})
        amount = max((dec(item["amount"]) for _code, item in refs), default=Decimal(0))
        if len(owners) == 1:
            out[owners[0]][1]["issues"].append(
                f"附件 {refs[0][1]['original']} 计入本单（唯一引用）"
            )
            out[owners[0]] = (out[owners[0]][0] + amount, out[owners[0]][1])
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
            out[code][1]["issues"].append(
                f"附件 {next(i['original'] for c, i in refs if c == code)} "
                f"被 {len(owners)} 个采购单共用，{basis}，本单分得 {share}"
            )
            out[code] = (out[code][0] + share, out[code][1])
    return out


def allocate_amounts(
    records: list[dict],
    total: Decimal,
    rmb_table: dict[tuple[str, str], Decimal],
    po_usd_total: Decimal | None = None,
    po_rmb_total: Decimal | None = None,
) -> list[tuple[Decimal, str]]:
    """给同一采购单的多条记录算分摊比例。

    优先级：
    1. **按入库单的实发区块对批次**——报关行合同号里的批次记号（`…-A`/`…B`）与
       入库单区块名（`实发数据-A`）能对上时，直接取该区块金额。
       实测 26MT-01P242Y：区块 A/B/C = 199,183.68 / 209,395.84 / 267,863.09，
       与飞书三张报关单的采购金额完全一致。
    2. 按记录实际装到的出运产品行人民币金额占比；
    3. 按出运金额(USD)占比。
    """
    batch_map = records[0].get("_batch_map") if records else None
    if batch_map:
        tokens = [str(record.get("_batch_token") or "") for record in records]
        if all(token and token in batch_map for token in tokens):
            picked = [batch_map[token] for token in tokens]
            unique_sum = sum((batch_map[token] for token in set(tokens)), Decimal(0))
            # 区块金额不可能超过这张采购单自己的成本——超过说明区块属于别的采购单
            # （典型：附加单目录里挂着主单的合并入库单），此时不能用区块直取。
            # 注意按「去重后的批次」比较：同一批次落在多张报关单上时按批次算一次。
            if unique_sum <= total * Decimal("1.05"):
                grouped: dict[str, list[int]] = collections.defaultdict(list)
                for position, token in enumerate(tokens):
                    grouped[token].append(position)
                results: list[tuple[Decimal, str] | None] = [None] * len(tokens)
                for token, positions in grouped.items():
                    amount = batch_map[token]
                    weights = [dec(records[p].get("出运采购金额合计")) for p in positions]
                    weight_sum = sum(weights, Decimal(0))
                    # 这个批次是否被当前报关单覆盖全：覆盖不全（例如这一批还出了别的报关单，
                    # 而那张单不在这批数据里）时不能把整批金额摊给眼前的记录，只能按各行金额取。
                    # 区块金额要和各报关行的人民币合计「对得上」才敢直取：
                    # 只在 ±15% 以内算一致。实测 26MT-06H189 的区块字母与报关批次
                    # 并不同名（B 区块 60,809.54 对应的是 775,100.46 的货），
                    # 单边阈值会把这种错配放过去。
                    covered = (
                        weight_sum > 0
                        and amount > 0
                        and abs(weight_sum - amount) / amount <= Decimal("0.15")
                    )
                    if len(positions) == 1 and weight_sum <= 0:
                        # 只有一条记录对到这个批次、而睿贝没给这行定价时，
                        # 批次记号已经唯一确定，直接取该区块金额
                        results[positions[0]] = (
                            amount.quantize(Decimal("0.01")),
                            f"入库单区块-{token}（睿贝该行无金额，按批次直取）",
                        )
                        continue
                    if not covered:
                        for position, weight in zip(positions, weights):
                            results[position] = (
                                weight.quantize(Decimal("0.01")),
                                f"入库单区块-{token} 只覆盖一部分（{weight_sum} / {amount}），按本行金额取",
                            )
                        continue
                    if len(positions) == 1:
                        position = positions[0]
                        results[position] = (
                            amount.quantize(Decimal("0.01")),
                            f"入库单区块-{token}",
                        )
                        continue
                    # 同一批次落在多张报关单上：按各报关行的出运采购金额占比拆这一个批次
                    if weight_sum <= 0:
                        weights = [Decimal(1)] * len(positions)
                        weight_sum = Decimal(len(positions))
                        basis = f"入库单区块-{token}（无权重，均分）"
                    else:
                        basis = f"入库单区块-{token}（按出运采购金额占比）"
                    for position, weight in zip(positions, weights):
                        results[position] = (
                            (amount * weight / weight_sum).quantize(Decimal("0.01")),
                            basis,
                        )
                if all(item is not None for item in results):
                    return [item for item in results if item is not None]
    # 权重一（首选）：拆单记录实际装到的出运产品行人民币金额
    weights = [dec(record.get("出运采购金额合计")) for record in records]
    source = "本记录出运采购金额(RMB)"
    records_usd = sum((dec(record.get("出运金额合计")) for record in records), Decimal(0))
    coverage = (records_usd / po_usd_total) if po_usd_total else None
    # 个别报关行在睿贝明细里没有采购金额（未定价）→ 该行金额留空，其余按比例分摊整单
    missing_weight = [
        position for position, weight in enumerate(weights) if weight <= 0
    ]
    if missing_weight and sum(weights, Decimal(0)) > 0:
        # 若这几张报关单已覆盖整单（≥50%），说明缺的只是睿贝那一行没定价，
        # 用出运金额(USD) 给全组分摊，比"某一行留空、其余摊满"更接近真实。
        if coverage is not None and coverage >= Decimal("0.5"):
            usd_weights = [dec(record.get("出运金额合计")) for record in records]
            usd_total = sum(usd_weights, Decimal(0))
            # 只有当「没定价的那几条」本身也装着有分量的货（≥5% 的出运金额）时才这么分；
            # 若只是一条 30 USD 的零头（26MT-02N126Y），留空让人工看更稳妥。
            shares = [(w / usd_total) if usd_total else Decimal(0) for w in usd_weights]
            if usd_total > 0 and all(
                shares[position] >= Decimal("0.05") for position in missing_weight
            ):
                return [
                    (
                        (total * weight / usd_total).quantize(Decimal("0.01")),
                        f"按出运金额(USD)占比分摊 {weight / usd_total * 100:.4f}%（睿贝部分行未定价）",
                    )
                    for weight in usd_weights
                ]
        kept_positions = [p for p in range(len(weights)) if p not in missing_weight]
        share_total = sum((weights[p] for p in kept_positions), Decimal(0))
        results: list[tuple[Decimal, str] | None] = [None] * len(weights)
        for position in missing_weight:
            results[position] = (
                Decimal(0),
                "睿贝该行未定价（采购金额留空，待人工）",
            )
        for position in kept_positions:
            results[position] = (
                (total * weights[position] / share_total).quantize(Decimal("0.01")),
                f"{source} {weights[position] / share_total * 100:.4f}%（其余行未定价）",
            )
        # 未定价的行金额留空由上层写 Excel 时用空串表示
        return [
            (item[0], item[1]) if item is not None else (Decimal(0), "")
            for item in results
        ]
    if sum(weights) <= 0:
        # 权重二：ERP 出运采购金额(RMB)，按 (采购单, 产品类型) 汇总
        weights = [
            rmb_table.get(
                (
                    str(record.get("采购单号") or "").strip(),
                    str(record.get("产品类型") or "").strip(),
                ),
                Decimal(0),
            )
            for record in records
        ]
        source = "ERP出运采购金额(RMB)"
    if sum(weights) <= 0:
        # 权重三：拆单记录里的出运金额(USD)
        weights = [dec(record.get("出运金额合计")) for record in records]
        source = "出运金额(USD)"
    total_weight = sum(weights)
    if total_weight <= 0:
        weights = [Decimal(1)] * len(records)
        total_weight = Decimal(len(records))
        source = "均分"
    else:
        # 是否覆盖了整张采购单的货：按出运金额(USD) 判断最可靠——
        # 睿贝的人民币单价可能是旧数据（如 26MT-03P075Y-HX 只有 364.20），
        # 用钱比钱会被带偏，用货量比货量不会。
        # 阈值 50%：实测「报关只装了整单一小部分」（10E101-Y 9.8%、05G019 15.5%）
        # 要按各行自己的金额取；而「装了一半以上」的（08C636 50%、03T203Y-XMLS 66.9%、
        # 03P075Y-HX 100%）飞书与单据都按整单摊满。
        partial = (
            coverage is not None and coverage < Decimal("0.5")
        ) or (
            coverage is None and total > 0 and total_weight < total * Decimal("0.5")
        )
    if total > 0 and partial:
        # 报关行只覆盖采购单的一部分（例：采购单出了 A/B/C/D 四批，当前报关单只装 A 批）
        # → 不能把整张采购单的金额按比例放大到这几条记录上，直接取出运采购金额本身，
        # 未覆盖的部分留给对应批次。
        return [
            (
                weight.quantize(Decimal("0.01")),
                f"{source} 直接取值（本单只覆盖采购单一部分：{total_weight} / {total}）",
            )
            for weight in weights
        ]
    # 覆盖度不到一半时已经在上面按各行金额取了；这里处理"装了大半、但仍有货落在
    # 我们没有的报关单上"的情况（26MT-02K065Y-GYL：装了 91% 的货）。
    # 判据：入库单总额更贴近"整单货值"还是"手里这几张报关单的货值"——
    # 更贴近整单 → 说明入库单里的钱是整单的，必须按整单占比取，否则会把
    # 别人那部分的钱摊到眼前这几条记录上。
    if (
        total > 0
        and po_rmb_total
        and po_rmb_total > total_weight * Decimal("1.005")
        and abs(total - po_rmb_total) < abs(total - total_weight)
    ):
        return [
            (
                (total * weight / po_rmb_total).quantize(Decimal("0.01")),
                f"{source} 按整单ERP占比 {weight / po_rmb_total * 100:.4f}%"
                f"（入库单覆盖整单 {po_rmb_total}，本单只装其中 {total_weight}）",
            )
            for weight in weights
        ]
    return [
        ((total * weight / total_weight).quantize(Decimal("0.01")), f"{source} {weight / total_weight * 100:.4f}%")
        for weight in weights
    ]
