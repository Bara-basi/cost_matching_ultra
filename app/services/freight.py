"""出运单费用（运费/保费）查询，用于背包容量计算。"""
from __future__ import annotations

import json
from pathlib import Path

DETAIL_DIR = Path(__file__).resolve().parents[2] / ".cache" / "erp" / "details" / "shipments"

# 客户费用的「加减方向」，即 报关金额 = 出运金额 + 方向 × 费用：
#   sub（加在报关金额里）= 报关金额 − 费用 才是出运金额/容量；
#   add（在报关金额之外）= 出运金额 = 报关金额 + 费用。
#
# 睿贝 ERP 的 `客户费用信息` 里确实有 `加减` 字段，但 MCP `shipment.find` 返回的
# 该字段**恒为空**（353 条缓存 + 12 张线上单据实测都是空串），所以这里按已闭合的单据
# 统计出的稳定规律把方向定下来（统计口径：227 个全精确闭合的合同）：
#   检验费/手续费/木箱费/试样费/模具费/内陆运费/仓储费/包装费 → sub（12/4/2/2/1/1/1/1，无 add）
#   罚金/FOB退费 → add（4/2，无 sub）
#   运费 → sub 13 : add 1；折扣 → sub 1 : add 2
# 方向定死能砍掉一半的组合，显著减少"费用口径不唯一"的假多解；
# 若按定死方向解不齐，运行时仍会退回两个方向都试。
FEE_DIRECTION_DEFAULTS: dict[str, str] = {
    "检验费": "sub",
    "手续费": "sub",
    "木箱费": "sub",
    "试样费": "sub",
    "模具费": "sub",
    "内陆运费": "sub",
    "仓储费": "sub",
    "包装费": "sub",
    "材料费": "sub",
    "运费": "sub",
    "罚金": "add",
    "FOB退费": "add",
    "折扣": "add",
}


def _safe(text: str) -> str:
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in str(text or ""))


def expense_cents(invoice_codes: list[str]) -> int:
    """汇总这些出运单的客户费用（运费等），单位分。"""
    total = 0
    for code in invoice_codes:
        path = DETAIL_DIR / f"{_safe(code)}.json"
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for row in payload.get("expenseList") or []:
            try:
                total += int(round(float(str(row.get("金额") or 0).replace(",", "")) * 100))
            except ValueError:
                continue
    return total


def expense_kinds(invoice_codes: list[str]) -> dict[str, float]:
    """按费用名称汇总本次出运的客户费用，用于判断是否「只有运费」。"""
    out: dict[str, float] = {}
    for code in invoice_codes:
        path = DETAIL_DIR / f"{_safe(code)}.json"
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for row in payload.get("expenseList") or []:
            name = str(row.get("费用名称") or "").strip() or "(未命名)"
            try:
                amount = float(str(row.get("金额") or 0).replace(",", ""))
            except ValueError:
                amount = 0.0
            if amount:
                out[name] = out.get(name, 0.0) + amount
    return out


def _num(value) -> float | None:
    text = str(value if value is not None else "").replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def fee_directions(invoice_codes: list[str]) -> dict[str, str]:
    """按睿贝自己的口径**定死**客户费用的加减方向。

    恒等式（ERP 自己成立）：

        出运总金额 − 产品行合计 = Σ(方向 × 客户费用)

    方向取「加(+)」→ 报关金额 = 出运金额 + 费用（本项目的 `sub`：算容量时扣掉）；
    「减(-)」→ 报关金额 = 出运金额 − 费用（本项目的 `add`：算容量时加回）。

    实测 1182 张出运明细里，有客户费用的 287 张中 **271 张能唯一解出方向组合**
    （15 张因费用金额相同而多解、1 张因半包含订单等口径对不上）。
    只有唯一解时返回；无法判定时返回空，由调用方退回经验表或两方向都试。
    """
    from itertools import product

    result: dict[str, str] = {}
    for code in invoice_codes:
        path = DETAIL_DIR / f"{_safe(code)}.json"
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        total = _num((payload.get("baseInfo") or {}).get("出运总金额"))
        if total is None:
            continue
        product_total = sum(
            _num(row.get("出运金额")) or 0.0 for row in payload.get("productList") or []
        )
        rows = [
            (str(row.get("费用名称") or "").strip(), _num(row.get("金额")) or 0.0)
            for row in payload.get("expenseList") or []
        ]
        rows = [(name, amount) for name, amount in rows if amount]
        if not rows:
            continue
        target = round(total - product_total, 2)
        amounts = [amount for _name, amount in rows]
        solutions = [
            signs
            for signs in product((1, -1), repeat=len(amounts))
            if abs(sum(sign * amount for sign, amount in zip(signs, amounts)) - target) <= 0.02
        ]
        if len(solutions) != 1:
            continue
        for (name, _amount), sign in zip(rows, solutions[0]):
            direction = "sub" if sign > 0 else "add"
            if result.get(name, direction) != direction:
                result.pop(name, None)  # 同名费用方向矛盾 → 不采用
            else:
                result[name] = direction
    return result
