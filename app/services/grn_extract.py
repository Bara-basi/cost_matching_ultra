"""入库单附件 → 结构化数据。

策略（2026-09-24 用户确认）：

1. **规则优先**：迈拓自制模板（编号 MT-QR-PD-002 / 001）版式稳定，
   用规则直接解析出明细与合计，快且可复现；
2. **模型兜底/复核**：版式不像模板、规则校验不自洽、或附件本身写得乱
   （典型：「金额算错，然后又给一行正确金额」）时，把表格文本或图片交给
   DeepSeek（`app/services/deepseek_client.py`）解析；
3. **结果落缓存**：`.cache/erp/attachments/parsed/<采购单号>/<文件名>.json`，
   按内容 sha256 前 16 位失效，支持断点续跑。
"""
from __future__ import annotations

import json
import re
import collections
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from app.services import deepseek_client
from app.services.erp_cache import CACHE_ROOT, ensure_dir
from app.services.sheet_reader import Attachment, read_attachment, sheets_to_text

PARSED_ROOT = CACHE_ROOT / "attachments" / "parsed"
SAFE_RE = re.compile(r"[^\w\u4e00-\u9fff.\-]")

SYSTEM_PROMPT = """你是外贸公司的单据录入员，负责把「入库单/发货单」附件读成结构化数据。

硬性要求：
1. 只回 JSON，不要解释、不要 Markdown 围栏。
2. 数字按单据原样抄写，不要换算、不要四舍五入、不要虚构。读不到就填 null。
3. 单据里如果出现"改错"痕迹——例如某行金额算错后又补了一行正确金额、
   手写批注改了数字、或有划线/删除——请在 `anomalies` 里如实说明，
   并把**最终有效**的数值填进字段。
4. 抬头品名、供应商、单号、日期尽量照抄原文。

JSON 结构：
{
  "doc_type": "入库单|发货单|出货单|对账单|其他",
  "contract_no": "单据上写的单号，多个用 & 连接",
  "supplier": "供应商全称",
  "item_name": "抬头品名",
  "doc_date": "YYYY-MM-DD 或原文",
  "batch": "单据上写的批次标记（如 A / 2 / 第二批），没有填 null",
  "lines": [
    {"material": "", "name": "", "spec": "", "qty": 0, "unit": "",
     "unit_price": 0, "amount": 0, "weight": null, "packages": null}
  ],
  "order_totals":  {"qty": 0, "amount": 0},
  "settled_totals": {"qty": 0, "amount": 0, "weight": null},
  "settled_blocks": [
    {"label": "实发数据-A+B", "amount": 0, "qty": 0, "weight": null}
  ],
  "extra_fees": [{"label": "木箱费", "amount": 0}],
  "batches": [{"batch": "A", "settled_amount": 0, "qty": 0, "weight": null}],
  "prepaid": 0,
  "payable_now": 0,
  "notes": "备注原文",
  "anomalies": ["发现的问题"]
}

字段说明：
- `qty` 与 `unit` 按单据的计量单位抄（公斤/米/只/支…）；
  单据另有「支数/件数/盘数」这类包装计数时填进 `packages`。
- `weight` 只在单据真的给了重量（KG/净重/理算重量）时填。
- `prepaid` 是备注里写的预付款；`payable_now` 是「本次付款/本次付尾款」。
- `extra_fees` 收单据里没有数量单价的纯费用行（包装费/木箱费/渗透实验费…）。
- **合计行下方的批注要逐行看**，两类都要收：
  ① 加项（「补款/补差/另加/另收 …」+ 木箱费/样品费/坡口费/吊绳…）→ `extra_fees` 填**正数**；
  ② **抵扣项**（「抵扣/扣款/扣减/冲减/减免/扣除/退回 …」，常见「抵扣运费与试验费用：1109」）
     → `extra_fees` 填**负数**（如实抄下来，不要丢）。
  实例：25MT-07F591-GYL 合计 170,213.04、表尾「抵扣运费与试验费用：1109」，
  实发 = 169,104.04（可用同页「本次付尾款 135,104.04 = 实发 − 预付款 34,000」交叉验证）。
- **一张表里有多个「实发数据-A / -B / -C」区块时**（一个文件覆盖多个发货批次），
  请**逐区块**把「合计行里该区块的金额」填进 `settled_blocks`（label 照抄区块名），
  `settled_totals.amount` 填这些区块金额之和；没有多区块时 `settled_blocks` 就一个元素。
- `settled_totals.amount` **必须是实发数据区的合计**，不是订单数据区的「金额（元）」。
  订单数据区是下单时的预期金额，只填进 `order_totals`，不要混进 settled。
"""


def safe_name(text: str, fallback: str = "file") -> str:
    cleaned = SAFE_RE.sub("_", str(text or "")).strip("_")
    return cleaned or fallback


def cache_path(purchase_code: str, file_name: str) -> Path:
    return PARSED_ROOT / safe_name(purchase_code) / (safe_name(file_name, "file") + ".json")


def load_cached(purchase_code: str, file_name: str, digest: str) -> dict[str, Any] | None:
    path = cache_path(purchase_code, file_name)
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None
    if payload.get("_digest") != digest:
        return None
    return payload


def save_cached(purchase_code: str, file_name: str, payload: dict[str, Any]) -> None:
    path = cache_path(purchase_code, file_name)
    ensure_dir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")


# --------------------------------------------------------------------------- 数值


def to_decimal(value: Any) -> Decimal | None:
    """把单元格内容转成 Decimal；取不到返回 None。"""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return Decimal(match.group())
    except InvalidOperation:
        return None


def dec_str(value: Decimal | None) -> str | None:
    if value is None:
        return None
    text = format(value.normalize(), "f")
    return text


NUMERIC_CELL = re.compile(r"^-?[\d,]+(?:\.\d+)?%?$")


def is_number_cell(text: str) -> bool:
    """整个单元格就是个数字（用于区分「金额数字」和「带数字的说明文字」）。"""
    return bool(NUMERIC_CELL.match(text.replace(" ", "")))


# --------------------------------------------------------------- 批次区块（解析/选单/成本三处共用）


def block_batch(label: str) -> str:
    """从「实发数据-A」「第2批」这类区块名里认出批次记号。

    合并区块（`A+B`、`A&B`）含义不唯一，一律返回空串。
    这是入库单区块名 → 批次记号的**唯一实现**：附件解析、选单（去旧版）、
    成本分摊（按批次直取）三处都调它，避免同一套规则三份写法各自漂移。
    """
    text = str(label or "").strip().upper()
    if not text or "+" in text or "&" in text:
        return ""
    chinese = re.search(r"第([一二三四五六七八九十\d]+)批", text)
    if chinese:
        return "第" + chinese.group(1) + "批"
    cleaned = re.sub(r"[\d\./\s\-_]+", " ", text)
    letters = re.findall(r"(?<![A-Z])([A-Z])(?![A-Z])", cleaned)
    return letters[-1] if letters else ""


def blocks_to_amounts(blocks: list[dict[str, Any]]) -> dict[str, Decimal]:
    """实发区块列表 → {批次记号: 金额}。

    同一份表里同一个记号出现两次且金额不同（供应商把两个区块都写成 `实发数据-B`）
    时，这个记号含义不唯一，**不放进映射**，避免按批次直取取错。
    """
    out: dict[str, Decimal] = {}
    seen: dict[str, Decimal] = {}
    ambiguous: set[str] = set()
    for block in blocks or []:
        if not isinstance(block, dict):
            continue
        token = block_batch(block.get("label"))
        amount = to_decimal(block.get("amount"))
        if not token or not amount:
            continue
        if token in seen and seen[token] != amount:
            ambiguous.add(token)
        seen[token] = amount
        out[token] = amount
    for token in ambiguous:
        out.pop(token, None)
    return out


def line_amounts(payload: dict[str, Any]) -> Decimal:
    """明细行的金额求和（rule 与 llm 两种结构都认）。"""
    total = Decimal(0)
    for line in payload.get("lines") or []:
        if not isinstance(line, dict):
            continue
        value = None
        for key in ("settled_amount", "amount", "金额", "金额 元", "含税金额（元）"):
            if line.get(key) not in (None, ""):
                value = to_decimal(line.get(key))
                if value is not None:
                    break
        if value is None:
            for label, raw in line.items():
                if "金额" in str(label) and "单价" not in str(label):
                    value = to_decimal(raw)
                    if value is not None:
                        break
        if value is not None:
            total += value
    return total


# ---- 数量核验（2026-09-28 新增）------------------------------------------------
# 「出运数量不能比入库数量多」的核验要按**单位**对齐：入库单里的数量列五花八门
# （数量 个 / 数量只 / 数量 PC / 数量 米 / KG…），先把列名归到三个"单位族"里再比。
COUNT_UNITS = {
    "ea", "pc", "pcs", "pce", "each", "piece", "set",
    "个", "只", "支", "件", "根", "套", "片", "条",
}
LENGTH_UNITS = {"m", "mtr", "mtrs", "meter", "metre", "米", "ft", "inch", "英寸"}
WEIGHT_UNITS = {"kg", "kgs", "千克", "公斤", "ton", "tons", "t", "lb", "磅", "克", "g"}
_QTY_COLUMN_HINT = ("数量", "支数", "只数", "件数", "个数", "米数", "重量", "kg", "KG")
_QTY_COLUMN_SKIP = ("单价", "金额", "总价", "备注")


def unit_family(text: Any) -> str:
    """把单位/列名归到 count / length / weight，认不出返回空串。"""
    key = re.sub(r"[（）()#\d]", " ", str(text or ""))
    parts = [part for part in key.replace("/", " ").split() if part]
    for part in reversed(parts):
        token = part.strip().lower()
        if token in COUNT_UNITS:
            return "count"
        if token in LENGTH_UNITS:
            return "length"
        if token in WEIGHT_UNITS:
            return "weight"
    return ""


def line_quantities(payload: dict[str, Any]) -> dict[str, Decimal]:
    """这份入库单的明细数量，**按单位族求和**：{"count": 6, "length": 294, "weight": 1282.5}。

    同一行里常有多个同族数量列（`数量 个` / `数量只` / `数量 个#2`…），它们写的是同一个数，
    所以每行、每族**只取最大值**再加总，避免重复计数。
    """
    totals: dict[str, Decimal] = collections.defaultdict(Decimal)
    for line in payload.get("lines") or []:
        if not isinstance(line, dict):
            continue
        best: dict[str, Decimal] = {}
        for label, raw in line.items():
            text = str(label)
            if any(skip in text for skip in _QTY_COLUMN_SKIP):
                continue
            if not any(hint in text for hint in _QTY_COLUMN_HINT):
                continue
            family = unit_family(text)
            if not family:
                continue
            value = to_decimal(raw)
            if value is None:
                continue
            if value > best.get(family, Decimal(0)):
                best[family] = value
        for family, value in best.items():
            totals[family] += value
    return dict(totals)


# --------------------------------------------------------------------------- 规则解析

QTY_LABELS = ("数量", "支数", "米数", "只数", "件数", "pcs", "kgs")
WEIGHT_LABELS = ("重量", "净重", "毛重", "理算重量", "做负重量", "kg")
AMOUNT_LABELS = ("金额（元）", "金额(元)", "金额", "含税金额", "理算金额")
PRICE_LABELS = ("单价", "理算单价", "最终单价", "含税单价")
ITEM_NAME_HINT = ("品名", "名称")


@dataclass
class RuleResult:
    ok: bool
    payload: dict[str, Any]
    reason: str = ""


def _after_label(cell: str) -> str:
    for mark in (":", "："):
        if mark in cell:
            return cell.split(mark, 1)[1].strip()
    return ""


def _next_value(row: list[str], col: int) -> str:
    for cell in row[col + 1:]:
        if cell:
            return cell
    return ""


DATE_RE = re.compile(r"(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})")


def parse_template(attachment: Attachment) -> RuleResult:
    """解析迈拓自制模板。ok=False 表示「不像模板」或「读不干净」，交给模型。"""
    if not attachment.sheets:
        return RuleResult(False, {}, "没有表格内容")
    grid = attachment.sheets[0].rows

    # ---- 表头行（含「序号」的那一行）
    header_row = None
    for index, row in enumerate(grid):
        if any(cell == "序号" for cell in row):
            header_row = index
            break
    if header_row is None:
        return RuleResult(False, {}, "找不到表头行（没有「序号」）")

    # ---- 抬头：只在表头之前找，且必须以「标签:」形式出现
    title = ""
    supplier = ""
    item_name = ""
    doc_date = ""
    for index, row in enumerate(grid[:header_row]):
        joined = " ".join(cell for cell in row if cell)
        if "入库单" in joined and not title:
            title = joined
        for col, cell in enumerate(row):
            if cell.startswith("品名") and (":" in cell or "：" in cell):
                item_name = _after_label(cell) or _next_value(row, col)
            elif cell.startswith("供应商") and (":" in cell or "：" in cell):
                supplier = _after_label(cell) or _next_value(row, col)
            elif cell.startswith("日期") and (":" in cell or "：" in cell):
                doc_date = _after_label(cell) or _next_value(row, col)
            elif not doc_date:
                found = DATE_RE.search(cell)
                if found:
                    doc_date = "-".join(
                        [found.group(1), found.group(2).zfill(2), found.group(3).zfill(2)]
                    )
    if "入库单" not in title and not item_name and not supplier:
        return RuleResult(False, {}, "不是迈拓模板（没有标题/品名/供应商抬头）")

    # ---- 明细起止：第一条「序号是数字」的行 → 第一条「合计」行
    data_start = None
    for index in range(header_row + 1, len(grid)):
        first = grid[index][0] if grid[index] else ""
        if first and first[0].isdigit():
            data_start = index
            break
    if data_start is None:
        return RuleResult(False, {}, "表头下面找不到明细行")

    total_index = None
    for index in range(data_start, len(grid)):
        cells = grid[index]
        first = cells[0] if cells else ""
        if first.startswith("合计") or first.startswith("总合计") or first.startswith("总计"):
            total_index = index
            break
        # 有的表「总合计」不在第一列（前面留空），只要行内有合计标记且带数字就算
        marker = next(
            (cell for cell in cells[:4] if cell.startswith(("合计", "总合计", "总计"))),
            "",
        )
        if marker and any(to_decimal(cell) not in (None, 0) for cell in cells):
            total_index = index
            break
    if total_index is None:
        return RuleResult(False, {}, "明细下面找不到「合计」行")

    # ---- 一张表可能有多条合计行：第一条是货值合计，费用行（木箱费/绳子费/过磅差抵扣…）
    #      常加在它后面、再给一条「最终合计」。实发金额以**最后一条有值的合计行**为准，
    #      否则会把合计之后的费用整块漏掉（实测 26MT-05G129 巨星补木箱费：少 590+17）。
    footer_start = len(grid)
    for index in range(total_index + 1, len(grid)):
        joined = " ".join(cell for cell in grid[index] if cell)
        if "制单" in joined or "审核" in joined:
            footer_start = index
            break
    total_rows = [total_index]
    for index in range(total_index + 1, footer_start):
        cells = grid[index]
        marker = next(
            (cell for cell in cells[:4] if cell.startswith(("合计", "总合计", "总计"))),
            "",
        )
        if marker:
            total_rows.append(index)
    last_total = total_rows[-1]
    total_row_set = set(total_rows)

    # ---- 列标签：表头行到明细行之间的所有行纵向拼接
    header_rows = grid[header_row:data_start]
    width = max(len(row) for row in grid[header_row:])
    labels: list[str] = []
    seen: dict[str, int] = {}
    for col in range(width):
        parts = [row[col] for row in header_rows if col < len(row) and row[col]]
        label = " ".join(parts).strip() or f"col{col + 1}"
        if label in seen:
            seen[label] += 1
            label = f"{label}#{seen[label]}"
        else:
            seen[label] = 1
        labels.append(label)

    def row_to_map(row: list[str]) -> dict[str, str]:
        return {
            label: (row[col] if col < len(row) else "")
            for col, label in enumerate(labels)
            if col < len(row) and row[col]
        }

    stop = last_total
    lines: list[dict[str, str]] = []
    extra_fees: list[dict[str, Any]] = []
    aux_rows: list[dict[str, Any]] = []
    fee_by_column: dict[str, Decimal] = {}
    fee_items: dict[str, list[Decimal]] = {}

    def amount_columns() -> list[tuple[int, str]]:
        return [
            (col, label)
            for col, label in enumerate(labels)
            if "金额" in label and "单价" not in label
        ]

    for index in range(data_start, stop):
        row = grid[index]
        if not any(row):
            continue
        if index in total_row_set:
            # 合计行本身不算明细、也不算费用
            continue
        first = row[0] if row else ""
        if not first or not first[0].isdigit():
            # 明细区里没有序号的行 = 费用/辅料行（管帽、吊绳、实验费…），
            # 它们的金额并进了合计，必须单独记账，否则合计对不上。
            amount = _row_amount(row, labels)
            for col, label in amount_columns():
                if col >= len(row):
                    continue
                value = to_decimal(row[col])
                if value:
                    fee_by_column[label] = fee_by_column.get(label, Decimal(0)) + value
                    fee_items.setdefault(label, []).append(value)
            label = first or next(
                (
                    cell
                    for cell in row
                    if cell and not is_number_cell(cell) and not cell.startswith("元")
                ),
                "",
            )
            if amount is not None:
                extra_fees.append({"label": label, "amount": dec_str(amount)})
            aux_rows.append({"label": label, "cells": [cell for cell in row if cell]})
            continue
        if first.startswith("按"):
            aux_rows.append({"label": first, "cells": [cell for cell in row if cell]})
            continue
        lines.append(row_to_map(row))

    # ---- 合计：逐列取「最后一条有值的合计行」；金额列若最后一条是 0（占位列），
    #      退回更早那条的非 0 值（05Q484-B 的 A 批只在第一条合计行里有数）
    totals: dict[str, str] = {}
    zeros: dict[str, str] = {}
    for index in reversed(total_rows):
        for label, value in row_to_map(grid[index]).items():
            if label in totals:
                continue
            if to_decimal(value) == 0 and "金额" in label:
                zeros.setdefault(label, value)
                continue
            totals[label] = value
    for label, value in zeros.items():
        totals.setdefault(label, value)

    # ---- 区块：一张表可能同时有订单数据 + 多个「实发数据-X」批次区块
    marker_index = None
    for index in range(header_row - 1, -1, -1):
        if any(cell.startswith("订单数据") for cell in grid[index]):
            marker_index = index
            break
    starts: list[tuple[int, str]] = []
    if marker_index is not None:
        for col, cell in enumerate(grid[marker_index]):
            # 区块名不一定是「实发数据」，也可能是「中通发溪流实发数据-裸管」「06H192A实发数据」
            # 或「实际数据」（实测 25MT-07F591-GYL 的回签件就写「实际数据」）
            if (
                cell.startswith("实发")
                or "实发数据" in cell
                or "实发总重" in cell
                or cell.startswith("实际数据")
                or "实际发货" in cell
            ):
                starts.append((col, cell))
    order_start = None
    if marker_index is not None:
        for col, cell in enumerate(grid[marker_index]):
            if cell.startswith("订单数据"):
                order_start = col

    def pick_column(labels_wanted: tuple[str, ...], low: int, high: int, *, avoid: tuple[str, ...] = ()) -> int | None:
        best: int | None = None
        fallback: int | None = None
        for col in range(max(low, 0), min(high, len(labels))):
            label = labels[col]
            if not any(key in label for key in labels_wanted):
                continue
            if any(bad in label for bad in avoid):
                continue
            # 「做负金额 / 做负奖励 / 理算金额」是派生的调整列，不是这一区块的结算金额
            if any(extra in label for extra in ("做负", "奖励", "调整", "退", "差异")):
                continue
            fallback = col  # 取最靠右的匹配列
            value = to_decimal(totals.get(label))
            # 有的表同一区块里金额列不止一个（订单金额/实发金额/做负金额），
            # 合计行里为 0 或空的那一列不是我们要的
            if value is not None and value != 0:
                best = col
        return best if best is not None else fallback

    def block_value(total_map: dict[str, str], col: int | None) -> str | None:
        if col is None:
            return None
        label = labels[col]
        return total_map.get(label) or None

    bounds = [col for col, _text in starts] + [len(labels)]
    settled_blocks: list[dict[str, Any]] = []
    settled_labels: set[str] = set()
    for position, (col, text) in enumerate(starts):
        low, high = col, bounds[position + 1]
        amount_col = pick_column(("金额",), low, high, avoid=("单价",))
        qty_col = pick_column(("数量", "支数", "米数"), low, high, avoid=("金额", "单价"))
        weight_col = pick_column(("重量", "净重"), low, high, avoid=("金额", "单价"))
        if amount_col is not None:
            settled_labels.add(labels[amount_col])
        settled_blocks.append(
            {
                "label": text,
                "amount": block_value(totals, amount_col),
                "qty": block_value(totals, qty_col),
                "weight": block_value(totals, weight_col),
            }
        )

    order_amount_col = pick_column(
        ("金额",), order_start if order_start is not None else 0,
        (starts[0][0] if starts else len(labels)), avoid=("单价",)
    )
    order_amount = block_value(totals, order_amount_col)

    # ---- 表尾补款/费用批注：合计行下方常有一行「补款 …（木箱费/样品费/坡口费…）」，
    #      这笔钱没有并进上面的合计，但确实要付给工厂。判据（2026-09-24 与飞书对账得出）：
    #        * 行内出现「补款/补差/另加/另收」时，取关键词右侧最近的数字；
    #        * 同一行里写了**费用性质**的词（木箱/样品/样管/坡口/绳子/包装/加工/成品/价格…）
    #          → 计入成本；其余（纯付款说明，如「补款 8820 2026.4.23」）只记进
    #          `tail_notes` 留痕，不动成本。
    #      实例：26MT-06M087 的「D补款2026.6.29 27079 木箱费 样品费 8公斤成品费」
    #      正是飞书比我方多的那 27,079；25MT-06H133 的「补款 8820」已经在它的合计里，
    #      行内没有费用词，因此不会被重复计入。
    TAIL_KEYWORDS = ("补款", "补差", "另加", "另收", "补收")
    # 表尾「抵扣类」批注：合计算的是货值，但实际付款时被扣掉了（运费/试验费/退货/过磅差…），
    # 实发要**减去**它。实例 25MT-07F591-GYL：「抵扣运费与试验费用：1109」，
    # 合计 170,213.04 − 1,109 = 169,104.04，与飞书分毫一致
    # （同一条还有「本次付尾款 135,104.04」= 169,104.04 − 预付款 34,000，可交叉验证）。
    DEDUCT_KEYWORDS = ("抵扣", "扣款", "扣减", "冲减", "减免", "扣除", "减掉", "退回")
    # 付款类词（预付款/尾款/已付…）：出现在同一行时，说明它是**付款说明**而不是费用扣减
    PAYMENT_WORDS = ("预付款", "预付", "尾款", "已付", "付款", "定金", "订金", "货款", "本次付")
    COST_WORDS = (
        "木箱", "样品", "样管", "试样", "坡口", "绳子", "吊绳", "包装", "加工",
        "成品", "价格", "运费", "装卸", "实验", "检测",
    )
    tail_notes: list[dict[str, Any]] = []
    tail_cost = Decimal(0)
    for index in range(last_total + 1, footer_start):
        row = grid[index]
        joined = " ".join(cell for cell in row if cell)
        if not joined or "制单" in joined or "审核" in joined:
            continue
        key_col = next(
            (
                col
                for col, cell in enumerate(row)
                if cell and any(word in cell for word in TAIL_KEYWORDS)
            ),
            None,
        )
        if key_col is None:
            continue
        amount = None
        for cell in row[key_col + 1:]:
            if cell and is_number_cell(str(cell).strip()):
                value = to_decimal(cell)
                if value:
                    amount = value
                    break
        if amount is None:
            continue
        tail_notes.append({"label": joined[:80], "amount": dec_str(amount), "row": index + 1})
        if any(word in joined for word in COST_WORDS):
            extra_fees.append(
                {"label": f"表尾补款：{joined[:60]}", "amount": dec_str(amount)}
            )
            tail_cost += amount

    # ---- 表尾抵扣：从合计里**减掉**（负向费用），同样在 extra_fees / tail_notes 留痕
    deduct_total = Decimal(0)
    for index in range(last_total + 1, footer_start):
        row = grid[index]
        joined = " ".join(cell for cell in row if cell)
        if not joined or "制单" in joined or "审核" in joined:
            continue
        if not any(word in joined for word in DEDUCT_KEYWORDS):
            continue
        # 只认**费用性质的抵扣**（抵扣运费/试验费/木箱费…）。
        # 「本次预付款抵扣 159,223.75」这类是付款信息（预付款/尾款/已付），
        # 口径上不进成本——动了它反而把成本算错（26MT-06M087 就这么被扣掉过 15.9 万）。
        if not any(word in joined for word in COST_WORDS):
            tail_notes.append({"label": joined[:80], "amount": None, "row": index + 1})
            continue
        if any(word in joined for word in PAYMENT_WORDS):
            tail_notes.append({"label": joined[:80], "amount": None, "row": index + 1})
            continue
        amount = None
        key_col = next(
            (col for col, cell in enumerate(row)
             if cell and any(word in cell for word in DEDUCT_KEYWORDS)),
            None,
        )
        # ① 先看关键词右侧的独立数字单元格
        if key_col is not None:
            for cell in row[key_col + 1:]:
                if cell and is_number_cell(str(cell).strip()):
                    value = to_decimal(cell)
                    if value:
                        amount = value
                        break
        # ② 金额常写在**同一个单元格的文字里**（「抵扣运费与试验费用：1109」）→ 取冒号后/
        #    文本里最后一段数字（排除 2026.4.23 这类日期）
        if amount is None:
            text = str(row[key_col] if key_col is not None else joined)
            tail = re.split(r"[:：]", text)[-1] if re.search(r"[:：]", text) else text
            numbers = re.findall(r"\d+(?:[.,]\d+)?", tail)
            for token in reversed(numbers):
                if re.fullmatch(r"20\d{2}", token):  # 年份不算金额
                    continue
                value = to_decimal(token)
                if value:
                    amount = value
                    break
        if amount is None:
            continue
        if any(note.get("label") == joined[:80] for note in tail_notes):
            continue  # 上面补款分支已经记过这一行
        deduct_total += amount
        extra_fees.append(
            {"label": f"表尾抵扣：{joined[:60]}", "amount": dec_str(-amount)}
        )
        tail_notes.append(
            {"label": joined[:80], "amount": dec_str(-amount), "row": index + 1}
        )
    tail_cost -= deduct_total

    if tail_cost:
        # 作为一条独立区块参与下游合计（区块标签不含批次字母，不会被误当批次）
        settled_blocks.append(
            {
                "label": "表尾补款",
                "amount": dec_str(tail_cost),
                "qty": None,
                "weight": None,
            }
        )

    settled_sum = Decimal(0)
    settled_qty_sum = Decimal(0)
    for block in settled_blocks:
        value = to_decimal(block.get("amount"))
        if value is not None:
            settled_sum += value
        qty_value = to_decimal(block.get("qty"))
        if qty_value is not None:
            settled_qty_sum += qty_value
    if settled_sum:
        # 表尾补款/抵扣已经作为独立区块（label="表尾补款"）计进 settled_blocks 了，
        # 这里**不能再加一次** tail_cost——曾经两边各加一次，金额被放大了整整一笔
        # （实测 26MT-06M087 多 27,079、25MT-07V034-JX 多 19,569.6、
        #   25MT-08C355B-MJ 多 2,933.5、25MT-07F591-GYL 多 1,109）。
        settled_amount = dec_str(settled_sum)
    else:
        settled_amount = (
            dec_str((to_decimal(order_amount) or Decimal(0)) + tail_cost)
            if tail_cost
            else order_amount
        )

    # ---- 表尾：合计之后（付款信息，不进成本）
    prepaid: Decimal | None = None
    payable: Decimal | None = None
    notes = ""
    for row in grid[last_total + 1:]:
        joined = " ".join(cell for cell in row if cell)
        if not joined:
            continue
        if "备注" in joined and not notes:
            notes = joined
            match = re.search(r"预付款\s*([\d,.]+)\s*(万元|元)?", joined)
            if match:
                value = Decimal(match.group(1).replace(",", ""))
                prepaid = value * 10000 if match.group(2) == "万元" else value
        if "本次付" in joined:
            numbers = [to_decimal(cell) for cell in row if to_decimal(cell) is not None]
            if numbers:
                payable = numbers[-1]
                if prepaid is None and len(numbers) > 1:
                    prepaid = numbers[0]
    for row in grid[data_start:stop]:
        if any(cell.startswith("备注") for cell in row) and not notes:
            notes = " ".join(cell for cell in row if cell)

    payload = {
        "doc_type": "入库单",
        "title": title,
        "contract_no": _contract_of(title),
        "batch": _batch_of(title),
        "supplier": supplier,
        "item_name": item_name,
        "doc_date": doc_date,
        "columns": labels,
        "lines": lines,
        "totals": totals,
        "extra_fees": extra_fees,
        "tail_notes": tail_notes,
        "tail_cost": dec_str(tail_cost) if tail_cost else None,
        "aux_rows": aux_rows,
        "fee_by_column": {k: dec_str(v) for k, v in fee_by_column.items()},
        "_fee_items": {
            key: [dec_str(value) for value in values] for key, values in fee_items.items()
        },
        "order_amount": order_amount,
        "settled_amount": settled_amount,
        "settled_blocks": settled_blocks,
        "_settled_labels": sorted(settled_labels),
        "settled_qty": dec_str(settled_qty_sum) if settled_qty_sum else None,
        "prepaid": dec_str(prepaid),
        "payable_now": dec_str(payable),
        "notes": notes,
        "source": "rule",
    }

    problems, order_notes = _validate(payload)
    payload["anomalies"] = problems + order_notes
    return RuleResult(not problems, payload, "；".join(problems))


def _subset_matches(values: list[Decimal], target: Decimal) -> bool:
    """target 是否恰好等于 values 的某个非空子集之和（用于「部分费用行进了合计」）。"""
    if not values or len(values) > 15:
        return False
    reachable: set[Decimal] = {Decimal(0)}
    for value in values:
        for current in list(reachable):
            reachable.add(current + value)
    reachable.discard(Decimal(0))
    return any(abs(candidate - target) <= Decimal("0.05") for candidate in reachable)


def _row_amount(row: list[str], labels: list[str]) -> Decimal | None:
    """取一行里金额列的值（优先「金额（元）」）。"""
    best: Decimal | None = None
    for col, label in enumerate(labels):
        if "金额" not in label or "单价" in label:
            continue
        if col >= len(row):
            continue
        value = to_decimal(row[col])
        if value is None:
            continue
        if "元" in label and "金额（元）" in label:
            return value
        best = value if best is None else best
    return best


def _validate(payload: dict[str, Any]) -> tuple[list[str], list[str]]:
    """校验明细与合计是否闭合；返回（阻断问题, 提示性说明）。

    允许「合计 = 明细求和 + 表内费用行」——包装费/木箱费/补款这类增量
    在单据里就是单独一行、并进了合计。

    只把**实发区块**的不一致当阻断问题：我们取的成本就是实发合计；
    订单数据区（下单时的预期金额）对不齐不动我们的结果，只留提示。
    """
    problems: list[str] = []
    notes: list[str] = []
    lines = payload.get("lines") or []
    totals = payload.get("totals") or {}
    settled_labels = set(payload.get("_settled_labels") or [])
    if not lines:
        problems.append("没有解析出明细行")

    extras = Decimal(0)
    for item in payload.get("extra_fees") or []:
        extras += to_decimal(item.get("amount")) or Decimal(0)
    payload["fee_delta"] = dec_str(extras) if extras else None
    fee_by_column = {
        key: (to_decimal(value) or Decimal(0))
        for key, value in (payload.get("fee_by_column") or {}).items()
    }
    fee_items = {
        key: [value for value in (to_decimal(item) for item in values) if value is not None]
        for key, values in (payload.get("_fee_items") or {}).items()
    }

    deltas: dict[str, str] = {}
    for label, value in totals.items():
        if "金额" not in label or "单价" in label:
            continue
        total = to_decimal(value)
        if total is None:
            continue
        summed = Decimal(0)
        for row in lines:
            for row_label, row_value in row.items():
                if row_label == label:
                    number = to_decimal(row_value)
                    if number is not None:
                        summed += number
                    break
        if summed == 0:
            continue
        diff = total - summed
        if abs(diff) <= Decimal("0.05"):
            continue
        # 费用行可能是加项（木箱费）也可能是减项（过磅差抵扣、合同赔款、做负奖励），
        # 两种符号都放行，所以比的是绝对值
        if label in fee_by_column and abs(abs(diff) - abs(fee_by_column[label])) <= Decimal("0.05"):
            continue
        if abs(abs(diff) - abs(extras)) <= Decimal("0.05"):
            deltas[label] = dec_str(diff) or ""
            continue
        # 有时只有**一部分**费用行进了合计（另一行是「补款/备注」，没并进合计），
        # 这时差额恰好等于本列费用行的一个子集——按子集精确匹配，不放宽容差
        if _subset_matches(fee_items.get(label) or [], diff):
            notes.append(
                f"「{label}」合计 {total} = 明细求和 {summed} + 部分费用行 {diff}（其余费用行未计入合计）"
            )
            continue
        message = f"「{label}」合计 {total} ≠ 明细求和 {summed}（差 {diff}）"
        if settled_labels and label not in settled_labels:
            notes.append(f"{message}（订单数据区，仅供参考）")
        elif abs(diff) <= max(Decimal("50"), abs(total) * Decimal("0.001")):
            # 单据自身的小残差（补款备注没进合计、四舍五入），按最后一个合计行取值即可
            notes.append(f"{message}（单据自身小残差，按最终合计行取值）")
        else:
            problems.append(message)
    if deltas:
        payload["fee_delta_columns"] = deltas
    return problems, notes


def _contract_of(title: str) -> str:
    match = re.search(r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:Y)?(?:[- ]?ADD\d*)?", title)
    return match.group(0) if match else ""


def _batch_of(title: str) -> str | None:
    official = re.search(r"第([一二三四五六七八九十\d]+)批", title)
    if official:
        return official.group(1)
    match = re.search(
        r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:Y)?(?:[- ]?ADD\d*)?[\s\-]*([A-Za-z0-9]{1,4})(?![\w])",
        title,
    )
    return match.group(1) if match else None


# --------------------------------------------------------------------------- 模型解析


def parse_with_model(
    attachment: Attachment,
    *,
    purchase_code: str,
    file_name: str,
    hint: str = "",
) -> dict[str, Any]:
    """把附件交给 DeepSeek 解析。表格给文本，图片给图。"""
    if attachment.kind == "table":
        body = sheets_to_text(attachment.sheets)
        if attachment.embedded:
            # 表格里几乎没有文字、但嵌了扫描件：图才是真正的入库单
            prompt = (
                f"采购单号（睿贝档案，仅作参考）：{purchase_code}\n"
                f"附件文件名：{file_name}\n"
                f"{hint}\n"
                "这张 Excel 的单元格几乎是空的，真正的入库单是**嵌在里面的图片**，"
                "请照图片读，金额以图片里的最终合计为准。\n"
                f"（附带表格里残留的文字如下，供定位：\n{body}）"
            )
            images = list(attachment.embedded)
        else:
            prompt = (
                f"采购单号（睿贝档案，仅作参考）：{purchase_code}\n"
                f"附件文件名：{file_name}\n"
                f"{hint}\n"
                f"下面是该附件的表格内容（rN 表示第 N 行）：\n\n{body}"
            )
            # 表格里嵌了扫描件/批注截图时一并给模型看：合计下边的「抵扣/补款」批注、
            # 手写改动经常只出现在图里（实测 26MT-07T265 这类回签件就是）
            images = list(attachment.embedded)
            if images:
                prompt += (
                    f"\n\n（另附该 Excel 内嵌的 {len(images)} 张图片，"
                    "请以图片里的金额与批注为准，尤其注意合计行下方的抵扣/补款说明。）"
                )
    elif attachment.kind == "image":
        prompt = (
            f"采购单号（睿贝档案，仅作参考）：{purchase_code}\n"
            f"附件文件名：{file_name}\n"
            f"{hint}\n"
            "下面是该附件的图片，请照着读。"
        )
        images = list(attachment.images)
    elif attachment.kind == "pdf":
        prompt = (
            f"采购单号（睿贝档案，仅作参考）：{purchase_code}\n"
            f"附件文件名：{file_name}\n"
            f"{hint}\n"
            f"下面是 PDF 抽取的文本：\n\n{attachment.text}"
        )
        images = []
    else:
        return {"anomalies": [f"不支持的附件类型：{attachment.note}"], "source": "skip"}

    payload = deepseek_client.chat_json(SYSTEM_PROMPT, prompt, images=images)
    if not isinstance(payload, dict):
        payload = {"anomalies": ["模型返回结构异常"], "raw": payload}
    payload["source"] = "llm"
    payload.setdefault("anomalies", [])
    normalize_settled(payload)
    return payload


def normalize_settled(payload: dict[str, Any]) -> None:
    """让 `settled_totals.amount` 始终等于各实发区块之和。

    模型偶尔把「第一个区块的金额」当成整单实发合计填进 `settled_totals`
    （实测 26MT-02N182-B/-C：填的只是 A 批 194539.9，区块求和其实是
    246623.12 / 270291.16）。下游取数一律走 `settled_blocks`，这里统一校正，
    避免两处数字打架，并在 anomalies 里留痕。

    另外：**没有实发区块**的附件（图片/扫描件的入库单通常只有一张明细表，
    模型把每行金额都读对了，却没填合计）用「明细求和」兜底，避免整单取不到金额。
    """
    blocks = payload.get("settled_blocks") or []
    values = [
        to_decimal(block.get("amount")) for block in blocks if isinstance(block, dict)
    ]
    values = [value for value in values if value is not None]
    totals = payload.get("settled_totals")
    if not isinstance(totals, dict):
        totals = {}
        payload["settled_totals"] = totals
    if not values:
        declared = to_decimal(totals.get("amount"))
        if declared:
            return
        # 注意：整张表只有一个合计（订单/实发不分栏）时，那个合计已经落在
        # order_amount 上，且通常**含费用行**；此时用"明细求和"兜底会算少
        # （实测 26MT-10E115：明细 48931.44 vs 合计 53431.44）。只有连订单
        # 金额都取不到时才用明细求和。
        order = to_decimal(payload.get("order_amount"))
        if order is None:
            order = to_decimal((payload.get("order_totals") or {}).get("amount"))
        if order:
            return
        detail = line_amounts(payload)
        if detail:
            totals["amount"] = dec_str(detail)
            payload.setdefault("anomalies", []).append(
                f"附件没有实发区块，实发合计按明细求和 {detail} 兜底"
            )
        return
    total = sum(values, Decimal(0))
    declared = to_decimal(totals.get("amount"))
    if declared is None or abs(declared - total) > Decimal("0.05"):
        totals["amount"] = dec_str(total)
        payload.setdefault("anomalies", []).append(
            f"模型填写的实发合计 {declared} 与各实发区块之和 {total} 不一致，"
            "已按区块之和修正"
        )


def extract_attachment(
    path: Path,
    *,
    purchase_code: str,
    use_model: bool = True,
    force_model: bool = False,
    force: bool = False,
    hint: str = "",
) -> dict[str, Any]:
    """解析单个附件：规则优先 → 校验不过或强制时交给模型。"""
    attachment = read_attachment(path)
    file_name = path.name
    digest = attachment.digest
    cached = None if (force or hint) else load_cached(purchase_code, file_name, digest)
    # 上一次是在「不允许用模型」的情况下跑出的失败/跳过结果，这次允许用模型就要重跑
    reusable = cached is not None and (
        cached.get("_model_allowed") or not use_model
    )
    if reusable:
        return cached

    payload: dict[str, Any]
    if attachment.kind == "table" and not force_model:
        result = parse_template(attachment)
        # 「结构读通了、但一个金额都没取到」也当失败处理（这类表的数据往往在嵌入图里）
        empty_amount = not (result.payload.get("settled_amount") or result.payload.get("order_amount"))
        if result.ok and not empty_amount:
            payload = result.payload
            payload["_how"] = "rule"
        elif use_model:
            payload = parse_with_model(
                attachment,
                purchase_code=purchase_code,
                file_name=file_name,
                hint=hint
                or (
                    f"规则解析失败（{result.reason}），请以表格内容为准重新读取。"
                    if not result.ok
                    else "规则只读出了表格结构、没读到任何金额，请重新读取（若数据在嵌入的图片里，以图片为准）。"
                ),
            )
            payload["_how"] = "llm_fallback"
        else:
            payload = result.payload
            payload["_how"] = "rule_failed"
    elif use_model:
        payload = parse_with_model(
            attachment,
            purchase_code=purchase_code,
            file_name=file_name,
            hint=hint or "请以附件内容为准读取。",
        )
        payload["_how"] = "llm"
    else:
        payload = {"anomalies": [attachment.note or "未解析"], "_how": "skipped"}

    payload.update(
        {
            "_purchase_code": purchase_code,
            "_file": file_name,
            "_digest": digest,
            "_kind": attachment.kind,
            "_note": attachment.note,
            "_model_allowed": bool(use_model),
        }
    )
    save_cached(purchase_code, file_name, payload)
    return payload
