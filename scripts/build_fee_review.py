r"""把「含额外费用」的入库单全部拉成一张待核表，供人工排查。

用法：
    & '.\.venv\Scripts\python.exe' scripts\build_fee_review.py

输出：`outputs/cost_match/额外费用待核.xlsx`

    费用待核  一行 = 一张入库单 × 一张报关单（同一张入库单涉及几张报关单就几行）
    说明      口径说明（怎么算费用、怎么判出运完成、候选落位规则是什么）

判据（全部只陈述事实，不做假设）：
  · 无归属费用 = 入库单「实发合计 − 产品明细金额之和」
  · 出运完成    = 该采购单**全部出运单**的产品行，按「规格 + 数量」覆盖该入库单每一行货；
                 全覆盖 = 已发完；有行覆盖不到 = 未发完；行认不出规格 = 无法判定
                 出运单清单来自出运单头原生的 `purchaseCode`（三单互查）
  · 飞书落位   = 飞书逐张报关单「采购金额 − 我方货值」= 飞书实际多给的钱
  · 候选落位   = 货值最大 / 报关金额最大（**未经验证的假设**，只作参考）
"""
from __future__ import annotations

import collections
import json
import re
import sys
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from app.services.cost_match import dec, unit_core  # noqa: E402
from app.services.erp_cache import CACHE_ROOT, read_jsonl  # noqa: E402
from app.services.grn_extract import PARSED_ROOT, line_amounts, safe_name  # noqa: E402
from app.services.grn_select import select  # noqa: E402
from app.services.shipment_index import canonical  # noqa: E402

SPLIT = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
REF = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs_full.json"
OUT = PROJECT_ROOT / "outputs" / "cost_match" / "额外费用待核.xlsx"
TOL = Decimal("1")

FULL_RE = re.compile(
    r"^\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:Y)?\s*[-_]?\s*([A-Z])(?:\s*[-_]?ADD\d*)?$"
)
TAIL_RE = re.compile(r"^\d{2,3}([A-Z])(?:\s*[-_]?ADD\d*)?$")
SPEC_TEXT = re.compile(r"(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)")

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
WARN_FILL = PatternFill("solid", fgColor="FFE699")
HEADER_FONT = Font(bold=True)


def read_rows(path: Path, sheet: str | None = None) -> list[dict]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    target = workbook[sheet] if sheet else workbook.active
    data = list(target.iter_rows(values_only=True))
    header = [str(cell) for cell in data[0]]
    workbook.close()
    return [dict(zip(header, row)) for row in data[1:]]


def num(value) -> Decimal | None:
    try:
        if value in (None, ""):
            return None
        return Decimal(str(value).replace(",", "").strip())
    except Exception:  # noqa: BLE001
        return None


def row_amount(line: dict) -> Decimal:
    """入库单明细行的金额（兼容中文列名与英文字段）。"""
    for key in ("settled_amount", "amount", "金额", "金额 元", "含税金额（元）"):
        if line.get(key) not in (None, ""):
            value = dec(line.get(key))
            if value:
                return value
    for key, raw in line.items():
        if "金额" in str(key) and "单价" not in str(key):
            value = dec(raw)
            if value:
                return value
    return Decimal(0)


def row_qty(line: dict) -> Decimal:
    for key in ("数量 总米数", "数量 米", "数量", "qty", "settled_qty"):
        if line.get(key) not in (None, ""):
            value = num(line.get(key))
            if value:
                return value
    for key, value in line.items():
        if str(key).startswith(("数量", "米数", "支数", "只数", "件数")):
            number = num(value)
            if number:
                return number
    return Decimal(0)


def spec_key(line: dict) -> str:
    """规格键 = 外径|壁厚。"""
    od = wt = None
    for key, value in line.items():
        name = str(key).upper().replace(" ", "")
        if od is None and name.startswith("外径"):
            od = num(value)
        elif wt is None and "壁厚" in name and "公差" not in name:
            wt = num(value)
    if od is None or wt is None:
        match = SPEC_TEXT.search(str(line.get("spec") or ""))
        if match:
            od = od if od is not None else num(match.group(1))
            wt = wt if wt is not None else num(match.group(2))
    if od is None and wt is None:
        return ""
    left = f"{od.normalize():f}" if od is not None else "?"
    right = f"{wt.normalize():f}" if wt is not None else "?"
    return f"{left}×{right}"


def strict_batches(text: str) -> set[str]:
    """合同号里严格标记的批次字母（尾部独立字母）。"""
    out: set[str] = set()
    for part in re.split(r"[&,，;；]", str(text or "")):
        part = part.strip()
        for pattern in (FULL_RE, TAIL_RE):
            match = pattern.match(part)
            if match:
                out.add(match.group(1).upper())
                break
    return out


def load_shipments() -> dict[str, dict]:
    """采购单号（归一）→ {出运单号（归一）: {invoice, date, status}}。"""
    out: dict[str, dict] = collections.defaultdict(dict)
    for row in read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl"):
        invoice = str(row.get("invoiceCode") or "")
        raw = str(row.get("purchaseCode") or "")
        codes = [c.strip() for c in raw.replace("&", ",").replace(";", ",").split(",") if c.strip()]
        for code in codes:
            out[canonical(code)][canonical(invoice)] = {
                "invoice": invoice,
                "date": str(row.get("shipDate") or ""),
                "status": str(row.get("statusName") or ""),
            }
    return out


def load_lines() -> dict[str, list[dict]]:
    """出运单号（归一）→ 产品行。"""
    out: dict[str, list[dict]] = collections.defaultdict(list)
    for path in sorted((CACHE_ROOT / "details" / "shipments").glob("*.json")):
        payload = json.loads(path.read_text(encoding="utf-8"))
        keys = {
            canonical(str(payload.get("invoiceCode") or "")),
            canonical(str((payload.get("baseInfo") or {}).get("外销发票号") or "")),
            canonical(path.stem),
        }
        keys.discard("")
        for key in keys:
            out[key].extend(payload.get("productList") or [])
    return out


def shipment_state(payload: dict, shipped_lines: list[dict]) -> tuple[str, str]:
    """(出运完成判定, 缺口说明)。"""
    qty_by_spec: dict[str, Decimal] = collections.defaultdict(Decimal)
    amount_by_spec: dict[str, Decimal] = collections.defaultdict(Decimal)
    for line in shipped_lines:
        key = spec_key(line)
        if key:
            qty_by_spec[key] += dec(line.get("出运数量"))
            amount_by_spec[key] += dec(line.get("出运采购金额(RMB)"))
    missing: list[str] = []
    unknown = 0
    for line in payload.get("lines") or []:
        amount = row_amount(line)
        if amount <= TOL:
            continue
        key = spec_key(line)
        qty = row_qty(line)
        if not key or (qty <= 0 and not qty_by_spec.get(key)):
            unknown += 1
            continue
        if qty_by_spec.get(key, Decimal(0)) + Decimal("0.01") >= qty:
            continue
        if amount_by_spec.get(key, Decimal(0)) + TOL >= amount:
            continue
        missing.append(f"{key}：需 {qty or amount}，已出运 {qty_by_spec.get(key, 0) or amount_by_spec.get(key, 0)}")
    if missing:
        return "未发完", "；".join(missing[:3])
    if unknown:
        return "无法判定", f"{unknown} 行货认不出规格"
    return "已发完", ""


def carrying_shipments(payload: dict, lines_with_invoice: list[tuple[str, dict]]) -> tuple[set[str], int]:
    """这张入库单的货是被哪几张出运单装走的（只看这张入库单的行）。

    逐行按「规格」找出运行：只有一张出运单有该规格 → 就是它；多张都有时用金额
    再认一次（该行金额等于某张单该规格的金额）；仍认不出就把候选都算上。
    返回（出运单集合, 认不出规格的货行数）。
    """
    carriers: set[str] = set()
    unknown = 0
    for line in payload.get("lines") or []:
        amount = row_amount(line)
        if amount <= TOL:
            continue
        key = spec_key(line)
        if not key:
            unknown += 1
            continue
        candidates = [
            (invoice, bline)
            for invoice, bline in lines_with_invoice
            if spec_key(bline) == key
        ]
        invoices = {invoice for invoice, _line in candidates}
        if len(invoices) <= 1:
            carriers |= invoices
            continue
        by_invoice: dict[str, Decimal] = collections.defaultdict(Decimal)
        for invoice, bline in candidates:
            by_invoice[invoice] += dec(bline.get("出运采购金额(RMB)"))
        exact = [inv for inv, value in by_invoice.items() if abs(value - amount) <= TOL]
        if len(exact) == 1:
            carriers.add(exact[0])
        else:
            carriers |= invoices  # 认不出是哪一张 → 保守地把候选都算上
    return carriers, unknown


def main() -> None:
    rows = read_rows(SPLIT)
    ref = json.loads(REF.read_text(encoding="utf-8"))["records"]
    shipments = load_shipments()
    lines_by_invoice = load_lines()

    feishu: dict[tuple[str, str, str], Decimal] = {}
    for record in ref:
        decl = str(record.get("报关单号") or "").strip()
        if decl:
            feishu[
                (
                    decl,
                    str(record.get("供应商简称") or "").strip(),
                    unit_core(record.get("合同号（应收表格）") or ""),
                )
            ] = dec(record.get("采购金额"))

    by_po: dict[str, list[dict]] = collections.defaultdict(list)
    for row in rows:
        by_po[str(row.get("采购单号") or "").strip()].append(row)

    out_rows: list[dict] = []
    for code, group in sorted(by_po.items()):
        vendor = str(group[0].get("供应商简称") or "")
        goods: dict[str, Decimal] = collections.defaultdict(Decimal)
        declared: dict[str, Decimal] = collections.defaultdict(Decimal)
        contract: dict[str, str] = {}
        for row in group:
            decl = str(row.get("报关单号") or "")
            goods[decl] += dec(row.get("出运采购金额合计"))
            declared[decl] += dec(row.get("报关金额"))
            contract[decl] = str(row.get("合同号_1") or "")
        po_shipments = shipments.get(canonical(code), {})
        shipped_lines: list[dict] = []
        lines_with_invoice: list[tuple[str, dict]] = []
        decl_date: dict[str, str] = {}
        decl_invoice: dict[str, str] = {}
        for key, info in po_shipments.items():
            for line in lines_by_invoice.get(key, []):
                if canonical(str(line.get("采购订单号") or "")) == canonical(code):
                    shipped_lines.append(line)
                    lines_with_invoice.append((info["invoice"], line))
            for decl, text in contract.items():
                if canonical(text) == key:
                    decl_date.setdefault(decl, info["date"])
                    decl_invoice.setdefault(decl, info["invoice"])
        for item in select(code)["kept"]:
            path = PARSED_ROOT / safe_name(code) / (safe_name(item["file"], "file") + ".json")
            if not path.exists():
                continue
            payload = json.loads(path.read_text(encoding="utf-8"))
            detail_sum = line_amounts(payload)
            fee = dec(item["amount"]) - detail_sum
            if not detail_sum or fee <= TOL:
                continue
            state, gap = shipment_state(payload, shipped_lines)
            carriers, unknown_rows = carrying_shipments(payload, lines_with_invoice)
            # 「货全在一次出运里」的判定（2026-09-24 用户口径）：
            # **三单互查（出运单头原生 purchaseCode）查到该采购单的出运单只有 1 张**
            # → 费用只能落在这一次出运（那一张报关单）上，不存在"归到哪一票"的问题。
            # 其余信号（该采购单有几张报关单、按规格匹配到几张出运单）只作为参考列。
            single_by_decl = len(goods) == 1
            single_by_shipment = len(po_shipments) == 1
            single_by_line = len(carriers) == 1 and unknown_rows == 0
            single_shipment = single_by_shipment
            if single_by_shipment:
                carrier_note = f"该采购单只有一张出运单（{next(iter(po_shipments.values()))['invoice']}）"
            elif single_by_decl:
                carrier_note = "该采购单只有一张报关单"
            elif single_by_line:
                carrier_note = f"按规格匹配到唯一出运单：{next(iter(carriers))}"
            elif carriers:
                carrier_note = "、".join(sorted(carriers))
            else:
                carrier_note = "出运单匹配不上（无法判定）"
            fee_detail = "；".join(
                f"{entry.get('label')} {entry.get('amount')}"
                for entry in payload.get("extra_fees") or []
                if isinstance(entry, dict)
            )
            excess = {d: feishu.get((d, vendor, unit_core(code)), Decimal(0)) - goods[d] for d in goods}
            whole = [d for d, value in excess.items() if abs(value - fee) <= TOL]
            others_zero = all(abs(excess[d]) <= TOL for d in goods if d not in whole)
            if len(whole) == 1 and others_zero:
                placement = f"整笔贴 {whole[0]}"
            elif all(abs(excess[d]) <= TOL for d in goods):
                placement = "飞书未含该费用"
            elif all(
                abs(excess[d] - fee * goods[d] / sum(goods.values(), Decimal(1))) <= max(TOL, Decimal(2))
                for d in goods
            ):
                placement = "飞书按货值比例摊开"
            else:
                placement = "其他（差额另有原因）"
            best_goods = max(goods, key=lambda d: goods[d])
            best_declared = max(declared, key=lambda d: declared[d])
            for decl in sorted(goods, key=lambda d: -goods[d]):
                out_rows.append(
                    {
                        "采购单号": code,
                        "供应商简称": vendor,
                        "入库单文件": item["original"],
                        "无归属费用": str(fee),
                        "费用行原文": fee_detail,
                        "入库单实发合计": str(dec(item["amount"])),
                        "产品明细求和": str(detail_sum),
                        "出运是否完成": state,
                        "出货缺口": gap,
                        "涉及出运单数": len(carriers),
                        "涉及出运单": carrier_note,
                        "三单互查出运单数": len(po_shipments),
                        "按规格匹配到出运单数": len(carriers),
                        "认不出规格的货行": unknown_rows,
                        "出运单": decl_invoice.get(decl, "—"),
                        "出运日期": decl_date.get(decl, "—"),
                        "报关单号": decl,
                        "合同号": contract.get(decl, ""),
                        "批次字母": "、".join(sorted(strict_batches(contract.get(decl, "")))) or "无",
                        "我方货值": str(goods[decl]),
                        "报关金额": str(declared[decl]),
                        "飞书采购金额": str(feishu.get((decl, vendor, unit_core(code)), "")),
                        "飞书多给": str(excess[decl]),
                        "飞书落位": placement,
                        "本单是货值最大": "是" if decl == best_goods else "",
                        "本单是报关金额最大": "是" if decl == best_declared else "",
                        "待确认": (
                            "费用整笔归到本单？"
                            if decl == best_goods and state == "已发完" and len(goods) > 1
                            else ""
                        ),
                        "_单次出运": single_shipment,
                    }
                )

    # 一张入库单的货全在一次出运里 → 不存在"费用归到哪一票"的问题，单列一张表
    multi = [row for row in out_rows if not row["_单次出运"]]
    single = [row for row in out_rows if row["_单次出运"]]
    # ---- 差额拆解：验证财务说的「按金额分摊」----
    grouped = collections.defaultdict(list)
    for row in out_rows:
        grouped[(row["采购单号"], row["入库单文件"])].append(row)
    breakdown: list[dict] = []
    for (_code, _file), items in grouped.items():
        fee = dec(items[0]["无归属费用"])
        total_goods = sum((dec(r["我方货值"]) for r in items), Decimal(0))
        total_declared = sum((dec(r["报关金额"]) for r in items), Decimal(0))
        total_excess = sum((dec(r["飞书多给"]) for r in items), Decimal(0))
        for row in items:
            goods = dec(row["我方货值"])
            declared = dec(row["报关金额"])
            excess = dec(row["飞书多给"])
            share_goods = (goods / total_goods) if total_goods else Decimal(0)
            share_declared = (declared / total_declared) if total_declared else Decimal(0)
            breakdown.append(
                {
                    "采购单号": row["采购单号"],
                    "供应商简称": row["供应商简称"],
                    "入库单文件": row["入库单文件"],
                    "无归属费用(我方算)": str(fee),
                    "飞书实际含的费用(Σ多给)": str(total_excess),
                    "差额(我方算−飞书含)": str(fee - total_excess),
                    "报关单号": row["报关单号"],
                    "合同号": row["合同号"],
                    "出运日期": row["出运日期"],
                    "出运是否完成": row["出运是否完成"],
                    "我方货值": str(goods),
                    "货值占比": f"{share_goods:.6f}",
                    "报关金额": str(declared),
                    "报关金额占比": f"{share_declared:.6f}",
                    "飞书多给": str(excess),
                    "按货值分摊应得": str((fee * share_goods).quantize(Decimal("0.01"))),
                    "按报关金额分摊应得": str((fee * share_declared).quantize(Decimal("0.01"))),
                    "与按货值分摊的差": str((excess - fee * share_goods).quantize(Decimal("0.01"))),
                    "与按报关金额分摊的差": str(
                        (excess - fee * share_declared).quantize(Decimal("0.01"))
                    ),
                }
            )
    workbook = Workbook()
    columns = [
        name for name in (list(out_rows[0].keys()) if out_rows else []) if not name.startswith("_")
    ]
    for position, (title, data_rows) in enumerate(
        (("费用待核", multi), ("单次出运（已排除）", single), ("差额拆解", breakdown))
    ):
        sheet = workbook.active if position == 0 else workbook.create_sheet()
        sheet.title = title
        sheet_columns = (
            columns
            if title != "差额拆解"
            else [name for name in breakdown[0].keys()] if breakdown else columns
        )
        sheet.append(sheet_columns)
        for cell in sheet[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        for row in data_rows:
            sheet.append([row.get(name, "") for name in sheet_columns])
        for index, name in enumerate(sheet_columns, 1):
            sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
                10, min(34, len(name) * 2 + 6)
            )
        sheet.freeze_panes = "A2"
    print(
        f"费用待核 {len(multi)} 行（{len({r['入库单文件'] for r in multi})} 张入库单）；"
        f"单次出运已排除 {len(single)} 行（{len({r['入库单文件'] for r in single})} 张入库单）"
    )
    # 分摊口径分类（只对多出运单案例，且飞书确实给了费用的）
    counts = collections.Counter()
    seen: set[tuple] = set()
    for row in breakdown:
        key = (row["采购单号"], row["入库单文件"])
        if key in seen:
            continue
        seen.add(key)
        items = [r for r in breakdown if (r["采购单号"], r["入库单文件"]) == key]
        fee = dec(items[0]["无归属费用(我方算)"])
        excess = dec(items[0]["飞书实际含的费用(Σ多给)"])
        if not any(abs(dec(r["飞书多给"])) > TOL for r in items):
            counts["飞书完全没含费用"] += 1
            continue
        if abs(excess - fee) > TOL:
            counts["飞书含的费用 ≠ 我方算的费用（金额本身不一致）"] += 1
            continue
        by_goods = all(abs(dec(r["与按货值分摊的差"])) <= Decimal("2") for r in items)
        by_declared = all(abs(dec(r["与按报关金额分摊的差"])) <= Decimal("2") for r in items)
        if by_goods and by_declared:
            counts["按货值分摊 = 按报关金额分摊（两者等价）"] += 1
        elif by_goods:
            counts["按货值（人民币）分摊"] += 1
        elif by_declared:
            counts["按报关金额（美元）分摊"] += 1
        else:
            counts["两种分摊都对不上"] += 1
    print("差额拆解口径分类（多出运单案例）：")
    for name, value in counts.most_common():
        print(f"   {value:>4} {name}")

    notes = workbook.create_sheet("说明")
    for line in (
        "本表只为人工排查用，不含任何猜测性规则。",
        "一行 = 一张入库单 × 一张报关单（同一张入库单涉及几张报关单就几行）。",
        "",
        "无归属费用 = 入库单「实发合计 − 产品明细金额之和」，即单据里的绳子费/管帽费/木箱费/",
        "木轴费/托盘费这类不属于任何产品行的独立费用行。",
        "费用行原文：解析出来的费用行标签与金额（可能比 实发−明细 小，差额是没并入实发的补款）. ",
        "数据来源：出运单头原生 purchaseCode（三单互查）+ 出运单明细 productList + 入库单解析缓存。",
        "",
        "出运是否完成的判定：把该采购单全部出运单的产品行按「规格（外径×壁厚）+ 数量」汇总，",
        "再逐行核对这张入库单的每一行货是否都被覆盖（数量够或金额够）；",
        "全覆盖=已发完；有行覆盖不到=未发完（缺口列写明差多少）；行认不出规格=无法判定。",
        "注意：不看出运单的「未出运」状态字段（实测该字段与是否真出运不一致）。",
        "",
        "飞书多给 = 飞书该报关单的采购金额 − 我方该报关单的货值；",
        "飞书落位 = 用「飞书多给」判断这笔费用实际被飞书贴在哪张报关单上。",
        "「涉及出运单数」= 装走这张入库单货物的出运单有几张；只有 1 张时不存在落位问题，",
        "已单独放到「单次出运（已排除）」表，不参与排查。",
        "「本单是货值最大 / 报关金额最大」两列只是把候选规则标出来，**不是结论**，",
        "请以费用行原文 + 出运完成情况自行判断。",
    ):
        notes.append([line])
    notes.column_dimensions["A"].width = 110

    OUT.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(OUT)
    except PermissionError:
        fallback = OUT.with_name(f"{OUT.stem}_new{OUT.suffix}")
        workbook.save(fallback)
        print(f"⚠ {OUT.name} 被 Excel/WPS 占用，已另存为 {fallback.name}；关闭后重跑即会写回。")
        print(f"written {fallback}（{len(multi)} + {len(single)} 行）")
        return
    print(f"written {OUT}（{len(multi)} + {len(single)} 行）")


if __name__ == "__main__":
    main()
