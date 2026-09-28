r"""采购金额匹配：拆单记录 × 入库单实发金额 × 飞书采购金额，三方对照落成四个文件。

用法：
    & '.\.venv\Scripts\python.exe' scripts\build_cost_match.py

输出（`outputs/cost_match/`）——一个文件回答一个问题：

    成本匹配_全部.xlsx     拆单记录逐条：我方金额、分摊依据、飞书金额、差异、判定
    成本匹配_正常.xlsx     与飞书一致（含「ADD 并单 / 合并行后一致」）
    成本匹配_异常.xlsx     **本轮新出现**的异常（已在待核实清单里的不再重复）；
                           另附「采购单对账」「缺报关单」「说明」三张表
    成本匹配_已核实.xlsx   人工已核实的冲突（`data/reference/verified_conflicts.json`）
    成本匹配_待核实.xlsx   待核实清单（**累计**）：所有未核对的异常，逐条写「错误原因 + 缺什么数据」

口径与分摊规则见 `app/services/cost_match.py` 模块开头。
本脚本只做三件事：取数（拆单记录 + 入库单 + 飞书）→ 判定 → 落表。
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

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from app.services.contract_shipments import cores_of  # noqa: E402
from app.services.cost_match import (  # noqa: E402
    IGNORABLE_DIFF,
    RecordCost,
    allocate_amounts,
    costs_for,
    dec,
    purchase_fees,
    record_batch_token,
    shipment_money,
    unit_core,
)
from app.services.erp_cache import CACHE_ROOT, read_jsonl  # noqa: E402
from app.services.grn_extract import PARSED_ROOT, safe_name  # noqa: E402
from app.services.grn_select import erp_amount, select  # noqa: E402
from app.services.supplier_names import SupplierNames  # noqa: E402

SPLIT = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
REF = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs_full.json"
VERIFIED = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
PARSE_SHEETS = (
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_出口退税联.xlsx",
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_预录单.xlsx",
)
OUT_DIR = PROJECT_ROOT / "outputs" / "cost_match"

# 额外费用（绳子费/管帽费/木箱费…）的分摊方案。
#
# `amount`（默认，**财务 2026-09-24 确认的口径**）：按金额分摊——
#   整张采购单的实发金额按各记录的**出运采购金额(RMB)** 占比摊到各条记录上，
#   费用自然跟着货值走。财务口径：**只要总成本不漏，费用摊到哪一票都可以**。
#
# 下面几个是排查时试过的**未验证假设**，只在复现历史贴法时用 `--fee-policy` 显式指定，
# 不作为正式口径：
#   max_goods / max_declared / max_goods_then_declared / last_shipment
FEE_POLICY = "amount"
UNVERIFIED_POLICIES = (
    "max_goods", "max_declared", "max_goods_then_declared", "last_shipment",
)

# ---- 判定口径（2026-09-24 用户确认）--------------------------------------------
# **最多允许 ±1 元的差额**，不用百分比容差。
# 旧脚本此处是「飞书 × 1%」的相对容差，会把 0.2%~0.9% 的真实差异当"一致"放过去
# （例：26MT-05Q196 少 14,720 被吞掉、26MT-03T203Y 的 981 分摊差被吞掉），已废除。
EQUAL = Decimal("1")    # 我方与飞书（或与睿贝）算「一致 / 相等」的最大差额：1 元
TOL = Decimal("0.01")   # 分位级严格相等，只用于「金额是否为零/是否有差额」这类判断
FEE_TOL = Decimal("1")  # 费用口径比对容差（元，与 EQUAL 同值但语义不同）
FAR = Decimal("0.10")   # **归因用**（不是判定口径）：「两边单据打架」的旁证阈值 >10%


def money(value: Decimal) -> str:
    """金额显示成两位小数（入库单里的原始数字常有 10 位小数）。"""
    return str(value.quantize(Decimal("0.01")))

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")


# --------------------------------------------------------------------------- 小工具


def read_rows(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    out = [dict(zip(header, row)) for row in rows[1:]]
    workbook.close()
    return out


def coarse_core(code: str) -> str:
    """再去掉 ADD 标记的核心号——飞书有时把附加单的金额并进主单那一行。"""
    return unit_core(code).split("-ADD")[0].split("ADD")[0]


def audit_core(code: str) -> str:
    """对账用的订单核心：`unit_core` 之后再剥掉尾部工厂后缀。

    `unit_core` 对标准单号（`26MT-02N182-JX`）已经剥了，但对 `26MT-DP002-XMLS`
    这类非标准单号不剥，而飞书写的是 `26MT-DP002`，不补一刀就永远配不上。
    """
    return re.sub(r"-[A-Z]{2,5}$", "", unit_core(code))


def feishu_core(record: dict) -> str:
    """飞书一行的采购单核心（飞书这一列叫「合同号（应收表格）」）。"""
    return unit_core(record.get("合同号（应收表格）") or record.get("合同号_1") or "")


def write_sheet(
    workbook: Workbook,
    title: str,
    columns: list[str],
    rows: list[dict],
    *,
    fill: PatternFill | None = None,
    first: bool = False,
) -> None:
    sheet = workbook.active if first else workbook.create_sheet()
    sheet.title = title
    sheet.append(columns)
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(name, "") for name in columns])
        if fill:
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    for index, name in enumerate(columns, 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            12, min(60, len(name) * 2 + 4)
        )
    sheet.freeze_panes = "A2"


def save(workbook: Workbook, path: Path) -> None:
    try:
        workbook.save(path)
    except PermissionError:
        # 文件正被 Excel 打开时另存一份，避免整轮跑白；但要大声提醒，
        # 否则目录里会多出一个 _new 文件、让人以为结果已经更新了。
        path = path.with_name(f"{path.stem}_new{path.suffix}")
        workbook.save(path)
        print(f"  ⚠ {path.name}：原文件被 Excel/WPS 占用，已另存为新文件；")
        print("    请关闭该工作簿后重跑本脚本，结果才会写回原文件。")
    print(f"  written {path.name}")


# --------------------------------------------------------------------------- 取数：我方金额


def shipment_dates() -> dict[str, str]:
    """出运单号（归一后）→ 出运日期。用睿贝出运单头，供「最后一批」策略使用。"""
    from app.services.erp_cache import CACHE_ROOT as _CACHE
    from app.services.shipment_index import canonical

    out: dict[str, str] = {}
    detail_dir = _CACHE / "details" / "shipments"
    if not detail_dir.exists():
        return out
    for path in sorted(detail_dir.glob("*.json")):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        base = payload.get("baseInfo") or {}
        date = str(base.get("出运日期") or "")
        for key in (payload.get("invoiceCode"), base.get("外销发票号"), path.stem):
            if key:
                out[canonical(str(key))] = date
    return out


def fee_target_for(
    code: str,
    indexes: list[int],
    split_rows: list[dict],
    policy: str,
    dates: dict[str, str],
) -> tuple[int | None, str]:
    """按策略挑出「无归属费用整笔归入」的那条记录，返回 (记录下标, 说明)。

    只在报关单全齐时调用。先在**报关单**层面挑目标（货值最大 / 出运日期最晚），
    再落到该报关单下货值最大的那条记录上（同一张报关单可能拆出多种产品类型）。
    """
    goods: dict[str, Decimal] = collections.defaultdict(Decimal)
    declared: dict[str, Decimal] = collections.defaultdict(Decimal)
    for index in indexes:
        decl = str(split_rows[index].get("报关单号") or "")
        goods[decl] += dec(split_rows[index].get("出运采购金额合计"))
        declared[decl] += dec(split_rows[index].get("报关金额"))
    if not goods:
        return None, ""
    target_decl = ""
    note = ""
    if policy == "max_declared":
        target_decl = max(declared, key=lambda decl: declared[decl])
        note = (
            "该采购单的报关单已全齐，费用整笔归入报关金额最大的那张报关单"
            f"（{target_decl}，报关金额 {money(declared[target_decl])}）"
        )
    elif policy == "max_goods_then_declared":
        top = max(goods.values())
        winners = [decl for decl in goods if goods[decl] == top]
        if len(winners) > 1:
            target_decl = max(winners, key=lambda decl: declared[decl])
            note = (
                "该采购单的报关单已全齐，货值并列（"
                f"{money(top)}），按报关金额兜底取（{target_decl}，报关金额 "
                f"{money(declared[target_decl])}）"
            )
        elif top > 0:
            target_decl = winners[0]
            note = f"该采购单的报关单已全齐，费用整笔归入货值最大的那张报关单（{target_decl}）"
        else:
            target_decl = max(declared, key=lambda decl: declared[decl])
            note = (
                "该采购单的报关单已全齐，但货值取不到，改按报关金额最大兜底"
                f"（{target_decl}，报关金额 {money(declared[target_decl])}）"
            )
    elif policy == "last_shipment":
        from app.services.shipment_index import canonical

        dated = {
            decl: dates.get(canonical(str(split_rows[index].get("合同号_1") or "")), "")
            for index in indexes
            for decl in [str(split_rows[index].get("报关单号") or "")]
        }
        known = {decl: date for decl, date in dated.items() if date}
        if not known:
            return None, ""
        target_decl = max(known, key=lambda decl: known[decl])
        note = f"该采购单的报关单已全齐，费用整笔归入最后一批（出运日期 {known[target_decl]}）"
    else:  # max_goods
        target_decl = max(goods, key=lambda decl: goods[decl])
        note = f"该采购单的报关单已全齐，费用整笔归入货值最大的那张报关单（{target_decl}）"
    candidates = [
        index for index in indexes
        if str(split_rows[index].get("报关单号") or "") == target_decl
    ]
    if not candidates:
        return None, ""
    # 同一张报关单下多条记录时，落到货值最大（取不到就报关金额最大）的那条
    target = max(
        candidates,
        key=lambda index: (
            dec(split_rows[index].get("出运采购金额合计")),
            dec(split_rows[index].get("报关金额")),
        ),
    )
    return target, note


def allocate_all(split_rows: list[dict], policy: str = FEE_POLICY) -> list[dict]:
    """给每条拆单记录算出「我方采购金额」。

    单位是采购单：同一采购单的记录一起送进 `allocate_amounts`，
    这样才看得清「这笔钱怎么分到各条记录」。
    """
    money = shipment_money()
    dates = shipment_dates() if policy == "last_shipment" else {}
    by_po: dict[str, list[int]] = collections.defaultdict(list)
    for index, row in enumerate(split_rows):
        by_po[str(row.get("采购单号") or "").strip()].append(index)
    totals = costs_for(list(by_po))

    out: list[dict] = [{} for _ in split_rows]
    for code, indexes in by_po.items():
        total, detail = totals[code]
        batch_amounts = {key: dec(value) for key, value in (detail["batch_amounts"] or {}).items()}
        # 无归属费用：入库单里没有批次归属的费用行（绳子费/管帽费/木箱费…）
        fee = fee_total(detail, code)
        fee_target: int | None = None
        fee_note = ""
        # 按金额分摊（amount/spread）时不需要挑目标票：整个实发金额按货值占比摊即可
        if fee > 0 and policy in UNVERIFIED_POLICIES:
            our_goods = sum(
                (dec(split_rows[index].get("出运采购金额合计")) for index in indexes),
                Decimal(0),
            )
            po_goods = money.rmb_by_po.get(code, Decimal(0))
            # 只在「报关单全齐」时归位：我方记录的人民币货值 = 睿贝整单出运人民币
            if po_goods and abs(our_goods - po_goods) <= EQUAL:
                fee_target, fee_note = fee_target_for(code, indexes, split_rows, policy, dates)
        # allocate_amounts 里用的是「本采购单内的位置」，这里把全局下标换过去
        local_target = indexes.index(fee_target) if fee_target is not None else None
        costs = [
            RecordCost(
                amount_rmb=dec(split_rows[index].get("出运采购金额合计")),
                amount_usd=dec(split_rows[index].get("出运金额合计")),
                batch=record_batch_token(split_rows[index].get("合同号_1"), code),
                product_type=str(split_rows[index].get("产品类型") or "").strip(),
                purchase_code=code,
            )
            for index in indexes
        ]
        amounts = allocate_amounts(
            costs,
            total,
            batch_amounts,
            money.rmb_by_kind,
            money.usd_by_po.get(code),
            money.rmb_by_po.get(code),
            unattributed_fee=fee if fee_target is not None else Decimal(0),
            fee_target=local_target,
            fee_note=fee_note,
        )
        for index, (amount, basis) in zip(indexes, amounts):
            unpriced = str(basis).startswith("睿贝该行未定价")
            out[index] = {
                "amount": Decimal(0) if unpriced else amount,
                "basis": basis,
                "blank": unpriced,
                "total": total,
                "detail": detail,
                "code": code,
            }
    return out


def fee_total(po_detail: dict, code: str) -> Decimal:
    """该采购单的「无归属费用」合计（元）。

    口径 = **实发合计 − 产品明细金额之和**：入库单里的绳子费/管帽费/木箱费/木轴费这类
    独立费用行不属于任何产品行，飞书（=工厂发票口径）常把它们整笔贴在某一票报关单上。
    金额在 `costs_for()` 里算好挂在明细的 `unattributed_fee` 上（那里才拿得到附件与解析）。

    为什么不用解析结果的 `extra_fees`：它可能含**没进实发合计**的补款/批注行，会把费用
    算大（实测 26MT-02N182-HD 算成 71,800，实际实发−明细只有 57,895）。
    """
    del code  # 保留参数只为调用处可读；金额已随明细返回
    return dec(po_detail.get("unattributed_fee") or 0)


# --------------------------------------------------------------------------- 判定


def po_note(
    *,
    mine: Decimal,
    feishu_local: Decimal,
    feishu_all: Decimal,
    feishu_rows: int,
    invoice: Decimal,
    erp: Decimal,
    grn_total: Decimal,
    fee: Decimal,
) -> str:
    """采购单级结论：这张采购单的钱对不对得上、对不对得上哪一侧。"""
    if not feishu_local and not feishu_all:
        return "飞书无对应行"
    diff = mine - feishu_local
    if abs(diff) <= EQUAL:
        return "整单一致"
    if feishu_all and abs(feishu_all - mine) <= EQUAL:
        return f"整单一致（飞书 {feishu_rows} 行合计 = 我方）"
    if erp and abs(erp - mine) <= EQUAL:
        return "我方 = 睿贝采购金额"
    if erp and abs(erp - feishu_local) <= EQUAL:
        return "飞书 = 睿贝采购金额（旧价）"
    if grn_total and mine < grn_total - EQUAL:
        return f"部分覆盖：我方只拿到该采购单的部分报关单，未分配 {money(grn_total - mine)}"
    if feishu_all and feishu_all > feishu_local + EQUAL:
        return (
            f"飞书另有 {money(feishu_all - feishu_local)} 落在我们没有的报关单上"
            f"（飞书全 {feishu_rows} 行 {money(feishu_all)}）"
        )
    if fee and abs(abs(diff) - abs(fee)) <= FEE_TOL:
        return f"费用口径：我方含入库单费用行 {fee}"
    if invoice and abs(grn_total - invoice) > EQUAL:
        return (
            f"发票 vs 入库单：工厂开票 {invoice}，入库单实发 {grn_total}"
            f"（差 {money(grn_total - invoice)}）"
        )
    return (
        f"整单不一致：我方=入库单实发 {money(grn_total)}，飞书各行合计 {money(feishu_local)}"
        f"（差 {money(diff)}），两边都不等于睿贝 {erp or '—'}"
    )


def po_reconciled(entry: dict) -> bool:
    """这张采购单的钱在飞书那边是不是齐的。

    判据：飞书该采购单**全部行**合计 = 入库单实发（差额 <1 个货币单位，2026-09-28
    用户口径：这种补充差额忽略不计）。这种单子即使我方没拆出全部报关单，差额也是
    「这笔钱摊到哪一票」的分摊口径问题，不是少钱。
    """
    grn = dec(entry.get("入库单实发总额"))
    feishu_all = dec(entry.get("飞书合计(含ADD)"))
    return bool(grn and feishu_all) and abs(grn - feishu_all) <= IGNORABLE_DIFF


def unit_fee_text(entry: dict) -> str:
    """这张采购单单据里能核到的费用（入库单费用行 + 采购订单「费用信息」）。

    2026-09-28 财务确认：睿贝「出运采购金额」会漏费用；飞书按出运采购金额给钱时，
    用这段文字说明"少掉的是哪几笔"，写进异常说明和「待核实」表。
    """
    parts: list[str] = []
    rows = str(entry.get("费用行明细") or "").strip()
    if rows:
        parts.append(f"入库单费用行：{rows}")
    po_fee = dec(entry.get("采购订单费用信息") or 0)
    if po_fee:
        parts.append(f"采购订单「费用信息」：{po_fee}")
    return "；".join(parts)


def fee_groups(entry: dict) -> tuple[list[tuple[str, Decimal]], list[tuple[str, Decimal]]]:
    """该采购单可以拿来「对号入座」的费用清单，分两组：

    1. 明细：入库单里解析出来的费用行 + 采购订单「费用信息」；
    2. 兜底：入库单「无归属费用」合计（费用行没解析全时，用整块费用合计去对）。

    先对明细、再对兜底——这样能命中"差额 = 包装费 6000"这种更具体的话。
    """
    out: list[tuple[str, Decimal]] = []
    for item in str(entry.get("费用行明细") or "").split("、"):
        text = item.strip()
        if not text:
            continue
        label, _sep, amount = text.rpartition(" ")
        try:
            value = Decimal(amount.replace(",", ""))
        except Exception:  # noqa: BLE001
            continue
        if value:
            out.append((label or "费用行", value.quantize(Decimal("0.01"))))
    po_fee = dec(entry.get("采购订单费用信息") or 0)
    if po_fee:
        out.append(("采购订单「费用信息」", po_fee.quantize(Decimal("0.01"))))
    bulk: list[tuple[str, Decimal]] = []
    grn_fee = dec(entry.get("无归属费用") or 0)
    if grn_fee > Decimal("1") and not out:
        bulk.append(("入库单费用合计（无归属费用）", grn_fee.quantize(Decimal("0.01"))))
    return out, bulk


def fee_explain(diff: Decimal, items: list[tuple[str, Decimal]]) -> str:
    """差额能不能被**一笔或几笔费用正好凑出来**？凑得出就返回说明，凑不出返回空串。

    2026-09-28 用户口径：只有"确定差异 = 某个/某几个额外费用（木箱费、包装费…）"的单子
    才丢进待核实表；其余照旧报异常。
    """
    if diff <= Decimal("1") or not items:
        return ""
    limit = Decimal("0.05")          # 分位舍入容差
    for label, amount in items:
        if abs(diff - amount) <= limit:
            return f"{label} {amount}"
    usable = [item for item in items if 0 < item[1] <= diff + limit]
    if not usable or len(usable) > 16:
        return ""
    # 费用行一般不超过十几笔，直接枚举组合（2^n）
    for mask in range(1, 1 << len(usable)):
        total = Decimal(0)
        chosen: list[str] = []
        for index, (label, amount) in enumerate(usable):
            if mask >> index & 1:
                total += amount
                chosen.append(f"{label} {amount}")
                if total > diff + limit:
                    break
        if abs(total - diff) <= limit:
            return " + ".join(chosen) if len(chosen) > 1 else chosen[0]
    return ""


# ----------------------------------------------------------------- 出运数量核验


def shipment_quantities() -> dict[str, tuple[Decimal, str]]:
    """采购单 → (出运数量合计, 出运单号列表)，给「出运数量不能比入库数量多」用。

    三条口径（都来自实测踩过的坑）：
    1. 只统计**计数单位**（EA/PC/PCS/个/只/支/件…）的行；长度、重量单位的单子不参与——单位和
       入库单对不齐（实测 `26MT-03P262Y` 睿贝写 EA、实际是米），会造出一堆假异常；
    2. 按「SKU + 数量 + 美元金额」去重：同一票货常被 PI 号和供应商自家 S 号各录一遍
       （实测 `26MT-02N105`、`26MT-05X167`、`26MT-03T094Y-HD`），不去重就凭空多一倍；
    3. 认不出单位的行不算，避免把脏数据算进来。
    """
    from app.services.grn_extract import unit_family
    from app.services.shipment_index import load_lines

    rows: dict[str, list[dict]] = collections.defaultdict(list)
    for line in load_lines():
        code = str(line.get("purchase_code") or "").strip()
        if code:
            rows[code].append(line)
    out: dict[str, tuple[Decimal, str]] = {}
    for code, items in rows.items():
        unique: dict[tuple, dict] = {}
        for line in items:
            key = (
                str(line.get("sku") or "").strip(),
                str(line.get("quantity") or "").strip(),
                str(line.get("amount_usd") or "").strip(),
            )
            unique.setdefault(key, line)
        kept = list(unique.values())
        if not kept or any(unit_family(line.get("unit")) != "count" for line in kept):
            continue
        total = sum((dec(line.get("quantity")) for line in kept), Decimal(0))
        if total <= 0:
            continue
        invoices = "、".join(
            sorted({str(line.get("invoice_code") or "").strip() for line in kept})
        )
        out[code] = (total, invoices)
    return out


def quantity_check(po_rows: dict[str, dict]) -> dict[str, dict]:
    """出运数量核验：**出运数量不能大于入库数量**（可以有剩余，不能凭空冒出来）。

    规则（2026-09-28 用户口径）：
    - 两边都按单位族对齐：入库单的数量从明细列按 count/length/weight 归族求和；
    - 只要**任一单位族**的入库数量与出运数量相等（±1），就认为只是单位口径不同 → 不算异常；
    - 入库单根本认不出数量（0）→ 不判定，避免误报；
    - 出运数量 > 计数单位入库数量 → 判异常（整组报错），说明里写出用到的出运单与入库单。
    """
    erp = shipment_quantities()
    out: dict[str, dict] = {}
    for code, entry in po_rows.items():
        pair = erp.get(code)
        if not pair:
            continue
        erp_qty, invoices = pair
        families = {
            family: dec(value)
            for family, value in (entry.get("数量(按单位族)") or {}).items()
        }
        count = families.get("count", Decimal(0))
        if count <= 0:
            continue
        if any(abs(value - erp_qty) <= IGNORABLE_DIFF for value in families.values()):
            continue
        block_qty = dec(entry.get("入库块数量") or 0)
        if block_qty > 0 and erp_qty <= block_qty + IGNORABLE_DIFF:
            continue      # 单据自己的「实发数量」够 → 只是明细列登记不全
        if erp_qty > count + IGNORABLE_DIFF:
            out[code] = {
                "出运数量": erp_qty,
                "入库数量": count,
                "出运单": invoices,
                "入库单": str(entry.get("入库单") or ""),
            }
    return out


def unit_reason(
    *,
    mine: Decimal,
    theirs: Decimal,
    erp_unit: Decimal,
    fee: Decimal,
    note: str,
    blank: bool,
    sibling: Decimal,
    sibling_rows: int,
    po_fee_info: Decimal = Decimal(0),
    invoice: Decimal = Decimal(0),
    po_reconciled: bool = False,
    fee_items_text: str = "",
    fee_ceiling: Decimal = Decimal(0),
) -> tuple[str, str]:
    """异常行的「原因分类 + 归因说明」。

    顺序 = 证据强度：先看有没有硬事实（未定价 / 费用行 / ADD 并单 / 采购单级结论），
    剩下的再看差额到底来自「这笔钱怎么分到各报关单」（分摊口径），
    还是「两边单据本身打架」（单据冲突：与飞书、睿贝都差 10% 以上）。
    最后统一附上「我方/飞书 是不是等于睿贝出运金额」的旁证。
    """
    if blank:
        category, text = "睿贝未定价", "睿贝该行未定价（采购金额留空，待人工补价）"
    elif po_fee_info and abs(abs(mine - theirs) - abs(po_fee_info)) <= FEE_TOL:
        # 差额正好等于采购订单界面的「费用信息」：这笔费用谁算谁不算，就是这一条的分歧点
        gap = money(abs(mine - theirs))
        fee_text = f"差额 {gap} 正好等于该采购单「费用信息」{money(po_fee_info)}"
        if mine < theirs:
            category = "采购订单费用信息未进入库单"
            text = f"{fee_text}：这笔费用只在采购订单界面记账，入库单实发里没有"
            if invoice and abs(invoice - theirs) <= EQUAL:
                text += f"；该采购单工厂开票 {money(invoice)}，与飞书一致"
        else:
            category = "飞书未含采购订单费用信息"
            text = f"{fee_text}：入库单实发把这笔费用算进去了，飞书没有"
    elif po_reconciled and note.startswith(("部分覆盖", "飞书另有")):
        # 这块钱在飞书那边是齐的（飞书该采购单全部行 = 入库单实发，差额 <1 个货币单位），
        # 只是我方没把全部报关单拆出来 → 差额属于"这笔钱摊到哪一票"，不是缺单。
        category = "分摊口径"
        text = (
            f"{note}；但该采购单的钱在飞书那边是齐的（飞书全部行合计 = 入库单实发，"
            "差额 <1 个货币单位），按分摊口径处理"
        )
    elif (
        erp_unit
        and abs(theirs - erp_unit) <= EQUAL          # 飞书给的就是睿贝出运采购金额
        and mine - theirs > EQUAL                    # 但比我们（入库单实发口径）少
        and fee_items_text                           # 单据里能明确看到费用
        and mine - theirs <= fee_ceiling + EQUAL     # 少掉的那部分由费用就能解释
    ):
        # 2026-09-28 财务确认：睿贝「出运采购金额」本身就会漏掉某些费用。
        # 飞书直接按出运采购金额给钱、而单据里明明有费用 → 一律报异常（不按分摊口径放行）。
        category = "飞书按出运采购金额计（漏费用）"
        text = (
            f"飞书 {money(theirs)} = 睿贝出运采购金额 {money(erp_unit)}（出运采购金额会漏费用，"
            f"已与财务确认）；我方按入库单实发 {money(mine)}，比飞书多 {money(mine - theirs)}；"
            f"单据里可核到的费用：{fee_items_text}"
        )
    elif fee and abs(abs(mine - theirs) - abs(fee)) <= FEE_TOL:
        category = "单据口径（实发含费用行）"
        text = f"我方含入库单费用行 {fee}，飞书只记货值"
    elif sibling_rows > 1 and sibling and abs(sibling - theirs) <= EQUAL:
        category = "分摊口径"
        text = f"ADD 并单：飞书一行 = 我方主单 + {sibling_rows - 1} 张附加单合计 {money(sibling)}"
    elif note.startswith("整单一致"):
        category, text = "分摊口径", f"拆行差异：{note}"
    elif note.startswith(("部分覆盖", "飞书另有")):
        category = "缺报关单（待业务补）"
        text = f"缺报关单：{note}"
    elif note.startswith(("我方 = 睿贝", "飞书 = 睿贝")):
        category = "单价口径"
        text = f"{note}（采购单级：我方 {money(mine)} / 飞书 {money(theirs)}）"
    elif note.startswith("费用口径"):
        category, text = "单据口径（实发含费用行）", note
    elif note.startswith("发票 vs 入库单"):
        category, text = "发票与入库单不一致", note
    elif (
        erp_unit
        and abs(mine - erp_unit) / erp_unit > FAR
        and abs(mine - theirs) / theirs > FAR
    ):
        category = "单据冲突"
        text = (
            f"我方 {money(mine)} 与飞书 {money(theirs)}、睿贝出运 {money(erp_unit)} "
            "都差 10% 以上，两侧单据本身对不上，待业务核"
        )
    else:
        category = "分摊口径"
        text = (
            f"差额来自这笔钱怎么分到各报关单/各行（我方 {money(mine)} / "
            f"飞书 {money(theirs)} / 睿贝出运 {money(erp_unit) if erp_unit else '—'}）"
        )
    tags: list[str] = []
    if erp_unit and abs(mine - erp_unit) <= EQUAL:
        tags.append(f"我方 = 睿贝出运采购金额 {money(erp_unit)}")
    if erp_unit and abs(theirs - erp_unit) <= EQUAL:
        tags.append(f"飞书 = 睿贝出运采购金额 {money(erp_unit)}（旧单价）")
    return category, "；".join([text] + tags)


# ----------------------------------------------------------------- 待核实表

# 「缺什么数据」按原因分类给一句可执行的话（2026-09-28 用户要求：每条注明错误原因 + 缺什么数据）
MISSING_DATA_HINT = {
    "飞书按出运采购金额计（漏费用）": (
        "需要财务确认：单据里这笔费用（见「单据里能核到的费用」）该不该计入本票成本"
        "——睿贝出运采购金额本身会漏费用（2026-09-28 财务确认）"
    ),
    "缺报关单（待业务补）": (
        "缺该采购单还没解析到的报关单（见异常工作簿的「缺报关单」表，可按入库单余额印证）"
    ),
    "采购订单费用信息未进入库单": "入库单里缺这笔费用行；需要入库单原件或补单",
    "飞书未含采购订单费用信息": "飞书这一票没算采购订单「费用信息」；需要财务确认口径",
    "单据口径（实发含费用行）": "飞书那侧未含入库单的费用行；需要确认是按实发还是按货值",
    "单价口径": "睿贝出运单价 / 飞书单价不一致；需要确认以哪个单价为准",
    "发票与入库单不一致": "工厂开票金额与入库单实发不一致；需要发票或差异说明",
    "飞书无对应行": "飞书里没有这一单元的对应行；需要飞书补行",
    "睿贝未定价": "睿贝该行没有采购单价；需要补价",
    "出运数量 > 入库数量": (
        "需要业务/仓库确认：多出来的这批货到底怎么回事（漏货、退换货，还是出运单多填）"
        "——入库单里的实发数量是够的凭证；出运单上这批货的美元金额还像是随手填的（如 0.01）"
    ),
    "分摊口径": "分摊口径待确认（整单金额已对得上）",
}


def verification_rows(
    bad: list[dict], po_rows: dict[str, dict]
) -> list[dict]:
    """把「未核对异常」整理成待核实清单：一条 = 一行记录。

    - 待核实类型 = 异常的原因分类（含 2026-09-28 新增的「飞书按出运采购金额计（漏费用）」）；
    - 错误原因 = 说明（谁跟谁差多少、差在哪）；
    - 缺什么数据 = 要拿到什么才能定案；
    - 「飞书=出运采购金额」「单据里能核到的费用」两列，方便直接筛出"财务按出运金额给钱、漏了费用"的单子。
    """
    out: list[dict] = []
    for row in bad:
        code = str(row.get("采购单号") or "")
        po = po_rows.get(code) or {}
        reason = str(row.get("异常原因") or "")
        mine = dec(row.get("采购金额_我方"))
        theirs = dec(row.get("采购金额_飞书"))
        erp = dec(row.get("ERP出运采购金额(本单元)"))
        out.append(
            {
                "待核实类型": reason,
                "待核实事由": str(row.get("标注") or ""),
                "缺什么数据": MISSING_DATA_HINT.get(reason, "待人工判断"),
                "差额对应的费用": str(row.get("_fee_explain") or ""),
                "错误原因": str(row.get("归因说明") or ""),
                "飞书=出运采购金额": (
                    "是" if (erp and theirs and abs(theirs - erp) <= EQUAL) else "否"
                ),
                "单据里能核到的费用": unit_fee_text(po) or "",
                "采购单号": code,
                "报关单号": row.get("报关单号"),
                "合同号_1": row.get("合同号_1"),
                "供应商简称": row.get("供应商简称"),
                "产品类型": row.get("产品类型"),
                "报关品名": row.get("报关品名"),
                "采购金额_我方": row.get("采购金额_我方"),
                "采购金额_飞书": row.get("采购金额_飞书"),
                "出运采购金额(本单元)": row.get("ERP出运采购金额(本单元)"),
                "差异": row.get("差异"),
                "比对口径": row.get("比对口径"),
                "采购单实发总额": row.get("采购单实发总额"),
                "工厂开票金额": po.get("工厂开票金额", ""),
                "睿贝采购金额": po.get("睿贝采购金额", ""),
                "分摊依据": row.get("分摊依据"),
            }
        )
    return out


VERIFICATION_COLUMNS = [
    "待核实类型", "待核实事由", "缺什么数据", "差额对应的费用", "错误原因",
    "飞书=出运采购金额", "单据里能核到的费用",
    "采购单号", "报关单号", "合同号_1", "供应商简称", "产品类型", "报关品名",
    "采购金额_我方", "采购金额_飞书", "出运采购金额(本单元)", "差异", "比对口径",
    "采购单实发总额", "工厂开票金额", "睿贝采购金额", "分摊依据",
]


VERIFICATION_PATH = OUT_DIR / "成本匹配_待核实.xlsx"


def verification_key(row: dict) -> tuple:
    """待核实清单里「同一条」的判据：报关单 + 采购单 + 供应商 + 我方金额。

    这四项一起才能区分同一张报关单下同一采购单拆出的多行记录；
    金额变了说明这条记录本身变了，按新的一条处理。
    """
    return (
        str(row.get("报关单号") or ""),
        str(row.get("采购单号") or ""),
        str(row.get("供应商简称") or ""),
        str(row.get("采购金额_我方") or ""),
    )


def verification_unit_key(row: dict) -> tuple:
    """「同一条记录」的粗判据（不含金额）：用来判断这条还在不在异常里。"""
    return (
        str(row.get("报关单号") or ""),
        str(row.get("采购单号") or ""),
        str(row.get("供应商简称") or ""),
    )


def load_verification(path: Path) -> tuple[list[str], list[dict]]:
    """读回上一轮的待核实清单（累计用）。

    保留原有的列（用户自己加的备注列也不会丢），只按表头取回来。
    """
    if not path.exists():
        return [], []
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "待核实" not in workbook.sheetnames:
            return [], []
        rows = list(workbook["待核实"].iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        return [], []
    header = [str(cell) if cell is not None else "" for cell in rows[0]]
    out: list[dict] = []
    for values in rows[1:]:
        if not any(value not in (None, "") for value in values):
            continue
        out.append({name: values[index] for index, name in enumerate(header) if name})
    return header, out


# --------------------------------------------------------------------------- 缺报关单


def grn_uncovered_index(
    ours: dict[tuple[str, str], Decimal],
    names: SupplierNames,
    wanted: set[tuple[str, str]],
) -> dict[tuple[str, str], tuple[Decimal, Decimal]]:
    """（订单核心 × 供应商）→ (入库单实发总额, 还没分配出去的余额)。

    缺报关单的钱能不能采信，就看「入库单余额」与「飞书缺单行的金额」对不对得上：
    对得上 → 财务填的钱正是入库单剩下的那块（可采信）。
    只算 `wanted` 里的组合——不相关的采购单不必去读它们的入库单。
    """
    purchases = {
        str(row.get("purchase_code") or "").replace(" ", "").replace("_", "").upper(): str(
            row.get("supplierName") or ""
        )
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
    }
    po_totals: dict[tuple[str, str], Decimal] = collections.defaultdict(Decimal)
    for folder in sorted(path for path in PARSED_ROOT.iterdir() if path.is_dir()):
        code = folder.name
        supplier = names.short(
            purchases.get(code.replace(" ", "").replace("_", "").upper(), "")
        )
        key = (audit_core(code), supplier)
        if key not in wanted:
            continue
        total = sum((dec(item.get("amount")) for item in select(code)["kept"]), Decimal(0))
        if not total:
            continue
        po_totals[key] += total
    return {
        key: (total, total - ours.get(key, Decimal(0))) for key, total in po_totals.items()
    }


def missing_declaration_sheet(
    rows: list[dict],
    split_rows: list[dict],
    feishu_by_core: dict[tuple[str, str], list[dict]],
    affected_cores: set[str],
) -> tuple[list[str], list[dict]]:
    """「飞书里有钱、我们没有报关单 PDF」的报关单清单。

    这些单子补上来之后，「缺报关单」这类差异会自然消掉。优先级按影响程度排；
    每行还带一句「入库单余额」判定，说明这块钱有没有凭证支撑。
    """
    ours: set[str] = set()
    for path in (*PARSE_SHEETS, SPLIT):
        if path.exists():
            ours |= {str(row.get("报关单号") or "").strip() for row in read_rows(path)}
    our_cores: set[str] = set()
    for row in split_rows:
        our_cores |= cores_of(str(row.get("合同号_1") or ""))

    mine: dict[tuple[str, str], Decimal] = collections.defaultdict(Decimal)
    matched: dict[tuple[str, str], set[str]] = collections.defaultdict(set)
    for row in rows:
        key = (audit_core(str(row["采购单号"])), str(row["供应商简称"] or ""))
        mine[key] += row["_mine"]
        matched[key].add(str(row["报关单号"]))

    grouped: dict[str, dict] = {}
    for record in json.loads(REF.read_text(encoding="utf-8"))["records"]:
        decl = str(record.get("报关单号") or "").strip()
        if not decl or decl in ours:
            continue
        contract = str(record.get("合同号（应收表格）") or record.get("合同号_1") or "")
        entry = grouped.setdefault(
            decl,
            {
                "合同号": set(),
                "供应商": set(),
                "品名": set(),
                "报关金额": Decimal(0),
                "采购金额": Decimal(0),
                "重量": Decimal(0),
                "核心": set(),
            },
        )
        entry["合同号"].add(contract)
        entry["供应商"].add(str(record.get("供应商简称") or "").strip())
        entry["品名"].add(str(record.get("报关品名") or ""))
        entry["报关金额"] += dec(record.get("报关金额"))
        entry["采购金额"] += dec(record.get("采购金额"))
        entry["重量"] += dec(record.get("报关重量"))
        entry["核心"] |= cores_of(contract)

    # 只有这些（订单核心 × 供应商）需要查入库单余额
    wanted = {
        (audit_core(next(iter(entry["合同号"]), "")), next(iter(sorted(entry["供应商"])), ""))
        for entry in grouped.values()
    }
    balances = grn_uncovered_index(mine, SupplierNames(), wanted)

    out: list[dict] = []
    for decl, entry in grouped.items():
        core = audit_core(next(iter(entry["合同号"]), ""))
        vendor = next(iter(sorted(entry["供应商"])), "")
        key = (core, vendor)
        bucket = feishu_by_core.get(key) or []
        total_money = sum((dec(r.get("采购金额")) for r in bucket), Decimal(0))
        matched_money = sum(
            (
                dec(r.get("采购金额"))
                for r in bucket
                if str(r.get("报关单号") or "") in matched.get(key, set())
            ),
            Decimal(0),
        )
        missing_money = total_money - matched_money
        po_total, balance = balances.get(key, (Decimal(0), Decimal(0)))
        if missing_money <= EQUAL:
            trust = "—"
        elif not po_total:
            trust = "无入库单"
        elif balance <= EQUAL:
            trust = "余额已用尽（飞书这行重复计或多算，需查）"
        elif abs(balance - missing_money) <= EQUAL:
            trust = "可采信"
        elif missing_money > balance:
            trust = "飞书多（需查）"
        else:
            trust = "飞书少（未出完）"
        hit = sorted(entry["核心"] & our_cores)
        cores = {audit_core(core) for core in entry["核心"]}
        contract_text = " ".join(sorted(entry["合同号"])).upper()
        if cores & affected_cores:
            level = "最高：涉及已知被缺单影响的采购单"
        elif hit:
            level = "高：涉及我方在做的采购单"
        elif any(token in contract_text for token in ("SP-", "SP", "CY-", "ZY", "24MT")):
            level = "低：历史遗留/样品/代理（按规则拦截）"
        else:
            level = "中：涉及的订单我方完全没有"
        out.append(
            {
                "优先级": level,
                "报关单号": decl,
                "合同号": "、".join(sorted(c for c in entry["合同号"] if c)),
                "飞书订单核心": core,
                "供应商简称": "、".join(sorted(s for s in entry["供应商"] if s)),
                "报关品名": "、".join(sorted(n for n in entry["品名"] if n)),
                "报关金额(USD)": str(entry["报关金额"].normalize()) if entry["报关金额"] else "",
                "飞书采购金额": str(entry["采购金额"].normalize()) if entry["采购金额"] else "",
                "报关重量": str(entry["重量"].normalize()) if entry["重量"] else "",
                "涉及的采购单核心": "、".join(hit),
                "该组合缺单行金额": str(missing_money),
                "入库单余额": str(balance) if balance else "",
                "余额判定": trust,
            }
        )
    order = {
        "最高：涉及已知被缺单影响的采购单": 0,
        "高：涉及我方在做的采购单": 1,
        "中：涉及的订单我方完全没有": 2,
        "低：历史遗留/样品/代理（按规则拦截）": 3,
    }
    out.sort(key=lambda row: (order.get(str(row["优先级"]), 9), str(row["报关单号"])))
    return (
        [
            "优先级", "报关单号", "合同号", "飞书订单核心", "供应商简称", "报关品名", "报关金额(USD)",
            "飞书采购金额", "报关重量", "涉及的采购单核心", "该组合缺单行金额",
            "入库单余额", "余额判定",
        ],
        out,
    )


# --------------------------------------------------------------------------- 主流程


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="采购金额匹配：生成 outputs/cost_match 四个文件")
    parser.add_argument(
        "--fee-policy",
        choices=("amount", "spread", *UNVERIFIED_POLICIES),
        default=FEE_POLICY,
        help="额外费用分摊方案：amount=按金额分摊（默认，财务口径）；"
             "其余为未验证假设，仅供排查",
    )
    args = parser.parse_args()
    print(f"额外费用分摊方案：{args.fee_policy}"
          + ("（未验证假设！）" if args.fee_policy in UNVERIFIED_POLICIES else ""))

    split_rows = read_rows(SPLIT)
    allocated = allocate_all(split_rows, args.fee_policy)
    print(f"拆单记录 {len(split_rows)}")

    ref_records = json.loads(REF.read_text(encoding="utf-8"))["records"]
    # 飞书按「报关单号 + 供应商简称 + 采购单核心」索引，就是我们的比对单元
    feishu: dict[tuple[str, str, str], list[dict]] = collections.defaultdict(list)
    for record in ref_records:
        decl = str(record.get("报关单号") or "").strip()
        if decl:
            feishu[
                (decl, str(record.get("供应商简称") or "").strip(), feishu_core(record))
            ].append(record)
    # 采购单级对账用（按订单核心 × 供应商，含 ADD 与开票金额）
    feishu_by_core: dict[tuple[str, str], list[dict]] = collections.defaultdict(list)
    for record in ref_records:
        core = audit_core(record.get("合同号（应收表格）") or record.get("合同号_1") or "")
        if core:
            feishu_by_core[(core, str(record.get("供应商简称") or "").strip())].append(record)

    # ---- 记录级行
    rows: list[dict] = []
    for index, row in enumerate(split_rows):
        info = allocated[index]
        code = info["code"]
        decl = str(row.get("报关单号") or "").strip()
        vendor = str(row.get("供应商简称") or "").strip()
        core = unit_core(code)
        f_rows = feishu.get((decl, vendor, core)) or []
        theirs = sum((dec(r.get("采购金额")) for r in f_rows), Decimal(0)) if f_rows else None
        rows.append(
            {
                "报关单号": decl,
                "合同号_1": row.get("合同号_1"),
                "采购单号": code,
                "采购单核心": core,
                "供应商简称": vendor,
                "产品类型": row.get("产品类型"),
                "报关品名": row.get("报关品名"),
                "出运金额合计(USD)": row.get("出运金额合计"),
                "采购金额_我方": "" if info["blank"] else str(info["amount"]),
                "分摊依据": info["basis"],
                "采购单实发总额": str(info["total"]),
                "采购金额_飞书": "" if theirs is None else str(theirs),
                "飞书行数": len(f_rows),
                "ERP出运采购金额(本单元)": "",
                "_unit": (decl, vendor, core),
                "_mine": info["amount"],
                "_theirs": theirs,
                "_blank": info["blank"],
                "_index": index,
                "_rmb": dec(row.get("出运采购金额合计")),
            }
        )

    # ---- 采购单级对账
    by_po: dict[str, list[dict]] = collections.defaultdict(list)
    for row in rows:
        by_po[str(row["采购单号"])].append(row)
    po_rows: dict[str, dict] = {}
    po_fee: dict[str, Decimal] = {}
    for code, group in by_po.items():
        core = audit_core(code)
        vendor = str(group[0]["供应商简称"] or "")
        bucket = feishu_by_core.get((core, vendor)) or []
        feishu_all = sum((dec(r.get("采购金额")) for r in bucket), Decimal(0))
        invoice = sum((dec(r.get("开票金额")) for r in bucket), Decimal(0))
        seen: set[tuple[str, str]] = set()
        feishu_local = Decimal(0)
        for row in group:
            key = (str(row["报关单号"]), str(row["采购金额_飞书"]))
            if not row["采购金额_飞书"] or key in seen:
                continue
            seen.add(key)
            feishu_local += dec(row["采购金额_飞书"])
        mine = sum((row["_mine"] for row in group), Decimal(0))
        grn_total = dec(group[0]["采购单实发总额"])
        po_fee[code] = fee_total(allocated[group[0]["_index"]]["detail"], code)
        # 入库单里解析出来的费用行明细（费用名 + 金额），只用于「待核实」表留痕
        fee_rows = [
            item
            for item in (allocated[group[0]["_index"]]["detail"].get("fee_items") or [])
            if dec(item.get("amount"))
        ]
        po_rows[code] = {
            "采购单号": code,
            "供应商": vendor,
            "入库单实发总额": str(grn_total),
            "我方合计": str(mine),
            "飞书合计(本单各行)": str(feishu_local),
            "飞书合计(含ADD)": str(feishu_all) if feishu_all else "",
            "飞书行数": len(bucket),
            "工厂开票金额": str(invoice) if invoice else "",
            "睿贝采购金额": str(erp_amount(code) or ""),
            "采购订单费用信息": str(purchase_fees().get(code, "") or ""),
            "费用行明细": "、".join(
                f"{item.get('label') or '(未命名)'} {dec(item.get('amount'))}"
                for item in fee_rows
            ),
            "费用行合计": str(sum((dec(item.get("amount")) for item in fee_rows), Decimal(0))),
            "无归属费用": str(po_fee[code]),
            "数量(按单位族)": dict(
                allocated[group[0]["_index"]]["detail"].get("qty_by_family") or {}
            ),
            "入库块数量": str(
                allocated[group[0]["_index"]]["detail"].get("block_qty") or ""
            ),
            "入库单": "、".join(
                allocated[group[0]["_index"]]["detail"].get("kept_names") or []
            ),
            "我方−飞书": str(mine - feishu_local),
            "结论": po_note(
                mine=mine,
                feishu_local=feishu_local,
                feishu_all=feishu_all,
                feishu_rows=len(bucket),
                invoice=invoice,
                erp=erp_amount(code) or Decimal(0),
                grn_total=grn_total,
                fee=po_fee[code],
            ),
        }

    # ---- 单元级判定（飞书是按「报关单 + 供应商 + 采购单」给钱的，我方同粒度比对）
    # 出运数量核验（2026-09-28）：出运数量 > 入库数量 = 凭空冒出来的货 → 整组报错
    quantity_flags = quantity_check(po_rows)
    if quantity_flags:
        print(
            "出运数量核验：出运 > 入库 的采购单 "
            + str(len(quantity_flags))
            + " 个 —— "
            + "、".join(
                f"{code}(出运 {flag['出运数量']} / 入库 {flag['入库数量']})"
                for code, flag in sorted(quantity_flags.items())
            )
        )
    units: dict[tuple[str, str, str], list[dict]] = collections.defaultdict(list)
    for row in rows:
        units[row["_unit"]].append(row)
    coarse_ours: dict[tuple[str, str, str], Decimal] = collections.defaultdict(Decimal)
    for row in rows:
        coarse_ours[
            (row["_unit"][0], row["_unit"][1], coarse_core(row["采购单核心"]))
        ] += row["_mine"]
    feishu_coarse: dict[tuple[str, str, str], Decimal] = collections.defaultdict(Decimal)
    for (decl, vendor, core), f_rows in feishu.items():
        feishu_coarse[(decl, vendor, coarse_core(core))] += sum(
            (dec(r.get("采购金额")) for r in f_rows), Decimal(0)
        )

    verified_cases = {}
    if VERIFIED.exists():
        verified_cases = {
            str(case.get("报关单号") or "").strip(): case
            for case in json.loads(VERIFIED.read_text(encoding="utf-8")).get("cases") or []
        }

    for (decl, vendor, core), group in units.items():
        mine = sum((row["_mine"] for row in group), Decimal(0))
        theirs = group[0]["_theirs"]
        blank = any(row["_blank"] for row in group)
        codes = sorted({str(row["采购单号"]) for row in group})
        # 单元级睿贝参照 = 本单元各记录自己的出运采购金额(RMB)；
        # 采购单级参照（erp_amount，睿贝采购单金额）只在「采购单对账」表里比。
        erp_unit = sum((row["_rmb"] for row in group), Decimal(0))
        for row in group:
            row["ERP出运采购金额(本单元)"] = str(erp_unit) if erp_unit else ""
        ckey = (decl, vendor, coarse_core(core))
        merged = feishu_coarse.get(ckey)
        merged_ok = bool(merged) and abs(coarse_ours.get(ckey, Decimal(0)) - merged) <= EQUAL
        sibling_rows = len(
            {row["采购单核心"] for row in rows if row["报关单号"] == decl
             and row["供应商简称"] == vendor and coarse_core(row["采购单核心"]) == coarse_core(core)}
        )
        scope = "单元"                 # 比对口径：单元 = 报关单+供应商+采购单核心
        merged_diff: Decimal | None = None   # 走「合并口径」时的差额
        if decl and decl in verified_cases:
            case = verified_cases[decl]
            verdict = "已核实"
            reason = str(case.get("类别") or "飞书错")
            evidence = f"{case.get('结论') or ''}｜{case.get('证据') or ''}"
        elif theirs is None:
            if merged_ok:
                # 飞书没有本单元的独立行，但把主单+各附加单并成一行给钱：
                # 比对改成"合并口径"，差额也按合并口径写，避免记录行显示虚高差额
                verdict, reason = "正常", ""
                scope = "合并(主单+ADD)"
                merged_diff = coarse_ours.get(ckey, Decimal(0)) - merged
                evidence = (
                    f"飞书没有本单元的独立行；按「去掉 ADD 的核心号」合并后一致"
                    f"（我方 {money(coarse_ours.get(ckey, Decimal(0)))} = 飞书 {money(merged)}）"
                )
            else:
                verdict = "异常"
                if blank:
                    reason = "睿贝未定价"
                    evidence = "睿贝该行未定价（采购金额留空，待人工补价）"
                else:
                    reason = "飞书无对应行"
                    evidence = f"飞书里找不到该单元（采购单级结论：{po_rows[codes[0]]['结论']}）"
        elif abs(mine - theirs) <= EQUAL:
            verdict, reason, evidence = "正常", "", ""
        elif merged_ok:
            verdict, reason = "正常", ""
            scope = "合并(主单+ADD)"
            merged_diff = coarse_ours.get(ckey, Decimal(0)) - merged
            evidence = (
                f"飞书把主单+附加单并成一行；按合并口径一致"
                f"（我方 {money(coarse_ours.get(ckey, Decimal(0)))} = 飞书 {money(merged)}）"
            )
        else:
            verdict = "异常"
            reason, evidence = unit_reason(
                mine=mine,
                theirs=theirs,
                erp_unit=erp_unit,
                fee=po_fee.get(codes[0], Decimal(0)),
                note=po_rows[codes[0]]["结论"],
                blank=blank,
                sibling=coarse_ours.get(ckey, Decimal(0)),
                sibling_rows=sibling_rows,
                po_fee_info=purchase_fees().get(codes[0], Decimal(0)),
                invoice=dec(po_rows[codes[0]].get("工厂开票金额") or 0),
                po_reconciled=po_reconciled(po_rows[codes[0]]),
                fee_items_text=unit_fee_text(po_rows[codes[0]]),
                fee_ceiling=max(
                    po_fee.get(codes[0], Decimal(0)),                  # 入库单无归属费用
                    dec(po_rows[codes[0]].get("费用行合计") or 0),        # 入库单费用行明细合计
                    purchase_fees().get(codes[0], Decimal(0)),         # 采购订单「费用信息」
                ),
            )
        # 财务 2026-09-24 口径：只要总成本不漏，额外费用摊到哪一票都可以。
        # 因此「整单金额对得上、只是分摊口径（费用摊法）不一样」的单元按**正常**处理，
        # 只加一个标注留痕，方便回头复核。
        mark = ""
        if verdict == "异常" and reason == "分摊口径":
            verdict = "正常"
            mark = "额外费用分摊口径不一致"
        elif verdict == "异常":
            # 2026-09-28 用户口径（二次放宽）：
            # ① 飞书采用「出运采购金额」、我方比它多（差额就是那些额外费用没算进去），
            #    且单据里**确实存在额外费用**（木箱费 / 包装费 / 样块费…）→ 丢进
            #    「成本匹配_待核实.xlsx」，不再要求差额与费用精确对上；
            # ② 早前在成本匹配阶段认定过异常、但没确认的（采购订单费用信息没进入库单）也进待核实；
            # 其余照旧留在「成本匹配_异常.xlsx」。
            explain = ""
            fee_text = unit_fee_text(po_rows[codes[0]])
            if (
                erp_unit
                and theirs is not None
                and abs(theirs - erp_unit) <= EQUAL
                and mine - theirs > EQUAL
                and fee_text
            ):
                itemized, bulk = fee_groups(po_rows[codes[0]])
                explain = fee_explain(mine - theirs, itemized) or fee_explain(
                    mine - theirs, bulk
                )
            if explain:
                verdict = "待核实"
                mark = f"飞书按出运采购金额计，差额 = {explain}（待核实）"
            elif fee_text and erp_unit and theirs is not None and abs(theirs - erp_unit) <= EQUAL and mine - theirs > EQUAL:
                # 放宽后：差额没能和某笔费用精确对上，但单据里确实有额外费用 → 一样进待核实
                verdict = "待核实"
                reason = "飞书按出运采购金额计（漏费用）"
                mark = (
                    f"飞书按出运采购金额计，比我们少 {mine - theirs}"
                    f"，单据里有额外费用（{fee_text}）（待核实）"
                )
            elif reason == "采购订单费用信息未进入库单":
                verdict = "待核实"
                mark = "早前认定为异常、未确认（待核实）"
            group[0]["_fee_explain"] = explain
        # 出运数量核验（2026-09-28）：出运数量比入库数量多 = 凭空冒出来的货。
        # 整组标成「待核实」（用户口径：这类先丢进待核实单，后续不再处理）。
        flag = next((quantity_flags[code] for code in codes if code in quantity_flags), None)
        if flag and verdict != "已核实":
            verdict = "待核实"
            mark = (
                f"出运数量 > 入库数量：出运 {flag['出运数量']}（{flag['出运单']}）"
                f"＞ 入库 {flag['入库数量']}，凭空多出 "
                f"{flag['出运数量'] - flag['入库数量']}"
            )
            reason = "出运数量 > 入库数量"
            evidence = (
                f"出运数量 {flag['出运数量']}（{flag['出运单']}）＞ 入库数量 {flag['入库数量']}"
                f"（{flag['入库单']}），凭空多出 "
                f"{flag['出运数量'] - flag['入库数量']}；按口径整组进待核实（出运单上多出来的货"
                "疑似随手填的数量/金额，需业务确认是漏货、退换货还是多填）"
            )
            scope = "单元"
        for row in group:
            row["判定"] = verdict
            row["标注"] = mark
            row["异常原因"] = reason
            row["归因说明"] = evidence
            row["比对口径"] = scope
            row["差异"] = (
                str(merged_diff)
                if merged_diff is not None
                else ("" if theirs is None else str(mine - theirs))
            )
            row["_verdict"] = verdict

    # ---- 落表
    columns = [
        "报关单号", "合同号_1", "采购单号", "采购单核心", "供应商简称", "产品类型", "报关品名",
        "出运金额合计(USD)", "采购金额_我方", "分摊依据", "采购单实发总额",
        "采购金额_飞书", "飞书行数", "差异", "比对口径", "ERP出运采购金额(本单元)",
        "判定", "标注",
    ]
    detail_columns = columns + ["异常原因", "归因说明"]
    normal = sorted(
        (row for row in rows if row["_verdict"] == "正常"),
        key=lambda row: (str(row["报关单号"]), str(row["采购单号"])),
    )
    bad = sorted(
        (row for row in rows if row["_verdict"] == "异常"),
        key=lambda row: (str(row["异常原因"]), str(row["报关单号"]), str(row["采购单号"])),
    )
    todo_rows = sorted(
        (row for row in rows if row["_verdict"] == "待核实"),
        key=lambda row: (str(row["异常原因"]), str(row["报关单号"]), str(row["采购单号"])),
    )
    done = sorted(
        (row for row in rows if row["_verdict"] == "已核实"),
        key=lambda row: (str(row["报关单号"]), str(row["采购单号"])),
    )
    print(
        f"记录 {len(rows)}：正常 {len(normal)} / 异常 {len(bad)} / "
        f"待核实 {len(todo_rows)} / 已核实 {len(done)}"
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    write_sheet(
        workbook, "全部", detail_columns,
        sorted(rows, key=lambda row: (str(row["报关单号"]), str(row["采购单号"]))),
        first=True,
    )
    save(workbook, OUT_DIR / "成本匹配_全部.xlsx")

    workbook = Workbook()
    write_sheet(workbook, "正常", columns, normal, first=True)
    save(workbook, OUT_DIR / "成本匹配_正常.xlsx")

    # 已知被缺单影响的采购单（整单只覆盖了一部分 / 飞书另有行落在我们没的单上）
    affected_cores = {
        audit_core(code)
        for code, entry in po_rows.items()
        if str(entry["结论"]).startswith(("部分覆盖", "飞书另有"))
    }
    # ---- 待核实清单：差额能被"一笔或几笔额外费用"整笔解释的 + 早前认定未确认的。
    #      人工可能在这张表里加了自己的备注列，写回时按同一条记录保留。
    prev_header, prev_rows = load_verification(VERIFICATION_PATH)
    prev_index = {verification_key(row): row for row in prev_rows}
    todo = []
    for row in verification_rows(todo_rows, po_rows):
        old = prev_index.get(verification_key(row))
        merged = dict(row)
        if old:
            for name, value in old.items():
                if name not in merged and value not in (None, ""):
                    merged[name] = value          # 用户自己加的列，保留
        todo.append(merged)
    workbook = Workbook()
    write_sheet(workbook, "异常明细", detail_columns, bad, fill=BAD_FILL, first=True)
    write_sheet(
        workbook, "采购单对账",
        [
            "采购单号", "供应商", "入库单实发总额", "我方合计", "飞书合计(本单各行)",
            "飞书合计(含ADD)", "飞书行数", "工厂开票金额", "睿贝采购金额", "采购订单费用信息",
            "费用行明细", "入库块数量", "我方−飞书", "结论",
        ],
        [po_rows[code] for code in sorted(po_rows)],
    )
    write_sheet(
        workbook,
        "缺报关单",
        *missing_declaration_sheet(rows, split_rows, feishu_by_core, affected_cores),
    )
    write_sheet(
        workbook,
        "说明",
        ["条目", "说明"],
        [
            {"条目": "本表放什么", "说明": "放**除待核实之外**的异常：普通口径差异、缺报关单、发票与入库单不一致、飞书无对应行等。"},
            {"条目": "什么被移走了", "说明": "两类不在这里：① 飞书按睿贝出运采购金额计、且该采购单单据里存在额外费用（木箱费、包装费…）的；② 早前认定异常但未确认的。这两类在 成本匹配_待核实.xlsx。"},
            {"条目": "已核实的", "说明": "人工核实过的冲突在 成本匹配_已核实.xlsx（来源 data/reference/verified_conflicts.json）。"},
            {"条目": "采购单对账 / 缺报关单", "说明": "两张全量台账，不受上面分类影响。"},
        ],
    )
    save(workbook, OUT_DIR / "成本匹配_异常.xlsx")

    # ---- 待核实清单（2026-09-28 用户口径）：① 飞书按「出运采购金额」给钱、且差额正好
    #      等于一笔/几笔额外费用的；② 早前认定异常但未确认的。每条写清错误原因 + 缺什么数据。
    workbook = Workbook()
    verify_columns = list(VERIFICATION_COLUMNS)
    verify_columns += [name for name in prev_header if name and name not in verify_columns]
    write_sheet(workbook, "待核实", verify_columns, todo, fill=BAD_FILL, first=True)
    counts = collections.Counter(str(row["待核实类型"]) for row in todo)
    write_sheet(
        workbook,
        "汇总",
        ["待核实类型", "条数", "缺什么数据"],
        [
            {"待核实类型": kind, "条数": count, "缺什么数据": MISSING_DATA_HINT.get(kind, "待人工判断")}
            for kind, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))
        ],
    )
    write_sheet(
        workbook,
        "说明",
        ["条目", "说明"],
        [
            {"条目": "本表放什么", "说明": "两类：① 飞书采用了睿贝「出运采购金额」、我方比它多、且该采购单单据里存在额外费用（木箱费/包装费/样块费…）的——不要求差额与费用精确对上；② 早前在成本匹配阶段认定过异常、但没有确认的。其余异常仍留在 成本匹配_异常.xlsx。"},
            {"条目": "怎么来的", "说明": "跑 scripts/build_cost_match.py 自动判定并写出；你在本表里加的备注列会按同一条记录保留。"},
            {"条目": "错误原因", "说明": "谁跟谁差多少、差在哪；「标注」列写明差额等于哪一笔（或哪几笔）费用。"},
            {"条目": "缺什么数据", "说明": "要去拿/去问什么才能定案。"},
            {"条目": "飞书=出运采购金额", "说明": "「是」= 飞书给的就是睿贝出运采购金额（该口径会漏费用，2026-09-28 财务确认）。"},
            {"条目": "单据里能核到的费用", "说明": "该采购单入库单里解析出来的费用行 + 采购订单「费用信息」+ 入库单费用合计，用来对出「差额 = 哪几笔费用」。"},
        ],
    )
    save(workbook, OUT_DIR / "成本匹配_待核实.xlsx")

    workbook = Workbook()
    write_sheet(
        workbook, "已核实",
        [
            "报关单号", "合同号_1", "采购单号", "供应商简称", "产品类型", "报关品名",
            "采购金额_我方", "采购金额_飞书", "差异", "类别", "核实结论", "证据", "核实日期",
        ],
        [
            {
                **row,
                "类别": str(verified_cases.get(str(row["报关单号"]), {}).get("类别") or "飞书错"),
                "核实结论": str(verified_cases.get(str(row["报关单号"]), {}).get("结论") or ""),
                "证据": str(verified_cases.get(str(row["报关单号"]), {}).get("证据") or ""),
                "核实日期": str(verified_cases.get(str(row["报关单号"]), {}).get("核实日期") or ""),
            }
            for row in done
        ],
        first=True,
    )
    save(workbook, OUT_DIR / "成本匹配_已核实.xlsx")

    print("异常构成：", dict(collections.Counter(str(row["异常原因"]) for row in bad)))
    print("输出目录：", OUT_DIR)


if __name__ == "__main__":
    main()
