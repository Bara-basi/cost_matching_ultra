"""采购金额匹配：拆单记录 × 入库单实发金额 × 飞书采购金额 三方对照。

用法：
    & '.\.venv\Scripts\python.exe' scripts\match_cost.py

输出（`outputs/cost_match/`）：
    采购金额_匹配.xlsx     记录级：我方采购金额、飞书采购金额、差异、差异原因
    采购金额_单元差异.xlsx  单元级（报关单 + 供应商 + 采购单）：两侧合计与差异
    采购金额_统计.json
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
    allocate_amounts,
    costs_for,
    dec,
    record_batch_token,
    shipment_rmb_by_kind,
    shipment_rmb_by_po,
    shipment_usd_by_po,
    unit_core,
)
from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.grn_select import erp_amount  # noqa: E402

SPLIT = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
NORMAL = PROJECT_ROOT / "outputs" / "shipments_match" / "匹配结果_正常.xlsx"
REF_CACHE = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs.json"
OUT_DIR = PROJECT_ROOT / "outputs" / "cost_match"

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFE699")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")

TOL = Decimal("0.01")


def read_rows(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    out = [dict(zip(header, row)) for row in rows[1:]]
    workbook.close()
    return out


def po_core(code: str) -> str:
    return unit_core(code)


def coarse_core(code: str) -> str:
    """去掉 ADD 标记的核心号——飞书有时把附加单的金额并进主单那一行。"""
    base = po_core(code)
    return base.split("-ADD")[0].split("ADD")[0]


def main() -> None:
    records = read_rows(SPLIT)
    print(f"拆单记录 {len(records)}")
    rmb_table = shipment_rmb_by_kind()
    usd_by_po = shipment_usd_by_po()
    rmb_by_po = shipment_rmb_by_po()
    print(
        f"ERP 出运产品行 (采购单, 产品类型) 组合 {len(rmb_table)}；"
        f"整单货值索引 {len(rmb_by_po)} 张采购单"
    )

    # ---- 飞书：按 (报关单号, 供应商简称, 采购单核心) 索引
    ref_records = json.loads(REF_CACHE.read_text(encoding="utf-8"))["records"]
    feishu: dict[tuple[str, str, str], list[dict]] = collections.defaultdict(list)
    for row in ref_records:
        decl = str(row.get("报关单号") or "").strip()
        vendor = str(row.get("供应商简称") or "").strip()
        code = str(row.get("采购单号") or "").strip()
        if not decl:
            continue
        feishu[(decl, vendor, po_core(code))].append(row)

    # ---- 按采购单聚合我方记录并分摊
    by_po: dict[str, list[dict]] = collections.defaultdict(list)
    for record in records:
        by_po[str(record.get("采购单号") or "").strip()].append(record)

    batch = costs_for(list(by_po))
    costs: dict[str, dict] = {}
    for code, group in by_po.items():
        total, detail = batch[code]
        batch_map = {k: dec(v) for k, v in (detail.get("batch_amounts") or {}).items()}
        for record in group:
            record["_batch_map"] = batch_map
            record["_batch_token"] = record_batch_token(
                record.get("合同号_1"), record.get("采购单号")
            )
            record["_block_count"] = detail.get("block_count") or 0
        amounts = allocate_amounts(
            group, total, rmb_table, usd_by_po.get(code), rmb_by_po.get(code)
        )
        costs[code] = {"total": total, "detail": detail, "amounts": amounts}

    # ---- 记录级输出
    out_rows: list[dict] = []
    for code, group in by_po.items():
        info = costs[code]
        total = info["total"]
        for record, (amount, basis) in zip(group, info["amounts"]):
            decl = str(record.get("报关单号") or "").strip()
            vendor = str(record.get("供应商简称") or "").strip()
            key = (decl, vendor, po_core(code))
            rows = feishu.get(key) or []
            feishu_total = sum((dec(r.get("采购金额")) for r in rows), Decimal(0))
            unpriced = str(basis).startswith("睿贝该行未定价")
            out_rows.append(
                {
                    "报关单号": decl,
                    "合同号_1": record.get("合同号_1"),
                    "采购单号": code,
                    "采购单核心": po_core(code),
                    "供应商简称": vendor,
                    "产品类型": record.get("产品类型"),
                    "报关品名": record.get("报关品名"),
                    "出运金额合计(USD)": record.get("出运金额合计"),
                    "采购金额_我方": "" if unpriced else str(amount),
                    "分摊依据": basis,
                    "采购单实发总额": str(total),
                    "入库单份数": info["detail"]["kept"],
                    "采购金额_飞书": str(feishu_total) if rows else "",
                    "飞书行数": len(rows),
                    "_key": key,
                    "_amount": Decimal(0) if unpriced else amount,
                    "_unpriced": unpriced,
                    "_feishu": feishu_total if rows else None,
                }
            )

    # ---- 单元级汇总（我方按 (报关单, 供应商, 采购单核心) 合计 vs 飞书）
    mine_units: dict[tuple[str, str, str], Decimal] = collections.defaultdict(Decimal)
    unit_meta: dict[tuple[str, str, str], dict] = {}
    for row in out_rows:
        key = row["_key"]
        mine_units[key] += row["_amount"]
        unit_meta.setdefault(
            key,
            {
                "报关单号": row["报关单号"],
                "采购单核心": row["采购单核心"],
                "供应商简称": row["供应商简称"],
                "产品类型": set(),
                "采购单号": set(),
            },
        )
        unit_meta[key]["产品类型"].add(str(row["产品类型"]))
        unit_meta[key]["采购单号"].add(str(row["采购单号"]))

    unit_rows: list[dict] = []
    stats = collections.Counter()
    # 已人工核实过的冲突：直接引用结论，不再算「未匹配」
    verified: dict[str, dict] = {}
    verified_path = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
    if verified_path.exists():
        payload = json.loads(verified_path.read_text(encoding="utf-8"))
        for case in payload.get("cases") or []:
            decl = str(case.get("报关单号") or "").strip()
            if decl:
                verified[decl] = case
    # 飞书有时把「主单 + 附加单」合并成一行（如 25MT-03P495Y-C 直接写两条记录之和），
    # 所以再按「去掉 ADD 的核心号」聚合一层，用于判定「合并后一致」
    coarse_mine: dict[tuple[str, str, str], Decimal] = collections.defaultdict(Decimal)
    coarse_meta: dict[tuple[str, str, str], list[str]] = collections.defaultdict(list)
    for row in out_rows:
        ckey = (row["报关单号"], row["供应商简称"], coarse_core(row["采购单核心"]))
        coarse_mine[ckey] += row["_amount"]
        coarse_meta[ckey].append(row["采购单核心"])
    feishu_coarse: dict[tuple[str, str, str], Decimal] = collections.defaultdict(Decimal)
    for (decl, vendor, core), rows_ in feishu.items():
        feishu_coarse[(decl, vendor, coarse_core(core))] += sum(
            (dec(r.get("采购金额")) for r in rows_), Decimal(0)
        )
    for key, mine in mine_units.items():
        feishu_rows = feishu.get(key) or []
        theirs = sum((dec(r.get("采购金额")) for r in feishu_rows), Decimal(0)) if feishu_rows else None
        meta = unit_meta[key]
        # 三方对照：E RP 采购金额（同一核心下各工厂采购单之和）作独立参照
        erp_total = Decimal(0)
        for code in meta["采购单号"]:
            erp_total += erp_amount(code) or Decimal(0)
        near_erp = erp_total > 0 and abs(mine - erp_total) / erp_total <= Decimal("0.01")
        near_feishu = theirs is not None and theirs > 0 and abs(mine - theirs) / theirs <= Decimal("0.01")
        if theirs is None:
            ckey = (key[0], key[1], coarse_core(key[2]))
            coarse_theirs = feishu_coarse.get(ckey)
            coarse_ours = coarse_mine.get(ckey)
            if (
                coarse_theirs
                and coarse_ours
                and abs(coarse_ours - coarse_theirs) <= max(
                    Decimal("0.01"), coarse_theirs * Decimal("0.01")
                )
            ):
                verdict = "一致（飞书无独立行，合并到主单后一致）"
                stats["合并后一致（飞书无独立行）"] += 1
            elif key[0] in verified:
                verdict = f"已核实：{verified[key[0]].get('结论')}（{verified[key[0]].get('核实日期')}）"
                stats["已核实"] += 1
            else:
                verdict = "飞书无对应行"
                stats["飞书无对应行"] += 1
        elif near_feishu:
            verdict = "一致"
            stats["一致"] += 1
        elif near_erp:
            verdict = "我方=ERP采购金额，飞书不一致"
            stats["我方与ERP一致、飞书不一致"] += 1
        elif key[0] in verified:
            verdict = f"已核实：{verified[key[0]].get('结论')}（{verified[key[0]].get('核实日期')}）"
            stats["已核实"] += 1
        else:
            diff = mine - theirs
            ckey = (key[0], key[1], coarse_core(key[2]))
            coarse_theirs = feishu_coarse.get(ckey)
            coarse_ours = coarse_mine.get(ckey)
            if (
                coarse_theirs
                and coarse_ours
                and abs(coarse_ours - coarse_theirs) <= max(
                    Decimal("0.01"), coarse_theirs * Decimal("0.01")
                )
            ):
                verdict = "一致（飞书把附加单并进主单后一致）"
                stats["合并后一致（附加单粒度差异）"] += 1
            elif erp_total > 0:
                verdict = f"三方都不一致（我方 {mine} / ERP {erp_total} / 飞书 {theirs}）"
                stats["三方都不一致"] += 1
            else:
                verdict = f"差 {diff:+}（无 ERP 参照）"
                stats["有差异（无 ERP 参照）"] += 1
        def ratio(text_mine: Decimal, other) -> str:
            if other in (None, "", 0):
                return ""
            base = Decimal(str(other))
            if base == 0:
                return ""
            return f"{abs(text_mine - base) / base * 100:.2f}%"

        unit_rows.append(
            {
                "报关单号": meta["报关单号"],
                "供应商简称": meta["供应商简称"],
                "采购单核心": meta["采购单核心"],
                "采购单号": "、".join(sorted(meta["采购单号"])),
                "产品类型": "、".join(sorted(x for x in meta["产品类型"] if x)),
                "采购金额_我方": str(mine),
                "采购金额_飞书": "" if theirs is None else str(theirs),
                "ERP采购金额": str(erp_total) if erp_total else "",
                "差异": "" if theirs is None else str(mine - theirs),
                "偏差_飞书": ratio(mine, theirs),
                "偏差_ERP": ratio(mine, erp_total if erp_total else None),
                "结论": verdict,
                "_diff": None if theirs is None else mine - theirs,
            }
        )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    write(
        out_rows,
        [
            "报关单号", "合同号_1", "采购单号", "采购单核心", "供应商简称", "产品类型", "报关品名",
            "出运金额合计(USD)", "采购金额_我方", "分摊依据", "分摊比例", "采购单实发总额",
            "入库单份数", "采购金额_飞书", "飞书行数",
        ],
        OUT_DIR / "采购金额_匹配.xlsx",
    )
    write(
        unit_rows,
        [
            "报关单号", "供应商简称", "采购单核心", "采购单号", "产品类型",
            "采购金额_我方", "采购金额_飞书", "ERP采购金额", "差异",
            "偏差_飞书", "偏差_ERP", "结论",
        ],
        OUT_DIR / "采购金额_单元差异.xlsx",
        highlight="结论",
    )

    total_records = len(out_rows)
    summary = {
        "拆单记录": total_records,
        "采购单": len(by_po),
        "入库单实发合计为 0 的采购单": sum(1 for x in costs.values() if x["total"] == 0),
        "单元数": len(unit_rows),
        **dict(stats),
    }
    (OUT_DIR / "采购金额_统计.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    print("输出目录:", OUT_DIR)


def write(rows: list[dict], columns: list[str], path: Path, highlight: str | None = None) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "结果"
    sheet.append(columns)
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(name, "") for name in columns])
        if highlight:
            value = str(row.get(highlight) or "")
            if value and value != "一致":
                fill = BAD_FILL if value.startswith("差") else WARN_FILL
                for cell in sheet[sheet.max_row]:
                    cell.fill = fill
    for index, name in enumerate(columns, 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            12, min(34, len(name) * 2 + 4)
        )
    sheet.freeze_panes = "A2"
    try:
        workbook.save(path)
    except PermissionError:
        workbook.save(path.with_name(f"{path.stem}_new{path.suffix}"))
    print(f"written {path} ({len(rows)} rows)")


if __name__ == "__main__":
    main()
