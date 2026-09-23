r"""核查「同一报关单 + 同一采购单 + 同一供应商 下有两种实际品名」的报关单。

这类报关单我们按三级拆分口径（采购订单 → 产品 → 供应商）会出两条记录，
需要逐张确认：睿贝采购单里是否真的有两种货、飞书是不是少写了一行。

输入
====
    outputs/shipments_split/拆单明细_全部.xlsx     我方拆单结果
    data/cache/ref_2026_customs.json               飞书对照表
    .cache/erp/index/shipment_lines.jsonl          睿贝出运产品行（采购单里的实际品名）
    .cache/erp/purchases/index.json                采购单总额

输出
====
    outputs/shipments_match/混装核查.xlsx
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font  # noqa: E402

SPLIT_XLSX = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
REF_CACHE = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs.json"
SHIPMENT_LINES = PROJECT_ROOT / ".cache" / "erp" / "index" / "shipment_lines.jsonl"
PURCHASE_INDEX = PROJECT_ROOT / ".cache" / "erp" / "purchases" / "index.json"
PRODUCT_MAP = PROJECT_ROOT / "data" / "reference" / "feishu_product_map.json"
VERIFIED = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
OUT = PROJECT_ROOT / "outputs" / "shipments_match" / "混装核查.xlsx"
MISSING_OUT = PROJECT_ROOT / "outputs" / "shipments_match" / "飞书漏拆待补行.xlsx"

COLUMNS = (
    "报关单号", "合同号_1", "供应商简称", "采购单号",
    "睿贝采购单内实际品名", "采购单总金额(RMB)",
    "我方记录数", "我方类型/出运金额",
    "飞书行数", "飞书品名(类型)/报关金额/采购金额",
    "飞书采购金额合计", "判定",
)

MISSING_COLUMNS = (
    "报关单号", "合同号_1", "供应商简称", "采购单号", "问题类型",
    "建议报关品名（实际品名）", "产品类型", "应填报关金额", "客户费用分摊",
    "飞书现有对应行", "依据",
)


def dec(value) -> Decimal:
    text = str(value or "").replace(",", "").strip()
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def money(value: Decimal) -> str:
    return f"{value.quantize(Decimal('0.01')):f}"


def main() -> None:
    workbook = load_workbook(SPLIT_XLSX, read_only=True)
    rows = list(workbook.active.iter_rows(values_only=True))
    header = [str(c) for c in rows[0]]
    ours = [dict(zip(header, row)) for row in rows[1:]]
    mapping = json.loads(PRODUCT_MAP.read_text(encoding="utf-8")) if PRODUCT_MAP.exists() else {}

    blocked: set[str] = set()
    if VERIFIED.exists():
        blocked = {
            str(case.get("报关单号") or "").strip()
            for case in (json.loads(VERIFIED.read_text(encoding="utf-8")).get("cases") or [])
            if case.get("报关单号")
        }
    ours = [row for row in ours if str(row.get("报关单号") or "").strip() not in blocked]

    ref = json.loads(REF_CACHE.read_text(encoding="utf-8"))["records"]
    ref_by_decl: dict[str, list[dict]] = defaultdict(list)
    for item in ref:
        decl = str(item.get("报关单号") or "").strip()
        if decl:
            ref_by_decl[decl].append(item)

    po_names: dict[str, set[str]] = defaultdict(set)
    with SHIPMENT_LINES.open(encoding="utf-8") as handle:
        for line in handle:
            item = json.loads(line)
            code = str(item.get("purchase_code") or "").strip()
            name = str(item.get("customs_name") or "").strip()
            if code and name:
                po_names[code].add(name)
    purchase_index = json.loads(PURCHASE_INDEX.read_text(encoding="utf-8"))

    groups: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for row in ours:
        key = (
            str(row.get("报关单号") or "").strip(),
            str(row.get("采购单号") or "").strip(),
            str(row.get("供应商简称") or "").strip(),
        )
        groups[key].append(row)
    targets = sorted(
        {
            key[0]
            for key, items in groups.items()
            if len({str(r.get("产品类型") or "") for r in items} - {""}) >= 2
        }
    )

    out_rows: list[dict] = []
    missing_rows: list[dict] = []
    for decl in targets:
        mine = [r for r in ours if str(r.get("报关单号") or "").strip() == decl]
        theirs = ref_by_decl.get(decl, [])
        codes = sorted({str(r.get("采购单号") or "") for r in mine})
        names = sorted({n for code in codes for n in po_names.get(code, set())})
        po_total = sum(
            (dec((purchase_index.get(code) or {}).get("amount")) for code in codes),
            Decimal(0),
        )
        ref_purchase = sum((dec(item.get("采购金额")) for item in theirs), Decimal(0))
        feishu_types = {t for item in theirs for t in (item.get("产品类型") or [])}
        gap = len(mine) - len(theirs)
        verdict = (
            f"飞书少拆 {gap} 行" if gap > 0
            else ("行数一致" if gap == 0 else f"飞书多 {-gap} 行")
        )
        out_rows.append(
            {
                "报关单号": decl,
                "合同号_1": mine[0].get("合同号_1"),
                "供应商简称": "、".join(sorted({str(r.get("供应商简称") or "") for r in mine})),
                "采购单号": "、".join(codes),
                "睿贝采购单内实际品名": "、".join(names),
                "采购单总金额(RMB)": money(po_total),
                "我方记录数": len(mine),
                "我方类型/出运金额": "；".join(
                    f"{r.get('产品类型')} {r.get('出运金额合计')}" for r in mine
                ),
                "飞书行数": len(theirs),
                "飞书品名(类型)/报关金额/采购金额": "；".join(
                    f"{item.get('报关品名')}({'、'.join(item.get('产品类型') or [])}) "
                    f"{item.get('报关金额')}/{item.get('采购金额')}"
                    for item in theirs
                ),
                "飞书采购金额合计": money(ref_purchase),
                "判定": verdict,
            }
        )

        # 我方有、飞书没有对应（同一供应商 + 同一实际类型）的记录 = 飞书该补的行
        for record in mine:
            kind = mapping.get(str(record.get("产品类型") or ""), str(record.get("产品类型") or ""))
            supplier = str(record.get("供应商简称") or "")
            amount = dec(record.get("出运金额合计"))
            if kind and kind in feishu_types:
                continue
            same_amount = [
                item for item in theirs
                if str(item.get("供应商简称") or "") == supplier
                and abs(dec(item.get("报关金额")) - amount) <= Decimal("0.02")
            ]
            problem = "品名/类型写错（飞书那行金额对得上）" if same_amount else "漏拆（飞书缺这一行）"
            missing_rows.append(
                {
                    "报关单号": decl,
                    "合同号_1": record.get("合同号_1"),
                    "供应商简称": supplier,
                    "采购单号": record.get("采购单号"),
                    "问题类型": problem,
                    "建议报关品名（实际品名）": record.get("产品类型"),
                    "产品类型": kind,
                    "应填报关金额": record.get("出运金额合计"),
                    "客户费用分摊": record.get("客户费用分摊"),
                    "飞书现有对应行": "；".join(
                        f"{item.get('报关品名')}({'、'.join(item.get('产品类型') or [])})"
                        f" {item.get('报关金额')} 采购{item.get('采购金额')}"
                        for item in theirs
                    ),
                    "依据": (
                        f"睿贝采购单 {'、'.join(codes)} 同时含 {'、'.join(names)}；"
                        f"飞书该单 {len(theirs)} 行、采购金额合计 {money(ref_purchase)}，"
                        f"采购单总额 {money(po_total)}"
                    ),
                }
            )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "混装核查"
    sheet.append(list(COLUMNS))
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for row in out_rows:
        sheet.append([row.get(name, "") for name in COLUMNS])
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), COLUMNS):
        sheet.column_dimensions[column[0].column_letter].width = max(12, min(46, len(name) * 2 + 4))
    sheet.freeze_panes = "A2"
    book.save(OUT)

    miss_book = Workbook()
    miss_sheet = miss_book.active
    miss_sheet.title = "飞书漏拆待补行"
    miss_sheet.append(list(MISSING_COLUMNS))
    for cell in miss_sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for row in missing_rows:
        miss_sheet.append([row.get(name, "") for name in MISSING_COLUMNS])
    for column, name in zip(miss_sheet.iter_cols(min_row=1, max_row=1), MISSING_COLUMNS):
        miss_sheet.column_dimensions[column[0].column_letter].width = max(12, min(52, len(name) * 2 + 4))
    miss_sheet.freeze_panes = "A2"
    miss_book.save(MISSING_OUT)

    counts: dict[str, int] = defaultdict(int)
    for row in out_rows:
        counts[row["判定"]] += 1
    print(f"命中 {len(out_rows)} 张，判定分布：")
    for key, value in sorted(counts.items()):
        print(f"   {key}: {value}")
    problem_counts: dict[str, int] = defaultdict(int)
    for row in missing_rows:
        problem_counts[row["问题类型"]] += 1
    print(f"待补行 {len(missing_rows)} 条：")
    for key, value in sorted(problem_counts.items()):
        print(f"   {key}: {value}")
    print(f"written {OUT}")
    print(f"written {MISSING_OUT}")


if __name__ == "__main__":
    main()
