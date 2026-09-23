r"""拆单记录 × 飞书报关金额对照（不含客户费用分摊）。

口径
====
    我方 报关金额 = 该记录的「出运金额合计」（= 睿贝出运产品行金额，**不加客户费用分摊**）
    飞书 报关金额 = 「迈拓财务部门数据 副本 / 2026年报关数据」里那一行的报关金额

带客户费用分摊的记录（`客户费用分摊 != 0`）本轮**不参与金额比对**，直接报异常
（费用口径未定，硬算会把差额混进结论里）。

要回答的问题
============
「产品类型对不上的记录，金额是不是也对不上？」——产品类型用**睿贝实际结果**口径
（`海关商品（中文）` 经「产品及类型」映射表折算），不是报关品名折算值。
输出一张交叉表：类型一致/不一致 × 金额一致/不一致。

比对层级
========
* 记录级：同一报关单内按 `供应商 + 产品类型` → `供应商 + 报关品名` → `供应商`
  逐级找**唯一**的飞书对应行；找不到唯一对应的记录不硬配。
* 报关单 × 供应商级：直接比**合计**（不需要配对，作为交叉表的基准）。

输出（`outputs/shipments_match/`）
----------------------------------
    报关金额_明细.xlsx        我方 460 条记录 + 飞书报关金额 + 差额 + 结论
    报关金额_正常.xlsx        报关单×供应商 分组，金额与类型都对得上且无费用
    报关金额_异常.xlsx        其余分组（含费用 / 金额不符 / 类型不符 / 无法配对）
    报关金额_统计.json        含「类型 × 金额」交叉表
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from app.services.feishu_client import FeishuClient  # noqa: E402
from app.services.contract_shipments import cores_of  # noqa: E402
from app.services.shipment_detail import shipment_invoices_for_contract  # noqa: E402
from match_feishu import (  # noqa: E402
    MAP_CACHE,
    REF_APP,
    REF_NAME,
    REF_TABLE,
    coarse_map,
    flat,
    product_option_map,
    split_multi,
)
from reverse_fee_allocation import shipment_fees  # noqa: E402

SPLIT_XLSX = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
REF_CACHE = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs.json"
VERIFIED = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
OUT_DIR = PROJECT_ROOT / "outputs" / "shipments_match"

TOL = Decimal("0.02")
NORMAL = "正常"
# 已确认口径（费用按销售订单号整笔归属到记录）的费用类型：先把这部分计入报关金额
CONFIRMED_FEES = ("罚金", "折扣")

DETAIL_COLUMNS = (
    "报关单号", "合同号_1", "报关品名", "海关编码",
    "产品类型（ERP）", "产品类型（ERP折算）", "产品类型（报关品名折算）",
    "供应商简称", "采购单号", "出运金额合计", "客户费用分摊",
    "已计入费用（罚金/折扣）", "未确认费用",
    "报关金额（我方）", "飞书报关金额", "差额（飞书-我方）", "配对方式", "结论",
)
GROUP_COLUMNS = (
    "报关单号", "合同号_1", "供应商简称",
    "系统产品类型", "飞书产品类型", "产品类型是否一致",
    "我方金额合计", "飞书金额合计", "差额（飞书-我方）",
    "含客户费用分摊", "对照情况", "结论",
)

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFE699")

BAD_FEE = "异常（含客户费用分摊）"
BAD_AMOUNT = "异常（金额不符）"
BAD_NOMATCH = "异常（飞书无对应行）"
BAD_PAIR = "异常（无法配对）"
NOT_PRODUCED = "异常（我方未拆出）"


def to_dec(value) -> Decimal:
    text = flat(value).replace(",", "")
    if not text:
        return Decimal(0)
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def money(value: Decimal) -> str:
    return f"{value.quantize(Decimal('0.01')):.2f}"


def set_text(values) -> str:
    return "、".join(sorted({v for v in values if v}))


# --------------------------------------------------------------------------- 读表


def read_ours() -> list[dict]:
    workbook = load_workbook(SPLIT_XLSX, read_only=True)
    rows = list(workbook.active.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(header)}
    out: list[dict] = []
    for row in rows[1:]:
        decl = flat(row[index["报关单号"]]).replace(" ", "")
        if not decl:
            continue
        out.append(
            {
                "报关单号": decl,
                "合同号_1": flat(row[index["合同号_1"]]),
                "报关品名": flat(row[index["报关品名"]]),
                "海关编码": flat(row[index["海关编码"]]),
                "产品类型": flat(row[index["产品类型"]]),
                "供应商简称": flat(row[index["供应商简称"]]),
                "采购单号": flat(row[index["采购单号"]]),
                "出运金额合计": to_dec(row[index["出运金额合计"]]),
                "客户费用分摊": to_dec(row[index["客户费用分摊"]]),
            }
        )
    return out


def fetch_ref(offline: bool) -> tuple[list[dict], dict[str, str]]:
    """飞书对照表：报关单号 + 报关品名 + 产品类型 + 供应商简称 + 报关金额。"""
    if offline and REF_CACHE.exists():
        payload = json.loads(REF_CACHE.read_text(encoding="utf-8"))
        mapping = payload.get("coarse_map") or {}
        if not mapping and MAP_CACHE.exists():
            mapping = json.loads(MAP_CACHE.read_text(encoding="utf-8"))
        return payload["records"], mapping

    client = FeishuClient()
    options = product_option_map(client)
    mapping = coarse_map(client)
    records: list[dict] = []
    for item in client.iter_records(REF_APP, REF_TABLE, page_size=500):
        fields = item.get("fields") or {}
        records.append(
            {
                "报关单号": fields.get("报关单号"),
                "合同号_1": fields.get("合同号_1"),
                "报关品名": fields.get("报关品名"),
                "产品类型": [options.get(t, t) for t in split_multi(fields.get("产品类型"))],
                "供应商简称": fields.get("供应商简称"),
                "报关金额": fields.get("报关金额"),
            }
        )
    REF_CACHE.parent.mkdir(parents=True, exist_ok=True)
    REF_CACHE.write_text(
        json.dumps(
            {"app": REF_APP, "table": REF_TABLE, "records": records, "coarse_map": mapping},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return records, mapping


def normalise_ref(records: list[dict]) -> list[dict]:
    out: list[dict] = []
    for record in records:
        decl = flat(record.get("报关单号")).replace(" ", "")
        if not decl:
            continue
        out.append(
            {
                "报关单号": decl,
                "合同号_1": flat(record.get("合同号_1")),
                "报关品名": flat(record.get("报关品名")),
                "产品类型": split_multi(record.get("产品类型")),
                "供应商简称": flat(record.get("供应商简称")),
                "报关金额": to_dec(record.get("报关金额")),
            }
        )
    return out


# ------------------------------------------------------------------------- 主流程


def pair_records(ours: list[dict], ref_rows: list[dict]) -> None:
    """在同一报关单内，按 供应商+产品类型 → 供应商+品名 → 供应商 找唯一对应行。"""
    candidates: dict[str, list[dict]] = defaultdict(list)
    for row in ref_rows:
        candidates[row["报关单号"]].append(row)

    used: set[int] = set()
    for record in ours:
        pool = candidates.get(record["报关单号"], [])
        pair = None
        how = ""
        for level in ("供应商+产品类型", "供应商+报关品名", "供应商"):
            hits: list[dict] = []
            for row in pool:
                if id(row) in used or row["供应商简称"] != record["供应商简称"]:
                    continue
                if level == "供应商+产品类型":
                    if record["产品类型（ERP折算）"] in row["产品类型"]:
                        hits.append(row)
                elif level == "供应商+报关品名":
                    if row["报关品名"] == record["报关品名"]:
                        hits.append(row)
                else:
                    hits.append(row)
            if len(hits) == 1:
                pair, how = hits[0], level
                break
        if pair is not None:
            used.add(id(pair))
        record["_pair"] = pair
        record["配对方式"] = how or ("无法配对" if pool else "飞书无同行")

    # 第二遍：同一（报关单，供应商）内两边都有多行且行数相同时，按金额降序一一配对
    leftover: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for record in ours:
        if record["_pair"] is None:
            leftover[(record["报关单号"], record["供应商简称"])].append(record)
    for key, records in leftover.items():
        pool = [
            row for row in candidates.get(key[0], [])
            if id(row) not in used and row["供应商简称"] == key[1]
        ]
        if not pool:
            continue
        records.sort(key=lambda r: r["报关金额（我方）"], reverse=True)
        pool.sort(key=lambda r: r["报关金额"], reverse=True)
        for record, row in zip(records, pool):
            record["_pair"] = row
            record["配对方式"] = "按金额降序配对"
            used.add(id(row))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--offline", action="store_true", help="用本地缓存，不拉飞书")
    args = parser.parse_args()

    blocked: set[str] = set()
    if VERIFIED.exists():
        blocked = {
            str(case.get("报关单号") or "").strip()
            for case in (json.loads(VERIFIED.read_text(encoding="utf-8")).get("cases") or [])
            if case.get("报关单号")
        }
    ours_all = read_ours()
    ours = [record for record in ours_all if record["报关单号"] not in blocked]
    ours_by_decl: dict[str, list[dict]] = defaultdict(list)
    for record in ours:
        ours_by_decl[record["报关单号"]].append(record)
    ref_records, mapping = fetch_ref(args.offline)
    ref_all = normalise_ref(ref_records)
    # 只比我们有拆单结果的报关单；飞书侧其余行（我方无 PDF）不参与
    our_decls = {record["报关单号"] for record in ours}
    ref_rows = [
        row for row in ref_all
        if row["报关单号"] in our_decls and row["报关单号"] not in blocked
    ]

    for record in ours:
        record["产品类型（ERP折算）"] = mapping.get(record["产品类型"], record["产品类型"])
        record["产品类型（报关品名折算）"] = mapping.get(record["报关品名"], "其他")
        # 本轮口径：报关金额 = 出运金额合计 + 已确认费用（罚金/折扣，按销售订单号整笔归属）
        record["已确认费用"] = Decimal(0)
        record["未确认费用"] = Decimal(0)
        record["_fee_notes"] = []
        record["报关金额（我方）"] = record["出运金额合计"]
        if not record["产品类型"]:
            record["_type_state"] = "我方缺失"
        elif record["产品类型"] not in mapping:
            record["_type_state"] = "对照表未收录"
        else:
            record["_type_state"] = "可比"

    # ---- 已确认费用层：罚金 / 折扣 ----
    fee_by_decl: dict[str, list[dict]] = {}
    for decl, records in ours_by_decl.items():
        invoices: list[str] = []
        for code in {r["合同号_1"] for r in records}:
            invoices.extend(shipment_invoices_for_contract(code))
        fee_by_decl[decl] = [
            fee for invoice in dict.fromkeys(invoices) for fee in shipment_fees(invoice)
        ]
    for record in ours:
        mine = cores_of(record["采购单号"])
        for fee in fee_by_decl.get(record["报关单号"], []):
            if not (cores_of(fee["订单"]) & mine):
                continue
            signed = (fee["方向"] or 1) * fee["金额"]
            if fee["名称"] in CONFIRMED_FEES:
                record["已确认费用"] += signed
                record["_fee_notes"].append(f"{fee['名称']} {signed:+}")
            else:
                record["未确认费用"] += signed
                record["_fee_notes"].append(f"{fee['名称']} {signed:+}（待定）")
        record["报关金额（我方）"] = record["出运金额合计"] + record["已确认费用"]

    pair_records(ours, ref_rows)

    # ---------------- 记录级结论 ----------------
    detail_rows: list[dict] = []
    for record in ours:
        pair = record["_pair"]
        mine = record["报关金额（我方）"]
        ref_amount = pair["报关金额"] if pair else Decimal(0)
        delta = ref_amount - mine

        problems: list[str] = []
        if record["未确认费用"]:
            problems.append(
                f"异常（含未确认费用 {money(record['未确认费用'])}："
                + "、".join(n for n in record["_fee_notes"] if "待定" in n)
                + "）"
            )
        if pair is None:
            problems.append(BAD_NOMATCH if record["配对方式"] == "飞书无同行" else BAD_PAIR)
        elif abs(delta) > TOL:
            problems.append(BAD_AMOUNT)

        detail_rows.append(
            {
                "报关单号": record["报关单号"],
                "合同号_1": record["合同号_1"],
                "报关品名": record["报关品名"],
                "海关编码": record["海关编码"],
                "产品类型（ERP）": record["产品类型"],
                "产品类型（ERP折算）": record["产品类型（ERP折算）"],
                "产品类型（报关品名折算）": record["产品类型（报关品名折算）"],
                "供应商简称": record["供应商简称"],
                "采购单号": record["采购单号"],
                "出运金额合计": money(record["出运金额合计"]),
                "客户费用分摊": money(record["客户费用分摊"]),
                "已计入费用（罚金/折扣）": money(record["已确认费用"]),
                "未确认费用": money(record["未确认费用"]),
                "报关金额（我方）": money(mine),
                "飞书报关金额": money(ref_amount) if pair else "",
                "差额（飞书-我方）": money(delta) if pair else "",
                "配对方式": record["配对方式"],
                "结论": "；".join(problems) if problems else NORMAL,
            }
        )

    # ---------------- 报关单 × 供应商 分组 ----------------
    ours_groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for record in ours:
        ours_groups[(record["报关单号"], record["供应商简称"])].append(record)
    ref_groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in ref_rows:
        ref_groups[(row["报关单号"], row["供应商简称"])].append(row)

    same_rows: list[dict] = []
    bad_rows: list[dict] = []
    cross: Counter = Counter()
    cross_clean: Counter = Counter()
    for key in sorted(set(ours_groups) | set(ref_groups)):
        mine = ours_groups.get(key, [])
        theirs = ref_groups.get(key, [])
        my_types = {r["产品类型（ERP折算）"] for r in mine} - {""}
        ref_types = {t for r in theirs for t in r["产品类型"]}
        my_amount = sum((r["报关金额（我方）"] for r in mine), Decimal(0))
        ref_amount = sum((r["报关金额"] for r in theirs), Decimal(0))
        delta = ref_amount - my_amount
        fee = sum((r["客户费用分摊"] for r in mine), Decimal(0))
        type_same = bool(mine) and bool(theirs) and my_types == ref_types
        amount_same = bool(mine) and bool(theirs) and abs(delta) <= TOL
        states = {r["_type_state"] for r in mine}
        if "我方缺失" in states:
            type_state = "我方缺失"
        elif "对照表未收录" in states:
            type_state = "对照表未收录"
        else:
            type_state = "一致" if type_same else "不符"

        if mine and theirs:
            cross[(type_state, amount_same)] += 1
            if not fee:
                cross_clean[(type_state, amount_same)] += 1

        problems: list[str] = []
        if not mine:
            problems.append(NOT_PRODUCED)
        if mine and not theirs:
            problems.append(BAD_NOMATCH)
        if fee:
            problems.append(BAD_FEE)
        if mine and theirs and not amount_same:
            problems.append(BAD_AMOUNT)
        if mine and theirs and not type_same:
            problems.append(
                f"异常（产品类型不符：我方 {set_text(my_types)} / 飞书 {set_text(ref_types)}）"
            )

        row = {
            "报关单号": key[0],
            "合同号_1": (mine[0]["合同号_1"] if mine else theirs[0]["合同号_1"]),
            "供应商简称": key[1],
            "系统产品类型": set_text(my_types),
            "飞书产品类型": set_text(ref_types),
            "产品类型是否一致": type_state,
            "我方金额合计": money(my_amount),
            "飞书金额合计": money(ref_amount),
            "差额（飞书-我方）": money(delta),
            "含客户费用分摊": money(fee) if fee else "",
            "对照情况": "双方都有" if (mine and theirs) else ("飞书有我方无" if theirs else "飞书无对应行"),
            "结论": NORMAL if not problems else "；".join(problems),
        }
        (same_rows if not problems else bad_rows).append(row)

    write(detail_rows, DETAIL_COLUMNS, OUT_DIR / "报关金额_明细.xlsx", highlight="结论")
    write(same_rows, GROUP_COLUMNS, OUT_DIR / "报关金额_正常.xlsx")
    write(bad_rows, GROUP_COLUMNS, OUT_DIR / "报关金额_异常.xlsx", highlight="结论")

    stats = {
        "对照表": REF_NAME,
        "我方口径": "报关金额 = 出运金额合计 + 已确认费用（罚金/折扣，按销售订单号整笔归属）",
        "已计入费用（罚金/折扣）的记录": sum(1 for r in ours if r["已确认费用"]),
        "仍带未确认费用的记录": sum(1 for r in ours if r["未确认费用"]),
        "我方记录数": len(ours),
        "上一步已核实/异常已排除的记录": len(ours_all) - len(ours),
        "飞书对照行数": len(ref_rows),
        "记录级结论": dict(Counter(r["结论"] for r in detail_rows)),
        "分组数(报关单×供应商)": len(set(ours_groups) | set(ref_groups)),
        "分组-正常": len(same_rows),
        "分组-异常": len(bad_rows),
        "交叉表(全部可比分组)": {
            f"{state}&金额{'一致' if amount else '不符'}": count
            for (state, amount), count in sorted(cross.items())
        },
        "交叉表(只看无客户费用的分组)": {
            f"{state}&金额{'一致' if amount else '不符'}": count
            for (state, amount), count in sorted(cross_clean.items())
        },
    }
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "报关金额_统计.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=1))


def write(rows: list[dict], columns, path: Path, highlight: str | None = None) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "结果"
    sheet.append(list(columns))
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(name, "") for name in columns])
        if highlight and row.get(highlight) and row[highlight] != NORMAL:
            fill = BAD_FILL if "金额不符" in str(row[highlight]) else WARN_FILL
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), columns):
        sheet.column_dimensions[column[0].column_letter].width = max(12, min(30, len(name) * 2 + 4))
    sheet.freeze_panes = "A2"
    try:
        workbook.save(path)
        print(f"written {path} ({len(rows)} rows)")
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_new{path.suffix}")
        workbook.save(fallback)
        print(f"!! {path.name} 被占用（Excel 打开中），已写到 {fallback.name} ({len(rows)} rows)")


if __name__ == "__main__":
    main()
