"""把全部「与飞书/睿贝不一致」的单元做成一份按原因分类的差异台账。

分类：
  * 单据冲突   —— 我方金额与飞书、睿贝都差很多（说明两边单据本身打架）
  * 分摊口径   —— 我方金额与其中一侧接近（≤3%），差异来自跨报关单怎么分钱
  * 睿贝未定价 —— 睿贝出运行没有金额，我方留空或按 USD 分摊
  * 已核实     —— `verified_conflicts.json` 里有结论
  * 飞书无行   —— 飞书里找不到对应记录

输出：`outputs/cost_match/差异台账.xlsx`
"""
from __future__ import annotations

import json
import re
import sys
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from app.services.cost_match import dec  # noqa: E402
from app.services.grn_extract import PARSED_ROOT, safe_name  # noqa: E402
from app.services.grn_select import erp_amount, select  # noqa: E402

OUT_DIR = PROJECT_ROOT / "outputs" / "cost_match"
SPLIT = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
FULL_REF = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs_full.json"
FILLS = {
    "单据冲突": PatternFill("solid", fgColor="FFC7CE"),
    "缺报关单（待业务补）": PatternFill("solid", fgColor="FCE4D6"),
    "发票与入库单不一致": PatternFill("solid", fgColor="D9D2E9"),
    "单价口径": PatternFill("solid", fgColor="DDEBF7"),
    "单据口径（实发含费用行）": PatternFill("solid", fgColor="FCE4D6"),
    "分摊口径": PatternFill("solid", fgColor="FFE699"),
    "睿贝未定价": PatternFill("solid", fgColor="DDEBF7"),
    "飞书无行": PatternFill("solid", fgColor="E2EFDA"),
    "已核实": PatternFill("solid", fgColor="EDEDED"),
}


def read_rows(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    out = [dict(zip(header, row)) for row in rows[1:]]
    workbook.close()
    return out


def ratio(value: Decimal, other) -> Decimal | None:
    if other in (None, ""):
        return None
    base = dec(other)
    if base == 0:
        return None
    return abs(value - base) / base


def main() -> None:
    units = read_rows(OUT_DIR / "采购金额_单元差异.xlsx")
    details = read_rows(OUT_DIR / "采购金额_匹配.xlsx")
    split_rows = read_rows(SPLIT) if SPLIT.exists() else []
    rmb_by_unit: dict[tuple[str, str], Decimal] = {}
    for row in split_rows:
        key = (str(row.get("报关单号") or ""), str(row.get("采购单号") or ""))
        rmb_by_unit[key] = rmb_by_unit.get(key, Decimal(0)) + dec(
            row.get("出运采购金额合计")
        )
    ref_records = (
        json.loads(FULL_REF.read_text(encoding="utf-8")).get("records") or []
        if FULL_REF.exists()
        else []
    )

    def ref_sum_for(code: str, supplier: str = "") -> tuple[Decimal, int]:
        """飞书里同一采购单（含 ADD 附加单）全部行的采购金额合计。"""
        rows = ref_rows_for(code, supplier)
        return sum((dec(record.get("采购金额")) for record in rows), Decimal(0)), len(rows)

    def ref_invoice_sum(code: str, supplier: str = "") -> tuple[Decimal, int]:
        """飞书里该采购单的**工厂开票金额**合计。

        实测 817 条有采购金额的记录里，`采购金额 = 开票金额 = 实际成本` **100% 一致** ——
        财务那一列就是工厂发票金额。所以「我方（入库单实发） vs 飞书（发票）」的差
        本质上是**发票与入库单的差**（补款、过磅差、费用行口径）。
        """
        rows = ref_rows_for(code, supplier)
        return sum((dec(record.get("开票金额")) for record in rows), Decimal(0)), len(rows)

    def ref_rows_for(code: str, supplier: str) -> list[dict]:
        """飞书里属于该采购单的行。

        单号要多归一一点：我们的采购单号带工厂后缀（`26MT-02N182-JX`），
        飞书那边常常不写后缀（`26MT-02N182A`），只按整串比会一条都匹配不上。
        所以先用"订单核心"匹配，再用供应商简称收敛；供应商这一层匹配不到就退回核心匹配。
        """
        core = re.sub(r"-ADD\d*$", "", str(code or "").upper())
        base = re.sub(r"[-_][A-Za-z]{2,5}$", "", core)
        norm = base.replace("-", "").replace("_", "").upper()
        if not norm:
            return []
        hits: list[dict] = []
        for record in ref_records:
            text = (
                str(record.get("合同号（应收表格）") or "")
                + str(record.get("合同号_1") or "")
            ).upper()
            if norm in text.replace("-", "").replace("_", ""):
                hits.append(record)
        if supplier:
            filtered = [
                record
                for record in hits
                if str(record.get("供应商简称") or "").strip() in ("", supplier)
                or supplier in str(record.get("供应商简称") or "")
            ]
            if filtered:
                return filtered
        return hits

    def ref_amount(decl: str, code: str) -> Decimal | None:
        """飞书里这条（报关单 × 本采购单）的采购金额。"""
        norm = code.replace("-", "").replace("_", "").upper()
        for record in ref_records:
            if str(record.get("报关单号") or "").strip() != decl:
                continue
            text = (
                str(record.get("合同号（应收表格）") or "")
                + str(record.get("合同号_1") or "")
            ).upper()
            if norm and norm in text.replace("-", "").replace("_", ""):
                return dec(record.get("采购金额"))
        return None

    def sibling_total(row: dict) -> tuple[Decimal, int]:
        """同一报关单上、同一采购单家族（主单+ADD）的我方金额合计。"""
        core = re.sub(r"-ADD\d*$", "", str(row.get("采购单号") or "").upper())
        total = Decimal(0)
        count = 0
        for other in details:
            if str(other.get("报关单号") or "") != str(row.get("报关单号") or ""):
                continue
            other_core = re.sub(r"-ADD\d*$", "", str(other.get("采购单号") or "").upper())
            if other_core == core:
                total += dec(other.get("采购金额_我方"))
                count += 1
        return total, count

    by_unit: dict[tuple, list[dict]] = {}
    for row in details:
        key = (row["报关单号"], row["供应商简称"], row["采购单核心"])
        by_unit.setdefault(key, []).append(row)

    verified = {}
    path = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
    if path.exists():
        payload = json.loads(path.read_text(encoding="utf-8"))
        verified = {
            str(case.get("报关单号") or "").strip(): case
            for case in payload.get("cases") or []
        }

    rows: list[list] = []
    fee_cache: dict[str, Decimal] = {}

    def fee_total(code: str) -> Decimal:
        """该采购单有效附件里费用行（包装费/木箱费…）的合计。"""
        if code in fee_cache:
            return fee_cache[code]
        total = Decimal(0)
        result = select(code)
        for item in result["kept"]:
            path_ = PARSED_ROOT / safe_name(code) / (
                safe_name(item["file"], "file") + ".json"
            )
            if not path_.exists():
                continue
            payload = json.loads(path_.read_text(encoding="utf-8"))
            by_column = payload.get("fee_by_column") or {}
            if by_column:
                total += sum((dec(v) for v in by_column.values()), Decimal(0))
            else:
                for fee in payload.get("extra_fees") or []:
                    if isinstance(fee, dict) and fee.get("amount"):
                        total += dec(fee.get("amount"))
        fee_cache[code] = total
        return total

    for unit in units:
        verdict = str(unit.get("结论") or "")
        if verdict.startswith("一致"):
            continue
        decl = str(unit["报关单号"])
        key = (decl, unit["供应商简称"], unit["采购单核心"])
        mine = dec(unit["采购金额_我方"])
        erp_r = ratio(mine, unit.get("ERP采购金额"))
        feishu_r = ratio(mine, unit.get("采购金额_飞书"))

        codes = [str(r["采购单号"]) for r in by_unit.get(key, [])]
        basis = "；".join(str(r.get("分摊依据") or "") for r in by_unit.get(key, []))
        files: list[str] = []
        for code in codes:
            result = select(code)
            for item in result["kept"]:
                files.append(f".cache/erp/attachments/grn/{code}/{item['file']}")

        if decl in verified:
            kind = "已核实"
        elif "未定价" in basis or not unit.get("采购金额_我方"):
            kind = "睿贝未定价"
        elif verdict == "飞书无对应行":
            kind = "飞书无行"
        elif (
            unit.get("ERP采购金额") not in (None, "")
            and fee_total(codes[0] if codes else "") > 0
            and abs((mine - dec(unit["ERP采购金额"])) - fee_total(codes[0] if codes else ""))
            <= Decimal("0.05")
        ):
            kind = "单据口径（实发含费用行）"
        elif (
            # 「单据冲突」要看**单元级**的睿贝出运金额：多单元的采购单里，
            # 拿「采购单级 ERP」比「单元金额」必然差一大截，那是分摊不是冲突。
            rmb_by_unit.get((decl, codes[0])) if codes else None
        ) and (
            abs(
                mine - (rmb_by_unit.get((decl, codes[0])) or Decimal(0))
            )
            / (rmb_by_unit.get((decl, codes[0])) or Decimal(1))
            > Decimal("0.10")
        ) and (feishu_r is not None and feishu_r > Decimal("0.10")):
            kind = "单据冲突"
        else:
            kind = "分摊口径"

        # ---- 自动归因（带证据）：让财务一眼看出差异是「口径」还是「数据问题」
        ref_money = dec(unit.get("采购金额_飞书"))
        fee = fee_total(codes[0]) if codes else Decimal(0)
        erp_unit = rmb_by_unit.get((decl, codes[0])) if codes else None
        tags: list[str] = []
        # ADD 并单最specific，先判：飞书一行 = 我方主单 + 各 ADD 附加单之和
        sib_total, sib_count = sibling_total(unit)
        if ref_money and sib_count > 1 and abs(sib_total - ref_money) <= Decimal("0.05"):
            tags.append(
                f"ADD 并单口径：飞书一行 = 我方主单+{sib_count - 1} 张附加单合计 {sib_total}"
            )
        if fee and abs(abs(mine - ref_money) - abs(fee)) <= Decimal("1"):
            tags.append(f"费用口径：我方含入库单费用行 {fee}，飞书只记货值")
        if erp_unit and abs(erp_unit - ref_money) <= Decimal("0.05"):
            tags.append("飞书=ERP出运采购金额（旧单价）")
        if erp_unit and abs(erp_unit - mine) <= Decimal("0.05"):
            tags.append("我方=ERP出运采购金额")
        if not tags and ref_money and codes:
            same_po, count = ref_sum_for(codes[0])
            if count > 1 and abs(same_po - mine) <= Decimal("0.05"):
                tags.append(f"采购单总额一致：飞书 {count} 行合计 = 我方金额（拆行/并单口径）")
        if not tags and ref_money and codes:
            # 财务口径：各单元取自己的 ERP 出运采购金额，差额（入库单总额 − ΣERP）挂在某一行
            pairs: dict[tuple[str, str], Decimal] = {}
            for other in details:
                if str(other.get("采购单号") or "") != codes[0]:
                    continue
                key = (str(other.get("报关单号") or ""), codes[0])
                pairs[key] = dec(other.get("采购金额_我方"))
            if len(pairs) > 1:
                po_total = sum(pairs.values(), Decimal(0))
                erp_sum = sum(
                    (rmb_by_unit.get(key, Decimal(0)) for key in pairs), Decimal(0)
                )
                if erp_sum:
                    residual = po_total - erp_sum
                    for key in pairs:
                        target = rmb_by_unit.get(key, Decimal(0)) + residual
                        value = ref_amount(key[0], codes[0])
                        others_ok = all(
                            abs(
                                rmb_by_unit.get(other, Decimal(0))
                                - (ref_amount(other[0], codes[0]) or Decimal(0))
                            )
                            <= Decimal("0.05")
                            for other in pairs
                            if other != key
                        )
                        if (
                            others_ok
                            and value is not None
                            and abs(target - value) <= Decimal("0.05")
                        ):
                            tags.append(
                                f"财务口径：各单元=各自ERP出运采购金额，差额 {residual} 挂在报关单 {key[0]}"
                            )
                            break
        if not tags:
            tags.append("待人工：需核单据或单价")

        po_totals: list[str] = []
        goods_units: list[Decimal] = []
        for code in codes:
            kept = select(code)["kept"]
            total = sum((dec(item["amount"]) for item in kept), Decimal(0))
            po_totals.append(str(total))
            if total:
                share = mine / total
                goods_units.append((total - fee_total(code)) * share)
        goods_value = goods_units[0] if goods_units else None
        evidence = (
            f"我方 {mine}｜飞书 {ref_money}｜ERP单元 "
            f"{erp_unit if erp_unit is not None else '—'}｜入库单费用行 {fee}"
            f"｜我方入库单总额 {'、'.join(po_totals) if po_totals else '—'}"
        )

        rows.append(
            [
                kind,
                decl,
                unit["供应商简称"],
                "、".join(codes),
                unit.get("采购金额_我方"),
                unit.get("采购金额_飞书"),
                unit.get("ERP采购金额"),
                "；".join(tags),
                evidence,
                str(goods_value.quantize(Decimal("0.01"))) if goods_value is not None else "",
                (
                    str((goods_value - ref_money).quantize(Decimal("0.01")))
                    if goods_value is not None and ref_money
                    else ""
                ),
                unit.get("偏差_飞书"),
                unit.get("偏差_ERP"),
                unit.get("结论"),
                basis,
                "；".join(files),
                verified.get(decl, {}).get("结论") if decl in verified else "",
            ]
        )

    rows.sort(key=lambda row: (row[0], str(row[1])))
    workbook = Workbook()
    workbook.remove(workbook.active)   # 两张表都自己建，去掉默认空表
    target = OUT_DIR / "差异台账.xlsx"

    # ---- 第二张表：采购单级对账（整单口径差 vs 纯拆行差）
    po_rows: dict[str, dict] = {}
    for row in details:
        code = str(row.get("采购单号") or "")
        if not code:
            continue
        entry = po_rows.setdefault(
            code,
            {
                "我方": Decimal(0),
                "飞书": Decimal(0),
                "飞书行": set(),
                "单位数": set(),
                "实发总额": dec(row.get("采购单实发总额")),
                "供应商": str(row.get("供应商简称") or ""),
            },
        )
        entry["我方"] += dec(row.get("采购金额_我方"))
        # 飞书金额是**行级**的（同一张报关单 × 同一个采购单可能对应我们多条产品类型记录），
        # 按「报关单 + 飞书金额」去重，既不漏行也不重复加
        unit_key = (
            str(row.get("报关单号") or ""),
            str(row.get("采购金额_飞书") or ""),
        )
        if unit_key not in entry["单位数"]:
            entry["单位数"].add(unit_key)
            entry["飞书"] += dec(row.get("采购金额_飞书"))
    summary_rows: list[list] = []
    for code, entry in sorted(po_rows.items()):
        ref_sum, ref_count = ref_sum_for(code, entry["供应商"])
        invoice_sum, invoice_count = ref_invoice_sum(code, entry["供应商"])
        erp_total = dec(erp_amount(code))
        mine = entry["我方"]
        diff = mine - entry["飞书"]
        if not entry["飞书"] and not ref_sum:
            note = "飞书无对应行"
        elif abs(diff) <= Decimal("0.05"):
            note = "整单一致"
        elif ref_sum and abs(ref_sum - mine) <= Decimal("0.05"):
            note = f"整单一致（飞书 {ref_count} 行合计 = 我方）"
        elif erp_total and abs(erp_total - mine) <= Decimal("0.05"):
            note = "我方 = 睿贝采购金额"
        elif erp_total and abs(erp_total - entry["飞书"]) <= Decimal("0.05"):
            note = "飞书 = 睿贝采购金额（旧价）"
        elif entry["实发总额"] and mine < entry["实发总额"] - Decimal("0.05"):
            note = (
                f"部分覆盖：我方只拿到该采购单的部分报关单，"
                f"未分配 {entry['实发总额'] - mine}"
            )
        elif ref_sum and ref_sum > entry["飞书"] + Decimal("0.05"):
            note = (
                f"飞书另有 {ref_sum - entry['飞书']} 落在我们没有的报关单上"
                f"（飞书全 {ref_count} 行 {ref_sum}）"
            )
        else:
            fee = fee_total(code)
            if fee and abs(abs(diff) - abs(fee)) <= Decimal("1"):
                note = f"费用口径：我方含费用行 {fee}"
            elif invoice_sum and abs(entry["实发总额"] - invoice_sum) > Decimal("0.05"):
                note = (
                    f"发票 vs 入库单：工厂开票 {invoice_sum}，入库单实发 {entry['实发总额']}"
                    f"（差 {entry['实发总额'] - invoice_sum}）"
                )
            else:
                note = (
                    f"整单不一致：我方=入库单实发 {entry['实发总额']}，"
                    f"飞书各行合计 {entry['飞书']}（差 {diff}），"
                    f"两边都不等于睿贝 {erp_total if erp_total is not None else '—'}"
                )
        summary_rows.append(
            [
                code,
                entry["供应商"],
                str(entry["实发总额"]),
                str(mine),
                str(entry["飞书"]),
                str(ref_sum) if ref_sum else "",
                ref_count,
                str(invoice_sum) if invoice_sum else "",
                str(erp_total) if erp_total is not None else "",
                str(diff) if entry["飞书"] else "",
                len(entry["单位数"]),
                note,
            ]
        )
    summary = workbook.create_sheet("采购单对账")
    summary.append(
        [
            "采购单号", "供应商", "入库单实发总额", "我方合计", "飞书合计(本单各行)",
            "飞书合计(含ADD)", "飞书行数", "工厂开票金额", "睿贝采购金额",
            "我方−飞书", "单元数", "结论",
        ]
    )
    for cell in summary[1]:
        cell.font = Font(bold=True)
    for row in summary_rows:
        summary.append(row)
    for index, name in enumerate(summary[1], 1):
        summary.column_dimensions[
            summary.cell(row=1, column=index).column_letter
        ].width = max(12, min(40, len(str(name)) * 2 + 6))
    summary.freeze_panes = "A2"

    # ---- 第一张表：单元级差异（多一列"采购单级结论"，把两级拉通）
    po_notes = {str(row[0]): str(row[-1]) for row in summary_rows}
    for row in rows:
        code = str(row[3]).split("、")[0]
        note = po_notes.get(code, "")
        row.append(note)
        # 「待人工」的行其实多数能在整单层面找到原因 —— 直接把那一层结论写进归因，
        # 并把原因分类从"单据冲突"这种粗标签改成可执行的类别。
        if (
            str(row[7]).startswith("待人工")
            and note
            and str(row[0]) in {"分摊口径", "单据冲突"}
        ):
            if note.startswith("整单一致"):
                row[7] = f"拆行差异：整单金额一致（{note}），差异来自怎么拆到各行"
                row[0] = "分摊口径"
            elif note.startswith(("部分覆盖", "飞书另有")):
                row[7] = f"缺报关单：{note}"
                row[0] = "缺报关单（待业务补）"
            elif note.startswith("发票 vs 入库单"):
                row[7] = f"发票与入库单不一致：{note}"
                row[0] = "发票与入库单不一致"
            elif note.startswith("我方 = 睿贝"):
                row[7] = "单价口径：我方=睿贝采购金额"
                row[0] = "单价口径"
            elif note.startswith("飞书 = 睿贝"):
                row[7] = "单价口径：飞书=睿贝采购金额（旧价）"
                row[0] = "单价口径"
        if str(row[7]).startswith("待人工"):
            kind = str(row[0])
            if kind == "已核实":
                conclusion = str(row[-1]) or str(row[16] or "")
                row[7] = f"已核实：{conclusion}" if conclusion else "已核实"
            elif kind == "睿贝未定价":
                row[7] = "睿贝该行未定价（采购金额留空，待人工补价）"
            elif kind == "飞书无行":
                row[7] = "飞书无对应行（无对照基准）"
    sheet = workbook.create_sheet("差异台账", 0)
    header = [
        "原因分类", "报关单号", "供应商简称", "采购单号", "我方采购金额", "飞书采购金额",
        "睿贝采购金额", "自动归因", "归因证据",
        "纯货值口径（我方）", "差额_纯货值_vs_飞书", "偏差_飞书", "偏差_ERP", "系统判定",
        "我方分摊依据", "附件路径", "已核实结论", "采购单级结论",
    ]
    sheet.append(header)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append(row)
        fill = FILLS.get(str(row[0]))
        if fill:
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    for index, name in enumerate(header, 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            12, min(64, len(name) * 2 + 6)
        )
    sheet.freeze_panes = "A2"

    workbook.save(target)
    counts: dict[str, int] = {}
    for row in rows:
        counts[str(row[0])] = counts.get(str(row[0]), 0) + 1
    print(f"written {target} ({len(rows)} rows) {counts}")


if __name__ == "__main__":
    main()
