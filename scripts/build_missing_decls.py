"""生成「待补报关单清单」：飞书里有钱、但我们没有 PDF 的报关单，按影响程度排序。

用途：财务/业务照着这张表去要原件，补齐后「部分覆盖 / 飞书另有行」这类差异会自然消掉。

排序依据：
  1. 高：这张报关单合同号涉及的订单**我们已经在拆**（缺了它，那条采购单的成本摊不全）
  2. 中：涉及的订单我们完全没有（目前不影响已有结果，但影响整体覆盖）
  3. 低：合同号是历史遗留/样品/代理（`SP-`/`CY-`/`ZY`/`24MT-`），按规则本来就拦截

输出：`outputs/cost_match/待补报关单清单.xlsx`
"""
from __future__ import annotations

import json
import sys
from decimal import Decimal
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from app.services.contract_shipments import cores_of  # noqa: E402

SPLIT = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
LEDGER = PROJECT_ROOT / "outputs" / "cost_match" / "差异台账.xlsx"
PARSE_SHEETS = (
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_出口退税联.xlsx",
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_预录单.xlsx",
)
REF = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs_full.json"
OUT = PROJECT_ROOT / "outputs" / "cost_match" / "待补报关单清单.xlsx"
LEGACY = ("SP-", "SP", "CY-", "ZY")


def read_rows(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    out = [dict(zip(header, row)) for row in rows[1:]]
    workbook.close()
    return out


def dec(value) -> Decimal:
    try:
        return Decimal(str(value).replace(",", "").strip())
    except Exception:  # noqa: BLE001
        return Decimal(0)


def our_declarations() -> set[str]:
    codes: set[str] = set()
    for path in (*PARSE_SHEETS, SPLIT):
        if not path.exists():
            continue
        codes |= {
            str(row.get("报关单号") or "").strip()
            for row in read_rows(path)
            if row.get("报关单号")
        }
    return codes


def main() -> None:
    ours = our_declarations()
    our_cores: set[str] = set()
    for row in read_rows(SPLIT):
        our_cores |= cores_of(str(row.get("合同号_1") or ""))

    # 「部分覆盖 / 飞书另有行」的采购单 = 已知被缺单影响的单，优先补它们的报关单
    affected_cores: set[str] = set()
    affected_pairs: set[tuple[str, str]] = set()
    if LEDGER.exists():
        workbook = load_workbook(LEDGER, read_only=True)
        if "采购单对账" in workbook.sheetnames:
            sheet = workbook["采购单对账"]
            rows = list(sheet.iter_rows(values_only=True))
            header = [str(cell) for cell in rows[0]]
            code_col = header.index("采购单号")
            note_col = header.index("结论")
            supplier_col = header.index("供应商")
            for row in rows[1:]:
                note = str(row[note_col] or "")
                if note.startswith(("部分覆盖", "飞书另有")):
                    affected_cores |= cores_of(str(row[code_col] or ""))
                    supplier = str(row[supplier_col] or "").strip()
                    for core in cores_of(str(row[code_col] or "")):
                        affected_pairs.add((core, supplier))
        workbook.close()

    records = json.loads(REF.read_text(encoding="utf-8"))["records"]
    grouped: dict[str, dict] = {}
    for record in records:
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
                "合同原文": set(),
            },
        )
        entry["合同号"].add(contract)
        entry["合同原文"].add(contract)
        entry["供应商"].add(str(record.get("供应商简称") or ""))
        entry["品名"].add(str(record.get("报关品名") or ""))
        entry["报关金额"] += dec(record.get("报关金额"))
        entry["采购金额"] += dec(record.get("采购金额"))
        entry["重量"] += dec(record.get("报关重量"))
        entry["核心"] |= cores_of(contract)

    rows: list[list] = []
    for decl, entry in grouped.items():
        hit = sorted(entry["核心"] & our_cores)
        critical = sorted(entry["核心"] & affected_cores)
        contract_text = "、".join(sorted(c for c in entry["合同原文"] if c))
        if critical:
            level = "最高：涉及已知被缺单影响的采购单"
        elif hit:
            level = "高：涉及我方在做的采购单"
        elif any(token in contract_text.upper() for token in LEGACY):
            level = "低：历史遗留/样品/代理（按规则拦截）"
        else:
            level = "中：涉及的订单我方完全没有"
        rows.append(
            [
                level,
                decl,
                contract_text,
                "、".join(sorted(s for s in entry["供应商"] if s)),
                "、".join(sorted(n for n in entry["品名"] if n)),
                str(entry["报关金额"].normalize()) if entry["报关金额"] else "",
                str(entry["采购金额"].normalize()) if entry["采购金额"] else "",
                str(entry["重量"].normalize()) if entry["重量"] else "",
                "、".join(critical or hit),
            ]
        )
    order = {
        "最高：涉及已知被缺单影响的采购单": 0,
        "高：涉及我方在做的采购单": 1,
        "中：涉及的订单我方完全没有": 2,
        "低：历史遗留/样品/代理（按规则拦截）": 3,
    }
    rows.sort(key=lambda row: (order.get(str(row[0]), 3), str(row[1])))

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "待补报关单"
    header = [
        "优先级", "报关单号", "合同号", "供应商简称", "报关品名",
        "报关金额(USD)", "飞书采购金额", "报关重量", "涉及的采购单核心",
    ]
    sheet.append(header)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append(row)
    fills = {
        "最高：涉及已知被缺单影响的采购单": PatternFill("solid", fgColor="FFC7CE"),
        "高：涉及我方在做的采购单": PatternFill("solid", fgColor="FCE4D6"),
        "中：涉及的订单我方完全没有": PatternFill("solid", fgColor="FFE699"),
        "低：历史遗留/样品/代理（按规则拦截）": PatternFill("solid", fgColor="EDEDED"),
    }
    for index, row in enumerate(rows, start=2):
        fill = fills.get(str(row[0]))
        if fill:
            for cell in sheet[index]:
                cell.fill = fill
    for index, name in enumerate(header, 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = max(
            12, min(40, len(name) * 2 + 6)
        )
    sheet.freeze_panes = "A2"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUT)

    counts: dict[str, int] = {}
    for row in rows:
        counts[str(row[0])] = counts.get(str(row[0]), 0) + 1
    print(f"written {OUT}（{len(rows)} 行）")
    for name, count in sorted(counts.items(), key=lambda kv: order.get(kv[0], 3)):
        print(f"   {count:>4}  {name}")


if __name__ == "__main__":
    main()
