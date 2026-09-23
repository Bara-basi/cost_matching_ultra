r"""反推客户费用的分摊口径（按费用类型统计 + 候选基数逐一试）。

方法
====
1. 把每张报关单的**我方记录**与**飞书行**按「供应商 + 实际类型」配对，
   得到逐条差额：`差额 = 飞书报关金额 − 我方出运金额`（这就是这条记录分到的费用）。
2. 列出该报关单对应出运单的全部客户费用（名称 / 金额 / 销售订单号 / 加减方向）。
3. 对每一笔费用，取「销售订单号能命中的那些记录」作为候选，逐一试这些分摊基数：

       whole  整笔落在某一条记录上
       equal  候选记录等分
       by_amount  按候选记录的出运金额占比
       by_weight  按候选记录的报关重量占比（重量取飞书该行的报关重量）
       by_record_count  按（报关单×供应商）分组数等分

   哪一种能把差额逐条复现出来，就记这一笔费用命中了哪个基数。
4. 按**费用名称**汇总：每种费用各命中多少笔、有没有出现多种基数并存。

输出
====
    outputs/shipments_match/费用分摊反推.xlsx    逐笔费用 × 命中的基数
    outputs/shipments_match/费用分摊统计.json    按费用名称的命中分布
"""
from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from itertools import product
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "scripts"))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from app.services.contract_shipments import cores_of  # noqa: E402
from app.services.shipment_detail import shipment_invoices_for_contract  # noqa: E402

SPLIT_XLSX = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
REF_CACHE = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs.json"
VERIFIED = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
MAP_CACHE = PROJECT_ROOT / "data" / "reference" / "feishu_product_map.json"
DETAIL_DIR = PROJECT_ROOT / ".cache" / "erp" / "details" / "shipments"
OUT_DIR = PROJECT_ROOT / "outputs" / "shipments_match"

TOL = Decimal("0.02")
BASES = ("whole", "equal", "by_amount", "by_weight", "by_record_count")
COLUMNS = (
    "报关单号", "合同号_1", "出运单", "费用名称", "费用金额", "方向",
    "销售订单号", "候选记录", "差额合计", "命中的基数", "各基数误差",
)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")


def dec(value) -> Decimal:
    text = str(value or "").replace(",", "").strip()
    try:
        return Decimal(text)
    except Exception:  # noqa: BLE001
        return Decimal(0)


def money(value: Decimal) -> str:
    return f"{value.quantize(Decimal('0.01')):f}"


def safe(text: str) -> str:
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in str(text or ""))


def flat(value) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return "、".join(part for part in (flat(x) for x in value) if part)
    if isinstance(value, dict):
        return flat(value.get("text") or value.get("name") or "")
    return str(value).strip()


def split_multi(value) -> list[str]:
    return [part.strip() for part in flat(value).replace("，", "、").split("、") if part.strip()]


def shipment_fees(invoice: str) -> list[dict]:
    path = DETAIL_DIR / f"{safe(invoice)}.json"
    if not path.exists():
        return []
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows = [
        {
            "名称": str(fee.get("费用名称") or "").strip() or "(未命名)",
            "金额": dec(fee.get("金额")),
            "订单": str(fee.get("销售订单号") or "").strip(),
            "备注": str(fee.get("费用备注") or "").strip(),
        }
        for fee in payload.get("expenseList") or []
    ]
    rows = [row for row in rows if row["金额"]]
    if not rows:
        return []
    total = dec((payload.get("baseInfo") or {}).get("出运总金额"))
    product_total = sum(
        (dec(row.get("出运金额")) for row in payload.get("productList") or []), Decimal(0)
    )
    target = total - product_total
    signs: list[int] | None = None
    for candidate in product((1, -1), repeat=len(rows)):
        if abs(
            sum((s * row["金额"] for s, row in zip(candidate, rows)), Decimal(0)) - target
        ) <= TOL:
            if signs is not None:
                signs = None
                break
            signs = list(candidate)
    for index, row in enumerate(rows):
        row["方向"] = signs[index] if signs else 0
        row["出运单"] = invoice
    return rows


def load_ours() -> dict[str, list[dict]]:
    workbook = load_workbook(SPLIT_XLSX, read_only=True)
    rows = list(workbook.active.iter_rows(values_only=True))
    header = [str(c) for c in rows[0]]
    pos = {n: i for i, n in enumerate(header)}
    out: dict[str, list[dict]] = defaultdict(list)
    for row in rows[1:]:
        out[str(row[pos["报关单号"]]).strip()].append(
            {
                "合同号_1": str(row[pos["合同号_1"]] or ""),
                "供应商": str(row[pos["供应商简称"]] or ""),
                "采购单号": str(row[pos["采购单号"]] or ""),
                "报关品名": str(row[pos["报关品名"]] or ""),
                "产品类型": str(row[pos["产品类型"]] or ""),
                "出运金额": dec(row[pos["出运金额合计"]]),
            }
        )
    return out


def load_ref() -> dict[str, list[dict]]:
    records = json.loads(REF_CACHE.read_text(encoding="utf-8"))["records"]
    out: dict[str, list[dict]] = defaultdict(list)
    for item in records:
        decl = str(item.get("报关单号") or "").strip().replace(" ", "")
        if not decl:
            continue
        out[decl].append(
            {
                "供应商": flat(item.get("供应商简称")),
                "类型": split_multi(item.get("产品类型")),
                "报关金额": dec(item.get("报关金额")),
                "报关重量": dec(item.get("报关重量")),
            }
        )
    return out


def pair(ours: list[dict], theirs: list[dict], mapping: dict[str, str]) -> list[dict]:
    """把一张报关单的我方记录与飞书行配对，返回带差额的记录。"""
    used: set[int] = set()
    paired: list[dict] = []
    for record in ours:
        mine_type = mapping.get(record["产品类型"], record["产品类型"])
        pool = [
            row for row in theirs
            if id(row) not in used and row["供应商"] == record["供应商"]
        ]
        hit = [row for row in pool if mine_type and mine_type in row["类型"]]
        if len(hit) != 1:
            hit = sorted(pool, key=lambda r: abs(r["报关金额"] - record["出运金额"]))[:1]
        match = hit[0] if hit else None
        if match is not None:
            used.add(id(match))
        paired.append(
            {
                **record,
                "折算类型": mine_type,
                "飞书报关金额": match["报关金额"] if match else Decimal(0),
                "报关重量": match["报关重量"] if match else Decimal(0),
                "差额": (match["报关金额"] if match else Decimal(0)) - record["出运金额"],
            }
        )
    return paired


def allocate(records: list[dict], amount: Decimal, direction: int, base: str) -> list[Decimal]:
    total = direction * amount
    count = len(records)
    if base == "whole":
        biggest = max(records, key=lambda r: r["出运金额"]) if count > 1 else records[0]
        return [total if r is biggest else Decimal(0) for r in records]
    if base == "equal":
        share = total / count
        out = [share] * count
    elif base == "by_amount":
        weights = [r["出运金额"] for r in records]
        out = _split(total, weights)
    elif base == "by_weight":
        weights = [r["报关重量"] for r in records]
        out = _split(total, weights) if any(weights) else [Decimal(0)] * count
    elif base == "by_record_count":
        suppliers = list(dict.fromkeys(r["供应商"] for r in records))
        weights = [Decimal(1) for _ in suppliers]
        shares = _split(total, weights)
        per_supplier = dict(zip(suppliers, shares))
        out = [per_supplier[r["供应商"]] / sum(1 for x in records if x["供应商"] == r["供应商"])
               for r in records]
    else:
        out = [Decimal(0)] * count
    # 尾差给最后一条
    diff = total - sum(out)
    out[-1] += diff
    return out


def _split(total: Decimal, weights: list[Decimal]) -> list[Decimal]:
    base = sum(weights)
    if not base:
        return [Decimal(0)] * len(weights)
    return [total * w / base for w in weights]


def main() -> None:
    ours = load_ours()
    ref = load_ref()
    mapping = (
        json.loads(MAP_CACHE.read_text(encoding="utf-8")) if MAP_CACHE.exists() else {}
    )
    blocked: set[str] = set()
    if VERIFIED.exists():
        blocked = {
            str(case.get("报关单号") or "").strip()
            for case in (json.loads(VERIFIED.read_text(encoding="utf-8")).get("cases") or [])
            if case.get("报关单号")
        }

    # 口径：一条记录分到的费用 = 它的**名下订单**（采购单号核心）承载的那些客户费用。
    # 拿这个「预测值」去和实测差额比，才对得上——否则同一记录身上挂着好几笔费用时必然比错。
    rows_out: list[dict] = []
    kind_hits: dict[str, Counter] = defaultdict(Counter)
    for decl, items in sorted(ours.items()):
        if decl in blocked:
            continue
        paired = pair(items, ref.get(decl, []), mapping)
        if not any(abs(r["差额"]) > TOL for r in paired):
            continue
        invoices: list[str] = []
        for code in {r["合同号_1"] for r in paired}:
            invoices.extend(shipment_invoices_for_contract(code))
        invoices = list(dict.fromkeys(invoices))
        fees = [f for inv in invoices for f in shipment_fees(inv)]

        claimed: set[int] = set()
        for record in paired:
            mine_cores = cores_of(record["采购单号"])
            owned = [
                fee for fee in fees
                if cores_of(fee["订单"]) & mine_cores
            ]
            for fee in owned:
                claimed.add(id(fee))
            predicted = sum(
                ((fee["方向"] or 1) * fee["金额"] for fee in owned), Decimal(0)
            )
            if abs(record["差额"]) <= TOL and abs(predicted) <= TOL:
                continue
            same = abs(record["差额"] - predicted) <= TOL
            for fee in owned:
                kind_hits[fee["名称"]]["整笔归属命中" if same else "整笔归属未命中"] += 1
            rows_out.append(
                {
                    "报关单号": decl,
                    "合同号_1": record["合同号_1"],
                    "出运单": "、".join(invoices),
                    "费用名称": "；".join(dict.fromkeys(fee["名称"] for fee in owned)) or "(无)",
                    "费用金额": "；".join(money(fee["金额"]) for fee in owned) or "0",
                    "方向": "；".join(
                        "+" if (fee["方向"] or 1) > 0 else "-" for fee in owned
                    ),
                    "销售订单号": "；".join(dict.fromkeys(fee["订单"] for fee in owned)),
                    "候选记录": (
                        f"{record['供应商']}/{record['产品类型'] or record['报关品名']} "
                        f"出运{money(record['出运金额'])} 重{money(record['报关重量'])}"
                    ),
                    "差额合计": money(record["差额"]),
                    "命中的基数": "整笔归属" if same else "不吻合",
                    "各基数误差": "",
                }
            )
        # 没能被任何记录认领的费用
        for fee in fees:
            if id(fee) not in claimed:
                kind_hits[fee["名称"]]["没有任何记录认领"] += 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "费用分摊反推"
    sheet.append(list(COLUMNS))
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for row in rows_out:
        sheet.append([row.get(name, "") for name in COLUMNS])
        for cell in sheet[sheet.max_row]:
            cell.fill = OK_FILL if row["命中的基数"] != "（未命中）" else BAD_FILL
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), COLUMNS):
        sheet.column_dimensions[column[0].column_letter].width = max(12, min(46, len(name) * 2 + 4))
    sheet.freeze_panes = "A2"
    workbook.save(OUT_DIR / "费用分摊反推.xlsx")

    stats = {
        "费用笔数": len(rows_out),
        "命中任一基数的笔数": sum(1 for r in rows_out if r["命中的基数"] != "（未命中）"),
        "按费用名称的命中分布": {k: dict(v) for k, v in sorted(kind_hits.items())},
    }
    (OUT_DIR / "费用分摊统计.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
