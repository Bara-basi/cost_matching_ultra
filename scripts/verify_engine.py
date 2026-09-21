"""验证拆单/分摊引擎：用飞书报关记录跑一遍，与「报关数据」人工金额逐条对比。"""
from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.cost_engine import CostEngine, DeclaredRow  # noqa: E402
from app.services.erp_data import ErpDataStore  # noqa: E402
from app.services.feishu_client import FeishuClient, get_config  # noqa: E402
from app.services.reference_data import ReferenceData  # noqa: E402

CACHE = PROJECT_ROOT / "data" / "cache"


def flatten(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return " | ".join(filter(None, (flatten(v) for v in value)))
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return flatten(value[key])
    return ""


def first_option(value: Any) -> str:
    """取单选字段的 option id。"""
    if isinstance(value, list) and value:
        value = value[0]
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        return str(value.get("id") or value.get("text") or "").strip()
    return ""


def to_decimal(value: Any) -> Decimal:
    text = flatten(value)
    if not text:
        return Decimal(0)
    try:
        return Decimal(text.replace(",", ""))
    except Exception:  # noqa: BLE001
        return Decimal(0)


def load_records(offline: bool) -> list[dict[str, Any]]:
    path = CACHE / "records_ai.json"
    if offline and path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    client = FeishuClient()
    records = list(
        client.iter_records(
            get_config("MT_FINANCE_AI_TBALE_APP_TOKEN"), get_config("MT_FINANCE_AI_TABLE_ID")
        )
    )
    CACHE.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    return records


def load_reference_costs() -> dict[tuple[str, str, str], Decimal]:
    path = CACHE / "t_customs2026.json"
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    out: dict[tuple[str, str, str], Decimal] = {}
    for record in data:
        fields = record.get("fields", {})
        key = (
            flatten(fields.get("合同号")),
            flatten(fields.get("报关单号")),
            flatten(fields.get("报关品名")),
        )
        if all(key):
            out[key] = to_decimal(fields.get("采购金额"))
    return out


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--offline", action="store_true", help="只用本地缓存，不联网")
    parser.add_argument("--limit", type=int, default=0, help="只跑前 N 组（0=全部）")
    args = parser.parse_args()

    records = load_records(args.offline)
    reference = ReferenceData()

    rows: list[DeclaredRow] = []
    for record in records:
        fields = record.get("fields", {})
        product_name = flatten(fields.get("报关品名"))
        short = reference.supplier_short(first_option(fields.get("供应商简称")))
        supplier = reference.supplier_full_by_short(short) or short
        rows.append(
            DeclaredRow(
                record_id=record["record_id"],
                contract=flatten(fields.get("合同号")),
                declaration_no=flatten(fields.get("报关单号")),
                product_name=product_name,
                weight=to_decimal(fields.get("报关重量")),
                amount=to_decimal(fields.get("报关金额")),
                product_type=reference.product_type(product_name),
                supplier=supplier,
            )
        )

    engine = CostEngine(store=ErpDataStore(offline=args.offline), reference=reference)
    if args.limit:
        keep = {r.contract for r in rows if r.declaration_no}
        keep = set(list(keep)[: args.limit])
        rows = [r for r in rows if r.contract in keep]
    results = engine.allocate(rows)

    reference_costs = load_reference_costs()
    lines: list[str] = []
    stats = defaultdict(int)
    for row, result in zip(rows, results):
        key = (row.contract, row.declaration_no, row.product_name)
        expected = reference_costs.get(key)
        got = result.purchase_amount
        stats[result.status] += 1
        if expected is None:
            stats["无参考"] += 1
            continue
        diff = got - expected
        if abs(diff) <= Decimal("0.02"):
            stats["一致"] += 1
        elif got == 0:
            stats["系统无金额"] += 1
        else:
            stats["不一致"] += 1
        lines.append(
            f"{row.contract}\t{row.declaration_no}\t{row.product_name}\t"
            f"系统={got}\t人工={expected}\t差={diff}\t{result.status}\t{result.message}"
        )

    header = " === ".join(f"{k}={v}" for k, v in sorted(stats.items()))
    (PROJECT_ROOT / "data" / "processed").mkdir(parents=True, exist_ok=True)
    (PROJECT_ROOT / "data" / "processed" / "verify_engine.txt").write_text(
        header + "\n\n" + "\n".join(lines), encoding="utf-8"
    )
    print(header)


if __name__ == "__main__":
    main()
