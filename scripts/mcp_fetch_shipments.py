"""用 MCP shipment.find 抓取出运单的产品行明细，落盘到 .cache/erp/details/shipments/。

productList 每行含：销售订单号、采购订单号、供应商编码、出运数量、
出运采购金额(RMB)、海关商品（中文）、产品编码 —— 正是拆单需要的字段。
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT, ensure_dir, now_iso, read_jsonl  # noqa: E402
from scripts.mcp_erp import McpClient  # noqa: E402

DETAIL_DIR = CACHE_ROOT / "details" / "shipments"
SAFE = staticmethod(lambda text: "".join(c if c.isalnum() or c in "-_." else "_" for c in str(text)))


def path_for(code: str) -> Path:
    safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in str(code or "unknown"))
    return DETAIL_DIR / f"{safe}.json"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--token", default="")
    parser.add_argument("--year", default="", help="留空表示抓全部年份")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--pause", type=float, default=0.3)
    args = parser.parse_args()

    from app.services.feishu_client import get_config

    token = args.token or get_config("ERP_API_KEY")
    shipments = read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl")
    if args.year:
        shipments = [s for s in shipments if s.get("_year") == args.year]
    if args.limit:
        shipments = shipments[: args.limit]
    print(f"待抓出运单 {len(shipments)}", flush=True)

    client = McpClient(token)
    client.initialize()
    ensure_dir(DETAIL_DIR)
    done = skipped = errors = lines_total = 0
    for index, row in enumerate(shipments, 1):
        code = str(row.get("invoiceCode") or "").strip()
        if not code:
            continue
        target = path_for(code)
        if target.exists() and not args.force:
            skipped += 1
            continue
        try:
            text = client.call_text(
                "shipment.find",
                {
                    "invoiceCode": code,
                    "includeProduct": "true",
                    "includeExpense": "true",
                    "includePurchaseExpense": "true",
                    "includeForwarderExpense": "true",
                },
            )
        except Exception as exc:  # noqa: BLE001
            errors += 1
            print(f"  [{index}/{len(shipments)}] {code} ERROR {exc}", flush=True)
            continue
        try:
            payload = json.loads(text)
        except ValueError:
            errors += 1
            continue
        value = payload.get("value") or {}
        rows = value.get("productList") or []
        target.write_text(
            json.dumps(
                {
                    "invoiceCode": code,
                    "shipmentId": (value.get("shipmentBaseInfo") or {}).get("shipmentId"),
                    "baseInfo": value.get("shipmentBaseInfo") or {},
                    "productList": rows,
                    "expenseList": value.get("expenseList") or [],
                    "purchaseExpenseList": value.get("purchaseExpenseList") or [],
                    "fetchedAt": now_iso(),
                },
                ensure_ascii=False,
                indent=1,
            ),
            encoding="utf-8",
        )
        done += 1
        lines_total += len(rows)
        time.sleep(args.pause)
        if index % 25 == 0 or index == len(shipments):
            print(
                f"  [{index}/{len(shipments)}] 已抓={done} 跳过={skipped} 错误={errors} 产品行={lines_total}",
                flush=True,
            )
    print(
        json.dumps(
            {"done": done, "skipped": skipped, "errors": errors, "lines": lines_total},
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
