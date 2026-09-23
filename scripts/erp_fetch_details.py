"""按订单逐单抓取产品行明细（断点续跑、幂等）。

用法：
    python scripts/erp_fetch_details.py --year 2026            # 抓 2026 年订单
    python scripts/erp_fetch_details.py --year 2026 --limit 5  # 只抓 5 个，先看效果
    python scripts/erp_fetch_details.py --year 2026 --force    # 忽略已有明细
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT, read_jsonl, now_iso  # noqa: E402
from app.services.erp_cache_index import build_indexes  # noqa: E402
from app.services.erp_detail import load_detail, save_detail  # noqa: E402
from app.services.erp_detail import fetch_order_details  # noqa: E402
from app.services.erp_replay import ErpSession  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--year", default="2026")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--pause", type=float, default=0.05)
    args = parser.parse_args()

    orders = [r for r in read_jsonl(CACHE_ROOT / "sale_orders" / "orders.jsonl") if r.get("_year") == args.year]
    orders = [r for r in orders if r.get("orderId")]
    if args.limit:
        orders = orders[: args.limit]

    todo = []
    for row in orders:
        code = str(row.get("orderCode") or "").strip()
        if not args.force and load_detail("orders", code) is not None:
            continue
        todo.append(row)

    print(f"2026 订单 {len(orders)} 个，待抓明细 {len(todo)} 个", flush=True)
    started = now_iso()
    done = 0
    with ErpSession() as session:
        for index, row in enumerate(todo, 1):
            code = str(row.get("orderCode") or "").strip()
            try:
                payload = fetch_order_details(session, str(row["orderId"]), code)
            except Exception as exc:  # noqa: BLE001
                print(f"  [{index}/{len(todo)}] {code} ERROR {exc}", flush=True)
                continue
            payload["_orderId"] = row["orderId"]
            payload["_fetchedAt"] = now_iso()
            save_detail("orders", code, payload)
            done += 1
            if index % 25 == 0 or index == len(todo):
                print(
                    f"  [{index}/{len(todo)}] {code} "
                    f"出运行={len(payload['shipment_items'])} 采购行={len(payload['purchase_items'])}",
                    flush=True,
                )
            time.sleep(args.pause)

    stats = build_indexes()
    meta_path = CACHE_ROOT / "details_meta.json"
    meta = {
        "started_at": started,
        "finished_at": now_iso(),
        "year": args.year,
        "orders_total": len(orders),
        "orders_fetched": done,
        "index": stats,
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False))


if __name__ == "__main__":
    main()
