"""三单互查报告：**全部基于睿贝原生关联字段**。

关系来源：
- 订单 ↔ 出运：明细接口 `/shipmentItem_selectShipFollow` 返回的 `shipmentId`
- 订单 ↔ 采购：明细接口 `/purchaseItems_listPlaceOrder` 返回的 `purchaseId`
- 出运 ↔ 采购：出运单头原生存的 `purchaseCode`
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.erp_crossref import build  # noqa: E402

OUT = CACHE_ROOT / "reports"


def main() -> None:
    cross = build()
    stats = cross.stats()
    out: list[str] = []
    out.append("=== 原生关联统计 ===")
    out.append(json.dumps(stats, ensure_ascii=False))

    both = [
        oid
        for oid in cross.orders
        if cross.order_shipments.get(oid) and cross.order_purchases.get(oid)
    ]
    out.append(f"订单同时有出运与采购（原生）= {len(both)}")

    multi = [
        (sid, oids)
        for sid, oids in cross.shipment_orders.items()
        if len(oids) > 1
    ]
    out.append(f"包含多个订单的出运单（原生 orderId）= {len(multi)}")
    demo = sorted(multi, key=lambda x: -len(x[1]))[:5]
    for sid, oids in demo:
        row = cross.shipments.get(sid, {})
        codes = [cross.orders[o].get("orderCode") for o in oids if o in cross.orders]
        out.append(f"  出运单 {row.get('invoiceCode')} (id={sid}) 含订单 {codes}")

    out.append("")
    out.append("=== 出运单 → 采购单（原生 purchaseCode）===")
    ok = 0
    for sid in cross.shipments:
        if cross.purchases_of_shipment(sid):
            ok += 1
    out.append(f"能反查到采购单的出运单 = {ok}/{len(cross.shipments)}")

    out.append("")
    out.append("=== 采购单 → 出运单（经订单原生外键）===")
    ok = sum(1 for pid in cross.purchases if cross.shipments_of_purchase(pid))
    out.append(f"能反查到出运单的采购单 = {ok}/{len(cross.purchases)}")

    out.append("")
    out.append("=== 半包含订单（同一订单出现在多张出运单）===")
    counts = Counter(
        {"orderId": oid for oid, sids in cross.order_shipments.items() for _ in sids}.values()
    )
    partial = [(oid, len(sids)) for oid, sids in cross.order_shipments.items() if len(sids) > 1]
    out.append(f"跨多张出运单的订单（原生）= {len(partial)}")
    for oid, n in partial[:6]:
        order = cross.orders.get(oid, {})
        codes = [cross.shipments[s].get("invoiceCode") for s in cross.order_shipments[oid]]
        out.append(f"  {order.get('orderCode')}: {n} 张 -> {codes[:8]}")

    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "crossref_report.txt").write_text("\n".join(out), encoding="utf-8")
    print("\n".join(out[:8]))


if __name__ == "__main__":
    main()
