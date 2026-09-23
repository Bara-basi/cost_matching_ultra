"""探测「按单据 ID 查明细」的接口写法，确认能覆盖出运单与采购单。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache_index import load_index  # noqa: E402
from app.services.erp_replay import ErpSession, columns_to_rows  # noqa: E402

OUT = PROJECT_ROOT / ".cache" / "erp" / "reports"


def main() -> None:
    index = load_index()
    target = next(
        (v for v in index.values() if v.get("order") and v.get("shipments") and v.get("purchases")),
        None,
    )
    order_id = target["order"]["orderId"]
    shipment_id = target["shipments"][0]["shipmentId"]
    purchase_id = target["purchases"][0]["purchase_id"]
    out = [f"ids: order={order_id} shipment={shipment_id} purchase={purchase_id}"]

    cases = [
        ("ship_items_by_shipment", "/shipmentItem_list", {"shipment_id": shipment_id}),
        ("ship_items_by_shipment2", "/shipmentItem_selectShipItem", {"shipment_id": shipment_id}),
        ("ship_items_by_shipment3", "/shipmentItem_select", {"shipment_id": shipment_id}),
        ("ship_items_by_order", "/shipmentItem_selectShipFollow", {"order_id": order_id, "userDefaultTableName": "salefollow_shipment"}),
        ("pur_items_by_purchase", "/purchaseItems_list", {"purchase_id": purchase_id}),
        ("pur_items_by_purchase2", "/purchaseItems_selectPurchaseItem", {"purchase_id": purchase_id}),
        ("pur_items_by_order", "/purchaseItems_listPlaceOrder", {"order_id": order_id}),
        ("pur_list_by_purchase", "/purchase_selectPurDetail", {"purchase_id": purchase_id}),
    ]
    with ErpSession() as session:
        for label, path, query in cases:
            try:
                payload = session.fetch_json(path, {**query, "p": 1})
            except Exception as exc:  # noqa: BLE001
                out.append(f"[{label}] ERROR {exc}")
                continue
            rows = columns_to_rows(payload.get("root"))
            raw = payload.get("_raw")
            out.append(
                f"[{label}] path={path} keys={list(payload)[:4]} total={payload.get('total')} "
                f"rows={len(rows)} raw_head={(raw or '')[:120]}"
            )
            if rows:
                out.append("   sample=" + json.dumps(rows[0], ensure_ascii=False)[:400])
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "detail_probe2.txt").write_text("\n".join(out), encoding="utf-8")
    print("written .cache/erp/reports/detail_probe2.txt")


if __name__ == "__main__":
    main()
