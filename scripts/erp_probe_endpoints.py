"""试跑各列表接口与「按订单」明细接口，确认可达性与数据量。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_lists import (  # noqa: E402
    PURCHASE_BY_ORDER,
    PURCHASES,
    SALE_ORDERS,
    SHIPMENT_BY_ORDER,
    SHIPMENTS,
)
from app.services.erp_replay import ErpSession, columns_to_rows  # noqa: E402

OUT = PROJECT_ROOT / ".cache" / "_probe"


def try_list(session: ErpSession, spec, out: list[str]) -> None:
    payload = session.fetch_json(spec.path, {**spec.base_query, "p": 1})
    rows = columns_to_rows(payload.get("root"))
    out.append(
        f"[{spec.name}] path={spec.path} status_keys={list(payload)[:6]} "
        f"total={payload.get('total')} rows={len(rows)}"
    )
    if rows:
        out.append("   sample=" + json.dumps(rows[0], ensure_ascii=False)[:400])


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    out: list[str] = []
    with ErpSession() as session:
        for spec in (SALE_ORDERS, SHIPMENTS, PURCHASES):
            try:
                try_list(session, spec, out)
            except Exception as exc:  # noqa: BLE001
                out.append(f"[{spec.name}] ERROR {exc}")

        # 取一个真实订单，试「按订单」明细接口
        payload = session.fetch_json(SALE_ORDERS.path, {**SALE_ORDERS.base_query, "p": 1})
        rows = columns_to_rows(payload.get("root"))
        order = next((r for r in rows if r.get("orderId")), None)
        if order:
            out.append(f"order sample: code={order.get('orderCode')} id={order.get('orderId')}")
            for label, path, table in (
                ("shipment_by_order", SHIPMENT_BY_ORDER, "salefollow_shipment"),
                ("purchase_by_order", PURCHASE_BY_ORDER, "purchase_order_waiting_purchase"),
            ):
                try:
                    rows2 = session.fetch_json(path, {"order_id": order["orderId"], "userDefaultTableName": table, "p": 1})
                    parsed = columns_to_rows(rows2.get("root"))
                    out.append(
                        f"[{label}] status={rows2.get('status')} keys={list(rows2)[:6]} rows={len(parsed)}"
                    )
                    out.append("   raw=" + json.dumps(rows2, ensure_ascii=False)[:500])
                except Exception as exc:  # noqa: BLE001
                    out.append(f"[{label}] ERROR {exc}")
    (OUT / "endpoint_probe.txt").write_text("\n".join(out), encoding="utf-8")
    print(f"written {OUT / 'endpoint_probe.txt'}")


if __name__ == "__main__":
    main()
