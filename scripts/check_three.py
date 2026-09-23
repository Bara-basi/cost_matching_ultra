"""核查三条未拆出订单：出运单列表里能否通过合同号找到出运单。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT, read_jsonl  # noqa: E402
from app.services.shipment_index import canonical, strip_pi  # noqa: E402

TARGETS = ["25MT-06H655R1", "25MT-07F591", "26MT-03P039Y-B", "26MT-03P039Y"]


def main() -> None:
    shipments = read_jsonl(CACHE_ROOT / "shipments" / "shipments.jsonl")
    detail_dir = CACHE_ROOT / "details" / "shipments"
    out: list[str] = [f"出运单列表条数={len(shipments)}"]
    for target in TARGETS:
        key = strip_pi(target)
        canon = canonical(target)
        hits = []
        for row in shipments:
            codes = {strip_pi(row.get("invoiceCode")), strip_pi(row.get("orderCode"))}
            codes |= {strip_pi(c) for c in str(row.get("purchaseCode") or "").replace("&", ",").split(",")}
            if key in codes or canon in {canonical(c) for c in codes if c}:
                hits.append(row)
        out.append(f"--- {target}: 出运单列表命中 {len(hits)} ---")
        for row in hits[:6]:
            invoice = str(row.get("invoiceCode") or "")
            detail_exists = (detail_dir / f"{invoice.replace('/', '_')}.json").exists()
            out.append(
                f"    invoice={invoice} order={row.get('orderCode')} "
                f"purchase={str(row.get('purchaseCode'))[:60]} 明细已缓存={detail_exists}"
            )
    target_path = PROJECT_ROOT / ".cache" / "erp" / "reports" / "check_three.txt"
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text("\n".join(out), encoding="utf-8")
    print("\n".join(out))


if __name__ == "__main__":
    main()
