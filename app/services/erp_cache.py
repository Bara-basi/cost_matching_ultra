"""睿贝 ERP 数据缓存：面向代码读取（JSONL + 索引），另附少量人类可读样例。

目录结构（.cache/erp/）：
    sale_orders/orders.jsonl          外销订单列表（每行一条 JSON）
    sale_orders/index.json            外销订单索引（orderCode -> 关键字段）
    sale_orders/samples/*.json        少量样例，便于人工查看
    shipments/shipments.jsonl         出运单列表
    shipments/index.json              出运单索引（invoiceCode/purchaseCode -> shipmentId）
    purchases/purchases.jsonl         采购订单列表
    purchases/index.json              采购订单索引（purchase_code -> purchase_id）
    links/order_graph.json            三类单据的关联图（orderCode ↔ purchaseCode ↔ invoiceCode）
    meta.json                         抓取时间、条数、参数
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

CACHE_ROOT = Path(__file__).resolve().parents[2] / ".cache" / "erp"


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_jsonl(path: Path, rows: Iterable[dict[str, Any]]) -> int:
    ensure_dir(path.parent)
    count = 0
    with path.open("w", encoding="utf-8") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")
            count += 1
    return count


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as handle:
        return [json.loads(line) for line in handle if line.strip()]


def write_json(path: Path, payload: Any) -> None:
    ensure_dir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")


def write_samples(rows: list[dict[str, Any]], out_dir: Path, limit: int = 3) -> None:
    """写少量人类可读样例（缩进 JSON），方便快速查看字段。"""
    ensure_dir(out_dir)
    for index, row in enumerate(rows[:limit], 1):
        key = row.get("orderCode") or row.get("invoiceCode") or row.get("purchase_code") or index
        safe = str(key).replace("/", "_").replace("\\", "_")
        write_json(out_dir / f"{index:02d}_{safe}.json", row)


def load_meta() -> dict[str, Any]:
    path = CACHE_ROOT / "meta.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def save_meta(meta: dict[str, Any]) -> None:
    write_json(CACHE_ROOT / "meta.json", meta)
