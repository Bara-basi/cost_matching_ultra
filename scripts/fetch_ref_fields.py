"""拉取「2026年报关数据」用到的目标表字段定义与记录，用于解码选项 ID。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient  # noqa: E402

APP_TOKEN = "KPvoblTTxakNtMsPSzMcIvfQnYg"  # 迈拓财务部门数据 副本
CACHE = PROJECT_ROOT / "data" / "cache"
TABLES = {
    "product_type": "tblnTbhye2EjcRUa",   # 产品及类型
    "supplier_ref": "tblclB7R2CHW9Yw4",   # 供应商（产品类型/供应商 lookup 的来源）
    "tax_rate": "tblfu5tTy3MeXAYv",
    "exchange": "tblLOwcrAkfk7IUJ",
}


def main() -> None:
    client = FeishuClient()
    for name, table_id in TABLES.items():
        try:
            fields = client.list_fields(APP_TOKEN, table_id)
        except Exception as exc:  # noqa: BLE001
            print(f"{name} 字段失败: {exc}")
            continue
        (CACHE / f"fields_{name}.json").write_text(
            json.dumps(fields, ensure_ascii=False, indent=1), encoding="utf-8"
        )
        try:
            records = list(client.iter_records(APP_TOKEN, table_id))
        except Exception as exc:  # noqa: BLE001
            print(f"{name} 记录失败: {exc}")
            records = []
        (CACHE / f"records_{name}.json").write_text(
            json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8"
        )
        print(f"{name}: 字段={len(fields)} 记录={len(records)}")


if __name__ == "__main__":
    main()
