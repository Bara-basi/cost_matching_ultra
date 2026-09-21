"""把副本表里的字典表（产品类型/供应商/退税率/汇率等）缓存到 data/reference/。

这些是拆分与分摊所需的静态参照数据，缓存后可离线复用。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

OUT_DIR = PROJECT_ROOT / "data" / "reference"

TABLES = {
    "product_type": "tblA34u43pAYSPMH",
    "supplier": "tblVRoMXPA5k7fhR",
    "tax_rate": "tblQh1D7L355Tjv2",
    "exchange_rate": "tblQOdhYE8VpfzCB",
    "product_map": "tblA34u43pAYSPMH",
    "customs_ref": "tbl2VGsoSAxhXRuv",
}


def fetch(app_token: str) -> None:
    client = FeishuClient()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for name, table_id in TABLES.items():
        if (OUT_DIR / f"{name}.json").exists():
            print(f"{name} 已存在，跳过")
            continue
        records = list(client.iter_records(app_token, table_id))
        (OUT_DIR / f"{name}.json").write_text(
            json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"{name}: {len(records)} 行")
    fields = client.list_fields(app_token, TABLES["product_map"])
    (OUT_DIR / "product_map_fields.json").write_text(
        json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    for name in ("customs_ref", "supplier", "product_type"):
        table_id = TABLES[name]
        (OUT_DIR / f"{name}_fields.json").write_text(
            json.dumps(client.list_fields(app_token, table_id), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


def main() -> None:
    app_token = get_config("MT_FINANCE_AI_TBALE_APP_TOKEN")
    fetch(app_token)
    print(f"输出目录 {OUT_DIR}")


if __name__ == "__main__":
    main()
