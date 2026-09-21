"""一次性拉取引擎所需数据：报关记录（AI 表）、报关数据（参考）、字典表。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

CACHE = PROJECT_ROOT / "data" / "cache"
CACHE_REF = PROJECT_ROOT / "data" / "reference"
CUSTOMS_TABLE = "tbl2VGsoSAxhXRuv"  # 2026年报关数据（人工参考）


def save(name: str, records: list[dict]) -> None:
    CACHE.mkdir(parents=True, exist_ok=True)
    (CACHE / f"{name}.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"{name}: {len(records)} 行")


def save_ref(name: str, payload) -> None:
    CACHE_REF.mkdir(parents=True, exist_ok=True)
    (CACHE_REF / f"{name}.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"{name}: 已保存")


def main() -> None:
    client = FeishuClient()
    app = get_config("MT_FINANCE_AI_TBALE_APP_TOKEN")
    table = get_config("MT_FINANCE_AI_TABLE_ID")
    save("records_ai", list(client.iter_records(app, table)))
    save("t_customs2026", list(client.iter_records(app, CUSTOMS_TABLE)))
    save_ref("ai_fields", client.list_fields(app, table))


if __name__ == "__main__":
    main()
