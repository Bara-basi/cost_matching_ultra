"""下载整张多维表格到 data/cache，供离线分析/服务启动预热。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

CACHE_DIR = PROJECT_ROOT / "data" / "cache"


def do_fetch(app_token: str, table_id: str, tag: str) -> Path:
    client = FeishuClient()
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    records = list(client.iter_records(app_token, table_id))
    target = CACHE_DIR / f"records_full_{tag}.json"
    target.write_text(
        json.dumps(records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return target


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tag", default="copy_full")
    parser.add_argument("--app-token", default=None)
    parser.add_argument("--table-id", default=None)
    parser.add_argument("--fields", action="store_true", help="同时导出字段定义")
    args = parser.parse_args()
    app_token = args.app_token or get_config("MT_FINANCE_AI_TBALE_APP_TOKEN")
    table_id = args.table_id or get_config("MT_FINANCE_AI_TABLE_ID")
    target = do_fetch(app_token, table_id, args.tag)
    if args.fields:
        client = FeishuClient()
        fields = client.list_fields(app_token, table_id)
        out = CACHE_DIR / f"fields_{args.tag}.json"
        out.write_text(json.dumps(fields, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"written {out}")
    print(f"written {target}")


if __name__ == "__main__":
    main()
