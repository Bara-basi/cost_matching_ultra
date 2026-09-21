"""按表名下载任意多维表格到 data/cache/<tag>.json，并打印记录原文。"""
from __future__ import annotations

import argparse
import json
import ssl
import sys
import urllib.request
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

CACHE_DIR = PROJECT_ROOT / "data" / "cache"


def fetch(app_token: str, table_id: str, tag: str) -> list[dict]:
    client = FeishuClient()
    records = list(client.iter_records(app_token, table_id))
    (CACHE_DIR / f"t_{tag}.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return records


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-token", default=None)
    parser.add_argument("--table-id", required=True)
    parser.add_argument("--tag", required=True)
    parser.add_argument("--print", type=int, default=0, dest="print_n")
    parser.add_argument("--fields", default=None, help="逗号分隔，仅打印这些字段")
    args = parser.parse_args()

    app_token = args.app_token or get_config("MT_FINANCE_AI_TBALE_APP_TOKEN")
    records = fetch(app_token, args.table_id, args.tag)
    print(f"rows={len(records)}")
    keep = args.fields.split(",") if args.fields else None
    out: list[str] = []
    for record in records[: args.print_n]:
        out.append(f"##### {record['record_id']} #####")
        for key, value in record.get("fields", {}).items():
            if keep and key not in keep:
                continue
            out.append(f"{key} = {json.dumps(value, ensure_ascii=False)}")
    if out:
        (PROJECT_ROOT / "tmp" / f"dump_{args.tag}.txt").write_text("\n".join(out), encoding="utf-8")
        print(f"wrote tmp/dump_{args.tag}.txt")


if __name__ == "__main__":
    main()
