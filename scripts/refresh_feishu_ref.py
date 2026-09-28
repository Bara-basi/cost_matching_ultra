r"""从飞书「迈拓财务部门数据 副本」重新拉一遍对照表，重建本地缓存。

用法：
    & '.\.venv\Scripts\python.exe' scripts\refresh_feishu_ref.py
    & '.\.venv\Scripts\python.exe' scripts\refresh_feishu_ref.py --app-token <token> --table-id <id>

输出（`data/cache/`）：
    ref_2026_customs_full.json   全字段（成本匹配、待核表、证据工具都用它）
    ref_2026_customs.json        精简版（报关单号/合同/品名/供应商/采购金额…）
    records_full_copy_raw.json   原始抓取结果（排查用）

配置来自 `.env`（原表/副本各用独立键，别再共用）：
    副本 app token = `MT_FINANCE_DATA_TABLE_COPY_APP_TOKEN`
    副本表 id      = `MT_FINANCE_DATA_COPY_TABLE_ID`
    （原表是 `MT_FINANCE_DATA_TABLE_APP_TOKEN` / `MT_FINANCE_DATA_TABLE_ID`）
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

CACHE = PROJECT_ROOT / "data" / "cache"


def flatten(value: Any) -> Any:
    """多维表字段值 → 纯文本/数字（与既有缓存保持同一种口味）。"""
    if value is None or isinstance(value, (int, float, bool)):
        return value
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        parts = [flatten(item) for item in value]
        parts = [part for part in parts if part not in (None, "")]
        if not parts:
            return None
        if all(isinstance(part, str) for part in parts):
            return " | ".join(parts)
        return parts[0] if len(parts) == 1 else parts
    if isinstance(value, dict):
        for key in ("text", "name", "value", "en_name"):
            if key in value and value[key] not in (None, ""):
                return flatten(value[key])
        return None
    return str(value)


def main() -> None:
    parser = argparse.ArgumentParser(description="刷新飞书对照表缓存")
    parser.add_argument("--app-token", default=None)
    parser.add_argument("--table-id", default=None)
    args = parser.parse_args()
    app_token = args.app_token or get_config("MT_FINANCE_DATA_TABLE_COPY_APP_TOKEN")
    table_id = args.table_id or get_config("MT_FINANCE_DATA_COPY_TABLE_ID")
    print(f"拉取：app_token={app_token[:6]}… table_id={table_id}")

    client = FeishuClient()
    raw = list(client.iter_records(app_token, table_id))
    CACHE.mkdir(parents=True, exist_ok=True)
    (CACHE / "records_full_copy_raw.json").write_text(
        json.dumps(raw, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    fields: list[str] = []
    records: list[dict[str, Any]] = []
    for item in raw:
        row: dict[str, Any] = {}
        for name, value in (item.get("fields") or {}).items():
            if name not in fields:
                fields.append(name)
            row[name] = flatten(value)
        records.append(row)
    (CACHE / "ref_2026_customs_full.json").write_text(
        json.dumps({"fields": fields, "records": records}, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )

    trimmed_fields = (
        "报关单号", "合同号_1", "报关品名", "海关编码", "产品类型", "供应商简称",
        "供应商", "采购金额", "报关金额", "报关重量",
    )
    trimmed = []
    for row in records:
        entry = {name: row.get(name) for name in trimmed_fields if row.get(name) not in (None, "")}
        code = row.get("合同号（应收表格）") or row.get("合同号_1")
        if code:
            entry["采购单号"] = code
        trimmed.append(entry)
    (CACHE / "ref_2026_customs.json").write_text(
        json.dumps({"records": trimmed}, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    print(f"记录数：{len(records)}；字段数：{len(fields)}")
    print(f"written {CACHE / 'ref_2026_customs_full.json'}")
    print(f"written {CACHE / 'ref_2026_customs.json'}")
    filled = sum(1 for row in records if str(row.get("采购金额") or "").strip())
    print(f"其中「采购金额」有值的行：{filled}")


if __name__ == "__main__":
    main()
