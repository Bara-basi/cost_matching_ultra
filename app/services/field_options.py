"""读取飞书字段定义里的选项字典（option id -> 名称）。

多维表里单选/多选字段返回的是 `optXXXX` 这类选项 ID，
需要配合字段定义 `property.options` 才能还原成可读名称。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def _collect(node: Any, out: dict[str, str]) -> None:
    if isinstance(node, dict):
        if isinstance(node.get("id"), str) and isinstance(node.get("name"), str):
            out[node["id"]] = node["name"]
        for value in node.values():
            _collect(value, out)
    elif isinstance(node, list):
        for value in node:
            _collect(value, out)


def load_option_map(fields_path: Path, field_names: tuple[str, ...] | None = None) -> dict[str, str]:
    """从字段定义文件里收集选项字典。"""
    if not fields_path.exists():
        return {}
    out: dict[str, str] = {}
    fields = json.loads(fields_path.read_text(encoding="utf-8"))
    for field in fields:
        if field_names and field.get("field_name") not in field_names:
            continue
        _collect(field.get("property"), out)
    return out
