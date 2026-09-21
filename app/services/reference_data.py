"""加载飞书副本表中的字典数据：产品类型映射、供应商名录、退税率、汇率。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

REF_DIR = Path(__file__).resolve().parents[2] / "data" / "reference"


def _flatten(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return " ".join(filter(None, (_flatten(v) for v in value)))
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return _flatten(value[key])
    return ""


def _load(name: str) -> list[dict[str, Any]]:
    path = REF_DIR / f"{name}.json"
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def _collect_options(node: Any, out: dict[str, str]) -> None:
    """递归收集 option id -> 名称。"""
    if isinstance(node, dict):
        if isinstance(node.get("id"), str) and isinstance(node.get("name"), str):
            out[node["id"]] = node["name"]
        for value in node.values():
            _collect_options(value, out)
    elif isinstance(node, list):
        for value in node:
            _collect_options(value, out)


class ReferenceData:
    """字典表缓存对象。"""

    def __init__(self) -> None:
        self.product_type_by_name: dict[str, str] = {}
        self.hs_by_name: dict[str, str] = {}
        self._load_products()
        self.supplier_by_code: dict[str, str] = {}
        self.supplier_short_by_name: dict[str, str] = {}
        self.supplier_option_by_id: dict[str, str] = {}
        self._load_suppliers()
        self._load_ai_fields()

    def _load_products(self) -> None:
        for record in _load("product_type"):
            fields = record.get("fields", {})
            name = _flatten(fields.get("海关品名"))
            kind = _flatten(fields.get("产品类型"))
            hs = _flatten(fields.get("对应海关编码"))
            if name:
                if kind:
                    self.product_type_by_name.setdefault(name, kind)
                if hs:
                    self.hs_by_name.setdefault(name, hs)

    def _load_suppliers(self) -> None:
        for record in _load("supplier"):
            fields = record.get("fields", {})
            name = _flatten(fields.get("供应商名称"))
            short = _flatten(fields.get("供应商简称"))
            code = _flatten(fields.get("文本"))
            if code and name:
                self.supplier_by_code[code] = name
            if name and short:
                self.supplier_short_by_name[name] = short

    def _load_ai_fields(self) -> None:
        """取「供应商简称」选项字典（option id -> 简称）。

        报关记录表里该字段是 Lookup，选项来自来源表，因此优先读报关参考表的字段定义。
        """
        for source in ("customs_ref_fields", "ai_fields"):
            for field in _load(source):
                if field.get("field_name") not in ("供应商简称", "供应商"):
                    continue
                mine: dict[str, str] = {}
                _collect_options(field.get("property"), mine)
                for key, value in mine.items():
                    self.supplier_option_by_id.setdefault(key, value)

    def product_type(self, declared_name: str) -> str:
        """按海关品名解析产品类型，未命中返回空串。"""
        text = (declared_name or "").strip()
        if not text:
            return ""
        if text in self.product_type_by_name:
            return self.product_type_by_name[text]
        for name, kind in self.product_type_by_name.items():
            if name and name in text:
                return kind
        return ""

    def supplier_name(self, code: str) -> str:
        return self.supplier_by_code.get((code or "").strip(), "")

    def supplier_short(self, option_id: str) -> str:
        """把报关记录里的供应商简称选项 ID 解析成简称文本。"""
        return self.supplier_option_by_id.get((option_id or "").strip(), "")

    def supplier_full_by_short(self, short: str) -> str:
        """简称 -> 全称。"""
        text = (short or "").strip()
        for name, value in self.supplier_short_by_name.items():
            if value == text:
                return name
        return ""
