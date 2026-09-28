"""商品资料缓存（按产品编码）：`产品编码 -> 产品类别 / 海关商品(中文) / 中文名`。

数据来源 = 睿贝 MCP `product.find(productQueryType=productCode)`，
抓取脚本 `scripts/fetch_product_categories.py`，落盘
`.cache/erp/details/products/<safe(产品编码)>.json`（原始返回，便于人工核对）。

用途（2026-09-28 用户口径）：出运单界面缺 `海关商品（中文）` 时，
用「产品编码 → 商品资料 → 产品类别」补 `产品类型`；
`产品类别` 可能带部门后缀（`板棒(管道事业部)`），去掉括号里的内容即可；
产品资料也没有该字段时，用报关行 `报关品名` 兜底。
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from app.services.erp_cache import CACHE_ROOT

PRODUCT_DIR = CACHE_ROOT / "details" / "products"

# 类别名称 → 飞书那套粗分类的写法归一（只收实测出现过的差异写法）
CATEGORY_ALIASES: dict[str, str] = {
    "不锈钢盘管": "盘管",
    "镍合金盘管": "盘管",
    "双相钢盘管": "盘管",
}

_PAREN = re.compile(r"[（(][^）)]*[）)]")


def safe_code(code: str) -> str:
    return re.sub(r"[^0-9A-Za-z._-]", "_", str(code or "unknown"))


def product_path(code: str) -> Path:
    return PRODUCT_DIR / f"{safe_code(code)}.json"


def load_product(code: str) -> dict | None:
    """按产品编码读缓存的商品资料（原始 payload）；没有缓存返回 None。"""
    path = product_path(code)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None


def _base(payload: dict | None) -> dict:
    if not payload or not payload.get("success"):
        return {}
    return (payload.get("value") or {}).get("productBase") or {}


def category_of(code: str) -> str:
    """产品编码 → 归一后的产品类别（去部门括号）；取不到返回空串。"""
    return normalise_category(_base(load_product(code)).get("类别名称"))


def customs_name_of(code: str) -> str:
    """产品编码 → 商品资料里的 `海关商品(中文)`（取不到返回空串）。"""
    return str(_base(load_product(code)).get("海关商品(中文)") or "").strip()


def normalise_category(raw: object) -> str:
    """类别名称归一：去掉括号（部门）内容，已知差异写法合并到粗分类。

    `板棒(管道事业部)` → `板棒`；`不锈钢盘管` → `盘管`；
    整段就是部门名（`管道事业部`）或为空 → 返回空串（交由报关品名兜底）。
    """
    text = _PAREN.sub("", str(raw or "")).strip()
    if not text or "事业部" in text:
        return ""
    return CATEGORY_ALIASES.get(text, text)


def load_index() -> dict[str, dict[str, str]]:
    """产品编码 → {类别, 海关商品(中文), 中文名}（只含已缓存的商品资料）。"""
    out: dict[str, dict[str, str]] = {}
    if not PRODUCT_DIR.exists():
        return out
    for path in PRODUCT_DIR.glob("*.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        base = _base(payload)
        code = str(base.get("产品编码") or "").strip()
        if not code:
            continue
        out[code] = {
            "类别": normalise_category(base.get("类别名称")),
            "海关商品": str(base.get("海关商品(中文)") or "").strip(),
            "中文名": str(base.get("中文名") or "").strip(),
        }
    return out
