"""供应商「全称 / 简称」互转。

来源：
- `data/reference/supplier.json`（飞书供应商名录：供应商名称 ↔ 供应商简称）
- 出运产品行自带的供应商全称
"""
from __future__ import annotations

import json
import re
from pathlib import Path

REF = Path(__file__).resolve().parents[2] / "data" / "reference" / "supplier.json"
SUFFIX = re.compile(r"[-－]?\s*(供应链|总部|本部|一部|二部)\s*$")

# 飞书供应商名录里同一家公司可能有多条记录、简称还不一致，这时以这里的人工裁决为准。
# 例：`浙江国泰萧星密封材料股份有限公司` 名录里有「国泰」（文本 195）和「萧星」（文本 203）两条，
# 业务口径是同一家、统一写「萧星」（2026-09 用户确认：国泰就是萧星）。
NAME_TO_SHORT_OVERRIDES = {
    "浙江国泰萧星密封材料股份有限公司": "萧星",
}
SHORT_ALIASES = {
    "国泰": "萧星",
}


def normalize(name: str) -> str:
    """去掉 `-供应链` 这类内部后缀。"""
    text = str(name or "").strip()
    cleaned = SUFFIX.sub("", text).strip()
    return cleaned or text


class SupplierNames:
    """供应商名称与简称的双向映射。"""

    def __init__(self) -> None:
        self.name_to_short: dict[str, str] = {}
        self.short_to_name: dict[str, str] = {}
        self._load()

    def _load(self) -> None:
        if not REF.exists():
            return
        for record in json.loads(REF.read_text(encoding="utf-8")):
            fields = record.get("fields", {})
            name = normalize(str(fields.get("供应商名称") or ""))
            short = str(fields.get("供应商简称") or "").strip()
            if not name or not short:
                continue
            self.name_to_short.setdefault(name, short)
            self.short_to_name.setdefault(short, name)
        # 人工裁决覆盖名录里的重复/歧义记录
        for full, short in NAME_TO_SHORT_OVERRIDES.items():
            cleaned = normalize(full)
            self.name_to_short[cleaned] = short
            self.short_to_name.setdefault(short, cleaned)
        for alias, short in SHORT_ALIASES.items():
            full = self.short_to_name.get(short)
            if full:
                self.short_to_name.setdefault(alias, full)

    def short(self, name: str) -> str:
        """全称 → 简称；支持一条记录里塞了多个供应商（逗号/顿号分隔）。

        例：`浙江希蒙雷斯钢管有限公司,浙江信得达特种管业` → `希蒙、信得达`
        """
        text = str(name or "").strip()
        if not text:
            return ""
        parts = [p.strip() for p in re.split(r"[,，、;；/]", text) if p.strip()]
        if len(parts) > 1:
            shorts: list[str] = []
            for part in parts:
                value = self.short(part)
                if value and value not in shorts:
                    shorts.append(value)
            return "、".join(shorts)
        cleaned = normalize(text)
        if cleaned in SHORT_ALIASES:
            return SHORT_ALIASES[cleaned]
        if cleaned in self.name_to_short:
            return self.name_to_short[cleaned]
        if text in self.short_to_name:
            return text
        # 名称可能被截断（如少了「有限公司」）：用包含关系找唯一匹配
        candidates = [
            (full, short)
            for full, short in self.name_to_short.items()
            if cleaned and (cleaned in full or full in cleaned)
        ]
        if len(candidates) == 1:
            return candidates[0][1]
        if candidates:
            # 多个候选时取名称最长的（最具体的那个）
            best = max(candidates, key=lambda x: len(x[0]))
            return best[1]
        # 未收录：退化为去掉「有限公司/公司」后的名称
        return re.sub(r"(有限责任公司|有限公司|公司)$", "", cleaned) or cleaned

    def full(self, short: str) -> str:
        return self.short_to_name.get(str(short or "").strip(), "")

    def same_supplier(self, left: str, right: str) -> bool:
        """两个供应商名是否指同一家（忽略后缀、全称/简称差异）。"""
        a, b = normalize(left), normalize(right)
        if not a or not b:
            return False
        if a == b:
            return True
        if self.short(a) == self.short(b):
            return True
        return self.full(a) == b or self.full(b) == a
