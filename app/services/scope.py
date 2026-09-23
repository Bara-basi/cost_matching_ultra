"""项目范围判定：拦截历史遗留订单，避免它们进入成本匹配链路。

不在范围内的合同号：
- `CY` / `SP` 开头（临时子公司、打样等中间程序号）
- `26SM` / `25SM` 等两位数字 + SM 开头（子公司 SM 系列）
- `ZY` 开头的中间程序号
- 24 年及更早年份（`2xMT-`，x ≤ 4）
- 这些号码在联合号里出现（如 `CY-26SM-S028`、`SP26MT-08C153`）同样拦截
"""
from __future__ import annotations

import re

# 形如 CY-26SM-S028 / SP-08C2605 / 26SM-J135 / 25SM-A134 / ZY20260508MT-XD
LEGACY_START_RE = re.compile(r"^(?:CY|SP|ZY|\d{2}SM)", re.IGNORECASE)
# 联合号（& 或 , 分隔）中任意一段命中
LEGACY_PART_RE = re.compile(r"(?:^|[&,，;；])\s*(?:CY|SP|ZY|\d{2}SM)", re.IGNORECASE)
# 24 年及更早的订单（25 年起在范围内：旧合同常有 26 年出运）
OLD_YEAR_RE = re.compile(r"^(?:1\d|2[0-4])MT-", re.IGNORECASE)


def out_of_scope_reason(contract: str) -> str:
    """返回拦截原因；在范围内返回空字符串。"""
    text = str(contract or "").strip()
    if not text:
        return ""
    if LEGACY_START_RE.match(text) or LEGACY_PART_RE.search(text):
        return "历史遗留合同号（CY/SM/SP/ZY）"
    if OLD_YEAR_RE.match(text):
        return "历史遗留年份（24 年及更早）"
    return ""


def in_scope(contract: str) -> bool:
    return not out_of_scope_reason(contract)
