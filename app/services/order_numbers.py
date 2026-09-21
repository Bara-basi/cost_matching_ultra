"""合同号 / 出运单号解析规则（依据 docs/合同号规范.md）。

合同号形态：`26MT-0KXXXX` + 可选 `Y`（本单有佣金）+ 可选 `ADD`/`ADD1`（附加单）。
出运单号可由多个采购单号用 `&` 连接，且遵循最小重复原则，例如：
  `26MT-01S180&190&241` -> 26MT-01S180 / 26MT-01S190 / 26MT-01S241
  `26MT-06H251&ADD1`    -> 26MT-06H251 / 26MT-06H251-ADD1
"""
from __future__ import annotations

import re
from typing import Iterable

CORE_RE = re.compile(r"^(\d{2})MT-(\d{2})([A-Z])(\d{3})", re.IGNORECASE)
EXCLUDED_PREFIX_RE = re.compile(r"^(DP|SP|ZY|CY|XM|SM)", re.IGNORECASE)
EXCLUDED_CONTAINS_RE = re.compile(r"(?:^|-)CY-?2[56]SM|2[56]SM", re.IGNORECASE)


def normalize(raw: object) -> str:
    """去掉空白与全角空格，统一大写。"""
    text = str(raw or "").strip().replace("\u3000", " ")
    text = re.sub(r"\s+", "", text)
    return text.upper()


def strip_pi(order: str) -> str:
    """去掉外销订单的 PI- 前缀。"""
    return re.sub(r"^PI-?", "", normalize(order))


def split_core(order: str) -> tuple[str, str]:
    """拆出 (主体基号, 尾缀)，尾缀含 Y / ADD 等信息。"""
    text = strip_pi(order)
    match = CORE_RE.match(text)
    if not match:
        return "", ""
    base = match.group(0).upper()
    return base, text[len(match.group(0)) :]


def is_excluded(contract: str) -> str:
    """返回跳过原因，空字符串表示需要处理。"""
    text = strip_pi(contract)
    if not text:
        return "合同号为空"
    if EXCLUDED_PREFIX_RE.match(text) or EXCLUDED_CONTAINS_RE.search(text):
        return "中间程序号（XM/CY/SM/DP/SP/ZY）"
    if text.startswith("24MT-"):
        return "历史24年订单"
    if not CORE_RE.match(text):
        return "不符合合同号规范"
    return ""


def expand_combined(raw: object) -> list[str]:
    """把出运单/联合合同号展开成单个订单号列表。

    支持最小重复原则与前缀继承：
      `26MT-01S180&190&241` -> [26MT-01S180, 26MT-01S190, 26MT-01S241]
      `26MT-06H251&ADD1`    -> [26MT-06H251, 26MT-06H251-ADD1]
    """
    text = normalize(raw)
    if not text:
        return []
    if "&" not in text:
        return [text]

    parts = [p for p in text.split("&") if p]
    out: list[str] = []
    previous = ""
    for index, part in enumerate(parts):
        if index == 0 or CORE_RE.match(part):
            previous = part
            out.append(part)
            continue
        base, _ = split_core(previous)
        prefix = base if base else previous
        candidate = f"{prefix}{part}"
        if re.fullmatch(r"\d+(?:[A-Z]\d*)?", part):
            prefix = re.sub(r"\d{3}$", "", prefix)
            candidate = f"{prefix}{part}"
        elif part.upper().startswith(("ADD", "Y", "R1")):
            candidate = f"{prefix}-{part}"
        previous = candidate
        out.append(candidate)
    return list(dict.fromkeys(out))


def purchase_candidates(order: str) -> list[str]:
    """采购单号候选：本体 + 常见 Y 后缀写法（用于容错匹配）。"""
    core = strip_pi(order)
    out = [core]
    if not core.endswith("Y") and re.search(r"\d$", core):
        out.append(f"{core}-Y")
    out.append(f"{core}Y")
    return list(dict.fromkeys(out))


def dedupe(values: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(v for v in values if v))
