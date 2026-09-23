"""输出 Excel 的列定义与合同号推导（对齐飞书「2026年报关记录 AI」表）。

飞书上由公式/查找得到的字段这里照常保留成列；PDF 无法解析到的字段留空，
保证输出结构与多维表一致，便于后续回填与比对。
"""
from __future__ import annotations

import re

# 输出列顺序：优先「PDF 可解析 + 与成本匹配直接相关」的字段，其余照例保留
OUTPUT_COLUMNS: tuple[str, ...] = (
    # —— 报关单与合同 ——
    "来源文件",
    "单据类型",
    "合同号_1",
    "合同号（应收表格）",
    "合同号",
    "报关单号",
    "商品序号",
    "出口日期",
    "申报日期",
    # —— 商品明细 ——
    "报关品名",
    "海关编码",
    "报关重量",
    "报关重量单位",
    "申报数量",
    "申报单位",
    "目的国",
    "单价",
    "总价",
    "币种",
    # —— 表头其它 ——
    "运输方式",
    "成交方式",
    "运费",
    "保费",
    "杂费",
    "贸易方式",
    "出口口岸",
    "指运港",
    "境内货源地",
    "境内收发货人",
    "生产销售单位",
    "件数",
    "包装种类",
    "毛重",
    "净重",
    # —— 多维表中由查找/公式得到、这里保留空列 ——
    "产品类型",
    "退税率",
    "供应商简称",
    "供应商",
    "报关金额",
    "汇率",
    "出运日期",
    "客户代码（中信保）",
    "买方英文名",
    "国家",
    "业务员💦",
    "部门",
    "采购金额",
    "解析警告",
    "重复来源文件",
)

_SUFFIX_RE = re.compile(r"^[\s\-_]*(ADD\d*)\b", re.IGNORECASE)


def derive_contract(text: str) -> str:
    """按飞书公式从「合同号（应收表格）」推导「合同号」。

    - 长度 ≤ 11：原样返回；
    - 第 12 位是 `Y`：取前 12 位；
    - 第 12 位起是 ADD/ ADD/-ADD：原样返回；
    - 其它：取前 11 位。
    """
    if not text:
        return ""
    upper = text.upper()
    if len(upper) <= 11:
        return text
    if upper[11:12] == "Y":
        return text[:12]
    tail = upper[11:]
    if tail.startswith("ADD") or tail.startswith(" ADD") or tail.startswith("-ADD"):
        return text
    return text[:11]


def strip_batch_suffix(text: str) -> str:
    """去掉 ADD 尾缀，得到主体合同号。"""
    if not text:
        return ""
    return _SUFFIX_RE.sub("", text).strip()
