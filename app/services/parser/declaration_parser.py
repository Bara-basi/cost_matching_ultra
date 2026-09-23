"""中华人民共和国海关出口货物报关单（PDF）解析器。

版式：`JG09` 打印的「出口退税联」，单页，表头字段 + 商品明细表。
pypdf 默认抽取模式在该版式下呈「一行一个字段」形态，因此按行取值最稳：
找到标签行后，取其后第一个「不是标签」的行作为取值。
"""
from __future__ import annotations

import re
from pathlib import Path

from app.services.parser.models import DeclarationHeader, DeclarationItem, ParsedDeclaration
from app.services.parser.pdf_text import extract_text

# 表头标签集合：取值时用于跳过中间夹着的其它标签
LABELS: tuple[str, ...] = (
    "预录入编号",
    "海关编号",
    "出口口岸",
    "备案号",
    "出口日期",
    "申报日期",
    "境内收发货人",
    "运输方式",
    "运输工具名称",
    "提运单号",
    "生产销售单位",
    "贸易方式",
    "征免性质",
    "结汇方式",
    "许可证号",
    "运抵国",
    "指运港",
    "境内货源地",
    "批准文号",
    "成交方式",
    "运费",
    "保费",
    "杂费",
    "合同协议号",
    "件数",
    "包装种类",
    "毛重（千克）",
    "净重（千克）",
    "集装箱号",
    "随附单据",
    "标记唛码及备注",
    "商品序号 商品编号 商品名称、规格型号",
    "法定数量/法定单位",
    "第二数量/第二单位",
    "申报数量/申报单位",
    "税费征收情况",
    "录入员",
    "录入单位",
    "海关审单批注及放行日期（签章）",
    "审单",
    "审价",
    "报关员",
    "申报单位（签章）",
    "征税",
    "统计",
    "单位地址",
    "查验",
    "放行",
    "邮编",
    "电话",
    "填制日期",
)

# 合同号可能因过长被截断，或整条被写进备注行
CONTRACT_TOKEN_RE = re.compile(r"[0-9]{2}[A-Z]{2}-[0-9]{2}[A-Z][0-9]{3}[A-Za-z0-9&.\-]*")
# 单个订单令牌（可含被截断的流水号）
ORDER_TOKEN_RE = re.compile(r"\d{2}[A-Z]\d{3}[A-Za-z0-9.\-]*")
ORDER_TOKEN_TAIL_RE = re.compile(r"\d{3}[A-Z]\d?(?:-[A-Za-z0-9]+|[A-Za-z0-9]*)?")
CONTRACT_REMARK_RE = re.compile(r"合同协议号\s*[:：]?\s*(.+)")

ITEM_RE = re.compile(r"^(\d{1,3}) (\d{8,10}) (.+?)\s+法定数量/法定单位$")
ITEM_V2_RE = re.compile(r"^(\d{1,3})\s+(\d{8,10})(\S.*)$")
QTY_RE = re.compile(r"^([\d,.]+\S*)\s+(\S+)\s+([\d.]+)\s+([\d.]+)\s+([A-Z]{3})\s*$")
# 未带代号的表头标签（对应「就近对齐」版式）
LABELS_V2: tuple[str, ...] = (
    "备案号",
    "申报日期出口日期",
    "出境关别",
    "境内发货人",
    "提运单号",
    "运输工具名称及航次号",
    "运输方式",
    "境外收货人",
    "许可证号",
    "征免性质",
    "监管方式",
    "生产销售单位",
    "离境口岸",
    "指运港",
    "运抵国",
    "贸易国",
    "合同协议号",
    "杂费保费",
    "运费",
    "成交方式",
    "净重(千克)",
    "毛重(千克)",
    "件数",
    "包装种类",
    "随附单证及编号",
)


def _is_label(line: str) -> bool:
    return any(line.startswith(label) for label in LABELS)


def _value_after(lines: list[str], label: str, occurrence: int = 0, max_scan: int = 4) -> str:
    """取标签后第一个非标签行。"""
    hits = [i for i, line in enumerate(lines) if line.startswith(label)]
    if len(hits) <= occurrence:
        return ""
    index = hits[occurrence]
    # 同一行标签后直接跟值（少数版式）
    inline = lines[index][len(label) :].strip(" :：")
    if inline:
        return inline
    for offset in range(1, max_scan + 1):
        if index + offset >= len(lines):
            break
        candidate = lines[index + offset].strip()
        if not candidate or _is_label(candidate):
            continue
        return candidate
    return ""


def _contract_after(lines: list[str]) -> str:
    """取「合同协议号」下方第一个看起来像合同号的行（可能被截断）。"""
    index = next((i for i, line in enumerate(lines) if line.startswith("合同协议号")), None)
    if index is None:
        return ""
    for offset in range(1, 4):
        if index + offset >= len(lines):
            break
        candidate = lines[index + offset].strip()
        if not candidate or _is_label(candidate):
            continue
        if CONTRACT_TOKEN_RE.match(candidate.strip(":：")):
            return candidate
    return _value_after(lines, "合同协议号")


def _clean_date(value: str) -> str:
    """yyyyMMdd / yyyy-MM-dd -> yyyy-MM-dd。"""
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    if len(digits) == 6:
        return f"{digits[:4]}-{digits[4:6]}"
    return (value or "").strip()


def recover_contract(header: DeclarationHeader, lines: list[str]) -> DeclarationHeader:
    """合同号被截断或写进备注行时的兼容处理。

    两种已知情形（实测 SMPdbRklEoyWGwxtzcacz6m9nUb / NOatb3G8koBTuNx8nRvco3Pcnug）：
    ① 表头只放了前半段（以 `&` 结尾），剩下半段进了备注行；
    ② 表头只放了前半段（`&` 后接半个令牌），后半段掉到了下一行，备注行另有完整值。
    处理顺序：备注行的完整值优先级最高；否则把下一行续接到表头值后面。
    """
    remark = ""
    for line in lines:
        match = CONTRACT_REMARK_RE.search(line)
        if match:
            candidate = match.group(1).strip()
            # 备注里可能只写「剩下半段」（如 `262Y-A`），因此放宽校验
            if CONTRACT_TOKEN_RE.search(candidate) or _order_tokens(candidate):
                remark = candidate
                break
    header.contract_remark = remark

    if remark:
        header.contract_raw = _merge_contract(header.contract_raw, remark)
        return header

    raw = header.contract_raw.strip()
    if not raw:
        return header
    # 表头行位置：取标签后的那一行
    index = next(
        (i for i, line in enumerate(lines) if line.startswith("合同协议号")), None
    )
    if index is None:
        return header
    next_line = lines[index + 1].strip() if index + 1 < len(lines) else ""
    incomplete = raw.endswith("&") or not re.search(r"[A-Za-z]\d*$", raw)
    if incomplete and next_line and not _is_label(next_line):
        merged = raw + next_line
        if CONTRACT_TOKEN_RE.search(merged):
            header.contract_raw = _merge_contract(raw, next_line)
    return header


def _segment_tokens(text: str) -> tuple[str, list[str]]:
    """拆出 (公司前缀, 订单令牌列表)。

    订单段支持三种写法（最小重复原则）：完整基号 `26MT-06N100`、
    只写流水 `&110`、只写批次 `&087A`/`&262Y-A`。后两种需要沿用前一段的前缀。
    """
    parts = [p.strip() for p in (text or "").split("&") if p.strip()]
    if not parts:
        return "", []
    prefix = ""
    head = re.fullmatch(r"(\d{2}[A-Z]{2}-)(.+)", parts[0])
    if head:
        prefix = head.group(1)
        parts[0] = head.group(2)
    tokens: list[str] = []
    for part in parts:
        if re.fullmatch(r"\d{2}[A-Z]{2}-.+", part):
            tokens.append(part)
            continue
        if ORDER_TOKEN_RE.fullmatch(part) or ORDER_TOKEN_TAIL_RE.fullmatch(part):
            tokens.append(part)
            continue
        if re.fullmatch(r"\d{1,3}", part):
            tokens.append(part)
            continue
        found = ORDER_TOKEN_RE.findall(part)
        if found:
            tokens.extend(x for x in found if x not in tokens)
        elif part:
            tokens.append(part)
    return prefix, tokens


def _order_tokens(text: str) -> list[str]:
    """只取订单令牌（去掉公司前缀）。"""
    return _segment_tokens(text)[1]


def _merge_contract(main: str, remark: str) -> str:
    """把表头片段与备注片段合并成规范联合号。

    实测两种情形：
    ① 表头 `26MT-03P200Y-A&229Y-A&245Y-A&`，备注 `262Y-A`；
    ② 表头 `26MT-03P200Y-A&229Y-A&245Y-A&`，备注含完整联合号。
    输出统一为 `26MT-03P200Y-A&229Y-A&245Y-A&262Y-A`。
    """
    prefix, main_tokens = _segment_tokens(main)
    if not prefix:
        prefix = _segment_tokens(remark)[0]
    remark_tokens = _order_tokens(remark)

    # 表头在 `&` 处被截断时，优先采纳备注（备注通常更完整）；
    # 否则以表头为准，仅把备注里新增的令牌补在后面。
    main_incomplete = main.rstrip().endswith("&")
    if main_incomplete and len(remark_tokens) >= len(main_tokens):
        tokens = list(remark_tokens)
        if not prefix:
            prefix = _segment_tokens(remark)[0]
    else:
        merged_tokens = list(main_tokens)
        for token in remark_tokens:
            if token not in merged_tokens:
                merged_tokens.append(token)
        tokens = merged_tokens

    tokens = _dedupe_tokens(tokens, prefix)

    if not tokens:
        return (remark or main or "").strip()

    # 续写出完整令牌：`03P` 前缀 + 流水/批次
    full_tokens: list[str] = []
    last_core = ""
    for token in tokens:
        core = re.match(r"^(\d{2}[A-Z]\d)\d{3}", token)
        if core:
            last_core = core.group(1)
            full_tokens.append(token)
            continue
        if re.fullmatch(r"\d{1,3}", token) and last_core:
            full_tokens.append(f"{last_core}{int(token):03d}")
            continue
        if last_core:
            full_tokens.append(f"{last_core}{token.lstrip('-')}")
        else:
            full_tokens.append(token)
    return f"{prefix}{'&'.join(full_tokens)}"


def _dedupe_tokens(tokens: list[str], prefix: str = "") -> list[str]:
    """同一订单只保留最完整的一版（`245Y-` 与 `245Y-A` 视为同一订单）。"""
    best: dict[str, str] = {}
    order: list[str] = []
    for token in tokens:
        match = re.search(r"\d{3}[A-Z]?", token)
        key = match.group(0) if match else token
        if key not in best:
            best[key] = token
            order.append(key)
            continue
        if len(token) > len(best[key]):
            best[key] = token
    return [best[key] for key in order]


def rebuild_contract(text: str) -> str:
    """兼容旧调用：把联合号按令牌规范化。"""
    tokens = _order_tokens(text)
    if not tokens:
        return ""
    base_match = re.match(r"^(\d{2}[A-Z]{2}-)", (text or "").strip())
    prefix = base_match.group(1) if base_match else ""
    return f"{prefix}{'&'.join(tokens)}"


def _strip_code(value: str) -> str:
    """去掉字段值里的括号代码，如 `(33039)`。"""
    return re.sub(r"[（(]\s*\d+\s*[)）]", "", value or "").strip()


def _find_line(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1).strip() if match else ""


def _parse_header(lines: list[str], text: str) -> DeclarationHeader:
    header = DeclarationHeader()
    header.declaration_no = _find_line(text, r"海关编号[:：]\s*(\d{12,20})")
    header.pre_entry_no = _find_line(text, r"预录入编号[:：]\s*([A-Za-z0-9]{8,24})")
    header.contract_raw = _contract_after(lines)

    header.export_date = _clean_date(_value_after(lines, "出口日期"))
    header.declare_date = _clean_date(_value_after(lines, "申报日期"))
    header.transport_mode = _value_after(lines, "运输方式")
    header.trade_mode = _strip_code(_trade_mode(lines))
    header.deal_mode = _value_after(lines, "成交方式")
    # 报关单上的运费/保费/杂费：取值紧跟标签行（如 `USD/2800/总价`、`/0/`）
    header.freight = _value_after(lines, "运费")
    header.insurance = _value_after(lines, "保费")
    header.misc_fee = _value_after(lines, "杂费")
    header.export_port = _after_code_line(lines, "出口口岸")
    header.destination_country = _after_code_line(lines, "运抵国")
    header.destination_port = _after_code_line(lines, "指运港")
    # 版式 1 该标签自带代码：`境内货源地(33039)`，取值在下一行
    header.domestic_source = _strip_code(_after_code_line(lines, "境内货源地"))
    header.consignor = _after_code_line(lines, "境内收发货人")
    header.producer = _value_after(lines, "生产销售单位")
    header.pieces = _value_after(lines, "件数")
    header.package_kind = _value_after(lines, "包装种类")
    header.gross_weight = _value_after(lines, "毛重（千克）")
    header.net_weight = _value_after(lines, "净重（千克）")
    return header


def _after_code_line(lines: list[str], label: str) -> str:
    """标签行本身只带机构代码，取值在**再下一行**（如 运抵国（地区）(501) -> 加拿大）。"""
    hits = [i for i, line in enumerate(lines) if line.startswith(label)]
    if not hits:
        return ""
    index = hits[0]
    for offset in (2, 1, 3):
        if index + offset >= len(lines):
            continue
        candidate = lines[index + offset].strip()
        if not candidate or _is_label(candidate):
            continue
        if re.fullmatch(r"[（(]\d+[)）]", candidate):
            continue
        return _strip_code(candidate)
    return ""


def _trade_mode(lines: list[str]) -> str:
    """`贸易方式(0110)` / `一般贸易` 两行形态。"""
    hits = [i for i, line in enumerate(lines) if line.startswith("贸易方式")]
    if not hits:
        return ""
    index = hits[0]
    for offset in (1, 2, 3):
        if index + offset >= len(lines):
            continue
        candidate = lines[index + offset].strip()
        if candidate and not _is_label(candidate):
            return candidate
    return _strip_code(lines[index])


def _item_section(lines: list[str]) -> list[str]:
    """截取商品明细区域。"""
    start = 0
    end = len(lines)
    for index, line in enumerate(lines):
        if line.startswith("商品序号 商品编号"):
            start = index + 1
        if line.startswith("税费征收情况") and start:
            end = index
            break
    return lines[start:end]


def _parse_items(lines: list[str]) -> list[DeclarationItem]:
    """商品明细（版式 1）。

    一条商品可能有三组数量：`法定数量/法定单位`（跟在商品行后面）、`第二数量/第二单位`、
    `申报数量/申报单位`。**报关重量要取"申报数量"**——例如 `223320260001346587`：
    法定数量 26 套、申报数量 1076 千克，按法定数量取会得到 26（错 40 倍）。
    没有申报数量时才退回法定数量；`报关重量` 若仍不是重量单位，再退回第二数量。
    """
    section = _item_section(lines)
    items: list[DeclarationItem] = []
    index = 0
    while index < len(section):
        text = section[index].strip()
        match = ITEM_RE.match(text)
        if not match:
            index += 1
            continue
        serial, hs_code, name = match.groups()
        item = DeclarationItem(serial=int(serial), hs_code=hs_code, product_name=name)
        items.append(item)
        # 收集本商品的行（到下一个商品行为止）
        block: list[str] = []
        cursor = index + 1
        while cursor < len(section) and not ITEM_RE.match(section[cursor].strip()):
            value = section[cursor].strip()
            if value:
                block.append(value)
            cursor += 1
        index = cursor

        for line in block:
            value = QTY_RE.match(line)
            if not value:
                continue
            quantity, country, price, total, currency = value.groups()
            item.quantity = quantity
            item.unit = _unit_of(quantity)
            item.declare_quantity = quantity
            item.declare_unit = item.unit
            item.destination_country = country
            item.unit_price = price
            item.total_price = total
            item.currency = currency
            break

        for label, slot in (("第二数量/第二单位", "second"), ("申报数量/申报单位", "declare")):
            for position, line in enumerate(block):
                if not line.startswith(label):
                    continue
                token = _next_quantity(block[position + 1 :])
                if not token:
                    break
                if slot == "second":
                    item.second_quantity = token
                    item.second_unit = _unit_of(token)
                else:
                    item.declare_quantity = token
                    item.declare_unit = _unit_of(token)
                break
    return items


def _next_quantity(lines: list[str]) -> str:
    """从若干行里挑出第一个「数字 + 单位」的令牌（如 `1076千克 (410) 美元` → `1076千克`）。"""
    for line in lines:
        token = line.strip().split()[0] if line.strip() else ""
        if token and re.match(r"^[\d,.]+\S*$", token):
            return token
        if line.startswith(("法定数量", "第二数量", "申报数量", "商品序号")):
            break
    return ""


def _unit_of(quantity: str) -> str:
    match = re.match(r"^[\d,.]+\s*(\D.*)$", quantity or "")
    return match.group(1).strip() if match else ""


def _detect_format(text: str) -> int:
    """1 = 标签在值上方（JG09 出口退税联）；2 = 值在标签上方（仅供核对版式）。"""
    first = text.splitlines()[0] if text.splitlines() else ""
    if "仅供核对用" in text or (first.startswith("*") and first.endswith("*")):
        return 2
    if "海关编号：" in text or "预录入编号：" in text or "境内收发货人(" in text:
        return 1
    return 2 if "监管方式" in text else 1


# ---------- 版式 2 ----------


def _v2_above(lines: list[str], label: str) -> str:
    """取值在上方一行，如 `北仑海关` / `出境关别 (3104)`。"""
    for index, line in enumerate(lines):
        if line.startswith(label):
            for offset in (1, 2):
                if index - offset < 0:
                    continue
                candidate = lines[index - offset].strip()
                if candidate and not _is_label_v2(candidate):
                    return candidate
    return ""


def _v2_same_before(lines: list[str], label: str) -> str:
    """取值在同一行的标签之前，如 `C&F` / `成交方式 (2)`。"""
    for line in lines:
        index = line.find(label)
        if index > 0:
            return line[:index].strip()
    return ""


def _v2_line_before(lines: list[str], label: str) -> str:
    """取值是标签行的**上一行整行**，如 `FOB` / `成交方式 (3)`。"""
    for index, line in enumerate(lines):
        if line.startswith(label):
            if index - 1 >= 0:
                return lines[index - 1].strip()
    return ""


def _is_label_v2(value: str) -> bool:
    return any(value.startswith(label) or label in value for label in LABELS_V2)


def _v2_number_before(lines: list[str], label: str) -> str:
    value = _v2_above(lines, label)
    return re.sub(r"[^\d.]", "", value)


def _parse_header_v2(lines: list[str], text: str) -> DeclarationHeader:
    header = DeclarationHeader()
    header.declaration_no = _find_line(text, r"^\*?(\d{18})\*?$")
    if not header.declaration_no:
        header.declaration_no = _find_line(text, r"(\d{18})")
    header.pre_entry_no = _find_line(text, r"预录入编号[:：]\s*(\d{18})")
    header.contract_raw = _v2_above(lines, "合同协议号")
    export = _v2_above(lines, "申报日期出口日期")
    if re.fullmatch(r"\d{8}", export or ""):
        header.export_date = _clean_date(export)
        header.declare_date = _clean_date(export)
    elif re.fullmatch(r"\d{16}", export or ""):
        header.declare_date = _clean_date(export[:8])
        header.export_date = _clean_date(export[8:])
    header.export_port = _strip_code(_v2_above(lines, "离境口岸"))
    header.transport_mode = _v2_above(lines, "运输方式")
    header.trade_mode = _strip_code(_v2_above(lines, "监管方式"))
    header.deal_mode = _v2_line_before(lines, "成交方式")
    # 仅供核对版式：值是标签行上方
    header.freight = _v2_above(lines, "运费")
    header.insurance = _v2_above(lines, "保费")
    header.misc_fee = _v2_above(lines, "杂费保费")
    header.destination_country = _strip_code(_v2_above(lines, "运抵国"))
    header.destination_port = _strip_code(_v2_above(lines, "指运港"))
    header.consignor = _strip_code(_v2_above(lines, "境内发货人"))
    header.producer = _strip_code(_v2_above(lines, "生产销售单位"))
    header.pieces = _v2_number_before(lines, "件数")
    header.package_kind = _v2_above(lines, "包装种类")
    header.gross_weight = _v2_number_before(lines, "毛重(千克)")
    header.net_weight = _v2_number_before(lines, "净重(千克)")
    header.domestic_source = _find_line(text, r"\(\d+\)(\S+?) 照章征税")
    return header


V2_QTY_RE = re.compile(r"^([\d,.]+)([^\d\s]+)$")


def _parse_items_v2(lines: list[str], default_country: str = "") -> list[DeclarationItem]:
    """版式 2 的商品明细：`1 7507120000镍合金管` + 后续单价/总价/币制/国家。"""
    items: list[DeclarationItem] = []
    for line in lines:
        text = line.strip()
        match = ITEM_V2_RE.match(text)
        if not match:
            continue
        serial, hs_code, name = match.groups()
        items.append(
            DeclarationItem(serial=int(serial), hs_code=hs_code, product_name=name)
        )

    for item in items:
        marker = f"{item.serial} {item.hs_code}{item.product_name}"
        start = next((i for i, l in enumerate(lines) if l.strip().startswith(marker)), None)
        if start is None:
            continue
        segment = [l.strip() for l in lines[start + 1 : start + 14] if l.strip()]
        prices = [s for s in segment if re.fullmatch(r"[\d.]+", s)]
        currencies = [s for s in segment if s in ("美元", "人民币", "欧元", "日元", "港元")]
        countries = [
            s
            for s in segment
            if re.fullmatch(r"[\u4e00-\u9fff（）()]{2,12}", s)
            and not s.startswith("(")
            and "申报" not in s
            and "单位" not in s
            and "照章" not in s
        ]
        if prices:
            item.unit_price = prices[0]
        if len(prices) > 1:
            item.total_price = prices[1]
        if currencies:
            item.currency = {"美元": "USD", "人民币": "CNY", "欧元": "EUR"}.get(
                currencies[0], currencies[0]
            )
        if countries:
            item.destination_country = _strip_code(countries[-1])
        elif default_country:
            item.destination_country = default_country

    # 该版式在明细之前按「法定数量、申报数量、法定数量、申报数量…」列出数量，
    # 即每个商品出现两次（第一轮=法定数量，第二轮=申报数量），因此按顺序隔一个取一个。
    quantities: list[tuple[str, str]] = []
    for line in lines:
        text = line.strip()
        if text.startswith("(") or "证书" in text or "页" in text:
            continue
        if ITEM_V2_RE.match(text):
            break
        match = V2_QTY_RE.match(text)
        if match and match.group(2) in ("千克", "公斤", "个", "件", "米", "套"):
            quantities.append((match.group(1), match.group(2)))
    legal = quantities[::2] if len(quantities) >= 2 * len(items) else quantities
    for item, (value, unit) in zip(items, legal[: len(items)]):
        item.quantity = f"{value}{unit}"
        item.unit = unit
        item.declare_quantity = item.quantity
        item.declare_unit = unit
    return items


def parse_declaration(path: Path) -> ParsedDeclaration:
    """解析单份报关单 PDF。"""
    text = extract_text(path)
    lines = text.splitlines()
    fmt = _detect_format(text)
    if fmt == 2:
        header = _parse_header_v2(lines, text)
        items = _parse_items_v2(lines, header.destination_country)
    else:
        header = _parse_header(lines, text)
        items = _parse_items(lines)
    header.sheet_type = "预录单（仅供核对用）" if fmt == 2 else "出口退税联"
    header = recover_contract(header, lines)
    warnings: list[str] = []
    if not header.declaration_no:
        warnings.append("未识别到海关编号")
    if not header.contract_raw:
        warnings.append("未识别到合同协议号")
    if not items:
        warnings.append("未识别到商品明细")
    if not header.declaration_no and not items:
        raise DeclarationParseError(f"{path.name}: 关键字段解析失败")
    return ParsedDeclaration(
        source_file=path.name, header=header, items=items, warnings=warnings
    )


def parse_many(paths: list[Path]) -> list[ParsedDeclaration]:
    """批量解析，单个失败不影响其余。"""
    out: list[ParsedDeclaration] = []
    for path in paths:
        try:
            out.append(parse_declaration(path))
        except Exception as exc:  # noqa: BLE001
            out.append(
                ParsedDeclaration(
                    source_file=path.name,
                    header=DeclarationHeader(),
                    items=[],
                    warnings=[f"解析失败: {exc}"],
                )
            )
    return out


class DeclarationParseError(RuntimeError):
    """无法解析出报关单关键信息时抛出。"""
