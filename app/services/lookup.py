r"""逐条核验的输入解析与补全。

粘贴策略（2026-09-29 与用户对齐）：
1. **只落可靠字段**：报关单号（18 位）、合同号 / 采购订单号（`xxMT-…`）、海关编码（10 位）、币种；
   带标签（`报关金额：123`）的才取金额 / 重量 / 品名，**不按位置猜数字**；
2. 其余不准的字段交给「从飞书与睿贝补全」：按报关单号 / 合同号精确查飞书；
   稀疏行（只给了单号或合同）直接展开成该单 / 该合同的全部商品行；
3. 飞书不可达时退回本地快照并明确提示来源，绝不静默猜值。
"""
from __future__ import annotations

import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal, InvalidOperation

from app.services.feishu_client import PROJECT_ROOT

FIELDS = ("报关单号", "合同号_1", "采购订单号", "商品序号", "海关编码", "报关品名", "报关金额", "币种",
          "报关重量", "申报数量", "申报单位", "单价", "出口日期", "申报日期")

DECLARATION_RE = re.compile(r"(?<!\d)(\d{18})(?!\d)")
# 联合合同号（`25MT-07F379Y-F&477G`）要整段吃掉，不能在 `&` 处截断——
# 截断会让后续按批次分摊时认错批次（实测：截成 `…-F` 后取了 F 区块，金额少了 2.3 万）
ORDER_RE = re.compile(r"\b(PI-)?(\d{2}MT-\d{2}[A-Za-z]\d{3}[A-Za-z0-9\-]*(?:&[A-Za-z0-9\-]+)*)")
HS_RE = re.compile(r"(?<!\d)(\d{10})(?!\d)")
CURRENCY = ("USD", "EUR", "CNY", "RMB", "JPY", "HKD")
ALIASES = {"合同号": "合同号_1", "采购单号": "采购订单号", "总价": "报关金额",
           "重量": "报关重量", "商品名称": "报关品名"}
TABLE_SCHEMA = PROJECT_ROOT / "app" / "services" / "feishu_copy_schema.json"

FILL_FIELDS = ("合同号_1", "报关品名", "报关金额", "报关重量", "币种", "海关编码", "商品序号")
MAX_ROWS = 50

# 飞书精确检索的短期缓存（同一批单号反复补全时不重复打接口）
_CACHE: dict[tuple, tuple[float, list]] = {}
CACHE_TTL = 300
FETCH_FIELDS = ["报关单号", "合同号_1", "报关品名", "报关金额", "报关重量", "币种", "海关编码", "商品序号"]
# 单次检索里最多放多少个条件（飞书 filter 条件数有上限，保守取 20）
CONDITION_CHUNK = 20


def _clean(text: str) -> str:
    return re.sub(r"[\u000b\u000c\u0085\u2028\u2029]+", "\n", str(text or "")).replace("\u200b", "")


def _text(value) -> str:
    """飞书单元格取值：兼容 `{"text": ...}` / 列表 / 数字。"""
    if value is None:
        return ""
    if isinstance(value, dict):
        return _text(value.get("value", value.get("text", value.get("name", ""))))
    if isinstance(value, list):
        return "、".join(filter(None, (_text(item) for item in value)))
    return str(value).strip()


def _amount_text(value) -> str:
    """金额 / 重量统一成不带千分位的字符串。"""
    raw = _text(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    return match.group(0) if match else ""


def _sum_text(*values) -> str:
    """金额 / 重量文本相加；全空返回空串（不凭空造 0）。"""
    total = Decimal(0)
    seen = False
    for value in values:
        raw = str(value or "").replace(",", "").strip()
        if not raw:
            continue
        try:
            total += Decimal(raw)
        except InvalidOperation:
            continue
        seen = True
    if not seen:
        return ""
    return f"{total:.4f}".rstrip("0").rstrip(".") or "0"


def _collapse_records(records: list[dict]) -> list[dict]:
    """同一「报关单号 + 合同号 + 报关品名」的多行飞书记录 → 合并成一条报关行。

    财务会把脚本拆出的结果再手工补回多维表，于是同一张报关单在表里出现多行拆完单
    的记录（报关金额、报关重量都被拆散）。这类数据不是「一张报关单多个商品」，
    要先把金额与重量并回一条，再当成一次报关输入系统（2026-09-29 用户口径）。
    返回的 record 是 fields 已合并的副本，并带 `_拆分行数` / `_成员金额`；
    报关单号为空时不合并，避免把不同报关单的行并到一起。
    """
    order: list[tuple[str, str, str]] = []
    groups: dict[tuple[str, str, str], list[dict]] = {}
    for record in records:
        fields = record.get("fields") or {}
        key = (_text(fields.get("报关单号")), _text(fields.get("合同号_1")),
               _text(fields.get("报关品名")))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(record)

    collapsed: list[dict] = []
    for key in order:
        group = groups[key]
        members = [_amount_text((record.get("fields") or {}).get("报关金额")) for record in group]
        merged = dict(group[0])
        merged["_成员金额"] = members
        if len(group) > 1 and key[0]:
            fields = dict(group[0].get("fields") or {})
            fields["报关金额"] = _sum_text(*members)
            fields["报关重量"] = _sum_text(
                *[_amount_text((record.get("fields") or {}).get("报关重量")) for record in group])
            # 合并后不再是原来那一条拆分记录
            fields.pop("商品序号", None)
            merged["fields"] = fields
            merged["_拆分行数"] = len(group)
        collapsed.append(merged)
    return collapsed


def _extract_row(line: str) -> dict:
    """从一行自由文本里抽取**可确认**的字段；不带标签的金额 / 重量一律不猜。"""
    row = {field: "" for field in FIELDS}
    match = DECLARATION_RE.search(line)
    if match:
        row["报关单号"] = match.group(1)
    orders = list(ORDER_RE.finditer(line))
    # 带工厂后缀（-YH/-HD/-GYL）或 PI- 前缀的是采购订单号，另一个是报关合同号
    # 采购单号带 2~5 位工厂后缀（-ZY/-YH/-GYL…）；单字母结尾是批次（…-F），属出运合同
    purchase = next((item for item in orders
                     if item.group(1) or re.search(r"-[A-Z]{2,5}$", item.group(2))), None)
    # 飞书一行里常同时出现「合同号」与「合同号_1」（`26MT-03R036` + `26MT-03R036F`），
    # 要取更长更具体的那个——退回没有批次的基号会把同订单其它批次当成这一单算
    others = [item for item in orders if item is not purchase]
    others_text = {item.group(2) for item in others}
    contract = next(
        (
            item
            for item in others
            if not any(
                text != item.group(2) and text.startswith(item.group(2))
                for text in others_text
            )
        ),
        None,
    )
    if purchase:
        row["采购订单号"] = purchase.group(2).upper()
    if contract:
        row["合同号_1"] = contract.group(2).upper()
    elif purchase:
        row["合同号_1"] = purchase.group(2).upper()
    match = HS_RE.search(line)
    if match:
        row["海关编码"] = match.group(1)
    for currency in CURRENCY:
        if currency in line.upper():
            row["币种"] = "CNY" if currency == "RMB" else currency
            break
    for label, field in (("报关金额", "报关金额"), ("总价", "报关金额"),
                         ("报关重量", "报关重量"), ("商品序号", "商品序号")):
        if match := re.search(rf"{label}\s*[：:=]?\s*([\d,]+(?:\.\d+)?)", line):
            row[field] = match.group(1).replace(",", "")
    # 品名只取到下一个字段标签为止，避免把「报关金额：…」并进品名
    if match := re.search(
        r"报关品名\s*[：:=]?\s*([^\t，,;；|]+?)"
        r"(?=\s*(?:报关金额|总价|报关重量|商品序号|币种|USD|EUR|CNY|RMB)|\s*$)",
        line,
    ):
        row["报关品名"] = match.group(1).strip()
    return row


def parse_paste(text: str) -> dict:
    """粘贴文本 → 商品行。带表头表格按列名精确取；自由文本只取可确认字段。"""
    source = _clean(text)
    lines = [line for line in source.splitlines() if line.strip()]
    if not lines:
        return {"rows": [], "evidence": [], "warnings": []}

    # 1) 带表头的制表符表格（从飞书整列复制 / 导出的场景）
    if "\t" in lines[0]:
        grid = [line.split("\t") for line in lines]
        header = [ALIASES.get(cell.strip(), cell.strip()) for cell in grid[0]]
        if set(header) & {"报关单号", "合同号_1", "报关金额", "报关品名", "采购订单号"}:
            rows: list[dict] = []
            evidence: list[dict] = []
            for cells in grid[1:]:
                if not any(cell.strip() for cell in cells):
                    continue
                row = {name: cells[position].strip()
                       for position, name in enumerate(header)
                       if position < len(cells) and name in FIELDS and cells[position].strip()}
                if any(row.get(key) for key in ("报关单号", "合同号_1", "采购订单号")):
                    rows.append(row)
                    evidence.append({"来源行": "\t".join(cells)[:300]})
            return {"rows": rows, "evidence": evidence,
                    "warnings": ["单次最多 50 行，请分批粘贴"] if len(rows) > MAX_ROWS else []}

    # 2) 自由文本：逐行抽取；没有主键的行当作续行，并进下一行一起识别
    rows = []
    evidence = []
    buffer = ""
    for raw in lines:
        line = f"{buffer} {raw.strip()}".strip() if buffer else raw.strip()
        row = _extract_row(line)
        if not any(row.get(key) for key in ("报关单号", "合同号_1", "采购订单号")):
            buffer = line
            continue
        buffer = ""
        rows.append(row)
        evidence.append({"来源行": line[:300]})
    if buffer:
        row = _extract_row(buffer)
        if any(row.get(key) for key in ("报关单号", "合同号_1", "采购订单号")):
            rows.append(row)
            evidence.append({"来源行": buffer[:300]})
    # 单号与合同号被换行拆开时（一行只有单号、下一行只有合同号），并成同一条记录
    def _merge_rows(first: dict, second: dict) -> dict:
        """只把 second 的非空字段补进 first（空值不覆盖已有值）。"""
        out = dict(first)
        for key, value in second.items():
            if str(value or "").strip() and not str(out.get(key) or "").strip():
                out[key] = value
        return out

    merged: list[dict] = []
    for row in rows:
        if (merged
                and merged[-1].get("报关单号") and not merged[-1].get("合同号_1")
                and row.get("合同号_1") and not row.get("报关单号")):
            merged[-1] = _merge_rows(merged[-1], row)
            continue
        if (merged
                and merged[-1].get("合同号_1") and not merged[-1].get("报关单号")
                and row.get("报关单号") and not row.get("合同号_1")):
            merged[-1] = _merge_rows(merged[-1], row)
            continue
        merged.append(row)
    rows = merged
    return {"rows": rows, "evidence": evidence,
            "warnings": ["单次最多 50 行，请分批粘贴"] if len(rows) > MAX_ROWS else []}


def _snapshot() -> tuple[list[dict], str]:
    """本地快照兜底（飞书实时读取不可用时）。"""
    path = PROJECT_ROOT / "data" / "cache" / "records_full_copy_raw.json"
    if not path.exists():
        return [], "飞书实时读取不可用，且本地没有快照，请手工补齐"
    return json.loads(path.read_text(encoding="utf-8")), "飞书实时读取不可用，使用本地快照补齐；请核对来源"


def _fetch(declarations: set[str], contracts: set[str]) -> tuple[dict, dict, str, list[str]]:
    """按报关单号 / 合同号精确取飞书记录（飞书端检索 + 短期缓存 + 并行）。"""
    from app.services.feishu_client import FeishuClient, workbench_table

    warnings: list[str] = []
    try:
        client = FeishuClient(timeout=30)
        app, table = workbench_table(required=False)
        if not app or not table:
            raise ValueError("飞书目标表未配置")

        def search(values: list[str], field: str, operator: str) -> list[dict]:
            """一次检索多个值（OR 条件，分批），带 5 分钟缓存。"""
            values = sorted({value for value in values if value})
            if not values:
                return []
            key = (table, field, operator, tuple(values))
            hit = _CACHE.get(key)
            if hit and time.time() - hit[0] < CACHE_TTL:
                return hit[1]
            records: list[dict] = []
            for start in range(0, len(values), CONDITION_CHUNK):
                chunk = values[start:start + CONDITION_CHUNK]
                query = {"conjunction": "or", "conditions": [
                    {"field_name": field, "operator": operator, "value": [value]} for value in chunk]}
                records.extend(client.search_records(app, table, query))
            _CACHE[key] = (time.time(), records)
            return records

        with ThreadPoolExecutor(max_workers=2) as pool:
            decl_future = pool.submit(search, list(declarations), "报关单号", "is")
            contract_future = pool.submit(search, list(contracts), "合同号_1", "contains")
            decl_records = decl_future.result()
            contract_records = contract_future.result()

        by_decl: dict[str, list] = {}
        for record in decl_records:
            declaration = _text((record.get("fields") or {}).get("报关单号"))
            if declaration in declarations:
                by_decl.setdefault(declaration, []).append(record)
        by_contract: dict[str, list] = {}
        for record in contract_records:
            value = _text((record.get("fields") or {}).get("合同号_1")).upper()
            for contract in contracts:
                if contract.upper() in value:
                    by_contract.setdefault(contract, []).append(record)
        return by_decl, by_contract, "飞书实时记录", warnings
    except Exception:  # noqa: BLE001  网络 / 权限不可用时退回快照
        records, note = _snapshot()
        if note:
            warnings.append(note)
        by_decl, by_contract = {}, {}
        for record in records:
            fields = record.get("fields") or {}
            declaration = _text(fields.get("报关单号"))
            if declaration:
                by_decl.setdefault(declaration, []).append(record)
            contract = _text(fields.get("合同号_1"))
            if contract:
                by_contract.setdefault(contract, []).append(record)
        return by_decl, by_contract, "飞书本地快照", warnings


def _merge(base: dict, record: dict) -> dict:
    """用飞书记录补齐一行（不覆盖已有值）。"""
    fields = record.get("fields") or {}
    out = dict(base)
    for name in FILL_FIELDS:
        value = _amount_text(fields.get(name)) if name in ("报关金额", "报关重量") else _text(fields.get(name))
        if value and not str(out.get(name) or "").strip():
            out[name] = value
    # 合同号例外：粘贴里被截断的联合号（`25MT-07F379Y-F` vs 飞书 `25MT-07F379Y-F&477G`）
    # 要以飞书为准，否则后面按批次分摊会认错批次
    feishu_contract = _text(fields.get("合同号_1"))
    current = str(out.get("合同号_1") or "").strip()
    if feishu_contract and feishu_contract != current:
        have = set(re.split(r"[&,，;；]", current)) - {""}
        want = set(re.split(r"[&,，;；]", feishu_contract)) - {""}
        # 「我方值是飞书值的子集」或「我方值是飞书值的前缀」（`26MT-03R036` vs
        # `26MT-03R036F`、`25MT-07F477` vs `25MT-07F477&26MT-07C330`）都算飞书更完整
        if (not have or have <= want
                or any(feishu_contract.upper().startswith(part.upper()) for part in have)):
            out["合同号_1"] = feishu_contract
    declaration = _text(fields.get("报关单号"))
    if declaration and not str(out.get("报关单号") or "").strip():
        out["报关单号"] = declaration
    return out


def enrich_rows(rows: list[dict], *, sources: list[dict] | None = None) -> dict:
    """按报关单号 / 合同号从飞书与睿贝补全；稀疏行展开成该单的全部商品行。"""
    from app.services.erp_cache_index import lookup as erp_lookup

    declarations = {str(row.get("报关单号") or "").strip() for row in rows if str(row.get("报关单号") or "").strip()}
    contracts = {str(row.get("合同号_1") or "").strip() for row in rows if str(row.get("合同号_1") or "").strip()}
    if sources is not None:
        by_decl: dict[str, list] = {}
        by_contract: dict[str, list] = {}
        label, warnings = "飞书记录", []
        for record in sources:
            fields = record.get("fields") or {}
            declaration = _text(fields.get("报关单号"))
            if declaration:
                by_decl.setdefault(declaration, []).append(record)
            contract = _text(fields.get("合同号_1"))
            if contract:
                by_contract.setdefault(contract, []).append(record)
    else:
        by_decl, by_contract, label, warnings = _fetch(declarations, contracts)

    output: list[dict] = []
    seen_contracts: set[str] = set()
    for index, raw in enumerate(rows, 1):
        row = {key: str(value or "").strip() for key, value in raw.items() if key in FIELDS}
        notes: list[str] = []

        # 只有采购单号 → 用睿贝出运单反查合同号
        purchase = row.get("采购订单号", "")
        if purchase and not row.get("合同号_1"):
            shipments = erp_lookup(purchase).get("shipments") or []
            codes = {code.strip() for item in shipments
                     for code in re.split(r"[,&，、;；]", _text(item.get("invoiceCode"))) if code.strip()}
            if len(codes) == 1:
                row["合同号_1"] = codes.pop()
                notes.append("合同号由睿贝出运单补全")
            elif len(codes) > 1:
                warnings.append(f"第 {index} 行采购单关联多个出运合同，请指定合同号")

        declaration = row.get("报关单号", "")
        contract = row.get("合同号_1", "")
        if contract and contract not in contracts:
            contracts.add(contract)
            if sources is None and not declaration:
                # 刚补出合同号时，把该合同的行也取回来
                _, extra, _, _ = _fetch(set(), {contract})
                by_contract.update(extra)
        sparse = not row.get("报关品名") and not row.get("报关金额")
        candidates: list[dict] = []
        if declaration:
            candidates = list(by_decl.get(declaration, []))
        elif contract:
            candidates = list(by_contract.get(contract, []))
            seen_contracts.add(contract)
        # 同一报关单的多行「拆完单记录」先并回一条，不要摊成多行
        candidates = _collapse_records(candidates)

        if declaration and not sparse and candidates:
            for name in ("报关品名", "报关金额"):
                if row.get(name):
                    wanted = row[name]
                    if name == "报关金额":
                        # 粘贴的可能是被拆散的那一行的金额，也要认得出它属于哪一条报关行
                        wanted_amount = _amount_text(wanted)
                        narrowed = [item for item in candidates
                                    if _amount_text((item.get("fields") or {}).get(name)) == wanted_amount
                                    or wanted_amount in (item.get("_成员金额") or [])]
                    else:
                        narrowed = [item for item in candidates
                                    if _text((item.get("fields") or {}).get(name)) == wanted]
                    if narrowed:
                        candidates = narrowed
            if len(candidates) == 1:
                candidate = candidates[0]
                split_rows = int(candidate.get("_拆分行数") or 1)
                before = dict(row)
                row = _merge(row, candidate)
                if split_rows > 1:
                    # 以合并后的报关金额 / 重量为准，否则只算到被拆散的那一部分
                    total = _amount_text((candidate.get("fields") or {}).get("报关金额"))
                    total_weight = _amount_text((candidate.get("fields") or {}).get("报关重量"))
                    notes.append(f"同报关单 {split_rows} 行拆分记录已并回一条，报关金额合计 {total or '待核对'}")
                    if total:
                        row["报关金额"] = total
                    if total_weight:
                        row["报关重量"] = total_weight
                    row["_已合并"] = True
                filled = [name for name in FILL_FIELDS
                          if row.get(name) and not str(before.get(name) or "").strip()]
                if filled:
                    notes.append(f"{'、'.join(filled)}由{label}补全")
            elif len(candidates) > 1:
                warnings.append(f"第 {index} 行有多个匹配商品行，请补充报关品名或商品序号")
            else:
                warnings.append(f"第 {index} 行在{label}里没有找到对应商品行")

        # 稀疏行（只给了报关单号或合同号）→ 展开成该单 / 该合同的全部商品行
        # （同一报关单的拆完单记录已先并回一条，所以这里展开的是「一张报关单一条记录」）
        if sparse and candidates:
            room = MAX_ROWS - len(output)
            expanded = candidates[:max(0, room)]
            for record in expanded:
                child = {"报关单号": declaration, "合同号_1": contract, "币种": "USD"}
                child = _merge(child, record)
                split_rows = int(record.get("_拆分行数") or 1)
                merged_note = f"，已把同报关单 {split_rows} 行拆分记录并回一条" if split_rows > 1 else ""
                child["_补全说明"] = f"按{'报关单号' if declaration else '合同号'}从{label}展开{merged_note}"
                if split_rows > 1:
                    child["_已合并"] = True
                output.append(child)
            if len(candidates) > len(expanded):
                warnings.append(f"第 {index} 行展开后超过 {MAX_ROWS} 行上限，请分批处理")
            continue

        row["_补全说明"] = "；".join(notes)
        output.append(row)

    # 粘贴了同一报关单的多行拆完单记录时，只保留合并后的那一条；
    # 否则下游按「单号 + 合同 + 品名」再合并求和会把同一批金额重复计算。
    unique: list[dict] = []
    seen_split_keys: set[tuple[str, str, str]] = set()
    duplicate_rows = 0
    for row in output:
        declaration = str(row.get("报关单号") or "").strip()
        key = (declaration, str(row.get("合同号_1") or "").strip(),
               str(row.get("报关品名") or "").strip())
        if row.pop("_已合并", False) and declaration:
            if key in seen_split_keys:
                duplicate_rows += 1
                continue
            seen_split_keys.add(key)
        unique.append(row)
    if duplicate_rows:
        warnings.append(f"同一报关单的 {duplicate_rows} 行重复拆分记录已合并为一条")
    output = unique

    if len(output) > MAX_ROWS:
        warnings.append(f"补全后共 {len(output)} 行，超过单批上限，请分批处理")
        output = output[:MAX_ROWS]
    return {"rows": output, "warnings": warnings}


def declarations_from_rows(rows: list[dict]) -> list[dict]:
    """页面表单行 → 拆单输入行（补齐拆单需要的列名）。"""
    out: list[dict] = []
    for row in rows:
        if not any(str(value or "").strip() for value in row.values()):
            continue
        out.append(
            {
                "来源文件": "逐单核验输入",
                "合同号_1": str(row.get("合同号_1") or "").strip(),
                "报关单号": str(row.get("报关单号") or "").strip(),
                "商品序号": str(row.get("商品序号") or "").strip(),
                "报关品名": str(row.get("报关品名") or "").strip(),
                "海关编码": str(row.get("海关编码") or "").strip(),
                "报关重量": str(row.get("报关重量") or "").strip(),
                "申报数量": str(row.get("申报数量") or "").strip(),
                "申报单位": str(row.get("申报单位") or "").strip(),
                "总价": str(row.get("报关金额") or row.get("总价") or "").strip(),
                "币种": str(row.get("币种") or "").strip(),
            }
        )
    return out


def validate_rows(rows: list[dict]) -> list[str]:
    """提交前的最小校验：缺关键字段就明确告诉用户缺什么。"""
    errors = []
    for index, row in enumerate(rows, 1):
        if not any(str(value or "").strip() for value in row.values()):
            continue
        missing = [name for name in ("报关单号", "合同号_1", "报关品名", "报关金额")
                   if not str(row.get(name) or "").strip()]
        if missing:
            errors.append(f"第 {index} 行缺少{'、'.join(missing)}")
            continue
        if not re.fullmatch(r"\d{18}", str(row["报关单号"]).strip()):
            errors.append(f"第 {index} 行报关单号须为 18 位数字")
        try:
            if Decimal(str(row["报关金额"]).replace(",", "")) <= 0:
                errors.append(f"第 {index} 行报关金额须大于 0")
        except InvalidOperation:
            errors.append(f"第 {index} 行报关金额不是数字")
    return errors
