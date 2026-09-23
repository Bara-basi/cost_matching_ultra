"""从解析好的入库单里挑出「有效附件」：分批 + 去旧版。

口径（2026-09-24 用户确认）：

1. **归属**：以睿贝采购单目录为准；附件名里的单号只做校验。
   名字里写了别的单号时，按「去掉 25MT/26MT 年份前缀」后比对，
   对不上就标记为别的订单的附件（不进本单成本）。
2. **分批**：批次令牌不同 → 各批次独立保留，绝不合并。
3. **组内定版**：
   - 完全重复（同名 / 同内容）→ 只留附件栏里最靠后的一份；
   - 费用补充件（补款/补木箱费/样品费…）→ 与正单并列保留；
   - 版本替换（更新/最终/最新/已修改）→ 取最靠后的一份，其余进「旧版」；
   - 名字判不出 → 比明细：后者是前者的超集或两者相同 → 取后者；
     两份明细互不相同且互补 → 视为两批，都保留并标记待人工确认。

顺序来源：`.cache/erp/attachments/lists/<采购单号>.json` 的 `attachmentList`
（睿贝附件栏的真实顺序，越靠后越新）。
"""
from __future__ import annotations

import collections
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT
from app.services.grn_extract import (
    PARSED_ROOT,
    normalize_settled,
    safe_name,
    to_decimal,
)

LIST_DIR = CACHE_ROOT / "attachments" / "lists"
MANIFEST = CACHE_ROOT / "attachments" / "grn_manifest.json"
PURCHASES = CACHE_ROOT / "purchases" / "purchases.jsonl"

_ERP_AMOUNTS: dict[str, Decimal] | None = None
_ERP_SUPPLIERS: dict[str, str] | None = None
_FACTORY_ALIASES: dict[str, set[str]] | None = None


def factory_aliases() -> dict[str, set[str]]:
    """工厂后缀（YH/HD/XMLS/…）→ 可能的中文简称集合。

    从睿贝采购单表里学：`26MT-03R411-HD` 的供应商是「浙江鸿迪管业有限公司」，
    于是 `HD ⇄ 鸿迪`。同一个缩写可能被不同工厂用过（`ZY` 既有卓业也有彰源），
    所以这里保留**集合**：只在唯一对应时才敢用来判定「附件是不是别家工厂的」。
    """
    global _FACTORY_ALIASES
    if _FACTORY_ALIASES is not None:
        return _FACTORY_ALIASES
    aliases: dict[str, set[str]] = collections.defaultdict(set)
    try:
        from app.services.supplier_names import SupplierNames

        names = SupplierNames()
    except Exception:  # noqa: BLE001
        names = None
    if PURCHASES.exists():
        for line in PURCHASES.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except ValueError:
                continue
            code = str(row.get("purchase_code") or "")
            match = re.search(r"[-_]([A-Za-z]{2,5})$", code)
            if not match:
                continue
            abbr = match.group(1).upper()
            if abbr not in FACTORY_ABBR:
                continue
            supplier = str(row.get("supplierName") or "")
            short = names.short(supplier) if names else supplier
            if short:
                aliases[abbr].add(short)
    _FACTORY_ALIASES = aliases
    return aliases


def factory_token(text: str, *, suffix_only: bool = False) -> str:
    """从单号/文件名里认出工厂后缀（统一成中文简称）；认不出返回空串。

    `suffix_only=True` 时只看单号末尾的缩写（采购单号用）；
    否则优先找文件名里的中文简称（附件名用）。含义不唯一的一律返回空串。
    """
    aliases = factory_aliases()
    raw = str(text or "")
    if suffix_only:
        match = re.search(r"[-_\s]([A-Za-z]{2,5})\s*$", raw)
        if match:
            names = aliases.get(match.group(1).upper())
            if names and len(names) == 1:
                return next(iter(names))
        return ""
    upper = raw.upper()
    for token in re.findall(r"[A-Za-z]{2,5}", upper):
        if token in {"MT", "ADD", "PI", "KG", "PCS"}:
            continue
        names = aliases.get(token)
        if names and len(names) == 1:
            return next(iter(names))
    try:
        from app.services.supplier_names import SupplierNames

        names = SupplierNames()
    except Exception:  # noqa: BLE001
        return ""
    for short in names.short_to_name:
        if short and short in str(text or ""):
            return short
    return ""


def erp_amount(purchase_code: str) -> Decimal | None:
    """睿贝采购单金额（旧数据，只作参照）。"""
    global _ERP_AMOUNTS
    if _ERP_AMOUNTS is None:
        _ERP_AMOUNTS = {}
        if PURCHASES.exists():
            for line in PURCHASES.read_text(encoding="utf-8").splitlines():
                if not line.strip():
                    continue
                try:
                    row = json.loads(line)
                except ValueError:
                    continue
                value = to_decimal(row.get("amount"))
                if row.get("purchase_code") and value is not None:
                    _ERP_AMOUNTS[row["purchase_code"]] = value
    return _ERP_AMOUNTS.get(purchase_code)


def erp_supplier(purchase_code: str) -> str:
    """睿贝采购单上的供应商名称（用于判断附件是不是别家工厂的）。"""
    global _ERP_SUPPLIERS
    if _ERP_SUPPLIERS is None:
        _ERP_SUPPLIERS = {}
        if PURCHASES.exists():
            for line in PURCHASES.read_text(encoding="utf-8").splitlines():
                if not line.strip():
                    continue
                try:
                    row = json.loads(line)
                except ValueError:
                    continue
                code = row.get("purchase_code")
                if code:
                    _ERP_SUPPLIERS[code] = str(row.get("supplierName") or "")
    return _ERP_SUPPLIERS.get(purchase_code, "")


# 名录里的「供应链/总部」这类不是工厂名，认工厂时要排掉
GENERIC_SHORTS = {"供应链", "总部", "本部", "一部", "二部", "迈拓", "公司", "采购", "加工"}


def named_factories(text: str) -> set[str]:
    """附件名里提到的工厂（中文简称）。"""
    try:
        from app.services.supplier_names import SupplierNames

        names = SupplierNames()
    except Exception:  # noqa: BLE001
        return set()
    found = {
        short
        for short in names.short_to_name
        if short and len(short) >= 2 and short not in GENERIC_SHORTS and short in str(text or "")
    }
    # 只保留最长的匹配，避免「希蒙」被「希蒙雷斯」之类的短名重复命中
    return {short for short in found if not any(short != other and short in other for other in found)}


def foreign_factory_attachment(stem: str, purchase_code: str, po_short: str) -> bool:
    """附件名点到别家工厂时返回 True（本单供应商是另一家工厂）。

    只在证据充分时才判：本单供应商是**具体工厂**、附件名提到**另一家**工厂，
    且两边名字不是「包含关系」的写法差异（`叁奇` ⊂ `新叁奇`、`格瑞` ⊂ `格瑞新`）。
    供应链/迈拓代采这类单子会牵涉多个主体，一律不判。
    """
    if not po_short or po_short in GENERIC_SHORTS:
        return False
    code_text = str(purchase_code or "")
    if "溪流" in code_text or "迈拓" in code_text or "GYL" in code_text.upper():
        return False
    mentioned = named_factories(stem)
    if not mentioned:
        return False
    for name in mentioned:
        if name in po_short or po_short in name:
            return False  # 同一家的不同写法
    return True

ORDER_CORE = r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}(?:Y)?(?:[- ]?ADD\d*)?"
ORDER_RE = re.compile(r"\d{2}[A-Za-z]{2}-?\d{2}[A-Za-z]\d{2,3}", re.IGNORECASE)
ORDER_KEY_RE = re.compile(
    r"(\d{2})[A-Za-z]{2}-?(\d{2}[A-Za-z]\d{2,3})(Y)?(?:[\s\-_]*(ADD\d*))?",
    re.IGNORECASE,
)
ORDER_FULL_RE = re.compile(ORDER_CORE, re.IGNORECASE)
JUNK_PAREN = re.compile(r"(?:\s*[(（]\d+[)）])+\s*$")
JUNK_PAREN_EXT = re.compile(r"(?:\s*[(（]\d+[)）])+(?=\.[A-Za-z]+$)")
BATCH_TEXT = re.compile(r"第([一二三四五六七八九十\d]+)批")

ABBR = {
    "YH", "HD", "XMLS", "GYL", "HX", "JX", "XL", "ZK", "MJ", "HT", "ZY", "JS",
    "JF", "JE", "KLX", "BF", "DS", "ZT", "SW", "LZ", "JH", "SJ", "WL", "ZX",
    "BND", "XM", "XW", "FH", "JT", "YW", "SK", "YX", "CL", "JJ", "XDD", "CY",
    "YT", "TA", "FK", "MCS", "ADD",
}
# ADD 是附加单标记、不是工厂后缀，认工厂时要排掉
FACTORY_ABBR = {token for token in ABBR if token != "ADD"}

FEE_WORDS = (
    "补款", "补样品费", "样品费", "补木箱费", "木箱费", "含木箱费", "补包装费", "包装费",
    "含包装", "补绳", "吊绳费", "含绳子费", "补坡口费", "加工费", "抵扣", "尾款",
    "材料款", "试样费", "渗透实验费", "补做", "补货", "装卸费", "定金",
)
VERSION_WORDS = ("更新版", "更新", "最终", "最新", "已修改", "修改", "最后一批", "2.0")


@dataclass
class Item:
    """一份附件的选择视角。"""

    file: str                 # 磁盘文件名（= 解析缓存里的 _file）
    original: str             # 睿贝里的原始附件名
    order: int                # 附件栏位置（越大越新）
    batch: str = ""
    is_fee: bool = False
    is_version: bool = False
    other_order: bool = False
    foreign_factory: bool = False
    payload: dict[str, Any] = field(default_factory=dict)

    @property
    def stem(self) -> str:
        return clean_stem(self.original)


def core(text: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", str(text or "").upper())


def strip_year(text: str) -> str:
    """去掉 25MT/26MT 这类年份前缀，便于跨年比对。"""
    return re.sub(r"^\d{2}MT", "MT", core(text))


def same_order(attachment_name: str, purchase_code: str) -> bool:
    """附件名里的单号是否指向本采购单（或它所属的同一份采购合同）。

    比对键 = 部门/业务员/流水（+ 工厂）：
    * 年份（25MT/26MT）忽略——年末年初常写错；
    * `Y`（佣金）忽略——入库单抬头经常不写 Y；
    * 批次字母忽略；
    * **工厂后缀必须一致**——同一单号下 -HD / -YH 是两家工厂、两张采购单，
      金额不能互相认领（历史病例：勇恒的入库单被算进鸿迪的采购单）；
    * `ADDn` 差异**不**判异常——同一张合并入库单经常同时覆盖主单与附加单，
      由成本匹配阶段按 ERP 采购金额拆分。
    """
    target = order_key(purchase_code)
    if not target:
        return True
    name = str(attachment_name or "")
    matches = list(ORDER_KEY_RE.finditer(name))
    if not matches:
        # 文件名里完全没有单号时不当异常
        return True
    target_base = target.split("#")[0]
    target_factory = factory_token(str(purchase_code or ""), suffix_only=True)
    name_factory = factory_token(name)
    for match in matches:
        key = order_key(match.group(0))
        if not key:
            continue
        if key.split("#")[0] != target_base:
            continue
        # 只有「采购单后缀与附件里写的工厂都能唯一认出来、且互不相同」时才判异常
        if target_factory and name_factory and name_factory != target_factory:
            continue
        return True
    return False


def order_key(code: str) -> str:
    """把单号压成「部门+业务员+流水+ADD+工厂」的比对键。"""
    match = ORDER_KEY_RE.search(str(code or ""))
    if not match:
        return ""
    # match.group(1) 是年份，去掉；保留 部门+业务员+流水 与 ADD
    body = match.group(2).upper()
    add = (match.group(4) or "").upper()
    if add == "ADD":
        # 命名规范：ADD1 有时被写成 ADD，视为同一单
        add = "ADD1"
    key = body + ("#" + add if add else "")
    factory = factory_token(str(code or ""), suffix_only=True)
    if factory:
        key += f"#F:{factory}"
    return key


def batch_token(original: str) -> str:
    """批次令牌：单号核心之后的短令牌，跳过供应商缩写；中文「第 N 批」另算。"""
    stem = clean_stem(original)
    chinese = BATCH_TEXT.search(stem)
    if chinese:
        return chinese.group(1)
    match = re.search(ORDER_CORE, stem, re.IGNORECASE)
    if not match:
        return ""
    rest = stem[match.end():].lstrip(" -_")
    window = rest[:24]
    # 批次令牌必须紧跟在单号后面（中间只允许空格/横杠/下划线）；
    # 一旦隔了中文（如「…入库单负奖励40%」），后面的数字就不是批次。
    while window:
        found = re.match(r"[\s\-_]*([A-Za-z0-9]+)", window)
        if not found:
            return ""
        token = found.group(1)
        following = window[found.end():found.end() + 1]
        upper = token.upper()
        # `-1-` 这种：数字后还有分隔符，说明属于单号本身（采购批次），不是发货批次
        if upper.isdigit() and following in {"-", "_"}:
            window = window[found.end():]
            continue
        if upper in ABBR or (upper.isdigit() and len(upper) >= 4):
            window = window[found.end():]
            continue
        for abbr in ABBR:
            if upper.startswith(abbr) and 0 < len(upper) - len(abbr) <= 2:
                return upper[len(abbr):]
        return upper
    return ""


def clean_stem(name: str) -> str:
    """去掉所有尾部 (1)(2) 以及扩展名。"""
    stem = Path(name).stem
    stem = JUNK_PAREN_EXT.sub("", stem)
    stem = JUNK_PAREN.sub("", stem)
    return stem.strip()


def load_order(purchase_code: str) -> dict[str, int]:
    """原始附件名 → 附件栏位置。"""
    path = LIST_DIR / f"{safe_name(purchase_code)}.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return {}
    names = [
        str(item.get("attachmentName") or "")
        for item in payload.get("attachmentList") or []
        if "入库单" in str(item.get("attachmentName") or "")
    ]
    return {name: index for index, name in enumerate(names)}


def disk_to_original() -> dict[str, str]:
    """磁盘文件名 → 原始附件名（manifest 里存了原始名）。"""
    mapping: dict[str, str] = {}
    if MANIFEST.exists():
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        for entry in manifest.values():
            for item in entry.get("files") or []:
                path = item.get("path") or ""
                if path:
                    mapping[Path(path).name] = item.get("name") or Path(path).name
    return mapping


def load_items(purchase_code: str) -> list[Item]:
    """读出一个采购单下所有已解析附件，按附件栏顺序（越靠后越新）。"""
    folder = PARSED_ROOT / safe_name(purchase_code)
    if not folder.exists():
        return []
    order_map = load_order(purchase_code)
    disk_map = disk_to_original()
    try:
        from app.services.supplier_names import SupplierNames

        po_factory = SupplierNames().short(erp_supplier(purchase_code))
    except Exception:  # noqa: BLE001
        po_factory = ""
    items: list[Item] = []
    for path in sorted(folder.glob("*.json")):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        # 老缓存里模型填的 settled_totals 可能只等于第一个区块，统一校正（下游取区块）
        normalize_settled(payload)
        file_name = payload.get("_file") or path.stem
        original = disk_map.get(file_name, file_name)
        stem = clean_stem(original)
        foreign_factory = foreign_factory_attachment(stem, purchase_code, po_factory)
        items.append(
            Item(
                file=file_name,
                original=original,
                order=order_map.get(original, 10_000 + len(items)),
                batch=batch_token(original),
                is_fee=any(word in stem for word in FEE_WORDS),
                is_version=any(word in stem for word in VERSION_WORDS),
                other_order=bool(ORDER_RE.search(original))
                and not same_order(original, purchase_code),
                foreign_factory=foreign_factory,
                payload=payload,
            )
        )
    items.sort(key=lambda item: (item.order, item.file))
    return items


def line_identity(payload: dict[str, Any]) -> set[tuple]:
    """明细的「货物身份」：材质 + 品名 + 规格 + 数量，**不含金额**。

    为什么不含金额：同一批货的单据常被改过价（先写错、后来更正），
    带上金额就判不出「后一份是前一份的续写」了。

    关键：同一张表可能一份是规则解析（中文列名）、一份是模型解析（英文字段），
    这里统一取键，否则两种来源的指纹永远对不上（实测 26MT-02N182 就是）。
    """
    out: set[tuple] = set()
    for line in payload.get("lines") or []:
        if not isinstance(line, dict):
            continue
        key = line_key(line)
        if key:
            out.add(key)
    return out


def goods_key(payload: dict[str, Any]) -> frozenset:
    """跨来源可比的货物指纹：材质 + 规格 + 数量 + 单位（**不含品名**）。

    规则解析的模板有的没有「品名」列，模型解析则有；把品名放进指纹会让
    同一批货的两种解析永远对不上。品名单独用 `names_compatible` 判断。
    """
    out = set()
    for line in payload.get("lines") or []:
        if not isinstance(line, dict):
            continue
        key = line_key(line)
        if not key:
            continue
        out.add((key[0], key[2], key[3], key[4]))
    return frozenset(out)


def line_names(payload: dict[str, Any]) -> set[str]:
    out: set[str] = set()
    for line in payload.get("lines") or []:
        if not isinstance(line, dict):
            continue
        key = line_key(line)
        if key and key[1]:
            out.add(key[1])
    return out


def names_compatible(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """两侧品名是否相容：任一侧没写品名就算相容；都写了则要求有交集或包含关系。"""
    a, b = line_names(left), line_names(right)
    if not a or not b:
        return True
    for x in a:
        for y in b:
            if x == y or x in y or y in x:
                return True
    return False


def _norm_text(value: Any) -> str:
    text = str(value or "").strip().upper()
    return re.sub(r"[\s×x*＊,，、_\-]+", "", text)


def line_key(line: dict) -> tuple | None:
    """一行明细的规范化身份（材质 / 品名 / 规格 / 数量）。"""

    def get(*keys: str) -> str:
        for name in keys:
            value = line.get(name)
            if value not in (None, ""):
                return str(value).strip()
        return ""

    material = get("material", "材质")
    name = get("name", "品名")
    spec = get("spec", "规格")
    if not spec:
        # 2026-09-24 修：原来只认「外径 MM / 外径 / OD」，遇到把外径拆成
        # 「外径1 / 外径2」两列的模板时 spec 退化成只剩壁厚，货物指纹变得极粗，
        # 于是把独立批次误判成「后一份的子集」丢掉（实锤：26MT-05X169-6，
        # 丢掉的 26,020.70 恰好是睿贝采购金额与飞书三行合计的缺口）。
        dims = [
            get("外径1", "外径 MM", "外径MM", "外径", "OD1", "OD"),
            get("外径2", "OD2"),
            get("壁厚 MM", "壁厚MM", "壁厚", "平均壁厚", "WT"),
            get("定尺 MM", "定尺 M", "定尺MM", "定尺M", "定尺"),
        ]
        spec = "".join(dim for dim in dims if dim)
    qty = get("qty", "settled_qty", "数量")
    unit = get("unit", "单位")
    if not qty:
        for label, value in line.items():
            if value in (None, ""):
                continue
            if label.startswith(("数量", "支数", "米数", "只数", "件数")):
                qty = str(value)
                if not unit:
                    unit = label.replace("数量", "").strip()
                break
    if not any((material, name, spec, qty)):
        return None
    return (_norm_text(material), _norm_text(name), _norm_text(spec), _norm_text(qty), _norm_text(unit))


def line_signature(payload: dict[str, Any]) -> tuple:
    """兼容旧口径：把身份集合转成可比较的元组。"""
    return tuple(sorted(line_identity(payload)))


def batch_set(payload: dict[str, Any]) -> set[str]:
    """文件里**真正有金额**的批次集合（用于识别「递进快照」）。"""
    out: set[str] = set()
    batches = payload.get("batches") or []
    if isinstance(batches, list):
        for item in batches:
            if not isinstance(item, dict):
                continue
            amount = to_decimal(item.get("settled_amount"))
            if amount:
                out.add(str(item.get("batch") or "").strip().upper())
    return out


def block_amounts(payload: dict[str, Any]) -> collections.Counter:
    """各「实发数据-X」区块的金额（分位量化），用于识别递进版本。

    同一张合并表会被反复另存：`-A`（只含 A 批）→ `-C`（A+B、C）→ `-D`（A+B、C、D），
    后一份的区块金额集合一定包含前一份。
    """
    counter: collections.Counter = collections.Counter()
    for block in payload.get("settled_blocks") or []:
        if not isinstance(block, dict):
            continue
        amount = to_decimal(block.get("amount"))
        if amount:
            counter[str(amount.quantize(Decimal("0.01")))] += 1
    return counter


def block_token_amounts(payload: dict[str, Any]) -> dict[str, Decimal]:
    """批次记号 → 金额（同一份表里重复且不同的记号视为不唯一，跳过）。"""
    out: dict[str, Decimal] = {}
    ambiguous: set[str] = set()
    seen: dict[str, Decimal] = {}
    for block in payload.get("settled_blocks") or []:
        if not isinstance(block, dict):
            continue
        token = _batch_of_label(str(block.get("label") or ""))
        amount = to_decimal(block.get("amount"))
        if not token or not amount:
            continue
        if token in seen and seen[token] != amount:
            ambiguous.add(token)
        seen[token] = amount
        out[token] = amount
    for token in ambiguous:
        out.pop(token, None)
    return out


def _batch_of_label(label: str) -> str:
    text = str(label or "").strip().upper()
    if not text or "+" in text or "&" in text:
        return ""
    chinese = re.search(r"第([一二三四五六七八九十\d]+)批", text)
    if chinese:
        return "第" + chinese.group(1) + "批"
    cleaned = re.sub(r"[\d\./\s\-_]+", " ", text)
    letters = re.findall(r"(?<![A-Z])([A-Z])(?![A-Z])", cleaned)
    return letters[-1] if letters else ""


def block_values(payload: dict[str, Any]) -> list[Decimal]:
    """实发区块金额列表（rule 与 llm 两种结构都认）。

    注意：模型返回的 JSON 里 `settled_blocks` 与 `batches` 常常是同一批数字的两种写法，
    两个都累加会把金额翻倍——所以只在没有 `settled_blocks` 时才用 `batches`。
    """
    values: list[Decimal] = []
    for block in payload.get("settled_blocks") or []:
        if isinstance(block, dict):
            value = to_decimal(block.get("amount"))
            if value is not None:
                values.append(value)
    if not values:
        for batch in payload.get("batches") or []:
            if isinstance(batch, dict):
                value = to_decimal(batch.get("settled_amount"))
                if value is not None:
                    values.append(value)
    return values


def order_value(payload: dict[str, Any]) -> Decimal | None:
    """订单数据区金额（下单时的预期金额），只作参照。"""
    value = to_decimal(payload.get("order_amount"))
    if value is not None:
        return value
    totals = payload.get("order_totals") or {}
    if isinstance(totals, dict):
        return to_decimal(totals.get("amount"))
    return None


def claim_reason(
    item: Item, purchase_code: str, kept_others: list[Item], foreign_list: list[Item]
) -> str:
    """附件名写了别的单号时，判断有没有证据说明它其实是本单的。

    只认两种强证据（都要能对上工厂，避免把别家的钱算进来）：

    1. **单据抬头的单号指向本单**——文件名写错是常见笔误
       （`25MT-03P459Y 兴旺 入库单.xlsx` 抬头写的是 25MT-03P495Y，
       实发 14,394.32 = 本单睿贝采购金额）；
    2. **本单只有这一份附件、没写别的工厂、实发金额 = 睿贝采购金额**
       （`26MT-06H269勇恒 入库单.xlsx`：勇恒 + 4,326.00 = 26MT-06H629-YH）。

    认领结果会写进 issues，交人工复核。
    """
    contract = str(item.payload.get("contract_no") or "").strip()
    if contract and same_order(contract + " " + item.original, purchase_code):
        return f"单据抬头单号 {contract} 指向本单"
    if item.foreign_factory:
        return ""
    if len(foreign_list) == 1 and not kept_others:
        amount = signature_amount(item.payload)
        reference = erp_amount(purchase_code)
        if amount is not None and reference is not None and abs(amount - reference) <= Decimal("1"):
            return f"本单只有这一份附件，实发 {amount} 与睿贝采购金额 {reference} 一致"
    return ""


def resolve_total(payload: dict[str, Any]) -> tuple[Decimal | None, str]:
    """定出一张入库单的「实发合计」，并给出判定依据。

    单据形态有两种，必须分开处理：

    * **分批累计型**（各区块是本单不同批次的货，互不重叠）：
      区块金额之和 ≈ 订单金额（多出来的是吊绳/管帽等费用）→ 取各区块之和。
      例：26MT-05G019-D 的 A+B / C / D = 5,828,644.86（订单 5,792,906.86 + 费用）。
    * **包含型**（后加的区块是整单的一部分，例如「06H192A 发货部分」）：
      区块之和远大于订单金额，但其中有区块等于订单金额 → 取该区块。
      例：26MT-06H192 的「裸管」区块 1,360,707.93 = 订单金额。
    """
    values = block_values(payload)
    order = order_value(payload)
    if not values:
        # 没有实发区块时，如果解析结果里已经确认了实发合计（含"明细求和"兜底），优先用它
        declared = to_decimal((payload.get("settled_totals") or {}).get("amount"))
        if declared:
            return declared, "无实发区块，取实发合计"
        return order, "无实发区块，取订单金额"
    total = sum(values, Decimal(0))
    if order and order > 0:
        if abs(total - order) / order <= Decimal("0.05"):
            return total, "区块求和≈订单金额"
        for value in values:
            if abs(value - order) / order <= Decimal("0.005"):
                return max(value, order), "某区块=订单金额"
    return total, "区块求和（与订单金额差异大，待核）"


def drop_superseded(
    items: list[Item], purchase_code: str = ""
) -> tuple[list[Item], list[Item], list[str]]:
    """跨批次组再筛一遍「递进快照 / 重复上传」。

    同一张合并表常被反复另存，判据按证据强度排序：

    1. **实发区块是后一份的子集**（A 批 → A+B → A+B+C）→ 早期版本；
    2. **货物身份是后一份的子集** → 早期版本；
    3. **实发总额与后一份完全相同** → 同一笔钱，留靠后的那份；
    4. **货物身份完全相同但金额不同**（典型：后一份补了吊绳费/木箱费）：
       用该采购单的 ERP 采购金额当参照——去掉前一份后更接近 ERP 就判重复，
       否则按两批保留（可能是同一批货分两次发货），并标注待人工确认。
    5. **后一份的实发区块是前一份的真子集** → 后一份是同一批货的**部分单/补充件**
       （2026-09-24 加）。实锤：`26MT-05Q012-1` 的顺序是
       A(49,767) → B(A+B) → C(A+B+C=98,872) → 「A 入库单（少付补款558元）」(49,767)，
       第 4 条只比了「两份都留 vs 只留靠后那份」，把最贴近睿贝 93,417 的 C 丢掉、
       留下 A 批的 49,767，整单少了 49,105。区块包含关系才是硬证据。
    """
    kept: list[Item] = []
    dropped: list[Item] = []
    notes: list[str] = []
    ordered = sorted(items, key=lambda item: item.order)
    reference = erp_amount(purchase_code) if purchase_code else None
    to_drop: set[int] = set()
    for index, item in enumerate(ordered):
        if id(item) in to_drop:
            dropped.append(item)
            continue
        later = ordered[index + 1:]
        # 2026-09-24 修：**不属于本单的附件**（附件名写了别的单号，稍后会整份剔除）
        # 不能拿来当「后一份」用——否则会出现「本单唯一的附件被别单附件判成旧版」
        # 的连锁错误：25MT-08C619 的 A 批被 26MT-08C020 的合并结算单「更正」掉，
        # 而那张合并单自己又被判成别单附件，最后本单一份都没留下。
        later = [other for other in later if not other.other_order]
        batches = batch_set(item.payload)
        identity = line_identity(item.payload)
        blocks = block_amounts(item.payload)
        tokens = block_token_amounts(item.payload)
        goods = goods_key(item.payload)
        item_total, _item_how = resolve_total(item.payload)
        superseded = False
        for other in later:
            if id(other) in to_drop:
                continue
            other_batches = batch_set(other.payload)
            other_identity = line_identity(other.payload)
            other_blocks = block_amounts(other.payload)
            other_goods = goods_key(other.payload)
            other_tokens = block_token_amounts(other.payload)
            other_total, _other_how = resolve_total(other.payload)
            # 同一批次标签在两份附件里金额不同 = 后来更正过 → 前一份整份判旧版
            if (
                tokens
                and other_tokens
                and set(tokens) <= set(other_tokens)
                and any(
                    tokens[token] != other_tokens[token]
                    for token in tokens
                    if token in other_tokens
                )
            ):
                superseded = True
                notes.append(
                    f"{item.original} 的批次区块 {sorted(tokens)} 已被 {other.original} 更正，判为旧版"
                )
                break
            if blocks and other_blocks and blocks != other_blocks and not (blocks - other_blocks):
                superseded = True
                notes.append(
                    f"{item.original} 的实发区块金额是 {other.original} 的子集，判为早期版本"
                )
                break
            # 反向：后一份的区块是我方的真子集 → 它是部分单/补充件，不该顶掉累计终版
            if blocks and other_blocks and blocks != other_blocks and not (other_blocks - blocks):
                to_drop.add(id(other))
                notes.append(
                    f"{other.original} 的实发区块金额是 {item.original} 的子集，"
                    f"判为部分单/补充件（保留 {item.original}）"
                )
                continue
            if batches and other_batches and batches < other_batches:
                superseded = True
                notes.append(
                    f"{item.original} 是 {other.original} 的早期版本"
                    f"（批次 {sorted(batches)} ⊂ {sorted(other_batches)}）"
                )
                break
            if (
                goods
                and other_goods
                and goods < other_goods
                and names_compatible(item.payload, other.payload)
            ):
                amount = signature_amount(item.payload)
                other_amount = signature_amount(other.payload)
                extra = ""
                if amount is not None and other_amount is not None and amount != other_amount:
                    extra = "（含金额更正）"
                superseded = True
                notes.append(f"{item.original} 的货是 {other.original} 的子集，判为旧版{extra}")
                break
            if item_total is not None and other_total is not None and item_total == other_total:
                superseded = True
                notes.append(
                    f"{item.original} 与 {other.original} 的实发金额相同（{item_total}），"
                    "按重复上传处理，留靠后的那份"
                )
                break
            if (
                goods
                and other_goods
                and goods == other_goods
                and names_compatible(item.payload, other.payload)
            ):
                # 货一样、钱不一样：靠 ERP 采购金额当参照决定留几份
                amount = signature_amount(item.payload)
                other_amount = signature_amount(other.payload)
                if (
                    reference is not None
                    and item_total is not None
                    and other_total is not None
                ):
                    keep_both = abs((item_total + other_total) - reference)
                    keep_later = abs(other_total - reference)
                    if keep_later < keep_both:
                        superseded = True
                        notes.append(
                            f"{item.original} 与 {other.original} 货物相同、金额差 "
                            f"{amount} → {other_amount}；去掉后更接近 ERP 采购金额 "
                            f"{reference}，判为旧版"
                        )
                        break
                notes.append(
                    f"{item.original} 与 {other.original} 货物相同但金额不同"
                    f"（{amount} → {other_amount}），已按两批保留，待人工确认"
                )
        (dropped if superseded else kept).append(item)
    return kept, dropped, notes


def calibrate_duplicates(
    purchase_code: str, items: list[Item]
) -> tuple[list[Item], list[Item], list[str]]:
    """「同货物、同金额」的跨批次副本：用采购单金额当参照决定留几份。

    这类附件（例：26MT-08C406 的 A/B 两份都是 2110000）光看内容分不清
    是「同一批货重复上传」还是「真分两批发同样的货」。取采购单金额当参照：
    去掉副本后更接近采购单金额 → 判重复；否则两份都留，交人工确认。
    """
    reference = erp_amount(purchase_code)
    if reference is None or len(items) < 2:
        return items, [], []

    clusters: dict[tuple, list[Item]] = collections.defaultdict(list)
    for item in items:
        identity = tuple(sorted(line_identity(item.payload)))
        amount = signature_amount(item.payload)
        if identity and amount is not None:
            clusters[(identity, str(amount))].append(item)

    extra: list[Item] = []
    for (_identity, _amount), members in clusters.items():
        if len(members) > 1:
            extra.extend(sorted(members, key=lambda item: item.order)[:-1])
    if not extra:
        return items, [], []
    extra_ids = {id(item) for item in extra}

    def total(rows: list[Item]) -> Decimal:
        return sum((signature_amount(item.payload) or Decimal(0) for item in rows), Decimal(0))

    keep_all = total(items)
    dedup = [item for item in items if id(item) not in extra_ids]
    with_dedup = total(dedup)
    if abs(with_dedup - reference) < abs(keep_all - reference):
        notes = [
            f"{item.original} 与另一份内容金额完全一致，且去掉后更接近采购单金额 "
            f"{reference}，判为重复上传"
            for item in extra
        ]
        return dedup, extra, notes
    notes = [
        f"{item.original} 与另一份内容金额完全一致，但两份合计更接近采购单金额 {reference}，"
        "按两批保留，待人工确认"
        for item in extra
    ]
    return items, [], notes


def signature_amount(payload: dict[str, Any]) -> Decimal | None:
    total, _how = resolve_total(payload)
    if total is not None:
        return total
    for key in ("settled_totals", "order_totals"):
        totals = payload.get(key) or {}
        if isinstance(totals, dict) and totals.get("amount") is not None:
            return to_decimal(totals.get("amount"))
    return to_decimal(payload.get("settled_amount"))


def resolve_group(items: list[Item]) -> tuple[list[Item], list[Item], list[str]]:
    """同一批次内的多份附件 → (保留, 丢弃, 说明)。"""
    kept: list[Item] = []
    dropped: list[Item] = []
    notes: list[str] = []
    if not items:
        return kept, dropped, notes

    # 1. 同名（剥掉 (n) 后同名）**且内容相同**才算重复上传 → 只留最靠后的
    #
    #    2026-09-24 修：原来只比文件名，结果 `25MT-06N294 良凡 入库单.xlsx`（正式入库单 2240）
    #    被 `…入库单(1).xlsx`（补做入库单 160）顶掉——名字只差一个 (1)，内容完全不同的两张单。
    #    现在只有「金额相同」或「金额与货物指纹都相同」才当重复，否则两份都留、交后续逻辑判断。
    by_stem: dict[str, list[Item]] = {}
    for item in items:
        by_stem.setdefault(item.stem, []).append(item)
    unique: list[Item] = []
    for members in by_stem.values():
        keep: list[Item] = []
        for item in sorted(members, key=lambda entry: entry.order):
            amount = signature_amount(item.payload)
            goods = goods_key(item.payload)
            duplicate: Item | None = None
            for previous in keep:
                previous_amount = signature_amount(previous.payload)
                previous_goods = goods_key(previous.payload)
                same_amount = (
                    amount is not None and previous_amount is not None and amount == previous_amount
                )
                same_goods = bool(goods) and goods == previous_goods
                if same_amount and (same_goods or not goods):
                    duplicate = previous
                    break
            if duplicate is not None:
                keep.remove(duplicate)
                dropped.append(duplicate)
                notes.append(
                    f"重复上传：{duplicate.original} 与 {item.original}"
                    f"（同名且金额相同 {amount}），保留靠后的"
                )
            keep.append(item)
        unique.extend(keep)
    unique.sort(key=lambda item: item.order)

    # 2. 费用补充件单独保留
    fees = [item for item in unique if item.is_fee and not item.is_version]
    mains = [item for item in unique if item not in fees]
    kept.extend(fees)

    # 3. 正单内部定版
    if len(mains) <= 1:
        kept.extend(mains)
        return sorted(kept, key=lambda item: item.order), dropped, notes

    versioned = [item for item in mains if item.is_version]
    plain = [item for item in mains if not item.is_version]
    if plain and versioned:
        kept.extend(versioned[-1:])
        dropped.extend(plain + versioned[:-1])
        notes.append("存在「更新/最终」版，取最靠后的一份：" + versioned[-1].original)
        return sorted(kept, key=lambda item: item.order), dropped, notes

    # 4. 名字判不出 → 比明细（按货物身份，不含金额）
    ordered = sorted(mains, key=lambda item: item.order)
    last = ordered[-1]
    earlier = ordered[:-1]
    last_sig = line_identity(last.payload)
    superset = False
    for item in earlier:
        sig = line_identity(item.payload)
        if sig and sig <= last_sig:
            dropped.append(item)
            superset = True
        elif sig == last_sig:
            dropped.append(item)
    if superset:
        notes.append(f"明细为前版超集，取最靠后的一份：{last.original}")
        # 2026-09-24 加：被丢的那份**金额反而更大**时点名，交人工复核。
        # 实锤：`26MT-02N295-KLX` 的「补」单里有一行没写标签的 −1,180.76，
        # 把实发从 2,821.50 压到 1,640.74；而被丢的正单是 3,066.76（
        # 与睿贝 3,180.15 更贴）。这类只在报告里点名，不改金额。
        last_amount = signature_amount(last.payload)
        bigger = [
            item
            for item in dropped
            if signature_amount(item.payload) is not None
            and last_amount is not None
            and signature_amount(item.payload) > last_amount
        ]
        if bigger:
            notes.append(
                "被丢弃的版本金额更大，请复核是否真为旧版："
                + "；".join(
                    f"{item.original}（{signature_amount(item.payload)} > {last_amount}）"
                    for item in bigger
                )
            )
        kept.append(last)
        return sorted(kept, key=lambda item: item.order), dropped, notes

    # 明细互补 → 视为两批，都保留
    kept.extend(ordered)
    notes.append("同批次内多份明细互不包含，按不同批次全部保留，待人工确认")
    return sorted(kept, key=lambda item: item.order), dropped, notes


def select(purchase_code: str) -> dict[str, Any]:
    """一个采购单的完整选择结果。"""
    items = load_items(purchase_code)
    groups: dict[str, list[Item]] = collections.defaultdict(list)
    for item in items:
        groups[item.batch or "(无批次)"].append(item)

    result: dict[str, Any] = {
        "purchaseCode": purchase_code,
        "files": len(items),
        "groups": [],
        "kept": [],
        "dropped": [],
        "issues": [],
    }
    resolved: dict[str, tuple[list[Item], list[Item], list[str]]] = {}
    dropped_objs: list[Item] = []
    for batch, members in sorted(groups.items()):
        resolved[batch] = resolve_group(members)

    # 跨批次组再筛一遍「递进快照」：同一张合并表的 A / A+B / A+B+C 只留最后一份
    all_kept = [item for kept, _dropped, _notes in resolved.values() for item in kept]
    survivors, superseded, cross_notes = drop_superseded(all_kept, purchase_code)
    survivors, dup_dropped, dup_notes = calibrate_duplicates(purchase_code, survivors)
    survivor_ids = {id(item) for item in survivors}
    result["issues"].extend(cross_notes)
    result["issues"].extend(dup_notes)

    for batch, members in sorted(groups.items()):
        kept, dropped, notes = resolved[batch]
        dropped = list(dropped) + [item for item in kept if id(item) not in survivor_ids]
        kept = [item for item in kept if id(item) in survivor_ids]
        # 附件自己写的单号与目录不符（典型：附加单目录里放了主单的入库单）→ 不计入本单成本
        foreign = [item for item in kept if item.other_order]
        kept = [item for item in kept if not item.other_order]
        # 但文件名写错单号、证据又支持属于本单的，认领回来（2026-09-24）
        for item in list(foreign):
            reason = claim_reason(item, purchase_code, kept, foreign)
            if reason:
                foreign.remove(item)
                kept.append(item)
                result["issues"].append(
                    f"附件单号写错、证据支持属于本单，已认领（{reason}）："
                    f"{item.original}（本单 {purchase_code}）"
                )
        dropped = list(dropped) + foreign
        for item in foreign:
            result["issues"].append(
                f"附件写的是别的单号，未计入本单成本：{item.original}（本单 {purchase_code}）"
            )
        # 附件名点到别家工厂（本单供应商是另一家）→ 不计入本单成本
        foreign_factory = [item for item in kept if item.foreign_factory]
        kept = [item for item in kept if not item.foreign_factory]
        dropped = list(dropped) + foreign_factory
        for item in foreign_factory:
            result["issues"].append(
                f"附件写的是别家工厂，未计入本单成本：{item.original}（本单 {purchase_code}）"
            )
        result["groups"].append(
            {
                "batch": batch,
                "files": [item.original for item in members],
                "kept": [item.original for item in kept],
                "dropped": [item.original for item in dropped],
                "notes": notes,
            }
        )
        result["kept"].extend(
            {
                "batch": batch,
                "file": item.file,
                "original": item.original,
                "order": item.order,
                "amount": str(signature_amount(item.payload) or ""),
                "is_fee": item.is_fee,
                "digest": item.payload.get("_digest") or "",
                "how": item.payload.get("_how") or "",
                "blocks": [
                    {"label": str(block.get("label") or ""), "amount": str(block.get("amount") or "")}
                    for block in (item.payload.get("settled_blocks") or [])
                    if isinstance(block, dict)
                ],
            }
            for item in kept
        )
        result["dropped"].extend({"batch": batch, "original": item.original} for item in dropped)
        dropped_objs.extend(dropped)
        result["issues"].extend(notes)
        for item in members:
            if item.other_order:
                result["issues"].append(f"附件写的是别的单号：{item.original}")
            elif item.foreign_factory:
                result["issues"].append(f"附件写的是别家工厂：{item.original}")

    # 被丢掉的附件如果表尾写着「补款」这类金额、而保留件里没有同一笔，点名留痕：
    # 例 26MT-05Q012-1 的「A 入库单（少付补款558元）」——多数情况它是付款批注
    # （按口径不进成本），但确实可能该加回，交财务判。
    kept_note_amounts = {
        str(note.get("amount"))
        for item in survivors
        for note in (item.payload.get("tail_notes") or [])
        if isinstance(note, dict) and note.get("amount")
    }
    for item in dropped_objs:
        for note in item.payload.get("tail_notes") or []:
            if not isinstance(note, dict) or not note.get("amount"):
                continue
            if str(note.get("amount")) in kept_note_amounts:
                continue
            result["issues"].append(
                f"被丢附件 {item.original} 表尾批注「{str(note.get('label') or '')[:40]}」"
                f"{note.get('amount')} 未计入成本（付款/补款批注口径），如财务认定属成本请人工加回"
            )
    return result
