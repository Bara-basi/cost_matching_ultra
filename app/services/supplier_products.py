"""工厂 ↔ 产品类别对照（公司知识库）。

`data/reference/supplier_products.json`：供应商简称 → 该厂能生产的产品类别，
例如 `{"沪新": ["不锈钢无缝管", "不锈钢焊管"], "勇恒": ["不锈钢法兰", "不锈钢管件"]}`。

用途：睿贝里偶尔会出现「供应商编码/名称」与「产品」互相矛盾的数据
（如 514120260001117786：法兰的货挂在沪新名下，而沪新只做焊管/无缝管）。
这时按**产品反推实际工厂**，并把这条数据同时标成异常——
不能因为最后算出了金额就当正常数据用。

反推顺序：
1. 采购单号后缀（`26MT-03T094Y-YH` → `YH` → 勇恒，后缀↔工厂由采购单缓存统计得出）；
2. 产品类别唯一命中的工厂；
3. 都定不下来 → 返回失败，交由上层报异常。
"""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from app.services.erp_cache import CACHE_ROOT, read_jsonl

REF = Path(__file__).resolve().parents[2] / "data" / "reference" / "supplier_products.json"
# 「名录没收录、但 ERP 采购历史里反复出现」的判定阈值：
# 出现 2 次及以上就认为该厂确实做这个产品（名录不全，不代表异常）
HISTORY_MIN_HITS = 2

# 产品类别归一：睿贝「海关商品（中文）」与知识库写法不完全一致
PRODUCT_ALIASES: tuple[tuple[tuple[str, ...], str], ...] = (
    (("螺丝", "螺栓", "螺母", "垫圈", "紧固件"), "紧固件"),
    (("圆棒", "棒材", "棒"), "棒材"),
    (("型材",), "型材"),
    (("焊丝",), "焊丝"),
    (("丝",), "丝"),
    (("阀门",), "阀门"),
    (("焊管",), "焊管"),
    (("无缝管", "钢管"), "无缝管"),
    (("管件",), "管件"),
    (("法兰",), "法兰"),
    (("板材", "板"), "板"),
)


def product_key(name: str) -> str:
    """把产品写法归一到知识库口径（去掉「不锈钢」前缀，合并同义写法）。"""
    text = str(name or "").strip().replace("不锈钢", "").replace("钢铁制", "")
    text = text.replace(" ", "")
    if not text:
        return ""
    # 非不锈钢材质（镍合金/碳钢/钴合金…）不在名录覆盖范围内，直接判为不可归类
    if any(material in str(name or "") for material in ("镍", "碳钢", "钴", "钛", "铜", "塑料", "橡胶")):
        return ""
    for aliases, key in PRODUCT_ALIASES:
        if any(alias in text for alias in aliases):
            return key
    return text


def order_base(purchase_code: str) -> str:
    """采购单号去掉尾部工厂/批次后缀，得到「订单基数」，用于按订单去重。

    `26MT-05Q171-GYL` → `26MT-05Q171`；`26MT-03P200Y-HD` → `26MT-03P200Y`。
    """
    parts = str(purchase_code or "").strip().split("-")
    while len(parts) > 1:
        tail = parts[-1].strip()
        if tail and (tail.isascii() and tail.isalpha() or len(tail) <= 3) and len(tail) <= 5:
            parts.pop()
            continue
        break
    return "-".join(parts).upper()


def _dump(payload: dict) -> str:
    """按原文件风格写回（一家工厂一行、产品内联），方便人工比对 diff。"""
    lines = ["{"]
    items = list(payload.items())
    for index, (key, values) in enumerate(items):
        body = ", ".join(json.dumps(str(value), ensure_ascii=False) for value in values)
        tail = "," if index < len(items) - 1 else ""
        lines.append(f"  {json.dumps(key, ensure_ascii=False)}: [{body}]{tail}")
    lines.append("}")
    return "\n".join(lines) + "\n"


class SupplierProducts:
    """工厂 → 产品类别，以及采购单号后缀 ↔ 工厂。"""

    def __init__(self) -> None:
        self.by_short: dict[str, set[str]] = {}
        self._labels: dict[str, str] = {}
        self._suffix: dict[str, Counter] = {}
        self._history: dict[tuple[str, str], int] | None = None
        self._load()
        self._load_suffixes()

    def _load(self) -> None:
        if not REF.exists():
            return
        payload = json.loads(REF.read_text(encoding="utf-8"))
        for short, items in payload.items():
            keys = {product_key(item) for item in items if str(item or "").strip()}
            keys.discard("")
            if short and keys:
                self.by_short[str(short).strip()] = keys
            for item in items:
                key = product_key(item)
                if key:
                    self._labels.setdefault(key, str(item).strip())

    def canonical_label(self, product: str) -> str:
        """产品类别在名录里的规范写法（如 key=焊管 → 「不锈钢焊管」）。"""
        key = product_key(product)
        if not key:
            return ""
        return self._labels.get(key) or str(product or "").strip()

    def learn(self, short: str, product: str) -> str:
        """把「名录没收录、但 ERP 采购历史反复出现」的产品**写回映射表**。

        只有同时满足下面三条才写：
        1. 该产品类别在名录里是已知类别（只是这家工厂没收录）；
        2. 这家工厂的名录里确实没有这个类别；
        3. 采购历史里这家工厂确实多次做过这个类别（≥ HISTORY_MIN_HITS 张订单）。

        返回写入的规范写法；没写入返回空字符串。
        """
        short = str(short or "").strip()
        key = product_key(product)
        if not short or not key:
            return ""
        # 工厂名对齐：报关/ERP 里常写简称（希蒙），名录里是全称（希蒙雷斯）
        target = short
        if target not in self.by_short:
            matches = [
                name
                for name in self.by_short
                if name.startswith(target) or target.startswith(name)
            ]
            if len(matches) != 1:
                return ""
            target = matches[0]
        if key in self.products_of(target) or not self.covered(product):
            return ""
        if self.history_hits(short, product) < HISTORY_MIN_HITS:
            return ""
        label = self.canonical_label(product)
        if not label or not REF.exists():
            return ""
        payload = json.loads(REF.read_text(encoding="utf-8"))
        items = payload.setdefault(target, [])
        if label in items:
            return ""
        items.append(label)
        REF.write_text(_dump(payload), encoding="utf-8")
        self.by_short[target] = self.products_of(target) | {key}
        return f"{target} += {label}" if target != short else label

    def _load_suffixes(self) -> None:
        """从采购单缓存统计「采购单号后缀 → 工厂简称」（如 YH→勇恒）。"""
        from app.services.supplier_names import SupplierNames

        names = SupplierNames()
        for row in read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl"):
            code = str(row.get("purchase_code") or "").strip()
            supplier = str(row.get("supplierName") or "").strip()
            if not code or not supplier:
                continue
            short = names.short(supplier)
            if not short:
                continue
            for token in code.split("-")[::-1]:
                token = token.strip()
                if not token:
                    continue
                if token.isascii() and token.isalpha() and len(token) <= 5:
                    self._suffix.setdefault(token.upper(), Counter())[short] += 1
                break

    # ---------- 查询 ----------

    def products_of(self, short: str) -> set[str]:
        return self.by_short.get(str(short or "").strip(), set())

    def factory_of_purchase(self, purchase_code: str) -> str:
        """采购单号 → 工厂简称（看尾缀，如 `...-GYL` → 溪流）。"""
        for token in str(purchase_code or "").split("-")[::-1]:
            token = token.strip()
            if not token:
                continue
            counter = self._suffix.get(token.upper())
            if counter:
                return counter.most_common(1)[0][0]
            if token.isascii() and token.isalpha() and len(token) <= 5:
                return ""
            break
        return ""

    def history_counts(self) -> dict[tuple[str, str], int]:
        """(工厂简称, 产品类别) → 出现在**多少张不同订单**里。

        工厂归属优先看采购单号尾缀（供应商名称字段经常是空的），
        这份统计用于判断「名录没收录」到底是名录不全还是睿贝挂错厂。

        按不同订单去重（而不是按行数）：同一张挂错厂的订单里的 4 行不能互相作证。
        """
        if self._history is not None:
            return self._history
        from app.services.shipment_index import load_lines
        from app.services.supplier_names import SupplierNames

        names = SupplierNames()
        orders: dict[tuple[str, str], set[str]] = {}
        for line in load_lines():
            purchase = str(line.get("purchase_code") or "")
            short = names.short(str(line.get("supplier") or "")) or self.factory_of_purchase(
                purchase
            )
            key = product_key(str(line.get("customs_name") or ""))
            if short and key:
                orders.setdefault((short, key), set()).add(order_base(purchase))
        self._history = {pair: len(bases) for pair, bases in orders.items()}
        return self._history

    def history_hits(self, short: str, product: str) -> int:
        key = product_key(product)
        if not key:
            return 0
        return self.history_counts().get((str(short or "").strip(), key), 0)

    def verdict(self, short: str, product: str) -> tuple[bool, str]:
        """判断「某工厂 × 某产品」是否可信，并给出依据。"""
        short = str(short or "").strip()
        keys = self.products_of(short)
        key = product_key(product)
        if not short or not key:
            return True, "工厂或产品无法归类，不做判断"
        if not self.covered(product):
            return True, f"名录未覆盖产品类别「{key}」"
        if not keys:
            return True, f"工厂「{short}」不在名录里"
        if key in keys:
            return True, f"名录内：{short} 生产「{key}」"
        hits = self.history_hits(short, product)
        if hits >= HISTORY_MIN_HITS:
            return True, (
                f"名录未收录，但 ERP 采购历史中 {short} 的「{key}」出现在 {hits} 张订单里，"
                "判为正常（名录不全）"
            )
        return False, (
            f"{short} 在名录里只做 {'、'.join(sorted(keys))}，"
            f"且采购历史里「{key}」只出现在 {hits} 张订单里，判为异常"
        )

    def factories_for(self, product: str) -> set[str]:
        key = product_key(product)
        if not key:
            return set()
        return {short for short, keys in self.by_short.items() if key in keys}

    def covered(self, product: str) -> bool:
        """知识库里是否有工厂生产这个产品类别。

        名录只覆盖不锈钢系列（镍合金、软管、控制管线等都不在里面），
        没覆盖到的类别不做判断，否则会把正常数据全标成异常。
        """
        return bool(self.factories_for(product))

    def consistent(self, short: str, product: str) -> bool:
        """该工厂是否做这个产品。工厂或产品不在名录里时不做判断（返回 True）。"""
        return self.verdict(short, product)[0]

    def infer_factory(self, product: str, purchase_code: str = "") -> tuple[str, str]:
        """按产品（可带采购单号后缀）反推实际工厂。

        返回 (工厂简称, 依据)；推不出来返回 ("", 原因)。
        """
        key = product_key(product)
        if not key:
            return "", f"产品「{product}」无法归类，推不出工厂"
        candidates = self.factories_for(product)
        if not candidates:
            return "", f"知识库未覆盖产品「{key}」（名录只有不锈钢系列），无法判断"
        if len(candidates) == 1:
            return next(iter(candidates)), f"产品「{key}」在知识库中只对应 {next(iter(candidates))}"
        suffix = ""
        for token in str(purchase_code or "").split("-")[::-1]:
            token = token.strip()
            if not token:
                continue
            if token.isascii() and token.isalpha() and len(token) <= 5:
                suffix = token.upper()
            break
        if suffix:
            counter = self._suffix.get(suffix)
            if counter:
                matched = candidates & set(counter)
                if len(matched) == 1:
                    return next(iter(matched)), f"采购单号后缀 {suffix} 指向 {next(iter(matched))}"
        return "", (
            f"产品「{key}」对应多家工厂（{'、'.join(sorted(candidates))}），"
            "采购单号后缀也无法区分"
        )
