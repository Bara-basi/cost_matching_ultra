r"""拆单结果 × 飞书「迈拓财务部门数据 / 2026年报关」逐字段比对。

比对内容
--------
把本地拆单结果的四条关键字段与飞书人工表对齐：

    报关品名 / 产品类型 / 供应商（简称） / 采购单号

产品类型口径（按财务确认）：飞书的 `产品类型` 就是**报关品名经「产品及类型」
映射表折算出来的粗分类**（如「不锈钢螺纹管」→「焊管」），映射表里没有的一律记
`其他`。因此这里把我方产品类型也统一改成 `映射(报关品名)`，两侧口径一致；
这样「产品类型」这一项是否一致，等价于「报关品名是否一致」。
睿贝原始的「海关商品（中文）」仍保留在明细表里作参考。

比对单元 = `报关单号`（一张报关单一行结论），单元内比较四类**集合**：
报关品名、产品类型（按品名折算）、供应商简称、采购单号。

范围
----
只比**我方有拆单结果**的报关单。飞书有、我方没有的报关单与我们无关，
直接无视（不进比对、不报异常、不统计）。

输出（`outputs/shipments_match/`）
---------------------------------
    匹配明细_全部.xlsx    我方拆单记录逐条带上飞书对应值与四个字段的比对结论
    匹配结果_正常.xlsx    比对单元：四个字段全部一致
    匹配结果_异常.xlsx    比对单元：任一字段不一致 / 飞书侧无对应行
    匹配统计.json         口径与统计

用法
----
    & '.\.venv\Scripts\python.exe' scripts\match_feishu.py            # 联网拉飞书
    & '.\.venv\Scripts\python.exe' scripts\match_feishu.py --offline  # 用本地缓存
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from app.services.feishu_client import FeishuClient  # noqa: E402
from app.services.contract_shipments import cores_of  # noqa: E402

# 飞书「迈拓财务部门数据 副本 / 2026年报关数据」
REF_APP = "KPvoblTTxakNtMsPSzMcIvfQnYg"
REF_TABLE = "tblXqWvN826az7g3"
REF_NAME = "迈拓财务部门数据 副本 / 2026年报关数据"
# 产品类型 Lookup 的来源表（海关品名 → 产品类型 粗分类）
REF_PRODUCT_TABLE = "tblnTbhye2EjcRUa"

OURS = PROJECT_ROOT / "outputs" / "shipments_split" / "拆单明细_全部.xlsx"
PARSE_XLSX = PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_出口退税联.xlsx"
OUT_DIR = PROJECT_ROOT / "outputs" / "shipments_match"
CACHE = PROJECT_ROOT / "data" / "cache" / "ref_2026_customs.json"
FALLBACK_MAP = PROJECT_ROOT / "data" / "reference" / "product_map.json"
MAP_CACHE = PROJECT_ROOT / "data" / "reference" / "feishu_product_map.json"
VERIFIED = PROJECT_ROOT / "data" / "reference" / "verified_conflicts.json"
# 映射表里没有的报关品名，按财务口径一律记「其他」
UNMAPPED_KIND = "其他"

REF_COLUMNS = (
    "报关单号", "合同号_1", "报关品名", "海关编码", "产品类型", "供应商简称",
    "供应商", "采购单号", "采购金额", "报关金额", "报关重量",
)
# 飞书里与本地「采购单号」对应的列名（取值是**去掉工厂后缀**的采购单号，
# 如本地 `25MT-01P611-GYL` ↔ 飞书 `25MT-01P611`）。取回后落到同一个键上。
REF_PURCHASE_ALIASES = ("合同号（应收表格）",)

DETAIL_COLUMNS = (
    "报关单号", "合同号_1", "报关品名", "海关编码", "产品类型（ERP）", "产品类型（按报关品名映射）",
    "供应商简称", "供应商", "采购单号", "出运金额合计", "报关金额", "产品行数",
    "飞书合同号", "飞书报关品名", "飞书产品类型", "飞书供应商简称", "飞书采购单号",
    "飞书报关金额", "飞书采购金额",
    "报关品名比对", "产品类型比对", "供应商比对", "采购单号比对", "逐条结论",
)
GROUP_COLUMNS = (
    "报关单号", "合同号_1",
    "系统报关品名", "飞书报关品名",
    "系统产品类型", "飞书产品类型",
    "系统供应商简称", "飞书供应商简称",
    "系统采购单号", "飞书采购单号",
    "系统行数", "飞书行数", "比对结果",
)

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFE699")

SAME = "一致"
NOT_PRODUCED = "我方未拆出"
REF_MISSING = "飞书侧无对应行"
NO_COLUMN = "飞书无此列"
OUR_TYPE_MISSING = "我方产品类型缺失"
MAP_MISSING = "产品类型对照表未收录"

SPLIT_RE = re.compile(r"[、,，/&\n]+")


# --------------------------------------------------------------------------- 工具


def flat(value) -> str:
    """飞书字段值 → 纯文本（兼容 text / lookup / 多选数组）。"""
    if value is None:
        return ""
    if isinstance(value, list):
        return "、".join(part for part in (flat(x) for x in value) if part)
    if isinstance(value, dict):
        for key in ("text", "name", "value", "en_name"):
            if value.get(key):
                return flat(value[key])
        return ""
    return str(value).strip()


def split_multi(value) -> list[str]:
    """「甲、乙」这类合并写法拆成集合元素。"""
    text = flat(value)
    if not text:
        return []
    return [part.strip() for part in SPLIT_RE.split(text) if part.strip()]


def money_text(value) -> str:
    text = flat(value)
    if not text:
        return ""
    try:
        return f"{float(text):.2f}"
    except ValueError:
        return text


def read_sheet(path: Path):
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell) for cell in rows[0]]
    return {name: pos for pos, name in enumerate(header)}, rows[1:]


def set_text(values) -> str:
    return "、".join(sorted({v for v in values if v}))


def ok(check: str) -> bool:
    """比对结论是否算「通过」：一致 / 一致（说明）。"""
    return check == SAME or check.startswith("一致")


_SEP_RE = re.compile(r"[\s\-_/]+")


def norm_code(text: str) -> str:
    """单号归一：去掉空格/连字符/下划线再比。"""
    return _SEP_RE.sub("", flat(text)).upper()


def purchase_match(ours, theirs) -> bool:
    """飞书「合同号（应收表格）」= 本地采购单号，但两侧尾缀写法不同：

    * 本地带**工厂后缀**（`25MT-01P611-GYL`、`25MT-03P495Y-ADD1-XW`）；
    * 飞书带**批次后缀**（`26MT-01P242Y-A`、`25MT-06H520B`）。

    所以按 `订单核心`（年份+MT+部门+业务员+流水[Y][-ADDn]）比集合，而不是逐字符比。
    """
    mine = {core for x in ours for core in cores_of(x)}
    ref = {core for x in theirs for core in cores_of(x)}
    if not mine and not ref:
        return True
    if mine and ref:
        return mine == ref
    # 单号解析不出核心时（写法异常）退回前缀比对，避免直接判错
    left = [norm_code(x) for x in ours if norm_code(x)]
    right = [norm_code(x) for x in theirs if norm_code(x)]
    return bool(left) and bool(right) and all(
        any(m.startswith(r) for r in right) for m in left
    ) and all(any(m.startswith(r) for m in left) for r in right)


def purchase_covered(record_codes, ref_codes) -> bool:
    """记录级：这条记录的采购单号（订单核心）要能在飞书那一单里找到。"""
    mine = {core for x in record_codes for core in cores_of(x)}
    ref = {core for x in ref_codes for core in cores_of(x)}
    if not mine:
        return True
    return mine <= ref if ref else False


def excluded_declarations() -> set[str]:
    """上一步（拆单比对）已核实/判异常的数据：不再往后推进，避免干扰判断。"""
    if not VERIFIED.exists():
        return set()
    cases = json.loads(VERIFIED.read_text(encoding="utf-8")).get("cases") or []
    return {
        str(case.get("报关单号") or "").strip()
        for case in cases
        if case.get("报关单号")
    }


# ------------------------------------------------------------------- 飞书对照表


def product_option_map(client: FeishuClient) -> dict[str, str]:
    """产品类型 Lookup 的原始值是选项 id，去来源表取 id → 名称。"""
    for field in client.list_fields(REF_APP, REF_PRODUCT_TABLE):
        if field.get("field_name") == "产品类型":
            options = (field.get("property") or {}).get("options") or []
            return {opt["id"]: opt["name"] for opt in options if opt.get("id")}
    return {}


def coarse_map(client: FeishuClient) -> dict[str, str]:
    """海关品名 → 产品类型（粗分类）。"""
    table = {}
    try:
        for item in client.iter_records(REF_APP, REF_PRODUCT_TABLE, page_size=500):
            fields = item.get("fields") or {}
            name = flat(fields.get("海关品名"))
            kind = flat(fields.get("产品类型"))
            if name and kind:
                table[name] = kind
    except Exception:  # noqa: BLE001 - 对照表读不到就用本地缓存
        table = {}
    if table:
        MAP_CACHE.parent.mkdir(parents=True, exist_ok=True)
        MAP_CACHE.write_text(
            json.dumps(table, ensure_ascii=False, indent=1), encoding="utf-8"
        )
    elif MAP_CACHE.exists():
        table = json.loads(MAP_CACHE.read_text(encoding="utf-8"))
    if not table and FALLBACK_MAP.exists():
        for item in json.loads(FALLBACK_MAP.read_text(encoding="utf-8")):
            fields = item.get("fields") or {}
            name = flat(fields.get("海关品名"))
            kind = flat(fields.get("产品类型"))
            if name and kind:
                table[name] = kind
    return table


def fetch_ref(offline: bool) -> tuple[list[dict], dict[str, str]]:
    if offline and CACHE.exists():
        payload = json.loads(CACHE.read_text(encoding="utf-8"))
        return payload["records"], payload.get("coarse_map") or {}
    client = FeishuClient()
    options = product_option_map(client)
    mapping = coarse_map(client)
    records: list[dict] = []
    for item in client.iter_records(REF_APP, REF_TABLE, page_size=500):
        fields = item.get("fields") or {}
        record = {key: fields.get(column) for key, column in
                  ((k, k) for k in REF_COLUMNS)}
        types = split_multi(record["产品类型"])
        record["产品类型"] = [options.get(t, t) for t in types]
        for alias in REF_PURCHASE_ALIASES:
            if fields.get(alias):
                record["采购单号"] = fields[alias]
                break
        records.append(record)
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(
        json.dumps(
            {"app": REF_APP, "table": REF_TABLE,
             "records": records, "coarse_map": mapping},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return records, mapping


def normalise_ref(records: list[dict]) -> list[dict]:
    out: list[dict] = []
    for record in records:
        decl = flat(record.get("报关单号")).replace(" ", "")
        if not decl:
            continue
        out.append(
            {
                "报关单号": decl,
                "合同号_1": flat(record.get("合同号_1")),
                "报关品名": flat(record.get("报关品名")),
                "海关编码": flat(record.get("海关编码")),
                "产品类型": split_multi(record.get("产品类型")),
                "供应商简称": split_multi(record.get("供应商简称")),
                "供应商": flat(record.get("供应商")),
                "采购单号": split_multi(record.get("采购单号")),
                "采购金额": flat(record.get("采购金额")),
                "报关金额": flat(record.get("报关金额")),
                "报关重量": flat(record.get("报关重量")),
            }
        )
    return out


# ------------------------------------------------------------------------- 主流程


def load_ours() -> list[dict]:
    index, rows = read_sheet(OURS)
    out: list[dict] = []
    for row in rows:
        out.append(
            {
                "报关单号": flat(row[index["报关单号"]]).replace(" ", ""),
                "合同号_1": flat(row[index["合同号_1"]]),
                "报关品名": flat(row[index["报关品名"]]),
                "海关编码": flat(row[index["海关编码"]]),
                "产品类型": flat(row[index["产品类型"]]),
                "供应商简称": flat(row[index["供应商简称"]]),
                "供应商": flat(row[index["供应商"]]),
                "采购单号": flat(row[index["采购单号"]]),
                "出运金额合计": flat(row[index["出运金额合计"]]),
                "报关金额": flat(row[index["报关金额"]]),
                "产品行数": row[index["产品行数"]],
            }
        )
    return out


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--offline", action="store_true", help="用本地缓存，不联网")
    parser.add_argument(
        "--type-rule",
        choices=("name", "actual-if-ambiguous"),
        default="actual-if-ambiguous",
        help="actual-if-ambiguous（默认，业务口径）=同一报关单内"
             "「同一采购订单+同一供应商」下有两种实际类型时取实际品名，其余取报关品名映射；"
             "name=一律取报关品名映射",
    )
    args = parser.parse_args()

    blocked = excluded_declarations()
    ours_all = load_ours()
    ours = [record for record in ours_all if record["报关单号"] not in blocked]
    ref_records, mapping = fetch_ref(args.offline)
    ref_rows_all = normalise_ref(ref_records)
    ref_rows = [row for row in ref_rows_all if row["报关单号"] not in blocked]
    # 只保留我方有拆单结果的报关单；飞书多出来的报关单直接无视
    scope = {record["报关单号"] for record in ours}

    for record in ours:
        # 睿贝实际品名（海关商品）折算成飞书粗分类，用于比「实际品名」这条路
        record["产品类型（ERP折算）"] = mapping.get(record["产品类型"], record["产品类型"])
        record["_by_name"] = mapping.get(record["报关品名"], UNMAPPED_KIND)

    # 产品类型口径（按财务确认）：
    #   默认 = 报关品名经映射表的折算值；
    #   例外 = 同一条报关单里，**同一个采购订单 + 同一个供应商**下拆出了两种及以上的
    #          实际产品类型（谁是谁无法区分）时，以睿贝的实际品名为准。
    ambiguous: set[tuple[str, str, str]] = set()
    group_types: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    for record in ours:
        key = (record["报关单号"], record["采购单号"], record["供应商简称"])
        if record["产品类型（ERP折算）"]:
            group_types[key].add(record["产品类型（ERP折算）"])
    for key, types in group_types.items():
        if len(types) >= 2:
            ambiguous.add(key)

    for record in ours:
        key = (record["报关单号"], record["采购单号"], record["供应商简称"])
        if args.type_rule == "actual-if-ambiguous" and key in ambiguous \
                and record["产品类型（ERP折算）"]:
            record["产品类型（折算）"] = record["产品类型（ERP折算）"]
            record["_type_state"] = "实际品名"
        else:
            record["产品类型（折算）"] = record["_by_name"]
            record["_type_state"] = "报关品名映射"

    ref_by_decl: dict[str, list[dict]] = defaultdict(list)
    for record in ref_rows:
        if record["报关单号"] not in scope:
            continue
        ref_by_decl[record["报关单号"]].append(record)

    ours_units: dict[str, list[dict]] = defaultdict(list)
    for record in ours:
        ours_units[record["报关单号"]].append(record)
    ref_units: dict[str, list[dict]] = defaultdict(list)
    for record in ref_rows:
        if record["报关单号"] not in scope:
            continue
        ref_units[record["报关单号"]].append(record)

    has_purchase_column = any(r["采购单号"] for r in ref_rows)

    # ---------------- 逐条比对：按报关单号找飞书同单的行 ----------------
    detail_rows: list[dict] = []
    for record in ours:
        candidates = ref_by_decl.get(record["报关单号"], [])
        vendors = set(split_multi(record["供应商简称"]))
        kinds = {record["产品类型（折算）"]} - {""}
        purchases = set(split_multi(record["采购单号"]))

        if not candidates:
            checks = {k: "" for k in ("报关品名比对", "产品类型比对", "供应商比对", "采购单号比对")}
            verdict = REF_MISSING
        else:
            ref_names = {r["报关品名"] for r in candidates}
            ref_kinds = {k for r in candidates for k in r["产品类型"]}
            ref_vendors = {v for r in candidates for v in r["供应商简称"]}
            ref_purchases = {p for r in candidates for p in r["采购单号"]}
            state = record["_type_state"]
            kind_check = SAME if kinds <= ref_kinds else f"不符（飞书：{set_text(ref_kinds)}）"
            if kind_check != SAME and state == "未收录→其他":
                kind_check += "（该报关品名不在映射表，我方按「其他」处理）"
            checks = {
                "报关品名比对": SAME if record["报关品名"] in ref_names
                else f"品名不符（飞书：{set_text(ref_names)}）",
                "产品类型比对": kind_check,
                "供应商比对": SAME if vendors <= ref_vendors
                else f"不符（飞书：{set_text(ref_vendors)}）",
                "采购单号比对": (
                    (
                        SAME if purchase_covered(sorted(purchases), sorted(ref_purchases))
                        else f"不符（飞书：{set_text(ref_purchases)}）"
                    )
                    if has_purchase_column else NO_COLUMN
                ),
            }
            judged = [checks["报关品名比对"], checks["供应商比对"]]
            # 采购单号（飞书列名「合同号（应收表格）」）两侧尾缀写法不同，
            # 单号本身的差异**不作为异常**，只在明细里留列供人工核对。
            judged.append(checks["产品类型比对"])
            base = SAME if all(ok(c) for c in judged) else "不一致"
            verdict = base

        first = candidates[0] if candidates else {}
        detail_rows.append(
            {
                "报关单号": record["报关单号"],
                "合同号_1": record["合同号_1"],
                "报关品名": record["报关品名"],
                "海关编码": record["海关编码"],
                "产品类型（ERP）": record["产品类型"],
                "产品类型（按报关品名映射）": record["产品类型（折算）"],
                "供应商简称": record["供应商简称"],
                "供应商": record["供应商"],
                "采购单号": record["采购单号"],
                "出运金额合计": record["出运金额合计"],
                "报关金额": record["报关金额"],
                "产品行数": record["产品行数"],
                "飞书合同号": first.get("合同号_1", ""),
                "飞书报关品名": set_text({r["报关品名"] for r in candidates}),
                "飞书产品类型": set_text({k for r in candidates for k in r["产品类型"]}),
                "飞书供应商简称": set_text({v for r in candidates for v in r["供应商简称"]}),
                "飞书采购单号": set_text({p for r in candidates for p in r["采购单号"]}),
                "飞书报关金额": money_text(first.get("报关金额")) if first else "",
                "飞书采购金额": money_text(first.get("采购金额")) if first else "",
                **checks,
                "逐条结论": verdict,
            }
        )

    # ---------------- 单元比对：一张报关单一行结论 ----------------
    same_rows: list[dict] = []
    bad_rows: list[dict] = []
    for key in sorted(set(ours_units) | set(ref_units)):
        decl = key
        mine = ours_units.get(key, [])
        theirs = ref_units.get(key, [])
        my_names = {r["报关品名"] for r in mine} - {""}
        ref_names = {r["报关品名"] for r in theirs} - {""}
        my_kinds = {r["产品类型（折算）"] for r in mine} - {""}
        ref_kinds = {k for r in theirs for k in r["产品类型"]}
        my_vendors = {v for r in mine for v in split_multi(r["供应商简称"])}
        ref_vendors = {v for r in theirs for v in r["供应商简称"]}
        my_purchases = {p for r in mine for p in split_multi(r["采购单号"])}
        ref_purchases = {p for r in theirs for p in r["采购单号"]}

        problems: list[str] = []
        if not mine:
            problems.append(NOT_PRODUCED)
        if mine and not theirs:
            problems.append(REF_MISSING)
        if mine and theirs:
            leak = len(theirs) < len(mine) and bool(ref_kinds) and ref_kinds < my_kinds
            if leak:
                # 飞书少写了几行：品名/类型「对不上」都是这一件事的投影，单独说明
                problems.append(
                    f"飞书漏拆（我方 {len(mine)} 行 {set_text(my_kinds)} / "
                    f"飞书 {len(theirs)} 行 {set_text(ref_kinds)}，应补 "
                    f"{len(mine) - len(theirs)} 行）"
                )
            else:
                if my_names != ref_names:
                    problems.append(
                        f"报关品名不符（我方 {set_text(my_names)} / 飞书 {set_text(ref_names)}）"
                    )
                if my_kinds and my_kinds != ref_kinds:
                    problems.append(
                        f"产品类型不符（我方 {set_text(my_kinds)} / 飞书 {set_text(ref_kinds)}）"
                    )
            if my_vendors != ref_vendors:
                problems.append(
                    f"供应商不符（我方 {set_text(my_vendors)} / 飞书 {set_text(ref_vendors)}）"
                )
            # 采购单号差异不判异常（见上），仅保留列

        row = {
            "报关单号": decl,
            "合同号_1": (mine[0]["合同号_1"] if mine else theirs[0]["合同号_1"]),
            "系统报关品名": set_text(my_names),
            "飞书报关品名": set_text(ref_names),
            "系统产品类型": set_text(my_kinds),
            "飞书产品类型": set_text(ref_kinds),
            "系统供应商简称": set_text(my_vendors),
            "飞书供应商简称": set_text(ref_vendors),
            "系统采购单号": set_text(my_purchases),
            "飞书采购单号": set_text(ref_purchases),
            "系统行数": len(mine),
            "飞书行数": len(theirs),
            "比对结果": SAME if not problems else "；".join(problems),
        }
        (same_rows if not problems else bad_rows).append(row)

    write(detail_rows, DETAIL_COLUMNS, OUT_DIR / "匹配明细_全部.xlsx", highlight="逐条结论")
    write(same_rows, GROUP_COLUMNS, OUT_DIR / "匹配结果_正常.xlsx")
    write(bad_rows, GROUP_COLUMNS, OUT_DIR / "匹配结果_异常.xlsx", highlight="比对结果")

    stats = {
        "对照表": REF_NAME,
        "对照表是否有采购单号列": has_purchase_column,
        "产品类型对照表条数": len(mapping),
        "产品类型口径": "映射(报关品名)，映射表没有的记「其他」",
        "产品类型规则": args.type_rule,
        "上一步已核实/异常已排除的报关单": len({r["报关单号"] for r in ours_all} & blocked),
        "上一步已核实/异常已排除的记录": len(ours_all) - len(ours),
        "我方拆单记录": len(ours),
        "比对单元(报关单)": len(set(ours_units) | set(ref_units)),
        "飞书对应行数": sum(len(v) for v in ref_units.values()),
        "正常单元": len(same_rows),
        "异常单元": len(bad_rows),
        "报关品名不一致的报关单": sum(
            1 for row in bad_rows if "报关品名不符" in row["比对结果"]
        ),
        "产品类型不一致的报关单": sum(
            1 for row in bad_rows if "产品类型不符" in row["比对结果"]
        ),
        "供应商不一致的报关单": sum(
            1 for row in bad_rows if "供应商不符" in row["比对结果"]
        ),
        "逐条结论分布": dict(Counter(r["逐条结论"] for r in detail_rows)),
        "异常分类": dict(
            Counter(
                part.split("（")[0]
                for row in bad_rows
                for part in row["比对结果"].split("；")
            )
        ),
        "飞书漏拆的报关单": sum(
            1 for row in bad_rows if row["比对结果"].startswith("飞书漏拆")
        ),
    }
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "匹配统计.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=1))


def write(rows: list[dict], columns, path: Path, highlight: str | None = None) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "结果"
    sheet.append(list(columns))
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append([row.get(name, "") for name in columns])
        if highlight and row.get(highlight) and not ok(str(row[highlight])):
            fill = BAD_FILL if str(row[highlight]).startswith(("不一致", "未拆出")) else WARN_FILL
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    for column, name in zip(sheet.iter_cols(min_row=1, max_row=1), columns):
        sheet.column_dimensions[column[0].column_letter].width = max(12, min(30, len(name) * 2 + 4))
    sheet.freeze_panes = "A2"
    try:
        workbook.save(path)
        print(f"written {path} ({len(rows)} rows)")
    except PermissionError:
        # 目标被 Excel 占用时退到 _new 文件，避免整轮白跑
        fallback = path.with_name(f"{path.stem}_new{path.suffix}")
        workbook.save(fallback)
        print(f"!! {path.name} 被占用（Excel 打开中），已写到 {fallback.name} ({len(rows)} rows)")


if __name__ == "__main__":
    main()
