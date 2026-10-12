"""工作台流水线：把「上传件 / 逐单输入 / 飞书同步」变成一轮拆单 + 成本匹配。

所有数据每次都重新跑（不读历史结果），输入是 PDF / Excel / 一维商品行，
输出是 `outputs/runs/<job>/` 下的一轮结果（拆单明细 + 财务核对行）。
"""
from __future__ import annotations

import json
import sys
import uuid
from collections import Counter
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402

from app.services.parser import parse_declaration  # noqa: E402
from app.services.parser.excel_writer import build_rows as build_parse_rows  # noqa: E402
from app.services.parser.excel_writer import write_workbook  # noqa: E402
from app.services.review import build_rows as build_review_rows  # noqa: E402
from app.services.scope import out_of_scope_reason  # noqa: E402

RUNS_DIR = PROJECT_ROOT / "outputs" / "runs"

# 逐单核验/同步输入的最小列（与 PDF 解析结果的列名一致，便于复用同一条拆单链路）
DECLARATION_COLUMNS = (
    "来源文件", "合同号_1", "报关单号", "商品序号", "报关品名", "海关编码",
    "报关重量", "申报数量", "申报单位", "总价", "币种", "产品类型",
)


def new_run_dir(job_id: str | None = None) -> Path:
    path = RUNS_DIR / (job_id or uuid.uuid4().hex[:12])
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_declarations(rows: list[dict], path: Path) -> int:
    """把一维商品行写成「解析结果」格式的 xlsx，供拆单直接消费。

    写之前先把「财务手工补回的拆单结果」合并回一条报关行：
    同一 `报关单号 + 合同号 + 报关品名 + 商品序号` 的历史拆分行先合并。
    真实不同商品序号保留为独立输入；不能仅因品名相同就把真实商品并成历史拆分行。
    只作用于飞书/手工输入（PDF 解析结果不走这里），所以真的多商品报关不会被并。
    """
    merged_rows, merged_count = merge_declaration_rows(rows)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "结果"
    sheet.append(list(DECLARATION_COLUMNS))
    kept = 0
    for row in merged_rows:
        contract = str(row.get("合同号_1") or row.get("合同号（应收表格）") or "").strip()
        if not contract or out_of_scope_reason(contract):
            continue
        sheet.append([str(row.get(column) or "") for column in DECLARATION_COLUMNS])
        kept += 1
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    if merged_count:
        print(f"[合并报关行] 把 {merged_count} 条手工拆分行并回了报关行（同报关单号+同品名）", flush=True)
    return kept


def _sum_text(*values) -> str:
    """把若干金额/重量文本相加，输出去掉尾零的字符串（空值当 0）。"""
    from decimal import Decimal, InvalidOperation

    total = Decimal(0)
    for value in values:
        try:
            total += Decimal(str(value).replace(",", "").strip() or "0")
        except (InvalidOperation, AttributeError):
            continue
    text = f"{total:.4f}".rstrip("0").rstrip(".")
    return text or "0"


def _dec_text(value) -> "Decimal":
    from decimal import Decimal, InvalidOperation

    try:
        return Decimal(str(value).replace(",", "").strip() or "0")
    except (InvalidOperation, AttributeError):
        return Decimal(0)


def _fmt_number(value) -> str:
    text = f"{value:.4f}".rstrip("0").rstrip(".")
    return text or "0"


def merge_group_totals(rows: list[dict]) -> tuple[str, str]:
    """同 `报关单号 + 合同号 + 报关品名` 的多行合计 → (报关金额, 报关重量)。

    飞书里同一个报关行的多行有两种含义，必须分开处理，否则会把金额算成两倍：

    - **完全重复的副本**（金额、重量都一样）→ 只算一次；
    - **原始行 + 财务拆出来的子行并存**：某一行金额 ≈ 其余去重行之和 → 取该行
      （原始行没被删掉时，子行是它的拆分，不能再加一遍）；
    - 其余情况（原始行已被拆分结果替换，只能看到子行）→ 各行相加。

    2026-10-08 案例：`26MT-05X216&217A` 每个报关行在表里都有两份完全相同的记录，
    相加后报关金额正好是睿贝出运金额的两倍，拆单直接报「子集和无解」。
    """
    from decimal import Decimal

    seen: set[tuple] = set()
    unique: list[tuple] = []
    for row in rows:
        value = row.get("报关金额")
        amount = _dec_text(row.get("总价") if value in (None, "") else value)
        weight = _dec_text(row.get("报关重量"))
        key = (amount, weight)
        if key in seen:
            continue
        seen.add(key)
        unique.append((amount, weight))
    if not unique:
        return "0", "0"
    if len(unique) > 2:
        amounts = [amount for amount, _weight in unique]
        biggest = max(amounts)
        others = sum(amounts, Decimal(0)) - biggest
        if abs(biggest - others) <= Decimal("0.05"):
            weight = next(weight for amount, weight in unique if amount == biggest)
            return f"{biggest:.2f}", _fmt_number(weight)
    amount_total = sum((amount for amount, _weight in unique), Decimal(0))
    weight_total = sum((weight for _amount, weight in unique), Decimal(0))
    return f"{amount_total:.2f}", _fmt_number(weight_total)


def merge_declaration_rows(rows: list[dict]) -> tuple[list[dict], int]:
    """同 `报关单号 + 合同号 + 报关品名` 的多行 → 合并成一条。

    金额/重量按 `merge_group_totals`（去重 + 取超集 + 相加）合并，
    避免「原始行与其拆分结果并存」或「行被复制」时把金额算成两倍。
    """
    key_of = _declaration_key
    slots: list[tuple[str, object]] = []
    index: dict[tuple[str, ...], list[dict]] = {}
    out: list[dict] = []
    merged = 0
    for row in rows:
        key = key_of(row)
        if key is None:
            slots.append(("single", row))
            continue
        if key not in index:
            index[key] = []
            slots.append(("group", key))
        index[key].append(row)
    for kind, payload in slots:
        if kind == "single":
            out.append(dict(payload))
            continue
        members = index[payload]
        target = dict(members[0])
        if len(members) > 1:
            amount, weight = merge_group_totals(members)
            target["总价"] = amount
            target["报关重量"] = weight
            quantities = [row for row in members if str(row.get("申报数量") or "").strip()]
            if len(quantities) > 1:
                target["申报数量"] = _fmt_number(
                    sum((_dec_text(row.get("申报数量")) for row in quantities), _dec_text(0))
                )
            target["商品序号"] = ""   # 合并后不再是原来那一条拆分记录
            merged += len(members) - 1
        out.append(target)
    return out, merged


def _declaration_key(row: dict) -> tuple[str, str, str, str] | None:
    declaration = str(row.get("报关单号") or "").strip()
    if not declaration:
        return None
    return (declaration,
            str(row.get("合同号_1") or row.get("合同号（应收表格）") or "").strip(),
            str(row.get("报关品名") or "").strip(),
            str(row.get("商品序号") or "").strip())


def parse_pdfs(pdf_paths: list[Path], out_xlsx: Path) -> dict:
    """解析报关单 PDF → 出口退税联 xlsx（与 `scripts/parse_customs.py` 同口径）。"""
    rows: list[dict] = []
    failed: list[dict] = []
    warnings: Counter = Counter()
    for path in pdf_paths:
        try:
            parsed = parse_declaration(path)
        except Exception as exc:  # noqa: BLE001
            failed.append({"文件": path.name, "原因": str(exc)})
            continue
        warnings.update(parsed.warnings)
        if parsed.header.sheet_type != "出口退税联":
            continue
        for row in build_parse_rows(parsed):
            if out_of_scope_reason(str(row.get("合同号_1") or "")):
                continue
            rows.append(row)
    if rows:
        write_workbook(rows, out_xlsx)
    return {"rows": len(rows), "failed": failed, "warnings": dict(warnings)}


def run_split(parse_xlsx: Path, out_dir: Path, only: str = "") -> dict:
    """调用拆单引擎（脚本的服务入口），返回统计。"""
    from scripts.run_split_knapsack import run as split_run

    return split_run(parse_xlsx=parse_xlsx, out_dir=out_dir, only=only) or {}


def _read_sheet(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True)
    try:
        rows = list(workbook.active.iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        return []
    header = [str(cell or "") for cell in rows[0]]
    return [dict(zip(header, values)) for values in rows[1:]]


def _num(value) -> "Decimal":
    from decimal import Decimal, InvalidOperation

    try:
        return Decimal(str(value).replace(",", "").strip() or "0")
    except (InvalidOperation, AttributeError):
        return Decimal(0)


def run_round(parse_xlsx: Path, out_dir: Path, only: str = "") -> dict:
    """与离线脚本一致，只对真实报关输入拆单一次，不生成 ERP 补充报关行。"""
    stats = run_split(parse_xlsx, out_dir, only)
    return {"stats": stats, "supplemented": 0}


class BatchCalculationError(ValueError):
    """本批无法核算的业务原因；完整明细随任务状态返回界面。"""

    def __init__(self, issues: list[dict], *, reason: str = "incomplete_input"):
        self.issues = issues
        self.reason = reason
        details = "；".join(str(item["原因"])[:110] for item in issues[:2])
        super().__init__(f"本批无法核算：{details}" + (f"；另有 {len(issues)-2} 项，请查看核验明细。" if len(issues) > 2 else ""))


def validate_input_completeness(parse_xlsx: Path) -> None:
    """按对应出运单核验真实输入，联合合同及共享同一出运单的写法一起比较。"""
    from decimal import Decimal, InvalidOperation
    from app.services.shipment_detail import load_shipment_detail, shipment_invoices_for_contract, strip_pi

    # 保留已确认的「报关金额 vs 出运总金额」容差；它与采购成本对账的人民币容差不同。
    tolerance = Decimal("5")

    def amount(value):
        try:
            result = Decimal(str(value).replace(",", "").strip())
            return result if result.is_finite() else None
        except (InvalidOperation, ValueError):
            return None

    issues: list[dict] = []
    by_contract: dict[str, dict] = {}
    for row in _read_sheet(parse_xlsx):
        contract = str(row.get("合同号_1") or "").strip()
        if not contract:
            continue
        entry = by_contract.setdefault(contract, {"declared": Decimal(0), "currencies": set()})
        value = row.get("总价")
        declared = amount(row.get("报关金额") if value in (None, "") else value)
        if declared is None or declared <= 0:
            issues.append({"合同号": contract, "报关单号": str(row.get("报关单号") or ""),
                           "原因": f"{contract}：报关金额缺失或不是有效正数，本批无法核算。"})
            continue
        entry["declared"] += declared
        currency = str(row.get("币种") or "").strip().upper()
        if currency:
            entry["currencies"].add({"美元": "USD", "欧元": "EUR", "人民币": "RMB", "CNY": "RMB"}.get(currency, currency))

    if issues:
        raise BatchCalculationError(issues)
    amounts: dict[str, Decimal] = {}
    currencies: dict[str, str] = {}
    groups: list[dict] = []
    for contract, entry in by_contract.items():
        invoices = shipment_invoices_for_contract(contract) or [contract]
        invoice_keys = set()
        valid = True
        for invoice in invoices:
            key = strip_pi(invoice)
            invoice_keys.add(key)
            if key in amounts:
                continue
            payload = next((value for candidate in (invoice, key, f"PI-{key}")
                            if (value := load_shipment_detail(candidate))), None)
            if not payload:
                issues.append({"合同号": contract, "出运单": invoice,
                               "原因": f"{contract}：缺少睿贝出运明细，无法确认本批报关单是否齐全，请先同步或核对出运单。"})
                valid = False
                continue
            base = payload.get("baseInfo") or {}
            value = base.get("出运总金额")
            if value in (None, ""):
                products = payload.get("productList") or []
                product_amounts = [amount(row.get("出运金额")) for row in products]
                expense_amounts = [amount(row.get("金额")) for row in payload.get("expenseList") or []]
                if not products or any(value is None for value in product_amounts + expense_amounts):
                    issues.append({"合同号": contract, "出运单": invoice,
                                   "原因": f"{contract}：睿贝出运金额不完整，无法核验本批输入。"})
                    valid = False
                    continue
                value = sum(product_amounts + expense_amounts, Decimal(0))
            parsed_amount = amount(value)
            if parsed_amount is None:
                issues.append({"合同号": contract, "出运单": invoice,
                               "原因": f"{contract}：睿贝出运总金额无效，无法核验本批输入。"})
                valid = False
                continue
            amounts[key] = parsed_amount
            currency = str(base.get("币种") or "").strip().upper()
            currencies[key] = {"美元": "USD", "欧元": "EUR", "人民币": "RMB", "CNY": "RMB"}.get(currency, currency)
        if not valid:
            continue
        combined = {"contracts": {contract}, "invoices": invoice_keys,
                    "declared": entry["declared"], "currencies": set(entry["currencies"])}
        remaining = []
        for group in groups:
            if group["invoices"] & combined["invoices"]:
                combined["invoices"].update(group["invoices"])
                combined["contracts"].update(group["contracts"])
                combined["declared"] += group["declared"]
                combined["currencies"].update(group["currencies"])
            else:
                remaining.append(group)
        groups = remaining + [combined]

    for group in groups:
        label = " / ".join(sorted(group["contracts"]))
        expected_currencies = {currencies[key] for key in group["invoices"] if currencies.get(key)}
        if len(group["currencies"] | expected_currencies) > 1:
            issues.append({"合同号": label, "原因": f"{label}：报关币种与出运币种不一致，无法直接核算本批。"})
            continue
        expected = sum((amounts[key] for key in group["invoices"]), Decimal(0))
        gap = group["declared"] - expected
        if abs(gap) > tolerance:
            direction = "少" if gap < 0 else "多"
            issues.append({"合同号": label, "报关金额": str(group["declared"]),
                           "出运金额": str(expected), "差额": str(gap),
                           "原因": f"{label}：报关合计 {group['declared']:.2f}，出运合计 {expected:.2f}，{direction} {abs(gap):.2f}；请提供该批完整报关商品行。"})
    if issues:
        raise BatchCalculationError(issues)


def validate_split_completeness(out_dir: Path) -> None:
    """拆单失败或真实货物未认领时拒绝整批；保留离线孪生行补记成本规则。"""
    from app.services.review import _friendly_local_note

    issues = []
    exception_path = out_dir / "拆单_本地异常.xlsx"
    if exception_path.exists():
        for row in _read_sheet(exception_path):
            contract = str(row.get("合同号_1") or "")
            declaration = str(row.get("报关单号") or "")
            _kind, note = _friendly_local_note(
                str(row.get("说明") or ""), str(row.get("类型") or ""))
            issues.append({"合同号": contract, "报关单号": declaration,
                           "原因": f"{contract} / {declaration}：{note}"})
    unclaimed_path = out_dir / "拆单_未认领产品行.xlsx"
    if unclaimed_path.exists():
        for row in _read_sheet(unclaimed_path):
            if str(row.get("处理") or "").startswith("已按孪生行规则补记成本") or _num(row.get("出运金额")) <= 0:
                continue
            contract = str(row.get("合同号_1") or "")
            issues.append({"合同号": contract, "采购单号": str(row.get("采购单号") or ""),
                           "原因": f"{contract}：仍有出运产品行未对应到报关行，请补齐该批报关单；本次不会从睿贝生成报关行。"})
    if issues:
        raise BatchCalculationError(issues, reason="split_failed")


def review_rows(split_xlsx: Path, amount_source: dict[str, str] | None = None) -> list[dict]:
    return build_review_rows(split_xlsx, amount_source)


def write_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
