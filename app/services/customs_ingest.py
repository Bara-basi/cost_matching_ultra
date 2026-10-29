"""报关 PDF 到飞书 AI 副本的最小闭环。

流程：PDF -> 当前解析器 -> 拆单 -> 成本匹配 -> 成本来源行 -> AI 展示行。

AI 表里的「供应商简称」「产品类型」「采购金额」是 Lookup，不能直接写入。
因此成本结果先写入同一 Base 的「2026年报关数据」，再由 AI 表公式回显。
"""
from __future__ import annotations

import json
import re
import time
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from app.services import pipeline
from app.services.feishu_client import FeishuClient, FeishuError, get_config
from app.services.parser.output_columns import OUTPUT_COLUMNS

CHINA_TZ = timezone(timedelta(hours=8))
DEFAULT_COST_TABLE_ID = "tbl2HV320SzNyXU4"
STATUS_FIELDS = (
    ("拆单状态", 1),
    ("成本匹配状态", 1),
)
ORIGINAL_AI_COST_FIELD = "采购金额AI"
PROTECTED_PURCHASE_FIELD = "采购金额"
AI_RESULT_FIELD_SPECS = (
    ("合同号_1", "合同号AI", 1),
    ("报关金额", "报关金额AI", 2),
    ("报关重量", "报关重量AI", 2),
    ("供应商", "供应商", 1),
    ("产品类型", "产品类型AI", 1),
    ("分摊依据", "采购金额分摊依据AI", 1),
    ("采购金额", "采购金额AI", 2),
)
ORIGINAL_PARSED_FIELDS = (
    "合同号_1",
    "合同号（应收表格）",
    "销售方式",
    "出口日期",
    "报关品名",
    "海关编码",
    "运输方式",
    "运费",
    "保费",
    "报关金额",
    "币种",
    "报关重量",
    "报关重量单位",
    "国家",
)
REQUIRED_TASK_FIELDS = (
    ("报关单号", "报关单号", 1),
    ("pdf原件", "pdf原件", 17),
    *(("", name, field_type) for name, field_type in STATUS_FIELDS),
)


def cost_table_id() -> str:
    return get_config("MT_FINANCE_AI_COPY_DATA_TABLE_ID", DEFAULT_COST_TABLE_ID)


def custom_cost_ai_table() -> tuple[str, str]:
    """报关核算结果目标表：(app_token, table_id)。

    旧版 ``MT_FINANCE_AI_*`` 继续指向原报关记录表，只作为 PDF 来源。
    所有解析、拆单和成本匹配结果统一写入独立的「2026年成本匹配 AI」表。
    """
    app = get_config("MT_FINANCE_COST_AI_APP_TOKEN", "") or get_config(
        "MT_FINANCE_AI_APP_TOKEN"
    )
    table = get_config("MT_FINANCE_COST_AI_TABLE_ID", "") or get_config(
        "MT_FINANCE_AI_TABLE_ID"
    )
    return app, table


LEGACY_COST_FIELDS = (
    "自动成本_采购金额",
    "自动成本_供应商简称",
    "自动成本_产品类型",
    "自动成本_采购订单号",
    "自动成本_分摊依据",
    "自动成本_匹配状态",
)


def decimal_value(value: Any) -> Decimal | None:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        result = Decimal(text)
    except InvalidOperation:
        return None
    return result if result.is_finite() else None


def fee_number(value: Any) -> Decimal:
    """报关单 `/0/`、`USD/2800/总价` 这类运费字段取金额段。"""
    text = str(value or "").strip()
    parts = [part.strip() for part in text.split("/")]
    number = decimal_value(parts[1] if len(parts) > 1 else text)
    return number or Decimal(0)


def date_millis(value: Any) -> int:
    if isinstance(value, datetime):
        moment = value.replace(tzinfo=CHINA_TZ) if value.tzinfo is None else value.astimezone(CHINA_TZ)
        return int(moment.timestamp() * 1000)
    raw = str(value or "").strip()
    match = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", raw)
    if not match:
        raise ValueError(f"无法转换出口日期：{value!r}")
    moment = datetime(int(match[1]), int(match[2]), int(match[3]), tzinfo=CHINA_TZ)
    return int(moment.timestamp() * 1000)


def read_sheet(path: Path) -> list[dict]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(workbook.active.iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        return []
    header = [str(value or "") for value in rows[0]]
    return [dict(zip(header, values)) for values in rows[1:]]


def assert_output_contract(parse_xlsx: Path) -> None:
    workbook = load_workbook(parse_xlsx, read_only=True)
    try:
        header = [str(value or "") for value in next(workbook.active.iter_rows(values_only=True))]
    finally:
        workbook.close()
    if tuple(header) != OUTPUT_COLUMNS:
        raise ValueError("本轮解析结果的字段与既有报关解析输出不一致")


def parse_formal_pdfs(pdf_dir: Path, work_dir: Path) -> Path:
    pdfs = sorted(pdf_dir.glob("*.pdf"))
    if not pdfs:
        raise ValueError(f"没有找到正式报关单 PDF：{pdf_dir}")
    output = work_dir / "parse" / "报关单解析结果_出口退税联.xlsx"
    stats = pipeline.parse_pdfs(pdfs, output)
    if stats.get("failed"):
        raise ValueError(f"解析失败 {len(stats['failed'])} 份")
    assert_output_contract(output)
    return output


def candidate_declarations(rows: list[dict], limit: int = 3) -> list[str]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        if row.get("报关单号"):
            grouped[str(row["报关单号"])].append(row)
    eligible = [
        declaration
        for declaration, group in grouped.items()
        if len(group) > 1
        and all(row.get("异常类型") == "正常" for row in group)
        and all(decimal_value(row.get("采购金额")) is not None for row in group)
        and all(row.get("供应商简称") and row.get("采购订单号") for row in group)
    ]
    eligible.sort(key=lambda value: (-len(grouped[value]), value))
    return eligible[:limit]


def _parse_rows(parse_xlsx: Path) -> tuple[dict[str, list[dict]], dict[str, str]]:
    rows = read_sheet(parse_xlsx)
    by_declaration: dict[str, list[dict]] = defaultdict(list)
    files: dict[str, str] = {}
    for row in rows:
        declaration = str(row.get("报关单号") or "").strip()
        if not declaration:
            continue
        by_declaration[declaration].append(row)
        source = str(row.get("来源文件") or "").strip()
        if source:
            files[declaration] = source
    return dict(by_declaration), files


def _allocated_fee(
    rows: list[dict], metadata: dict, field: str
) -> list[Decimal]:
    total = fee_number(metadata.get(field))
    declared = [decimal_value(row.get("报关金额")) or Decimal(0) for row in rows]
    denominator = sum(declared, Decimal(0))
    out: list[Decimal] = []
    allocated = Decimal(0)
    for index, amount in enumerate(declared):
        if index == len(declared) - 1:
            share = total - allocated
        else:
            share = (
                (total * amount / denominator).quantize(Decimal("0.01"))
                if denominator
                else Decimal(0)
            )
            allocated += share
        out.append(share)
    return out


def build_records(
    selected: list[str],
    result_rows: list[dict],
    parse_xlsx: Path,
    pdf_dir: Path,
    marker: str,
) -> tuple[list[dict], list[dict], dict[str, Path]]:
    """构造来源表和 AI 表的记录，尚未上传附件或写飞书。"""
    metadata, source_names = _parse_rows(parse_xlsx)
    grouped: dict[str, list[dict]] = defaultdict(list)
    for row in result_rows:
        declaration = str(row.get("报关单号") or "")
        if declaration in selected:
            grouped[declaration].append(row)
    source_rows: list[dict] = []
    ai_rows: list[dict] = []
    pdfs: dict[str, Path] = {}
    for declaration in selected:
        rows = grouped.get(declaration) or []
        if not rows:
            raise ValueError(f"{declaration} 没有拆单结果")
        source_name = source_names.get(declaration, "")
        pdf = pdf_dir / source_name
        if not pdf.is_file():
            raise ValueError(f"{declaration} 的原始 PDF 不存在：{pdf}")
        pdfs[declaration] = pdf
        meta = (metadata.get(declaration) or [{}])[0]
        sales_mode = (
            "预录"
            if str(meta.get("单据类型") or "").startswith("预录")
            else "出口"
        )
        freight = _allocated_fee(rows, meta, "运费")
        insurance = _allocated_fee(rows, meta, "保费")
        for index, row in enumerate(rows):
            purchase = str(row.get("采购订单号") or "").strip()
            supplier = str(row.get("供应商简称") or "").strip()
            product_type = str(row.get("产品类型") or "").strip()
            common = {
                "合同号_1": str(row.get("合同号_1") or ""),
                "合同号（应收表格）": str(row.get("合同号（应收表格）") or ""),
                "销售方式": sales_mode,
                "报关单号": declaration,
                "出口日期": date_millis(meta.get("出口日期")),
                "报关品名": str(row.get("报关品名") or ""),
                "海关编码": str(meta.get("海关编码") or ""),
                "运输方式": str(meta.get("成交方式") or ""),
                "运费": float(freight[index]),
                "保费": float(insurance[index]),
                "报关金额": float(decimal_value(row.get("报关金额")) or 0),
                "币种": str(meta.get("币种") or ""),
                "报关重量": float(decimal_value(row.get("报关重量")) or 0),
                "报关重量单位": str(meta.get("报关重量单位") or ""),
                "国家": str(meta.get("目的国") or ""),
                "备注": marker,
            }
            source_rows.append(
                {
                    **common,
                    "供应商简称": supplier,
                    "供应商": str(row.get("供应商名称") or ""),
                    "产品类型": [product_type] if product_type else [],
                    "采购金额": float(decimal_value(row.get("采购金额")) or 0),
                }
            )
            issue_type = str(row.get("异常类型") or "正常")
            issue_detail = str(row.get("异常明细") or "").strip()
            split_errors = {
                "找不到出运单", "拆单失败", "产品类型缺失", "缺报关单",
                "睿贝缺数据", "睿贝数据异常（供应商挂错厂）",
            }
            if issue_type == "正常":
                split_status = cost_status = "正常"
            elif issue_type in split_errors:
                split_status = f"异常：{issue_detail or issue_type}"
                cost_status = "未执行"
            else:
                split_status = "正常"
                cost_status = f"异常：{issue_detail or issue_type}"
            ai_rows.append(
                {
                    **common,
                    "供应商简称": supplier,
                    "供应商": str(row.get("供应商名称") or ""),
                    "产品类型": product_type,
                    "采购订单号": purchase,
                    "分摊依据": str(row.get("分摊依据") or ""),
                    "采购金额": float(decimal_value(row.get("采购金额")) or 0),
                    "拆单状态": split_status,
                    "成本匹配状态": cost_status,
                    "备注": marker,
                }
            )
    return source_rows, ai_rows, pdfs


@dataclass
class WriteSummary:
    marker: str
    declarations: list[str]
    sourceRecords: int
    aiRecords: int
    uploadedFiles: int
    workDir: str
    mode: str


def write_direct_to_ai_copy(
    *,
    declarations: list[str],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 30,
) -> WriteSummary:
    """只操作 AI 副本表：采购金额为数字字段，PDF 每单只附第一条。"""
    app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
    table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    client = FeishuClient()
    existing_fields = {field.get("field_name"): field for field in client.list_fields(app, table)}
    for name, field_type in STATUS_FIELDS:
        if name not in existing_fields:
            client.create_field(app, table, name, field_type)

    existing = _records_for_declarations(client, app, table, declarations)
    upload_tokens: dict[str, str] = {}
    updates: list[dict] = []
    creates: list[dict] = []
    used: set[str] = set()
    for declaration in declarations:
        desired = [
            dict(row) for row in ai_rows
            if str(row.get("报关单号") or "") == declaration
        ]
        existing_records = list(existing.get(declaration) or [])
        primary = next(
            (
                record for record in existing_records
                if (record.get("fields") or {}).get("pdf原件")
            ),
            existing_records[0] if existing_records else None,
        )
        existing_token = ""
        if primary:
            attachments = (primary.get("fields") or {}).get("pdf原件") or []
            if attachments:
                existing_token = str((attachments[0] or {}).get("file_token") or "")
        token = existing_token or client.upload_bitable_file(app, pdfs[declaration])
        upload_tokens[declaration] = token

        for index, desired_row in enumerate(desired):
            amount = decimal_value(desired_row.get("报关金额"))
            weight = decimal_value(desired_row.get("报关重量"))
            matches = [
                record for record in existing_records
                if record.get("record_id") not in used
                and decimal_value(_record_value(record, "报关金额")) == amount
                and decimal_value(_record_value(record, "报关重量")) == weight
            ]
            record = matches[0] if len(matches) == 1 else None
            desired_row["pdf原件"] = [{"file_token": token}] if index == 0 else []
            if record:
                used.add(record["record_id"])
                updates.append({"record_id": record["record_id"], "fields": desired_row})
            else:
                creates.append(desired_row)
        extras = [
            record for record in existing_records
            if record.get("record_id") not in used
        ]
        if extras:
            raise RuntimeError(
                f"{declaration} 已有 {len(extras)} 条无法对应的 AI 记录，拒绝重复写入"
            )

    if updates:
        client.batch_update_records(app, table, updates)
    if creates:
        client.batch_create_records(app, table, creates)

    deadline = time.monotonic() + wait_seconds
    expected = Counter(str(row["报关单号"]) for row in ai_rows)
    expected_amounts = {
        declaration: sorted(
            (
                f"{decimal_value(row.get('采购金额')).normalize():f}"
                if decimal_value(row.get("采购金额")) is not None
                else ""
            )
            for row in ai_rows
            if str(row.get("报关单号") or "") == declaration
        )
        for declaration in declarations
    }
    problem = ""
    while time.monotonic() < deadline:
        found = _records_for_declarations(client, app, table, declarations)
        counts_ok = {key: len(value) for key, value in found.items()} == dict(expected)
        if not counts_ok:
            problem = "AI 记录数与目标不一致"
            time.sleep(2)
            continue
        problems = []
        for declaration, records in found.items():
            costs = sorted(
                (
                    f"{decimal_value(_record_value(record, '采购金额')).normalize():f}"
                    if decimal_value(_record_value(record, "采购金额")) is not None
                    else ""
                )
                for record in records
            )
            attachments = [
                len((record.get("fields") or {}).get("pdf原件") or [])
                for record in records
            ]
            if costs != expected_amounts[declaration]:
                problems.append(f"{declaration}: 采购金额不一致")
            if sum(attachments) != 1:
                problems.append(f"{declaration}: PDF 附件不是 1 个")
        if not problems:
            summary = WriteSummary(
                marker=marker,
                declarations=declarations,
                sourceRecords=0,
                aiRecords=len(ai_rows),
                uploadedFiles=len(set(upload_tokens.values())),
                workDir=str(work_dir),
                mode="direct-ai-copy",
            )
            (work_dir / "direct_write_result.json").write_text(
                json.dumps(
                    {
                        **asdict(summary),
                        "fileTokens": upload_tokens,
                        "recordIds": [
                            str(record.get("record_id") or "")
                            for records in found.values()
                            for record in records
                        ],
                    },
                    ensure_ascii=False,
                    indent=1,
                ),
                encoding="utf-8",
            )
            return summary
        problem = "；".join(problems)
        time.sleep(2)
    raise RuntimeError(problem or "直接写入副本表后验证超时")


def _records_for_declarations(client: FeishuClient, app: str, table: str, declarations: list[str]) -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = {value: [] for value in declarations}
    wanted = set(declarations)
    records: list[dict] = []
    for start in range(0, len(declarations), 10):
        chunk = declarations[start : start + 10]
        records.extend(
            client.search_records(
                app,
                table,
                {
                    "conjunction": "or",
                    "conditions": [
                        {
                            "field_name": "报关单号",
                            "operator": "is",
                            "value": [declaration],
                        }
                        for declaration in chunk
                    ],
                },
            )
        )
    seen: set[str] = set()
    for record in records:
        record_id = str(record.get("record_id") or "")
        if record_id and record_id in seen:
            continue
        if record_id:
            seen.add(record_id)
        fields = record.get("fields") or {}
        declaration = _first_text(fields.get("报关单号"))
        if declaration in wanted:
            out[declaration].append(record)
    return out


def _first_text(value: Any) -> str:
    if isinstance(value, list):
        return " ".join(
            str(item.get("text") or item.get("name") or "")
            if isinstance(item, dict) else str(item)
            for item in value
        ).strip()
    return str(value or "").strip()


def write_to_copy(
    *,
    declarations: list[str],
    source_rows: list[dict],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 30,
) -> WriteSummary:
    app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
    ai_table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    client = FeishuClient()
    source_table = cost_table_id()
    existing_source = _records_for_declarations(client, app, source_table, declarations)
    existing_ai = _records_for_declarations(client, app, ai_table, declarations)
    duplicates = {
        "source": {key: len(value) for key, value in existing_source.items() if value},
        "ai": {key: len(value) for key, value in existing_ai.items() if value},
    }
    if any(duplicates.values()):
        raise ValueError(f"目标表已存在测试单号，拒绝重复写入：{duplicates}")

    tokens: dict[str, str] = {}
    for declaration in declarations:
        tokens[declaration] = client.upload_bitable_file(app, pdfs[declaration])
    for row in source_rows:
        token = tokens[str(row["报关单号"])]
        row["附件"] = [{"file_token": token}]
    for row in ai_rows:
        token = tokens[str(row["报关单号"])]
        row["pdf原件"] = [{"file_token": token}]
    (work_dir / "upload_result.json").write_text(
        json.dumps(tokens, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    # 先形成成本来源，再让 AI 表的 Lookup 取数。
    source_written = False
    try:
        source_result = client.batch_create_records(app, source_table, source_rows)
        source_written = True
    except FeishuError as exc:
        if exc.code != 800074911 and "ConnectorSyncTableActionForbidden" not in str(exc):
            raise
        # 同步表只能由上游同步任务写入；当前测试改为在 AI 行备注保存完整成本结果。
        source_result = {"records": []}
    time.sleep(3)
    ai_result = client.batch_create_records(app, ai_table, ai_rows)

    deadline = time.monotonic() + wait_seconds
    expected = Counter(str(row["报关单号"]) for row in ai_rows)
    expected_pairs = {
        declaration: sorted(
            (
                str(decimal_value(row.get("报关金额"))),
                str(decimal_value(row.get("采购金额"))),
            )
            for row in source_rows
            if str(row.get("报关单号") or "") == declaration
        )
        for declaration in declarations
    }
    expected_source = {
        declaration: sorted(
            (
                str(decimal_value(row.get("报关金额"))),
                str(decimal_value(row.get("采购金额"))),
                str(row.get("供应商简称") or ""),
            )
            for row in source_rows
            if str(row.get("报关单号") or "") == declaration
        )
        for declaration in declarations
    }
    problem = ""
    while time.monotonic() < deadline:
        found_source = _records_for_declarations(client, app, source_table, declarations)
        found_ai = _records_for_declarations(client, app, ai_table, declarations)
        source_counts_ok = (
            not source_written
            or {key: len(value) for key, value in found_source.items()} == dict(expected)
        )
        if not source_counts_ok or {
            key: len(value) for key, value in found_ai.items()
        } != dict(expected):
            problem = "写入后的记录数与计划不一致"
            time.sleep(2)
            continue
        missing_cost = []
        missing_pdf = []
        amount_mismatch = []
        for declaration, records in found_ai.items():
            if source_written:
                actual_pairs = sorted(
                    (
                        str(decimal_value(fields.get("报关金额"))),
                        str(decimal_value(_first_text(fields.get("采购金额")))),
                    )
                    for fields in [record.get("fields") or {} for record in records]
                )
                if actual_pairs != expected_pairs[declaration]:
                    amount_mismatch.append(declaration)
            for record in records:
                fields = record.get("fields") or {}
                if source_written and not _first_text(fields.get("采购金额")):
                    missing_cost.append(declaration)
                if not fields.get("pdf原件"):
                    missing_pdf.append(declaration)
        if source_written:
            for declaration, records in found_source.items():
                actual_pairs = sorted(
                    (
                        str(decimal_value(fields.get("报关金额"))),
                        str(decimal_value(fields.get("采购金额"))),
                        str(fields.get("供应商简称") or ""),
                    )
                    for fields in [record.get("fields") or {} for record in records]
                )
                if actual_pairs != expected_source[declaration]:
                    amount_mismatch.append(f"source:{declaration}")
        else:
            for declaration, records in found_ai.items():
                expected_notes = {
                    str(row.get("备注") or "")
                    for row in ai_rows
                    if str(row.get("报关单号") or "") == declaration
                }
                actual_notes = {
                    str((record.get("fields") or {}).get("备注") or "")
                    for record in records
                }
                if actual_notes != expected_notes:
                    amount_mismatch.append(f"remark:{declaration}")
        if not missing_cost and not missing_pdf and not amount_mismatch:
            summary = WriteSummary(
                marker=marker,
                declarations=declarations,
                sourceRecords=len(source_result.get("records") or []),
                aiRecords=len(ai_result.get("records") or []),
                uploadedFiles=len(tokens),
                workDir=str(work_dir),
                mode="source+ai" if source_written else "ai-only-sync-source",
            )
            (work_dir / "write_result.json").write_text(
                json.dumps(
                    {
                        **asdict(summary),
                        "sourceRecordIds": [
                            item.get("record_id") for item in source_result.get("records") or []
                        ],
                        "aiRecordIds": [
                            item.get("record_id") for item in ai_result.get("records") or []
                        ],
                        "fileTokens": tokens,
                    },
                    ensure_ascii=False,
                    indent=1,
                ),
                encoding="utf-8",
            )
            return summary
        problem = (
            f"写入校验未完成：采购金额缺失 {len(missing_cost)} 条，"
            f"附件缺失 {len(missing_pdf)} 条，金额/供应商不一致 {len(amount_mismatch)} 组"
        )
        time.sleep(3)
    raise RuntimeError(problem or "写入后验证超时")


def verify_existing_ai_copy(work_dir: Path) -> WriteSummary:
    """校验已经写入的 AI 副本记录；不重复创建记录或上传附件。"""
    plan_path = work_dir / "write_plan.json"
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    declarations = list(plan["declarations"])
    expected_rows = list(plan["aiRecords"])
    client = FeishuClient()
    app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
    table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    found = _records_for_declarations(client, app, table, declarations)
    expected_counts = Counter(str(row["报关单号"]) for row in expected_rows)
    actual_counts = {key: len(value) for key, value in found.items()}
    if actual_counts != dict(expected_counts):
        raise RuntimeError(f"AI 记录数不一致：expected={dict(expected_counts)} actual={actual_counts}")

    problems = []
    record_ids: list[str] = []
    file_tokens = set()
    compare_fields = (
        "合同号_1", "合同号（应收表格）", "报关品名", "海关编码",
        "报关金额", "报关重量", "采购金额", "备注",
    )

    def canonical(field: str, value: Any) -> str:
        if field in ("报关金额", "报关重量", "采购金额"):
            number = decimal_value(value)
            return f"{number.normalize():f}" if number is not None else ""
        return str(value or "")

    for declaration in declarations:
        expected = sorted(
            tuple(canonical(field, row.get(field)) for field in compare_fields)
            for row in expected_rows
            if str(row.get("报关单号") or "") == declaration
        )
        actual = sorted(
            tuple(
                canonical(field, record.get("fields", {}).get(field))
                for field in compare_fields
            )
            for record in found[declaration]
        )
        if actual != expected:
            problems.append(f"{declaration}: 字段内容不一致")
        attachment_count = 0
        for record in found[declaration]:
            record_ids.append(str(record.get("record_id") or ""))
            attachments = (record.get("fields") or {}).get("pdf原件") or []
            attachment_count += len(attachments)
            if attachments:
                file_tokens.update(str(item.get("file_token") or "") for item in attachments)
        if attachment_count != 1:
            problems.append(f"{declaration}: PDF 附件数量应为 1，实际 {attachment_count}")
    if problems:
        raise RuntimeError("；".join(problems))
    summary = WriteSummary(
        marker=str(plan["marker"]),
        declarations=declarations,
        sourceRecords=0,
        aiRecords=sum(actual_counts.values()),
        uploadedFiles=len(file_tokens),
        workDir=str(work_dir),
        mode="direct-ai-copy",
    )
    (work_dir / "write_result.json").write_text(
        json.dumps(
            {
                **asdict(summary),
                "aiRecordIds": record_ids,
                "fileTokens": sorted(file_tokens),
                "sourceTableTouched": False,
            },
            ensure_ascii=False,
            indent=1,
        ),
        encoding="utf-8",
    )
    return summary


def _record_value(record: dict, field: str) -> Any:
    return (record.get("fields") or {}).get(field)


def _write_to_original_legacy(
    *,
    declarations: list[str],
    source_rows: list[dict],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 60,
) -> WriteSummary:
    """写原表成本来源和 AI 表；同名业务键更新，缺失拆分行新增。"""
    app = get_config("MT_FINANCE_AI_APP_TOKEN")
    ai_table = get_config("MT_FINANCE_AI_TABLE_ID")
    source_table = get_config("MT_FINANCE_AI_DATA_TABLE_ID")
    client = FeishuClient()

    source_fields = client.list_fields(app, source_table)
    synced_business = [
        str(field.get("field_name") or "")
        for field in source_fields
        if field.get("is_synced")
        and str(field.get("field_name") or "") in {
            "合同号（应收表格）", "报关单号", "报关品名", "供应商简称",
            "产品类型", "采购金额", "报关金额", "报关重量",
        }
    ]
    if synced_business:
        raise RuntimeError(
            f"来源表 {source_table} 是同步表，业务字段不可写："
            + "、".join(synced_business)
            + "。请写入其上游同步源，或把 AI 表的成本字段改为可写数字字段。"
        )

    existing_source = _records_for_declarations(client, app, source_table, declarations)
    existing_ai = _records_for_declarations(client, app, ai_table, declarations)
    fields = {field.get("field_name"): field for field in client.list_fields(app, ai_table)}
    for name, field_type in STATUS_FIELDS:
        if name not in fields:
            client.create_field(app, ai_table, name, field_type)

    tokens = {
        declaration: client.upload_bitable_file(app, path)
        for declaration, path in pdfs.items()
    }
    (work_dir / "original_upload_result.json").write_text(
        json.dumps(tokens, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    # 来源表：先按采购金额精确配对，再按同一报关单和供应商复用旧汇总行。
    source_updates: list[dict] = []
    source_creates: list[dict] = []

    def original_source_fields(row: dict) -> dict:
        values = dict(row)
        product_type = values.get("产品类型")
        if isinstance(product_type, list):
            values["产品类型"] = "、".join(str(item) for item in product_type if item)
        return values

    for declaration in declarations:
        desired = [
            original_source_fields(row) for row in source_rows
            if str(row.get("报关单号") or "") == declaration
        ]
        existing = list(existing_source.get(declaration) or [])
        pairs: list[tuple[dict, dict]] = []
        for desired_row in list(desired):
            target_amount = decimal_value(desired_row.get("采购金额"))
            matched = next(
                (
                    record for record in existing
                    if decimal_value(_record_value(record, "采购金额")) == target_amount
                ),
                None,
            )
            if matched:
                pairs.append((desired_row, matched))
                desired.remove(desired_row)
                existing.remove(matched)
        for desired_row in list(desired):
            supplier = str(desired_row.get("供应商简称") or "")
            matched = next(
                (
                    record for record in existing
                    if _first_text(_record_value(record, "供应商简称")) == supplier
                ),
                None,
            )
            if matched:
                pairs.append((desired_row, matched))
                desired.remove(desired_row)
                existing.remove(matched)
        if existing:
            raise RuntimeError(
                f"{declaration} 来源表存在 {len(existing)} 条无法对应的旧行，拒绝自动删除"
            )
        for desired_row, record in pairs:
            source_updates.append(
                {"record_id": record["record_id"], "fields": desired_row}
            )
        source_creates.extend(desired)

    client.batch_update_records(app, source_table, source_updates) if source_updates else None
    client.batch_create_records(app, source_table, source_creates) if source_creates else None

    # AI 表：每个报关单保留一条原行作为第一条拆分行，其余新增。
    ai_updates: list[dict] = []
    ai_creates: list[dict] = []
    for declaration in declarations:
        desired = [
            dict(row) for row in ai_rows
            if str(row.get("报关单号") or "") == declaration
        ]
        existing = list(existing_ai.get(declaration) or [])
        if len(existing) > len(desired):
            raise RuntimeError(
                f"{declaration} AI 表已有 {len(existing)} 行，多于目标 {len(desired)} 行"
            )
        for index, desired_row in enumerate(desired):
            desired_row["pdf原件"] = (
                [{"file_token": tokens[declaration]}] if index == 0 else []
            )
            if index < len(existing):
                ai_updates.append(
                    {"record_id": existing[index]["record_id"], "fields": desired_row}
                )
            else:
                ai_creates.append(desired_row)
    client.batch_update_records(app, ai_table, ai_updates) if ai_updates else None
    client.batch_create_records(app, ai_table, ai_creates) if ai_creates else None

    deadline = time.monotonic() + wait_seconds
    problems: list[str] = []
    while time.monotonic() < deadline:
        found_source = _records_for_declarations(client, app, source_table, declarations)
        found_ai = _records_for_declarations(client, app, ai_table, declarations)
        expected = Counter(str(row["报关单号"]) for row in ai_rows)
        source_ok = {key: len(value) for key, value in found_source.items()} == dict(expected)
        ai_ok = {key: len(value) for key, value in found_ai.items()} == dict(expected)
        if source_ok and ai_ok:
            missing_lookup = []
            attachment_counts = Counter()
            for declaration, records in found_ai.items():
                for record in records:
                    fields = record.get("fields") or {}
                    if not _first_text(fields.get("采购金额")):
                        missing_lookup.append(declaration)
                    if fields.get("pdf原件"):
                        attachment_counts[declaration] += 1
            if not missing_lookup and all(attachment_counts[d] == 1 for d in declarations):
                summary = WriteSummary(
                    marker=marker,
                    declarations=declarations,
                    sourceRecords=len(source_rows),
                    aiRecords=len(ai_rows),
                    uploadedFiles=len(tokens),
                    workDir=str(work_dir),
                    mode="original-source+ai",
                )
                (work_dir / "original_write_result.json").write_text(
                    json.dumps(
                        {
                            **asdict(summary),
                            "sourceUpdates": len(source_updates),
                            "sourceCreates": len(source_creates),
                            "aiUpdates": len(ai_updates),
                            "aiCreates": len(ai_creates),
                            "fileTokens": tokens,
                        },
                        ensure_ascii=False,
                        indent=1,
                    ),
                    encoding="utf-8",
                )
                return summary
            problems = [
                f"Lookup 未回填 {len(missing_lookup)} 条",
                f"附件数量不符 {dict(attachment_counts)}",
            ]
        else:
            problems = ["来源表或 AI 表记录数尚未达到目标"]
        time.sleep(3)
    raise RuntimeError("；".join(filter(None, problems)) or "原表写入验证超时")


def ensure_ai_result_fields(
    client: FeishuClient,
    app: str,
    table: str,
) -> dict[str, dict[str, Any]]:
    """确保任务写回所需的字段存在；字段位置由用户手动调整。"""
    existing = {
        str(field.get("field_name") or ""): field
        for field in client.list_fields(app, table)
    }
    field_specs = (*REQUIRED_TASK_FIELDS, *AI_RESULT_FIELD_SPECS)
    for _, name, field_type in field_specs:
        field = existing.get(name)
        if field is None:
            field = client.create_field(app, table, name, field_type)
            existing[name] = field
            time.sleep(0.12)
        elif int(field.get("type") or 0) != field_type:
            raise RuntimeError(
                f"{name} 已存在，但类型为 {field.get('type')}，期望 {field_type}"
            )
    # 创建接口的 data 外壳在不同飞书版本间不一致；以字段列表回读结果为准。
    confirmed = {
        str(field.get("field_name") or ""): field
        for field in client.list_fields(app, table)
    }
    return {
        name: confirmed[name]
        for _, name, _ in AI_RESULT_FIELD_SPECS
    }


def original_ai_fields(row: dict) -> dict:
    """把拆单、成本结果和报关解析事实写入正式表字段。"""
    values: dict[str, Any] = {"报关单号": str(row.get("报关单号") or "")}
    for field in ORIGINAL_PARSED_FIELDS:
        if field not in row:
            continue
        raw = row.get(field)
        if isinstance(raw, list):
            raw = "、".join(
                str(item.get("text") or item.get("name") or item)
                if isinstance(item, dict)
                else str(item)
                for item in raw
                if item
            )
        values[field] = raw
    for source, target, field_type in AI_RESULT_FIELD_SPECS:
        raw = row.get(source)
        if field_type == 2:
            number = decimal_value(raw)
            values[target] = float(number) if number is not None else None
            continue
        if field_type == 5:
            if raw not in (None, ""):
                values[target] = int(raw)
            continue
        if isinstance(raw, list):
            raw = "、".join(str(item) for item in raw if item)
        text = str(raw or "").strip()
        values[target] = text
    values["拆单状态"] = str(row.get("拆单状态") or "正常")
    values["成本匹配状态"] = str(row.get("成本匹配状态") or "正常")
    return values


def write_to_original(
    *,
    declarations: list[str],
    source_rows: list[dict],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 60,
    _target: str = "original",
) -> WriteSummary:
    """写一个 AI 结果表：拆单/成本字段、两个状态及去重后的 PDF。"""
    del source_rows  # 兼容既有调用签名；来源表明确不在本链路中访问。
    if _target == "copy":
        app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
        table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
        mode = "copy-ai-fields-only"
        result_name = "copy_ai_write_result.json"
    elif _target == "original":
        app, table = custom_cost_ai_table()
        mode = "cost-ai-fields-only"
        result_name = "cost_ai_write_result.json"
    else:
        raise ValueError(f"未知写回目标：{_target}")
    client = FeishuClient()

    ai_field_map = ensure_ai_result_fields(client, app, table)
    existing = _records_for_declarations(client, app, table, declarations)
    protected_before = {
        str(record.get("record_id") or ""): _record_value(
            record, PROTECTED_PURCHASE_FIELD
        )
        for records in existing.values()
        for record in records
    }

    upload_tokens: dict[str, str] = {}
    for declaration in declarations:
        token = ""
        for record in existing.get(declaration) or []:
            attachments = (record.get("fields") or {}).get("pdf原件") or []
            if attachments:
                token = str((attachments[0] or {}).get("file_token") or "")
                if token:
                    break
        upload_tokens[declaration] = token or client.upload_bitable_file(
            app, pdfs[declaration]
        )

    updates: list[dict] = []
    creates: list[dict] = []
    for declaration in declarations:
        desired_rows = [
            row
            for row in ai_rows
            if str(row.get("报关单号") or "") == declaration
        ]
        remaining = list(existing.get(declaration) or [])
        matched_records: list[dict | None] = []
        for row in desired_rows:
            amount = decimal_value(row.get("报关金额"))
            weight = decimal_value(row.get("报关重量"))
            match = next(
                (
                    record
                    for record in remaining
                    if decimal_value(
                        _record_value(record, "报关金额AI")
                        or _record_value(record, "报关金额")
                    )
                    == amount
                    and decimal_value(
                        _record_value(record, "报关重量AI")
                        or _record_value(record, "报关重量")
                    )
                    == weight
                ),
                None,
            )
            matched_records.append(match)
            if match:
                remaining.remove(match)

        if remaining:
            primary = remaining.pop(0)
            for index, record in enumerate(matched_records):
                if record is None:
                    matched_records[index] = primary
                    break
        if remaining:
            raise RuntimeError(
                f"{declaration} 原表已有 {len(remaining) + 1} 条无法对应的记录，"
                "拒绝自动删除"
            )

        for index, (row, record) in enumerate(zip(desired_rows, matched_records)):
            fields = original_ai_fields(row)
            fields["pdf原件"] = (
                [{"file_token": upload_tokens[declaration]}] if index == 0 else []
            )
            if record:
                updates.append(
                    {"record_id": str(record["record_id"]), "fields": fields}
                )
            else:
                creates.append(fields)

    if updates:
        client.batch_update_records(app, table, updates)
    if creates:
        client.batch_create_records(app, table, creates)

    expected = Counter(str(row["报关单号"]) for row in ai_rows)
    problem = ""
    deadline = time.monotonic() + wait_seconds
    while time.monotonic() < deadline:
        found = _records_for_declarations(client, app, table, declarations)
        counts = {key: len(value) for key, value in found.items()}
        if counts != dict(expected):
            problem = f"原表记录数未达到目标：{counts}"
            time.sleep(2)
            continue

        problems: list[str] = []
        for declaration, records in found.items():
            desired_rows = [
                row
                for row in ai_rows
                if str(row.get("报关单号") or "") == declaration
            ]
            def canonical(field_type: int, value: Any) -> Any:
                if field_type == 2:
                    number = decimal_value(value)
                    return number.normalize() if number is not None else None
                if field_type == 5:
                    number = decimal_value(value)
                    return int(number) if number is not None else None
                if isinstance(value, list):
                    value = "、".join(
                        str(item.get("text") or item.get("name") or "")
                        if isinstance(item, dict)
                        else str(item)
                        for item in value
                    )
                return str(value or "").strip()

            expected_rows = sorted(
                tuple(
                    canonical(field_type, row.get(source))
                    for source, _, field_type in AI_RESULT_FIELD_SPECS
                )
                for row in desired_rows
            )
            actual_rows = sorted(
                tuple(
                    canonical(field_type, _record_value(record, target))
                    for _, target, field_type in AI_RESULT_FIELD_SPECS
                )
                for record in records
            )
            if actual_rows != expected_rows:
                problems.append(f"{declaration}: AI结果字段不一致")
            expected_status = sorted(
                (
                    str(row.get("拆单状态") or "正常"),
                    str(row.get("成本匹配状态") or "正常"),
                )
                for row in desired_rows
            )
            actual_status = sorted(
                (
                    _first_text(_record_value(record, "拆单状态")),
                    _first_text(_record_value(record, "成本匹配状态")),
                )
                for record in records
            )
            if actual_status != expected_status:
                problems.append(f"{declaration}: 状态字段不一致")
            attachment_count = sum(
                len((record.get("fields") or {}).get("pdf原件") or [])
                for record in records
            )
            if attachment_count != 1:
                problems.append(f"{declaration}: PDF 附件数量为 {attachment_count}")
            else:
                attached = next(
                    record
                    for record in records
                    if (record.get("fields") or {}).get("pdf原件")
                )
                first_row = desired_rows[0]
                if (
                    decimal_value(_record_value(attached, "报关金额AI"))
                    != decimal_value(first_row.get("报关金额"))
                    or decimal_value(_record_value(attached, "报关重量AI"))
                    != decimal_value(first_row.get("报关重量"))
                ):
                    problems.append(f"{declaration}: PDF 未附在第一条拆分行")
        if not problems:
            summary = WriteSummary(
                marker=marker,
                declarations=declarations,
                sourceRecords=0,
                aiRecords=len(ai_rows),
                uploadedFiles=len(set(upload_tokens.values())),
                workDir=str(work_dir),
                mode=mode,
            )
            (work_dir / result_name).write_text(
                json.dumps(
                    {
                        **asdict(summary),
                        "aiFields": {
                            name: str(field["field_id"])
                            for name, field in ai_field_map.items()
                        },
                        "recordIds": [
                            str(record.get("record_id") or "")
                            for records in found.values()
                            for record in records
                        ],
                        "fileTokens": upload_tokens,
                        "protectedPurchaseFieldWritten": False,
                        "sourceTableTouched": False,
                        "protectedPurchaseSnapshotBefore": protected_before,
                    },
                    ensure_ascii=False,
                    indent=1,
                ),
                encoding="utf-8",
            )
            return summary
        problem = "；".join(problems)
        time.sleep(2)
    raise RuntimeError(problem or "原 AI 表写入验证超时")


def _match_desired_records(
    desired_rows: list[dict],
    records: list[dict],
) -> list[dict | None]:
    """按报关单号、供应商、产品类型匹配；同键多行再用金额/重量区分。"""
    remaining = list(records)
    matched: list[dict | None] = []
    for index, row in enumerate(desired_rows):
        amount = decimal_value(row.get("报关金额"))
        weight = decimal_value(row.get("报关重量"))
        cost = decimal_value(row.get("采购金额"))
        basis = str(row.get("分摊依据") or "").strip()
        supplier = str(row.get("供应商") or "").strip()
        product = str(row.get("产品类型") or "").strip()
        product_name = str(row.get("报关品名") or "").strip()

        def same_key(record: dict) -> bool:
            fields = record.get("fields") or {}
            return (
                _first_text(fields.get("供应商AI") or fields.get("供应商")) == supplier
                and _first_text(
                    fields.get("产品类型AI") or fields.get("产品类型(对内)")
                )
                == product
            )

        candidates = [record for record in remaining if same_key(record)]
        exact = [
            record
            for record in candidates
            if decimal_value(
                _record_value(record, "报关金额AI")
                if _record_value(record, "报关金额AI") not in (None, "")
                else _record_value(record, "报关金额")
            )
            == amount
            and decimal_value(
                _record_value(record, "报关重量AI")
                if _record_value(record, "报关重量AI") not in (None, "")
                else _record_value(record, "报关重量")
            )
            == weight
        ]
        refined = [
            record
            for record in exact
            if decimal_value(_record_value(record, "采购金额AI")) == cost
        ]
        if len(refined) > 1:
            basis_matches = [
                record
                for record in refined
                if _first_text(_record_value(record, "采购金额分摊依据AI")) == basis
            ]
            if basis_matches:
                refined = basis_matches
        if len(refined) > 1 and product_name:
            name_matches = [
                record
                for record in refined
                if _first_text(_record_value(record, "报关品名")) == product_name
            ]
            if name_matches:
                refined = name_matches
        if refined:
            choice = sorted(
                refined,
                key=lambda record: str(record.get("record_id") or ""),
            )[0]
        else:
            choice = (
                exact[0]
                if len(exact) == 1
                else candidates[0]
                if len(candidates) == 1
                else None
            )

        if choice is None:
            exact_any = [
                record
                for record in remaining
                if decimal_value(
                    _record_value(record, "报关金额AI")
                    if _record_value(record, "报关金额AI") not in (None, "")
                    else _record_value(record, "报关金额")
                )
                == amount
                and decimal_value(
                    _record_value(record, "报关重量AI")
                    if _record_value(record, "报关重量AI") not in (None, "")
                    else _record_value(record, "报关重量")
                )
                == weight
            ]
            if len(exact_any) == 1:
                choice = exact_any[0]
            elif exact_any:
                refined_any = [
                    record
                    for record in exact_any
                    if decimal_value(_record_value(record, "采购金额AI")) == cost
                    and _first_text(_record_value(record, "采购金额分摊依据AI")) == basis
                ]
                if len(refined_any) > 1 and product_name:
                    name_matches = [
                        record
                        for record in refined_any
                        if _first_text(_record_value(record, "报关品名")) == product_name
                    ]
                    if name_matches:
                        refined_any = name_matches
                if refined_any:
                    choice = sorted(
                        refined_any,
                        key=lambda record: str(record.get("record_id") or ""),
                    )[0]

        if choice is None and index == 0:
            empty_keys = [
                record
                for record in remaining
                if not _first_text(
                    (record.get("fields") or {}).get("供应商AI")
                    or (record.get("fields") or {}).get("供应商")
                )
                or not _first_text(
                    (record.get("fields") or {}).get("产品类型AI")
                    or (record.get("fields") or {}).get("产品类型(对内)")
                )
            ]
            if empty_keys:
                choice = empty_keys[0]
            elif len(desired_rows) == 1 and len(remaining) == 1:
                choice = remaining[0]

        matched.append(choice)
        if choice:
            remaining.remove(choice)
    return matched


def write_backfill_to_cost_table(
    *,
    declarations: list[str],
    source_rows: list[dict],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 90,
) -> WriteSummary:
    """按业务三元组更新或新增成本表行，并补齐解析事实和 AI 结果。"""
    app, table = custom_cost_ai_table()
    client = FeishuClient()
    ai_field_map = ensure_ai_result_fields(client, app, table)
    existing = _records_for_declarations(client, app, table, declarations)

    upload_tokens: dict[str, str] = {}
    for declaration in declarations:
        token = ""
        for record in existing.get(declaration) or []:
            attachments = (record.get("fields") or {}).get("pdf原件") or []
            if attachments:
                token = str((attachments[0] or {}).get("file_token") or "")
                if token:
                    break
        upload_tokens[declaration] = token or client.upload_bitable_file(
            app, pdfs[declaration]
        )

    updates: list[dict] = []
    creates: list[dict] = []
    expected: dict[str, list[str]] = {}
    for declaration in declarations:
        desired_rows = [
            row
            for row in ai_rows
            if str(row.get("报关单号") or "") == declaration
        ]
        records = list(existing.get(declaration) or [])
        matched = _match_desired_records(desired_rows, records)
        expected[declaration] = [
            str(record.get("record_id") or "") if record else ""
            for record in matched
        ]
        for index, (row, record) in enumerate(zip(desired_rows, matched)):
            fields = original_ai_fields(row)
            fields["pdf原件"] = (
                [{"file_token": upload_tokens[declaration]}] if index == 0 else []
            )
            if record:
                updates.append(
                    {"record_id": str(record["record_id"]), "fields": fields}
                )
            else:
                creates.append(fields)

    if updates:
        client.batch_update_records(app, table, updates)
    if creates:
        client.batch_create_records(app, table, creates)

    deadline = time.monotonic() + wait_seconds
    problem = ""
    while time.monotonic() < deadline:
        found = _records_for_declarations(client, app, table, declarations)
        matched_by_declaration: dict[str, list[dict]] = {}
        missing: list[str] = []
        for declaration in declarations:
            desired_rows = [
                row
                for row in ai_rows
                if str(row.get("报关单号") or "") == declaration
            ]
            matched_records = _match_desired_records(
                desired_rows, list(found.get(declaration) or [])
            )
            if any(record is None for record in matched_records):
                missing.append(declaration)
                continue
            matched_by_declaration[declaration] = [
                record for record in matched_records if record is not None
            ]
        if missing:
            problem = f"以下报关单尚未匹配到目标行：{missing}"
            time.sleep(2)
            continue

        problems: list[str] = []
        for declaration, records in matched_by_declaration.items():
            desired_rows = [
                row
                for row in ai_rows
                if str(row.get("报关单号") or "") == declaration
            ]
            for row, record in zip(desired_rows, records):
                numeric_parsed_fields = {
                    "出口日期",
                    "运费",
                    "保费",
                    "报关金额",
                    "报关重量",
                }
                for source, target, field_type in AI_RESULT_FIELD_SPECS:
                    expected_value = (
                        decimal_value(row.get(source))
                        if field_type == 2
                        else _first_text(row.get(source))
                    )
                    actual_value = (
                        decimal_value(_record_value(record, target))
                        if field_type == 2
                        else _first_text(_record_value(record, target))
                    )
                    if expected_value != actual_value:
                        problems.append(f"{declaration}: {target} 未正确回填")
                        break
                for field in ORIGINAL_PARSED_FIELDS:
                    if field not in row:
                        continue
                    if field in numeric_parsed_fields:
                        expected_value = decimal_value(row.get(field))
                        actual_value = decimal_value(_record_value(record, field))
                    else:
                        expected_value = _first_text(row.get(field))
                        actual_value = _first_text(_record_value(record, field))
                    if expected_value != actual_value:
                        problems.append(f"{declaration}: {field} 未正确回填")
                        break
                if _first_text(_record_value(record, "拆单状态")) != str(
                    row.get("拆单状态") or "正常"
                ):
                    problems.append(f"{declaration}: 拆单状态未正确回填")
                if _first_text(_record_value(record, "成本匹配状态")) != str(
                    row.get("成本匹配状态") or "正常"
                ):
                    problems.append(f"{declaration}: 成本匹配状态未正确回填")
            attachments = [
                len((record.get("fields") or {}).get("pdf原件") or [])
                for record in records
            ]
            if sum(attachments) != 1:
                problems.append(
                    f"{declaration}: PDF 应仅附在一条目标行，实际 {sum(attachments)} 个"
                )

        if not problems:
            record_ids = [
                str(record.get("record_id") or "")
                for records in matched_by_declaration.values()
                for record in records
            ]
            summary = WriteSummary(
                marker=marker,
                declarations=declarations,
                sourceRecords=len(source_rows),
                aiRecords=len(ai_rows),
                uploadedFiles=len(set(upload_tokens.values())),
                workDir=str(work_dir),
                mode="cost-ai-backfill-by-business-key",
            )
            (work_dir / "cost_ai_backfill_result.json").write_text(
                json.dumps(
                    {
                        **asdict(summary),
                        "recordIds": record_ids,
                        "fileTokens": upload_tokens,
                        "matchedBefore": expected,
                        "protectedPurchaseFieldWritten": False,
                        "sourceTableTouched": False,
                    },
                    ensure_ascii=False,
                    indent=1,
                ),
                encoding="utf-8",
            )
            return summary
        problem = "；".join(problems[:20])
        time.sleep(2)
    raise RuntimeError(problem or "成本匹配 AI 表业务键回填验证超时")


def write_backfill_to_original(**kwargs: Any) -> WriteSummary:
    """兼容旧调用名；现在写到「2026年成本匹配 AI」表。"""
    return write_backfill_to_cost_table(**kwargs)


def write_to_copy_ai_fields(
    *,
    declarations: list[str],
    source_rows: list[dict],
    ai_rows: list[dict],
    pdfs: dict[str, Path],
    work_dir: Path,
    marker: str,
    wait_seconds: int = 60,
) -> WriteSummary:
    """显式写回飞书副本表，供副本联调使用；正式任务链写原表。"""
    return write_to_original(
        declarations=declarations,
        source_rows=source_rows,
        ai_rows=ai_rows,
        pdfs=pdfs,
        work_dir=work_dir,
        marker=marker,
        wait_seconds=wait_seconds,
        _target="copy",
    )


def remove_duplicate_copy_attachments(work_dir: Path) -> dict:
    """清理副本测试记录：每张报关单只在第一条拆分行保留 PDF。"""
    plan = json.loads((work_dir / "write_plan.json").read_text(encoding="utf-8"))
    declarations = list(plan["declarations"])
    client = FeishuClient()
    app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
    table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    records = _records_for_declarations(client, app, table, declarations)
    updates = []
    kept = []
    for declaration in declarations:
        ordered = list(records.get(declaration) or [])
        if not ordered:
            continue
        kept.append(str(ordered[0].get("record_id") or ""))
        for record in ordered[1:]:
            if (record.get("fields") or {}).get("pdf原件"):
                updates.append({"record_id": record["record_id"], "fields": {"pdf原件": []}})
    if updates:
        client.batch_update_records(app, table, updates)
    result = {
        "declarations": declarations,
        "keptRecordIds": kept,
        "clearedRecords": len(updates),
    }
    (work_dir / "copy_attachment_cleanup.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    return result


def fill_ai_status_fields(work_dir: Path) -> dict:
    """清理临时成本列，只保留拆单状态和成本匹配状态。"""
    plan = json.loads((work_dir / "write_plan.json").read_text(encoding="utf-8"))
    declarations = list(plan["declarations"])
    planned = list(plan["aiRecords"])
    client = FeishuClient()
    app = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
    table = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    existing_fields = {field.get("field_name"): field for field in client.list_fields(app, table)}
    deleted_fields = []
    for name in LEGACY_COST_FIELDS:
        field = existing_fields.get(name)
        if field:
            client.delete_field(app, table, str(field["field_id"]))
            deleted_fields.append(name)
            existing_fields.pop(name, None)

    created_fields = []
    for name, field_type in STATUS_FIELDS:
        if name not in existing_fields:
            created = client.create_field(app, table, name, field_type)
            existing_fields[name] = created
            created_fields.append(name)

    records = _records_for_declarations(client, app, table, declarations)
    used: set[str] = set()
    updates = []
    for planned_row in planned:
        declaration = str(planned_row["报关单号"])
        amount = decimal_value(planned_row.get("报关金额"))
        weight = decimal_value(planned_row.get("报关重量"))
        matches = [
            record
            for record in records[declaration]
            if record.get("record_id") not in used
            and decimal_value((record.get("fields") or {}).get("报关金额")) == amount
            and decimal_value((record.get("fields") or {}).get("报关重量")) == weight
        ]
        if len(matches) != 1:
            raise RuntimeError(
                f"{declaration} 金额 {amount} / 重量 {weight} 无法唯一定位 AI 记录"
            )
        record = matches[0]
        used.add(record["record_id"])
        updates.append(
            {
                "record_id": record["record_id"],
                "fields": {
                    "拆单状态": str(planned_row.get("拆单状态") or "正常"),
                    "成本匹配状态": str(planned_row.get("成本匹配状态") or "正常"),
                },
            }
        )
    if len(updates) != len(planned):
        raise RuntimeError("存在未定位到 AI 记录的计划行")
    result = client.batch_update_records(app, table, updates)
    payload = {
        "updatedRecords": len(result.get("records") or []),
        "deletedFields": deleted_fields,
        "createdFields": created_fields,
        "recordIds": [item["record_id"] for item in updates],
    }
    (work_dir / "status_field_update.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    return payload
