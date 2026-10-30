"""执行财务电脑上传的报关 PDF：解析、拆单、成本匹配并写回成本匹配 AI 表。"""
from __future__ import annotations

import json
import re
import shutil
import threading
import traceback
from dataclasses import asdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

from app.services import customs_tasks as customs_task_queue, pipeline
from app.services.feishu_client import get_config

_LOCK = threading.Lock()
_PROCESS_LOCK = threading.Lock()
_RUNNING: set[str] = set()
_DOCUMENT_EVENTS: dict[str, threading.Event] = {}
DOCUMENT_NAMES_FILE = "document_names.json"


def _upload_name_wait_seconds() -> int:
    try:
        return max(
            1,
            int(get_config("CUSTOMS_DOCUMENT_NAME_WAIT_SECONDS", "300")),
        )
    except (TypeError, ValueError):
        return 300


UPLOAD_NAME_WAIT_SECONDS = _upload_name_wait_seconds()


def _write_result(task_id: str, payload: dict[str, Any]) -> None:
    path = customs_task_queue._task_dir(task_id) / "result.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=1, default=str),
        encoding="utf-8",
    )


def _filename_part(value: Any) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value or "").strip())
    return text.rstrip(" .")


def _year_month(row: dict[str, Any]) -> tuple[str, str]:
    for field in ("出口日期", "申报日期"):
        value = row.get(field)
        if isinstance(value, (datetime, date)):
            return value.strftime("%Y%m"), field
        match = re.search(r"(\d{4})\D?(\d{1,2})", str(value or ""))
        if match:
            month = int(match.group(2))
            if 1 <= month <= 12:
                return f"{match.group(1)}{month:02d}", field
    return "", ""


def _document_name(row: dict[str, Any]) -> tuple[str, str, str]:
    year_month, date_field = _year_month(row)
    contract = ""
    for field in ("合同号_1", "合同号（应收表格）", "合同号"):
        contract = str(row.get(field) or "").strip()
        if contract:
            break
    declaration = str(row.get("报关单号") or "").strip()
    missing = []
    if not year_month:
        missing.append("出口日期/申报日期")
    if not contract:
        missing.append("合同号")
    if len(declaration) < 4:
        missing.append("报关单号")
    if missing:
        return "", date_field, f"缺少{'、'.join(missing)}，无法生成财务文件名"
    safe_contract = _filename_part(contract)
    if not safe_contract:
        return "", date_field, "合同号包含不可用于文件名的字符"
    return (
        f"{year_month}_{safe_contract}_{declaration[-4:]}.pdf",
        date_field,
        "",
    )


def build_document_names(
    task_id: str,
    parse_xlsx: Path,
    upload_entries: list[dict[str, Any]],
    parse_stats: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """按上传接收顺序，把解析事实映射为财务要求的目标文件名。"""
    from app.services.customs_ingest import read_sheet

    parse_rows = read_sheet(parse_xlsx) if parse_xlsx.is_file() else []
    failed_reasons = {
        str(item.get("文件") or ""): str(item.get("原因") or "").strip()
        for item in (parse_stats or {}).get("failed") or []
        if isinstance(item, dict)
    }
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in parse_rows:
        source = Path(str(row.get("来源文件") or "")).name
        if source:
            grouped.setdefault(source, []).append(row)

    documents: list[dict[str, Any]] = []
    name_map: dict[str, str | None] = {}
    for position, entry in enumerate(upload_entries, 1):
        original = str(entry.get("originalFileName") or entry.get("name") or "")
        stored = str(entry.get("storedFileName") or entry.get("name") or "")
        rows = grouped.get(stored) or []
        error = ""
        month_date_field = ""
        new_name = ""
        if rows:
            pairs = {
                (
                    str(row.get("合同号_1") or ""),
                    str(row.get("报关单号") or ""),
                )
                for row in rows
            }
            if len(pairs) > 1:
                error = "一个 PDF 解析出多个合同号/报关单号，无法唯一命名"
            else:
                new_name, month_date_field, error = _document_name(rows[0])
        else:
            error = failed_reasons.get(stored) or "未解析出可命名的报关单"
        documents.append(
            {
                "index": int(entry.get("index") or position),
                "originalFileName": original,
                "storedFileName": stored,
                "newFileName": new_name or None,
                "monthDateField": month_date_field or None,
                "error": error or None,
            }
        )
        name_map[stored] = new_name or None
    return {
        "taskId": task_id,
        "readyAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "orderBasis": "multipart-files-field-received-order",
        "documents": documents,
        "documentNames": [item["newFileName"] for item in documents],
        "nameMap": name_map,
        "declarations": list(
            dict.fromkeys(
                str(row.get("报关单号") or "").strip()
                for row in parse_rows
                if str(row.get("报关单号") or "").strip()
            )
        ),
    }


def _is_parse_only_task(task: dict[str, Any]) -> bool:
    source = str(task.get("source") or "").strip().lower()
    if not source:
        source = str((task.get("parameters") or {}).get("source") or "").strip().lower()
    return source == "test"


def _publish_document_names(
    task_id: str, payload: dict[str, Any], *, notify: bool = True
) -> None:
    path = customs_task_queue._task_dir(task_id) / DOCUMENT_NAMES_FILE
    customs_task_queue._atomic_json(path, payload)
    if notify:
        with _LOCK:
            event = _DOCUMENT_EVENTS.get(task_id)
        if event:
            event.set()


def _publish_empty_document_names(task_id: str, reason: str) -> None:
    entries = customs_task_queue.ordered_upload_entries(task_id)
    documents = [
        {
            "index": int(entry.get("index") or position),
            "originalFileName": str(
                entry.get("originalFileName") or entry.get("name") or ""
            ),
            "storedFileName": str(
                entry.get("storedFileName") or entry.get("name") or ""
            ),
            "newFileName": None,
            "monthDateField": None,
            "error": reason,
        }
        for position, entry in enumerate(entries, 1)
    ]
    _publish_document_names(
        task_id,
        {
            "taskId": task_id,
            "readyAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "orderBasis": "multipart-files-field-received-order",
            "documents": documents,
            "documentNames": [None for _item in documents],
            "nameMap": {
                item["storedFileName"]: None
                for item in documents
                if item["storedFileName"]
            },
            "declarations": [],
        },
        notify=not _is_parse_only_task(customs_task_queue.get_task(task_id)),
    )


def document_names(task_id: str) -> dict[str, Any] | None:
    path = customs_task_queue._task_dir(task_id) / DOCUMENT_NAMES_FILE
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return payload if isinstance(payload, dict) else None


def wait_for_document_names(
    task_id: str, timeout: float = UPLOAD_NAME_WAIT_SECONDS
) -> dict[str, Any] | None:
    """等待解析完成为文件命名；超时后仍可从状态接口补取。"""
    with _LOCK:
        event = _DOCUMENT_EVENTS.get(task_id)
    if event and not event.wait(max(0.0, float(timeout))):
        return None
    return document_names(task_id)


def _process(task_id: str) -> dict[str, Any]:
    task = customs_task_queue.mark_processing(task_id)
    task_dir = customs_task_queue._task_dir(task_id)
    uploads = task_dir / "uploads"
    work_dir = task_dir / "work"
    upload_entries = customs_task_queue.ordered_upload_entries(task_id)
    pdfs = [
        uploads / str(entry.get("storedFileName") or entry.get("name") or "")
        for entry in upload_entries
    ]
    pdfs = [path for path in pdfs if path.is_file() and path.suffix.lower() == ".pdf"]
    if not pdfs:
        raise ValueError("任务目录中没有 PDF")
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    customs_task_queue.update_processing(task_id, progress=12, message="正在解析 PDF")
    parse_xlsx = work_dir / "parse" / "报关单解析结果.xlsx"
    try:
        parse_stats = pipeline.parse_pdfs(pdfs, parse_xlsx)
    except Exception as exc:
        _publish_empty_document_names(task_id, f"{type(exc).__name__}: {exc}")
        raise
    naming = build_document_names(
        task_id, parse_xlsx, upload_entries, parse_stats
    )
    parse_only = _is_parse_only_task(task)
    _publish_document_names(task_id, naming, notify=not parse_only)
    failed = list(parse_stats.get("failed") or [])
    if failed:
        detail = "；".join(
            f"{item.get('文件', '')}: {item.get('原因', '')}" for item in failed[:5]
        )
        raise ValueError(f"{len(failed)} 份 PDF 解析失败：{detail}")
    if not parse_xlsx.exists():
        raise ValueError("PDF 未解析出可核算的正式出口报关单")

    if parse_only:
        result = {
            "taskId": task_id,
            "mode": "test-parse-only",
            "declarations": naming.get("declarations") or [],
            "pdfCount": len(pdfs),
            "parseRows": parse_stats.get("rows", 0),
            "documentNames": naming["documentNames"],
            "documentRenameMap": naming["nameMap"],
            "finishedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "request": task.get("parameters") or {},
        }
        _write_result(task_id, result)
        return result

    customs_task_queue.update_processing(task_id, progress=30, message="正在拆单")
    split_dir = work_dir / "split"
    split_stats = pipeline.run_round(parse_xlsx, split_dir)
    split_path = split_dir / "拆单明细_全部.xlsx"
    if not split_path.exists():
        raise ValueError("拆单阶段没有产出行级结果")

    customs_task_queue.update_processing(task_id, progress=58, message="正在成本匹配")
    rows = pipeline.review_rows(split_path)
    declarations = list(
        dict.fromkeys(
            str(row.get("报关单号") or "").strip()
            for row in rows
            if str(row.get("报关单号") or "").strip()
        )
    )
    if not declarations:
        raise ValueError("拆单结果中没有报关单号")

    from app.services.customs_ingest import build_records, write_backfill_to_cost_table

    marker = f"飞书任务-{task_id}-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    source_rows, ai_rows, declaration_pdfs = build_records(
        declarations,
        rows,
        parse_xlsx,
        uploads,
        marker,
    )
    customs_task_queue.update_processing(task_id, progress=76, message="正在写回成本匹配 AI 表")
    summary = write_backfill_to_cost_table(
        declarations=declarations,
        source_rows=source_rows,
        ai_rows=ai_rows,
        pdfs=declaration_pdfs,
        work_dir=work_dir,
        marker=marker,
        wait_seconds=60,
    )
    write_summary = asdict(summary)
    write_summary["workDir"] = "(任务结束后已清理)"
    write_detail_path = work_dir / "cost_ai_backfill_result.json"
    result = {
        "taskId": task_id,
        "declarations": declarations,
        "pdfCount": len(pdfs),
        "parseRows": parse_stats.get("rows", 0),
        "splitStats": split_stats,
        "documentNames": naming["documentNames"],
        "documentRenameMap": naming["nameMap"],
        "write": write_summary,
        "writeDetails": (
            json.loads(write_detail_path.read_text(encoding="utf-8"))
            if write_detail_path.exists()
            else {}
        ),
        "finishedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "request": task.get("parameters") or {},
    }
    _write_result(task_id, result)
    return result


def _run(task_id: str) -> None:
    with _LOCK:
        if task_id in _RUNNING:
            return
        _RUNNING.add(task_id)
    try:
        with _PROCESS_LOCK:
            result = _process(task_id)
        write = result.get("write") or {}
        customs_task_queue.mark_completed(
            task_id,
            {
                "mode": result.get("mode") or "full",
                "declarations": result.get("declarations") or [],
                "pdfCount": result.get("pdfCount", 0),
                "aiRecords": write.get("aiRecords", 0),
                "uploadedFiles": write.get("uploadedFiles", 0),
                "documentNamesReady": sum(
                    bool(item) for item in (result.get("documentNames") or [])
                ),
                "resultFile": "result.json",
            },
        )
    except Exception as exc:  # noqa: BLE001
        error_detail = f"{type(exc).__name__}: {exc}"
        traceback.print_exc()
        try:
            _write_result(
                task_id,
                {
                    "taskId": task_id,
                    "error": error_detail,
                    "failedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
                },
            )
        finally:
            customs_task_queue.mark_failed(task_id, error_detail)
    finally:
        try:
            customs_task_queue.clear_pdf_cache(task_id)
        except Exception:  # noqa: BLE001
            # 清理失败不能覆盖已经落地的业务结果；错误会留在进程日志。
            traceback.print_exc()
        with _LOCK:
            _RUNNING.discard(task_id)
            event = _DOCUMENT_EVENTS.get(task_id)
        if event:
            event.set()


def start(task_id: str) -> None:
    with _LOCK:
        _DOCUMENT_EVENTS[task_id] = threading.Event()
    thread = threading.Thread(target=_run, args=(task_id,), daemon=True)
    thread.start()


def resume_uploaded() -> list[str]:
    ids = []
    for task in customs_task_queue.list_tasks(limit=500):
        if task.get("state") == "uploaded":
            start(str(task["id"]))
            ids.append(str(task["id"]))
    return ids
