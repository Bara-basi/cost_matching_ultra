"""执行财务电脑上传的报关 PDF：解析、拆单、成本匹配并写回飞书副本。"""
from __future__ import annotations

import json
import shutil
import threading
import traceback
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

from app.services import customs_tasks as customs_task_queue, pipeline

_LOCK = threading.Lock()
_PROCESS_LOCK = threading.Lock()
_RUNNING: set[str] = set()


def _write_result(task_id: str, payload: dict[str, Any]) -> None:
    path = customs_task_queue._task_dir(task_id) / "result.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=1, default=str),
        encoding="utf-8",
    )


def _process(task_id: str) -> dict[str, Any]:
    task = customs_task_queue.mark_processing(task_id)
    task_dir = customs_task_queue._task_dir(task_id)
    uploads = task_dir / "uploads"
    work_dir = task_dir / "work"
    pdfs = sorted(uploads.glob("*.pdf"))
    if not pdfs:
        raise ValueError("任务目录中没有 PDF")
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    customs_task_queue.update_processing(task_id, progress=12, message="正在解析 PDF")
    parse_xlsx = work_dir / "parse" / "报关单解析结果.xlsx"
    parse_stats = pipeline.parse_pdfs(pdfs, parse_xlsx)
    failed = list(parse_stats.get("failed") or [])
    if failed:
        detail = "；".join(
            f"{item.get('文件', '')}: {item.get('原因', '')}" for item in failed[:5]
        )
        raise ValueError(f"{len(failed)} 份 PDF 解析失败：{detail}")
    if not parse_xlsx.exists():
        raise ValueError("PDF 未解析出可核算的正式出口报关单")

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

    from app.services.customs_ingest import build_records, write_to_copy_ai_fields

    marker = f"飞书任务-{task_id}-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    source_rows, ai_rows, declaration_pdfs = build_records(
        declarations,
        rows,
        parse_xlsx,
        uploads,
        marker,
    )
    customs_task_queue.update_processing(task_id, progress=76, message="正在写回飞书副本")
    summary = write_to_copy_ai_fields(
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
    write_detail_path = work_dir / "copy_ai_write_result.json"
    result = {
        "taskId": task_id,
        "declarations": declarations,
        "pdfCount": len(pdfs),
        "parseRows": parse_stats.get("rows", 0),
        "splitStats": split_stats,
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
        customs_task_queue.mark_completed(
            task_id,
            {
                "declarations": result["declarations"],
                "pdfCount": result["pdfCount"],
                "aiRecords": result["write"]["aiRecords"],
                "uploadedFiles": result["write"]["uploadedFiles"],
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


def start(task_id: str) -> None:
    thread = threading.Thread(target=_run, args=(task_id,), daemon=True)
    thread.start()


def resume_uploaded() -> list[str]:
    ids = []
    for task in customs_task_queue.list_tasks(limit=500):
        if task.get("state") == "uploaded":
            start(str(task["id"]))
            ids.append(str(task["id"]))
    return ids
