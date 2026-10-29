"""报关代理任务队列：飞书建单、财务电脑领取、PDF 上传与状态持久化。"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import threading
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from app.services.feishu_client import get_config

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TASKS_DIR = Path(get_config("CUSTOMS_TASKS_DIR", str(PROJECT_ROOT / ".runtime" / "customs_tasks")))
MAX_FILES = int(get_config("CUSTOMS_MAX_PDF_FILES", "500"))
MAX_FILE_BYTES = int(get_config("CUSTOMS_MAX_PDF_MB", "20")) * 1024 * 1024
MAX_TOTAL_BYTES = int(get_config("CUSTOMS_MAX_TOTAL_MB", "500")) * 1024 * 1024
DEFAULT_LEASE_SECONDS = int(get_config("CUSTOMS_AGENT_LEASE_SECONDS", "900"))
TASK_ID_RE = re.compile(r"^ct_[0-9a-f]{16}$")

_LOCK = threading.RLock()


class TaskNotFound(LookupError):
    pass


class TaskConflict(RuntimeError):
    pass


class TaskForbidden(PermissionError):
    pass


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _parse_time(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def _task_dir(task_id: str) -> Path:
    if not TASK_ID_RE.fullmatch(str(task_id or "")):
        raise TaskNotFound("任务不存在")
    return TASKS_DIR / task_id


def _task_path(task_id: str) -> Path:
    return _task_dir(task_id) / "task.json"


def _atomic_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.part")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=1, default=str),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _read_task(path: Path) -> dict[str, Any] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def get_task(task_id: str) -> dict[str, Any]:
    task = _read_task(_task_path(task_id))
    if task is None:
        raise TaskNotFound("任务不存在")
    return task


def _all_tasks() -> list[dict[str, Any]]:
    if not TASKS_DIR.exists():
        return []
    tasks = []
    for path in TASKS_DIR.glob("ct_*/task.json"):
        task = _read_task(path)
        if task:
            tasks.append(task)
    tasks.sort(key=lambda item: (str(item.get("createdAt") or ""), str(item.get("id") or "")))
    return tasks


def _save_task(task: dict[str, Any], event: str = "", detail: dict[str, Any] | None = None) -> dict[str, Any]:
    task["updatedAt"] = _now()
    history = list(task.get("history") or [])
    if event:
        history.append({"at": task["updatedAt"], "event": event, "detail": detail or {}})
    task["history"] = history[-50:]
    _atomic_json(_task_path(str(task["id"])), task)
    return task


def _string_list(value: Any) -> list[str]:
    if value in (None, ""):
        return []
    if isinstance(value, str) and value.strip().startswith("["):
        try:
            parsed = json.loads(value)
        except ValueError:
            parsed = None
        if isinstance(parsed, list):
            value = parsed
    items = value if isinstance(value, list) else re.split(r"[,，\n]", str(value))
    return list(dict.fromkeys(str(item).strip() for item in items if str(item).strip()))


def _normalize_date(value: Any, label: str) -> str:
    if isinstance(value, (int, float)) or (
        isinstance(value, str) and value.strip().isdigit() and len(value.strip()) >= 13
    ):
        try:
            return datetime.fromtimestamp(float(value) / 1000, tz=timezone.utc).astimezone(
                timezone(timedelta(hours=8))
            ).strftime("%Y-%m-%d")
        except (ValueError, OverflowError, OSError):
            raise ValueError(f"{label}时间戳无效")
    text = str(value or "").strip()
    if not text:
        return ""
    if "T" in text:
        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00")).strftime("%Y-%m-%d")
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"{label}格式应为 YYYY-MM-DD 或 YYYYMMDD")


def normalize_request(payload: dict[str, Any]) -> dict[str, Any]:
    """保留扩展字段，同时规范化财务电脑需要的检索参数。"""
    if not isinstance(payload, dict):
        raise ValueError("任务请求必须是 JSON 对象")
    flow = str(payload.get("flow") or "both").strip().lower()
    if flow not in {"export", "pre_record", "both"}:
        raise ValueError("flow 只能是 export、pre_record 或 both")
    start = _normalize_date(payload.get("startDate"), "startDate")
    end = _normalize_date(payload.get("endDate"), "endDate")
    if start and end and start > end:
        raise ValueError("startDate 不能晚于 endDate")
    declarations = _string_list(payload.get("declarations"))
    contracts = _string_list(payload.get("contracts"))
    if len(declarations) > 500 or len(contracts) > 500:
        raise ValueError("declarations 和 contracts 分别最多 500 项")
    parameters = {
        **payload,
        "flow": flow,
        "startDate": start,
        "endDate": end,
        "declarations": declarations,
        "contracts": contracts,
    }
    for control_field in ("idempotencyKey", "authToken", "createToken", "token"):
        parameters.pop(control_field, None)
    return parameters


def create_task(
    payload: dict[str, Any],
    *,
    requested_by: str = "",
    source: str = "feishu_button",
) -> tuple[dict[str, Any], bool]:
    parameters = normalize_request(payload)
    idempotency_key = str(
        payload.get("idempotencyKey")
        or payload.get("requestId")
        or payload.get("recordId")
        or ""
    ).strip()
    with _LOCK:
        if idempotency_key:
            for existing in reversed(_all_tasks()):
                if existing.get("idempotencyKey") != idempotency_key:
                    continue
                if existing.get("parameters") != parameters:
                    raise TaskConflict("同一幂等键对应了不同任务参数")
                return existing, True

        task_id = f"ct_{uuid.uuid4().hex[:16]}"
        task = {
            "id": task_id,
            "state": "queued",
            "createdAt": _now(),
            "updatedAt": _now(),
            "requestedBy": requested_by,
            "source": source,
            "idempotencyKey": idempotency_key,
            "parameters": parameters,
            "progress": 0,
            "message": "等待财务电脑领取",
            "history": [],
        }
        task = _save_task(task, "created")
        task["queuePosition"] = sum(
            1 for item in _all_tasks()
            if item.get("state") == "queued" and item.get("id") != task_id
        ) + 1
        return task, False


def claim_next(*, agent_name: str, lease_seconds: int = DEFAULT_LEASE_SECONDS) -> dict[str, Any] | None:
    now = datetime.now(timezone.utc)
    with _LOCK:
        candidates: list[tuple[int, dict[str, Any]]] = []
        for task in _all_tasks():
            state = task.get("state")
            eligible = state == "queued"
            if state == "claimed":
                expires = str(task.get("leaseExpiresAt") or "")
                try:
                    eligible = not expires or _parse_time(expires) <= now
                except ValueError:
                    eligible = True
            if eligible:
                candidates.append((0 if state == "queued" else 1, task))
        if not candidates:
            return None
        candidates.sort(
            key=lambda item: (
                item[0],
                str(item[1].get("createdAt") or ""),
                str(item[1].get("id") or ""),
            )
        )
        task = candidates[0][1]
        task.update(
            {
                "state": "claimed",
                "claimedAt": now.isoformat(timespec="seconds"),
                "agentName": agent_name,
                "leaseExpiresAt": (now + timedelta(seconds=lease_seconds)).isoformat(
                    timespec="seconds"
                ),
                "progress": 2,
                "message": "财务电脑已领取任务",
            }
        )
        return _save_task(task, "claimed", {"agentName": agent_name})


def heartbeat(
    task_id: str,
    *,
    agent_name: str,
    lease_seconds: int = DEFAULT_LEASE_SECONDS,
) -> dict[str, Any]:
    with _LOCK:
        task = get_task(task_id)
        if task.get("state") not in {"claimed", "processing"}:
            raise TaskConflict(f"任务状态 {task.get('state')} 不能续租")
        owner = str(task.get("agentName") or "")
        if owner and owner != agent_name:
            raise TaskConflict("任务已由其他财务电脑领取")
        task["leaseExpiresAt"] = (
            datetime.now(timezone.utc) + timedelta(seconds=lease_seconds)
        ).isoformat(timespec="seconds")
        return _save_task(task, "heartbeat", {"agentName": agent_name})


def report_task_error(
    task_id: str,
    *,
    agent_name: str,
    error_code: str,
    message: str,
    retryable: bool,
    reported_at: str,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """记录财务电脑无法继续任务的业务原因，不启动 PDF 处理。"""
    error_code = str(error_code or "").strip()
    message = str(message or "").strip()
    agent_name = str(agent_name or "").strip()
    if not error_code or len(error_code) > 64:
        raise ValueError("errorCode 必须为 1 到 64 个字符")
    if not message or len(message) > 2000:
        raise ValueError("message 必须为 1 到 2000 个字符")
    if not agent_name or len(agent_name) > 100:
        raise ValueError("agentName 必须为 1 到 100 个字符")
    detail_payload = details or {}
    if not isinstance(detail_payload, dict):
        raise ValueError("details 必须是 JSON 对象")
    if len(json.dumps(detail_payload, ensure_ascii=False, default=str)) > 8000:
        raise ValueError("details 过大")

    with _LOCK:
        task = get_task(task_id)
        state = str(task.get("state") or "")
        owner = str(task.get("agentName") or "")
        if state in {"failed", "blocked"}:
            error_agent = str(task.get("errorAgentName") or "")
            if error_agent != agent_name:
                raise TaskForbidden("任务失败信息已由其他 Agent 上报")
            if (
                str(task.get("errorCode") or "") == error_code
                and error_agent == agent_name
            ):
                return task
            raise TaskConflict(f"任务已处于 {state} 状态，不能覆盖原失败信息")
        if state != "claimed":
            raise TaskConflict(f"任务状态 {state} 不允许上报失败")
        if owner != agent_name:
            raise TaskForbidden("任务不是由该 Agent 领取")
        lease_expires = str(task.get("leaseExpiresAt") or "")
        try:
            if lease_expires and _parse_time(lease_expires) <= datetime.now(timezone.utc):
                raise TaskConflict("任务租约已过期，请重新领取")
        except ValueError:
            raise TaskConflict("任务租约无效，请重新领取")

        task_dir = _task_dir(task_id)
        for path in task_dir.glob(".uploads.*"):
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
        uploads = task_dir / "uploads"
        if uploads.exists():
            shutil.rmtree(uploads)

        received_at = _now()
        next_state = "blocked" if retryable else "failed"
        task.update(
            {
                "state": next_state,
                "message": f"任务{('阻塞' if retryable else '失败')}：{message}",
                "errorCode": error_code,
                "errorMessage": message,
                "errorRetryable": bool(retryable),
                "errorReportedAt": reported_at,
                "errorReceivedAt": received_at,
                "errorAgentName": agent_name,
                "errorDetails": detail_payload,
                "error": message,
                "blockedAt" if retryable else "failedAt": received_at,
            }
        )
        if not retryable:
            task["progress"] = 100
        return _save_task(
            task,
            "error-reported",
            {
                "errorCode": error_code,
                "retryable": bool(retryable),
                "agentName": agent_name,
            },
        )


def _write_pdf(path: Path, content: bytes) -> None:
    if not content or len(content) < 5 or not content.startswith(b"%PDF-"):
        raise ValueError(f"{path.name} 不是有效 PDF")
    path.write_bytes(content)
    if path.stat().st_size != len(content):
        path.unlink(missing_ok=True)
        raise OSError(f"{path.name} 写入不完整")


def save_uploads(
    task_id: str,
    files: list[tuple[str, bytes]],
    *,
    agent_name: str = "",
) -> dict[str, Any]:
    if not files:
        raise ValueError("至少上传一个 PDF")
    if len(files) > MAX_FILES:
        raise ValueError(f"单任务最多上传 {MAX_FILES} 个 PDF")
    if sum(len(content) for _, content in files) > MAX_TOTAL_BYTES:
        raise ValueError(f"单任务 PDF 总量超过 {MAX_TOTAL_BYTES // 1024 // 1024} MB")

    with _LOCK:
        task = get_task(task_id)
        if task.get("state") in {"processing", "completed"}:
            raise TaskConflict("任务已进入处理阶段或已完成")
        if task.get("state") not in {"claimed", "uploaded"}:
            raise TaskConflict("任务尚未被财务电脑领取")
        owner = str(task.get("agentName") or "")
        if owner and owner != agent_name:
            raise TaskConflict("任务已由其他财务电脑领取")

        task_dir = _task_dir(task_id)
        staging = task_dir / f".uploads.{uuid.uuid4().hex}"
        staging.mkdir(parents=True, exist_ok=False)
        manifest = []
        used: set[str] = set()
        try:
            total = 0
            for index, (raw_name, content) in enumerate(files, 1):
                safe_name = Path(str(raw_name).replace("\\", "/")).name
                if not safe_name.lower().endswith(".pdf"):
                    raise ValueError(f"{safe_name or f'第{index}个文件'} 不是 PDF")
                if len(content) > MAX_FILE_BYTES:
                    raise ValueError(f"{safe_name} 超过 {MAX_FILE_BYTES // 1024 // 1024} MB")
                name = safe_name
                if name in used:
                    name = f"{Path(safe_name).stem}_{index}{Path(safe_name).suffix}"
                used.add(name)
                _write_pdf(staging / name, content)
                total += len(content)
                manifest.append(
                    {
                        "name": name,
                        "bytes": len(content),
                        "sha256": hashlib.sha256(content).hexdigest(),
                    }
                )
            if total <= 0:
                raise ValueError("上传内容为空")
            uploads = task_dir / "uploads"
            if uploads.exists():
                shutil.rmtree(uploads)
            os.replace(staging, uploads)
        except Exception:
            shutil.rmtree(staging, ignore_errors=True)
            raise

        task.update(
            {
                "state": "uploaded",
                "uploadedAt": _now(),
                "fileCount": len(manifest),
                "files": manifest,
                "progress": 3,
                "message": "PDF 已上传，等待服务器核算",
            }
        )
        return _save_task(task, "uploaded", {"fileCount": len(manifest)})


def mark_processing(task_id: str) -> dict[str, Any]:
    with _LOCK:
        task = get_task(task_id)
        if task.get("state") not in {"uploaded", "processing"}:
            raise TaskConflict(f"任务状态 {task.get('state')} 不能开始处理")
        task.update({"state": "processing", "progress": 5, "message": "开始解析 PDF"})
        return _save_task(task, "processing-started")


def update_processing(
    task_id: str,
    *,
    progress: int,
    message: str,
) -> dict[str, Any]:
    with _LOCK:
        task = get_task(task_id)
        if task.get("state") not in {"processing", "uploaded"}:
            raise TaskConflict(f"任务状态 {task.get('state')} 不能更新进度")
        task.update(
            {
                "state": "processing",
                "progress": max(0, min(99, int(progress))),
                "message": str(message or ""),
            }
        )
        return _save_task(task, "progress", {"progress": task["progress"], "message": message})


def mark_completed(task_id: str, summary: dict[str, Any]) -> dict[str, Any]:
    with _LOCK:
        task = get_task(task_id)
        task.update(
            {
                "state": "completed",
                "progress": 100,
                "message": "飞书副本写回完成",
                "completedAt": _now(),
                "summary": summary,
            }
        )
        return _save_task(task, "completed")


def mark_failed(task_id: str, error: str) -> dict[str, Any]:
    with _LOCK:
        task = get_task(task_id)
        task.update(
            {
                "state": "failed",
                "progress": 100,
                "message": "任务失败",
                "failedAt": _now(),
                "error": str(error)[:4000],
            }
        )
        return _save_task(task, "failed")


def clear_pdf_cache(task_id: str) -> dict[str, Any]:
    """删除本任务上传 PDF 和中间产物，仅保留 task.json 状态。"""
    task_dir = _task_dir(task_id)
    removed = []
    for name in ("uploads", "work"):
        path = task_dir / name
        if path.exists():
            shutil.rmtree(path)
            removed.append(name)
    with _LOCK:
        task = get_task(task_id)
        task["cacheClearedAt"] = _now()
        task["cacheRemoved"] = removed
        return _save_task(task, "cache-cleared", {"removed": removed})


def recover_interrupted() -> list[str]:
    """服务器重启后，将仍有 PDF 的任务重新置为可处理。"""
    recoverable = []
    with _LOCK:
        for task in _all_tasks():
            task_dir = _task_dir(str(task["id"]))
            for staging in task_dir.glob(".uploads.*"):
                if staging.is_dir():
                    shutil.rmtree(staging, ignore_errors=True)
            uploads = task_dir / "uploads"
            has_pdf = uploads.exists() and any(uploads.glob("*.pdf"))
            if task.get("state") in {"processing", "claimed"} and has_pdf:
                task.update(
                    {
                        "state": "uploaded",
                        "progress": 3,
                        "message": "服务重启，任务自动恢复",
                    }
                )
                _save_task(task, "recovered")
                recoverable.append(str(task["id"]))
            elif task.get("state") == "processing":
                task.update(
                    {
                        "state": "failed",
                        "progress": 100,
                        "message": "服务重启且上传文件不存在",
                        "error": "服务重启发生在上传完成前，请重新上传 PDF",
                    }
                )
                _save_task(task, "recovery-failed")
    return recoverable


def list_tasks(limit: int = 100) -> list[dict[str, Any]]:
    return _all_tasks()[-max(1, min(int(limit), 500)):][::-1]
