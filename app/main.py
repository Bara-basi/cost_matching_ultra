r"""成本匹配工作台：FastAPI 服务（前端静态资源 + /api/*）。

启动：
  & '.\.venv\Scripts\python.exe' -m uvicorn app.main:app --host 127.0.0.1 --port 8765
"""
from __future__ import annotations

import json
import os
import secrets
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from pydantic import BaseModel, ConfigDict, Field, field_validator

from app.services import erp_sync, lookup, workspace
from app.services.feishu_client import get_config

PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_VERSION = "2026.10.29-original-ai.1"

app = FastAPI(title="成本匹配工作台", version=APP_VERSION)
WEB_DIST = PROJECT_ROOT / "web" / "dist"
if (WEB_DIST / "assets").exists():
    app.mount("/assets", StaticFiles(directory=WEB_DIST / "assets"), name="web-assets")


class OverrideRequest(BaseModel):
    index: int
    amount: str
    operator: str
    reason: str = ""


class ReviewRequest(BaseModel):
    index: int
    operator: str


class FlagExceptionRequest(BaseModel):
    index: int
    operator: str
    reason: str


class ResubmitExceptionRequest(BaseModel):
    index: int
    operator: str


class SourceMappingRequest(BaseModel):
    row_index: int
    source_id: str
    operator: str


class CustomsTaskErrorRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    agentName: str = Field(min_length=1, max_length=100)
    errorCode: str = Field(min_length=1, max_length=64)
    message: str = Field(min_length=1, max_length=2000)
    retryable: bool
    reportedAt: datetime
    details: dict[str, object] = Field(default_factory=dict)

    @field_validator("agentName", "errorCode", "message")
    @classmethod
    def strip_required_text(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("不能为空")
        return cleaned

    @field_validator("reportedAt")
    @classmethod
    def require_timezone(cls, value: datetime) -> datetime:
        if value.tzinfo is None or value.utcoffset() is None:
            raise ValueError("reportedAt 必须包含时区")
        return value


# 本机（127.0.0.1 / ::1）视为可信：财务在自己电脑上操作，不再需要额外口令。
# 只有把工作台暴露到局域网 / 公网时才要求 WORKBENCH_WRITE_TOKEN。
LOCAL_HOSTS = {"127.0.0.1", "::1", "localhost", "testclient"}


def _trust_all() -> bool:
    """`WORKBENCH_TRUST_ALL=1`：整站视为可信，不再要求网关口令。

    给「直接在 IP:端口 上给内部同事用、没有域名也没有反向代理」的部署用：
    任何能访问到该端口的人都能写回飞书、查看睿贝凭证，所以只在内部环境打开。
    """
    from app.services.feishu_client import get_config

    return str(get_config("WORKBENCH_TRUST_ALL", "")).strip().lower() in {"1", "true", "yes", "on"}


def _is_local(request: Request) -> bool:
    if _trust_all():
        return True
    host = (request.client.host if request.client else "") or ""
    return host in LOCAL_HOSTS


def _require_write_token(request: Request, provided: str = "") -> None:
    """写回：本机直接放行；远端必须由可信网关注入 `WORKBENCH_WRITE_TOKEN`。"""
    if _is_local(request):
        return
    from app.services.feishu_client import get_config

    configured = get_config("WORKBENCH_WRITE_TOKEN", "")
    if not configured or not provided or not secrets.compare_digest(configured, provided):
        raise HTTPException(status_code=403, detail="远端写回未启用或缺少内网网关授权")


def _require_evidence_token(request: Request, provided: str = "") -> None:
    if _is_local(request):
        return
    from app.services.feishu_client import get_config

    configured = get_config("WORKBENCH_EVIDENCE_TOKEN", "")
    if not configured or not provided or not secrets.compare_digest(configured, provided):
        raise HTTPException(status_code=403, detail="远端查看 ERP 凭证未启用或缺少内网网关授权")


def _require_customs_create_token(provided: str) -> None:
    configured = get_config("CUSTOMS_TASK_CREATE_TOKEN", "")
    if not configured:
        raise HTTPException(status_code=503, detail="未配置 CUSTOMS_TASK_CREATE_TOKEN")
    if not provided or not secrets.compare_digest(configured, provided):
        raise HTTPException(status_code=403, detail="飞书任务请求鉴权失败")


def _bearer_token(authorization: str) -> str:
    scheme, _, token = str(authorization or "").partition(" ")
    return token.strip() if scheme.lower() == "bearer" else ""


def _require_customs_agent_token(authorization: str) -> None:
    configured = get_config("CUSTOMS_AGENT_TOKEN", "")
    if not configured:
        raise HTTPException(status_code=503, detail="未配置 CUSTOMS_AGENT_TOKEN")
    provided = _bearer_token(authorization)
    if not provided or not secrets.compare_digest(configured, provided):
        raise HTTPException(status_code=403, detail="财务电脑代理鉴权失败")


@app.on_event("startup")
def _startup() -> None:
    workspace.recover_interrupted()
    from app.services import customs_tasks as customs_task_queue, customs_task_runner

    customs_task_queue.recover_interrupted()
    customs_task_runner.resume_uploaded()
    # 定时全量同步：启动时按需跑一次 + 后台每 12 小时一轮（可用环境变量调整）
    try:
        erp_sync.start_scheduler()
    except Exception:  # noqa: BLE001
        pass


@app.get("/api/health")
def health(request: Request) -> dict:
    from app.services.feishu_client import get_config

    return {
        "status": "ok",
        "appVersion": APP_VERSION,
        "projectPath": str(PROJECT_ROOT),
        "erpConfigured": bool(get_config("ERP_API_KEY") or get_config("ERP_USERNAME")),
        "feishuConfigured": bool(get_config("LARK_APP_ID") and get_config("LARK_APP_SECRET")),
        # 本机操作默认允许写回 / 查看凭证；对外暴露时才需要网关 token
        "writeEnabled": _is_local(request) or bool(get_config("WORKBENCH_WRITE_TOKEN", "")),
        "evidenceEnabled": _is_local(request) or bool(get_config("WORKBENCH_EVIDENCE_TOKEN", "")),
        "local": _is_local(request),
        # 打开了整站信任开关（没有域名/网关的直连部署）：写回与凭证不再要口令
        "trustAll": _trust_all(),
        "erpSync": erp_sync.status(),
        "customsTaskQueue": {
            "createConfigured": bool(get_config("CUSTOMS_TASK_CREATE_TOKEN", "")),
            "agentConfigured": bool(get_config("CUSTOMS_AGENT_TOKEN", "")),
        },
    }


@app.post("/api/customs/tasks")
async def create_customs_task(
    request: Request,
    x_customs_task_token: str = Header(""),
) -> dict:
    """飞书按钮入口：校验请求并持久化到财务电脑任务队列。"""
    content_type = str(request.headers.get("content-type") or "").lower()
    if "application/json" in content_type:
        try:
            body = await request.json()
        except ValueError as exc:
            raise HTTPException(status_code=400, detail="请求体不是有效 JSON") from exc
    else:
        form = await request.form()
        body = {key: value for key, value in form.items()}
    if not isinstance(body, dict):
        raise HTTPException(status_code=400, detail="请求体必须是对象")

    if body.get("type") == "url_verification":
        expected = get_config("CUSTOMS_FEISHU_VERIFICATION_TOKEN", "") or get_config(
            "CUSTOMS_TASK_CREATE_TOKEN", ""
        )
        if not expected:
            raise HTTPException(status_code=503, detail="飞书回调验证令牌未配置")
        provided = str(body.get("token") or "")
        if not provided or not secrets.compare_digest(expected, provided):
            raise HTTPException(status_code=403, detail="飞书回调验证失败")
        return {"challenge": str(body.get("challenge") or "")}

    _require_customs_create_token(
        x_customs_task_token
        or str(body.get("authToken") or body.get("createToken") or "")
    )
    nested = body.get("task")
    if isinstance(nested, dict):
        task_body = dict(nested)
        for key in ("idempotencyKey", "requestId", "recordId", "source"):
            if key in body and key not in task_body:
                task_body[key] = body[key]
    else:
        task_body = body
    idempotency_header = str(request.headers.get("idempotency-key") or "").strip()
    if idempotency_header and not task_body.get("idempotencyKey"):
        task_body["idempotencyKey"] = idempotency_header

    from app.services import customs_tasks as customs_task_queue

    try:
        task, duplicate = customs_task_queue.create_task(
            task_body,
            requested_by=str(
                body.get("operator")
                or body.get("requestedBy")
                or request.headers.get("x-feishu-user-id")
                or ""
            ),
            source=str(body.get("source") or "feishu_button"),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except customs_task_queue.TaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    return {
        "taskId": task["id"],
        "state": task["state"],
        "duplicate": duplicate,
        "queuePosition": task.get("queuePosition"),
        "statusUrl": f"/api/customs/tasks/{task['id']}",
    }


@app.get("/api/customs/tasks/claim")
def claim_customs_task(
    authorization: str = Header(""),
    x_agent_name: str = Header("finance-pc"),
) -> dict:
    """财务电脑轮询入口：原子领取最早的一条待处理任务。"""
    _require_customs_agent_token(authorization)
    from app.services import customs_tasks as customs_task_queue

    task = customs_task_queue.claim_next(agent_name=x_agent_name or "finance-pc")
    if task is None:
        return {"hasTask": False, "task": None}
    return {
        "hasTask": True,
        "task": task,
        "uploadUrl": f"/api/customs/tasks/{task['id']}/upload",
        "statusUrl": f"/api/customs/tasks/{task['id']}",
        "heartbeatUrl": f"/api/customs/tasks/{task['id']}/heartbeat",
    }


@app.post("/api/customs/tasks/{task_id}/heartbeat")
def heartbeat_customs_task(
    task_id: str,
    authorization: str = Header(""),
    x_agent_name: str = Header("finance-pc"),
) -> dict:
    _require_customs_agent_token(authorization)
    from app.services import customs_tasks as customs_task_queue

    try:
        task = customs_task_queue.heartbeat(
            task_id, agent_name=x_agent_name or "finance-pc"
        )
    except customs_task_queue.TaskNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except customs_task_queue.TaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    return {"taskId": task["id"], "state": task["state"], "leaseExpiresAt": task.get("leaseExpiresAt")}


@app.post("/api/customs/tasks/{task_id}/upload", status_code=202)
async def upload_customs_task_files(
    task_id: str,
    files: list[UploadFile] = File(...),
    agent_name: str = Form(""),
    x_agent_name: str = Header(""),
    authorization: str = Header(""),
) -> dict:
    """上传本任务的全部 PDF；保存成功后立即启动正式 AI 表写回。"""
    _require_customs_agent_token(authorization)
    from app.services import customs_tasks as customs_task_queue, customs_task_runner

    payload: list[tuple[str, bytes]] = []
    for item in files:
        content = await item.read(customs_task_queue.MAX_FILE_BYTES + 1)
        if len(content) > customs_task_queue.MAX_FILE_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"{item.filename or 'PDF'} 超过单文件大小限制",
            )
        payload.append((item.filename or "upload.pdf", content))
    try:
        task = customs_task_queue.save_uploads(
            task_id, payload, agent_name=agent_name or x_agent_name
        )
    except customs_task_queue.TaskNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except customs_task_queue.TaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    customs_task_runner.start(task_id)
    return {
        "taskId": task["id"],
        "state": task["state"],
        "fileCount": task["fileCount"],
        "statusUrl": f"/api/customs/tasks/{task['id']}",
    }


@app.post("/api/customs/tasks/{task_id}/report-error")
def report_customs_task_error(
    task_id: str,
    request: CustomsTaskErrorRequest,
    authorization: str = Header(""),
    x_agent_name: str = Header(""),
) -> dict:
    """财务电脑无法继续任务时上报业务原因，不触发 PDF 解析。"""
    _require_customs_agent_token(authorization)
    header_agent = str(x_agent_name or "").strip()
    if not header_agent or not secrets.compare_digest(header_agent, request.agentName):
        raise HTTPException(status_code=403, detail="X-Agent-Name 与请求体 agentName 不一致")
    from app.services import customs_tasks as customs_task_queue

    try:
        task = customs_task_queue.report_task_error(
            task_id,
            agent_name=request.agentName,
            error_code=request.errorCode,
            message=request.message,
            retryable=request.retryable,
            reported_at=request.reportedAt.isoformat(),
            details=request.details,
        )
    except customs_task_queue.TaskNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except customs_task_queue.TaskForbidden as exc:
        raise HTTPException(status_code=403, detail=str(exc)) from exc
    except customs_task_queue.TaskConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "taskId": task["id"],
        "state": task["state"],
        "errorCode": task["errorCode"],
        "message": task["errorMessage"],
        "retryable": task["errorRetryable"],
        "errorReportedAt": task["errorReportedAt"],
        "errorReceivedAt": task["errorReceivedAt"],
        "agentName": task["errorAgentName"],
    }


@app.get("/api/customs/tasks/{task_id}")
def get_customs_task(task_id: str, authorization: str = Header("")) -> dict:
    _require_customs_agent_token(authorization)
    from app.services import customs_tasks as customs_task_queue

    try:
        return customs_task_queue.get_task(task_id)
    except customs_task_queue.TaskNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@app.get("/api/customs/tasks/{task_id}/result")
def get_customs_task_result(task_id: str, authorization: str = Header("")) -> dict:
    _require_customs_agent_token(authorization)
    from app.services import customs_tasks as customs_task_queue

    try:
        task = customs_task_queue.get_task(task_id)
        path = customs_task_queue._task_dir(task_id) / "result.json"
    except customs_task_queue.TaskNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    if not path.exists():
        raise HTTPException(
            status_code=404,
            detail=f"结果尚未生成，当前状态：{task.get('state')}",
        )
    return json.loads(path.read_text(encoding="utf-8"))


@app.get("/api/erp/sync")
def erp_sync_status() -> dict:
    return erp_sync.status()


@app.post("/api/erp/sync")
def erp_sync_start(lists_only: bool = False, refresh_days: float = 7.0) -> dict:
    return erp_sync.start(refresh_days=refresh_days, lists_only=lists_only)


@app.get("/api/sync/target")
def sync_target() -> dict:
    from app.services.feishu_client import workbench_table

    app_token, table_id = workbench_table(required=False)
    return {
        "table": "迈拓财务部门数据 副本 / 2026年报关数据",
        "appToken": app_token,
        "tableId": table_id,
    }


@app.post("/api/lookup/parse")
async def lookup_parse(text: str = Form(""), use_ai: bool = Form(False)) -> dict:
    result = lookup.parse_paste(text)
    if result["rows"] and len(result["rows"]) <= 50:
        enriched = await run_in_threadpool(lookup.enrich_rows, result["rows"])
        result["rows"] = enriched["rows"]
        result["warnings"].extend(enriched["warnings"])
    return result


@app.post("/api/lookup/enrich")
async def lookup_enrich(rows: str = Form("[]")) -> dict:
    try:
        items = json.loads(rows)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="商品行格式不正确") from exc
    if not isinstance(items, list) or not 1 <= len(items) <= 50:
        raise HTTPException(status_code=400, detail="每次可补全 1 至 50 行")
    return await run_in_threadpool(lookup.enrich_rows, items)


@app.post("/api/lookup")
async def lookup_start(rows: str = Form("[]"), use_ai: bool = Form(False)) -> dict:
    try:
        items = json.loads(rows or "[]")
    except ValueError:
        raise HTTPException(status_code=400, detail="商品行格式不正确")
    declarations = lookup.declarations_from_rows(items)
    if not declarations:
        raise HTTPException(status_code=400, detail="请至少填写一行报关商品")
    if len(declarations) > 50:
        raise HTTPException(status_code=400, detail="逐条核验每次最多 50 条；同一完整出运批次超过上限时请使用模板文件")
    errors = lookup.validate_rows(items)
    if errors:
        raise HTTPException(status_code=400, detail="；".join(errors[:5]))
    job_id = workspace.start_job(kind="lookup", declarations=declarations, use_ai=use_ai)
    return {"id": job_id, "status": f"/api/jobs/{job_id}"}


@app.post("/api/lookup/with-evidence")
async def lookup_with_evidence(rows: str = Form("[]"), files: list[UploadFile] = File(default=[])) -> dict:
    try:
        items = json.loads(rows)
    except ValueError:
        raise HTTPException(status_code=400, detail="商品行格式不正确")
    declarations = lookup.declarations_from_rows(items)
    if not 1 <= len(declarations) <= 50:
        raise HTTPException(status_code=400, detail="逐条核验须有 1 至 50 条商品行")
    errors = lookup.validate_rows(items)
    if errors:
        raise HTTPException(status_code=400, detail="；".join(errors[:5]))
    payload = [(item.filename or "evidence", await item.read()) for item in files]
    job_id = workspace.start_job(kind="lookup", declarations=declarations, files=payload)
    return {"id": job_id}


@app.post("/api/lookup/import")
async def lookup_import(file: UploadFile = File(...)) -> dict:
    from app.services.template_import import read_template
    from tempfile import NamedTemporaryFile
    if not (file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传固定模板 .xlsx")
    with NamedTemporaryFile(suffix=".xlsx", delete=False) as temp:
        temp.write(await file.read())
        name = temp.name
    try:
        rows, errors = read_template(Path(name), limit=50)
        return {"rows": rows, "errors": errors}
    finally:
        Path(name).unlink(missing_ok=True)


@app.get("/api/template")
def download_template():
    from app.services.template_import import template_bytes
    return Response(template_bytes(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename=feishu-2026-ai-copy-template.xlsx"})


@app.post("/api/template/check")
async def check_template(file: UploadFile = File(...)) -> dict:
    from app.services.template_import import read_template
    from tempfile import NamedTemporaryFile
    if not (file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 .xlsx 模板")
    content = await file.read()
    with NamedTemporaryFile(suffix=".xlsx", delete=False) as temp:
        temp.write(content)
        name = temp.name
    try:
        rows, errors = read_template(Path(name))
        return {"rows": len(rows), "errors": errors}
    finally:
        Path(name).unlink(missing_ok=True)


@app.get("/api/jobs")
def jobs() -> dict:
    return {"jobs": workspace.list_jobs()}


@app.post("/api/jobs")
async def create_job(files: list[UploadFile] = File(...), use_ai: bool = Form(False)) -> dict:
    payload: list[tuple[str, bytes]] = []
    for item in files:
        payload.append((item.filename or "upload", await item.read()))
    if not payload:
        raise HTTPException(status_code=400, detail="没有收到文件")
    job_id = workspace.start_job(kind="files", files=payload, use_ai=use_ai)
    return {"id": job_id, "status": f"/api/jobs/{job_id}"}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict:
    status = workspace.get_job(job_id)
    if not status:
        raise HTTPException(status_code=404, detail="任务不存在")
    return status


@app.get("/api/jobs/{job_id}/rows")
def job_rows(job_id: str) -> dict:
    rows = workspace.result_rows(job_id)
    if rows is None:
        raise HTTPException(status_code=404, detail="结果尚未生成")
    from app.services.review import refresh_source_links
    return {"rows": [refresh_source_links(row) for row in rows]}


@app.get("/api/exceptions")
def exception_records() -> dict:
    return {"rows": workspace.all_exceptions()}


@app.post("/api/jobs/{job_id}/flag-exception")
def job_flag_exception(job_id: str, request: FlagExceptionRequest) -> dict:
    try:
        return workspace.flag_exception(job_id, request.index, request.operator, request.reason)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/{job_id}/resubmit-exception")
def job_resubmit_exception(job_id: str, request: ResubmitExceptionRequest) -> dict:
    try:
        return workspace.resubmit_exception(job_id, request.index, request.operator)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/{job_id}/override")
def job_override(job_id: str, request: OverrideRequest) -> dict:
    try:
        return workspace.override_row(job_id, request.index, request.amount, request.operator, request.reason)
    except (ValueError, ArithmeticError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/{job_id}/reviewed")
def job_reviewed(job_id: str, request: ReviewRequest) -> dict:
    try:
        return workspace.mark_reviewed(job_id, request.index, request.operator)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/{job_id}/rerun")
async def rerun_job(job_id: str, rows: str = Form("[]")) -> dict:
    prior = workspace.get_job(job_id)
    if not prior:
        raise HTTPException(status_code=404, detail="原任务不存在")
    try:
        items = json.loads(rows)
    except ValueError:
        raise HTTPException(status_code=400, detail="商品行格式不正确")
    declarations = lookup.declarations_from_rows(items)
    if not declarations:
        raise HTTPException(status_code=400, detail="没有可重算的商品行")
    source_path = workspace.JOBS_DIR / job_id / "source_records.json"
    sources = json.loads(source_path.read_text(encoding="utf-8")) if source_path.exists() else None
    new_id = workspace.start_job(kind=prior.get("kind", "lookup"), declarations=declarations,
                                 source_records=sources)
    workspace.append_audit(job_id, "rerun", "", {"newJobId": new_id})
    return {"id": new_id}


@app.get("/api/jobs/{job_id}/audit")
def job_audit(job_id: str) -> dict:
    if not workspace.get_job(job_id):
        raise HTTPException(status_code=404, detail="任务不存在")
    return {"events": workspace.audit(job_id)}


@app.get("/api/jobs/{job_id}/input")
def job_input(job_id: str) -> dict:
    path = workspace.JOBS_DIR / job_id / "input_rows.json"
    if not path.exists():
        return {"rows": [], "editable": False}
    return {"rows": json.loads(path.read_text(encoding="utf-8")), "editable": True}


@app.get("/api/jobs/{job_id}/erp-evidence/{index}")
def job_erp_evidence(job_id: str, index: int, request: Request,
                     x_workbench_evidence_token: str = Header("")) -> dict:
    _require_evidence_token(request, x_workbench_evidence_token)
    rows = workspace.result_rows(job_id)
    if rows is None or not 0 <= index < len(rows):
        raise HTTPException(status_code=404, detail="结果记录不存在")
    from app.services.review import erp_sources
    row = rows[index]
    sources = erp_sources(row.get("采购订单号", ""), row.get("合同号_1", ""))
    return {"purchase": sources["purchase"], "shipment": sources["shipment"]}


@app.get("/api/jobs/{job_id}/evidence")
def job_evidence(job_id: str) -> dict:
    root = workspace.JOBS_DIR / job_id / "uploads"
    return {"files": [{"name": p.name, "url": f"/api/jobs/{job_id}/evidence/{p.name}"}
                      for p in root.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]} if root.exists() else {"files": []}


@app.get("/api/jobs/{job_id}/evidence-comparison")
def evidence_comparison(job_id: str, request: Request,
                        x_workbench_evidence_token: str = Header("")) -> dict:
    _require_evidence_token(request, x_workbench_evidence_token)
    path = workspace.JOBS_DIR / job_id / "evidence_comparison.json"
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {"state": "unavailable", "rows": []}


@app.get("/api/jobs/{job_id}/evidence/{name}")
def evidence_file(job_id: str, name: str):
    safe = Path(name).name
    path = workspace.JOBS_DIR / job_id / "uploads" / safe
    if not path.is_file() or path.suffix.lower() != ".pdf":
        raise HTTPException(status_code=404, detail="凭证不存在")
    return FileResponse(path, media_type="application/pdf", filename=safe)


@app.post("/api/jobs/{job_id}/cancel")
def job_cancel(job_id: str) -> dict:
    status = workspace.cancel_job(job_id)
    if not status:
        raise HTTPException(status_code=404, detail="任务不存在")
    return status


@app.get("/api/jobs/{job_id}/download/{kind}")
def job_download(job_id: str, kind: str):
    path = workspace.download_path(job_id, kind)
    if not path:
        raise HTTPException(status_code=404, detail="文件尚未生成")
    return FileResponse(path, filename=path.name)


@app.post("/api/jobs/{job_id}/feishu-comparison")
def feishu_comparison_start(job_id: str) -> dict:
    from app.services import feishu_compare

    return feishu_compare.start(job_id)


@app.get("/api/jobs/{job_id}/feishu-comparison")
def feishu_comparison_get(job_id: str) -> dict:
    from app.services import feishu_compare

    result = feishu_compare.get(job_id)
    if result is None:
        raise HTTPException(status_code=404, detail="核对尚未开始")
    return result


@app.post("/api/sync/scan")
async def sync_scan(
    start_date: str = Form(""),
    end_date: str = Form(""),
    declarations: str = Form(""),
    contracts: str = Form(""),
    limit: int = Form(50),
) -> dict:
    from app.services import feishu_workflow
    from app.services.feishu_client import FeishuError
    import urllib.error

    try:
        return feishu_workflow.start_scan(
            start_date=start_date, end_date=end_date,
            declarations=declarations, contracts=contracts, limit=limit,
        )
    except (FeishuError, urllib.error.URLError) as exc:
        raise HTTPException(status_code=503, detail=f"飞书读取暂不可用：{exc}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/sync/preview")
async def sync_preview(
    start_date: str = Form(""), end_date: str = Form(""), declarations: str = Form(""),
    contracts: str = Form(""), limit: int = Form(50),
) -> dict:
    from app.services import feishu_workflow
    from app.services.feishu_client import FeishuError
    import urllib.error
    try:
        result = feishu_workflow.select_records(start_date=start_date, end_date=end_date,
                                                declarations=declarations, contracts=contracts, limit=limit)
    except (FeishuError, urllib.error.URLError) as exc:
        raise HTTPException(status_code=503, detail=f"飞书读取暂不可用：{exc}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    result["selected"] = [{k: value for k, value in item.items() if k not in ("pdf原件", "pdfTokens")}
                          for item in result["selected"]]
    result["context"] = [{k: value for k, value in item.items() if k not in ("pdf原件", "pdfTokens")}
                         for item in result["context"]]
    return result


@app.get("/api/sync/{job_id}/plan")
def sync_plan(job_id: str) -> dict:
    from app.services import feishu_workflow

    plan = feishu_workflow.plan(job_id)
    if plan is None:
        raise HTTPException(status_code=404, detail="没有可同步的结果")
    return plan


@app.post("/api/sync/{job_id}/mapping")
def sync_mapping(job_id: str, request: SourceMappingRequest) -> dict:
    from app.services import feishu_workflow
    try:
        return feishu_workflow.set_mapping(job_id, request.row_index, request.source_id, request.operator)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/sync/{job_id}/push")
def sync_push(job_id: str, request: Request, operator: str = Form(""),
              x_workbench_write_token: str = Header("")) -> dict:
    _require_write_token(request, x_workbench_write_token)
    from app.services import feishu_workflow
    try:
        return feishu_workflow.push(job_id, operator)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/")
def index():
    index_file = WEB_DIST / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return JSONResponse({"message": "前端尚未构建，请在 web 目录运行 npm install && npm run build",
                         "api": "/api/health", "version": APP_VERSION})
