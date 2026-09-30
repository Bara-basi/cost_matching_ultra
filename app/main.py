r"""成本匹配工作台：FastAPI 服务（前端静态资源 + /api/*）。

启动：
  & '.\.venv\Scripts\python.exe' -m uvicorn app.main:app --host 127.0.0.1 --port 8765
"""
from __future__ import annotations

import json
import os
import secrets
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from pydantic import BaseModel

from app.services import erp_sync, lookup, workspace

PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_VERSION = "2026.09.30-api.2"

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


class SourceMappingRequest(BaseModel):
    row_index: int
    source_id: str
    operator: str


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


@app.on_event("startup")
def _startup() -> None:
    workspace.recover_interrupted()
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
    }


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
        raise HTTPException(status_code=400, detail="逐条核验每次最多 50 条")
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
    return {"rows": rows}


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
    from app.services.erp_cache_index import lookup
    row = rows[index]
    purchase = str(row.get("采购订单号") or "")
    shipment = str(row.get("合同号_1") or "")
    return {"purchase": lookup(purchase) if purchase else {},
            "shipment": lookup(shipment) if shipment else {}}


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
