r"""飞书人工结果核对（只读，不覆盖系统结果）。

判定口径（2026-09-28 用户确认）：
1. 采购金额相同 → 正确；
2. 采购金额不同、但只是**同一采购组内分配位置不同**（该组合计一致）
   → 「正常，但额外金额分配不一致」；
3. 报关金额与采购金额都允许行级不同，只要**合计一致**就按正常处理，分配差异不给提示。

比较单元 = 报关单号 + 采购单核心（合同号（应收表格））。
"""
from __future__ import annotations

import collections
import json
import re
import threading
from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.services.erp_cache import CACHE_ROOT
from app.services.feishu_client import FeishuClient, workbench_table
from app.services.workspace import JOBS_DIR, result_rows

TOLERANCE = Decimal("1")
_LOCK = threading.Lock()
_RUNNING: set[str] = set()


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        return text(value.get("value", value.get("text", value.get("name", ""))))
    if isinstance(value, list):
        return "&".join(filter(None, (text(item) for item in value)))
    return str(value).strip()


def number(value) -> Decimal:
    """飞书的数值字段可能是 `[15740]`、`15740`、`1,234.50` 等写法，取第一个数值。"""
    raw = text(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    if not match:
        return Decimal(0)
    try:
        return Decimal(match.group(0))
    except InvalidOperation:
        return Decimal(0)


def core(value) -> str:
    """采购单核心：去掉 PI- 前缀、工厂后缀与 ADD 尾缀。"""
    raw = re.sub(r"^PI-", "", text(value).upper())
    parts = raw.split("-")
    while parts:
        token = parts[-1]
        if token.isalpha() and token.isascii() and len(token) <= 5:
            parts.pop()
            continue
        break
    return "-".join(parts).split("-ADD")[0]


def _path(job_id: str) -> Path:
    return JOBS_DIR / job_id / "feishu_comparison.json"


def get(job_id: str) -> dict | None:
    path = _path(job_id)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None


def start(job_id: str) -> dict:
    if not result_rows(job_id):
        raise ValueError("本轮还没有结果，不能核对")
    if job_id in _RUNNING:
        return {"state": "running", "message": "正在核对"}
    with _LOCK:
        _RUNNING.add(job_id)
    _path(job_id).write_text(
        json.dumps({"state": "running", "message": "正在读取飞书人工记录…"}, ensure_ascii=False),
        encoding="utf-8",
    )
    threading.Thread(target=_run, args=(job_id,), daemon=True).start()
    return {"state": "running", "message": "已开始核对"}


def _fetch(declarations: set[str]) -> list[dict]:
    client = FeishuClient()
    app, table = workbench_table(required=False)
    if not app or not table:
        raise ValueError("飞书目标表未配置，请联系管理员")
    out: list[dict] = []
    # 副本表的列会变动。field_names 中只要有一列不存在，飞书就返回 FieldNameNotFound。
    for record in client.iter_records(app, table):
        data = record.get("fields") or {}
        if text(data.get("报关单号")) in declarations:
            out.append(record)
    return out


def _run(job_id: str) -> None:
    try:
        ours = result_rows(job_id) or []
        declarations = {str(row.get("报关单号") or "") for row in ours if row.get("报关单号")}
        records = _fetch(declarations)

        def group_key(decl: str, contract: str) -> tuple[str, str]:
            return (decl, core(contract))

        mine: dict[tuple[str, str], dict] = collections.defaultdict(
            lambda: {"采购": Decimal(0), "报关": Decimal(0), "行数": 0}
        )
        for row in ours:
            key = group_key(str(row.get("报关单号") or ""), str(row.get("合同号（应收表格）") or row.get("采购订单号") or ""))
            mine[key]["采购"] += number(row.get("采购金额"))
            mine[key]["报关"] += number(row.get("报关金额"))
            mine[key]["行数"] += 1

        theirs: dict[tuple[str, str], dict] = collections.defaultdict(
            lambda: {"采购": Decimal(0), "报关": Decimal(0), "行数": 0, "id": "", "link": ""}
        )
        for record in records:
            data = record.get("fields") or {}
            key = group_key(text(data.get("报关单号")), text(data.get("合同号（应收表格）")) or text(data.get("合同号_1")))
            theirs[key]["采购"] += number(data.get("采购金额"))
            theirs[key]["报关"] += number(data.get("报关金额"))
            theirs[key]["行数"] += 1
            theirs[key]["id"] = record.get("record_id") or ""

        rows: list[dict] = []
        for key, mine_group in sorted(mine.items()):
            decl, purchase = key
            other = theirs.get(key)
            base = {
                "报关单号": decl,
                "合同号（应收表格）": purchase,
                "合同号_1": purchase,
                "系统采购金额": f"{mine_group['采购']:.2f}",
                "飞书采购金额": f"{other['采购']:.2f}" if other else "",
                "系统报关金额": f"{mine_group['报关']:.2f}",
                "飞书报关金额": f"{other['报关']:.2f}" if other else "",
                "飞书记录ID": other["id"] if other else "",
                "飞书链接": "",
            }
            if other is None:
                base.update({"核对状态": "需人工确认", "异常类型": "飞书没有这一行",
                             "异常明细": "飞书里找不到这张采购单的记录，请人工核对。"})
            elif abs(mine_group["采购"] - other["采购"]) <= TOLERANCE:
                if mine_group["行数"] == other["行数"]:
                    base.update({"核对状态": "正常", "异常类型": "正常", "异常明细": ""})
                else:
                    base.update({"核对状态": "正常", "异常类型": "正常，但额外金额分配不一致",
                                 "异常明细": "采购金额合计一致，只是分摊到了不同行。"})
            else:
                base.update({"核对状态": "需人工确认", "异常类型": "采购金额对不上",
                             "异常明细": "我方与飞书的采购金额合计不一致，请核对入库结算与费用。"})
            rows.append(base)

        summary = {
            "consistent": sum(1 for row in rows if row["核对状态"] == "正常"),
            "exceptions": sum(1 for row in rows if row["核对状态"] != "正常"),
            "localRows": len(ours),
            "onlineRows": len(records),
        }
        _path(job_id).write_text(
            json.dumps(
                {
                    "state": "complete", "rows": rows, "summary": summary,
                    "amountBasis": "系统采购金额（采购单合计） vs 飞书采购金额（采购单合计）",
                    "message": f"{summary['consistent']} 条正常 · {summary['exceptions']} 条需要人工确认",
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except Exception as exc:  # noqa: BLE001
        _path(job_id).write_text(
            json.dumps({"state": "failed", "message": str(exc)[:200]}, ensure_ascii=False),
            encoding="utf-8",
        )
    finally:
        _RUNNING.discard(job_id)
