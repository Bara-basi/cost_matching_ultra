r"""睿贝数据同步：把 ERP 三类单据补全并更新到本地 `.cache/erp`（断点续跑）。

阶段：
  1) 列表刷新：外销订单 / 出运单 / 采购订单。走 ERP 直连，分页抓全后**按主键合并**
     （已存在的更新、新出现的追加），保证不漏也不丢历史；
  2) 出运明细：走 MCP `shipment.find`，缺的必抓；已有的按「文件时间超过 N 天」重抓一遍，
     保证旧数据也能更新；
  3) 采购附件（默认 `--with-attachments`）：走 MCP 抓采购单附件清单并下载其中的「入库单」，
     已下载的跳过（断点续跑），这是成本匹配的唯一凭证来源；
  4) 入库单解析（默认启用，可用 `--no-grn-parse` 跳过）：把新下载的入库单解析成结构化金额；
  5) 重建出运产品行索引（拆单/成本匹配直接消费）。

进度写在 `.cache/erp/sync_state.json`，工作台启动时可据此判断是否需要同步。

用法：
  & '.\.venv\Scripts\python.exe' scripts\erp_sync.py                 # 默认：列表 + 出运明细补漏
  & '.\.venv\Scripts\python.exe' scripts\erp_sync.py --refresh-days 7
  & '.\.venv\Scripts\python.exe' scripts\erp_sync.py --lists-only
  & '.\.venv\Scripts\python.exe' scripts\erp_sync.py --no-attachments      # 跳过附件阶段
  & '.\.venv\Scripts\python.exe' scripts\erp_sync.py --with-grn-parse      # 顺带解析新入库单
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT, now_iso, read_jsonl, write_json, write_jsonl  # noqa: E402
from app.services.erp_lists import PURCHASES, SALE_ORDERS, SHIPMENTS  # noqa: E402

STATE_PATH = CACHE_ROOT / "sync_state.json"
TARGETS = {
    "sale_orders": CACHE_ROOT / "sale_orders" / "orders.jsonl",
    "shipments": CACHE_ROOT / "shipments" / "shipments.jsonl",
    "purchases": CACHE_ROOT / "purchases" / "purchases.jsonl",
}
ID_FIELDS = {"sale_orders": "orderId", "shipments": "shipmentId", "purchases": "purchase_id"}


class SyncInterrupted(RuntimeError):
    """连续远端失败：暂停自动重试，等待手动恢复。"""


def write_state(**values) -> dict:
    state = {}
    if STATE_PATH.exists():
        try:
            state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except ValueError:
            state = {}
    state.update(values)
    state["updatedAt"] = now_iso()
    write_json(STATE_PATH, state)
    return state


def merge_rows(old: list[dict], new: list[dict], id_field: str) -> tuple[list[dict], int, int]:
    """按主键合并：返回 (合并结果, 新增数, 更新数)。"""
    by_id: dict[str, dict] = {}
    for row in old:
        key = str(row.get(id_field) or "").strip()
        if key:
            by_id[key] = row
    added = updated = 0
    for row in new:
        key = str(row.get(id_field) or "").strip()
        if not key:
            continue
        if key in by_id:
            combined = {**by_id[key], **row}
            if by_id[key] != combined:
                updated += 1
            by_id[key] = combined
        else:
            added += 1
            by_id[key] = row
    return list(by_id.values()), added, updated


def sync_lists(state: dict) -> None:
    from app.services.erp_replay import ErpSession
    from scripts.erp_fetch_all import build_graph, build_indexes, fetch_list, year_of

    with ErpSession() as session:
        for spec in (SALE_ORDERS, SHIPMENTS, PURCHASES):
            write_state(phase=f"列表：{spec.name}", message=f"正在刷新 {spec.name}")
            fresh = fetch_list(session, spec)
            if spec.name == "sale_orders":
                for row in fresh:
                    row["_year"] = year_of(row, ["orderCode"], ["contractDate"])
            elif spec.name == "shipments":
                for row in fresh:
                    row["_year"] = year_of(row, ["invoiceCode", "purchaseCode"], ["shipDate"])
            else:
                for row in fresh:
                    row["_year"] = year_of(row, ["purchase_code", "orderCode"], ["purchase_date"])
            old = read_jsonl(TARGETS[spec.name])
            merged, added, updated = merge_rows(old, fresh, ID_FIELDS[spec.name])
            current_ids = {str(row.get(ID_FIELDS[spec.name]) or "") for row in fresh}
            # 保留历史缓存，但区分远端当前列表已不再提供的记录。
            for row in merged:
                row["_remotePresent"] = str(row.get(ID_FIELDS[spec.name]) or "") in current_ids
            write_jsonl(TARGETS[spec.name], merged)
            state.setdefault("lists", {})[spec.name] = {
                "total": len(merged), "added": added, "updated": updated,
            }
            write_state(phase=f"列表：{spec.name}", message=f"{spec.name} 共 {len(merged)} 条（+{added}/~{updated}）",
                        lists=state.get("lists"))
            print(f"  {spec.name}: 共 {len(merged)} 条（新增 {added} / 更新 {updated}）", flush=True)

    orders = read_jsonl(TARGETS["sale_orders"])
    shipments = read_jsonl(TARGETS["shipments"])
    purchases = read_jsonl(TARGETS["purchases"])
    indexes = build_indexes(orders, shipments, purchases)
    write_json(CACHE_ROOT / "sale_orders" / "index.json", indexes["orders"])
    write_json(CACHE_ROOT / "shipments" / "index.json", indexes["shipments"])
    write_json(CACHE_ROOT / "purchases" / "index.json", indexes["purchases"])
    write_json(CACHE_ROOT / "links" / "order_graph.json", build_graph(orders, shipments, purchases))


def sync_shipment_details(refresh_days: float, pause: float) -> dict:
    """MCP 抓出运明细：缺的必抓，已有的按过期天数重抓。"""
    from app.services.feishu_client import get_config
    from scripts.mcp_erp import McpClient

    detail_dir = CACHE_ROOT / "details" / "shipments"
    detail_dir.mkdir(parents=True, exist_ok=True)
    shipments = read_jsonl(TARGETS["shipments"])
    now = time.time()
    todo: list[str] = []
    for row in shipments:
        code = str(row.get("invoiceCode") or "").strip()
        if not code:
            continue
        safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in code)
        path = detail_dir / f"{safe}.json"
        if not path.exists():
            todo.append(code)
            continue
        age_days = (now - path.stat().st_mtime) / 86400
        if refresh_days < 0 or (refresh_days > 0 and age_days >= refresh_days):
            todo.append(code)
    write_state(phase="出运明细", message=f"待抓 {len(todo)} 张出运单")
    print(f"出运明细：待抓 {len(todo)} 张", flush=True)
    if not todo:
        return {"requested": 0, "done": 0, "failed": 0}

    client = McpClient(get_config("ERP_API_KEY"))
    client.initialize()
    done = failed = 0
    consecutive_failures = 0
    interrupted = False
    for index, code in enumerate(todo, 1):
        try:
            text = client.call_text(
                "shipment.find",
                {"invoiceCode": code, "includeProduct": "true", "includeExpense": "true",
                 "includePurchaseExpense": "true", "includeForwarderExpense": "true"},
            )
            payload = json.loads(text)
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"  [{index}/{len(todo)}] {code} ERROR {exc}", flush=True)
            consecutive_failures += 1
            if consecutive_failures >= 3:
                interrupted = True
                break
            time.sleep(2)
            continue
        value = payload.get("value") or {}
        if not value:
            failed += 1
            consecutive_failures += 1
            if consecutive_failures >= 3:
                interrupted = True
                break
            continue
        consecutive_failures = 0
        safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in code)
        (detail_dir / f"{safe}.json").write_text(
            json.dumps(
                {
                    "invoiceCode": code,
                    "shipmentId": (value.get("shipmentBaseInfo") or {}).get("shipmentId"),
                    "baseInfo": value.get("shipmentBaseInfo") or {},
                    "productList": value.get("productList") or [],
                    "expenseList": value.get("expenseList") or [],
                    "purchaseExpenseList": value.get("purchaseExpenseList") or [],
                    "fetchedAt": now_iso(),
                },
                ensure_ascii=False, indent=1,
            ),
            encoding="utf-8",
        )
        done += 1
        time.sleep(pause)
        if index % 25 == 0 or index == len(todo):
            write_state(
                phase="出运明细",
                message=f"出运明细 {index}/{len(todo)}（成功 {done}、失败 {failed}）",
            )
            print(f"  [{index}/{len(todo)}] 成功={done} 失败={failed}", flush=True)
    return {"requested": len(todo), "done": done, "failed": failed, "interrupted": interrupted}


def main() -> None:
    # 强制 UTF-8 输出：子进程 stdout 常被重定向到文件/管道，Windows 默认 GBK 遇到
    # 代理字符（\ufffd）会直接抛 UnicodeEncodeError，把已完成的同步误标成 failed。
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass

    parser = argparse.ArgumentParser()
    parser.add_argument("--refresh-days", type=float, default=7.0,
                        help="出运明细缓存超过这么多天就重抓（0=只补缺，负数=全部刷新）")
    parser.add_argument("--pause", type=float, default=0.3)
    parser.add_argument("--lists-only", action="store_true")
    parser.add_argument("--resume", action="store_true", help="复用上次已完成的三单列表，从明细/附件阶段接续；上次列表不完整时仍重抓")
    parser.add_argument("--with-attachments", dest="attachments", action="store_true", default=True,
                        help="抓采购附件并下载入库单（默认开）")
    parser.add_argument("--no-attachments", dest="attachments", action="store_false",
                        help="跳过附件阶段")
    parser.add_argument("--with-grn-parse", dest="with_grn_parse", action="store_true", default=True,
                        help="把新下载的入库单解析成结构化金额（默认开）")
    parser.add_argument("--no-grn-parse", dest="with_grn_parse", action="store_false", help="只同步原始凭证，不解析；同步状态标为不完整")
    parser.add_argument("--attachments-limit", type=int, default=0, help="附件阶段最多处理多少个采购单（0=全部）")
    args = parser.parse_args()

    previous = json.loads(STATE_PATH.read_text(encoding="utf-8")) if args.resume and STATE_PATH.exists() else {}
    reuse_lists = bool(previous.get("listsComplete") and not previous.get("scopedSync")
                      and all(path.exists() for path in TARGETS.values()))
    started = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    write_state(state="running", startedAt=started, message="开始同步", phase="准备",
                listsComplete=False, detailsComplete=False, attachmentsComplete=False, parseComplete=False, fullSyncComplete=False, interrupted=False, scopedSync=None)
    try:
        if reuse_lists:
            print("[1/5] 断点接续：复用已完整刷新并建立索引的三类列表", flush=True)
        else:
            print("[1/5] 刷新三类单据列表", flush=True)
            sync_lists({})
        write_state(listsComplete=True)
        detail_stats = {"requested": 0, "done": 0, "failed": 0}
        if not args.lists_only:
            print("[2/5] 补抓出运明细（MCP）", flush=True)
            detail_stats = sync_shipment_details(args.refresh_days, args.pause)
            if detail_stats.get("interrupted"):
                write_state(details=detail_stats)
                raise SyncInterrupted("出运明细连续三次失败或为空，按要求暂停同步；成本凭证尚未同步完整")

        attachment_stats = {"skipped": True}
        if args.attachments and not args.lists_only:
            print("[3/5] 补抓采购附件（入库单）", flush=True)
            write_state(phase="采购附件", message="正在补抓采购附件（入库单）")
            attachment_stats = run_attachments(args.attachments_limit)

        if args.with_grn_parse and not args.lists_only:
            print("[4/5] 解析新入库单", flush=True)
            write_state(phase="入库单解析", message="正在解析新入库单")
            attachment_stats["parse"] = run_script([str(PROJECT_ROOT / "scripts" / "parse_grn.py")])
            write_state(phase="入库单筛选", message="正在选择有效入库单，排除旧版和重复附件")
            attachment_stats["selection"] = run_script([str(PROJECT_ROOT / "scripts" / "select_grn.py")])

        print("[5/5] 重建出运产品行索引", flush=True)
        write_state(phase="索引", message="正在重建出运产品行索引")
        from app.services.shipment_index import build as build_index

        index_stats = build_index()
        details_complete = not args.lists_only and not detail_stats.get("failed")
        attachments_complete = bool(args.attachments and not args.lists_only and not args.attachments_limit)
        parse_complete = bool(args.with_grn_parse and not args.lists_only)
        full_complete = details_complete and attachments_complete and parse_complete
        write_state(
            state="complete" if full_complete else "incomplete", finishedAt=now_iso(), phase="完成",
            detailsComplete=details_complete, attachmentsComplete=attachments_complete,
            parseComplete=parse_complete, fullSyncComplete=full_complete,
            message=f"{'同步完成' if full_complete else '同步不完整'}（出运明细成功 {detail_stats['done']}，失败 {detail_stats['failed']}）",
            details=detail_stats, attachments=attachment_stats, index=index_stats,
        )
        try:  # 兜底：控制台/日志的编码问题绝不能把「已完成的同步」改写成失败
            print(json.dumps({"details": detail_stats, "attachments": attachment_stats,
                              "index": index_stats}, ensure_ascii=False))
        except Exception:  # noqa: BLE001
            pass
    except Exception as exc:  # noqa: BLE001
        write_state(state="paused" if isinstance(exc, SyncInterrupted) else "failed",
                    interrupted=isinstance(exc, SyncInterrupted), finishedAt=now_iso(), message=str(exc)[:200])
        raise


def run_script(args: list[str]) -> dict:
    """跑一个子脚本（复用已有实现），把输出写进同步日志。"""
    process = subprocess.run(
        [sys.executable, *args],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env={**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"},
    )
    tail = (process.stdout or "").strip().splitlines()[-3:]
    summary = {}
    for line in reversed((process.stdout or "").splitlines()):
        try:
            value = json.loads(line)
            if isinstance(value, dict):
                summary = value
                break
        except ValueError:
            continue
    if process.returncode or summary.get("errors") or (Path(args[0]).name == "parse_grn.py" and (summary.get("异常") or summary.get("rule_failed") or summary.get("skipped"))):
        tail = ((process.stderr or "").strip().splitlines()[-3:] or tail)
        write_state(phase="子任务失败", message=f"{args[0]} 退出码 {process.returncode}")
        error_type = SyncInterrupted if summary.get("interrupted") else RuntimeError
        raise error_type(f"同步子任务未完成：{Path(args[0]).name}，退出码 {process.returncode}，失败统计 {summary}")
    return {"returncode": process.returncode, "tail": tail, "summary": summary}


def run_attachments(limit: int) -> dict:
    args = [str(PROJECT_ROOT / "scripts" / "erp_fetch_grn_mcp.py"), "--list-max-age-days", "0", "--force"]
    if limit:
        args += ["--limit", str(limit)]
    return run_script(args)


if __name__ == "__main__":
    main()
