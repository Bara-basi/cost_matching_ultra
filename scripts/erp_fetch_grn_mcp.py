"""用 MCP 抓取全部采购单的附件清单，并下载其中的「入库单」文件。

用法：
    python scripts/erp_fetch_grn_mcp.py --limit 5       # 试跑 5 个采购单
    python scripts/erp_fetch_grn_mcp.py                 # 全量（发生过入库的采购单）
    python scripts/erp_fetch_grn_mcp.py --force         # 忽略已抓清单
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from playwright.sync_api import sync_playwright  # noqa: E402

from app.services.erp_attachment_mcp import (  # noqa: E402
    GRN_DIR,
    download_file,
    grn_attachments,
    load_list,
    load_manifest,
    now_iso,
    parse_purchase_payload,
    safe_name,
    save_list,
    save_manifest,
)
from app.services.erp_cache import CACHE_ROOT, read_jsonl  # noqa: E402
from app.services.erp_web import launch_browser, login  # noqa: E402
from scripts.mcp_erp import McpClient  # noqa: E402


def num(value) -> float:
    try:
        return float(str(value or 0).replace(",", ""))
    except ValueError:
        return 0.0


def list_is_stale(cached: dict, max_age_days: float) -> bool:
    """附件清单缓存是否过期（判断依据：抓取时间；判不出时间就当过期）。"""
    from datetime import datetime, timezone

    if max_age_days <= 0:
        return True
    stamp = str(cached.get("fetchedAt") or "").strip()
    if not stamp:
        return True
    try:
        when = datetime.fromisoformat(stamp)
    except ValueError:
        return True
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    age_days = (datetime.now(timezone.utc) - when.astimezone(timezone.utc)).total_seconds() / 86400
    return age_days >= max_age_days


def candidates() -> list[dict]:
    """发生过入库（有入库数量或入库日期）的采购单，按入库时间倒序。"""
    rows = read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
    # 旧采购单可能在本月出运/补传附件；不能只看采购日期或入库标志。
    from app.services.shipment_index import normalise_purchase_code
    from app.services.contract_shipments import order_core
    import re
    # 先以真实采购列表为准；合写/批次别名只扩展到已存在的同核心采购单。
    wanted = [row for row in rows if str(row.get("purchase_code") or "").strip()
              and row.get("_remotePresent") is not False]
    known = {normalise_purchase_code(str(row["purchase_code"])) for row in wanted}
    by_core = {}
    for code in known:
        by_core.setdefault(order_core(code), set()).add(code)
    additional = set()
    for path in (CACHE_ROOT / "details" / "shipments").glob("*.json"):
        try:
            detail = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        for item in detail.get("productList") or []:
            raw_code = str(item.get("采购订单号") or "")
            # &ADD 是真实编号的一部分；先匹配完整编号再拆联合写法。
            if normalise_purchase_code(raw_code.strip()) in known:
                continue
            for token in re.split(r"[,，、;；&]+", raw_code):
                code = normalise_purchase_code(token.strip())
                if not code or code in known or order_core(code) in by_core:
                    continue
                # 不把「麦金3吨/增补」等备注当成采购编号发给接口。
                if re.fullmatch(r"\d{2}MT-\d{2}[A-Za-z]\d{3}[A-Za-z0-9-]*", code):
                    additional.add(code)
    wanted.extend({"purchase_code": code} for code in sorted(additional))
    wanted.sort(key=lambda r: str(r.get("lastGrnDate") or ""), reverse=True)
    return wanted


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--token", default="")
    parser.add_argument("--code", default="", help="只跑指定的采购单号（可多个，用逗号分隔）")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument(
        "--list-max-age-days",
        type=float,
        default=3.0,
        help="附件清单缓存超过这么多天就重抓（睿贝 tempFile 下载链接会过期，0=永不复用旧清单）",
    )
    parser.add_argument("--pause", type=float, default=0.25)
    args = parser.parse_args()

    from app.services.feishu_client import get_config

    token = args.token or get_config("ERP_API_KEY")
    todo = candidates()
    if args.code:
        wanted = {part.strip() for part in args.code.split(",") if part.strip()}
        rows = read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
        from_code = [r for r in rows if str(r.get("purchase_code") or "").strip() in wanted]
        known = {str(r.get("purchase_code") or "").strip() for r in from_code}
        # 单号不在采购单缓存里也照样抓一次（清单接口以单号为准）
        missing = [{"purchase_code": code} for code in sorted(wanted - known)]
        todo = from_code + missing
    if args.limit:
        todo = todo[: args.limit]
    print(f"待检查采购单 {len(todo)} 个", flush=True)

    client = McpClient(token)
    client.initialize()
    manifest = load_manifest()
    stats = {"scanned": 0, "withGrn": 0, "downloaded": 0, "skipped": 0, "errors": 0}
    consecutive_failures = 0
    interrupted = False
    download_failures = 0

    with sync_playwright() as playwright:
        browser = launch_browser(playwright)
        page = browser.new_page(accept_downloads=True)
        login(page)
        for index, row in enumerate(todo, 1):
            code = str(row.get("purchase_code") or "").strip()
            if not code:
                continue
            cached = None if args.force else load_list(code)
            if cached is not None and list_is_stale(cached, args.list_max_age_days):
                # 清单里的 downloadUrl 指向 tempFile，过期后服务端返回 200 + 空体
                # （下到 0 字节文件，成本匹配就永远「缺入库单」）→ 过期清单必须重抓。
                cached = None
            if cached is None:
                try:
                    text = client.call_text(
                        "purchase.find", {"purchaseCode": code, "includeAttachment": "true"}
                    )
                except Exception as exc:  # noqa: BLE001
                    stats["errors"] += 1
                    consecutive_failures += 1
                    if consecutive_failures >= 3:
                        interrupted = True
                        break
                    if index % 20 == 0:
                        print(f"  [{index}/{len(todo)}] {code} 清单失败 {exc}", flush=True)
                    continue
                value = parse_purchase_payload(text)
                if not value:
                    stats["errors"] += 1
                    consecutive_failures += 1
                    if consecutive_failures >= 3:
                        interrupted = True
                        break
                    continue
                consecutive_failures = 0
                cached = {
                    "purchaseCode": code,
                    "purchaseId": (value.get("purchaseBaseInfo") or {}).get("purchaseId"),
                    "attachmentList": value.get("attachmentList") or [],
                    "fetchedAt": now_iso(),
                }
                save_list(code, cached)
                time.sleep(args.pause)

            stats["scanned"] += 1
            grn_items = grn_attachments(cached)
            if not grn_items:
                continue
            stats["withGrn"] += 1
            entry = manifest.setdefault(code, {"purchaseCode": code, "files": []})
            entry["purchaseId"] = cached.get("purchaseId")
            for item in grn_items:
                name = item["name"]
                target = GRN_DIR / safe_name(code) / safe_name(name, "grn.xlsx")
                previous = next((record for record in reversed(entry["files"])
                                 if record.get("name") == name and record.get("status") == "downloaded"), {})
                if (target.exists() and target.stat().st_size > 0 and not args.force
                        and previous.get("sourceUrl") == item["url"]):
                    stats["skipped"] += 1
                    continue
                try:
                    size = download_file(page, item["url"], target)
                except Exception as exc:  # noqa: BLE001
                    stats["errors"] += 1
                    entry["files"].append({"name": name, "status": "error", "error": str(exc)[:120]})
                    download_failures += 1
                    if download_failures >= 3:
                        interrupted = True
                        break
                    continue
                if not size:
                    # 200 + 空体：下载链接已过期或会话失效，不能算成功
                    stats["errors"] += 1
                    entry["files"].append(
                        {"name": name, "status": "empty", "error": "空文件（下载链接可能已过期）"}
                    )
                    download_failures += 1
                    if download_failures >= 3:
                        interrupted = True
                        break
                    continue
                download_failures = 0
                stats["downloaded"] += 1
                entry["files"].append(
                    {"name": name, "status": "downloaded", "bytes": size, "path": str(target), "sourceUrl": item["url"]}
                )
            if interrupted:
                break
            if index % 20 == 0 or index == len(todo):
                save_manifest(manifest)
                print(
                    f"  [{index}/{len(todo)}] 已查={stats['scanned']} 含入库单={stats['withGrn']} "
                    f"下载={stats['downloaded']} 跳过={stats['skipped']} 错误={stats['errors']}",
                    flush=True,
                )
        browser.close()

    save_manifest(manifest)
    summary = {"finishedAt": now_iso(), "total": len(todo), "interrupted": interrupted, **stats}
    (CACHE_ROOT / "reports").mkdir(parents=True, exist_ok=True)
    (CACHE_ROOT / "reports" / "grn_mcp_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False))
    if stats["errors"] or interrupted:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
