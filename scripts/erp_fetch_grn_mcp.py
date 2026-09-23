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


def candidates() -> list[dict]:
    """发生过入库（有入库数量或入库日期）的采购单，按入库时间倒序。"""
    rows = read_jsonl(CACHE_ROOT / "purchases" / "purchases.jsonl")
    wanted = [r for r in rows if num(r.get("grnQuantity")) > 0 or str(r.get("lastGrnDate") or "").strip()]
    wanted.sort(key=lambda r: str(r.get("lastGrnDate") or ""), reverse=True)
    return wanted


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--token", default="")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--pause", type=float, default=0.25)
    args = parser.parse_args()

    from app.services.feishu_client import get_config

    token = args.token or get_config("ERP_API_KEY")
    todo = candidates()
    if args.limit:
        todo = todo[: args.limit]
    print(f"发生过入库的采购单 {len(todo)} 个", flush=True)

    client = McpClient(token)
    client.initialize()
    manifest = load_manifest()
    stats = {"scanned": 0, "withGrn": 0, "downloaded": 0, "skipped": 0, "errors": 0}

    with sync_playwright() as playwright:
        browser = launch_browser(playwright)
        page = browser.new_page(accept_downloads=True)
        login(page)
        for index, row in enumerate(todo, 1):
            code = str(row.get("purchase_code") or "").strip()
            if not code:
                continue
            cached = None if args.force else load_list(code)
            if cached is None:
                try:
                    text = client.call_text(
                        "purchase.find", {"purchaseCode": code, "includeAttachment": "true"}
                    )
                except Exception as exc:  # noqa: BLE001
                    stats["errors"] += 1
                    if index % 20 == 0:
                        print(f"  [{index}/{len(todo)}] {code} 清单失败 {exc}", flush=True)
                    continue
                value = parse_purchase_payload(text)
                if not value:
                    stats["errors"] += 1
                    continue
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
                if target.exists() and not args.force:
                    stats["skipped"] += 1
                    continue
                try:
                    size = download_file(page, item["url"], target)
                except Exception as exc:  # noqa: BLE001
                    stats["errors"] += 1
                    entry["files"].append({"name": name, "status": "error", "error": str(exc)[:120]})
                    continue
                stats["downloaded"] += 1
                entry["files"].append(
                    {"name": name, "status": "downloaded", "bytes": size, "path": str(target)}
                )
            if index % 20 == 0 or index == len(todo):
                save_manifest(manifest)
                print(
                    f"  [{index}/{len(todo)}] 已查={stats['scanned']} 含入库单={stats['withGrn']} "
                    f"下载={stats['downloaded']} 跳过={stats['skipped']} 错误={stats['errors']}",
                    flush=True,
                )
        browser.close()

    save_manifest(manifest)
    summary = {"finishedAt": now_iso(), "total": len(todo), **stats}
    (CACHE_ROOT / "reports").mkdir(parents=True, exist_ok=True)
    (CACHE_ROOT / "reports" / "grn_mcp_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
