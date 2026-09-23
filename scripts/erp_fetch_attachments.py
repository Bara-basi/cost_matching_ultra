"""批量下载采购附件（只取含「入库单」的）。

用法：
    python scripts/erp_fetch_attachments.py --limit 3      # 试跑 3 个
    python scripts/erp_fetch_attachments.py                # 全量（含入库单的采购单）
    python scripts/erp_fetch_attachments.py --force        # 重新下载
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

from app.services.erp_attachments import (  # noqa: E402
    GRN_KEYWORD,
    download_purchase_attachments,
    load_manifest,
    save_manifest,
)
from app.services.erp_cache import CACHE_ROOT, now_iso, read_jsonl  # noqa: E402
from app.services.erp_lists import PURCHASES_WITH_GRN  # noqa: E402
from app.services.erp_replay import ErpSession, columns_to_rows  # noqa: E402
from app.services.erp_web import launch_browser, login  # noqa: E402


def grn_candidates() -> list[dict]:
    """取「含入库单附件」的采购单：用 grnChoose=Y 的列表筛选结果。"""
    cached = CACHE_ROOT / "purchases" / "with_grn.jsonl"
    if cached.exists():
        return read_jsonl(cached)
    rows: list[dict] = []
    with ErpSession() as session:
        page = 1
        while page <= 50:
            payload = session.fetch_json(PURCHASES_WITH_GRN.path, {**PURCHASES_WITH_GRN.base_query, "p": page})
            page_rows = columns_to_rows(payload.get("root"))
            if not page_rows:
                break
            rows.extend(page_rows)
            total = int(payload.get("total") or 0)
            if total and len(rows) >= total:
                break
            page += 1
    from app.services.erp_cache import write_jsonl

    write_jsonl(cached, rows)
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--pause", type=float, default=0.2)
    args = parser.parse_args()

    candidates = grn_candidates()
    if args.limit:
        candidates = candidates[: args.limit]
    print(f"含入库单的采购单 {len(candidates)} 个", flush=True)

    manifest = load_manifest()
    results = []
    with sync_playwright() as playwright:
        browser = launch_browser(playwright)
        page = browser.new_page(accept_downloads=True)
        login(page)
        for index, row in enumerate(candidates, 1):
            code = str(row.get("purchase_code") or "").strip()
            pid = str(row.get("purchase_id") or "").strip()
            if not pid:
                continue
            try:
                result = download_purchase_attachments(page, pid, code, overwrite=args.force)
            except Exception as exc:  # noqa: BLE001
                result = {
                    "purchaseCode": code,
                    "purchaseId": pid,
                    "status": "download-error",
                    "error": str(exc)[:200],
                    "hasGrn": False,
                }
            manifest[code] = result
            results.append(result)
            if index % 5 == 0 or index == len(candidates):
                ok = sum(1 for r in results if r.get("status") in ("downloaded", "cached"))
                grn = sum(1 for r in results if r.get("hasGrn"))
                print(
                    f"  [{index}/{len(candidates)}] 成功={ok} 含入库单={grn} 最近={code} "
                    f"状态={result.get('status')}",
                    flush=True,
                )
                save_manifest(manifest)
            time.sleep(args.pause)
        browser.close()

    save_manifest(manifest)
    summary = {
        "finishedAt": now_iso(),
        "total": len(candidates),
        "downloaded": sum(1 for r in results if r.get("status") == "downloaded"),
        "cached": sum(1 for r in results if r.get("status") == "cached"),
        "errors": sum(1 for r in results if r.get("status") == "download-error"),
        "withGrn": sum(1 for r in results if r.get("hasGrn")),
    }
    (CACHE_ROOT / "reports").mkdir(parents=True, exist_ok=True)
    (CACHE_ROOT / "reports" / "attachment_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
