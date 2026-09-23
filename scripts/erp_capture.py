"""打开睿贝页面并抓取其 XHR 接口，用于确定全量数据的取数方式。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from playwright.sync_api import sync_playwright  # noqa: E402

from app.services.erp_web import capture_xhr, launch_browser, login  # noqa: E402


def main() -> None:
    tag = sys.argv[1] if len(sys.argv) > 1 else "sale_order"
    path = sys.argv[2] if len(sys.argv) > 2 else "/saleOrder?menuCode=80300"
    out_dir = PROJECT_ROOT / ".cache" / "_probe"
    with sync_playwright() as playwright:
        browser = launch_browser(playwright)
        page = browser.new_page()
        login(page)
        records = capture_xhr(page, path, out_dir, tag, wait_ms=8000)
        print(f"captured {len(records)} requests -> {out_dir / (tag + '.json')}")
        for record in records:
            print(" ", record["method"], record["status"], record["url"][:140])
        browser.close()


if __name__ == "__main__":
    main()
