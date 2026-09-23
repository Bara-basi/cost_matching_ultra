"""用 MCP 的 downloadUrl 下载单个附件，验证能否直接取到入库单。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from playwright.sync_api import sync_playwright  # noqa: E402

from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.erp_web import launch_browser, login  # noqa: E402
from scripts.mcp_erp import McpClient  # noqa: E402

OUT = CACHE_ROOT / "probe_attachments"


def main() -> None:
    token = sys.argv[1] if len(sys.argv) > 1 else ""
    code = sys.argv[2] if len(sys.argv) > 2 else "26MT-03T271-HD"
    OUT.mkdir(parents=True, exist_ok=True)

    client = McpClient(token)
    client.initialize()
    text = client.call_text("purchase.find", {"purchaseCode": code, "includeAttachment": "true"})
    payload = json.loads(text)
    attachments = payload["value"].get("attachmentList") or []
    grn = [a for a in attachments if "入库单" in a.get("attachmentName", "")]
    print(f"{code}: 附件 {len(attachments)} 个，其中入库单 {len(grn)} 个")

    results: list[str] = []
    with sync_playwright() as playwright:
        browser = launch_browser(playwright)
        page = browser.new_page(accept_downloads=True)
        login(page)
        for item in grn[:3]:
            url = item["downloadUrl"]
            name = item["attachmentName"]
            try:
                response = page.request.get(url, timeout=60_000)
                body = response.body()
                target = OUT / name
                target.write_bytes(body)
                results.append(f"  {name}: status={response.status} bytes={len(body)} -> {target.name}")
            except Exception as exc:  # noqa: BLE001
                results.append(f"  {name}: ERROR {exc}")
        browser.close()
    print("\n".join(results))
    (CACHE_ROOT / "reports").mkdir(parents=True, exist_ok=True)
    (CACHE_ROOT / "reports" / "mcp_download_probe.txt").write_text(
        "\n".join(results), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
