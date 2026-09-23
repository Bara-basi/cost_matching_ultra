"""睿贝 ERP Web 端访问层：Playwright 登录 + 常用页面的 XHR 捕获。

睿贝同一账号同时只能一人登录，因此抓取应尽量一次性完成；
这里把「登录」和「浏览器启动」独立出来，便于脚本复用同一会话。
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Callable

from app.services.feishu_client import load_env

ERP_ROOT = os.environ.get("ERP_BASE_URL", "https://erp.mtholdinggroup.com").rstrip("/")

PAGES = {
    "sale_order": "/saleOrder?menuCode=80300",
    "purchase": "/purchase?menuCode=80400",
    "shipment": "/shipment?menuCode=80500",
}


def erp_credentials() -> tuple[str, str]:
    """从 .env / 环境变量取睿贝账号。"""
    env = load_env()
    user = os.environ.get("ERP_USERNAME") or env.get("ERP_USERNAME", "")
    password = os.environ.get("ERP_PASSWORD") or env.get("ERP_PASSWORD", "")
    if not user or not password:
        raise RuntimeError("缺少睿贝账号：ERP_USERNAME / ERP_PASSWORD")
    return user, password


def launch_browser(playwright, headless: bool = True):
    """优先用 Playwright 自带 Chromium，其次本机 Chrome。"""
    try:
        return playwright.chromium.launch(headless=headless)
    except Exception:  # noqa: BLE001
        return playwright.chromium.launch(headless=headless, channel="chrome")


def login(page) -> None:
    """登录睿贝；已登录时直接返回。"""
    user, password = erp_credentials()
    page.goto(f"{ERP_ROOT}/", wait_until="domcontentloaded", timeout=60_000)
    if not page.locator("input[type=password]").count():
        return
    user_box = (
        page.locator("#userId")
        if page.locator("#userId").count()
        else page.locator("input[type=text]").first
    )
    pwd_box = (
        page.locator("#password")
        if page.locator("#password").count()
        else page.locator("input[type=password]").first
    )
    user_box.fill(user)
    pwd_box.fill(password)
    button = page.get_by_role("button", name="立即登录")
    if button.count():
        button.click()
    else:
        page.locator("button[type=submit],input[type=submit]").first.click()
    page.wait_for_load_state("domcontentloaded", timeout=60_000)
    page.locator("input[type=password]").wait_for(state="hidden", timeout=30_000)


def capture_xhr(
    page,
    url: str,
    out_dir: Path,
    tag: str,
    *,
    wait_ms: int = 6000,
    filter_fn: Callable[[str], bool] | None = None,
) -> list[dict[str, Any]]:
    """打开页面并记录所有 XHR/fetch 请求与响应，落盘到 out_dir。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []

    def on_response(response) -> None:
        try:
            request = response.request
            if request.resource_type not in ("xhr", "fetch"):
                return
            if filter_fn and not filter_fn(response.url):
                return
            body: Any = None
            try:
                body = response.json()
            except Exception:  # noqa: BLE001
                try:
                    body = response.text()[:4000]
                except Exception:  # noqa: BLE001
                    body = None
            records.append(
                {
                    "url": response.url,
                    "method": request.method,
                    "status": response.status,
                    "post_data": request.post_data,
                    "body": body,
                }
            )
        except Exception:  # noqa: BLE001
            return

    page.on("response", on_response)
    page.goto(f"{ERP_ROOT}{url}", wait_until="domcontentloaded", timeout=60_000)
    page.wait_for_timeout(wait_ms)
    page.remove_listener("response", on_response)
    (out_dir / f"{tag}.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    return records
