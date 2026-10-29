"""睿贝 ERP Web 端访问层：Playwright 登录 + 常用页面的 XHR 捕获。

睿贝同一账号同时只能一人登录，因此抓取应尽量一次性完成；
这里把「登录」和「浏览器启动」独立出来，便于脚本复用同一会话。
"""
from __future__ import annotations

import json
import os
import sys
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


def chromium_executables() -> list[Path]:
    """返回可用的 Chromium/Chrome 可执行文件，优先显式配置。"""
    candidates: list[Path] = []
    configured = os.environ.get("PLAYWRIGHT_CHROMIUM_EXECUTABLE", "").strip()
    if configured:
        candidates.append(Path(configured))
    if sys.platform.startswith("linux"):
        cache = Path.home() / ".cache" / "ms-playwright"
        candidates.extend(
            sorted(cache.glob("chromium-*/chrome-linux64/chrome"), reverse=True)
        )
        candidates.extend(
            Path(path)
            for path in (
                "/usr/bin/chromium",
                "/usr/bin/chromium-browser",
                "/usr/bin/google-chrome",
                "/usr/bin/google-chrome-stable",
            )
        )
    elif sys.platform.startswith("win"):
        for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")):
            if base:
                candidates.append(
                    Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe"
                )
    return [path for path in candidates if path.is_file()]


def launch_browser(playwright, headless: bool = True):
    """优先用 Playwright 默认浏览器，再回退到本机或缓存的 Chromium。"""
    try:
        return playwright.chromium.launch(headless=headless)
    except Exception:  # noqa: BLE001
        pass
    errors: list[str] = []
    for executable in chromium_executables():
        try:
            return playwright.chromium.launch(
                headless=headless,
                executable_path=str(executable),
                args=["--no-sandbox"],
            )
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{executable}: {exc}")
    try:
        return playwright.chromium.launch(headless=headless, channel="chrome")
    except Exception as exc:  # noqa: BLE001
        errors.append(f"channel=chrome: {exc}")
    raise RuntimeError("无法启动 Chromium：" + "；".join(errors))


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
    # 登录页同时带着「修改密码」弹窗的两个隐藏密码框（#diy_password / #diy_passwordConfirm），
    # `input[type=password]` 会命中 3 个元素，Playwright 严格模式直接抛错；
    # 这里只等登录框 #password 消失（取 .first 同样是为了避开严格模式）。
    page.locator("#password").first.wait_for(state="hidden", timeout=30_000)


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
