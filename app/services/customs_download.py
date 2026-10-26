"""中国国际贸易单一窗口出口报关单下载。

本模块只复现影刀的“网页取 PDF”阶段，不解析、不写飞书：

1. 使用持久化 Chrome 会话打开出口报关单打印页面；
2. 按日期查询并逐页读取报关单号；
3. 全选当前页并点击“批量预览/打印”；
4. 从浏览器下载事件、PDF 响应或打印预览页生成 PDF；
5. 把原始分页 PDF 和断点清单保存在调用方指定的目录。

登录仍使用单一窗口的卡介质。会话可复用时不会再次登录；过期时需要人工
完成一次登录，卡口令可通过 ``SINGLEWINDOW_CARD_PASSWORD`` 环境变量提供。
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import time
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from app.services.feishu_client import load_env


LOGIN_URL = "https://app.singlewindow.cn/cas/login?service"
DASHBOARD_URL = "https://swapp.singlewindow.cn/deskserver/sw/deskIndex?menu_id=dec001#"
EXPORT_URL = (
    "https://swapp.singlewindow.cn/deskserver/sw/deskIndex"
    "?menu_id=txrpjstop&show_id=txrpjstop9001"
)

RESULT_TABLE = "#resultTable"
DECLARATION_RE = re.compile(r"(?<!\d)(\d{18})(?!\d)")
NEXT_SELECTORS = (
    "li.page-next:not(.disabled) a",
    "li.page-next a",
    ".pagination .page-next:not(.disabled) a",
    "a[title='下一页']",
    "a[aria-label='下一页']",
)


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def normalize_date(value: str) -> str:
    """把常见日期写法规范成 ``YYYYMMDD``。"""
    raw = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    raise ValueError(f"日期格式不正确：{value!r}，请使用 YYYY-MM-DD 或 YYYYMMDD")


def date_range(start: str, end: str, *, days: int = 30) -> tuple[str, str]:
    """计算下载日期区间。

    未指定结束日期时取今天；未指定开始日期时从结束日期向前 ``days`` 天。
    日期只用于网站查询条件，不做自然月推断。
    """
    end_value = normalize_date(end) if end else date.today().strftime("%Y%m%d")
    if start:
        start_value = normalize_date(start)
    else:
        if days < 1:
            raise ValueError("days 必须大于 0")
        start_value = (
            datetime.strptime(end_value, "%Y%m%d") - timedelta(days=days)
        ).strftime("%Y%m%d")
    if start_value > end_value:
        raise ValueError("开始日期不能晚于结束日期")
    return start_value, end_value


def default_run_id(start: str, end: str, *, flow: str = "export") -> str:
    """同一日期区间默认复用同一运行目录，便于断点续跑。"""
    return f"{flow}_{start}_{end}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def valid_pdf(path: Path) -> bool:
    """检查文件非空且确实以 PDF 头开始。"""
    try:
        if path.stat().st_size < 5:
            return False
        with path.open("rb") as handle:
            return handle.read(5) == b"%PDF-"
    except OSError:
        return False


def write_pdf_atomic(path: Path, content: bytes) -> None:
    """先写临时文件，再原子替换，避免中断后把半个 PDF 当成成功。"""
    if not content.startswith(b"%PDF"):
        raise ValueError("网页返回的内容不是 PDF")
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".part")
    temporary.write_bytes(content)
    if not valid_pdf(temporary):
        temporary.unlink(missing_ok=True)
        raise ValueError("下载得到的是空文件或损坏 PDF")
    os.replace(temporary, path)


def extract_declaration_numbers(row_texts: Iterable[str]) -> list[str]:
    """从结果表格每行文本中提取并去重 18 位报关单号，保留页面顺序。"""
    out: list[str] = []
    for text in row_texts:
        for value in DECLARATION_RE.findall(str(text or "")):
            if value not in out:
                out.append(value)
    return out


@dataclass
class PageRecord:
    page: int
    declaration_numbers: list[str]
    item_id: str = ""
    flow: str = ""
    pdf: str = ""
    bytes: int = 0
    sha256: str = ""
    status: str = "pending"
    error: str = ""
    downloaded_at: str = ""


@dataclass
class DownloadManifest:
    run_id: str
    flow: str
    start_date: str
    end_date: str
    output_dir: str
    status: str = "running"
    created_at: str = field(default_factory=now_iso)
    updated_at: str = field(default_factory=now_iso)
    pages: list[PageRecord] = field(default_factory=list)
    error: str = ""

    def page(self, number: int) -> PageRecord | None:
        return next((item for item in self.pages if item.page == number), None)

    def put_page(self, record: PageRecord) -> None:
        for index, item in enumerate(self.pages):
            same_item = record.item_id and item.item_id == record.item_id
            if same_item or (not record.item_id and item.page == record.page):
                self.pages[index] = record
                break
        else:
            self.pages.append(record)
        self.pages.sort(key=lambda item: item.page)
        self.updated_at = now_iso()


def load_manifest(path: Path) -> DownloadManifest | None:
    if not path.exists():
        return None
    payload = json.loads(path.read_text(encoding="utf-8"))
    payload["pages"] = [PageRecord(**item) for item in payload.get("pages") or []]
    return DownloadManifest(**payload)


def save_manifest(path: Path, manifest: DownloadManifest) -> None:
    manifest.updated_at = now_iso()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".part")
    temporary.write_text(
        json.dumps(asdict(manifest), ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def declaration_index(manifest: DownloadManifest, run_dir: Path) -> dict[str, dict[str, Any]]:
    """生成“报关单号 -> 所在分页 PDF”的索引。"""
    out: dict[str, dict[str, Any]] = {}
    for page in manifest.pages:
        if page.status != "downloaded" or not page.pdf:
            continue
        for declaration in page.declaration_numbers:
            entry = out.setdefault(
                declaration, {"declaration": declaration, "references": []}
            )
            entry["references"].append({
                "flow": page.flow or manifest.flow,
                "page": page.page,
                "pdf": page.pdf,
                "absolutePdf": str((run_dir / page.pdf).resolve()),
                "sha256": page.sha256,
            })
    return out


def save_declaration_index(manifest: DownloadManifest, run_dir: Path) -> Path:
    path = run_dir / "declaration_index.json"
    temporary = path.with_name(path.name + ".part")
    temporary.write_text(
        json.dumps(declaration_index(manifest, run_dir), ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    os.replace(temporary, path)
    return path


def _frame_has(frame: Any, selector: str) -> bool:
    try:
        return frame.locator(selector).count() > 0
    except Exception:  # noqa: BLE001
        return False


def find_frame(page: Any, selector: str, *, timeout_ms: int = 30_000) -> Any:
    """在主页面的所有 frame 中寻找包含目标控件的 frame。"""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for frame in page.frames:
            if _frame_has(frame, selector):
                return frame
        page.wait_for_timeout(250)
    raise RuntimeError(f"网页中找不到控件 {selector}；请检查登录状态或页面是否改版")


def _visible_password_input(page: Any) -> Any:
    for selector in ("input#password", "input[type='password']"):
        locator = page.locator(selector)
        try:
            if locator.count() and locator.first.is_visible():
                return locator.first
        except Exception:  # noqa: BLE001
            continue
    return None


def ensure_logged_in(page: Any, *, manual_timeout: int = 300) -> None:
    """打开单一窗口；能安全自动填写的登录信息才自动填写。

    卡介质必须插在运行程序的本机。没有卡口令时会等待操作员在真实浏览器里
    手工完成登录，避免在代码中保存或打印口令。
    """
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90_000)
    page.wait_for_timeout(1500)
    if "cas/login" not in page.url:
        return

    env = load_env()
    password = os.environ.get("SINGLEWINDOW_CARD_PASSWORD") or env.get(
        "SINGLEWINDOW_CARD_PASSWORD", ""
    )
    card_tab = page.locator("#cardTabBtn")
    if card_tab.count():
        try:
            card_tab.first.click()
            page.wait_for_timeout(300)
        except Exception:  # noqa: BLE001
            pass
    password_box = _visible_password_input(page)
    if password and password_box is not None:
        password_box.fill(password)
        agreement = page.locator("#agreementCardId")
        try:
            if agreement.count():
                inner = agreement.locator("input")
                if not inner.count() or not inner.first.is_checked():
                    agreement.first.click()
        except Exception:  # noqa: BLE001
            pass
        login_button = page.locator("#loginbutton")
        if login_button.count():
            login_button.first.click()

    print(
        "请在打开的 Chrome 窗口中完成单一窗口卡介质登录"
        f"（最多等待 {manual_timeout} 秒）...",
        flush=True,
    )
    deadline = time.monotonic() + manual_timeout
    while time.monotonic() < deadline:
        if "cas/login" not in page.url:
            return
        page.wait_for_timeout(1000)
    raise TimeoutError("单一窗口登录未完成")


def open_export_page(page: Any, *, url: str = EXPORT_URL) -> Any:
    page.goto(url, wait_until="domcontentloaded", timeout=90_000)
    frame = find_frame(page, 'input[name="startDate"]', timeout_ms=60_000)
    return frame


def _click_first_visible(frame: Any, selectors: Iterable[str], *, timeout_ms: int = 30_000) -> Any:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for selector in selectors:
            locator = frame.locator(selector)
            try:
                for index in range(locator.count()):
                    item = locator.nth(index)
                    if item.is_visible():
                        return item
            except Exception:  # noqa: BLE001
                continue
        frame.page.wait_for_timeout(250)
    raise RuntimeError(f"找不到可点击控件：{' / '.join(selectors)}")


def _set_autocomplete(frame: Any, selector: str, value: str) -> None:
    """填写报关数据查询页的自动完成输入框并选择精确选项。"""
    box = frame.locator(selector)
    if not box.count():
        raise RuntimeError(f"报关数据查询页缺少 {selector}")
    box.first.click()
    box.first.fill(value)
    frame.page.wait_for_timeout(600)
    candidates = frame.locator(
        "[role='option'], .autocompleter-item, .dropdown-menu li, .typeahead li"
    )
    for index in range(candidates.count()):
        item = candidates.nth(index)
        try:
            if item.is_visible() and value in item.inner_text():
                item.click()
                return
        except Exception:  # noqa: BLE001
            continue
    exact = frame.get_by_text(value, exact=True)
    for index in range(exact.count()):
        item = exact.nth(index)
        try:
            if item.is_visible():
                item.click()
                return
        except Exception:  # noqa: BLE001
            continue
    raise RuntimeError(f"自动完成框 {selector} 中找不到选项 {value}")


def open_pre_record_page(page: Any) -> Any:
    """复现影刀的“综合查询 -> 报关数据查询”入口。"""
    page.goto(DASHBOARD_URL, wait_until="domcontentloaded", timeout=90_000)
    menu = None
    try:
        menu = _click_first_visible(
            page,
            (
                "a:has-text('综合查询')",
                "a[title='综合查询']",
            ),
            timeout_ms=10_000,
        )
    except RuntimeError:
        pass
    if menu is not None:
        menu.click()
        page.wait_for_timeout(500)
    query_link = _click_first_visible(
        page,
        (
            "a[menuid='dec001th06leaf011']",
            "a[title='报关数据查询']",
            "a:has-text('报关数据查询')",
        ),
        timeout_ms=30_000,
    )
    query_link.click()
    frame = find_frame(page, "#etpsCategoryName", timeout_ms=60_000)
    advanced = frame.locator("#cusAdvancedQuery")
    if advanced.count():
        try:
            advanced.first.click()
            page.wait_for_timeout(500)
        except Exception:  # noqa: BLE001
            pass
    return frame


def configure_pre_record_query(
    frame: Any,
    *,
    enterprise_category: str = "C-报关收发货人",
    lookback_months: int = 6,
) -> None:
    """复现影刀预录查询：收发货人 + 出口 + 放行 + 最近六个月。"""
    if lookback_months != 6:
        raise ValueError("单一窗口页面仅复刻“六个月”快捷条件；其他区间请先不要指定")
    _set_autocomplete(frame, "#etpsCategoryName", enterprise_category)
    _set_autocomplete(frame, "#cusIEFlagName", "E-出口")
    _set_autocomplete(frame, "#cusDecStatusName", "9-放行")
    date_radio = frame.locator("#operateDate8")
    if not date_radio.count():
        raise RuntimeError("报关数据查询页缺少“六个月”条件 #operateDate8")
    try:
        date_radio.first.check(force=True)
    except Exception:  # noqa: BLE001
        date_radio.first.click(force=True)
    query = frame.locator("#decQuery")
    if not query.count():
        raise RuntimeError("报关数据查询页缺少 #decQuery")
    query.first.click()
    find_frame(frame.page, "#declareTable", timeout_ms=60_000)
    wait_for_pre_record_result_state(frame)


def current_pre_record_numbers(frame: Any) -> list[str]:
    try:
        return extract_declaration_numbers(
            frame.locator("#declareTable tbody tr").all_inner_texts()
        )
    except Exception:  # noqa: BLE001
        return []


def wait_for_pre_record_result_state(frame: Any, *, timeout_ms: int = 30_000) -> list[str]:
    deadline = time.monotonic() + timeout_ms / 1000
    empty_markers = ("暂无数据", "没有数据", "无数据", "未查询到", "no matching records", "no data")
    last_text = ""
    while time.monotonic() < deadline:
        declarations = current_pre_record_numbers(frame)
        if declarations:
            return declarations
        try:
            last_text = frame.locator("body").inner_text(timeout=3_000)
        except Exception:  # noqa: BLE001
            last_text = ""
        if any(marker.lower() in last_text.lower() for marker in empty_markers):
            return []
        frame.page.wait_for_timeout(500)
    raise TimeoutError(
        "预录报关数据未加载完成"
        + (f"；当前页面提示：{last_text[:160]}" if last_text else "")
    )


def select_pre_record_row(frame: Any, declaration_no: str) -> None:
    rows = frame.locator("#declareTable tbody tr")
    for index in range(rows.count()):
        row = rows.nth(index)
        try:
            if declaration_no not in row.inner_text():
                continue
            checkbox = row.locator("input[type='checkbox']")
            if not checkbox.count():
                raise RuntimeError(f"报关单 {declaration_no} 所在行没有选择框")
            try:
                checkbox.first.check(force=True)
            except Exception:  # noqa: BLE001
                checkbox.first.click(force=True)
            return
        except RuntimeError:
            raise
        except Exception:  # noqa: BLE001
            continue
    raise RuntimeError(f"预录结果页找不到报关单 {declaration_no}")


def clear_pre_record_row(frame: Any, declaration_no: str) -> None:
    rows = frame.locator("#declareTable tbody tr")
    for index in range(rows.count()):
        row = rows.nth(index)
        try:
            if declaration_no not in row.inner_text():
                continue
            checkbox = row.locator("input[type='checkbox']").first
            if checkbox.is_checked():
                checkbox.uncheck(force=True)
            return
        except Exception:  # noqa: BLE001
            continue


def configure_export_query(frame: Any, start: str, end: str) -> None:
    """按影刀原流程选择日期区间并查询。"""
    radio = frame.locator('input[name="dateRadio"]')
    if radio.count():
        try:
            radio.first.check(force=True)
        except Exception:  # noqa: BLE001
            radio.first.click(force=True)

    start_box = frame.locator('input[name="startDate"]')
    end_box = frame.locator('input[name="endDate"]')
    if not start_box.count() or not end_box.count():
        raise RuntimeError("报关单查询页缺少 startDate / endDate 输入框")
    start_box.first.fill(start)
    start_box.first.press("Enter")
    end_box.first.fill(end)
    end_box.first.press("Enter")
    frame.page.wait_for_timeout(500)

    search = frame.locator("#searchBtn")
    if not search.count():
        raise RuntimeError("报关单查询页缺少 #searchBtn")
    search.first.click()
    find_frame(frame.page, RESULT_TABLE, timeout_ms=60_000)
    wait_for_result_state(frame)


def current_declaration_numbers(frame: Any) -> list[str]:
    rows = frame.locator(f"{RESULT_TABLE} tbody tr")
    try:
        return extract_declaration_numbers(rows.all_inner_texts())
    except Exception:  # noqa: BLE001
        return []


def wait_for_result_state(frame: Any, *, timeout_ms: int = 30_000) -> list[str]:
    """等待查询结果稳定；无数据时返回空列表，加载超时则明确报错。"""
    deadline = time.monotonic() + timeout_ms / 1000
    empty_markers = (
        "暂无数据",
        "没有数据",
        "无数据",
        "未查询到",
        "no matching records",
        "no data",
    )
    last_text = ""
    while time.monotonic() < deadline:
        declarations = current_declaration_numbers(frame)
        if declarations:
            return declarations
        try:
            last_text = frame.locator("body").inner_text(timeout=3_000)
        except Exception:  # noqa: BLE001
            last_text = ""
        lowered = last_text.lower()
        if any(marker.lower() in lowered for marker in empty_markers):
            return []
        frame.page.wait_for_timeout(500)
    raise TimeoutError(
        "报关单查询结果未加载完成"
        + (f"；当前页面提示：{last_text[:160]}" if last_text else "")
    )


def select_current_page(frame: Any) -> None:
    checkbox = frame.locator('input[name="btSelectAll"]')
    if not checkbox.count():
        raise RuntimeError("报关单结果页缺少批量选择框")
    target = checkbox.first
    try:
        if not target.is_checked():
            target.check(force=True)
    except Exception:  # noqa: BLE001
        target.click(force=True)


def _pdf_from_download(download: Any, target: Path) -> None:
    temporary = target.with_name(target.name + ".download")
    download.save_as(str(temporary))
    content = temporary.read_bytes()
    temporary.unlink(missing_ok=True)
    write_pdf_atomic(target, content)


def _pdf_from_response(response: Any, target: Path) -> bool:
    try:
        content = response.body()
    except Exception:  # noqa: BLE001
        return False
    if not content.startswith(b"%PDF"):
        return False
    write_pdf_atomic(target, content)
    return True


def _pdf_from_page(page: Any, target: Path) -> bool:
    """从打印预览/PDF 查看页得到 PDF。

    优先读取当前 URL 的原始 PDF 响应；如果它是 HTML 打印页或 Blob，再使用
    Chromium 的 print-to-PDF 渲染当前页面。后者是为了替代影刀的 ``Ctrl+S``
    原生保存对话框。
    """
    url = str(page.url or "")
    if url.startswith(("http://", "https://")):
        try:
            response = page.context.request.get(url, timeout=60_000)
            if response.ok:
                content = response.body()
                if content.startswith(b"%PDF"):
                    write_pdf_atomic(target, content)
                    return True
        except Exception:  # noqa: BLE001
            pass
    if url.startswith("blob:"):
        try:
            encoded = page.evaluate(
                """async (url) => {
                    const data = await (await fetch(url)).arrayBuffer();
                    const bytes = new Uint8Array(data);
                    let binary = '';
                    for (let i = 0; i < bytes.length; i += 32768) {
                        binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
                    }
                    return btoa(binary);
                }""",
                url,
            )
            content = base64.b64decode(encoded)
            if content.startswith(b"%PDF"):
                write_pdf_atomic(target, content)
                return True
        except Exception:  # noqa: BLE001
            pass
    try:
        content = page.pdf(format="A4", print_background=True, prefer_css_page_size=True)
    except Exception:  # noqa: BLE001
        return False
    if not content.startswith(b"%PDF"):
        return False
    write_pdf_atomic(target, content)
    return True


def capture_pdf_after_action(
    page: Any,
    target: Path,
    action: Any,
    *,
    timeout_ms: int = 60_000,
    same_page_allowed: bool = False,
) -> Path:
    """执行打印动作，并从下载事件、PDF 响应或预览页取得 PDF。"""
    downloads: list[Any] = []
    pages: list[Any] = []
    responses: list[Any] = []
    context = page.context

    def on_download(item: Any) -> None:
        downloads.append(item)

    def on_page(item: Any) -> None:
        pages.append(item)

    def on_response(item: Any) -> None:
        try:
            content_type = str(item.headers.get("content-type") or "").lower()
        except Exception:  # noqa: BLE001
            content_type = ""
        if "pdf" in content_type or ".pdf" in str(item.url or "").lower():
            responses.append(item)

    page.on("download", on_download)
    context.on("page", on_page)
    context.on("response", on_response)
    try:
        action()
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            if downloads or responses:
                # 响应正文和打印预览通常比 response 事件稍晚就绪。
                page.wait_for_timeout(1500)
                break
            page.wait_for_timeout(500)

        if downloads:
            _pdf_from_download(downloads[-1], target)
        else:
            saved = any(_pdf_from_response(item, target) for item in reversed(responses))
            if not saved:
                candidates = [item for item in pages if not item.is_closed()]
                if not candidates and same_page_allowed:
                    candidates = [page]
                for candidate in reversed(candidates):
                    try:
                        candidate.wait_for_load_state("domcontentloaded", timeout=20_000)
                    except Exception:  # noqa: BLE001
                        pass
                    if _pdf_from_page(candidate, target):
                        saved = True
                        break
            if not saved:
                raise RuntimeError("未能从批量打印窗口取得 PDF")
        if not valid_pdf(target):
            raise RuntimeError(f"保存后的文件不是有效 PDF：{target}")
        return target
    finally:
        page.remove_listener("download", on_download)
        context.remove_listener("page", on_page)
        context.remove_listener("response", on_response)
        for extra in pages:
            if not extra.is_closed() and extra is not page:
                try:
                    extra.close()
                except Exception:  # noqa: BLE001
                    pass


def capture_current_page_pdf(
    page: Any,
    frame: Any,
    target: Path,
    *,
    timeout_ms: int = 60_000,
) -> Path:
    """点击批量打印并保存当前结果页的合并 PDF。"""
    button = frame.locator("#batchPrintBtn")
    if not button.count():
        raise RuntimeError("报关单结果页缺少 #batchPrintBtn")
    return capture_pdf_after_action(
        page,
        target,
        lambda: button.first.click(timeout=30_000),
        timeout_ms=timeout_ms,
        same_page_allowed=False,
    )


def capture_single_pre_record_pdf(
    page: Any,
    frame: Any,
    target: Path,
    *,
    declaration_no: str,
    timeout_ms: int = 60_000,
) -> Path:
    """复现预录单“打印报关单 -> 打印预览”的单张下载。"""
    print_button = frame.locator("#decPdfPrint")
    if not print_button.count():
        raise RuntimeError("报关数据查询页缺少 #decPdfPrint")

    def print_preview() -> None:
        print_button.first.click(timeout=30_000)
        preview = _click_first_visible(
            frame,
            (
                "a:has-text('打印预览')",
                "button:has-text('打印预览')",
                "a[href='javascript:;']:has-text('打印预览')",
            ),
            timeout_ms=60_000,
        )
        preview.click(timeout=30_000)

    try:
        return capture_pdf_after_action(
            page,
            target,
            print_preview,
            timeout_ms=timeout_ms,
            same_page_allowed=False,
        )
    except Exception as exc:
        raise RuntimeError(f"预录单 {declaration_no} 下载失败：{exc}") from exc


def _next_control(frame: Any, selectors: Iterable[str]) -> Any | None:
    for selector in selectors:
        locator = frame.locator(selector)
        try:
            if locator.count():
                target = locator.first
                if target.is_visible():
                    parent_class = target.locator("xpath=..").get_attribute("class") or ""
                    if "disabled" in parent_class.split():
                        continue
                    return target
        except Exception:  # noqa: BLE001
            continue
    return None


def advance_page(frame: Any, previous: list[str], *, timeout_ms: int = 30_000) -> bool:
    """点击下一页；页面未变化或不存在下一页时返回 False。"""
    next_link = _next_control(frame, NEXT_SELECTORS)
    if next_link is None:
        return False
    try:
        if next_link.get_attribute("aria-disabled") == "true":
            return False
    except Exception:  # noqa: BLE001
        pass
    next_link.click()
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        frame.page.wait_for_timeout(500)
        current = current_declaration_numbers(frame)
        if current and current != previous:
            return True
    return False


def download_export_pdfs(
    page: Any,
    *,
    start: str,
    end: str,
    run_dir: Path,
    manifest: DownloadManifest,
    max_pages: int = 100,
    page_wait_ms: int = 1200,
    preview_timeout_ms: int = 60_000,
    force: bool = False,
    next_selectors: Iterable[str] = NEXT_SELECTORS,
) -> DownloadManifest:
    """下载日期区间内的全部出口正式单，每页结果保存为一个 PDF。"""
    run_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = run_dir / "manifest.json"
    frame = open_export_page(page)
    configure_export_query(frame, start, end)

    for page_number in range(1, max_pages + 1):
        declarations = wait_for_result_state(frame)
        if not declarations:
            if page_number == 1:
                print("查询区间内没有出口报关单", flush=True)
            break

        target = run_dir / "raw" / f"{start}_{end}_page_{page_number:04d}.pdf"
        prior = manifest.page(page_number)
        can_resume = (
            not force
            and prior is not None
            and prior.declaration_numbers == declarations
            and valid_pdf(target)
        )
        if can_resume:
            prior.pdf = str(target.relative_to(run_dir))
            prior.status = "downloaded"
            prior.bytes = target.stat().st_size
            prior.sha256 = sha256_file(target)
            manifest.put_page(prior)
            print(
                f"[{page_number}/{max_pages}] 已完成，跳过 {len(declarations)} 张报关单",
                flush=True,
            )
        else:
            select_current_page(frame)
            print(
                f"[{page_number}/{max_pages}] 下载本页 {len(declarations)} 张报关单",
                flush=True,
            )
            try:
                capture_current_page_pdf(
                    page, frame, target, timeout_ms=preview_timeout_ms
                )
            except Exception as exc:
                manifest.put_page(
                    PageRecord(
                        page=page_number,
                        declaration_numbers=declarations,
                        flow="export",
                        status="failed",
                        error=str(exc)[:1000],
                    )
                )
                save_manifest(manifest_path, manifest)
                raise
            manifest.put_page(
                PageRecord(
                    page=page_number,
                    declaration_numbers=declarations,
                    flow="export",
                    pdf=str(target.relative_to(run_dir)),
                    bytes=target.stat().st_size,
                    sha256=sha256_file(target),
                    status="downloaded",
                    downloaded_at=now_iso(),
                )
            )
        save_manifest(manifest_path, manifest)
        save_declaration_index(manifest, run_dir)
        page.wait_for_timeout(page_wait_ms)

        if not advance_page(frame, declarations):
            break
    else:
        raise RuntimeError(f"达到最大页数 {max_pages}，下载未结束")

    save_declaration_index(manifest, run_dir)
    return manifest


def _manifest_item(manifest: DownloadManifest, item_id: str) -> PageRecord | None:
    return next((item for item in manifest.pages if item.item_id == item_id), None)


def download_pre_record_pdfs(
    page: Any,
    *,
    run_dir: Path,
    manifest: DownloadManifest,
    lookback_months: int = 6,
    enterprise_category: str = "C-报关收发货人",
    max_pages: int = 20,
    page_wait_ms: int = 800,
    preview_timeout_ms: int = 60_000,
    force: bool = False,
    next_selectors: Iterable[str] = NEXT_SELECTORS,
) -> DownloadManifest:
    """复现影刀预录单流程：逐张勾选、打印预览并保存单张 PDF。"""
    run_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = run_dir / "manifest.json"
    frame = open_pre_record_page(page)
    configure_pre_record_query(
        frame,
        enterprise_category=enterprise_category,
        lookback_months=lookback_months,
    )

    sequence = max((item.page for item in manifest.pages), default=0)
    for screen_page in range(1, max_pages + 1):
        declarations = wait_for_pre_record_result_state(frame)
        if not declarations:
            if screen_page == 1:
                print("预录查询中没有放行报关单", flush=True)
            break
        for ordinal, declaration_no in enumerate(declarations, 1):
            target = run_dir / "pre_record" / f"{declaration_no}.pdf"
            prior = _manifest_item(manifest, f"pre:{declaration_no}")
            can_resume = not force and prior is not None and valid_pdf(target)
            if can_resume:
                prior.pdf = str(target.relative_to(run_dir))
                prior.status = "downloaded"
                prior.bytes = target.stat().st_size
                prior.sha256 = sha256_file(target)
                sequence = max(sequence, prior.page)
                manifest.put_page(prior)
                print(
                    f"预录 {declaration_no} 已完成，跳过",
                    flush=True,
                )
                continue

            sequence += 1
            print(
                f"[预录 {screen_page}/{max_pages}] {ordinal}/{len(declarations)} "
                f"{declaration_no}",
                flush=True,
            )
            try:
                select_pre_record_row(frame, declaration_no)
                capture_single_pre_record_pdf(
                    page,
                    frame,
                    target,
                    declaration_no=declaration_no,
                    timeout_ms=preview_timeout_ms,
                )
            except Exception as exc:
                manifest.put_page(
                    PageRecord(
                        page=sequence,
                        declaration_numbers=[declaration_no],
                        item_id=f"pre:{declaration_no}",
                        flow="pre_record",
                        status="failed",
                        error=str(exc)[:1000],
                    )
                )
                save_manifest(manifest_path, manifest)
                raise
            finally:
                clear_pre_record_row(frame, declaration_no)
            manifest.put_page(
                PageRecord(
                    page=sequence,
                    declaration_numbers=[declaration_no],
                    item_id=f"pre:{declaration_no}",
                    flow="pre_record",
                    pdf=str(target.relative_to(run_dir)),
                    bytes=target.stat().st_size,
                    sha256=sha256_file(target),
                    status="downloaded",
                    downloaded_at=now_iso(),
                )
            )
            save_manifest(manifest_path, manifest)
            save_declaration_index(manifest, run_dir)
            page.wait_for_timeout(page_wait_ms)
        if not advance_page(frame, declarations, timeout_ms=30_000):
            break
    else:
        raise RuntimeError(f"预录查询达到最大页数 {max_pages}，下载未结束")

    save_declaration_index(manifest, run_dir)
    return manifest


def launch_persistent_browser(playwright: Any, *, profile_dir: Path, channel: str = "chrome") -> Any:
    """启动有界面的持久化浏览器，保留登录 Cookie 和本机卡控件环境。"""
    profile_dir.mkdir(parents=True, exist_ok=True)
    options: dict[str, Any] = {
        "headless": False,
        "accept_downloads": True,
        "viewport": {"width": 1440, "height": 1000},
        "args": ["--start-maximized"],
    }
    if channel:
        options["channel"] = channel
    return playwright.chromium.launch_persistent_context(str(profile_dir), **options)
