"""用 Playwright 登录后，在页面上下文里重放睿贝内部列表接口。

睿贝是服务端渲染 + jQuery 分页：
- 列表接口形如 `POST /saleOrder_list`，body 为 query string 形式；
- 返回 `{total, root: [[column...]]}`，每个 column 含 `beanName` 与 `columnValues`。
本模块负责登录、调用、翻页以及把这种「列式」响应转成普通记录列表。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any, Iterable

from app.services.erp_web import ERP_ROOT, launch_browser, login


@dataclass
class ListResult:
    """一次列表请求的结果。"""

    total: int
    rows: list[dict[str, Any]]
    raw: dict[str, Any]


def columns_to_rows(root: Any) -> list[dict[str, Any]]:
    """把睿贝列式表格 `root` 转成 [{beanName: value}]。

    root 可能是平铺的 column 数组，也可能按行分组，两种都兼容。
    """
    if not root:
        return []
    # 情形一：扁平 column 列表
    if isinstance(root, list) and root and isinstance(root[0], dict) and "columnValues" in root[0]:
        return _flatten_columns(root)
    # 情形二：按行分组
    rows: list[dict[str, Any]] = []
    for group in root:
        if isinstance(group, list):
            rows.extend(_flatten_columns(group))
    return rows


def _flatten_columns(columns: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """把一组 column 展开成多行记录。

    睿贝的 `beanName` 是**与 `columnValues` 一一对应的字段名列表**，例如
    `beanName="orderId,comId,shipDeadline,status"` 且
    `columnValues=["6260","5562","2026-10-03","1450"]` 表示第一条记录的
    orderId=6260、comId=5562、shipDeadline=2026-10-03、status=1450；
    第二条记录则从下标 4 开始。即 `columnValues` 按「记录 × 字段」行优先排列。
    """
    def field_names(column: dict[str, Any]) -> list[str]:
        raw = str(column.get("beanName") or column.get("columnName") or "")
        return [n.strip() for n in raw.split(",") if n.strip()]

    # 单字段列取值为「一条记录一个值」，多字段列取值为「一条记录 * 字段数」
    # 因此行数 = 取值个数 ÷ 字段个数（单字段列直接等于取值个数）。
    length = 0
    for column in columns:
        values = column.get("columnValues") or []
        names = field_names(column)
        if not names:
            continue
        length = max(length, len(values) // len(names))

    rows: list[dict[str, Any]] = []
    for index in range(length):
        row: dict[str, Any] = {}
        for column in columns:
            values = column.get("columnValues") or []
            names = field_names(column)
            if not names:
                continue
            span = len(names)
            for position, name in enumerate(names):
                slot = index * span + position if span > 1 else index
                if slot >= len(values):
                    continue
                if name not in row or not row[name]:
                    row[name] = values[slot]
        rows.append(row)
    return rows


def parse_lenient_json(text: str) -> Any:
    """解析睿贝返回的非标准 JSON。

    部分接口（如 `shipment_select`）返回 `{total:4,root:[[{"id":"",...}]]}`：
    key 没有引号、布尔/空值写法也较随意，标准 json 解析会失败。
    """
    if not text:
        return {}
    try:
        return json.loads(text)
    except ValueError:
        pass
    fixed = text.strip()
    # 给未加引号的 key 补引号：{key: -> {"key":
    fixed = re.sub(r"([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:", r'\1"\2":', fixed)
    fixed = re.sub(r":\s*undefined\b", ":null", fixed)
    fixed = re.sub(r":\s*NaN\b", ":null", fixed)
    try:
        return json.loads(fixed)
    except ValueError:
        return {"_raw": text}


class ErpSession:
    """维护一个已登录的浏览器页面，并提供接口重放能力。"""

    def __init__(self, headless: bool = True) -> None:
        self.headless = headless
        self._playwright = None
        self._browser = None
        self.page = None

    def __enter__(self) -> "ErpSession":
        from playwright.sync_api import sync_playwright

        self._playwright = sync_playwright().start()
        self._browser = launch_browser(self._playwright, headless=self.headless)
        self.page = self._browser.new_page()
        login(self.page)
        return self

    def __exit__(self, *exc: object) -> None:
        try:
            if self._browser:
                self._browser.close()
        finally:
            if self._playwright:
                self._playwright.stop()

    def fetch_json(self, path: str, query: dict[str, Any], *, method: str = "POST") -> dict[str, Any]:
        """在页面上下文里调用接口，带上浏览器的 Cookie 与请求头。"""
        assert self.page is not None
        body = "&".join(f"{key}={value}" for key, value in query.items())
        script = """
        async ([path, body, method]) => {
          const options = {method: method, headers: {"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8", "X-Requested-With": "XMLHttpRequest"}};
          if (method !== "GET") { options.body = body; }
          const url = method === "GET" && body ? path + "?" + body : path;
          const response = await fetch(url, options);
          const text = await response.text();
          return {status: response.status, text: text};
        }
        """
        result = self.page.evaluate(script, [path, body, method])
        text = result.get("text") or ""
        payload = parse_lenient_json(text)
        return payload if isinstance(payload, dict) else {"root": payload}

    def iter_list(
        self,
        path: str,
        query: dict[str, Any],
        *,
        page_param: str = "p",
        max_pages: int = 1000,
    ) -> Iterable[ListResult]:
        """逐页拉取列表接口。"""
        page = 1
        while page <= max_pages:
            params = dict(query)
            params[page_param] = page
            payload = self.fetch_json(path, params)
            rows = columns_to_rows(payload.get("root"))
            total = int(payload.get("total") or 0)
            yield ListResult(total=total, rows=rows, raw=payload)
            if not rows:
                break
            if total and page * len(rows) >= total:
                break
            page += 1
