"""睿贝 ERP MCP 客户端：通过 JSON-RPC 直接调用 MCP 工具。

ERP 暴露的是 MCP(HTTP) 服务，这里做最简封装：initialize 拿 session，
再逐个调用 tools/call。仅依赖标准库。
"""
from __future__ import annotations

import json
import time
import urllib.error
import urllib.request
from typing import Any

from app.services.feishu_client import get_config

DEFAULT_URL = "https://erp.mtholdinggroup.com/AgentServerSimple/mcp"
PROTOCOL_VERSION = "2025-06-18"


class ErpError(RuntimeError):
    """ERP MCP 返回错误时抛出。"""


class ErpClient:
    """睿贝 ERP MCP 客户端。"""

    def __init__(self, url: str | None = None, token: str | None = None, timeout: int = 60) -> None:
        self.url = url or get_config("ERP_MCP_URL", DEFAULT_URL)
        self.token = token or get_config("RUIBEI_ERP_API_KEY", get_config("ERP_API_KEY"))
        self.timeout = timeout
        self._session_id: str | None = None
        self._next_id = 1
        self._last_call_at = 0.0
        self.min_interval = 1.05

    def _post(self, payload: dict[str, Any], *, notify: bool = False) -> tuple[dict[str, Any] | None, dict[str, str]]:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            "Authorization": f"Bearer {self.token}",
        }
        if self._session_id:
            headers["Mcp-Session-Id"] = self._session_id
        req = urllib.request.Request(self.url, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
                resp_headers = {k.lower(): v for k, v in resp.headers.items()}
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise ErpError(f"HTTP {exc.code}: {detail[:400]}") from exc
        if notify or not raw.strip():
            return None, resp_headers
        return self._parse(raw), resp_headers

    @staticmethod
    def _parse(raw: str) -> dict[str, Any]:
        """兼容 application/json 与 text/event-stream 两种返回。"""
        text = raw.strip()
        if text.startswith("{"):
            return json.loads(text)
        for line in text.splitlines():
            line = line.strip()
            if line.startswith("data:"):
                candidate = line[5:].strip()
                if candidate.startswith("{"):
                    return json.loads(candidate)
        raise ErpError(f"无法解析响应: {text[:200]}")

    def _rpc(self, method: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        payload = {"jsonrpc": "2.0", "id": self._next_id, "method": method}
        self._next_id += 1
        if params is not None:
            payload["params"] = params
        data, headers = self._post(payload)
        if "mcp-session-id" in headers and not self._session_id:
            self._session_id = headers["mcp-session-id"]
        if data is None:
            raise ErpError(f"{method} 无响应")
        if "error" in data:
            raise ErpError(f"{method} 失败: {json.dumps(data['error'], ensure_ascii=False)}")
        return data.get("result", {})

    def initialize(self) -> dict[str, Any]:
        result = self._rpc(
            "initialize",
            {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "cost-matching", "version": "0.1.0"},
            },
        )
        self._post({"jsonrpc": "2.0", "method": "notifications/initialized"}, notify=True)
        return result

    def list_tools(self) -> list[dict[str, Any]]:
        if not self._session_id:
            self.initialize()
        return self._rpc("tools/list").get("tools", [])

    def call_tool(self, name: str, arguments: dict[str, Any] | None = None) -> Any:
        if not self._session_id:
            self.initialize()
        # 服务端限制 1 请求/秒，这里做客户端节流并允许一次重试
        for attempt in range(2):
            self._throttle()
            try:
                result = self._rpc("tools/call", {"name": name, "arguments": arguments or {}})
            except ErpError as exc:
                if "Rate limit" in str(exc) and attempt == 0:
                    time.sleep(1.2)
                    continue
                raise
            return self._unwrap(result)
        raise ErpError(f"{name} 调用失败")

    def _throttle(self) -> None:
        wait = self.min_interval - (time.time() - self._last_call_at)
        if wait > 0:
            time.sleep(wait)
        self._last_call_at = time.time()

    @staticmethod
    def _unwrap(result: dict[str, Any]) -> Any:
        """把 MCP tool 结果解包成 Python 对象（优先解析 content 中的 JSON 文本）。"""
        if isinstance(result.get("structuredContent"), dict):
            return result["structuredContent"]
        parts: list[Any] = []
        for item in result.get("content", []):
            text = item.get("text")
            if text is None:
                parts.append(item)
                continue
            try:
                parts.append(json.loads(text))
            except (ValueError, TypeError):
                parts.append(text)
        if len(parts) == 1:
            return parts[0]
        return parts
