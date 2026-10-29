"""睿贝官方 MCP 客户端（AgentServerSimple），用于按采购单号取附件等。

token 从 .env 的 ERP_API_KEY 或参数传入；这里只做最小实现。
"""
from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

URL = "https://erp.mtholdinggroup.com/AgentServerSimple/mcp"


class SessionLost(RuntimeError):
    """MCP 会话失效：返回登录页/HTML/空数据而不是 JSON-RPC 结果。

    睿贝一个账号只允许一处登录，别处登录会把这里的会话顶掉，接口随即返回登录页。
    这种响应要重新 `initialize()` 再重试，而不是当成业务错误直接失败。
    """


class McpClient:
    def __init__(self, token: str, url: str = URL) -> None:
        self.token = token
        self.url = url
        self.session: str | None = None
        self._id = 1

    def _post(self, payload: dict, notify: bool = False):
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            "Authorization": f"Bearer {self.token}",
            # 2025-06-18 Streamable HTTP transport requires this header after
            # initialization; some ERP gateways started enforcing it in 2026-10.
            "MCP-Protocol-Version": "2025-06-18",
        }
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(self.url, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=120) as response:
            text = response.read().decode("utf-8", "replace")
            self.session = response.headers.get("Mcp-Session-Id") or self.session
        if notify:
            return None
        return self._parse(text)

    @staticmethod
    def _parse(text: str):
        """解析 JSON-RPC 响应：SSE 的 `data:` 行与裸 JSON 都支持。

        非 JSON（登录页 HTML、被截断的响应）抛 `SessionLost`，由 `call()` 重登重试。
        """
        if not text.strip():
            return None
        for line in text.splitlines():
            if not line.startswith("data:"):
                continue
            chunk = line[5:].strip()
            if not chunk:
                continue
            try:
                return json.loads(chunk)
            except ValueError as exc:
                raise SessionLost(f"MCP 返回非 JSON 数据：{chunk.lstrip()[:80]!r}") from exc
        try:
            return json.loads(text)
        except ValueError as exc:
            raise SessionLost(f"MCP 返回非 JSON 响应：{text.lstrip()[:80]!r}") from exc

    def _reinitialize(self) -> None:
        """丢掉旧会话重新握手（失败不抛，交给调用方按原错误处理）。"""
        self.session = None
        try:
            self.initialize()
        except Exception:  # noqa: BLE001  重连失败时让上层重试/兜底
            pass

    def initialize(self) -> None:
        self._post(
            {
                "jsonrpc": "2.0",
                "id": self._id,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "cost-matching", "version": "0.1.0"},
                },
            }
        )
        self._id += 1
        self._post({"jsonrpc": "2.0", "method": "notifications/initialized"}, notify=True)

    def list_tools(self) -> list[dict]:
        data = self._post({"jsonrpc": "2.0", "id": self._id, "method": "tools/list"})
        self._id += 1
        return (data or {}).get("result", {}).get("tools", [])

    def call(self, name: str, arguments: dict) -> dict:
        for attempt in range(3):
            try:
                data = self._post(
                    {
                        "jsonrpc": "2.0",
                        "id": self._id,
                        "method": "tools/call",
                        "params": {"name": name, "arguments": arguments},
                    }
                )
            except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, ConnectionError) as exc:
                if attempt < 2:
                    time.sleep(2)
                    continue
                raise
            except SessionLost:
                # 会话被顶掉（返回登录页）：重新握手后重试同一调用。
                self._reinitialize()
                if attempt < 2:
                    time.sleep(2)
                    continue
                raise
            self._id += 1
            result = (data or {}).get("result")
            if result is None:
                error = (data or {}).get("error")
                if error and "Rate limit" in json.dumps(error):
                    time.sleep(2)
                    continue
                raise RuntimeError(f"{name} 调用失败: {error}")
            return result
        raise RuntimeError(f"{name} 重试耗尽")

    def call_text(self, name: str, arguments: dict) -> str:
        for attempt in range(3):
            result = self.call(name, arguments)
            parts = []
            for item in result.get("content", []) or []:
                if item.get("type") == "text":
                    parts.append(str(item.get("text") or ""))
                else:
                    parts.append(json.dumps(item, ensure_ascii=False))
            text = "\n".join(parts)
            if text.strip() or attempt == 2:
                return text
            # 睿贝偶尔返回成功RPC但空文本；先重试，不把空content伪装成业务JSON。
            time.sleep(2)
        return ""


def main() -> None:
    token = sys.argv[1] if len(sys.argv) > 1 else ""
    if not token:
        from app.services.feishu_client import get_config

        token = get_config("ERP_API_KEY")
    client = McpClient(token)
    client.initialize()
    tools = client.list_tools()
    out = [f"工具数={len(tools)}"]
    for tool in tools:
        out.append(f"{tool['name']}\t{(tool.get('description') or '').strip()[:120]}")
    (PROJECT_ROOT / ".cache" / "erp" / "reports").mkdir(parents=True, exist_ok=True)
    (PROJECT_ROOT / ".cache" / "erp" / "reports" / "mcp_tools.txt").write_text(
        "\n".join(out), encoding="utf-8"
    )
    print("\n".join(out))


if __name__ == "__main__":
    main()
