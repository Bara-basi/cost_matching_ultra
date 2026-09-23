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
        }
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(self.url, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=120) as response:
            text = response.read().decode("utf-8", "replace")
            self.session = response.headers.get("Mcp-Session-Id") or self.session
        if notify or not text.strip():
            return None
        for line in text.splitlines():
            if line.startswith("data:"):
                return json.loads(line[5:].strip())
        return json.loads(text)

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
            except urllib.error.HTTPError as exc:
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
        result = self.call(name, arguments)
        parts = []
        for item in result.get("content", []) or []:
            parts.append(item.get("text") or json.dumps(item, ensure_ascii=False))
        return "\n".join(parts)


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
