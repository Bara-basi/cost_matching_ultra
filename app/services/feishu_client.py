"""飞书开放平台客户端：负责鉴权与多维表格（Bitable）读写。

仅依赖标准库，避免在受限环境下额外安装依赖。
配置从项目根目录的 .env 读取，环境变量优先。
"""
from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Iterator

BASE_URL = "https://open.feishu.cn/open-apis"
PROJECT_ROOT = Path(__file__).resolve().parents[2]


def load_env(path: Path | None = None) -> dict[str, str]:
    """读取 .env 文件（UTF-8），返回键值对。同名字段后者覆盖前者。"""
    env_path = path or (PROJECT_ROOT / ".env")
    values: dict[str, str] = {}
    if not env_path.exists():
        return values
    for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        values[key.strip()] = value.strip()
    return values


def get_config(key: str, default: str | None = None) -> str:
    """取配置项：优先环境变量，其次 .env。"""
    if key in os.environ and os.environ[key]:
        return os.environ[key]
    env = load_env()
    if key in env and env[key]:
        return env[key]
    if default is not None:
        return default
    raise KeyError(f"缺少配置项: {key}")


class FeishuError(RuntimeError):
    """飞书接口返回非 0 code 时抛出。"""

    def __init__(self, code: int, msg: str, payload: dict[str, Any] | None = None):
        super().__init__(f"飞书接口错误 code={code} msg={msg}")
        self.code = code
        self.msg = msg
        self.payload = payload or {}


class FeishuClient:
    """飞书多维表格客户端，内部缓存 tenant_access_token。"""

    def __init__(
        self,
        app_id: str | None = None,
        app_secret: str | None = None,
        timeout: int = 30,
    ) -> None:
        self.app_id = app_id or get_config("LARK_APP_ID")
        self.app_secret = app_secret or get_config("LARK_APP_SECRET")
        self.timeout = timeout
        self._token: str | None = None
        self._token_expire_at: float = 0.0

    # ---------- 基础请求 ----------

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        body: dict[str, Any] | None = None,
        auth: bool = True,
    ) -> dict[str, Any]:
        url = f"{BASE_URL}{path}"
        if params:
            cleaned = {k: v for k, v in params.items() if v is not None}
            if cleaned:
                url = f"{url}?{urllib.parse.urlencode(cleaned)}"
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if auth:
            headers["Authorization"] = f"Bearer {self.tenant_access_token()}"
        data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:  # 保留服务端返回体，便于排查
            detail = exc.read().decode("utf-8", errors="replace")
            raise FeishuError(exc.code, f"HTTP {exc.code}: {detail}") from exc
        if payload.get("code") != 0:
            # token 失效时清空缓存，交由调用方重试
            if payload.get("code") in (99991663, 99991661, 99991664):
                self._token = None
            raise FeishuError(payload.get("code", -1), payload.get("msg", "unknown"), payload)
        return payload

    def tenant_access_token(self) -> str:
        if self._token and time.time() < self._token_expire_at:
            return self._token
        payload = self._request(
            "POST",
            "/auth/v3/app_access_token/internal",
            body={"app_id": self.app_id, "app_secret": self.app_secret},
            auth=False,
        )
        self._token = payload["tenant_access_token"]
        self._token_expire_at = time.time() + int(payload.get("expire", 7200)) - 300
        return self._token

    # ---------- 多维表格 ----------

    def list_tables(self, app_token: str) -> list[dict[str, Any]]:
        return self._request("GET", f"/bitable/v1/apps/{app_token}/tables", params={"page_size": 100})["data"][
            "items"
        ]

    def list_fields(self, app_token: str, table_id: str) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        page_token: str | None = None
        while True:
            data = self._request(
                "GET",
                f"/bitable/v1/apps/{app_token}/tables/{table_id}/fields",
                params={"page_size": 100, "page_token": page_token},
            )["data"]
            items.extend(data.get("items", []))
            if not data.get("has_more"):
                break
            page_token = data.get("page_token")
        return items

    def iter_records(
        self,
        app_token: str,
        table_id: str,
        *,
        view_id: str | None = None,
        page_size: int = 500,
        field_names: list[str] | None = None,
    ) -> Iterator[dict[str, Any]]:
        """按页迭代记录，逐条 yield。"""
        page_token: str | None = None
        while True:
            params: dict[str, Any] = {"page_size": page_size, "page_token": page_token}
            if view_id:
                params["view_id"] = view_id
            if field_names:
                params["field_names"] = json.dumps(field_names, ensure_ascii=False)
            data = self._request(
                "GET",
                f"/bitable/v1/apps/{app_token}/tables/{table_id}/records",
                params=params,
            )["data"]
            for item in data.get("items", []):
                yield item
            if not data.get("has_more"):
                break
            page_token = data.get("page_token")

    def list_records(self, app_token: str, table_id: str, **kwargs: Any) -> list[dict[str, Any]]:
        return list(self.iter_records(app_token, table_id, **kwargs))
