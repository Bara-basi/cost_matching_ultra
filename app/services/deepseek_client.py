"""DeepSeek 客户端（OpenAI 兼容接口）。

用途：解析版式不固定的入库单附件（表格文本 / 图片）。
key 取 `.env` 的 `DEEPSEEK_API_KEY`；限速与重试按项目惯例（长线任务需可续跑）。

实测（2026-09-24）：`deepseek-chat` 会被路由到 `deepseek-flash`
（1M 上下文，`input_modalities` 含 image，可直接读图）。
"""
from __future__ import annotations

import base64
import json
import re
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

API_ROOT = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"

MIN_INTERVAL = 0.8  # 秒；官方限速之外自己也留一点余量
_lock = threading.Lock()
_last_call = [0.0]


def set_min_interval(seconds: float) -> None:
    """并发跑时把间隔调小（每个线程各自间隔即可）。"""
    global MIN_INTERVAL
    MIN_INTERVAL = max(0.0, float(seconds))


class DeepSeekError(RuntimeError):
    """调用失败（网络 / 限速 / 返回异常）。"""


def _throttle() -> None:
    with _lock:
        wait = MIN_INTERVAL - (time.time() - _last_call[0])
        if wait > 0:
            time.sleep(wait)
        _last_call[0] = time.time()


def api_key() -> str:
    from app.services.feishu_client import get_config

    return get_config("DEEPSEEK_API_KEY")


def _post(path: str, payload: dict[str, Any], timeout: int = 180) -> dict[str, Any]:
    request = urllib.request.Request(
        f"{API_ROOT}{path}",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key()}",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def chat(
    messages: list[dict[str, Any]],
    *,
    model: str = DEFAULT_MODEL,
    json_mode: bool = False,
    temperature: float = 0.0,
    max_tokens: int = 8192,
    retries: int = 3,
    timeout: int = 300,
) -> str:
    """调用对话接口，返回文本内容。失败自动重试（网络抖动 / 限速）。"""
    payload: dict[str, Any] = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}

    last_error: Exception | None = None
    for attempt in range(retries):
        _throttle()
        try:
            body = _post("/chat/completions", payload, timeout)
            return body["choices"][0]["message"]["content"] or ""
        except urllib.error.HTTPError as exc:  # HTTP 层
            detail = ""
            try:
                detail = exc.read().decode("utf-8", "replace")[:300]
            except Exception:  # noqa: BLE001
                pass
            last_error = DeepSeekError(f"HTTP {exc.code}: {detail}")
            if exc.code in {429, 500, 502, 503, 504}:
                time.sleep(2 * (attempt + 1))
                continue
            raise last_error
        except Exception as exc:  # noqa: BLE001  网络抖动
            last_error = exc
            time.sleep(2 * (attempt + 1))
    raise DeepSeekError(f"重试耗尽: {last_error}")


def _strip_fence(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*\s*", "", cleaned)
        cleaned = re.sub(r"```\s*$", "", cleaned)
    return cleaned.strip()


def parse_json(text: str) -> Any:
    """把模型返回的文本解析成 JSON；容忍 ```json 围栏与前后多余文字。"""
    cleaned = _strip_fence(text)
    try:
        return json.loads(cleaned)
    except ValueError:
        pass
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start >= 0 and end > start:
        return json.loads(cleaned[start : end + 1])
    raise DeepSeekError(f"返回不是合法 JSON: {cleaned[:200]}")


def chat_json(
    system: str,
    user: str,
    *,
    images: list[Path | str] | None = None,
    model: str = DEFAULT_MODEL,
    max_tokens: int = 8192,
    retries: int = 3,
) -> Any:
    """要求模型返回 JSON。`images` 里的文件会被转成 base64 一并发送。"""
    content: list[dict[str, Any]] = [{"type": "text", "text": user}]
    for item in images or []:
        path = Path(item)
        mime = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
        }.get(path.suffix.lower(), "image/png")
        encoded = base64.b64encode(path.read_bytes()).decode("ascii")
        content.append(
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{encoded}"}}
        )
    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": content},
    ]
    text = chat(
        messages,
        model=model,
        json_mode=True,
        max_tokens=max_tokens,
        retries=retries,
    )
    return parse_json(text)
