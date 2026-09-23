"""通过睿贝 MCP 抓取采购附件清单并下载「入库单」文件。

MCP `purchase.find` 的 `attachmentList` 给出每个附件的 `downloadUrl`
（形如 `https://erp.mtholdinggroup.com/tempFile/<HASH>`），
可带浏览器会话直接 GET 下载，比网页上的「打包下载」更完整、更可靠。
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.services.erp_cache import CACHE_ROOT, ensure_dir

ATTACH_ROOT = CACHE_ROOT / "attachments"
LIST_DIR = ATTACH_ROOT / "lists"          # 每个采购单的附件清单
GRN_DIR = ATTACH_ROOT / "grn"             # 解出的入库单文件
MANIFEST = ATTACH_ROOT / "grn_manifest.json"
GRN_KEYWORD = "入库单"
SAFE_RE = re.compile(r"[^\w\u4e00-\u9fff.\-]")


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def safe_name(text: str, fallback: str = "file") -> str:
    cleaned = SAFE_RE.sub("_", str(text or "")).strip("_")
    return cleaned or fallback


def list_path(purchase_code: str) -> Path:
    return LIST_DIR / f"{safe_name(purchase_code)}.json"


def load_list(purchase_code: str) -> dict[str, Any] | None:
    path = list_path(purchase_code)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None


def save_list(purchase_code: str, payload: dict[str, Any]) -> None:
    path = list_path(purchase_code)
    ensure_dir(path.parent)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")


def parse_purchase_payload(text: str) -> dict[str, Any]:
    """把 MCP 返回的 purchase.find 文本解析成 dict。"""
    try:
        payload = json.loads(text)
    except ValueError:
        return {}
    if not payload.get("success"):
        return {}
    return payload.get("value") or {}


def grn_attachments(value: dict[str, Any]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for item in value.get("attachmentList") or []:
        name = str(item.get("attachmentName") or "")
        url = str(item.get("downloadUrl") or "")
        if GRN_KEYWORD in name and url:
            out.append({"name": name, "url": url, "group": item.get("attachmentGroup") or ""})
    return out


def download_file(page, url: str, destination: Path) -> int:
    """用浏览器会话下载单个附件。"""
    ensure_dir(destination.parent)
    response = page.request.get(url, timeout=120_000)
    if response.status != 200:
        raise RuntimeError(f"HTTP {response.status}")
    body = response.body()
    destination.write_bytes(body)
    return len(body)


def load_manifest() -> dict[str, Any]:
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text(encoding="utf-8"))
    return {}


def save_manifest(manifest: dict[str, Any]) -> None:
    ensure_dir(ATTACH_ROOT)
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
