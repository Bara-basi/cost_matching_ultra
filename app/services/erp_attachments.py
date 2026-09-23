"""睿贝采购附件下载。

路径：`/purchase_toUpdate?openWindow=Y&type=view&id=<purchaseId>`
页面上的下载按钮带 `downloadfileinfo` 属性（内含 fileName 列表），
点击后用 Playwright 的 download 事件保存 ZIP。
按要求只保留文件名包含「入库单」的附件（在清单里标注，ZIP 原样保存）。
"""
from __future__ import annotations

import json
import re
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from app.services.erp_cache import CACHE_ROOT, ensure_dir
from app.services.erp_web import ERP_ROOT

ATTACH_DIR = CACHE_ROOT / "attachments"
MANIFEST = ATTACH_DIR / "manifest.json"
GRN_KEYWORD = "入库单"


def load_manifest() -> dict:
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text(encoding="utf-8"))
    return {}


def save_manifest(manifest: dict) -> None:
    ensure_dir(ATTACH_DIR)
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def list_attachment_names(page) -> list[str]:
    """读下载按钮上的 fileName 列表。"""
    button = page.locator('a[onclick*="downloadZip"][downloadfileinfo]').first
    button.wait_for(state="visible", timeout=45_000)
    attr = button.get_attribute("downloadfileinfo") or ""
    names = re.findall(r"fileName\s*:\s*'([^']+)'", attr)
    if not names:
        names = re.findall(r'fileName\s*:\s*"([^"]+)"', attr)
    return names


def download_purchase_attachments(
    page,
    purchase_id: str,
    purchase_code: str,
    *,
    overwrite: bool = False,
) -> dict:
    """下载单个采购单的附件 ZIP，返回结果描述。"""
    safe = re.sub(r"[^\w\u4e00-\u9fff.\-]", "_", purchase_code) or purchase_id
    target = ATTACH_DIR / "downloads" / f"{safe}.zip"
    ensure_dir(target.parent)
    if target.exists() and not overwrite:
        with zipfile.ZipFile(target) as archive:
            names = archive.namelist()
        return {
            "purchaseCode": purchase_code,
            "purchaseId": purchase_id,
            "status": "cached",
            "attachments": names,
            "hasGrn": any(GRN_KEYWORD in n for n in names),
            "path": str(target),
        }

    page.goto(
        f"{ERP_ROOT}/purchase_toUpdate?openWindow=Y&type=view&id={purchase_id}",
        wait_until="domcontentloaded",
        timeout=60_000,
    )
    names = list_attachment_names(page)
    temporary = target.with_suffix(".zip.part")
    with page.expect_download(timeout=120_000) as event:
        page.locator('a[onclick*="downloadZip"][downloadfileinfo]').first.click()
    event.value.save_as(str(temporary))
    with zipfile.ZipFile(temporary) as archive:
        if archive.testzip() is not None:
            raise ValueError("附件压缩包校验失败")
        inner = archive.namelist()
    temporary.replace(target)
    return {
        "purchaseCode": purchase_code,
        "purchaseId": purchase_id,
        "status": "downloaded",
        "attachments": inner,
        "hasGrn": any(GRN_KEYWORD in n for n in inner),
        "listedNames": names,
        "path": str(target),
        "capturedAt": now_iso(),
    }


def extract_grn_files(purchase_code: str) -> list[Path]:
    """从 ZIP 里解出含「入库单」的文件（供解析入库单用）。"""
    safe = re.sub(r"[^\w\u4e00-\u9fff.\-]", "_", purchase_code)
    archive_path = ATTACH_DIR / "downloads" / f"{safe}.zip"
    if not archive_path.exists():
        return []
    out_dir = ATTACH_DIR / "grn" / safe
    ensure_dir(out_dir)
    extracted: list[Path] = []
    with zipfile.ZipFile(archive_path) as archive:
        for name in archive.namelist():
            if GRN_KEYWORD not in name or name.endswith("/"):
                continue
            member = Path(name).name
            destination = out_dir / member
            with archive.open(name) as src, destination.open("wb") as dst:
                shutil.copyfileobj(src, dst)
            extracted.append(destination)
    return extracted
