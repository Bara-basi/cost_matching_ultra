"""从飞书「2026年海关AI副本」表下载报关单 PDF 原件（只读、断点续跑）。

背景：项目里 `data/raw/customs_declaration/` 只有 377 份 PDF，财务表里
有 546 张报关单，缺的 178 张里只有少数在这张表上留了原件附件。本脚本按
「报关单号」把能拿到的原件补下来，落盘文件名用附件的 file_token，
并写一份 `报关单号 → 文件` 的清单，供 parse_customs.py 解析。

用法：
    python -X utf8 scripts/fetch_customs_pdfs.py --from-parsed   # 只补「我方已解析集合」之外的
    python -X utf8 scripts/fetch_customs_pdfs.py --decl 224420260011166496
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.request
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

APP_TOKEN = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
TABLE_ID = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
RAW_DIR = PROJECT_ROOT / "data" / "raw" / "customs_declaration"
CACHE = PROJECT_ROOT / "data" / "cache"
MANIFEST = CACHE / "customs_pdf_manifest.json"
PARSED_SHEETS = (
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_出口退税联.xlsx",
    PROJECT_ROOT / "outputs" / "customs_parse" / "报关单解析结果_预录单.xlsx",
)


def parsed_declarations() -> set[str]:
    from openpyxl import load_workbook

    out: set[str] = set()
    for path in PARSED_SHEETS:
        if not path.exists():
            continue
        workbook = load_workbook(path, read_only=True)
        sheet = workbook.active
        rows = list(sheet.iter_rows(values_only=True))
        header = [str(cell) for cell in rows[0]]
        index = header.index("报关单号")
        out |= {str(row[index]).strip() for row in rows[1:] if row[index]}
        workbook.close()
    return out


def attachment_items(value) -> list[dict]:
    items = value if isinstance(value, list) else ([value] if value else [])
    return [item for item in items if isinstance(item, dict) and item.get("file_token")]


def _get(url: str, client: FeishuClient) -> bytes:
    if not url.startswith("http"):
        url = f"https://open.feishu.cn/open-apis{url}"
    request = urllib.request.Request(
        url, headers={"Authorization": f"Bearer {client.tenant_access_token()}"}
    )
    with urllib.request.urlopen(request, timeout=120) as response:
        return response.read()


def download(url: str, destination: Path, token: str, client: FeishuClient) -> int:
    """飞书的附件要跳两跳：先拿 tmp_download_url，再下真正的内容。"""
    data = _get(url, client)
    if data.lstrip()[:1] == b"{":
        payload = json.loads(data.decode("utf-8", "ignore"))
        links = ((payload.get("data") or {}).get("tmp_download_urls")) or []
        inner = str((links[0] or {}).get("tmp_download_url") or "") if links else ""
        if inner:
            data = _get(inner, client)
    destination.write_bytes(data)
    return len(data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--decl", action="append", default=[], help="指定报关单号（可多个）")
    parser.add_argument("--from-parsed", action="store_true", help="补「我方已解析集合」之外的")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    known = parsed_declarations()
    want = {code.strip() for code in args.decl if code.strip()}
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    client = FeishuClient()
    manifest = (
        json.loads(MANIFEST.read_text(encoding="utf-8")) if MANIFEST.exists() else {}
    )
    stats = {"scanned": 0, "target": 0, "downloaded": 0, "skipped": 0, "errors": 0}
    started = time.time()
    for record in client.iter_records(APP_TOKEN, TABLE_ID):
        fields = record.get("fields") or {}
        decl = str(fields.get("报关单号") or "").strip()
        items = attachment_items(fields.get("pdf原件"))
        stats["scanned"] += 1
        if not items or not decl:
            continue
        if want and decl not in want:
            continue
        if args.from_parsed and not want and decl in known:
            continue
        stats["target"] += 1
        for item in items:
            token = str(item["file_token"])
            target = RAW_DIR / f"{token}.pdf"
            entry = manifest.setdefault(decl, {})
            if target.exists() and not args.force:
                stats["skipped"] += 1
                entry[token] = {"file": target.name, "name": item.get("name"), "status": "cached"}
                continue
            url = str(item.get("tmp_url") or "")
            if not url:
                stats["errors"] += 1
                entry[token] = {"name": item.get("name"), "status": "no-url"}
                continue
            try:
                size = download(url, target, token, client)
            except Exception as exc:  # noqa: BLE001
                stats["errors"] += 1
                entry[token] = {"name": item.get("name"), "status": f"error:{exc}"[:120]}
                continue
            stats["downloaded"] += 1
            entry[token] = {
                "file": target.name,
                "name": item.get("name"),
                "size": size,
                "status": "downloaded",
            }
            print(f"  ✔ {decl} {item.get('name')} → {target.name} ({size} B)", flush=True)
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    stats["用时秒"] = round(time.time() - started)
    print(json.dumps(stats, ensure_ascii=False))
    print("清单:", MANIFEST)


if __name__ == "__main__":
    main()
