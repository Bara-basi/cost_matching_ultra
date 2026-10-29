"""从「2026年报关记录 AI」下载指定月份的 PDF 附件到 tmp/（只读）。

默认读取主表 ``MT_FINANCE_AI_*``，筛选出口日期所在月份和销售方式后，
只下载 ``pdf原件`` 中的 PDF。不会修改飞书。

用法：
    python -X utf8 scripts/fetch_ai_customs_attachments.py
    python -X utf8 scripts/fetch_ai_customs_attachments.py \
        --month 2026-10 --sales-mode 出口 --sales-mode 预录
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import threading
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.feishu_client import FeishuClient, get_config  # noqa: E402

CHINA_TZ = timezone(timedelta(hours=8))
PDF_MAGIC = b"%PDF-"


def plain_text(value: Any) -> str:
    """把飞书文本、多选、公式等字段压成可比较文本。"""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return " ".join(filter(None, (plain_text(item) for item in value))).strip()
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return plain_text(value[key])
    return str(value).strip()


def date_text(value: Any) -> str:
    """飞书日期通常为毫秒时间戳；同时兼容字符串日期。"""
    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value / 1000, tz=CHINA_TZ).strftime("%Y-%m-%d")
    raw = plain_text(value)
    if re.fullmatch(r"\d{13}", raw):
        return datetime.fromtimestamp(int(raw) / 1000, tz=CHINA_TZ).strftime("%Y-%m-%d")
    match = re.match(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", raw)
    return f"{int(match[1]):04d}-{int(match[2]):02d}-{int(match[3]):02d}" if match else raw[:10]


def normalize_month(value: str) -> str:
    raw = str(value or "").strip()
    match = re.fullmatch(r"(\d{4})[-/]?(\d{1,2})", raw)
    if not match:
        raise ValueError("月份格式应为 YYYY-MM")
    month = int(match.group(2))
    if not 1 <= month <= 12:
        raise ValueError("月份应在 1 至 12 之间")
    return f"{int(match.group(1)):04d}-{month:02d}"


def safe_filename(value: str, fallback: str = "attachment") -> str:
    name = re.split(r"[\\/]", str(value or ""))[-1]
    name = re.sub(r'[<>:"/\\|?*]+', "_", name).strip(" .")
    return name or fallback


def attachment_items(value: Any) -> list[dict]:
    items = value if isinstance(value, list) else ([value] if value else [])
    return [
        item
        for item in items
        if isinstance(item, dict)
        and item.get("file_token")
        and str(item.get("type") or "application/pdf").lower() in ("application/pdf", "pdf")
    ]


def _urlopen(url: str, token: str, *, timeout: int = 120) -> bytes:
    request = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def download_attachment_bytes(attachment: dict, token: str) -> bytes:
    """兼容飞书临时下载地址与官方 media 下载地址。"""
    errors: list[str] = []
    tmp_url = str(attachment.get("tmp_url") or "")
    direct_url = str(attachment.get("url") or "")
    if tmp_url:
        try:
            data = _urlopen(tmp_url, token)
            if data.lstrip().startswith(b"{"):
                payload = json.loads(data.decode("utf-8", "ignore"))
                links = ((payload.get("data") or {}).get("tmp_download_urls") or [])
                inner = str((links[0] or {}).get("tmp_download_url") or "") if links else ""
                if inner:
                    data = _urlopen(inner, token)
            if data.startswith(PDF_MAGIC):
                return data
            errors.append("临时下载地址返回的不是 PDF")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"临时地址：{exc}")
    if direct_url:
        try:
            data = _urlopen(direct_url, token)
            if data.startswith(PDF_MAGIC):
                return data
            errors.append("官方下载地址返回的不是 PDF")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"官方地址：{exc}")
    raise RuntimeError("；".join(errors) or "附件没有可用的下载地址")


def _write_atomic(path: Path, content: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".part")
    temporary.write_bytes(content)
    os.replace(temporary, path)


def is_pdf_file(path: Path) -> bool:
    try:
        if path.stat().st_size < 5:
            return False
        with path.open("rb") as handle:
            return handle.read(5) == PDF_MAGIC
    except OSError:
        return False


def _save_manifest(path: Path, payload: dict) -> None:
    payload["updatedAt"] = datetime.now(CHINA_TZ).isoformat(timespec="seconds")
    temporary = path.with_name(path.name + ".part")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(temporary, path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="下载飞书主表指定月份的报关单 PDF")
    parser.add_argument("--month", default=datetime.now(CHINA_TZ).strftime("%Y-%m"))
    parser.add_argument(
        "--sales-mode",
        action="append",
        choices=("出口", "预录"),
        default=None,
        help="可重复指定；默认出口和预录都下载",
    )
    parser.add_argument(
        "--source",
        choices=("main", "copy"),
        default="main",
        help="main=主表，copy=副本，默认主表",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_ai_attachments",
    )
    parser.add_argument("--workers", type=int, default=6)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    month = normalize_month(args.month)
    modes = set(args.sales_mode or ["出口", "预录"])
    if args.source == "main":
        app_token = get_config("MT_FINANCE_AI_APP_TOKEN")
        table_id = get_config("MT_FINANCE_AI_TABLE_ID")
    else:
        app_token = get_config("MT_FINANCE_AI_COPY_APP_TOKEN")
        table_id = get_config("MT_FINANCE_AI_COPY_TABLE_ID")
    output_dir = (args.output_root / month).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output_dir / "manifest.json"
    manifest = (
        json.loads(manifest_path.read_text(encoding="utf-8"))
        if manifest_path.exists()
        else {"source": "main" if args.source == "main" else "copy", "month": month,
              "records": {}, "files": {}, "stats": {}}
    )
    source_name = "main" if args.source == "main" else "copy"
    if manifest.get("month") not in (None, month):
        raise SystemExit(
            f"输出目录已有其他月份的清单：{manifest.get('month')}，请换 --output-root"
        )
    if manifest.get("source") not in (None, source_name):
        raise SystemExit(
            f"输出目录已有其他来源表的清单：{manifest.get('source')}，请换 --output-root"
        )
    manifest["source"] = source_name
    manifest["month"] = month

    client = FeishuClient(timeout=120)
    records = list(client.iter_records(app_token, table_id))
    selected: list[dict] = []
    for record in records:
        fields = record.get("fields") or {}
        export_date = date_text(fields.get("出口日期"))
        sales_mode = plain_text(fields.get("销售方式"))
        if not export_date.startswith(month) or sales_mode not in modes:
            continue
        declaration_no = plain_text(fields.get("报关单号"))
        attachments = attachment_items(fields.get("pdf原件"))
        item = {
            "record_id": str(record.get("record_id") or ""),
            "export_date": export_date,
            "sales_mode": sales_mode,
            "declaration_no": declaration_no,
            "contract": plain_text(fields.get("合同号_1")),
            "product": plain_text(fields.get("报关品名")),
            "attachments": attachments,
        }
        selected.append(item)
        manifest["records"][item["record_id"]] = {
            key: value for key, value in item.items() if key != "attachments"
        }
        for attachment in attachments:
            token = str(attachment["file_token"])
            manifest["files"].setdefault(
                token,
                {
                    "file_token": token,
                    "name": str(attachment.get("name") or ""),
                    "size": int(attachment.get("size") or 0),
                    "sales_mode": sales_mode,
                    "declaration_no": declaration_no,
                    "status": "pending",
                },
            )

    unique_tokens = {
        str(item["file_token"]): item
        for record in selected
        for item in record["attachments"]
    }
    print(
        f"月份 {month}，销售方式 {'/'.join(sorted(modes))}："
        f"记录 {len(selected)} 条，附件 {len(unique_tokens)} 个",
        flush=True,
    )
    if args.dry_run:
        return

    access_token = client.tenant_access_token()
    hash_owners: dict[str, str] = {}
    hash_lock = threading.Lock()
    for token, meta in manifest["files"].items():
        if meta.get("status") in ("downloaded", "duplicate") and meta.get("sha256"):
            hash_owners.setdefault(str(meta["sha256"]), str(meta.get("file") or ""))

    def download_one(file_token: str, attachment: dict) -> dict:
        prior = manifest["files"].get(file_token) or {}
        prior_path = output_dir / str(prior.get("file") or "")
        if (
            not args.force
            and prior.get("status") in ("downloaded", "duplicate")
            and is_pdf_file(prior_path)
        ):
            return {
                "file_token": file_token,
                "status": "cached" if prior.get("status") == "downloaded" else "duplicate",
                "file": prior["file"],
                "sha256": prior.get("sha256", ""),
                "bytes": prior.get("bytes", 0),
            }
        content = download_attachment_bytes(attachment, access_token)
        digest = hashlib.sha256(content).hexdigest()
        mode_dir = "export" if prior.get("sales_mode") == "出口" else "pre_record"
        declaration_no = str(prior.get("declaration_no") or "unknown")
        original = safe_filename(str(attachment.get("name") or f"{file_token}.pdf"))
        stem = Path(original).stem
        name = f"{declaration_no}__{digest[:8]}__{stem}.pdf"
        target = output_dir / mode_dir / name
        with hash_lock:
            if digest in hash_owners and hash_owners[digest]:
                return {
                    "file_token": file_token,
                    "status": "duplicate",
                    "file": hash_owners[digest],
                    "sha256": digest,
                }
            _write_atomic(target, content)
            hash_owners[digest] = str(target.relative_to(output_dir))
        return {
            "file_token": file_token,
            "status": "downloaded",
            "file": str(target.relative_to(output_dir)),
            "sha256": digest,
            "bytes": len(content),
        }

    started = time.time()
    failure_count = 0
    with ThreadPoolExecutor(max_workers=max(1, min(args.workers, 12))) as pool:
        futures = {
            pool.submit(download_one, token, attachment): token
            for token, attachment in unique_tokens.items()
        }
        for position, future in enumerate(as_completed(futures), 1):
            token = futures[future]
            try:
                result = future.result()
            except Exception as exc:  # noqa: BLE001
                failure_count += 1
                result = {
                    "file_token": token,
                    "status": "failed",
                    "error": str(exc)[:500],
                }
            entry = manifest["files"].setdefault(token, {})
            entry.update(result)
            if result.get("sha256"):
                hash_owners.setdefault(result["sha256"], result.get("file", ""))
            _save_manifest(manifest_path, manifest)
            if position % 10 == 0 or position == len(futures):
                print(f"  附件进度 {position}/{len(futures)}", flush=True)

    statuses: dict[str, int] = {}
    for token in unique_tokens:
        item = manifest["files"].get(token) or {}
        status = str(item.get("status") or "pending")
        statuses[status] = statuses.get(status, 0) + 1
    manifest["stats"] = {
        "scanned": len(records),
        "selectedRecords": len(selected),
        "selectedByMode": {
            mode: sum(record["sales_mode"] == mode for record in selected)
            for mode in sorted(modes)
        },
        "uniqueAttachments": len(unique_tokens),
        "statuses": statuses,
        "elapsedSeconds": round(time.time() - started, 1),
    }
    _save_manifest(manifest_path, manifest)
    print(json.dumps(manifest["stats"], ensure_ascii=False, indent=2))
    print(f"清单：{manifest_path}")
    if failure_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
