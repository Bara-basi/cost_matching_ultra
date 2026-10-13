"""只读检查区间内出运单、报关任务及其历史采购依赖的缓存缺口。"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT, now_iso, read_jsonl
from app.services.erp_attachment_mcp import safe_name
from app.services.shipment_index import normalise_purchase_code, strip_pi
from app.services.shipment_detail import load_shipment_detail, shipment_invoices_for_contract


def audit(start: str, end: str) -> dict:
    issues = []
    shipments = [row for row in read_jsonl(CACHE_ROOT / "shipments/shipments.jsonl")
                 if start <= str(row.get("shipDate") or "")[:10] <= end]
    purchases = read_jsonl(CACHE_ROOT / "purchases/purchases.jsonl")
    new_purchases = [row for row in purchases
                     if start <= str(row.get("purchase_date") or "")[:10] <= end]
    invoices = {str(row.get("invoiceCode") or "") for row in shipments}
    contracts = set()
    # 报关日期与出运日期不同：同时检查该区间实际核算任务引用的旧出运单。
    jobs = PROJECT_ROOT / ".runtime/jobs"
    for directory in jobs.iterdir() if jobs.exists() else []:
        status_path, source_path = directory / "status.json", directory / "source_records.json"
        if not status_path.exists() or not source_path.exists():
            continue
        status = json.loads(status_path.read_text(encoding="utf-8"))
        filters = status.get("selectionFilters") or {}
        if status.get("kind") != "sync" or not filters.get("start_date"):
            continue
        if filters["start_date"] > end or (filters.get("end_date") or end) < start:
            continue
        for row in json.loads(source_path.read_text(encoding="utf-8")):
            contract = str(row.get("合同号_1") or "").strip()
            if contract:
                contracts.add(contract)
                invoices.update(shipment_invoices_for_contract(contract) or [contract])
    dependencies = set()
    for invoice in sorted(invoices):
        detail = load_shipment_detail(invoice)
        if not detail:
            issues.append({"kind": "出运明细缺失", "code": invoice})
            continue
        for row in detail.get("productList") or []:
            code = str(row.get("采购订单号") or "").strip()
            if code:
                dependencies.add(normalise_purchase_code(code))
    # 缺明细时仍从已有三单索引和采购列表识别历史采购依赖。
    from app.services.contract_shipments import cores_of
    wanted_cores = {core for contract in contracts | invoices for core in cores_of(contract)}
    for row in purchases:
        if cores_of(str(row.get("orderCode") or "")) & wanted_cores:
            dependencies.add(normalise_purchase_code(str(row.get("purchase_code") or "")))
    all_codes = dependencies | {normalise_purchase_code(str(row.get("purchase_code") or ""))
                                for row in new_purchases}
    lists = {}
    for path in (CACHE_ROOT / "attachments/lists").glob("*.json"):
        record = json.loads(path.read_text(encoding="utf-8"))
        lists[normalise_purchase_code(record.get("purchaseCode") or path.stem)] = record
    files_checked = 0
    for code in sorted(all_codes - {""}):
        record = lists.get(code)
        if not record:
            issues.append({"kind": "附件清单未获取", "code": code,
                           "costDependency": code in dependencies})
        else:
            stamp = str(record.get("fetchedAt") or "")[:10]
            if stamp < start:
                issues.append({"kind": "附件清单早于核算区间，需刷新", "code": code, "fetchedAt": stamp})
            for item in record.get("attachmentList") or []:
                name = str(item.get("attachmentName") or "")
                if "入库" not in name:
                    continue
                paths = [CACHE_ROOT / "attachments/grn" / safe_name(alias) / safe_name(name)
                         for alias in (code, record.get("purchaseCode") or code)]
                if not any(path.exists() and path.stat().st_size for path in paths):
                    issues.append({"kind": "清单中的入库附件未下载", "code": code, "file": name})
        folders = {CACHE_ROOT / "attachments/grn" / safe_name(code)}
        if record:
            folders.add(CACHE_ROOT / "attachments/grn" / safe_name(record.get("purchaseCode") or code))
        for folder in folders:
            for path in folder.iterdir() if folder.exists() else []:
                if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg"}:
                    continue
                files_checked += 1
                parsed_path = CACHE_ROOT / "attachments/parsed" / folder.name / (safe_name(path.name) + ".json")
                if not parsed_path.exists():
                    issues.append({"kind": "已下载附件未解析", "code": code, "file": path.name})
                    continue
                parsed = json.loads(parsed_path.read_text(encoding="utf-8"))
                if parsed.get("_digest") != hashlib.sha256(path.read_bytes()).hexdigest()[:16]:
                    issues.append({"kind": "附件已更新但解析缓存过期", "code": code, "file": path.name})
                if parsed.get("_how") in {"rule_failed", "skipped"}:
                    issues.append({"kind": "附件解析失败", "code": code, "file": path.name})
    counts = {kind: sum(row["kind"] == kind for row in issues) for kind in sorted({row["kind"] for row in issues})}
    return {"checkedAt": now_iso(), "scope": [start, end], "shipmentListRows": len(shipments),
            "customsContracts": len(contracts), "invoicesChecked": len(invoices),
            "newPurchaseRows": len(new_purchases), "dependentPurchaseCodes": sorted(dependencies),
            "purchaseCodesChecked": len(all_codes), "attachmentFilesChecked": files_checked,
            "counts": counts, "issues": issues,
            "note": "仅检查本地可证实的缺口。远端未知新增附件无法离线识别；未获取清单不等于存在入库凭证。"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", required=True)
    parser.add_argument("--end", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = audit(args.start, args.end)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in result.items() if key not in {"issues", "dependentPurchaseCodes"}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
