"""PDF → 解析 → 拆单 → 成本匹配 → 飞书 AI 副本测试写入。"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services import pipeline  # noqa: E402
from app.services.customs_ingest import (  # noqa: E402
    build_records,
    candidate_declarations,
    fill_ai_status_fields,
    parse_formal_pdfs,
    remove_duplicate_copy_attachments,
    verify_existing_ai_copy,
    write_to_copy_ai_fields,
    write_to_original,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="报关 PDF 到飞书 AI 副本测试链路")
    parser.add_argument(
        "--pdf-dir",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_ai_attachments" / "2026-10" / "export",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_pipeline" / datetime.now().strftime("%Y%m%d_%H%M%S"),
    )
    parser.add_argument("--declarations", default="", help="逗号分隔的报关单号；留空自动选择多拆记录")
    parser.add_argument("--limit", type=int, default=3)
    parser.add_argument("--apply", action="store_true", help="实际写飞书；不加只生成计划")
    parser.add_argument("--target", choices=("copy", "original"), default="original")
    parser.add_argument(
        "--verify-existing",
        action="store_true",
        help="只校验指定 work-dir 中已经写入的记录，不重复创建",
    )
    parser.add_argument(
        "--fill-status-fields",
        action="store_true",
        help="清理临时成本列，并为已写入的 AI 记录补齐拆单/成本状态",
    )
    parser.add_argument(
        "--cleanup-copy-attachments",
        action="store_true",
        help="副本测试记录每张报关单只在第一条保留 PDF",
    )
    parser.add_argument("--wait-seconds", type=int, default=45)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    work_dir = args.work_dir.resolve()
    if args.cleanup_copy_attachments:
        result = remove_duplicate_copy_attachments(work_dir)
        print(json.dumps(result, ensure_ascii=False, indent=1))
        return
    if args.fill_status_fields:
        result = fill_ai_status_fields(work_dir)
        print(json.dumps(result, ensure_ascii=False, indent=1))
        return
    if args.verify_existing:
        summary = verify_existing_ai_copy(work_dir)
        print(json.dumps(summary.__dict__, ensure_ascii=False, indent=1))
        return
    work_dir.mkdir(parents=True, exist_ok=True)
    parse_xlsx = parse_formal_pdfs(args.pdf_dir.resolve(), work_dir)
    split_dir = work_dir / "split"
    pipeline.run_round(parse_xlsx, split_dir)
    rows = pipeline.review_rows(split_dir / "拆单明细_全部.xlsx")
    requested = [item.strip() for item in args.declarations.split(",") if item.strip()]
    selected = requested or candidate_declarations(rows, args.limit)
    if len(selected) < args.limit:
        raise SystemExit(f"只有 {len(selected)} 张符合“多记录且全部正常”的报关单")
    marker = f"自动测试-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    source_rows, ai_rows, pdfs = build_records(
        selected, rows, parse_xlsx, args.pdf_dir.resolve(), marker
    )
    plan = {
        "marker": marker,
        "declarations": selected,
        "sourceRecords": source_rows,
        "aiRecords": ai_rows,
        "pdfs": {key: str(value) for key, value in pdfs.items()},
    }
    (work_dir / "write_plan.json").write_text(
        json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "declarations": selected,
                "sourceRows": len(source_rows),
                "aiRows": len(ai_rows),
                "plan": str(work_dir / "write_plan.json"),
                "apply": bool(args.apply),
            },
            ensure_ascii=False,
        )
    )
    if not args.apply:
        return
    if args.target == "original":
        summary = write_to_original(
            declarations=selected,
            source_rows=source_rows,
            ai_rows=ai_rows,
            pdfs=pdfs,
            work_dir=work_dir,
            marker=marker,
            wait_seconds=args.wait_seconds,
        )
    else:
        summary = write_to_copy_ai_fields(
            declarations=selected,
            source_rows=source_rows,
            ai_rows=ai_rows,
            pdfs=pdfs,
            work_dir=work_dir,
            marker=marker,
            wait_seconds=args.wait_seconds,
        )
    print(json.dumps(summary.__dict__, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
