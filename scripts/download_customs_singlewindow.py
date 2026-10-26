"""从中国国际贸易单一窗口下载报关单 PDF 到 tmp/。

复现影刀的正式出口单与预录单两条网页下载流程，不解析、不写飞书：

    python -X utf8 scripts/download_customs_singlewindow.py --flow both --days 30
    python -X utf8 scripts/download_customs_singlewindow.py \
        --flow export --start-date 2026-10-01 --end-date 2026-10-26
    python -X utf8 scripts/download_customs_singlewindow.py \
        --resume both_20261001_20261026

首次运行或登录过期时，在打开的 Chrome 里用卡介质完成一次登录。卡口令可选
通过环境变量 ``SINGLEWINDOW_CARD_PASSWORD`` 提供，代码不会保存口令。
"""
from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from playwright.sync_api import sync_playwright  # noqa: E402

from app.services.customs_download import (  # noqa: E402
    DownloadManifest,
    date_range,
    declaration_index,
    default_run_id,
    download_export_pdfs,
    download_pre_record_pdfs,
    ensure_logged_in,
    launch_persistent_browser,
    load_manifest,
    save_declaration_index,
    save_manifest,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="单一窗口报关单 PDF 下载")
    parser.add_argument(
        "--flow",
        choices=("export", "pre_record", "both"),
        default="both",
        help="export=正式出口单，pre_record=预录放行单，both=按影刀顺序全部执行",
    )
    parser.add_argument("--start-date", default="", help="开始日期，YYYY-MM-DD 或 YYYYMMDD")
    parser.add_argument("--end-date", default="", help="结束日期，默认今天")
    parser.add_argument("--days", type=int, default=30, help="未给开始日期时向前取多少天")
    parser.add_argument(
        "--run-id",
        default="",
        help="断点运行号；默认按 流程_开始_结束 命名",
    )
    parser.add_argument(
        "--resume",
        default="",
        help="只按指定运行号续跑；日期范围从原 manifest 读取",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "customs_download",
        help="输出根目录，默认 tmp/customs_download",
    )
    parser.add_argument(
        "--profile-dir",
        type=Path,
        default=PROJECT_ROOT / "tmp" / "singlewindow_profile",
        help="浏览器持久化目录，默认 tmp/singlewindow_profile",
    )
    parser.add_argument("--channel", default="chrome", help="浏览器通道，默认 chrome")
    parser.add_argument("--max-pages", type=int, default=100)
    parser.add_argument("--login-timeout", type=int, default=300)
    parser.add_argument("--preview-timeout", type=int, default=60)
    parser.add_argument("--page-wait-ms", type=int, default=1200)
    parser.add_argument(
        "--pre-lookback-months",
        type=int,
        default=6,
        help="预录查询固定复刻影刀的“六个月”快捷条件",
    )
    parser.add_argument(
        "--pre-enterprise-category",
        default="C-报关收发货人",
        help="预录查询的企业类别",
    )
    parser.add_argument("--force", action="store_true", help="忽略已下载页并重新下载")
    parser.add_argument("--trace", action="store_true", help="失败时保存 Playwright trace")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run_id = args.resume.strip()
    if run_id:
        run_dir = (args.output_root / run_id).resolve()
        manifest_path = run_dir / "manifest.json"
        manifest = load_manifest(manifest_path)
        if manifest is None:
            raise SystemExit(f"续跑目录缺少 manifest.json：{manifest_path}")
        start, end = manifest.start_date, manifest.end_date
        flow = manifest.flow
    else:
        end_value = args.end_date or datetime.now().strftime("%Y-%m-%d")
        start, end = date_range(args.start_date, end_value, days=args.days)
        flow = args.flow
        run_id = args.run_id.strip() or default_run_id(start, end, flow=flow)
        run_dir = (args.output_root / run_id).resolve()
        manifest_path = run_dir / "manifest.json"
        existing = load_manifest(manifest_path)
        if existing and (
            existing.start_date != start
            or existing.end_date != end
            or existing.flow != flow
        ):
            raise SystemExit(
                f"运行号 {run_id} 已有不同日期或流程的清单，请换 --run-id"
            )
        manifest = existing or DownloadManifest(
            run_id=run_id,
            flow=flow,
            start_date=start,
            end_date=end,
            output_dir=str(run_dir),
        )

    run_dir.mkdir(parents=True, exist_ok=True)
    manifest.status = "running"
    manifest.error = ""
    save_manifest(manifest_path, manifest)
    print(
        f"运行号：{run_id}\n日期：{start} 至 {end}\n输出：{run_dir}",
        flush=True,
    )

    context = None
    trace_path = run_dir / "playwright-trace.zip"
    try:
        with sync_playwright() as playwright:
            context = launch_persistent_browser(
                playwright, profile_dir=args.profile_dir.resolve(), channel=args.channel
            )
            if args.trace:
                context.tracing.start(screenshots=True, snapshots=True)
            page = context.pages[0] if context.pages else context.new_page()
            ensure_logged_in(page, manual_timeout=args.login_timeout)
            if flow in ("export", "both"):
                download_export_pdfs(
                    page,
                    start=start,
                    end=end,
                    run_dir=run_dir,
                    manifest=manifest,
                    max_pages=args.max_pages,
                    page_wait_ms=args.page_wait_ms,
                    preview_timeout_ms=args.preview_timeout * 1000,
                    force=args.force,
                )
            if flow in ("pre_record", "both"):
                download_pre_record_pdfs(
                    page,
                    run_dir=run_dir,
                    manifest=manifest,
                    lookback_months=args.pre_lookback_months,
                    enterprise_category=args.pre_enterprise_category,
                    max_pages=min(args.max_pages, 20),
                    page_wait_ms=args.page_wait_ms,
                    preview_timeout_ms=args.preview_timeout * 1000,
                    force=args.force,
                )
            manifest.status = "complete"
            save_manifest(manifest_path, manifest)
            if args.trace:
                context.tracing.stop(path=str(trace_path))
    except Exception as exc:  # noqa: BLE001
        manifest.status = "failed"
        manifest.error = str(exc)[:1000]
        save_manifest(manifest_path, manifest)
        save_declaration_index(manifest, run_dir)
        if context is not None:
            try:
                context.pages[0].screenshot(path=str(run_dir / "failure.png"), full_page=True)
            except Exception:  # noqa: BLE001
                pass
            if args.trace:
                try:
                    context.tracing.stop(path=str(trace_path))
                except Exception:  # noqa: BLE001
                    pass
        print(f"下载失败：{exc}", file=sys.stderr)
        if args.trace:
            print(f"Playwright trace：{trace_path}", file=sys.stderr)
        traceback.print_exc()
        raise SystemExit(1) from exc
    finally:
        if context is not None:
            try:
                context.close()
            except Exception:  # noqa: BLE001
                pass

    pages = [item for item in manifest.pages if item.status == "downloaded"]
    index = declaration_index(manifest, run_dir)
    print("", flush=True)
    print(
        f"完成：{len(pages)} 个 PDF 下载项，{len(index)} 张报关单，"
        f"{sum(item.bytes for item in pages)} 字节",
        flush=True,
    )
    for item in pages:
        print(
            f"  page {item.page}: {len(item.declaration_numbers)} 张 -> {item.pdf}",
            flush=True,
        )
    print(f"清单：{manifest_path}", flush=True)
    print(f"报关单索引：{run_dir / 'declaration_index.json'}", flush=True)


if __name__ == "__main__":
    main()
