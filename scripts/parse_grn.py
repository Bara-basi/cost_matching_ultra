"""解析采购单的入库单附件（规则优先 + DeepSeek 兜底）。

用法：
    python scripts/parse_grn.py --limit 5          # 试跑 5 个采购单
    python scripts/parse_grn.py --code 26MT-03R411-HD
    python scripts/parse_grn.py                    # 全量（有缓存则跳过）
    python scripts/parse_grn.py --force            # 忽略缓存重跑
    python scripts/parse_grn.py --no-model         # 只跑规则，不联网

结果缓存在 `.cache/erp/attachments/parsed/<采购单号>/<文件名>.json`。
"""
from __future__ import annotations

import argparse
import collections
import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.services.erp_cache import CACHE_ROOT  # noqa: E402
from app.services.grn_extract import PARSED_ROOT, extract_attachment, safe_name  # noqa: E402

GRN_DIR = CACHE_ROOT / "attachments" / "grn"
REPORT = CACHE_ROOT / "reports" / "grn_parse_summary.json"


def targets(code: str | None, limit: int) -> list[Path]:
    dirs = sorted(d for d in GRN_DIR.iterdir() if d.is_dir())
    if code:
        dirs = [d for d in dirs if d.name == code]
    return dirs[: limit or None]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--code", default="")
    parser.add_argument("--codes-file", default="", help="只跑文件里列出的采购单（每行一个）")
    parser.add_argument(
        "--files-file",
        default="",
        help="定向强制重跑：文件里每行写 `<采购单号>/<附件文件名>`，忽略缓存重解析这几份",
    )
    parser.add_argument("--only-failed", action="store_true", help="只补跑规则失败的附件")
    parser.add_argument(
        "--rerun-rule",
        action="store_true",
        help="按新规则重算一遍（忽略缓存里的 rule 结果）；模型解析结果继续复用、不联网",
    )
    parser.add_argument(
        "--hint",
        default="",
        help="定向重跑时额外告诉模型的提示（只在 --files-file 下有意义）",
    )
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--no-model", action="store_true")
    parser.add_argument("--workers", type=int, default=1, help="并发解析线程数（模型调用占大头）")
    parser.add_argument("--pause", type=float, default=0.0)
    args = parser.parse_args()

    if args.workers > 1:
        from app.services import deepseek_client

        # 每个线程各自节流，整体间隔按并发数缩小
        deepseek_client.set_min_interval(0.8 / args.workers)

    if args.codes_file:
        wanted = {
            line.strip()
            for line in Path(args.codes_file).read_text(encoding="utf-8").splitlines()
            if line.strip()
        }
        dirs = sorted(d for d in GRN_DIR.iterdir() if d.is_dir() and d.name in wanted)
    else:
        dirs = targets(args.code, args.limit)

    forced: set[tuple[str, str]] = set()
    if args.files_file:
        for line in Path(args.files_file).read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            code, _, name = line.partition("/")
            forced.add((code.strip(), name.strip()))
        wanted_codes = {code for code, _name in forced}
        dirs = [d for d in dirs if d.name in wanted_codes]
        print(f"定向重跑附件 {len(forced)} 份（采购单 {len(wanted_codes)} 个）", flush=True)
    todo = dirs if not args.limit else dirs[: args.limit]
    print(f"待处理采购单 {len(todo)} 个", flush=True)
    stats: collections.Counter = collections.Counter()
    problems: list[dict] = []
    started = time.time()
    lock = __import__("threading").Lock()

    def handle(path: Path, code: str, force: bool = False) -> tuple[str, dict]:
        try:
            payload = extract_attachment(
                path,
                purchase_code=code,
                use_model=not args.no_model,
                force=force,
                hint=args.hint if force else "",
            )
        except Exception as exc:  # noqa: BLE001
            return "异常", {"code": code, "file": path.name, "error": str(exc)[:200]}
        record = None
        if payload.get("anomalies"):
            record = {
                "code": code,
                "file": path.name,
                "how": payload.get("_how"),
                "anomalies": payload["anomalies"],
            }
        return payload.get("_how") or "unknown", record

    def cached_how(directory: Path, path: Path) -> str:
        cached = PARSED_ROOT / safe_name(directory.name) / (safe_name(path.name, "file") + ".json")
        if not cached.exists():
            return ""
        try:
            return str(json.loads(cached.read_text(encoding="utf-8")).get("_how") or "")
        except ValueError:
            return ""

    tasks: list[tuple[Path, str, bool]] = []
    for directory in todo:
        for path in sorted(f for f in directory.iterdir() if f.is_file()):
            if forced:
                if (directory.name, path.name) not in forced:
                    continue
                tasks.append((path, directory.name, True))
                continue
            if args.rerun_rule:
                # 只重算规则解析的附件；模型解析结果（含图片/PDF）原样复用，不联网
                if cached_how(directory, path) == "rule":
                    tasks.append((path, directory.name, True))
                continue
            if args.only_failed:
                # 缓存文件名经过 safe_name 清洗，直接拼 path.name 会找不到
                cached = PARSED_ROOT / safe_name(directory.name) / (
                    safe_name(path.name, "file") + ".json"
                )
                if cached.exists():
                    try:
                        payload = json.loads(cached.read_text(encoding="utf-8"))
                    except ValueError:
                        payload = {}
                    if payload.get("_how") not in {"rule_failed", "skipped"}:
                        continue
            tasks.append((path, directory.name, False))
    print(f"待处理附件 {len(tasks)} 份", flush=True)

    def record(how: str, item: dict | None) -> None:
        with lock:
            stats[how] += 1
            if item is not None:
                stats["有疑点"] += 1
                problems.append(item)
            done = sum(stats.values())
            if done % 25 == 0:
                print(
                    f"  [{done}/{len(tasks)}] 规则={stats['rule']} 模型兜底={stats['llm_fallback']} "
                    f"模型={stats['llm']} 疑点={stats['有疑点']} 异常={stats['异常']} "
                    f"用时={time.time() - started:.0f}s",
                    flush=True,
                )

    if args.workers > 1:
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = [
                pool.submit(handle, path, code, force) for path, code, force in tasks
            ]
            for future in as_completed(futures):
                how, item = future.result()
                record(how, item)
    else:
        for path, code, force in tasks:
            how, item = handle(path, code, force)
            record(how, item)
            if args.pause:
                time.sleep(args.pause)

    summary = {
        "finishedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "purchaseCodes": len(todo),
        **{k: v for k, v in stats.items()},
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(
        json.dumps({"summary": summary, "problems": problems[:2000]}, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False))
    print("疑点/异常明细：", REPORT)
    print("解析缓存目录：", PARSED_ROOT)


if __name__ == "__main__":
    main()
