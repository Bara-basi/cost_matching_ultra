"""睿贝数据同步的启动器、状态读取与**定时全量同步**。

定时策略（2026-09-29 用户确认：12 小时启动一次重跑即可）：
- 应用启动时先做一次「按需同步」（数据还新就不跑）；
- 之后由后台调度线程每 10 分钟检查一次，距上次同步超过 `ERP_SYNC_INTERVAL_HOURS`
  （默认 12 小时）就再跑一轮，保证本地缓存持续跟上睿贝；
- 同步任务本身在子进程里跑（`scripts/erp_sync.py`），失败不影响工作台。

可用环境变量：
  ERP_SYNC_ENABLED=0/1            是否启用定时同步（默认 1）
  ERP_SYNC_INTERVAL_HOURS=12      多久同步一次
  ERP_SYNC_REFRESH_DAYS=7         出运明细超过多少天重抓
  ERP_SYNC_ATTACHMENTS=0/1        同步时是否补抓采购附件/入库单（默认 1）
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

from app.services.erp_cache import CACHE_ROOT

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SYNC_SCRIPT = PROJECT_ROOT / "scripts" / "erp_sync.py"
STATE_PATH = CACHE_ROOT / "sync_state.json"
LOG_PATH = CACHE_ROOT / "reports" / "erp_sync.log"

_LOCK = threading.Lock()
_PROC: subprocess.Popen | None = None
_SCHEDULER: threading.Thread | None = None
_SCHEDULER_LOCK = threading.Lock()

# 调度参数（可用环境变量覆盖）
SYNC_ENABLED = str(os.environ.get("ERP_SYNC_ENABLED", "1")).strip().lower() not in {"0", "false", "no", "off"}
INTERVAL_HOURS = float(os.environ.get("ERP_SYNC_INTERVAL_HOURS", "12") or 12)
REFRESH_DAYS = float(os.environ.get("ERP_SYNC_REFRESH_DAYS", "7") or 7)
WITH_ATTACHMENTS = str(os.environ.get("ERP_SYNC_ATTACHMENTS", "1")).strip().lower() not in {"0", "false", "no", "off"}
CHECK_SECONDS = 600  # 每 10 分钟检查一次是否到了下一次同步时间


def _read_state() -> dict:
    """直接读同步状态文件（不走 status()，避免递归）。"""
    if not STATE_PATH.exists():
        return {"state": "never", "message": "还没有同步过睿贝数据"}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except ValueError:
        return {"state": "unknown", "message": "同步状态文件损坏"}


def status() -> dict:
    payload = _read_state()
    age = staleness_hours()
    payload.update({
        "enabled": SYNC_ENABLED,
        "intervalHours": INTERVAL_HOURS,
        "refreshDays": REFRESH_DAYS,
        "withAttachments": WITH_ATTACHMENTS,
        "running": is_running(),
        "ageHours": round(age, 2) if age is not None else None,
        "nextDueAt": _next_due(age),
    })
    return payload


def _next_due(age: float | None) -> str:
    """下一次计划同步时间（ISO 时间串）；没开调度就返回空。"""
    if not SYNC_ENABLED:
        return ""
    if age is None:
        return "尽快（尚未同步过）"
    remaining = INTERVAL_HOURS - age
    if remaining <= 0:
        return "尽快（已到同步时间）"
    when = datetime.now(timezone.utc) + timedelta(hours=remaining)
    return when.astimezone().isoformat(timespec="seconds")


def is_running() -> bool:
    global _PROC
    return _PROC is not None and _PROC.poll() is None


def staleness_hours() -> float | None:
    """上次同步完成距今多少小时；从未同步返回 None。"""
    state = _read_state()
    stamp = state.get("finishedAt") or state.get("updatedAt")
    if not stamp:
        return None
    try:
        when = datetime.fromisoformat(str(stamp))
    except ValueError:
        return None
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - when.astimezone(timezone.utc)).total_seconds() / 3600


def start(*, refresh_days: float = 7.0, lists_only: bool = False,
          with_attachments: bool | None = None) -> dict:
    """后台启动一次同步。已在跑时直接返回当前状态。"""
    global _PROC
    if with_attachments is None:
        with_attachments = WITH_ATTACHMENTS
    with _LOCK:
        if is_running():
            return {"state": "running", "message": "同步正在进行中"}
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        args = [sys.executable, str(SYNC_SCRIPT), "--refresh-days", str(refresh_days)]
        if lists_only:
            args.append("--lists-only")
        if not with_attachments:
            args.append("--no-attachments")
        log = LOG_PATH.open("a", encoding="utf-8")
        _PROC = subprocess.Popen(
            args,
            cwd=str(PROJECT_ROOT),
            stdout=log,
            stderr=subprocess.STDOUT,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    return {"state": "running", "message": "已开始同步睿贝数据"}


def ensure_fresh(*, max_age_hours: float = 12.0, refresh_days: float = 7.0) -> dict:
    """启动时调用：缓存太旧或从未同步才自动跑一次（不阻塞启动）。"""
    age = staleness_hours()
    if is_running():
        return {"state": "running", "message": "同步正在进行中"}
    if age is not None and age < max_age_hours:
        return {"state": "fresh", "message": f"{age:.1f} 小时前刚同步过"}
    started = start(refresh_days=refresh_days, lists_only=age is None)
    return {**started, "reason": "首次启动或数据已过期"}


def start_scheduler() -> dict:
    """启动定时同步线程（应用启动时调用一次，幂等）。

    每 `CHECK_SECONDS` 检查一次：距上次同步超过 `INTERVAL_HOURS` 就跑一轮，
    跑的时候带上出运明细补抓与采购附件补抓，保证本地缓存持续完整。
    """
    global _SCHEDULER
    if not SYNC_ENABLED:
        return {"state": "disabled", "message": "定时同步已关闭（ERP_SYNC_ENABLED=0）"}
    with _SCHEDULER_LOCK:
        if _SCHEDULER is not None and _SCHEDULER.is_alive():
            return {"state": "running", "message": "定时同步线程已在运行"}
        _SCHEDULER = threading.Thread(target=_scheduler_loop, name="erp-sync-scheduler", daemon=True)
        _SCHEDULER.start()
    # 启动时先按需同步一次（数据还新就跳过）
    first = ensure_fresh(max_age_hours=INTERVAL_HOURS, refresh_days=REFRESH_DAYS)
    return {"state": "started", "intervalHours": INTERVAL_HOURS, "first": first}


def _scheduler_loop() -> None:
    """后台调度：到点就跑一轮同步（子进程里跑，失败也不影响工作台）。"""
    while True:
        try:
            age = staleness_hours()
            if not is_running() and (age is None or age >= INTERVAL_HOURS):
                start(refresh_days=REFRESH_DAYS)
        except Exception:  # noqa: BLE001  调度线程不允许因为异常退出
            pass
        time.sleep(CHECK_SECONDS)
