"""Bounded worker diagnostics that never persist raw exception text or output."""

from __future__ import annotations

from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import time
import traceback
from typing import Any

MAX_LOG_BYTES = 1024 * 1024


def log_path(cfg: dict, job_id: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise ValueError("Invalid job_id")
    return Path(cfg["data_dir"]) / "jobs" / (job_id + ".log")


def exception_detail(exc: BaseException) -> str:
    """Safe actionable codes; messages may contain keys, mail or provider bodies."""
    if type(exc).__name__ == "HTTPStatusError":
        return f"Embedding provider returned HTTP {exc.response.status_code}."  # type: ignore[attr-defined]
    if (
        isinstance(exc, ModuleNotFoundError)
        and exc.name
        and re.fullmatch(r"[\w.]+", exc.name)
    ):
        return "Required module unavailable: " + exc.name + "."
    for field, label in (
        ("wmi_return_code", "WMI process creation code"),
        ("hresult", "Outlook COM HRESULT"),
        ("winerror", "Windows error"),
        ("errno", "OS error"),
    ):
        value = getattr(exc, field, None)
        if isinstance(value, int):
            return f"{label} {value}."
    name = getattr(exc, "sqlite_errorname", None)
    if isinstance(name, str) and re.fullmatch(r"SQLITE_[A-Z_]+", name):
        return "Local database error: " + name + "."
    return f"Worker failed ({type(exc).__name__}); inspect diagnostic stack frames."


def exception_info(exc: BaseException) -> dict:
    frames = [
        {
            "file": Path(frame.f_code.co_filename).name,
            "line": line,
            "function": frame.f_code.co_name,
        }
        for frame, line in traceback.walk_tb(exc.__traceback__)
    ]
    return {
        "error_type": type(exc).__name__,
        "error_detail": exception_detail(exc),
        "stack": frames,
    }


def append(cfg: dict, job_id: str, event: str, **fields: Any) -> None:
    """Callers pass only counters/state or sanitized exception_info, never mail."""
    path = log_path(cfg, job_id)
    value = {"at": datetime.now(timezone.utc).isoformat(), "event": event, **fields}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and path.stat().st_size > MAX_LOG_BYTES:
            with path.open("rb") as stream:
                stream.seek(-MAX_LOG_BYTES // 2, os.SEEK_END)
                tail = stream.read().split(b"\n", 1)[-1]
            path.write_bytes(tail)
        with path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(value, ensure_ascii=False) + "\n")
    except OSError:
        # A diagnostic failure must not turn a completed sync into a failure.
        pass


def recent(cfg: dict, job_id: str) -> list[str]:
    path = log_path(cfg, job_id)
    try:
        with path.open("rb") as stream:
            stream.seek(max(0, path.stat().st_size - 65536))
            return stream.read().decode("utf-8", "replace").splitlines()[-20:]
    except OSError:
        return []


def prune(cfg: dict, active_job_id: str | None = None, days: int = 14) -> int:
    cutoff = time.time() - days * 86400
    removed = 0
    for path in (Path(cfg["data_dir"]) / "jobs").glob("*.log"):
        if (
            re.fullmatch(r"[0-9a-f]{32}", path.stem)
            and path.stem != active_job_id
            and path.stat().st_mtime < cutoff
        ):
            path.unlink(missing_ok=True)
            removed += 1
    return removed
