"""Opt-in attachment capture and provenance, with isolated bounded parsing."""
from __future__ import annotations
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import time
from typing import Any

SUPPORTED = {".txt", ".csv", ".md", ".pdf", ".docx", ".xlsx", ".pptx"}


def run_worker(path: Path, env: dict, seconds: float) -> subprocess.CompletedProcess:
    import psutil
    command = [sys.executable, "-m", "outlook_rag.attachment_worker", str(path)]
    deadline = time.perf_counter() + seconds
    with subprocess.Popen(command, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8") as process:
        monitored = psutil.Process(process.pid)
        try:
            while True:
                remaining = deadline - time.perf_counter()
                if remaining <= 0:
                    raise subprocess.TimeoutExpired(command, seconds)
                try:
                    if monitored.memory_info().rss > 1024 * 1024 * 1024:
                        raise MemoryError("Attachment parser memory limit")
                except psutil.NoSuchProcess:
                    pass
                try:
                    output, errors = process.communicate(timeout=min(.05, remaining))
                    return subprocess.CompletedProcess(command, process.returncode, output, errors)
                except subprocess.TimeoutExpired:
                    pass
        except BaseException:
            process.kill()
            process.communicate()
            raise


def initialize(sql: sqlite3.Connection) -> None:
    sql.executescript("""
    CREATE TABLE IF NOT EXISTS chunk_sources(chunk_id TEXT PRIMARY KEY,attachment_name TEXT);
    CREATE TRIGGER IF NOT EXISTS remove_chunk_sources AFTER DELETE ON chunks
      BEGIN DELETE FROM chunk_sources WHERE chunk_id=OLD.id; END;
    CREATE TABLE IF NOT EXISTS attachment_reports(email_id TEXT PRIMARY KEY,report TEXT);
    CREATE TRIGGER IF NOT EXISTS remove_attachment_reports AFTER DELETE ON emails
      BEGIN DELETE FROM attachment_reports WHERE email_id=OLD.id; END;
    """)


def read(item: Any, cfg: dict) -> tuple[list[dict], list[dict]]:
    from .outlook import safe_attr
    found: list[dict] = []
    errors: list[dict] = []
    collection = safe_attr(item, "Attachments", None)
    count = safe_attr(collection, "Count", 0)
    if not isinstance(count, int):
        return found, errors
    env = {key: value for key, value in os.environ.items() if key.upper() in {"PATH", "SYSTEMROOT", "WINDIR", "TEMP", "TMP", "HOME", "USERPROFILE", "LANG", "LC_ALL", "PYTHONPATH"}}
    env["PYTHONIOENCODING"] = "utf-8"
    with tempfile.TemporaryDirectory(prefix="outlook-rag-attachment-", ignore_cleanup_errors=True) as folder:
        for position in range(1, min(count, 10) + 1):
            if cfg.get("_cancel_check", lambda: False)():
                break
            remaining = cfg.get("_embedding_deadline", time.perf_counter() + 10) - time.perf_counter()
            if remaining <= 0:
                errors.append(dict(error_type="TimeBudget"))
                break
            name = ""
            path: Path | None = None
            try:
                attachment = collection.Item(position)
                name = str(attachment.FileName)[:255]
                suffix = Path(name).suffix.casefold()
                if suffix not in SUPPORTED:
                    continue
                if not 0 < int(attachment.Size) <= 10 * 1024 * 1024:
                    raise ValueError("Attachment size limit")
                path = Path(folder) / (str(position) + suffix)
                attachment.SaveAsFile(str(path))
                # Avoid inheriting API credentials; output is bounded inside the worker.
                result = run_worker(path, env, min(10, remaining))
                if result.returncode:
                    raise ValueError("Attachment parser failed")
                text = json.loads(result.stdout)["text"]
                if not isinstance(text, str) or len(text) > 100000:
                    raise ValueError("Attachment text limit")
                if text.strip():
                    found.append(dict(name=name, text=text))
            except Exception as exc:
                errors.append(dict(attachment_name=name, error_type=type(exc).__name__))
            finally:
                if path is not None:
                    try:
                        path.unlink(missing_ok=True)
                    except OSError:
                        errors.append(dict(attachment_name=name, error_type="TemporaryFileCleanup"))
    if count > 10:
        errors.append(dict(error_type="AttachmentCountLimit", omitted=count - 10))
    return found, errors


def sources(sql: sqlite3.Connection, ids: list[str]) -> dict[str, str]:
    if not sql.execute("SELECT 1 FROM sqlite_master WHERE name='chunk_sources'").fetchone():
        return {}
    result: dict[str, str] = {}
    for offset in range(0, len(ids), 500):
        batch = ids[offset:offset + 500]
        if batch:
            result.update(sql.execute("SELECT chunk_id,attachment_name FROM chunk_sources WHERE chunk_id IN (" + ",".join("?" for _ in batch) + ")", batch))
    return result
