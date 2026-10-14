"""Pickle-safe worker entry. Native libraries and COM stay outside MCP stdio."""
from __future__ import annotations

from typing import Any


def execute(name: str, arguments: dict[str, Any], timeout: float) -> dict:
    from outlook_rag import diagnostics

    try:
        from outlook_rag import app

        # Keep foreground sync inside its cooperative budget before the hard
        # process deadline. Bulk work belongs in the detached sync_index job.
        if name in {"sync_emails", "retry_failed_emails"}:
            arguments = dict(arguments)
            arguments["max_seconds"] = min(arguments.get("max_seconds") or 30, max(1, timeout - 5))
        return {"ok": True, "result": app._tools[name](**arguments)}
    except BaseException as exc:
        # Never pickle COM exceptions or propagate SystemExit/KeyboardInterrupt
        # into the transport. Raw messages can contain credentials or mail.
        return {"ok": False, "error": diagnostics.exception_info(exc)}
