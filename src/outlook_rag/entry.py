"""Lightweight executable entry: initialize MCP without native/API/COM work."""
from __future__ import annotations

import argparse
import json
import logging
import math
import os
from pathlib import Path
from typing import Any


def transport_options() -> tuple[str, float]:
    """Only transport settings: no database, model, credentials or Outlook probes."""
    cfg: dict[str, Any] = {}
    try:
        if path := os.environ.get("OUTLOOK_RAG_CONFIG"):
            cfg = json.loads(Path(path).expanduser().read_text(encoding="utf-8-sig"))
            if not isinstance(cfg, dict):
                raise ValueError("Configuration must be an object")
        profile = os.environ.get("OUTLOOK_RAG_MCP_TOOL_PROFILE", cfg.get("mcp_tool_profile", "standard"))
        raw = os.environ.get("OUTLOOK_RAG_MCP_TOOL_TIMEOUT_SECONDS", cfg.get("mcp_tool_timeout_seconds", 30))
        timeout = float(raw)
        if profile not in {"standard", "full"} or isinstance(raw, bool) or not math.isfinite(timeout) or not 1 <= timeout <= 300:
            raise ValueError("Invalid transport settings")
        return profile, timeout
    except (OSError, ValueError, TypeError, OverflowError) as exc:
        # Keep the handshake alive so a bad config can be corrected without
        # restarting Desktop. Operations still validate their full config.
        logging.getLogger(__name__).warning("Invalid MCP settings (%s); using standard profile and 30-second deadline.", type(exc).__name__)
        return "standard", 30


def serve(mcp: Any = None, adapters: dict | None = None) -> None:
    import anyio
    from mcp.server.fastmcp import FastMCP
    from outlook_rag import server, tool_runtime

    if mcp is None:
        mcp = FastMCP("outlook-rag", instructions=server.MAIL_SEARCH_INSTRUCTIONS)
        adapters = server.register_tools(mcp, None)
    profile, timeout = transport_options()
    assert adapters is not None
    runtime = tool_runtime.Pool()
    server.configure_tools(mcp, adapters, profile, runtime=runtime, timeout=timeout)
    async def run() -> None:
        try:
            await mcp.run_stdio_async()
        finally:
            await runtime.aclose()
    anyio.run(run)


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--config")
    args, rest = parser.parse_known_args()
    if not rest:
        if args.config:
            os.environ["OUTLOOK_RAG_CONFIG"] = args.config
        serve()
    else:
        from outlook_rag import app
        app.main()


if __name__ == "__main__":
    main()
