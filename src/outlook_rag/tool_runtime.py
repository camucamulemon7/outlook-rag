"""Bounded, cancellable process isolation for synchronous MCP operations."""
from __future__ import annotations

import json
from typing import Any, Callable

import anyio
from anyio import to_process
from mcp.server.fastmcp.exceptions import ToolError

from outlook_rag.tool_worker import execute

# Native/COM failures must stay confined to helpers. Expose only JSON-like tool
# results, never a database handle, client, exception object or COM proxy.


async def call(name: str, arguments: dict[str, Any], timeout: float,
               limiter: anyio.CapacityLimiter, worker: Callable = execute) -> dict:
    try:
        with anyio.fail_after(timeout):
            response = await to_process.run_sync(worker, name, arguments, timeout,
                                                 cancellable=True, limiter=limiter)
    except TimeoutError as exc:
        raise ToolError(f"{name} exceeded the {timeout:g}-second MCP deadline. Its helper was stopped; this MCP connection remains usable. Check sync_index(action='status') before restarting indexing.") from exc
    except (anyio.BrokenWorkerProcess, anyio.ClosedResourceError, OSError) as exc:
        raise ToolError(f"{name} helper stopped unexpectedly ({type(exc).__name__}). This MCP connection remains usable; retry the operation or inspect index_status.") from exc
    if not response["ok"]:
        raise ToolError(json.dumps(response["error"], ensure_ascii=False))
    return response["result"]
