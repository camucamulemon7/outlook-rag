"""Owned, bounded subprocesses for native MCP operations, including startup."""
from __future__ import annotations

from dataclasses import dataclass
import json
import subprocess
import sys
import time
from typing import Any

import anyio
from anyio.abc import Process
from anyio.streams.buffered import BufferedByteReceiveStream
from mcp.server.fastmcp.exceptions import ToolError

from outlook_rag.tool_worker import MAX_FRAME_BYTES


@dataclass(eq=False)
class Helper:
    process: Process
    output: BufferedByteReceiveStream
    used_at: float = 0


class Pool:
    """Own every child from creation until reaping; never abandon startup."""
    def __init__(self, command: list[str] | None = None) -> None:
        self.command = command or [sys.executable, "-u", "-m", "outlook_rag.tool_worker"]
        self.limiter = anyio.CapacityLimiter(4)
        self.helpers: set[Helper] = set()
        self.idle: list[Helper] = []

    async def stop(self, helper: Helper) -> None:
        self.helpers.discard(helper)
        if helper in self.idle:
            self.idle.remove(helper)
        with anyio.CancelScope(shield=True):
            try:
                if helper.process.returncode is None:
                    helper.process.kill()
            except ProcessLookupError:
                pass
            await helper.process.aclose()

    async def aclose(self) -> None:
        for helper in list(self.helpers):
            await self.stop(helper)

    async def call(self, name: str, arguments: dict[str, Any], timeout: float) -> dict:
        try:
            with anyio.fail_after(timeout):
                async with self.limiter:
                    helper = None
                    try:
                        while self.idle:
                            candidate = self.idle.pop()
                            if candidate.process.returncode is None and time.monotonic() - candidate.used_at < 300:
                                helper = candidate
                                break
                            await self.stop(candidate)
                        if helper is None:
                            # Track ownership before delivering cancellation after
                            # OS spawn, including before the first protocol read.
                            with anyio.CancelScope(shield=True):
                                process = await anyio.open_process(self.command, stdin=subprocess.PIPE,
                                                                  stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
                                assert process.stdout is not None
                                helper = Helper(process, BufferedByteReceiveStream(process.stdout))
                                self.helpers.add(helper)
                        assert helper is not None
                        request = json.dumps({"name": name, "arguments": arguments, "timeout": timeout}).encode() + b"\n"
                        if len(request) > MAX_FRAME_BYTES:
                            raise ToolError("MCP operation exceeds the helper message size limit.")
                        assert helper.process.stdin is not None
                        await helper.process.stdin.send(request)
                        frame = await helper.output.receive_until(b"\n", MAX_FRAME_BYTES)
                        response = json.loads(frame)
                        if (not isinstance(response, dict) or type(response.get('ok')) is not bool
                                or not isinstance(response.get('result' if response['ok'] else 'error'), dict)):
                            raise ValueError("Invalid helper response")
                    except BaseException:
                        if helper is not None:
                            await self.stop(helper)
                        raise
                    else:
                        helper.used_at = time.monotonic()
                        self.idle.append(helper)
            if not response["ok"]:
                raise ToolError(json.dumps(response["error"], ensure_ascii=False))
            return response["result"]
        except TimeoutError as exc:
            raise ToolError(f"{name} exceeded the {timeout:g}-second MCP deadline. Its helper was stopped; this MCP connection remains usable. Check sync_index(action='status') before restarting indexing.") from exc
        except (anyio.EndOfStream, anyio.BrokenResourceError, anyio.ClosedResourceError,
                anyio.IncompleteRead, anyio.DelimiterNotFound, OSError, ValueError) as exc:
            raise ToolError(f"{name} helper stopped unexpectedly ({type(exc).__name__}). This MCP connection remains usable; retry the operation or inspect index_status.") from exc
