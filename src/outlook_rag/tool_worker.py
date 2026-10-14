"""JSON worker entry. Native libraries and COM stay outside MCP stdio."""
from __future__ import annotations

import json
import os
import sys
from typing import Any, Callable

MAX_FRAME_BYTES = 16 * 1024 * 1024


def redirect_native_handles() -> None:
    if sys.platform == 'win32':
        import ctypes
        from ctypes import wintypes
        import msvcrt
        setter = ctypes.WinDLL('kernel32', use_last_error=True).SetStdHandle
        setter.argtypes = [wintypes.DWORD, wintypes.HANDLE]
        setter.restype = wintypes.BOOL
        for identifier, fd in ((-10, 0), (-11, 1), (-12, 2)):
            if not setter(identifier & 0xffffffff, msvcrt.get_osfhandle(fd)):
                raise OSError(ctypes.get_last_error(), 'Cannot redirect helper standard handles')


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


def run(handler: Callable = execute) -> None:
    """Private JSON channel; both Python and native stdio go to the null device."""
    from outlook_rag import diagnostics
    from outlook_rag.process_guard import join_job

    join_job()

    incoming = os.fdopen(os.dup(sys.stdin.fileno()), "rb")
    outgoing = os.fdopen(os.dup(sys.stdout.fileno()), "wb", buffering=0)
    with open(os.devnull, "rb") as null_in, open(os.devnull, "wb") as null_out:
        os.dup2(null_in.fileno(), 0)
        os.dup2(null_out.fileno(), 1)
        os.dup2(null_out.fileno(), 2)
    redirect_native_handles()
    sys.stdin = open(os.devnull)
    sys.stdout = open(os.devnull, "w")
    sys.stderr = open(os.devnull, "w")
    with incoming, outgoing:
        while line := incoming.readline(MAX_FRAME_BYTES + 1):
            try:
                if len(line) > MAX_FRAME_BYTES or not line.endswith(b"\n"):
                    raise ValueError("Invalid helper frame")
                request = json.loads(line)
                response = handler(request["name"], request["arguments"], request["timeout"])
                encoded = json.dumps(response, ensure_ascii=True).encode() + b"\n"
                if len(encoded) > MAX_FRAME_BYTES:
                    raise ValueError("Helper response too large")
            except BaseException as exc:
                encoded = json.dumps({"ok": False, "error": diagnostics.exception_info(exc)}).encode() + b"\n"
            outgoing.write(encoded)


if __name__ == "__main__":
    run()
