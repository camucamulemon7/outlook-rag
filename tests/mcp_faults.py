"""Importable subprocess faults, never enabled through production settings."""
import os
from pathlib import Path
import time


def execute(name, arguments, timeout):
    query = arguments.get("query", "ok")
    if path := arguments.get("marker"):
        marker = Path(path)
        temporary = marker.with_name(marker.name + '.tmp')
        temporary.write_text(str(os.getpid()), encoding="utf-8")
        temporary.replace(marker)
    if query == "hang":
        time.sleep(60)
    if query == "exit":
        os._exit(7)
    if query == "native-output":
        os.write(1, b"native output is not an MCP message\n")
    if query == "print":
        print("This output must never reach MCP stdio", flush=True)
    if query in {"system-exit", "com-error", "attribute-error"}:
        from unittest.mock import patch
        from outlook_rag import app, tool_worker
        def fail(**kwargs):
            if query == "system-exit":
                raise SystemExit("private secret")
            if query == "attribute-error":
                raise AttributeError("private secret")
            if os.name == "nt":
                import pywintypes
                raise pywintypes.com_error(-2147418111, "private secret", None, None)
            exc = RuntimeError("private secret")
            exc.hresult = -2147418111
            raise exc
        with patch.dict(app._tools, {name: fail}):
            return tool_worker.execute(name, arguments, timeout)
    return {"ok": True, "result": {"pid": os.getpid(), "query": query}}
