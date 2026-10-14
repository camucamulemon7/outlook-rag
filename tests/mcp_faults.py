"""Importable subprocess faults, never enabled through production settings."""
import os
from pathlib import Path
import time
import subprocess
import sys


def execute(name, arguments, timeout):
    query = arguments.get("query", "ok")
    if path := arguments.get("marker"):
        marker = Path(path)
        temporary = marker.with_name(marker.name + '.tmp')
        temporary.write_text(str(os.getpid()), encoding="utf-8")
        temporary.replace(marker)
    if query == "spawn-hang":
        child_marker = str(arguments["marker"]) + ".child"
        subprocess.Popen([sys.executable, "-c",
                          "import os,sys,time; from pathlib import Path; "
                          "Path(sys.argv[1]).write_text(str(os.getpid())); time.sleep(60)", child_marker])
        time.sleep(60)
    if query == "hang":
        time.sleep(60)
    if query == "exit":
        os._exit(7)
    if query == "native-output":
        os.write(1, b"native output is not an MCP message\n")
        if os.name == 'nt':
            import ctypes
            from ctypes import wintypes
            kernel = ctypes.WinDLL('kernel32', use_last_error=True)
            kernel.GetStdHandle.argtypes = [wintypes.DWORD]
            kernel.GetStdHandle.restype = wintypes.HANDLE
            kernel.WriteFile.argtypes = [wintypes.HANDLE, wintypes.LPCVOID, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID]
            written = wintypes.DWORD()
            data = b'Windows native output\n'
            kernel.WriteFile(kernel.GetStdHandle(-11 & 0xffffffff), data, len(data), ctypes.byref(written), None)
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
