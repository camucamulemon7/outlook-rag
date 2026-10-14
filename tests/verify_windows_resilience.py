"""Opt-in Windows 11 checks with real COM/WMI and aggregate-only output.

Run with --real-outlook for a bounded real-mail read into a temporary index.
Embedding requests go only to a local synthetic HTTP endpoint. No mail values,
credentials or Outlook mutations are printed or sent to an external provider.
The outer watchdog kills only this run's test processes when a case stalls.
"""
import argparse
import asyncio
from contextlib import asynccontextmanager
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
import psutil


ROOT = Path(__file__).resolve().parents[1]


def clean_env():
    return {key: value for key, value in os.environ.items()
            if not key.startswith(("OUTLOOK_RAG_", "_OUTLOOK_RAG_"))}


def kill_python(processes):
    """Leave Outlook itself running; it is a shared application, not a fixture."""
    for process in reversed(processes):
        try:
            if process.name().lower() in {"python.exe", "pythonw.exe"}:
                process.kill()
        except psutil.NoSuchProcess:
            pass
    psutil.wait_procs(processes, timeout=3)


def bounded_case(case, seconds=90):
    directory = tempfile.TemporaryDirectory(prefix="outlook-rag-watchdog-")
    watchdog_root = Path(directory.name)
    process = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), "--case", case],
                               stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                               env=dict(clean_env(), PYTHONUTF8="1", OUTLOOK_RAG_TEST_DIR=str(watchdog_root)), cwd=ROOT)
    observed = {}
    deadline = time.monotonic() + seconds
    try:
        while process.poll() is None and time.monotonic() < deadline:
            try:
                for child in psutil.Process(process.pid).children(recursive=True):
                    observed[child.pid] = child
            except psutil.NoSuchProcess:
                pass
            for record in watchdog_root.rglob("*.job.json"):
                try:
                    pid = json.loads(record.read_text()).get("pid")
                    if pid and pid not in observed:
                        worker = psutil.Process(pid)
                        if str(watchdog_root) in " ".join(worker.cmdline()):
                            observed[pid] = worker
                except (OSError, ValueError, psutil.NoSuchProcess):
                    pass
            time.sleep(.05)
        if process.poll() is None:
            raise TimeoutError("Validation case exceeded its watchdog deadline")
        output, _ = process.communicate(timeout=3)
        if process.returncode:
            raise AssertionError("Validation case failed: " + case)
        result = json.loads(output)
        return result
    finally:
        kill_python(list(observed.values()))
        if process.poll() is None:
            process.kill()
        process.wait(timeout=5)
        directory.cleanup()


def wait_marker(path, seconds=12):
    deadline = time.monotonic() + seconds
    while not path.exists():
        if time.monotonic() >= deadline:
            raise TimeoutError("Fixture did not start")
        time.sleep(.03)
    return psutil.Process(int(path.read_text()))


def parent_case(folder, query, nested=False):
    from outlook_rag.tool_runtime import Pool
    host_job = None
    if nested:
        from outlook_rag.process_guard import create_job, attach_process
        _, host_job = create_job()
        attach_process(host_job, os.getpid())
    async def run():
        (folder / "parent").write_text(str(os.getpid()))
        pool = Pool([sys.executable, "-u", str(ROOT / "tests/mcp_fault_worker.py")])
        try:
            await pool.call("search_emails", {"query": query, "marker": str(folder / "native")}, 30)
        finally:
            await pool.aclose()
    anyio.run(run)
    _ = host_job


def hard_kill_case():
    checks = []
    for query, nested in (("hang", False), ("spawn-hang", False), ("spawn-hang", True)):
        with tempfile.TemporaryDirectory(prefix="outlook-rag-hardkill-") as directory:
            folder = Path(directory)
            process = subprocess.Popen([sys.executable, str(Path(__file__).resolve()),
                                       "--parent-case", str(folder), "--query", query, *(["--nested"] if nested else [])],
                                      env=clean_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            tracked = []
            try:
                owner = wait_marker(folder / "parent")
                tracked.append(owner)
                native = wait_marker(folder / "native")
                tracked.append(native)
                if query == "spawn-hang":
                    tracked.append(wait_marker(folder / "native.child"))
                tracked.extend(owner.children(recursive=True))
                started = time.perf_counter()
                owner.kill()
                _, alive = psutil.wait_procs(tracked, timeout=5)
                assert not alive, "A native helper/descendant survived transport termination"
                checks.append(dict(fault=query, nested_job=nested,
                                   reap_ms=round((time.perf_counter() - started) * 1000, 1), survivors=0))
            finally:
                kill_python(tracked)
                if process.poll() is None:
                    process.kill()
                process.wait(timeout=5)
    return dict(case="hard_kill", passed=True, checks=checks)


def payload(result):
    assert not result.isError, "MCP operation returned a tool error"
    return result.structuredContent or json.loads(result.content[0].text)


@asynccontextmanager
async def connect(config, environment):
    params = StdioServerParameters(command=sys.executable,
                                  args=["-m", "outlook_rag.entry", "--config", str(config)],
                                  env=environment)
    with anyio.fail_after(45):
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                yield session


def real_outlook_case():
    import pythoncom
    import win32com.client
    from outlook_rag import app
    from verify_outlook_metadata import probe

    pythoncom.CoInitialize()
    namespace = folder = targets = None
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        report, targets = app.discover_sources(namespace, {"exclude_system_folders": True})
        assert not report["errors"], "Outlook discovery failed"
        path, folder = next((path, folder) for path, folder in targets if folder.Items.Count > 65)
        metadata = probe(folder, 65, namespace)
        stores = len(report["stores"])
    finally:
        targets = folder = namespace = None
        pythoncom.CoUninitialize()

    class Handler(BaseHTTPRequestHandler):
        calls = 0
        blocked = threading.Event()
        release = threading.Event()
        pause = False
        def do_POST(self):
            value = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            rows = []
            for index, text in enumerate(value["input"]):
                digest = hashlib.sha256(text.encode()).digest()
                rows.append(dict(index=index, embedding=[(byte + 1) / 256 for byte in digest[:4]]))
            data = json.dumps(dict(data=rows)).encode()
            type(self).calls += 1
            if type(self).pause:
                type(self).blocked.set()
                type(self).release.wait(timeout=10)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        def log_message(self, *args):
            pass

    api = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=api.serve_forever, daemon=True)
    thread.start()
    try:
        with tempfile.TemporaryDirectory(prefix="outlook-rag-real-com-", dir=os.environ.get("OUTLOOK_RAG_TEST_DIR")) as directory:
            root = Path(directory)
            index = root / "index"
            config = root / "config.json"
            cfg = dict(model="windows-real-com-fixture", dimensions=4, data_dir=str(index),
                       embedding_url=f"http://127.0.0.1:{api.server_port}/embeddings", warmup_on_start=False,
                       mcp_tool_profile="full", mcp_tool_timeout_seconds=30, sync_max_emails=3,
                       sync_max_total_emails=3, sync_max_scanned=8, sync_batch_emails=1, sync_max_seconds=4)
            config.write_text(json.dumps(cfg), encoding="utf-8")
            env = dict(clean_env(), OUTLOOK_RAG_API_KEY="local-fixture-only")
            async def exercise():
                async with connect(config, env) as session:
                    before = payload(await session.call_tool("search_emails", {"query": "test"}))
                    assert before["state"] == "not_initialized" and not index.exists()
                    sources = payload(await session.call_tool("list_outlook_sources", {}))
                    assert len(sources["stores"]) == stores and not sources["errors"]
                    foreground = payload(await session.call_tool("sync_emails", {"folders": [path],
                                         "max_total_emails": 3, "max_scanned": 8, "max_seconds": 4}))
                    assert foreground["indexed_total"] > 0
                    count = payload(await session.call_tool("index_status", {}))["emails"]
                    assert 0 < count <= 3
                    hits = payload(await session.call_tool("search_emails", {"query": "", "mode": "metadata"}))
                    assert hits["count"] == count
                    # Real WMI + authenticated named-pipe launch.
                    job = payload(await session.call_tool("sync_index", {"action": "start", "folders": ["inbox"], "since_days": 0, "confirmed": True}))
                    job_id = job["job_id"]
                    for _ in range(100):
                        job = payload(await session.call_tool("sync_index", {"action": "status", "job_id": job_id}))
                        if job["state"] not in {"starting", "running", "waiting"}:
                            break
                        await asyncio.sleep(.1)
                    assert job["state"] == "completed", "Real WMI worker did not finish"
                    await session.send_ping()
                    Handler.pause = True
                    job = payload(await session.call_tool("sync_index", {"action": "start", "folders": [path], "since_days": 0, "confirmed": True}))
                    job_id = job["job_id"]
                    for _ in range(120):
                        job = payload(await session.call_tool("sync_index", {"action": "status", "job_id": job_id}))
                        if Handler.blocked.is_set() and job["state"] == "running":
                            break
                        await asyncio.sleep(.05)
                    assert Handler.blocked.is_set() and job["state"] == "running", "WMI job did not reach the local API barrier"
                    worker = psutil.Process(job["pid"])
                assert worker.is_running(), "Detached WMI job died on MCP EOF"
                async with connect(config, env) as session:
                    state = payload(await session.call_tool("sync_index", {"action": "status", "job_id": job_id}))
                    assert state["state"] in {"running", "waiting"}
                    cancelled = payload(await session.call_tool("sync_index", {"action": "cancel", "job_id": job_id}))
                    assert cancelled["cancel_requested"]
                    Handler.release.set()
                    for _ in range(120):
                        state = payload(await session.call_tool("sync_index", {"action": "status", "job_id": job_id}))
                        if state["state"] not in {"starting", "running", "waiting"}:
                            break
                        await asyncio.sleep(.05)
                    assert state["state"] == "cancelled", "WMI job did not honor cancellation"
                    await session.send_ping()
                    current = payload(await session.call_tool("index_status", {}))["emails"]
                worker.wait(timeout=5)
                return dict(indexed_real_mails=current, connected_stores=stores, metadata_rows=metadata["rows_checked"],
                                wmi_worker_state=state["state"], synthetic_local_api_calls=Handler.calls,
                                wmi_survived_mcp_eof=True, wmi_cancelled=state["state"] == "cancelled",
                                writes_to_outlook=False, external_api_used=False)
            result = asyncio.run(exercise())
            # A second session validates release of COM apartments, DB handles
            # and idle helpers on transport EOF.
            async def restart():
                async with connect(config, env) as session:
                    assert payload(await session.call_tool("index_status", {}))["emails"] == result["indexed_real_mails"]
            asyncio.run(restart())
            return dict(case="real_outlook", passed=True, **result)
    finally:
        api.shutdown()
        api.server_close()
        thread.join(timeout=3)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--real-outlook", action="store_true")
    parser.add_argument("--case", choices=["hard_kill", "real_outlook"])
    parser.add_argument("--parent-case", type=Path)
    parser.add_argument("--query", default="hang")
    parser.add_argument("--nested", action="store_true")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.error("This opt-in probe requires Windows")
    if args.parent_case:
        parent_case(args.parent_case, args.query, args.nested)
        return
    if args.case:
        print(json.dumps(hard_kill_case() if args.case == "hard_kill" else real_outlook_case()))
        return
    cases = ["hard_kill"] + (["real_outlook"] if args.real_outlook else [])
    result = dict(platform=sys.platform, python=sys.version.split()[0], cases=[bounded_case(case) for case in cases])
    if args.report:
        args.report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result))


if __name__ == "__main__":
    main()
