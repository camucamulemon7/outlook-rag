"""Exercise actual stdio handshakes, fault recovery, cancellation and EOF."""
import asyncio
from contextlib import asynccontextmanager
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
import psutil


ROOT = Path(__file__).resolve().parents[1]


@asynccontextmanager
async def connect(args, env=None):
    params = StdioServerParameters(command=sys.executable, args=args, env=env)
    with anyio.fail_after(20):
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                yield session


def payload(result):
    return result.structuredContent or json.loads(result.content[0].text)


class StdioResilience(unittest.TestCase):
    def test_cold_start_without_dimensions_key_api_or_outlook(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                root = Path(folder)
                config = root / "config.json"
                config.write_text(json.dumps({"model": "fixture", "data_dir": str(root / "index"),
                                              "embedding_url": "http://127.0.0.1:1/embeddings"}), encoding="utf-8")
                env = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RAG_")}
                for module in ("outlook_rag.entry", "outlook_rag.app"):
                    async with connect(["-m", module, "--config", str(config)], env) as session:
                        self.assertEqual(len((await session.list_tools()).tools), 6)
                        for mode in ("hybrid", "metadata", "thread"):
                            result = await session.call_tool("search_emails", {"query": "test", "mode": mode})
                            self.assertFalse(result.isError, result)
                            self.assertEqual(payload(result)["state"], "not_initialized")
                        for name in ("index_status", "sync_index"):
                            result = await session.call_tool(name, {})
                            self.assertFalse(result.isError, result)
                        missing = await session.call_tool("get_indexed_mail", {"mail_id": "missing"})
                        self.assertTrue(missing.isError)
                        await session.send_ping()
                        self.assertFalse((await session.call_tool("index_status", {})).isError)
                self.assertFalse((root / "index").exists())
        asyncio.run(exercise())

    def test_timeout_crash_com_and_output_do_not_disconnect_stdio(self):
        async def exercise():
            async with connect([str(ROOT / "tests" / "mcp_fault_server.py")]) as session:
                for query in ("print", "exit", "native-output", "system-exit", "com-error", "attribute-error", "hang"):
                    started = time.perf_counter()
                    result = await session.call_tool("search_emails", {"query": query})
                    self.assertEqual(result.isError, query != "print", result)
                    self.assertNotIn("private secret", str(result))
                    self.assertLess(time.perf_counter() - started, 6)
                    await session.send_ping()
                    self.assertFalse((await session.call_tool("index_status", {})).isError)
        asyncio.run(exercise())

    def test_cancel_and_eof_reap_helpers_and_status_stays_responsive(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                marker = Path(folder) / "pid"
                async with connect([str(ROOT / "tests" / "mcp_fault_server.py")]) as session:
                    task = asyncio.create_task(session.call_tool("search_emails", {"query": "hang", "marker": str(marker)}))
                    while not marker.exists():
                        await asyncio.sleep(.02)
                    cancelled_pid = int(marker.read_text())
                    started = time.perf_counter()
                    await session.send_ping()
                    status = await session.call_tool("index_status", {})
                    self.assertFalse(status.isError, status)
                    self.assertLess(time.perf_counter() - started, 1.8)
                    task.cancel()
                    with self.assertRaises(asyncio.CancelledError):
                        await task
                    for _ in range(100):
                        if not psutil.pid_exists(cancelled_pid):
                            break
                        await asyncio.sleep(.02)
                    self.assertFalse(psutil.pid_exists(cancelled_pid))
                    alive_pid = payload(await session.call_tool("index_status", {}))["pid"]
                for _ in range(100):
                    if not psutil.pid_exists(alive_pid):
                        break
                    await asyncio.sleep(.02)
                self.assertFalse(psutil.pid_exists(alive_pid))
        asyncio.run(exercise())

    def test_transport_import_does_not_load_native_engines(self):
        code = "from outlook_rag import entry; from mcp.server.fastmcp import FastMCP; from outlook_rag import server, tool_runtime; import sys; server.register_tools(FastMCP('test'),None); assert not any(m in sys.modules for m in ('outlook_rag.app','lancedb','pyarrow','pythoncom','win32com.client','win32crypt'))"
        result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_bad_configuration_can_be_fixed_in_the_same_session(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                config = Path(folder) / 'config.json'
                config.write_text('{broken private secret', encoding='utf-8')
                env = {k: v for k, v in os.environ.items() if not k.startswith('OUTLOOK_RAG_')}
                async with connect(['-m', 'outlook_rag.entry', '--config', str(config)], env) as session:
                    self.assertEqual(len((await session.list_tools()).tools), 6)
                    result = await session.call_tool('index_status', {})
                    self.assertTrue(result.isError)
                    self.assertNotIn('private secret', str(result))
                    await session.send_ping()
                    config.write_text(json.dumps({'model': 'fixture', 'data_dir': str(Path(folder) / 'index')}), encoding='utf-8')
                    result = await session.call_tool('index_status', {})
                    self.assertFalse(result.isError, result)
                    self.assertEqual(payload(result)['state'], 'not_initialized')
        asyncio.run(exercise())

    def test_saturated_pool_has_bounded_wait_and_recovers(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                async with connect([str(ROOT / 'tests' / 'mcp_fault_server.py')]) as session:
                    requests = [asyncio.create_task(session.call_tool('search_emails', {'query': 'hang', 'marker': str(Path(folder) / str(i))})) for i in range(6)]
                    await asyncio.sleep(.3)
                    await session.send_ping()
                    results = await asyncio.gather(*requests)
                    self.assertTrue(all(result.isError for result in results))
                    pids = {int(path.read_text()) for path in Path(folder).iterdir() if path.name.isdigit()}
                    self.assertLessEqual(len(pids), 4)
                    self.assertTrue(all(not psutil.pid_exists(pid) for pid in pids))
                    self.assertFalse((await session.call_tool('index_status', {})).isError)
        asyncio.run(exercise())

    def test_eof_during_blocked_call_reaps_worker_before_deadline(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                marker = Path(folder) / "pid"
                started = time.perf_counter()
                async with connect([str(ROOT / "tests" / "mcp_fault_server.py"), "30"]) as session:
                    task = asyncio.create_task(session.call_tool("search_emails", {"query": "hang", "marker": str(marker)}))
                    while not marker.exists():
                        await asyncio.sleep(.02)
                    pid = int(marker.read_text())
                task.cancel()
                await asyncio.gather(task, return_exceptions=True)
                for _ in range(100):
                    if not psutil.pid_exists(pid):
                        break
                    await asyncio.sleep(.02)
                self.assertFalse(psutil.pid_exists(pid))
                self.assertLess(time.perf_counter() - started, 8)
        asyncio.run(exercise())
