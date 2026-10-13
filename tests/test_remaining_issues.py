"""Operational regressions raised in issues 16, 18, 19, 21, 23 and 25."""
import asyncio
import base64
import hashlib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
from array import array

import httpx
from filelock import FileLock
from mcp.server.fastmcp import FastMCP
from mcp.shared.memory import create_connected_server_and_client_session
from outlook_rag import app, embeddings, jobs, outlook, paths, server, settings


class OperationalRegressions(unittest.TestCase):
    def test_settings_fail_before_probe_and_accept_documented_booleans(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "config.json"
            env = dict(OUTLOOK_RAG_CONFIG=str(path), OUTLOOK_RAG_MODEL="fixture")
            for field, value in (("sync_max_seconds", "secret-invalid"), ("excluded_folders", [1]),
                                 ("dimensions", True), ("semantic_weight", "NaN")):
                path.write_text(json.dumps({field: value}))
                with patch.dict(os.environ, env, clear=True), patch.object(app, "detect_dimensions") as probe:
                    with self.assertRaisesRegex(ValueError, field) as raised:
                        app.config()
                    self.assertNotIn("secret-invalid", str(raised.exception))
                    probe.assert_not_called()
            for value in (True, "True", "YES", 1):
                self.assertTrue(settings.boolean(value))
            self.assertFalse(settings.boolean("False"))

    def test_transport_retry_base64_and_capability_fallback(self):
        cfg = dict(model="fixture-retries", dimensions=4, embedding_url="http://fixture.invalid")
        vector = base64.b64encode(array("f", [3., 4., 0., 0.]).tobytes()).decode()
        response = httpx.Response(200, json={"data": [{"index": 0, "embedding": vector}]}, request=httpx.Request("POST", cfg["embedding_url"]))
        client = Mock()
        client.post.side_effect = [httpx.ConnectError("fixture"), response]
        with patch.object(app, "api_key", return_value="fixture"), patch.object(embeddings, "pause") as pause:
            self.assertEqual(app.embed(["mail"], cfg, client=client), [[.6, .8, 0., 0.]])
            self.assertEqual(client.post.call_count, 2)
            pause.assert_called_once()
        rejected = httpx.Response(422, request=httpx.Request("POST", cfg["embedding_url"]))
        client.post.side_effect = [rejected, response, response]
        cfg["embedding_encoding_format"] = "base64"
        embeddings._ENCODING.clear()
        with patch.object(app, "api_key", return_value="fixture"):
            app.embed(["a"], cfg, client=client)
            app.embed(["b"], cfg, client=client)
        self.assertEqual([c.kwargs["json"].get("encoding_format") for c in client.post.call_args_list[-3:]], ["base64", "float", None])

    def test_retry_after_budget_and_cancellation(self):
        cfg = dict(embedding_url="http://fixture.invalid", _embedding_deadline=time.perf_counter() + .06)
        response = httpx.Response(429, headers={"Retry-After": "30"})
        client = Mock(post=Mock(return_value=response))
        started = time.perf_counter()
        with self.assertRaises(httpx.ReadTimeout):
            embeddings.request(client, {}, cfg, "fixture")
        self.assertLess(time.perf_counter() - started, .5)
        self.assertEqual(client.post.call_count, 1)
        self.assertEqual(embeddings.retry_delay(response, 0), 30)
        with self.assertRaises(embeddings.SyncCancelled):
            embeddings.pause(dict(_cancel_check=lambda: True), 30)
        with self.assertRaises(ValueError):
            embeddings.decode_vector("!!invalid!!")

    def test_onedrive_notice_and_local_fallback(self):
        with tempfile.TemporaryDirectory() as root, patch.dict(os.environ, {"OneDrive": root, "LOCALAPPDATA": str(Path(root).parent / "local-fixture")}, clear=True):
            self.assertTrue(paths.warnings({"data_dir": str(Path(root) / "mail" / "db")}))
            self.assertFalse(paths.warnings({"data_dir": str(paths.fallback_root())}))

    def test_corrupt_job_is_skipped_and_terminal_files_are_bounded(self):
        with tempfile.TemporaryDirectory() as root:
            cfg = dict(data_dir=root)
            folder = jobs.root(cfg)
            (folder / ("e" * 32 + ".job.json")).write_text("{broken")
            for n in range(30):
                jid = f"{n:032x}"
                jobs.write(folder / (jid + ".job.json"), dict(job_id=jid, state="completed", created_at=n, updated_at=time.time()))
                for suffix in (".config.json", ".cancel"):
                    (folder / (jid + suffix)).touch()
            self.assertEqual(jobs.load(cfg)[1]["job_id"], f"{29:032x}")
            self.assertEqual(jobs.prune_records(cfg), 10)
            self.assertEqual(len(list(folder.glob("*.config.json"))), 20)
            self.assertEqual(len(list(folder.glob("*.cancel"))), 20)

    def test_short_status_lock_does_not_fail_worker(self):
        with tempfile.TemporaryDirectory() as root:
            cfg = dict(data_dir=root)
            jid = "a" * 32
            path, _ = jobs.paths(cfg, jid)
            jobs.write(path, dict(job_id=jid, state="starting", created_at=time.time(), updated_at=time.time(),
                                 known=[], complete=[], options={}, cycles=0, indexed_total=0, scanned_total=0, retried_total=0))
            lock = FileLock(jobs.root(cfg) / "worker.lock", timeout=0)
            lock.acquire()
            with patch.object(app, "api_key", return_value="fixture"), patch.object(app, "sync", return_value=dict(folders=[], indexed_total=0)), ThreadPoolExecutor() as pool:
                future = pool.submit(jobs.run, jid, cfg)
                time.sleep(.1)
                lock.release()
                future.result(timeout=5)
            self.assertEqual(jobs.load(cfg, jid)[1]["state"], "completed")

    def test_changed_key_file_invalidates_in_memory_decryption_cache(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "key.dpapi"
            path.write_bytes(b"first")
            cfg = dict(key_file=str(path))
            decrypt = Mock(side_effect=lambda value, *args: (None, b"decoded-" + value))
            settings._KEY_CACHE.clear()
            with patch.dict(os.environ, {}, clear=True), patch.dict("sys.modules", {"win32crypt": SimpleNamespace(CryptUnprotectData=decrypt)}):
                self.assertEqual(app.api_key(cfg), "decoded-first")
                self.assertEqual(app.api_key(cfg), "decoded-first")
                self.assertEqual(decrypt.call_count, 1)
                path.write_bytes(b"second-value")
                self.assertEqual(app.api_key(cfg), "decoded-second-value")
                self.assertEqual(decrypt.call_count, 2)
                with patch.dict(os.environ, {"OUTLOOK_RAG_API_KEY": "environment-value"}):
                    self.assertEqual(app.api_key(cfg), "environment-value")
                self.assertEqual(decrypt.call_count, 2)

    def test_source_date_update_publishes_vectors_without_embeddings(self):
        from outlook_rag import storage
        with tempfile.TemporaryDirectory() as root:
            cfg = dict(model="fixture", dimensions=4, data_dir=root, embedding_url="http://fixture.invalid")
            sql, table = app.open_store(cfg)
            try:
                mail = dict(id=hashlib.sha256(b"date-boundary").hexdigest(), subject="Date", body="Original text", received="2026-10-12T23:30:00+00:00", modified="2026-10-13T00:00:00+00:00")
                with patch.object(app, "embed", return_value=[[1., 0., 0., 0.]]):
                    app.upsert_mail(mail, sql, table, cfg)
                with patch.object(app, "embed", side_effect=AssertionError("Dates must not embed")):
                    storage.update_dates(sql, table, mail["id"], "2026-10-13T08:30:00+09:00", "2026-10-13T09:00:00+09:00", cfg)
                self.assertEqual(sql.execute("SELECT substr(received,1,10) FROM emails").fetchone()[0], "2026-10-13")
                self.assertEqual(table.search().select(["received_date"]).limit(1).to_list()[0]["received_date"], "2026-10-13")
                self.assertEqual(sql.execute("SELECT value FROM state WHERE key='vector_version'").fetchone()[0], str(table.version))
            finally:
                sql.close()

    def test_table_namespace_and_item_properties_preserve_date_semantics(self):
        value = datetime(2026, 10, 13, 9, 30, tzinfo=timezone.utc)
        accessor = Mock()
        accessor.GetProperty.return_value = value
        item = SimpleNamespace(PropertyAccessor=accessor, ReceivedTime=value)
        actual = datetime.fromisoformat(outlook.item_date(item, "ReceivedTime", outlook.RECEIVED))
        self.assertEqual(actual.astimezone(timezone.utc), value)
        accessor.GetProperty.assert_called_once_with(outlook.RECEIVED)
        table = Mock()
        table.EndOfTable = True
        list(outlook.mail_metadata(SimpleNamespace(GetTable=lambda _: table), []))
        self.assertEqual([c.args[0] for c in table.Columns.Add.call_args_list], ["EntryID", "MessageClass", outlook.RECEIVED, outlook.MODIFIED])

    def test_extended_filters_deduplicate_within_scope_and_recover_legacy_rows(self):
        with tempfile.TemporaryDirectory() as root:
            cfg = dict(model="fixture", dimensions=4, data_dir=root, embedding_url="http://fixture.invalid")
            sql, table = app.open_store(cfg)
            try:
                mails = [dict(id=hashlib.sha256(str(n).encode()).hexdigest(), subject="PCIe", body="timing approval", sender="address@example.invalid",
                              sender_name="田中 太郎", cc="工程チーム", internet_message_id="<copy@example.invalid>", has_attachments=True,
                              received=f"2026-10-{13 + n:02d}T00:00:00+00:00", folder=f"Archive/{n}") for n in range(2)]
                with patch.object(app, "embed", side_effect=lambda texts, *a, **k: [[1., 0., 0., 0.] for _ in texts]):
                    app.upsert_mails(mails, sql, table, cfg)
                with patch.object(app, "local_config", return_value=cfg), patch.object(app, "embed", side_effect=AssertionError("Metadata must be offline")):
                    result = app.search("", mode="metadata", sender="田中", recipient="工程", folder="Archive/*", has_attachments=True)
                    self.assertEqual(result["count"], 1)
                    self.assertEqual(result["items"][0]["mail_id"], mails[1]["id"])
                    self.assertEqual(app.search("", mode="metadata", folder="Archive/0")["items"][0]["mail_id"], mails[0]["id"])
                    self.assertEqual(app.search("", mode="metadata", deduplicate=False)["count"], 2)
                    self.assertEqual(app.get_mail(mails[0]["id"])["cc"], "工程チーム")
                    self.assertEqual(app.search("timing", mode="lexical", sender="田中", folder="Archive/")["count"], 1)
                with sql:
                    sql.execute("DROP TRIGGER remove_mail_metadata")
                    sql.execute("DROP TABLE mail_metadata")
                sql.close()
                sql, table = app.open_store(cfg)
                self.assertEqual(sql.execute("SELECT count(*) FROM emails").fetchone()[0], 2)
                self.assertIsNotNone(sql.execute("SELECT name FROM sqlite_master WHERE name='mail_metadata'").fetchone())
            finally:
                sql.close()

    def test_mcp_blocked_search_leaves_original_and_status_responsive(self):
        async def run():
            started, release = threading.Event(), threading.Event()
            def slow(*args, **kwargs):
                started.set()
                if not release.wait(5):
                    raise TimeoutError()
                return {"count": 0}
            backend = SimpleNamespace(search=slow, status=lambda: {"emails": 1}, get_mail=lambda *args: {"body": "saved"})
            mcp = FastMCP("fixture")
            server.register_tools(mcp, backend)
            async with create_connected_server_and_client_session(mcp) as session:
                task = asyncio.create_task(session.call_tool("search_emails", {"query": "slow"}))
                try:
                    self.assertTrue(await asyncio.to_thread(started.wait, 2))
                    await asyncio.wait_for(session.send_ping(), .5)
                    await asyncio.wait_for(session.call_tool("index_status", {}), .5)
                    await asyncio.wait_for(session.call_tool("get_indexed_mail", {"mail_id": "a" * 64}), .5)
                    self.assertFalse(task.done())
                finally:
                    release.set()
                    await task
        asyncio.run(run())
