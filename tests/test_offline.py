"""No email account, API server or API key required. Run with unittest discovery."""
import hashlib
import json
import math
import os
import tempfile
import threading
import time
import unittest
from datetime import datetime, timezone
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

import httpx

from outlook_rag import app


class IndexTests(unittest.TestCase):
    def setUp(self):
        scratch = Path(os.environ.get("OUTLOOK_RAG_TEST_DIR", "work"))
        scratch.mkdir(exist_ok=True, parents=True)
        self.temp = tempfile.TemporaryDirectory(dir=scratch, ignore_cleanup_errors=True)
        self.cfg = {"model": "fixture", "dimensions": 16, "embedding_url": "http://fixture.invalid/embeddings",
                    "data_dir": str(Path(self.temp.name) / "index"), "embedding_batch_size": 4}
        self.sql, self.table = app.open_store(self.cfg)
        self.calls = []
        def fake_embed(texts, cfg, **kwargs):
            self.calls.append(list(texts))
            output = []
            for text in texts:
                values = [v + 1.0 for v in hashlib.sha256(text.encode()).digest()[:16]]
                norm = math.sqrt(sum(v*v for v in values))
                output.append([v / norm for v in values])
            return output
        self.patch = patch.object(app, "embed", side_effect=fake_embed)
        self.patch.start()

    def tearDown(self):
        self.patch.stop()
        self.sql.close()
        self.temp.cleanup()

    def mail(self, index, body="納期が延期されました。"):
        return dict(id=hashlib.sha256(str(index).encode()).hexdigest(), entry_id=str(index), store_id="fixture",
                    folder="inbox", subject="納期について", body=body, sender="fixture@example.invalid",
                    recipients="fixture@example.invalid", received="2026-09-01T12:00:00+09:00", modified="1", conversation_id="fixture")

    def test_batch_deduplicates_and_reuses_cached_embeddings(self):
        mails = [self.mail(i) for i in range(8)]
        stats = app.upsert_mails(mails, self.sql, self.table, self.cfg)
        self.assertEqual(stats["embedded_chunks"], 1)
        self.assertEqual(len(self.calls), 1)
        self.assertEqual(self.table.count_rows(), 8)
        stats = app.upsert_mails([self.mail(9)], self.sql, self.table, self.cfg)
        self.assertEqual(stats["embedded_chunks"], 0)
        self.assertEqual(stats["cache_hits"], 1)
        self.assertEqual(len(self.calls), 1)

    def test_metadata_change_does_not_reembed(self):
        mail = self.mail(1)
        app.upsert_mail(mail, self.sql, self.table, self.cfg)
        mail.update(modified="2", sender="updated@example.invalid")
        stats = app.upsert_mail(mail, self.sql, self.table, self.cfg)
        self.assertEqual(stats["unchanged_content"], 1)
        self.assertEqual(stats["embedding_requests"], 0)
        self.assertEqual(self.sql.execute("SELECT sender FROM emails").fetchone()[0], "updated@example.invalid")
        self.assertEqual(self.table.count_rows(), 1)

    def test_reduced_dimensions_require_separate_index(self):
        with self.assertRaises(ValueError):
            app.open_store(dict(self.cfg, storage_dimensions=8))

    def test_invalid_outlook_unicode_does_not_block_hashing_embedding_or_sql(self):
        mail = self.mail(1,body="before\udce2after\ud83d\ude00")
        mail.update(subject="title\ud800",sender="sender\udcff")
        stats=app.upsert_mails([mail],self.sql,self.table,self.cfg)
        self.assertEqual(stats["sanitized_emails"],1)
        row=self.sql.execute("SELECT subject,body,sender FROM emails").fetchone()
        self.assertEqual(row['body'],"before\ufffdafter\U0001f600")
        self.assertEqual(row['subject'],"title\ufffd")
        self.assertEqual(row['sender'],"sender\ufffd")
        self.assertEqual(mail['body'],"before\udce2after\ud83d\ude00")
        for batch in self.calls:
            for text in batch:
                text.encode('utf-8')

    def test_all_time_sync_has_no_received_date_lower_bound(self):
        restrictions = []
        class Items(list):
            def Restrict(self, expression):
                restrictions.append(expression)
                return self
            def Sort(self, *args):
                pass
        mail = SimpleNamespace(Class=43, EntryID="ancient", Subject="old", Body="old body",
            ReceivedTime=datetime(1995, 1, 1, tzinfo=timezone.utc), LastModificationTime=datetime.now(timezone.utc))
        folder = SimpleNamespace(EntryID="folder", StoreID="store", Items=Items([mail]))
        outlook = SimpleNamespace(GetNamespace=lambda _: object())
        cfg = dict(self.cfg, since_days="all", folders=["fixture"])
        with patch.object(app, "config", return_value=cfg), patch.object(app, "resolve_folder", return_value=folder), patch("win32com.client.Dispatch", return_value=outlook):
            result = app.sync(max_emails=1)
            self.assertEqual(result["folders"][0]["indexed"], 1)
            self.assertEqual(restrictions, [])
            result = app.sync(max_emails=1)
            self.assertTrue(result["folders"][0]["incremental"])
            self.assertTrue(all("ReceivedTime" not in value for value in restrictions))
            folder.Items.clear()
            result = app.sync(reconcile=True)
            self.assertEqual(result["folders"][0]["removed_local"], 1)

    def test_global_sync_budget_prevents_large_multi_folder_calls(self):
        class Items(list):
            def Restrict(self, expression): return self
            def Sort(self,*args): pass
        def folder(name):
            messages=[SimpleNamespace(Class=43,EntryID=name+str(i),Subject=name,Body="body",
                ReceivedTime=datetime(1995,1,1,tzinfo=timezone.utc),LastModificationTime=datetime.now(timezone.utc)) for i in range(4)]
            return SimpleNamespace(EntryID=name,StoreID="store",Items=Items(messages))
        folders={name:folder(name) for name in ('first','second')}
        cfg=dict(self.cfg,folders=list(folders),since_days=0)
        outlook=SimpleNamespace(GetNamespace=lambda _:object())
        with patch.object(app,'config',return_value=cfg),patch.object(app,'resolve_folder',side_effect=lambda namespace,name:folders[name]),patch('win32com.client.Dispatch',return_value=outlook):
            result=app.sync(max_emails=500,max_total_emails=2)
        self.assertEqual(result['indexed_total'],2)
        self.assertEqual(result['pending_folders'],['second'])
        self.assertFalse(result['folders'][0]['window_complete'])

    def test_bounded_concurrent_embedding(self):
        active = maximum = 0
        lock = threading.Lock()
        original = app.embed.side_effect
        def slow_embed(texts, cfg, **kwargs):
            nonlocal active, maximum
            with lock:
                active += 1
                maximum = max(maximum, active)
            try:
                time.sleep(0.02)
                return original(texts, cfg, **kwargs)
            finally:
                with lock:
                    active -= 1
        cfg = dict(self.cfg, embedding_batch_size=1, embedding_concurrent_requests=2)
        mails = [self.mail(i, body=f"異なる本文 {i}") for i in range(8)]
        with patch.object(app, "embed", side_effect=slow_embed):
            stats = app.upsert_mails(mails, self.sql, self.table, cfg)
        self.assertEqual(maximum, 2)
        self.assertEqual(stats["embedding_requests"], 8)
        self.assertEqual(self.table.count_rows(), 8)
        for mail in mails:
            text = app.chunks(mail["subject"], mail["body"])[0]
            digest = hashlib.sha256(text.encode()).hexdigest()
            self.assertIsNotNone(self.sql.execute("SELECT vector FROM embedding_cache WHERE hash=?", (digest,)).fetchone())

    def test_embedding_failure_keeps_previous_index(self):
        mail = self.mail(1)
        app.upsert_mail(mail, self.sql, self.table, self.cfg)
        updated = dict(mail, body="出荷予定を再度変更しました。", modified="2")
        with patch.object(app, "embed", side_effect=RuntimeError("fixture failure")):
            with self.assertRaises(RuntimeError):
                app.upsert_mail(updated, self.sql, self.table, self.cfg)
        self.assertEqual(self.sql.execute("SELECT body FROM emails").fetchone()[0], mail["body"])
        self.assertEqual(self.table.count_rows(), 1)
        app.upsert_mail(updated, self.sql, self.table, self.cfg)
        self.assertEqual(self.sql.execute("SELECT body FROM emails").fetchone()[0], updated["body"])
        self.assertEqual(self.table.count_rows(), 1)

    def test_partial_vector_write_recovers_on_retry(self):
        mail = self.mail(1)
        app.upsert_mail(mail, self.sql, self.table, self.cfg)
        updated = dict(mail, body="出荷が来月まで遅れます。", modified="2")
        with patch.object(self.table, "add", side_effect=RuntimeError("fixture write failure")):
            with self.assertRaises(RuntimeError):
                app.upsert_mail(updated, self.sql, self.table, self.cfg)
        self.assertEqual(self.sql.execute("SELECT modified FROM emails").fetchone()[0], "1")
        app.upsert_mail(updated, self.sql, self.table, self.cfg)
        self.assertEqual(self.table.count_rows(), 1)
        self.assertEqual(self.sql.execute("SELECT modified FROM emails").fetchone()[0], "2")


class ConfigTests(unittest.TestCase):
    def test_minimal_settings_detect_dimensions_once_and_default_to_all_sources(self):
        app._DIMENSIONS.clear()
        transport = httpx.MockTransport(lambda request: httpx.Response(200,json={"data":[{"index":0,"embedding":[1.0]*16}]}))
        actual_client = httpx.Client
        with patch.dict(os.environ, {"OUTLOOK_RAG_MODEL":"fixture-auto", "OUTLOOK_RAG_API_KEY":"fixture", "USERPROFILE":str(Path('work/profile').resolve()), "LOCALAPPDATA":str(Path('work').resolve())},clear=True):
            with patch.object(app.httpx,"Client",side_effect=lambda **kwargs:actual_client(transport=transport,**kwargs)) as clients:
                cfg=app.config()
                app.config()
                self.assertEqual(clients.call_count,1)
        self.assertEqual(cfg["dimensions"],16)
        self.assertEqual(app.vector_dimensions(cfg),16)
        self.assertEqual(cfg["folders"],"auto")
        self.assertEqual(cfg["since_days"],0)
        self.assertEqual(cfg["embedding_batch_size"],8)
        self.assertEqual(cfg["embedding_concurrent_requests"],2)
    def test_default_database_location_and_legacy_compatibility(self):
        with tempfile.TemporaryDirectory() as root:
            home = Path(root) / "profile"
            local = Path(root) / "local"
            env = dict(LOCALAPPDATA=str(local), OUTLOOK_RAG_MODEL="fixture-path",
                       OUTLOOK_RAG_DIMENSIONS="16")
            with patch.dict(os.environ, env, clear=True), patch.object(Path, "home", return_value=home):
                cfg = app.config()
                preferred = Path(cfg["data_dir"])
                self.assertEqual(preferred.parent, home / "Documents" / "Outlook\u30d5\u30a1\u30a4\u30eb" / "outlook-rag")
                self.assertEqual(Path(cfg["key_file"]), local / "outlook-rag" / "api-key.dpapi")
                legacy = local / "outlook-rag" / hashlib.sha256(app.index_identity(cfg).encode()).hexdigest()[:12]
                legacy.mkdir(parents=True)
                self.assertEqual(Path(app.config()["data_dir"]), preferred)
                (legacy / "metadata.sqlite").touch()
                self.assertEqual(Path(app.config()["data_dir"]), legacy)
                preferred.mkdir(parents=True)
                (preferred / "metadata.sqlite").touch()
                self.assertEqual(Path(app.config()["data_dir"]), preferred)
                explicit = Path(root) / "custom"
                with patch.dict(os.environ, {"OUTLOOK_RAG_DATA_DIR": str(explicit)}):
                    self.assertEqual(Path(app.config()["data_dir"]), explicit)

    def test_mcp_environment_without_configuration_file(self):
        env = dict(LOCALAPPDATA=str(Path('work').resolve()), OUTLOOK_RAG_EMBEDDING_URL="http://fixture.invalid", OUTLOOK_RAG_MODEL="generic",
                   OUTLOOK_RAG_DIMENSIONS="768", OUTLOOK_RAG_DATA_DIR="./fixture-data",
                   OUTLOOK_RAG_SINCE_DAYS="all", OUTLOOK_RAG_FOLDERS='["inbox"]')
        with patch.dict(os.environ, env, clear=True):
            cfg = app.config()
        self.assertEqual(cfg["since_days"], 0)
        self.assertEqual(cfg["dimensions"], 768)
        self.assertEqual(cfg["folders"], ["inbox"])
        self.assertEqual(app.vector_dimensions(cfg), 768)
        self.assertEqual(app.query_instruction(cfg), "")
    def test_relative_runtime_paths_resolve_from_config(self):
        scratch = Path(os.environ.get("OUTLOOK_RAG_TEST_DIR", "work"))
        scratch.mkdir(exist_ok=True, parents=True)
        with tempfile.TemporaryDirectory(dir=scratch, ignore_cleanup_errors=True) as root:
            path = Path(root) / "config.json"
            path.write_text(json.dumps(dict(embedding_url="http://fixture.invalid", model="fixture", dimensions=16,
                                           data_dir="./data", key_file="./secret.dpapi")), encoding="utf-8")
            with patch.dict(os.environ, {"OUTLOOK_RAG_CONFIG": str(path)}):
                cfg = app.config()
            self.assertEqual(Path(cfg["data_dir"]), path.resolve().parent / "data")
            self.assertEqual(Path(cfg["key_file"]), path.resolve().parent / "secret.dpapi")


class EmbeddingTests(unittest.TestCase):
    def test_generic_model_prefixes_are_applied_once(self):
        cfg = dict(embedding_url="http://fixture.invalid", model="generic", dimensions=2,
                   query_prefix="query: ", document_prefix="passage: ")
        inputs = []
        def respond(request):
            inputs.append(json.loads(request.content)["input"])
            return httpx.Response(200, json={"data": [{"index": 0, "embedding": [1., 0.]}]})
        with httpx.Client(transport=httpx.MockTransport(respond)) as client:
            with patch.object(app, "api_key", return_value="fixture"):
                app.embed(["text"], cfg, query=True, client=client)
                app.embed(["text"], cfg, client=client)
        self.assertEqual(inputs, [["query: text"], ["passage: text"]])

    def test_mrl_prefix_is_renormalized_for_queries_and_documents(self):
        cfg = dict(embedding_url="http://fixture.invalid", model="fixture", dimensions=4, storage_dimensions=2)
        data = {"data": [{"index": 0, "embedding": [3., 4., 8., 9.]}]}
        with httpx.Client(transport=httpx.MockTransport(lambda request: httpx.Response(200, json=data))) as client:
            with patch.object(app, "api_key", return_value="fixture"):
                for query in (False, True):
                    vector = app.embed(["one"], cfg, query=query, client=client)[0]
                    self.assertEqual(vector, [0.6, 0.8])
        with self.assertRaises(ValueError):
            app.vector_dimensions(dict(cfg, storage_dimensions=5))

    def test_rejects_duplicate_response_indices(self):
        cfg = dict(embedding_url="http://fixture.invalid", model="fixture", dimensions=2)
        data = {"data": [{"index": 0, "embedding": [1., 0.]}, {"index": 0, "embedding": [0., 1.]}]}
        with httpx.Client(transport=httpx.MockTransport(lambda request: httpx.Response(200, json=data))) as client:
            with patch.object(app, "api_key", return_value="fixture"):
                with self.assertRaises(ValueError):
                    app.embed(["one", "two"], cfg, client=client)


class CleaningTests(unittest.TestCase):
    def test_reply_header_requires_correlated_labels(self):
        body = "新しい締切は金曜日です。\n\n差出人: old\n送信日時: yesterday\n宛先: user\n件名: old\n古い締切は水曜日です。"
        self.assertEqual(app.clean_body(body), "新しい締切は金曜日です。")
        self.assertIn("From: source", app.clean_body("本文\nFrom: source\n重要な追加情報"))

    def test_cleanup_keeps_content_after_unsubscribe_banner(self):
        body = "配信停止はこちら\nキャンペーンは金曜日まで。\n価格は1000円です。"
        cleaned = app.clean_body(body)
        self.assertNotIn("配信停止はこちら", cleaned)
        self.assertIn("金曜日", cleaned)
        self.assertIn("1000円", cleaned)

    def test_signature_and_tracking_noise(self):
        body = "重要な通知\n" + "\u034f\u200c " * 20 + "\n<https://images.example.invalid/logo.png>\nhttps://tracking.example.invalid/" + "x" * 200 + "\n-- \n署名"
        cleaned = app.clean_body(body)
        self.assertIn("重要な通知", cleaned)
        self.assertIn("https://tracking.example.invalid/", cleaned)
        self.assertNotIn("logo.png", cleaned)
        self.assertNotIn("署名", cleaned)
        self.assertNotIn("\u034f", cleaned)
        self.assertIn("logo.png", app.clean_body(body, version=1))


class DiscoveryTests(unittest.TestCase):
    def test_archive_and_subfolders_are_discovered_without_scanning_mail_bodies(self):
        def folder(name, children=(), mail=True):
            return SimpleNamespace(Name=name,EntryID=name,DefaultItemType=0 if mail else 1,
                Items=SimpleNamespace(Count=1),Folders=list(children))
        archive=folder("Archive",[folder("2010"),folder("Calendar",mail=False)])
        root=folder("Root",[archive,folder("Junk"),folder("Search")])
        store=SimpleNamespace(DisplayName="Archive PST",FilePath="C:/archive.pst",GetRootFolder=lambda:root,
            GetDefaultFolder=lambda kind:folder("Junk"),GetSearchFolders=lambda:[folder("Search")])
        report,targets=app.discover_sources(SimpleNamespace(Stores=[store]),{})
        paths={path for path,_ in targets}
        self.assertIn("Archive PST/Archive/2010",paths)
        self.assertNotIn("Archive PST/Junk",paths)
        self.assertNotIn("Archive PST/Search",paths)
        self.assertNotIn("Archive PST/Archive/Calendar",paths)
        self.assertEqual(report["stores"][0]["file_path"],"C:/archive.pst")
        self.assertEqual(report["errors"],[])


if __name__ == "__main__":
    unittest.main()
