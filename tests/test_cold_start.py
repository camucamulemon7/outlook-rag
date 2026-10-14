"""First-use and partially initialized indexes must remain safe to inspect."""
import json
from contextlib import closing
import os
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from outlook_rag import app, features


class ColdIndex(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.index = self.root / "index"
        (self.root / "config.json").write_text("{}", encoding="utf-8")
        self.env = patch.dict(os.environ, {
            "OUTLOOK_RAG_MODEL": "fixture",
            "OUTLOOK_RAG_DATA_DIR": str(self.index),
            "OUTLOOK_RAG_CONFIG": str(self.root / "config.json"),
            "OUTLOOK_RAG_EMBEDDING_URL": "http://127.0.0.1:1/embeddings",
            "LOCALAPPDATA": str(self.root),
        }, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)
        probe = patch.object(app, "detect_dimensions", side_effect=AssertionError("API probe forbidden"))
        probe.start()
        self.addCleanup(probe.stop)

    def assert_unready(self, state):
        with patch("outlook_rag.storage.open_index", side_effect=AssertionError("Vector open forbidden")):
            self.assertEqual(app.status()["state"], state)
            for mode in ("hybrid", "semantic", "lexical", "metadata", "thread"):
                with self.subTest(mode=mode):
                    result = app.search("example", mode=mode)
                    self.assertEqual(result["state"], state)
                    self.assertEqual(result["count"], 0)
                    self.assertFalse(result["absence_confirmed"])

    def test_new_index_does_not_create_files_or_probe_api(self):
        self.assert_unready("not_initialized")
        self.assertFalse(self.index.exists())

    def test_schema_creation_window_is_not_a_corrupt_index(self):
        self.index.mkdir()
        with closing(sqlite3.connect(self.index / "metadata.sqlite")):
            pass
        self.assert_unready("initializing")
        self.assertEqual(app.get_sync_progress()["folders"], [])

    def test_identity_committed_before_vector_creation(self):
        self.index.mkdir()
        cfg = app.config(resolve_dimensions=False)
        identity = app.index_identity(dict(cfg, dimensions=4))
        with closing(sqlite3.connect(self.index / "metadata.sqlite")) as sql, sql:
            sql.execute("CREATE TABLE state(key TEXT PRIMARY KEY,value TEXT)")
            sql.execute("CREATE TABLE emails(id TEXT)")
            sql.execute("INSERT INTO state VALUES('embedding_identity',?)", (identity,))
        self.assert_unready("initializing")

    def test_invalid_database_is_not_reported_as_no_mail(self):
        self.index.mkdir()
        (self.index / "metadata.sqlite").write_bytes(b"not sqlite")
        with self.assertRaises(ValueError):
            app.status()

    def test_populated_metadata_without_identity_needs_repair(self):
        self.index.mkdir()
        with closing(sqlite3.connect(self.index / "metadata.sqlite")) as sql, sql:
            sql.execute("CREATE TABLE emails(id TEXT)")
            sql.execute("INSERT INTO emails VALUES('a')")
        with self.assertRaises(ValueError):
            features.local_config()

    def test_failed_vector_initialization_closes_sqlite_and_remains_inspectable(self):
        created = []
        connect = sqlite3.connect
        def record(*args, **kwargs):
            connection = connect(*args, **kwargs)
            created.append(connection)
            return connection
        cfg = dict(app.config(resolve_dimensions=False), dimensions=4)
        with patch('outlook_rag.database.sqlite3.connect', side_effect=record), patch.object(
                app.lancedb, 'connect', side_effect=RuntimeError('Synthetic native initialization failure')):
            with self.assertRaises(RuntimeError):
                app.open_store(cfg)
        self.assertEqual(len(created), 1)
        with self.assertRaises(sqlite3.ProgrammingError):
            created[0].execute('SELECT 1')
        self.assert_unready('initializing')
