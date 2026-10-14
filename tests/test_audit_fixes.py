"""Regression coverage for issues #3-#7; no real mail or provider required."""

import hashlib
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from outlook_rag import app, features, jobs
import test_improvements as fixtures


class MetadataArrays(unittest.TestCase):
    def test_opt_in_outlook_probe_compares_rows_without_reading_bodies(self):
        from verify_outlook_metadata import probe
        from outlook_rag import outlook

        for mismatch in (False, True):
            with self.subTest(mismatch=mismatch):

                class Table:
                    Columns = SimpleNamespace(
                        RemoveAll=lambda: None, Add=lambda _: None
                    )
                    position = 0
                    source = [
                        ("id1", "IPM.Note", "2026-09-01", "2026-09-02"),
                        ("id2", "IPM.Note", "2026-09-01", "2026-09-02"),
                    ]

                    def Sort(self, *_):
                        pass

                    @property
                    def EndOfTable(self):
                        return self.position >= len(self.source)

                    def GetArray(self, n):
                        rows = self.source[self.position : self.position + n]
                        self.position += len(rows)
                        return tuple(zip(*rows))

                    def MoveToStart(self):
                        self.position = 0

                    def GetNextRow(self):
                        values = dict(
                            zip(
                                (
                                    "EntryID",
                                    "MessageClass",
                                    outlook.RECEIVED,
                                    outlook.MODIFIED,
                                ),
                                self.source[self.position],
                            )
                        )
                        self.position += 1
                        if mismatch:
                            values[outlook.RECEIVED] = "wrong"
                        return SimpleNamespace(Item=lambda name: values[name])

                table = Table()
                folder = SimpleNamespace(GetTable=lambda _: table)
                if mismatch:
                    with self.assertRaisesRegex(ValueError, "metadata differ"):
                        probe(folder)
                else:
                    result = probe(folder)
                    self.assertEqual(result["rows_checked"], 2)
                    self.assertFalse(result["mail_bodies_read"])

    def test_legacy_column_major_batches(self):
        for count in (0, 1, 2, 4, 63, 64, 65, 129):
            with self.subTest(count=count):
                source = [
                    (
                        str(i),
                        "IPM.Note" if i % 2 else "IPM.Appointment",
                        "2026-09-01T12:00:00",
                        "2026-09-02T12:00:00",
                    )
                    for i in range(count)
                ]

                class Table:
                    Columns = SimpleNamespace(
                        RemoveAll=lambda: None, Add=lambda _: None
                    )
                    position = 0

                    def Sort(self, *_):
                        pass

                    @property
                    def EndOfTable(self):
                        return self.position >= count

                    def GetArray(self, n):
                        rows = source[self.position : self.position + n]
                        self.position += len(rows)
                        columns = tuple(zip(*rows))
                        return columns

                table = Table()
                rows = list(
                    app.mail_metadata(SimpleNamespace(GetTable=lambda _: table), [])
                )
                self.assertEqual(
                    [r["entry_id"] for r, _ in rows], [str(i) for i in range(count)]
                )
                self.assertEqual(
                    [r["mail"] for r, _ in rows], [bool(i % 2) for i in range(count)]
                )
                self.assertTrue(
                    all(r["received"] == "2026-09-01T12:00:00" for r, _ in rows)
                )
                self.assertEqual(table.position, count)

    def test_ragged_or_wrong_column_count_is_rejected(self):
        for values in (
            (("id",), ("IPM.Note",)),
            (("id",), ("IPM.Note",), (), ("date",)),
        ):
            table = SimpleNamespace(
                Columns=SimpleNamespace(RemoveAll=lambda: None, Add=lambda _: None),
                Sort=lambda *_: None,
                EndOfTable=False,
                GetArray=lambda _: values,
            )
            with (
                self.subTest(values=values),
                self.assertRaisesRegex(ValueError, "table shape"),
            ):
                list(app.mail_metadata(SimpleNamespace(GetTable=lambda _: table), []))

    def test_empty_columns_do_not_loop(self):
        table = SimpleNamespace(
            Columns=SimpleNamespace(RemoveAll=lambda: None, Add=lambda _: None),
            Sort=lambda *_: None,
            EndOfTable=False,
            GetArray=lambda _: ((), (), (), ()),
        )
        self.assertEqual(
            list(app.mail_metadata(SimpleNamespace(GetTable=lambda _: table), [])), []
        )


class PublishedIndex(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.cfg = dict(
            model="fixture",
            dimensions=2,
            data_dir=str(Path(self.temp.name) / "index"),
            embedding_url="http://fixture.invalid",
            query_cache_size=0,
            folders=["inbox"],
        )
        self.sql, self.table = app.open_store(self.cfg)
        self.embeddings = patch.object(app, "embed", side_effect=self.vectors)
        self.embeddings.start()
        self.configuration = patch.object(app, "config", return_value=self.cfg)
        self.configuration.start()

    def tearDown(self):
        self.configuration.stop()
        self.embeddings.stop()
        self.sql.close()
        self.temp.cleanup()

    @staticmethod
    def vectors(texts, *_args, **_kwargs):
        return [[0.0, 1.0] if "new" in t else [1.0, 0.0] for t in texts]

    def mail(self, i=1, body="old body", **fields):
        return dict(
            dict(
                id=hashlib.sha256(str(i).encode()).hexdigest(),
                entry_id=str(i),
                store_id="store",
                folder="inbox",
                subject="fixture",
                body=body,
                modified="1",
                sender="fixture@example.invalid",
                recipients="fixture@example.invalid",
                received="2026-09-01T12:00:00+09:00",
                conversation_id=str(i),
            ),
            **fields,
        )

    def assert_search(self, query, text, similarity=1.0):
        found = app.search(query, hybrid=False)
        self.assertEqual(found["count"], 1)
        self.assertIn(text, found["items"][0]["snippet"])
        self.assertEqual(found["items"][0]["semantic_similarity"], similarity)

    def test_vector_add_failure_preserves_old_search_and_reverted_content(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)

        def failed_add(_):
            # Read through a different DB/table while the writer has deleted rows.
            self.assert_search("old", "old body")
            raise RuntimeError("synthetic write failure")

        with (
            patch.object(self.table, "add", side_effect=failed_add),
            self.assertRaises(RuntimeError),
        ):
            app.upsert_mail(
                dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
            )
        self.assertEqual(self.table.count_rows(), 1)
        self.assert_search("old", "old body")
        app.upsert_mail(dict(old, modified="3"), self.sql, self.table, self.cfg)
        self.assert_search("old", "old body")
        self.assertEqual(self.table.count_rows(), 1)
        self.assertEqual(
            self.sql.execute("SELECT modified FROM emails").fetchone()[0], "3"
        )

    def test_sql_publication_failure_keeps_matching_vectors_and_reuses_cache(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        with self.sql:
            self.sql.execute(
                "CREATE TRIGGER fail_publish BEFORE INSERT ON emails WHEN NEW.modified='2' BEGIN SELECT RAISE(ABORT, 'synthetic failure'); END"
            )
        with self.assertRaises(sqlite3.IntegrityError):
            app.upsert_mail(
                dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
            )
        self.assert_search("new", "old body", 0.0)
        with self.sql:
            self.sql.execute("DROP TRIGGER fail_publish")
        stats = app.upsert_mail(
            dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
        )
        self.assertEqual(stats["embedding_requests"], 0)
        self.assert_search("new", "new body")

    def test_commit_time_constraint_failure_rolls_back_publication_marker(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        published = app.pin_vectors(self.sql, self.table)
        self.sql.execute("PRAGMA foreign_keys=ON")
        self.sql.executescript("""
            CREATE TABLE commit_parent(id INTEGER PRIMARY KEY);
            CREATE TABLE commit_child(id INTEGER REFERENCES commit_parent(id) DEFERRABLE INITIALLY DEFERRED);
            CREATE TRIGGER fail_at_commit AFTER INSERT ON emails WHEN NEW.modified='2'
            BEGIN INSERT INTO commit_child VALUES (999); END;
        """)
        statements = []
        self.sql.set_trace_callback(statements.append)
        with self.assertRaises(sqlite3.IntegrityError):
            app.upsert_mail(
                dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
            )
        self.sql.set_trace_callback(None)
        self.assertTrue(
            any(
                "INSERT OR REPLACE INTO state VALUES ('vector_version'" in statement
                for statement in statements
            )
        )
        self.assertIn("COMMIT", statements)
        self.assertEqual(app.pin_vectors(self.sql, self.table), published)
        self.assertEqual(
            self.sql.execute("SELECT count(*) FROM commit_child").fetchone()[0], 0
        )
        self.assert_search("new", "old body", 0.0)
        with self.sql:
            self.sql.execute("DROP TRIGGER fail_at_commit")
        stats = app.upsert_mail(
            dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
        )
        self.assertEqual(stats["embedding_requests"], 0)
        self.assert_search("new", "new body")
        self.assertEqual(self.table.count_rows(), 1)

    def test_process_exit_before_sql_commit_reopens_old_version_and_resumes(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        config_path = Path(self.temp.name) / "crash-input.json"
        config_path.write_text(
            json.dumps(
                dict(cfg=self.cfg, mail=dict(old, body="new body", modified="2"))
            ),
            encoding="utf-8",
        )
        code = """
import json, os, sys
from unittest.mock import patch
from outlook_rag import app
value = json.load(open(sys.argv[1], encoding='utf-8'))
sql, table = app.open_store(value['cfg'])
original = table.add
def crash(rows):
    original(rows)
    os._exit(7)
with patch.object(app, 'embed', side_effect=lambda texts, *a, **kw: [[0., 1.] for _ in texts]), patch.object(table, 'add', side_effect=crash):
    app.upsert_mail(value['mail'], sql, table, value['cfg'])
"""
        process = subprocess.run(
            [sys.executable, "-c", code, str(config_path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        self.assertEqual(process.returncode, 7, process.stderr)
        self.sql.close()
        self.sql, self.table = app.open_store(self.cfg)
        self.assert_search("old", "old body")
        stats = app.upsert_mail(
            dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
        )
        self.assertEqual(stats["embedding_requests"], 0)
        self.assert_search("new", "new body")
        self.assertEqual(self.table.count_rows(), 1)

    def test_failed_mail_deletion_preserves_published_content(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        with self.sql:
            self.sql.execute(
                "CREATE TRIGGER fail_delete BEFORE DELETE ON emails BEGIN SELECT RAISE(ABORT, 'synthetic failure'); END"
            )
        with self.assertRaises(sqlite3.IntegrityError):
            app.remove_mail(old["id"], self.sql, self.table)
        self.assert_search("old", "old body")
        with self.sql:
            self.sql.execute("DROP TRIGGER fail_delete")
        app.remove_mail(old["id"], self.sql, self.table)
        self.assertEqual(self.table.count_rows(), 0)

    def test_reader_snapshot_survives_concurrent_successful_publication(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        sql, table = app.open_store(self.cfg)
        errors = []
        ready, release = threading.Event(), threading.Event()

        def writer():
            try:
                writer_sql, writer_table = app.open_store(self.cfg)
                ready.set()
                try:
                    if not release.wait(5):
                        raise AssertionError("Reader did not release the writer")
                    with app.mutation_lock(self.cfg):
                        app.upsert_mail(
                            dict(old, body="new body", modified="2"),
                            writer_sql,
                            writer_table,
                            self.cfg,
                        )
                finally:
                    writer_sql.close()
            except BaseException as exc:
                errors.append(exc)

        worker = threading.Thread(target=writer, daemon=True)
        worker.start()
        try:
            self.assertTrue(ready.wait(5))
            with app.indexed_snapshot(sql, table, self.cfg):
                release.set()
                worker.join(5)
                self.assertFalse(worker.is_alive())
                self.assertEqual(errors, [])
                self.assertEqual(
                    sql.execute("SELECT body FROM emails").fetchone()[0], "old body"
                )
                rows = table.search([1.0, 0.0]).distance_type("cosine").to_list()
                self.assertEqual(len(rows), 1)
                self.assertAlmostEqual(rows[0]["_distance"], 0.0)
            self.assert_search("new", "new body")
        finally:
            release.set()
            worker.join(5)
            sql.close()

    def test_legacy_index_adoption_and_missing_vector_repair(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        self.table.checkout_latest()
        self.table.delete(f"email_id='{old['id']}'")
        with self.sql:
            self.sql.execute("DELETE FROM state WHERE key='vector_version'")
        self.sql.close()
        self.sql, self.table = app.open_store(self.cfg)
        self.assertEqual(self.table.count_rows(), 0)
        stats = app.upsert_mail(dict(old, modified="2"), self.sql, self.table, self.cfg)
        self.assertEqual(stats["embedding_requests"], 0)
        self.assertEqual(stats["cache_hits"], 1)
        self.assert_search("old", "old body")

    def test_maintenance_protects_published_version_and_unfinished_caches(self):
        old = self.mail()
        app.upsert_mail(old, self.sql, self.table, self.cfg)
        with patch.object(self.table, "add", side_effect=RuntimeError("failed")):
            with self.assertRaises(RuntimeError):
                app.upsert_mail(
                    dict(old, body="new body", modified="2"),
                    self.sql,
                    self.table,
                    self.cfg,
                )
        result = app.maintain()
        self.assertEqual(result["removed_cached_embeddings"], 0)
        self.assert_search("old", "old body")
        self.assertIn("outlook-rag-published", self.table.tags.list())
        stats = app.upsert_mail(
            dict(old, body="new body", modified="2"), self.sql, self.table, self.cfg
        )
        self.assertEqual(stats["embedding_requests"], 0)
        app.maintain()
        self.assert_search("new", "new body")

    def long_mail_index(self, count=110, threads=False, lexical=False):
        first = self.mail(0, body="long", id="0" * 64)
        others = [
            self.mail(i, body="keyword " if lexical else "short", id=f"{i:064x}")
            for i in range(1, 10)
        ]

        def parts(subject, body, *args, **kwargs):
            return (
                ["keyword" if lexical else "long"] * count if body == "long" else [body]
            )

        def ranked(texts, *args, **kwargs):
            return [
                [1.0, 0.0]
                if kwargs.get("query") or t in ("keyword", "long")
                else [0.8, 0.6]
                for t in texts
            ]

        mails = [first] + others
        if threads:
            mails = [
                self.mail(i + 20, body="long", conversation_id="big")
                for i in range(count)
            ] + others
            parts = lambda subject, body, *args, **kwargs: [body]
        with (
            patch.object(app, "chunks", side_effect=parts),
            patch.object(app, "embed", side_effect=ranked),
        ):
            app.upsert_mails(mails, self.sql, self.table, self.cfg)
        return ranked

    def test_long_mail_does_not_starve_distinct_semantic_results(self):
        ranked = self.long_mail_index()
        with patch.object(app, "embed", side_effect=ranked):
            result = app.search("query", limit=10, hybrid=False)
        self.assertEqual(result["count"], 10)
        self.assertEqual(len({row["mail_id"] for row in result["items"]}), 10)
        self.assertEqual(result["candidate_limit"], 200)
        self.assertFalse(result["candidate_limit_reached"])

    def test_large_thread_does_not_starve_other_conversations(self):
        ranked = self.long_mail_index(threads=True)
        with patch.object(app, "embed", side_effect=ranked):
            result = app.search("query", limit=10, hybrid=False, group_by_thread=True)
        self.assertEqual(result["count"], 10)
        self.assertEqual(result["items"][0]["thread_mail_count"], 110)
        self.assertEqual(result["candidate_limit"], 200)

    def test_lexical_candidates_expand_and_metadata_filters_remain_applied(self):
        ranked = self.long_mail_index(lexical=True)
        self.cfg.update(semantic_weight=0.0, lexical_weight=1.0)
        with patch.object(app, "embed", side_effect=ranked):
            result = app.search("keyword", limit=10)
            filtered = app.search("keyword", limit=10, sender="missing")
        self.assertEqual(result["count"], 10)
        self.assertEqual(result["candidate_limit"], 200)
        self.assertEqual(filtered["count"], 0)

    def test_candidate_expansion_has_a_hard_cap(self):
        ranked = self.long_mail_index(count=5001)
        with patch.object(app, "embed", side_effect=ranked):
            result = app.search("query", limit=2, hybrid=False)
        self.assertEqual(result["count"], 1)
        self.assertEqual(result["candidate_limit"], 5000)
        self.assertTrue(result["candidate_limit_reached"])

    def test_candidate_expansion_has_a_time_budget(self):
        ranked = self.long_mail_index()
        clock = SimpleNamespace(
            perf_counter=Mock(side_effect=[0.0, 0.0, 0.0, 0.0, 6.0, 7.0]),
            time=app.time.time,
        )
        with (
            patch.object(app, "embed", side_effect=ranked),
            patch.object(app, "time", clock),
        ):
            result = app.search("query", limit=10, hybrid=False)
        self.assertEqual(result["count"], 1)
        self.assertEqual(result["candidate_limit"], 100)
        self.assertTrue(result["candidate_time_limit_reached"])


class OfflineCandidates(unittest.TestCase):
    def setUp(self):
        from outlook_test_support import isolated_environment
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.environment = patch.dict(
            os.environ,
            isolated_environment(dict(
                OUTLOOK_RAG_MODEL="cached",
                OUTLOOK_RAG_DIMENSIONS="4",
                LOCALAPPDATA=str(self.root / "local"),
            )),
            clear=True,
        )
        self.environment.start()
        self.addCleanup(self.environment.stop)
        self.home = patch.object(Path, "home", return_value=self.root / "home")
        self.home.start()
        self.addCleanup(self.home.stop)
        self.cfg = app.config()
        self.normal_directory = Path(self.cfg["data_dir"])
        sql, _ = app.open_store(self.cfg)
        sql.close()
        os.environ["OUTLOOK_RAG_DIMENSIONS"] = "auto"
        self.probe = patch.object(
            app, "detect_dimensions", side_effect=AssertionError("API probe forbidden")
        )
        self.probe.start()
        self.addCleanup(self.probe.stop)

    def bad_candidate(self, name, body):
        folder = Path(self.cfg["data_dir"]).parent / name
        folder.mkdir()
        (folder / "metadata.sqlite").write_bytes(body)
        return folder

    def test_unrelated_partial_corrupt_and_invalid_identity_candidates_are_isolated(
        self,
    ):
        self.bad_candidate("partial", b"")
        self.bad_candidate("broken", b"not SQLite")
        for name, value in (("bad-json", "{private-token"), ("bad-shape", "[]")):
            folder = self.bad_candidate(name, b"")
            with sqlite3.connect(folder / "metadata.sqlite") as sql:
                sql.execute("CREATE TABLE state(key TEXT PRIMARY KEY, value TEXT)")
                sql.execute(
                    "INSERT INTO state VALUES (?,?)", ("embedding_identity", value)
                )
        with self.assertLogs(features.__name__, level="WARNING") as logs:
            resolved = features.local_config()
        self.assertEqual(resolved["data_dir"], self.cfg["data_dir"])
        self.assertEqual(resolved["dimensions"], 4)
        self.assertNotIn("private-token", "\n".join(logs.output))

    def test_bad_job_snapshot_does_not_hide_valid_snapshot_in_same_directory(self):
        folder = self.root / "snapshot-only"
        self.cfg["data_dir"] = str(folder)
        with patch.object(jobs, "launch_worker"):
            job = jobs.start(self.cfg)["job_id"]
        (jobs.root(self.cfg) / "bad.config.json").write_text(
            "{private-token", encoding="utf-8"
        )
        with patch.dict(os.environ, {"OUTLOOK_RAG_DATA_DIR": str(folder)}):
            # Explicit selections report damaged snapshots clearly.
            with self.assertRaisesRegex(
                ValueError, "Cannot read cached index settings"
            ):
                features.local_config()
        destination = self.normal_directory.parent / "snapshots"
        folder.rename(destination)
        (self.normal_directory / "metadata.sqlite").unlink()
        with self.assertLogs(features.__name__, level="WARNING"):
            resolved = features.local_config()
        self.assertEqual(resolved["data_dir"], str(destination))
        self.assertEqual(resolved["dimensions"], 4)
        self.assertEqual(
            jobs.load(dict(self.cfg, data_dir=str(destination)), job)[1]["job_id"], job
        )

    def test_explicit_damaged_selection_has_a_clear_error(self):
        damaged = self.bad_candidate("broken", b"not SQLite")
        with patch.dict(os.environ, {"OUTLOOK_RAG_DATA_DIR": str(damaged)}):
            with self.assertRaisesRegex(
                ValueError, "Cannot read cached index settings"
            ):
                features.local_config()

    def test_explicit_dimensions_do_not_hide_damaged_selected_index(self):
        damaged = self.bad_candidate("broken", b"not SQLite")
        with patch.dict(
            os.environ,
            {"OUTLOOK_RAG_DATA_DIR": str(damaged), "OUTLOOK_RAG_DIMENSIONS": "4"},
        ):
            with self.assertRaisesRegex(
                ValueError, "Cannot read cached index settings"
            ):
                app.list_sync_failures()

    def test_smaller_unrelated_model_does_not_break_storage_dimension_matching(self):
        folder = self.bad_candidate("different-model", b"")
        identity = app.index_identity(dict(model="other", dimensions=2))
        with sqlite3.connect(folder / "metadata.sqlite") as sql:
            sql.execute("CREATE TABLE state(key TEXT PRIMARY KEY, value TEXT)")
            sql.execute(
                "INSERT INTO state VALUES (?,?)", ("embedding_identity", identity)
            )
        with patch.dict(os.environ, {"OUTLOOK_RAG_STORAGE_DIMENSIONS": "4"}):
            self.assertEqual(
                features.local_config()["data_dir"], str(self.normal_directory)
            )

    def test_no_usable_match_reports_damaged_candidates(self):
        self.bad_candidate("broken", b"not SQLite")
        with (
            patch.dict(os.environ, {"OUTLOOK_RAG_MODEL": "missing"}),
            self.assertLogs(features.__name__, level="WARNING"),
        ):
            with self.assertRaisesRegex(ValueError, "No usable cached index"):
                features.local_config()


class ChangingJobSources(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.cfg = self.fixture.cfg
        self.cfg.update(folders="auto", sync_max_scanned=1, sync_prefetch_emails=0)

    def tearDown(self):
        self.fixture.tearDown()

    def run_job(self, discovery):
        with (
            patch.object(app, "config", return_value=self.cfg),
            patch.object(app, "api_key", return_value="fixture-key"),
            patch("win32com.client.Dispatch", return_value=self.fixture.dispatch),
            patch.object(app, "discover_sources", side_effect=discovery),
            patch.object(jobs, "launch_worker"),
        ):
            job = jobs.start(self.cfg)["job_id"]
            jobs.run(job, self.cfg)
        return jobs.status(self.cfg, job)

    def discovery(self, entries):
        return dict(
            errors=[],
            stores=[
                dict(
                    folders=[
                        dict(path=name, included=True, item_count=3)
                        for name, _ in entries
                    ]
                )
            ],
        ), entries

    def test_deleted_or_disconnected_pending_folder_ends_with_errors(self):
        self.fixture.folder("inbox", 2)
        self.fixture.folder("archive", 2)
        calls = []

        def discover(*_):
            calls.append(1)
            entries = [("inbox", self.fixture.folders["inbox"])]
            if len(calls) == 1:
                entries.append(("archive", self.fixture.folders["archive"]))
            return self.discovery(entries)

        state = self.run_job(discover)
        self.assertEqual(state["state"], "completed_with_errors")
        self.assertEqual(state["reason"], "folder_unavailable")
        self.assertEqual(state["cycles"], 2)
        self.assertEqual(state["missing_folders"][0]["folder"], "archive")

    def test_renamed_folder_keeps_its_completion_and_updates_cached_paths(self):
        self.fixture.folder("old", 3, same_time=True)
        folder = self.fixture.folders["old"]
        calls = []

        def discover(*_):
            calls.append(1)
            return self.discovery([("old" if len(calls) == 1 else "renamed", folder)])

        state = self.run_job(discover)
        self.assertEqual(state["state"], "completed")
        self.assertEqual(state["known_folders"], 1)
        self.assertEqual(state["completed_folders"], 1)
        self.assertEqual(state["cycles"], 3)
        self.assertEqual(state["missing_folders"], [])
        self.assertEqual(
            {r[0] for r in self.fixture.sql.execute("SELECT folder FROM emails")},
            {"renamed"},
        )

    def test_folders_with_identical_display_paths_are_tracked_separately(self):
        inventory = [dict(folder="same", folder_id=fid) for fid in ("first", "second")]
        cycles = [
            dict(
                folders=[dict(folder="same", folder_id=fid, window_complete=True)],
                target_folders=inventory,
                pending_folders=["same"] if fid == "first" else [],
                scanned_total=1,
                stop_reason=None,
            )
            for fid in ("first", "second")
        ]
        with (
            patch.object(app, "api_key", return_value="fixture-key"),
            patch.object(jobs, "launch_worker"),
        ):
            job = jobs.start(self.cfg)["job_id"]
        with (
            patch.object(app, "api_key", return_value="fixture-key"),
            patch.object(app, "sync", side_effect=cycles),
        ):
            jobs.run(job, self.cfg)
        state = jobs.status(self.cfg, job)
        self.assertEqual(state["state"], "completed")
        self.assertEqual(state["cycles"], 2)
        self.assertEqual(state["known_folders"], 2)
        self.assertEqual(state["completed_folders"], 2)

    def test_missing_inventory_stops_even_when_other_folders_keep_scanning(self):
        inbox = dict(folder="inbox", folder_id="inbox-id")
        archive = dict(folder="archive", folder_id="archive-id")
        cycle = dict(
            folders=[dict(inbox, window_complete=True)],
            pending_folders=[],
            scanned_total=1,
        )
        responses = [
            dict(cycle, target_folders=[inbox, archive], pending_folders=["archive"]),
            dict(cycle, target_folders=[inbox]),
            dict(cycle, stop_reason="cancelled"),
        ]
        with (
            patch.object(app, "api_key", return_value="fixture-key"),
            patch.object(jobs, "launch_worker"),
        ):
            job = jobs.start(self.cfg)["job_id"]
        with (
            patch.object(app, "api_key", return_value="fixture-key"),
            patch.object(app, "sync", side_effect=responses) as sync,
        ):
            jobs.run(job, self.cfg)
        state = jobs.status(self.cfg, job)
        self.assertEqual(state["state"], "completed_with_errors")
        self.assertEqual(state["reason"], "folder_unavailable")
        self.assertEqual(sync.call_count, 2)


if __name__ == "__main__":
    unittest.main()
