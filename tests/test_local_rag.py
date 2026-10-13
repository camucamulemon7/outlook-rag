"""Real local databases exercise publication, recovery and query behavior."""

from array import array
from contextlib import contextmanager
import hashlib
import io
import json
import os
from pathlib import Path
import sqlite3
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

import lancedb
import pyarrow as pa

from outlook_rag import app, cli, search, storage
import evaluate_search


class LocalRag(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.temp.cleanup)
        self.cfg = dict(
            model="fixture",
            dimensions=4,
            embedding_url="http://fixture.invalid",
            data_dir=str(Path(self.temp.name) / "index"),
            embedding_batch_size=8,
            embedding_concurrent_requests=2,
            sync_batch_emails=16,
            text_cleaning_version=2,
        )
        self.configuration = patch.object(app, "config", return_value=self.cfg)
        self.configuration.start()
        self.addCleanup(self.configuration.stop)
        self.embedding = patch.object(
            app,
            "embed",
            side_effect=lambda texts, *a, **k: [[1.0, 0.0, 0.0, 0.0] for _ in texts],
        )
        self.embedding.start()
        self.addCleanup(self.embedding.stop)
        self.sql, self.table = app.open_store(self.cfg)
        self.addCleanup(self.sql.close)

    def mail(self, key, **values):
        return dict(
            dict(
                id=hashlib.sha256(key.encode()).hexdigest(),
                entry_id=key,
                store_id="store",
                folder="inbox",
                sender="Sender@example.invalid",
                received="2026-10-13T12:00:00+09:00",
                subject=key,
                body="synthetic body",
                conversation_id="thread",
            ),
            **values,
        )

    def index(self, *mails):
        return app.upsert_mails(list(mails), self.sql, self.table, self.cfg)

    def test_native_filter_preserves_recall_without_id_shards(self):
        mails = [
            self.mail(str(i), folder="in'box", body="eligible") for i in range(1001)
        ]
        mails.append(self.mail("excluded", folder="archive", body="excluded"))
        self.index(*mails)
        with (
            patch.object(
                search, "vector_candidates", wraps=search.vector_candidates
            ) as vector,
            patch.object(
                search, "id_filter", side_effect=AssertionError("ID shards used")
            ),
        ):
            result = app.search(
                "question",
                folder="in'box",
                sender="SENDER@",
                since="2026-10-13",
                until="2026-10-13",
                mode="semantic",
                limit=5,
            )
        self.assertEqual(result["count"], 5)
        self.assertTrue(result["metadata_prefilter"])
        self.assertTrue(all(row["folder"] == "in'box" for row in result["items"]))
        self.assertEqual(vector.call_count, 1)
        self.assertIsNone(vector.call_args.args[3])
        self.assertIn("in''box", vector.call_args.kwargs["predicate"])

    def test_metadata_only_update_and_move_do_not_reembed(self):
        mail = self.mail("mail")
        self.index(mail)
        updated = dict(
            mail,
            folder="archive",
            sender="OTHER@example.invalid",
            received="2026-09-01",
        )
        stats = self.index(updated)
        self.assertEqual(stats["embedded_chunks"], 0)
        self.assertEqual(
            app.search("question", folder="inbox", mode="semantic")["count"], 0
        )
        self.assertEqual(
            app.search(
                "question",
                folder="archive",
                sender="other@",
                until="2026-09-01",
                mode="semantic",
            )["count"],
            1,
        )
        storage.update_folder(
            self.sql, self.table, "renamed", store_id="store", old_folder="archive"
        )
        self.assertEqual(
            app.search("question", folder="renamed", mode="semantic")["count"], 1
        )
        storage.update_folder(self.sql, self.table, "moved", mail_id=mail["id"])
        self.assertEqual(
            app.search("question", folder="moved", mode="semantic")["count"], 1
        )

    def test_failed_metadata_publication_preserves_old_filter(self):
        mail = self.mail("mail")
        self.index(mail)
        with self.sql:
            self.sql.execute(
                "CREATE TRIGGER fail_meta BEFORE INSERT ON emails BEGIN SELECT RAISE(ABORT, 'fixture'); END"
            )
        with self.assertRaises(sqlite3.IntegrityError):
            self.index(dict(mail, folder="archive"))
        self.assertEqual(
            app.search("question", folder="inbox", mode="semantic")["count"], 1
        )
        self.assertEqual(
            app.search("question", folder="archive", mode="semantic")["count"], 0
        )

    def test_null_date_and_unicode_sender_match_sql_filter_semantics(self):
        self.index(
            self.mail("dated", sender="ÄBC@EXAMPLE.INVALID"),
            self.mail("undated", received=None),
        )
        result = app.search("question", until="2026-10-13", mode="semantic")
        self.assertEqual(result["count"], 1)
        self.assertEqual(
            app.search("question", sender="Äbc@", mode="semantic")["count"], 1
        )
        self.assertEqual(
            app.search("question", sender="äbc@", mode="semantic")["count"], 0
        )

    def test_single_identifier_is_exact_offline_and_miss_returns_zero(self):
        target = self.mail("hotel", body="予約番号 HT-5639 です")
        self.index(
            target,
            self.mail("flight", body="予約番号 HN-5639 です"),
            self.mail("suffix", body="番号 HT-56390 です"),
        )
        with patch.object(app, "embed", side_effect=AssertionError("unexpected API")):
            found = app.search("ht-5639 の予約", mode="auto")
            self.assertEqual([row["mail_id"] for row in found["items"]], [target["id"]])
            self.assertEqual(found["search_mode"], "identifier")
            self.assertEqual(app.search("HT-9999", mode="auto")["count"], 0)
            self.assertEqual(
                app.search("HT-5639", folder="archive", mode="identifier")["count"], 0
            )
            self.assertEqual(app.search("予約", mode="lexical")["count"], 2)
        self.assertEqual(
            search.resolve_mode("HT-5639 と HN-5639 の比較", "auto")[0], "semantic"
        )
        self.assertEqual(
            search.resolve_mode("HT-5639 以外の予約", "auto")[0], "semantic"
        )
        with self.assertRaises(ValueError):
            app.search("HT-5639 HN-5639", mode="identifier")
        self.index(dict(target, body="different identifier HT-7788"))
        self.assertEqual(app.search("HT-5639", mode="identifier")["count"], 0)
        app.remove_mail(target["id"], self.sql, self.table)
        self.assertEqual(app.search("HT-7788", mode="identifier")["count"], 0)

    def test_search_read_path_does_not_initialize_or_wait_for_sql_writer(self):
        self.index(self.mail("mail"))
        app.search("cached", mode="semantic")
        self.sql.execute("BEGIN IMMEDIATE")
        try:
            with (
                patch.object(
                    app, "open_store", side_effect=AssertionError("read attempted DDL")
                ),
                patch.object(
                    app, "embed", side_effect=AssertionError("cached query called API")
                ),
            ):
                started = time.perf_counter()
                self.assertEqual(app.search("cached", mode="semantic")["count"], 1)
                self.assertLess(time.perf_counter() - started, 2)
                status = app.status()
                self.assertEqual(status["emails"], 1)
        finally:
            self.sql.rollback()

    def test_candidate_metadata_and_thread_counts_are_batched(self):
        self.index(*[self.mail(str(i), conversation_id=str(i)) for i in range(30)])
        statements = []
        original = storage.open_index

        def tracked(*args, **kwargs):
            sql, table = original(*args, **kwargs)
            sql.set_trace_callback(statements.append)
            return sql, table

        with patch.object(storage, "open_index", side_effect=tracked):
            result = app.search(
                "question", mode="semantic", group_by_thread=True, limit=10
            )
        self.assertEqual(result["count"], 10)
        self.assertTrue(all(row["thread_mail_count"] == 1 for row in result["items"]))
        self.assertEqual(sum("AS chunk_id" in stmt for stmt in statements), 1)
        self.assertEqual(
            sum("GROUP BY store_id,conversation_id" in stmt for stmt in statements), 1
        )
        self.assertFalse(any("SELECT e.*" in stmt for stmt in statements))

    def test_lexical_limit_is_applied_after_date_and_archive_filters(self):
        self.index(
            *[
                self.mail(
                    f"old-{i}", body="needle " * 20,
                    folder="archive", received="2008-06-01",
                )
                for i in range(110)
            ],
            self.mail("inbox", body="needle " * 20),
            self.mail("weak", folder="archive", body="needle unrelated words"),
            self.mail("strong", folder="archive", body="needle " * 10),
        )
        result = app.search(
            "needle", mode="lexical", folder="archive",
            since="2026-01-01", until="2026-12-31", limit=2,
        )
        self.assertEqual(
            [row["entry_id"] for row in result["items"]], ["strong", "weak"]
        )
        # The default search still reaches the older archive without dates.
        all_dates = app.search("needle", mode="lexical", limit=50)
        self.assertTrue(any(row["received"] == "2008-06-01" for row in all_dates["items"]))

    def test_lexical_bm25_survives_custom_fts_rank_and_orphan_rows(self):
        self.index(
            self.mail("weak", body="needle unrelated words"),
            self.mail("strong", body="needle " * 10),
            self.mail("noise", body="unrelated words"),
        )
        with self.sql:
            # An old/custom FTS rank mapping must not change the search scorer.
            self.sql.execute(
                "INSERT INTO lexical(lexical,rank) VALUES ('rank','bm25(99.0,0.0)')"
            )
            self.sql.executemany(
                "INSERT INTO lexical VALUES (?,?)",
                [(f"orphan-{i}", "needle " * 30) for i in range(110)],
            )
        result = app.search("needle", mode="lexical", limit=2)
        self.assertEqual(
            [row["entry_id"] for row in result["items"]], ["strong", "weak"]
        )

    def test_rebuild_migrates_legacy_schema_and_repairs_lexical_without_api(self):
        target = self.mail("target", body="予約 HT-5639")
        self.index(target)
        self.table.checkout_latest()
        self.table.drop_columns(sorted(storage.METADATA_COLUMNS))
        with self.sql:
            app.publish_vectors(self.sql, self.table)
            self.sql.execute(
                "DELETE FROM state WHERE key IN ('vector_metadata','identifiers_ready')"
            )
            self.sql.execute("DELETE FROM identifiers")
            self.sql.execute("DELETE FROM lexical")
        self.assertEqual(app.search("HT-5639", mode="identifier")["count"], 1)
        with patch.object(
            app, "embed", side_effect=AssertionError("rebuild called API")
        ):
            report = storage.rebuild_vectors(self.cfg)
            self.assertEqual(report["embedding_requests"], 0)
            self.assertTrue(storage.doctor(self.cfg)["healthy"])
            result = app.search("予約", mode="lexical", folder="inbox")
            self.assertEqual(result["count"], 1)
            self.assertTrue(result["metadata_prefilter"])
            self.assertEqual(app.search("HT-5639", mode="identifier")["count"], 1)

    def test_invalid_cache_refuses_rebuild_and_keeps_publication(self):
        self.index(self.mail("mail"))
        previous = storage.table_name(self.sql)
        for blob in (
            b"short",
            array("f", [float("nan"), 0.0, 0.0, 0.0]).tobytes(),
            array("f", [0.0] * 4).tobytes(),
        ):
            with self.sql:
                self.sql.execute("UPDATE embedding_cache SET vector=?", (blob,))
            with self.assertRaises(ValueError):
                storage.rebuild_vectors(self.cfg)
            self.assertEqual(storage.table_name(self.sql), previous)
            self.assertEqual(app.search("question", mode="semantic")["count"], 1)
        with self.sql:
            self.sql.execute("DELETE FROM embedding_cache")
        with self.assertRaises(ValueError):
            storage.rebuild_vectors(self.cfg)
        self.assertFalse(storage.doctor(self.cfg)["cached_rebuild_available"])
        self.assertEqual(
            lancedb.connect(str(Path(self.cfg["data_dir"]) / "vectors"))
            .list_tables()
            .tables,
            ["mail_chunks"],
        )

    def test_rebuild_sql_failure_leaves_old_generation_readable(self):
        self.index(self.mail("mail"))
        with self.sql:
            self.sql.execute(
                "CREATE TRIGGER fail_rebuild BEFORE INSERT ON state WHEN NEW.key='vector_table' BEGIN SELECT RAISE(ABORT, 'fixture'); END"
            )
        with self.assertRaises(sqlite3.IntegrityError):
            storage.rebuild_vectors(self.cfg)
        self.assertEqual(storage.table_name(self.sql), "mail_chunks")
        self.assertEqual(app.search("question", mode="semantic")["count"], 1)
        self.assertTrue(storage.doctor(self.cfg)["healthy"])

    def test_interruption_after_rebuild_commit_keeps_published_table(self):
        self.index(self.mail("mail"))
        original = app.reader_lock

        @contextmanager
        def interrupted(cfg):
            with original(cfg):
                yield
            raise RuntimeError("interrupted just after commit")

        with (
            patch.object(app, "reader_lock", side_effect=interrupted),
            self.assertRaises(RuntimeError),
        ):
            storage.rebuild_vectors(self.cfg)
        self.assertNotEqual(storage.table_name(self.sql), "mail_chunks")
        self.assertTrue(storage.doctor(self.cfg)["healthy"])
        self.assertEqual(app.search("synthetic", mode="lexical")["count"], 1)

    def test_doctor_finds_metadata_mismatch_and_rebuild_repairs_it(self):
        self.index(self.mail("mail"))
        with self.sql:
            self.sql.execute("UPDATE emails SET folder='archive'")
        self.assertIn("vector_metadata", storage.doctor(self.cfg)["problems"])
        storage.rebuild_vectors(self.cfg)
        self.assertTrue(storage.doctor(self.cfg)["healthy"])
        self.assertEqual(
            app.search("question", folder="archive", mode="semantic")["count"], 1
        )

    def test_rebuild_while_embedding_refreshes_reader_table(self):
        self.index(self.mail("mail"))

        def rebuild_and_embed(*args, **kwargs):
            storage.rebuild_vectors(self.cfg)
            app.maintain()
            return [[1.0, 0.0, 0.0, 0.0]]

        with patch.object(app, "embed", side_effect=rebuild_and_embed):
            self.assertEqual(app.search("question", mode="semantic")["count"], 1)
        self.assertNotEqual(storage.table_name(self.sql), "mail_chunks")

    def test_rebuild_publication_waits_for_active_reader_snapshot(self):
        self.index(self.mail("mail"))
        reader, table = storage.open_index(self.cfg)
        self.addCleanup(reader.close)
        waiting = threading.Event()
        errors = []
        original = app.reader_lock

        def observed(cfg):
            if threading.current_thread().name == "rebuild-fixture":
                waiting.set()
            return original(cfg)

        def rebuild():
            try:
                storage.rebuild_vectors(self.cfg)
            except BaseException as exc:
                errors.append(exc)

        worker = threading.Thread(target=rebuild, name="rebuild-fixture", daemon=True)
        with patch.object(app, "reader_lock", side_effect=observed):
            with app.indexed_snapshot(reader, table, self.cfg) as table:
                worker.start()
                self.assertTrue(waiting.wait(5))
                self.assertTrue(worker.is_alive())
                self.assertEqual(storage.table_name(reader), "mail_chunks")
                self.assertEqual(table.count_rows(), 1)
            worker.join(5)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertNotEqual(storage.table_name(reader), "mail_chunks")
        self.assertTrue(storage.doctor(self.cfg)["healthy"])

    def test_backup_restore_pins_committed_publication_and_refuses_overwrite(self):
        mail = self.mail("mail")
        self.index(mail)
        # An unpublished Lance update must not appear in a restored backup.
        self.table.checkout_latest()
        self.table.delete(f"email_id='{mail['id']}'")
        backup, restored = (
            Path(self.temp.name) / "backup",
            Path(self.temp.name) / "restored",
        )
        with patch.object(
            app, "embed", side_effect=AssertionError("offline maintenance called API")
        ):
            self.assertTrue(storage.backup_index(self.cfg, str(backup))["backed_up"])
            self.assertTrue(
                storage.restore_index(self.cfg, str(backup), str(restored))["restored"]
            )
        with patch.object(
            app, "config", return_value=dict(self.cfg, data_dir=str(restored))
        ):
            self.assertEqual(app.search("synthetic", mode="lexical")["count"], 1)
        with self.assertRaises(FileExistsError):
            storage.backup_index(self.cfg, str(backup))
        with self.assertRaises(ValueError):
            storage.restore_index(self.cfg, str(backup), str(restored))
        with self.assertRaises(ValueError):
            storage.backup_index(self.cfg, str(Path(self.cfg["data_dir"]) / "backup"))
        self.assertTrue((backup / "metadata.sqlite").is_file())

    def test_backup_restore_after_rebuild_and_missing_vector_recovery(self):
        self.index(self.mail("mail"))
        storage.rebuild_vectors(self.cfg)
        backup, restored = (
            Path(self.temp.name) / "backup",
            Path(self.temp.name) / "restored",
        )
        storage.backup_index(self.cfg, str(backup))
        storage.restore_index(self.cfg, str(backup), str(restored))
        self.assertTrue(
            storage.doctor(dict(self.cfg, data_dir=str(restored)))["healthy"]
        )
        db = lancedb.connect(str(Path(self.cfg["data_dir"]) / "vectors"))
        db.drop_table(storage.table_name(self.sql))
        report = storage.doctor(self.cfg)
        self.assertIn("vector_publication_unavailable", report["problems"])
        self.assertTrue(report["cached_rebuild_available"])
        storage.rebuild_vectors(self.cfg)
        self.assertTrue(storage.doctor(self.cfg)["healthy"])

    def test_doctor_reports_corruption_without_changing_publication(self):
        self.index(self.mail("mail"))
        before = self.sql.execute(
            "SELECT value FROM state WHERE key='vector_version'"
        ).fetchone()[0]
        self.table.checkout_latest()
        self.table.delete("1=1")
        with self.sql:
            app.publish_vectors(self.sql, self.table)
        report = storage.doctor(self.cfg)
        self.assertFalse(report["healthy"])
        self.assertIn("vector_count", report["problems"])
        self.assertIn("vector_ids", report["problems"])
        self.assertNotEqual(
            before,
            self.sql.execute(
                "SELECT value FROM state WHERE key='vector_version'"
            ).fetchone()[0],
        )
        self.assertEqual(self.table.count_rows(), 0)

    def test_ann_controls_and_exact_baseline(self):
        self.index(*[self.mail(str(i)) for i in range(256)])
        app.optimize()
        self.assertEqual(app.status()["vector_coverage"]["indexed_rows"], 256)
        self.sql.close()
        self.sql, self.table = app.open_store(self.cfg)
        self.addCleanup(self.sql.close)
        self.index(self.mail("new-row"))
        self.assertEqual(app.status()["vector_coverage"]["unindexed_rows"], 1)
        self.cfg.update(search_nprobes=1, search_ef=10)
        rows = search.vector_candidates(
            self.table, [1.0, 0.0, 0.0, 0.0], 100, None, cfg=self.cfg
        )
        exact = search.vector_candidates(
            self.table,
            [1.0, 0.0, 0.0, 0.0],
            100,
            None,
            cfg=dict(self.cfg, search_exact=True),
        )
        self.assertEqual(len(rows), 100)
        self.assertEqual(len(exact), 100)
        self.assertTrue(all(abs(row["_distance"]) < 1e-5 for row in exact))

    def test_cli_offline_commands_do_not_require_key_or_api(self):
        self.index(self.mail("mail"))
        for command in (
            ["doctor"],
            ["rebuild-vectors"],
            ["search", "synthetic", "--mode", "lexical"],
        ):
            with (
                patch("sys.argv", ["outlook-rag", *command]),
                patch("sys.stdout", io.StringIO()) as output,
                patch.object(
                    app, "api_key", side_effect=AssertionError("unexpected key request")
                ),
                patch.object(
                    app, "embed", side_effect=AssertionError("unexpected API")
                ),
            ):
                cli.main(app)
                self.assertIsInstance(json.loads(output.getvalue()), dict)

    def test_evaluation_cold_warm_and_ann_are_separate(self):
        mail = self.mail("mail")
        self.index(mail)
        cases = [
            dict(query="synthetic", expected_mail_ids=[mail["id"]]),
            dict(query="HT-9999", expected_mail_ids=[]),
        ]
        cold = evaluate_search.evaluate(cases, mode="auto", cache="cold")
        warm = evaluate_search.evaluate(cases, mode="auto", cache="warm")
        self.assertEqual(cold["query_cache_hits"], 0)
        self.assertEqual(warm["query_cache_hits"], 1)
        self.assertEqual(warm["hit_at_1"], 1)
        self.assertEqual(warm["no_answer_rate"], 1)
        ann = evaluate_search.evaluate_ann(cases)
        self.assertEqual(ann["chunk_recall"], 1)
        self.assertIn("exact_median_ms", ann)

    def test_evaluation_metric_denominators_and_no_answer(self):
        cases = [
            dict(expected_mail_ids=["a", "b"]),
            dict(expected_mail_ids=[]),
            dict(expected_mail_ids=["c"]),
        ]
        results = [
            dict(items=[dict(mail_id="z"), dict(mail_id="a")]),
            dict(items=[]),
            dict(items=[]),
        ]
        measured = evaluate_search.metrics(cases, results, [2.0, 3.0, 4.0])
        self.assertEqual(measured["hit_at_1"], 0)
        self.assertEqual(measured["hit_at_5"], 0.5)
        self.assertEqual(measured["mrr_at_10"], 0.25)
        self.assertEqual(measured["labeled_recall_at_10"], 0.25)
        self.assertEqual(measured["no_answer_rate"], 1)


class SearchConfiguration(unittest.TestCase):
    def test_modes_and_ann_validation_do_not_change_embedding_identity(self):
        env = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RAG_")}
        env.update(OUTLOOK_RAG_MODEL="fixture", OUTLOOK_RAG_DIMENSIONS="4")
        with patch.dict(os.environ, env, clear=True):
            base = app.config()
            with patch.dict(
                os.environ,
                dict(
                    OUTLOOK_RAG_SEARCH_MODE="auto",
                    OUTLOOK_RAG_SEARCH_NPROBES="32",
                    OUTLOOK_RAG_SEARCH_EF="200",
                    OUTLOOK_RAG_SEARCH_EXACT="true",
                ),
            ):
                changed = app.config()
            self.assertEqual(app.index_identity(base), app.index_identity(changed))
            self.assertTrue(changed["search_exact"])
            for key, value in (
                ("SEARCH_MODE", "bad"),
                ("SEARCH_NPROBES", "0"),
                ("SEARCH_EF", "-1"),
                ("SEARCH_EXACT", "1"),
            ):
                with (
                    patch.dict(os.environ, {"OUTLOOK_RAG_" + key: value}),
                    self.assertRaises(ValueError),
                ):
                    app.config()


if __name__ == "__main__":
    unittest.main()
