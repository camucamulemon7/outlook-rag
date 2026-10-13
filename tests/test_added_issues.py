"""Regressions for additional search, worker diagnostics and CLI issues."""

import asyncio
from contextlib import redirect_stdout
import hashlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from outlook_rag import app, diagnostics, jobs, search, worker
import test_improvements as fixtures


class AddedIssues(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.cfg, self.sql, self.table = (
            self.fixture.cfg,
            self.fixture.sql,
            self.fixture.table,
        )
        self.configuration = patch.object(app, "config", return_value=self.cfg)
        self.configuration.start()

    def tearDown(self):
        self.configuration.stop()
        self.fixture.tearDown()

    def mail(self, i, body="body", **values):
        return dict(
            dict(
                id=hashlib.sha256(str(i).encode()).hexdigest(),
                entry_id=str(i),
                store_id="store",
                folder="inbox",
                subject="fixture",
                body=body,
                sender="sender@example.invalid",
                recipients="recipient",
                received="2026-09-05T00:00:00+00:00",
                modified="1",
                conversation_id="thread",
            ),
            **values,
        )

    def start(self):
        with patch.object(jobs, "launch_worker"):
            return jobs.start(self.cfg)["job_id"]

    def test_large_filter_preserves_recall_and_bounds_every_expression(self):
        mails = [self.mail(i, "weak") for i in range(1000)] + [
            self.mail(1000, "strong")
        ]
        # Unfiltered candidates would be filled with stronger, excluded mail.
        mails += [self.mail(i + 2000, "outside", folder="archive") for i in range(120)]

        def vectors(texts, *_args, **kwargs):
            return [
                [1.0, 0.0, 0.0, 0.0]
                if kwargs.get("query") or "outside" in t
                else [0.8, 0.6, 0.0, 0.0]
                if "strong" in t
                else [0.0, 1.0, 0.0, 0.0]
                for t in texts
            ]

        with patch.object(app, "embed", side_effect=vectors):
            app.upsert_mails(mails, self.sql, self.table, self.cfg)
            # Legacy tables lack native metadata; keep the bounded fallback covered.
            with self.sql:
                self.sql.execute("DELETE FROM state WHERE key='vector_metadata'")
            expressions = []
            original = search.id_filter

            def capture(column, ids):
                value = original(column, ids)
                expressions.append((len(ids), len(value)))
                return value

            with patch.object(search, "id_filter", side_effect=capture):
                result = app.search(
                    "query",
                    limit=5,
                    folder="inbox",
                    sender="sender@",
                    since="2026-09-01",
                    until="2026-09-30",
                    hybrid=False,
                )
        self.assertEqual(result["count"], 5)
        self.assertEqual(result["items"][0]["mail_id"], mails[1000]["id"])
        self.assertTrue(all(row["folder"] == "inbox" for row in result["items"]))
        self.assertEqual(len(expressions), 2)
        self.assertLessEqual(max(n for n, _ in expressions), 1000)
        self.assertLessEqual(max(n for _, n in expressions), 67013)

    def test_similarity_threshold_filters_without_changing_default(self):
        def vectors(texts, *_args, **kwargs):
            return [
                [1.0, 0.0, 0.0, 0.0]
                if kwargs.get("query") or "relevant" in t
                else [0.0, 1.0, 0.0, 0.0]
                for t in texts
            ]

        with patch.object(app, "embed", side_effect=vectors):
            app.upsert_mails(
                [self.mail(1, "relevant"), self.mail(2, "noise")],
                self.sql,
                self.table,
                self.cfg,
            )
            self.assertEqual(app.search("query", hybrid=False)["count"], 2)
            filtered = app.search("query", hybrid=False, min_similarity=0.5)
            self.assertEqual(filtered["count"], 1)
            self.assertEqual(filtered["items"][0]["mail_id"], self.mail(1)["id"])
            self.assertEqual(
                app.search("query", folder="archive", min_similarity=0.5)["count"], 0
            )
        for invalid in (float("nan"), float("inf"), -1.01, 1.01, True):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                app.search("query", min_similarity=invalid)

    def test_bm25_only_candidate_gets_real_similarity_for_threshold(self):
        def vectors(texts, *_args, **kwargs):
            return [
                [1.0, 0.0, 0.0, 0.0]
                if kwargs.get("query")
                else [0.75, 0.6614378278, 0.0, 0.0]
                if "keyword" in t
                else [0.8, 0.6, 0.0, 0.0]
                for t in texts
            ]

        with patch.object(app, "embed", side_effect=vectors):
            app.upsert_mails(
                [self.mail(i, "unrelated") for i in range(105)]
                + [self.mail(999, "keyword")],
                self.sql,
                self.table,
                self.cfg,
            )
            self.cfg.update(semantic_weight=0.0, lexical_weight=1.0)
            hit = app.search("keyword", limit=1, min_similarity=0.7)
            self.assertEqual(hit["items"][0]["mail_id"], self.mail(999)["id"])
            self.assertAlmostEqual(hit["items"][0]["semantic_similarity"], 0.75)
            self.assertEqual(app.search("keyword", min_similarity=0.9)["count"], 0)

    def test_long_chunk_snippet_includes_late_match_and_mcp_exposes_threshold(self):
        self.sql.close()
        self.cfg.update(
            chunk_size=16000, data_dir=str(Path(self.fixture.temp.name) / "long-chunks")
        )
        self.sql, self.table = app.open_store(self.cfg)
        self.fixture.sql, self.fixture.table = self.sql, self.table
        app.upsert_mail(
            self.mail(1, "x" * 8000 + " 運賃改定 " + "y" * 3000),
            self.sql,
            self.table,
            self.cfg,
        )
        result = app.search("運賃改定")
        self.assertIn("運賃改定", result["items"][0]["snippet"])
        self.assertLessEqual(len(result["items"][0]["snippet"]), 1802)
        schemas = asyncio.run(app.mcp.list_tools())
        tool = next(tool for tool in schemas if tool.name == "search_emails")
        self.assertIn("min_similarity", tool.inputSchema["properties"])

    def test_unicode_case_expansion_keeps_snippet_offsets_in_original_text(self):
        text = "İ" * 3000 + " cargochange " + "X" * 3000
        snippet = search.contextual_snippet(text, "cargochange", [])
        self.assertIn("cargochange", snippet)
        self.assertLessEqual(len(snippet), 1802)

    def test_zero_weight_sources_never_add_unmatched_results_or_provider_calls(self):
        app.upsert_mail(self.mail(1, "cached body"), self.sql, self.table, self.cfg)
        self.cfg.update(semantic_weight=0.0, lexical_weight=1.0)
        self.assertEqual(app.search("notpresent")["count"], 0)
        with patch.object(
            app,
            "embed",
            side_effect=AssertionError(
                "Disabled semantic source must not call provider"
            ),
        ):
            self.assertEqual(app.search("body")["count"], 1)
        self.cfg.update(semantic_weight=1.0, lexical_weight=0.0)
        with patch.object(
            app,
            "tokens",
            side_effect=AssertionError("Disabled lexical source must not tokenize"),
        ):
            self.assertEqual(app.search("query")["count"], 1)

    def test_empty_metadata_filter_returns_without_query_embedding(self):
        app.upsert_mail(self.mail(1), self.sql, self.table, self.cfg)
        with patch.object(
            app,
            "embed",
            side_effect=AssertionError("Empty filtered index needs no provider"),
        ):
            self.assertEqual(app.search("unseen", folder="absent")["count"], 0)

    def test_wmi_return_code_is_retained_in_safe_launch_diagnostics(self):
        arguments = SimpleNamespace()
        startup = SimpleNamespace()
        process = SimpleNamespace(
            Methods_=lambda _: SimpleNamespace(
                InParameters=SimpleNamespace(SpawnInstance_=lambda: arguments)
            ),
            ExecMethod_=lambda *_: SimpleNamespace(ReturnValue=9),
        )
        wmi = SimpleNamespace(
            Get=lambda name: (
                SimpleNamespace(SpawnInstance_=lambda: startup)
                if name == "Win32_ProcessStartup"
                else process
            )
        )
        platform = SimpleNamespace(
            name="nt",
            urandom=lambda _: b"x" * 32,
            environ=os.environ,
            replace=os.replace,
        )
        with (
            patch.dict(os.environ, {"OUTLOOK_RAG_API_KEY": "fixture-secret"}),
            patch.object(jobs, "os", platform),
            patch("win32com.client.GetObject", return_value=wmi, create=True),
        ):
            with self.assertRaises(OSError):
                jobs.start(self.cfg)
        state = jobs.status(self.cfg)
        self.assertEqual(state["state"], "failed")
        self.assertIn("WMI", state["error_detail"])
        self.assertIn("9", state["error_detail"])
        self.assertNotIn("fixture-secret", json.dumps(state))

    def test_worker_failure_retains_safe_stack_and_recent_log(self):
        job_id = self.start()
        with (
            patch.object(app, "api_key", return_value="fixture-secret"),
            patch.object(
                app,
                "sync",
                side_effect=RuntimeError("fixture-secret private-mail-body"),
            ),
        ):
            jobs.run(job_id, self.cfg)
        state = jobs.status(self.cfg, job_id)
        self.assertEqual(state["state"], "failed")
        self.assertEqual(state["error_type"], "RuntimeError")
        self.assertTrue(state["error_detail"])
        logs = "\n".join(state["recent_logs"])
        self.assertIn("worker_failed", logs)
        self.assertIn('"stack"', logs)
        self.assertIn("jobs.py", logs)
        self.assertNotIn("fixture-secret", logs)
        self.assertNotIn("private-mail-body", logs)
        self.assertTrue(Path(state["log_path"]).is_file())

    def test_launcher_failure_is_diagnosed_without_raw_error_text(self):
        with patch.object(
            jobs, "launch_worker", side_effect=PermissionError("fixture-secret")
        ):
            with self.assertRaises(PermissionError):
                jobs.start(self.cfg)
        state = jobs.status(self.cfg)
        self.assertEqual(state["state"], "failed")
        self.assertIn("launch_failed", "\n".join(state["recent_logs"]))
        self.assertNotIn("fixture-secret", json.dumps(state))

    def test_real_bootstrap_captures_import_crash_before_app_loads(self):
        job_id = self.start()
        snapshot = jobs.root(self.cfg) / (job_id + ".config.json")
        with tempfile.TemporaryDirectory() as root:
            package = Path(root) / "outlook_rag"
            package.mkdir()
            (package / "__init__.py").write_text("", encoding="utf-8")
            (package / "app.py").write_text(
                "raise ImportError('fixture-secret private-mail-body')",
                encoding="utf-8",
            )
            environment = dict(
                os.environ, PYTHONPATH=root, OUTLOOK_RAG_API_KEY="fixture-secret"
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(Path(worker.__file__).resolve()),
                    "--config",
                    str(snapshot.resolve()),
                    "_job-worker",
                    job_id,
                ],
                env=environment,
                capture_output=True,
                text=True,
                timeout=15,
            )
        self.assertEqual(result.returncode, 1, result.stderr)
        state = jobs.status(self.cfg, job_id)
        self.assertEqual(state["state"], "failed")
        self.assertEqual(state["error_type"], "ImportError")
        self.assertIn("bootstrap_failed", "\n".join(state["recent_logs"]))
        self.assertNotIn(
            "fixture-secret", json.dumps(state) + result.stdout + result.stderr
        )
        self.assertNotIn("private-mail-body", json.dumps(state))

    def test_wmi_bootstrap_environment_transfer_and_diagnostics(self):
        job_id = self.start()
        snapshot = jobs.root(self.cfg) / (job_id + ".config.json")
        connection = Mock()
        connection.__enter__ = Mock(return_value=connection)
        connection.__exit__ = Mock(return_value=False)
        connection.poll.return_value = True
        connection.recv.return_value = {"OUTLOOK_RAG_API_KEY": "fixture-secret"}
        listener = Mock()
        listener.__enter__ = Mock(return_value=listener)
        listener.__exit__ = Mock(return_value=False)
        listener.accept.return_value = connection
        argv = [
            "worker",
            "--pipe",
            "synthetic-pipe",
            "--authkey",
            "00" * 32,
            "--config",
            str(snapshot.resolve()),
            "_job-worker",
            job_id,
        ]
        with (
            patch.dict(os.environ, {}, clear=True),
            patch.object(sys, "argv", argv),
            patch(
                "multiprocessing.connection.Listener", return_value=listener
            ) as factory,
            patch.object(jobs, "run") as run,
        ):
            self.assertEqual(worker.main(), 0)
            self.assertEqual(os.environ["OUTLOOK_RAG_API_KEY"], "fixture-secret")
        factory.assert_called_once_with(
            "synthetic-pipe", family="AF_PIPE", authkey=b"\0" * 32
        )
        run.assert_called_once_with(job_id, self.cfg)
        logs = "\n".join(jobs.status(self.cfg, job_id)["recent_logs"])
        self.assertIn("bootstrap_ready", logs)
        self.assertNotIn("fixture-secret", logs)

    def test_both_launch_paths_use_diagnostic_bootstrap_and_keep_key_in_memory(self):
        from unittest.mock import MagicMock

        command = [
            sys.executable,
            "-m",
            "outlook_rag.app",
            "--config",
            "snapshot.json",
            "_job-worker",
            "a" * 32,
        ]
        environment = {"OUTLOOK_RAG_API_KEY": "fixture-secret"}
        with (
            patch.object(jobs, "os", SimpleNamespace(name="posix")),
            patch.object(jobs.subprocess, "Popen") as popen,
        ):
            jobs.launch_worker(command, env=environment)
        invoked = popen.call_args.args[0]
        self.assertTrue(invoked[1].endswith("worker.py"))
        self.assertNotIn("fixture-secret", repr(invoked))
        self.assertEqual(popen.call_args.kwargs["env"], environment)
        self.assertTrue(popen.call_args.kwargs["start_new_session"])

        arguments = SimpleNamespace()
        startup = SimpleNamespace()
        process = SimpleNamespace(
            Methods_=lambda _: SimpleNamespace(
                InParameters=SimpleNamespace(SpawnInstance_=lambda: arguments)
            ),
            ExecMethod_=lambda *_: SimpleNamespace(ReturnValue=0),
        )
        wmi = SimpleNamespace(
            Get=lambda name: (
                SimpleNamespace(SpawnInstance_=lambda: startup)
                if name == "Win32_ProcessStartup"
                else process
            )
        )
        connection = MagicMock()
        connection.__enter__.return_value = connection
        with (
            patch.object(
                jobs, "os", SimpleNamespace(name="nt", urandom=lambda _: b"x" * 32)
            ),
            patch("win32com.client.GetObject", return_value=wmi, create=True),
            patch("multiprocessing.connection.Client", return_value=connection),
        ):
            jobs.launch_worker(command, env=environment)
        self.assertIn("worker.py", arguments.CommandLine)
        self.assertIn("--pipe", arguments.CommandLine)
        self.assertNotIn("fixture-secret", arguments.CommandLine)
        self.assertEqual(startup.ShowWindow, 0)
        connection.send.assert_called_once_with(environment)

    def test_log_retention_protects_running_worker_and_bounds_file_size(self):
        old, active, recent = ["a" * 32, "b" * 32, "c" * 32]
        for job_id in (old, active, recent):
            diagnostics.append(self.cfg, job_id, "fixture")
        stale = time.time() - 15 * 86400
        for job_id in (old, active):
            os.utime(diagnostics.log_path(self.cfg, job_id), (stale, stale))
        jobs.write(jobs.root(self.cfg) / "worker.owner.json", {"job_id": active})
        with patch.object(jobs, "worker_running", return_value=True):
            self.assertEqual(jobs.prune_logs(self.cfg), 1)
        self.assertFalse(diagnostics.log_path(self.cfg, old).exists())
        self.assertTrue(diagnostics.log_path(self.cfg, active).exists())
        self.assertTrue(diagnostics.log_path(self.cfg, recent).exists())
        diagnostics.log_path(self.cfg, recent).write_text(
            "x\n" * (diagnostics.MAX_LOG_BYTES // 2 + 10), encoding="utf-8"
        )
        diagnostics.append(self.cfg, recent, "bounded")
        self.assertLess(
            diagnostics.log_path(self.cfg, recent).stat().st_size,
            diagnostics.MAX_LOG_BYTES,
        )

    def cli(self, arguments):
        environment = {
            key: value
            for key, value in os.environ.items()
            if not key.startswith("OUTLOOK_RAG_")
        }
        environment.update(
            OUTLOOK_RAG_MODEL="fixture",
            OUTLOOK_RAG_DATA_DIR=str(Path(self.cfg["data_dir"]).resolve()),
            OUTLOOK_RAG_EMBEDDING_URL="http://127.0.0.1:1/embeddings",
        )
        result = subprocess.run(
            [sys.executable, "-m", "outlook_rag.app", *arguments],
            env=environment,
            capture_output=True,
            text=True,
            timeout=15,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def test_cli_offline_job_failure_progress_and_thread_commands(self):
        mail = self.mail(1, "cached original")
        app.upsert_mail(mail, self.sql, self.table, self.cfg)
        app.record_sync_failure(
            self.sql,
            {"entry_id": "missing", "modified": "1"},
            "store",
            "inbox",
            ValueError("private"),
        )
        job_id = self.start()
        self.assertEqual(self.cli(["job", "status", job_id])["job_id"], job_id)
        self.assertTrue(self.cli(["job", "cancel", job_id])["cancel_requested"])
        self.assertEqual(
            self.cli(["failures", "list", "--folder", "inbox", "--limit", "1"])[
                "total"
            ],
            1,
        )
        self.assertEqual(self.cli(["progress"])["folders"][0]["indexed_emails"], 1)
        thread = self.cli(
            [
                "thread",
                mail["id"],
                "--before",
                "0",
                "--after",
                "0",
                "--max-body-chars",
                "6",
            ]
        )
        self.assertEqual(thread["items"][0]["body"], "cached")

    def test_cli_start_retry_and_search_dispatch_share_mcp_operations(self):
        for arguments, method, expected in (
            (
                [
                    "job",
                    "start",
                    "--folders",
                    "inbox",
                    "sent",
                    "--since-days",
                    "7",
                    "--reconcile",
                ],
                "start_sync_job",
                (["inbox", "sent"], 7, True),
            ),
            (
                ["failures", "retry", "a" * 64, "b" * 64, "--max-seconds", "12"],
                "retry_failed_emails",
                (["a" * 64, "b" * 64], 12.0),
            ),
            (
                [
                    "search",
                    "query",
                    "--folder",
                    "inbox",
                    "--min-similarity",
                    ".6",
                    "--semantic-only",
                ],
                "search",
                ("query", 10, "inbox", None, None, None, False, False, 0.6),
            ),
        ):
            output = io.StringIO()
            with (
                self.subTest(method=method),
                patch.object(sys, "argv", ["outlook-rag", *arguments]),
                patch.object(app, "api_key", return_value="fixture"),
                patch.object(app, method, return_value={"ok": True}) as call,
                redirect_stdout(output),
            ):
                app.main()
            call.assert_called_once_with(*expected)
            self.assertEqual(json.loads(output.getvalue()), {"ok": True})


if __name__ == "__main__":
    unittest.main()
