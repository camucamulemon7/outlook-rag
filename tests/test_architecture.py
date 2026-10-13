"""Attribute lookup, scoped chronological evidence and derived-thread recovery."""

import base64
import hashlib
import io
import json
import os
from pathlib import Path
import unittest
from unittest.mock import patch

from outlook_rag import app, cli, features, search, storage, text, threads
import test_local_rag


class Architecture(unittest.TestCase):
    def setUp(self):
        fixture = test_local_rag.LocalRag()
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        self.cfg, self.sql, self.table = fixture.cfg, fixture.sql, fixture.table
        self.mail, self.index = fixture.mail, fixture.index

    def test_metadata_is_exact_offline_and_keyset_pages_use_utc_instants(self):
        mails = [
            self.mail(
                "local-clock",
                received="2026-10-13T10:00:00+09:00",
                sender="Tanaka@example.invalid",
                recipients="team",
                subject="PCIe Gen7",
                folder="archive",
            ),
            self.mail(
                "utc-later",
                received="2026-10-13T03:00:00+00:00",
                sender="Tanaka@example.invalid",
                recipients="team",
                subject="PCIe Gen7",
                folder="archive",
            ),
            self.mail(
                "other-offset",
                received="2026-10-13T05:00:00+02:00",
                sender="Tanaka@example.invalid",
                recipients="team",
                subject="PCIe Gen7",
                folder="archive",
            ),
            self.mail(
                "old-archive",
                received="2008-06-01",
                sender="Tanaka@example.invalid",
                recipients="team",
                subject="PCIe Gen7",
                folder="archive",
            ),
            self.mail(
                "other-sender",
                sender="Other@example.invalid",
                recipients="team",
                subject="PCIe Gen7",
            ),
        ]
        self.index(*mails)
        options = dict(
            mode="metadata", sender="tanaka@", subject="pcie", recipient="team", limit=2
        )
        with (
            patch.object(
                app, "embed", side_effect=AssertionError("attribute search called API")
            ),
            patch.object(
                storage,
                "open_index",
                side_effect=AssertionError("attribute search opened vectors"),
            ),
        ):
            first = app.search("", **options)
            second = app.search("", cursor=first["next_cursor"], **options)
        expected_first = sorted([mails[1]["id"], mails[2]["id"]], reverse=True)
        self.assertEqual([row["mail_id"] for row in first["items"]], expected_first)
        self.assertEqual(
            [row["entry_id"] for row in second["items"]], ["local-clock", "old-archive"]
        )
        self.assertIsNone(second["next_cursor"])
        self.assertEqual(first["embedding_ms"], 0)
        self.assertEqual(first["search_scope"]["stores"], "all_indexed_stores")
        with self.assertRaises(ValueError):
            app.search(
                "", cursor=first["next_cursor"], **dict(options, sender="other@")
            )
        for invalid in ("invalid", base64.urlsafe_b64encode(b"[]").decode()):
            with self.assertRaises(ValueError):
                app.search("", mode="metadata", cursor=invalid)

    def test_thread_context_contains_latest_reversal_and_original_bodies(self):
        messages = [
            self.mail(
                "approval",
                store_id="pst",
                conversation_id="spec",
                folder="archive",
                received="2026-10-01",
                body="PCIe Gen7 仕様を承認しました。採用を承認します。",
            ),
            self.mail(
                "withdrawal",
                store_id="pst",
                conversation_id="spec",
                folder="archive",
                received="2026-10-05",
                body="前の承認は撤回します。性能試験が必要です。",
            ),
            self.mail(
                "latest",
                store_id="pst",
                conversation_id="spec",
                folder="archive",
                received="2026-10-08",
                body="試験に合格したので再承認しました。",
            ),
            self.mail(
                "other-store",
                store_id="other",
                conversation_id="spec",
                body="PCIe Gen7 別の会話。",
            ),
        ]
        self.index(*messages)
        result = app.search("PCIe Gen7 承認の変更経緯", mode="thread")
        item = next(row for row in result["items"] if row["store_id"] == "pst")
        evidence = item["thread_context"]["items"]
        self.assertIn(messages[2]["id"], [row["mail_id"] for row in evidence])
        self.assertEqual(
            [row["received"] for row in evidence],
            sorted(row["received"] for row in evidence),
        )
        originals = {mail["id"]: mail["body"] for mail in messages}
        self.assertTrue(
            all(row["body"] == originals[row["mail_id"]] for row in evidence)
        )
        self.assertTrue(all(row["store_id"] == "pst" for row in evidence))
        self.assertFalse(item["thread_context"]["summary_is_generated"])

    def test_thread_neighbors_compare_instants_across_timezones(self):
        earlier = self.mail("local", received="2026-10-13T10:00:00+09:00")
        later = self.mail("utc", received="2026-10-13T03:00:00+00:00")
        self.index(earlier, later)
        result = features.thread(self.cfg, earlier["id"])
        self.assertEqual(
            [m["mail_id"] for m in result["items"]], [earlier["id"], later["id"]]
        )
        self.assertEqual([m["position"] for m in result["items"]], ["focus", "next"])
        self.assertTrue(all("instant" not in m for m in result["items"]))

    def test_thread_rebuild_failure_rolls_back_and_cli_is_offline(self):
        self.index(self.mail("mail", subject="PCIe"))
        before = list(self.sql.execute("SELECT key,fingerprint FROM thread_documents"))
        with patch.object(threads, "refresh", side_effect=RuntimeError("fault")):
            with self.assertRaises(RuntimeError):
                threads.rebuild(self.cfg)
        self.assertEqual(
            list(self.sql.execute("SELECT key,fingerprint FROM thread_documents")),
            before,
        )
        for args in (
            ["rebuild-threads"],
            ["search", "--mode", "metadata", "--subject", "PCIe"],
        ):
            with (
                patch("sys.argv", ["outlook-rag", *args]),
                patch("sys.stdout", io.StringIO()) as output,
                patch.object(app, "embed", side_effect=AssertionError("CLI used API")),
                patch.object(
                    app, "api_key", side_effect=AssertionError("CLI requested key")
                ),
            ):
                cli.main(app)
                self.assertIsInstance(json.loads(output.getvalue()), dict)

    def test_thread_search_respects_disabled_lexical_source_and_excerpt_bound(self):
        self.index(self.mail("long-header", subject="題" * 8000, body="最新の決定"))
        self.assertLessEqual(
            self.sql.execute(
                "SELECT length(search_text) FROM thread_documents"
            ).fetchone()[0],
            6000,
        )
        self.cfg.update(lexical_weight=0, semantic_weight=1)
        with patch.object(
            threads, "candidates", side_effect=AssertionError("disabled source queried")
        ):
            result = app.search("決定", mode="thread")
        self.assertFalse(result["thread_parent_search_used"])
        self.assertTrue(
            result["items"][0]["thread_context"]["items"][0]["is_latest_scoped"]
        )

    def test_thread_filters_also_filter_context_and_disable_whole_thread_text(self):
        self.index(
            self.mail(
                "old", received="2008-06-01", body="PCIe Gen7 SECRET", folder="archive"
            ),
            self.mail(
                "allowed",
                received="2026-10-10",
                body="PCIe Gen7 permitted",
                folder="archive",
            ),
            self.mail(
                "inbox",
                received="2026-10-12",
                body="PCIe Gen7 excluded",
                folder="inbox",
            ),
            self.mail(
                "different-sender",
                received="2026-10-11",
                body="PCIe Gen7 excluded",
                sender="other",
                folder="archive",
            ),
        )
        result = app.search(
            "PCIe Gen7",
            mode="thread",
            folder="archive",
            sender="Sender@",
            since="2026-01-01",
        )
        self.assertFalse(result["thread_parent_search_used"])
        for item in result["items"]:
            self.assertEqual(item["thread_mail_count"], 1)
            self.assertEqual(
                [row["entry_id"] for row in item["thread_context"]["items"]],
                ["allowed"],
            )
        self.assertNotIn("SECRET", json.dumps(result))

    def test_thread_refresh_is_local_incremental_and_tracks_move_delete(self):
        a = self.mail("a", body="approved", conversation_id="a")
        b = self.mail("b", body="unrelated", conversation_id="b")
        self.index(a, b)
        fingerprints = dict(
            self.sql.execute("SELECT conversation_id,fingerprint FROM thread_documents")
        )
        same = self.index(a)
        self.assertEqual(same["threads_updated"], 0)
        reply = self.mail(
            "reply", body="withdrawn", conversation_id="a", received="2026-10-14"
        )
        self.assertEqual(self.index(reply)["threads_updated"], 1)
        self.assertEqual(
            self.sql.execute(
                "SELECT fingerprint FROM thread_documents WHERE conversation_id='b'"
            ).fetchone()[0],
            fingerprints["b"],
        )
        moved = dict(reply, conversation_id="b")
        self.assertEqual(self.index(moved)["threads_updated"], 2)
        app.remove_mail(reply["id"], self.sql, self.table)
        doc = self.sql.execute(
            "SELECT anchor_mail_id,member_count FROM thread_documents WHERE conversation_id='b'"
        ).fetchone()
        self.assertEqual(tuple(doc), (b["id"], 1))

    def test_thread_publish_failure_keeps_old_mail_and_thread_document(self):
        mail = self.mail("mail", body="old decision")
        self.index(mail)
        old = tuple(self.sql.execute("SELECT * FROM thread_documents").fetchone())
        with patch.object(
            threads, "refresh", side_effect=RuntimeError("publication fault")
        ):
            with self.assertRaises(RuntimeError):
                self.index(dict(mail, body="new decision"))
        self.assertEqual(
            tuple(self.sql.execute("SELECT * FROM thread_documents").fetchone()), old
        )
        self.assertEqual(
            self.sql.execute("SELECT body FROM emails").fetchone()[0], "old decision"
        )

    def test_legacy_thread_search_never_backfills_and_offline_rebuild_restores_it(self):
        self.index(self.mail("thread"), self.mail("standalone", conversation_id=""))
        with self.sql:
            self.sql.execute("DROP TABLE thread_documents")
            self.sql.execute("DROP TABLE thread_lexical")
        with patch.object(
            threads, "initialize", side_effect=AssertionError("search wrote schema")
        ):
            result = app.search("body", mode="thread")
        self.assertEqual(result["thread_index"]["state"], "missing")
        with patch.object(
            app, "embed", side_effect=AssertionError("thread rebuild called API")
        ):
            report = threads.rebuild(self.cfg)
        self.assertEqual(report["threads"], 2)
        self.assertEqual(report["embedding_requests"], 0)
        self.assertEqual(threads.status(self.sql)["state"], "complete")

    def test_thread_context_global_body_budget_is_bounded(self):
        self.index(
            *[
                self.mail(
                    str(i),
                    body="long original " * 400,
                    conversation_id=str(i // 3),
                    received=f"2026-10-{i + 1:02d}",
                )
                for i in range(15)
            ]
        )
        result = app.search("original", mode="thread")
        self.assertLessEqual(result["thread_context"]["body_chars"], 12000)
        self.assertTrue(
            any(row["thread_context"]["truncated"] for row in result["items"])
        )


class Paragraphs(unittest.TestCase):
    def test_v3_selects_separate_identity_and_new_modes_do_not_reembed(self):
        env = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RAG_")}
        env.update(OUTLOOK_RAG_MODEL="fixture", OUTLOOK_RAG_DIMENSIONS="4")
        with patch.dict(os.environ, env, clear=True):
            original = app.config()
            self.assertEqual(original["text_cleaning_version"], 2)
            for mode in ("thread", "metadata"):
                with patch.dict(os.environ, {"OUTLOOK_RAG_SEARCH_MODE": mode}):
                    self.assertEqual(
                        app.index_identity(app.config()), app.index_identity(original)
                    )
            with patch.dict(os.environ, {"OUTLOOK_RAG_TEXT_CLEANING_VERSION": "3"}):
                changed = app.config()
                self.assertNotEqual(
                    app.index_identity(changed), app.index_identity(original)
                )
                self.assertNotEqual(changed["data_dir"], original["data_dir"])

    def test_short_mail_is_single_chunk_and_paragraphs_are_not_cut(self):
        body = (
            "段落一。" * 15
            + "\n\n"
            + "重要な決定を撤回しました。" * 8
            + "\n\n"
            + "段落三。" * 15
        )
        parts = text.chunks("変更", body, 150, 20, cleaning_version=3)
        for paragraph in body.split("\n\n"):
            self.assertTrue(any(paragraph in chunk for chunk in parts))
        self.assertEqual(
            len(text.chunks("短文", "承認します。", cleaning_version=3)), 1
        )

    def test_oversized_paragraphs_preserve_all_content_with_bounded_chunks(self):
        for body in ("英語 single line " * 100, "日本語の長文です。" * 100, "x" * 2500):
            parts = text.chunks("s", body, 150, 0, cleaning_version=3)
            bodies = [part.split("\n\n", 1)[1] for part in parts]
            self.assertTrue(all(len(part) <= 150 for part in bodies))
            self.assertEqual("".join(bodies).replace("\n", ""), body.strip())

    def test_v3_keeps_forwarded_facts_and_v2_is_unchanged(self):
        body = "ご確認ください。\n--- Original Message ---\n仕様は32 GT/sです。"
        self.assertNotIn(
            "32 GT/s", "\n".join(text.chunks("Fwd: 仕様", body, cleaning_version=2))
        )
        self.assertIn(
            "32 GT/s", "\n".join(text.chunks("Fwd: 仕様", body, cleaning_version=3))
        )
        self.assertIn("原文の決定", text.clean_body("> 原文の決定", 3))
        self.assertNotIn("旧決定", text.clean_body("新規の撤回。\n> 旧決定", 3))


if __name__ == "__main__":
    unittest.main()
