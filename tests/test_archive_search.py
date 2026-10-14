"""Archive scope, bounded all-store sync and truthful no-match responses."""

from datetime import datetime, timezone
import json
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from outlook_rag import app, coverage, jobs
import test_improvements as fixtures


class ArchiveSearch(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        self.cfg, self.sql = self.fixture.cfg, self.fixture.sql
        self.cfg.update(
            text_cleaning_version=2,
            embedding_batch_size=8,
            embedding_concurrent_requests=2,
        )
        self.fixture.folder("inbox", 1)
        rows = self.fixture.folder("archive", 1)
        old = datetime(2008, 6, 1, tzinfo=timezone.utc)
        rows[0] = (rows[0][0], rows[0][1], old, rows[0][3])
        self.fixture.mails["archive0"].ReceivedTime = old
        self.fixture.mails["archive0"].Subject = "旧契約 ARC-2008 の保管"
        inbox, archive = self.fixture.folders["inbox"], self.fixture.folders["archive"]
        for folder, name, store_id in (
            (inbox, "Inbox", "primary-store"),
            (archive, "2008", "pst-store"),
        ):
            folder.Name, folder.StoreID, folder.DefaultItemType = name, store_id, 0
            folder.Folders = []
            folder.Items = SimpleNamespace(Count=1)

        def store(name, store_id, path, child):
            root = SimpleNamespace(
                Name=name,
                EntryID=store_id + "-root",
                StoreID=store_id,
                DefaultItemType=1,
                Folders=[child],
                Items=SimpleNamespace(Count=0),
            )
            child.Parent = root

            def default(kind):
                if kind == 6:
                    return child
                raise ValueError("fixture has no system folder")

            return SimpleNamespace(
                DisplayName=name,
                StoreID=store_id,
                FilePath=path,
                GetRootFolder=lambda: root,
                GetDefaultFolder=default,
                GetSearchFolders=lambda: [],
            )

        primary = store("Primary", "primary-store", "C:/primary.ost", inbox)
        pst = store("Archive PST", "pst-store", "C:/old-mail.PST", archive)
        self.fixture.namespace.Stores = [primary, pst]
        self.fixture.namespace.GetDefaultFolder = primary.GetDefaultFolder
        self.configuration = patch.object(app, "config", return_value=self.cfg)
        self.configuration.start()
        self.addCleanup(self.configuration.stop)
        self.outlook = patch(
            "win32com.client.Dispatch", return_value=self.fixture.dispatch
        )
        self.outlook.start()
        self.addCleanup(self.outlook.stop)

    def all_sync(self):
        for _ in range(6):
            app.sync(folders=["auto"], since_days=0, max_total_emails=1)
            if coverage.index_coverage(self.sql, self.cfg)["known_inventory_complete"]:
                break

    def test_all_store_sync_overrides_inbox_config_and_indexes_old_pst(self):
        app.sync()
        before = app.search_emails("ARC-2008", mode="identifier")
        self.assertEqual(before["count"], 0)
        self.assertFalse(before["absence_confirmed"])
        self.assertIsNone(before["coverage"]["known_inventory_complete"])
        self.all_sync()
        after = app.search_emails("ARC-2008", mode="identifier")
        self.assertEqual(after["count"], 1)
        self.assertEqual(after["items"][0]["store_id"], "pst-store")
        self.assertEqual(after["items"][0]["folder"], "Archive PST/2008")
        self.assertTrue(after["items"][0]["received"].startswith("2008-06-01"))
        self.assertEqual(after["search_scope"]["stores"], "all_indexed_stores")
        self.assertTrue(after["coverage"]["known_inventory_complete"])
        self.assertEqual(after["coverage"]["known_pst_stores"], ["Archive PST"])
        self.assertFalse(after["coverage"]["live_sources_checked"])
        self.assertEqual(self.cfg["folders"], ["inbox"])
        status = app.index_status()
        self.assertIn(
            "Archive PST/2008",
            {folder["folder"] for folder in status["folder_progress"]},
        )
        self.assertTrue(status["search_coverage"]["known_inventory_complete"])

    def test_explicit_inbox_filter_is_honored_and_scope_is_reported(self):
        self.all_sync()
        result = app.search_emails(
            "ARC-2008", folder="Primary/Inbox", mode="identifier"
        )
        self.assertEqual(result["count"], 0)
        self.assertEqual(result["search_scope"]["stores"], "selected_folder")
        self.assertEqual(result["search_scope"]["folder"], "Primary/Inbox")
        self.assertFalse(result["absence_confirmed"])
        self.assertEqual(app.search_emails("ARC-2008", mode="identifier")["count"], 1)

    def test_sources_detects_unindexed_pst_without_body_reads_or_api(self):
        app.sync()
        self.fixture.reads.clear()
        with (
            patch.object(
                app, "embed", side_effect=AssertionError("source probe called API")
            ),
            patch.object(
                app,
                "api_key",
                side_effect=AssertionError("source probe required a key"),
            ),
        ):
            result = app.list_outlook_sources()
        self.assertEqual(self.fixture.reads, [])
        self.assertTrue(result["coverage"]["live_sources_checked"])
        self.assertFalse(result["coverage"]["all_date_scan_complete"])
        self.assertIn("Archive PST/2008", result["coverage"]["pending_folders"])
        pst = next(store for store in result["stores"] if store["is_pst"])
        self.assertEqual(pst["store_id"], "pst-store")
        self.assertEqual(pst["indexed_emails"], 0)
        self.all_sync()
        result = app.list_outlook_sources()
        self.assertTrue(result["coverage"]["all_date_scan_complete"])
        self.assertEqual(
            next(s for s in result["stores"] if s["is_pst"])["indexed_emails"], 1
        )

    def test_partial_scan_reports_pending_archive_even_when_primary_has_hits(self):
        result = app.sync(folders=["auto"], max_total_emails=1, since_days=0)
        saved = json.loads(
            self.sql.execute(
                "SELECT value FROM state WHERE key='folder_inventory'"
            ).fetchone()[0]
        )
        self.assertEqual(saved["selection"], "auto")
        self.assertEqual(
            {f["folder"] for f in saved["folders"]},
            {"Primary/Inbox", "Archive PST/2008"},
        )
        found = app.search_emails("body", mode="lexical")
        self.assertGreater(found["count"], 0)
        self.assertFalse(found["coverage"]["known_inventory_complete"])
        self.assertGreater(found["coverage"]["pending_folder_count"], 0)
        self.assertIn("Archive PST/2008", found["coverage"]["pending_folders"])

    def test_legacy_inventory_is_unknown_and_empty_filter_does_not_call_api(self):
        app.sync()
        with self.sql:
            self.sql.execute(
                "UPDATE state SET value=? WHERE key='folder_inventory'",
                (json.dumps(dict(folders=[dict(folder="inbox")], since_days=0)),),
            )
        with patch.object(
            app, "embed", side_effect=AssertionError("empty filter called API")
        ):
            result = app.search_emails("question", folder="Archive PST/2008")
        self.assertEqual(result["count"], 0)
        self.assertFalse(result["absence_confirmed"])
        self.assertEqual(result["coverage"]["basis"], "unknown")
        self.assertIsNone(result["coverage"]["known_inventory_complete"])

    def test_no_index_reports_unknown_coverage_without_initialization(self):
        cfg = dict(self.cfg, data_dir=str(Path(self.fixture.temp.name) / "missing"))
        with patch.object(app, "config", return_value=cfg):
            result = app.search_emails("old mail")
        self.assertEqual(result["count"], 0)
        self.assertFalse(result["absence_confirmed"])
        self.assertEqual(result["search_scope"]["stores"], "all_indexed_stores")
        self.assertFalse((Path(cfg["data_dir"]) / "metadata.sqlite").exists())

    def test_recent_scan_and_retained_failures_never_claim_complete_archive(self):
        self.all_sync()
        with self.sql:
            row = self.sql.execute(
                "SELECT value FROM state WHERE key='folder_inventory'"
            ).fetchone()
            saved = json.loads(row[0])
            saved["since_days"] = 30
            self.sql.execute(
                "UPDATE state SET value=? WHERE key='folder_inventory'",
                (json.dumps(saved),),
            )
        self.assertFalse(
            coverage.index_coverage(self.sql, self.cfg)["known_inventory_complete"]
        )
        saved["since_days"] = 0
        with self.sql:
            self.sql.execute(
                "UPDATE state SET value=? WHERE key='folder_inventory'",
                (json.dumps(saved),),
            )
            self.sql.execute(
                "INSERT INTO sync_failures(mail_id,attempts) VALUES ('failed',1)"
            )
        self.assertFalse(
            coverage.index_coverage(self.sql, self.cfg)["known_inventory_complete"]
        )
        self.assertFalse(app.sources()["coverage"]["all_date_scan_complete"])

    def test_source_probe_before_index_initialization_never_detects_dimensions(self):
        cfg = dict(
            self.cfg,
            dimensions=None,
            data_dir=str(Path(self.fixture.temp.name) / "uninitialized"),
        )
        with (
            patch.object(app, "config", return_value=cfg),
            patch.object(
                app,
                "detect_dimensions",
                side_effect=AssertionError("dimension probe called API"),
            ),
            patch.object(
                app,
                "api_key",
                side_effect=AssertionError("source probe required credentials"),
            ),
        ):
            result = app.sources()
        self.assertEqual(len(result["stores"]), 2)
        self.assertFalse(result["coverage"]["all_date_scan_complete"])
        self.assertFalse((Path(cfg["data_dir"]) / "metadata.sqlite").exists())

    def test_archive_job_option_survives_snapshot_and_overrides_folder_setting(self):
        with (
            patch.object(jobs, "launch_worker"),
            patch.object(app, "api_key", return_value="fixture-key"),
        ):
            started = app.start_sync_job(folders=["auto"], since_days=0, confirmed=True)
            _, record = jobs.load(self.cfg, started["job_id"])
            self.assertEqual(record["options"]["folders"], ["auto"])
            self.assertEqual(record["options"]["since_days"], 0)
            jobs.run(started["job_id"], self.cfg)
        self.assertEqual(jobs.status(self.cfg, started["job_id"])["state"], "completed")
        self.assertEqual(app.search_emails("ARC-2008", mode="identifier")["count"], 1)

    def test_invalid_auto_mixture_does_not_launch_job(self):
        with (
            patch.object(jobs, "launch_worker") as launch,
            self.assertRaises(ValueError),
        ):
            jobs.start(self.cfg, folders=["auto", "inbox"])
        launch.assert_not_called()


if __name__ == "__main__":
    unittest.main()
