"""Public MCP surface and durable PST-adjacent index selection."""
from outlook_test_support import isolated_environment
import asyncio
from contextlib import ExitStack
import json
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from mcp.server.fastmcp import FastMCP

from outlook_rag import app, features, jobs, paths, server
from outlook_test_support import synthetic_com


class ToolProfiles(unittest.TestCase):
    def setUp(self):
        self.backend = SimpleNamespace(search=Mock(return_value={"count": 0}),
            config=Mock(return_value={"data_dir": "unused"}))
        self.mcp = FastMCP("fixture")
        self.adapters = server.register_tools(self.mcp, self.backend)

    def names(self):
        return {tool.name: tool for tool in asyncio.run(self.mcp.list_tools())}

    def test_standard_surface_and_explicit_full_compatibility(self):
        tools = self.names()
        self.assertEqual(set(tools), server.STANDARD_TOOLS)
        self.assertEqual(len(tools), 6)
        self.assertEqual(tools["sync_index"].inputSchema["properties"]["action"]["enum"], ["start", "status", "cancel"])
        self.assertFalse(tools["sync_index"].annotations.readOnlyHint)
        self.assertTrue(tools["get_indexed_mail"].annotations.readOnlyHint)
        server.configure_tools(self.mcp, self.adapters, "full")
        self.assertEqual(set(self.names()), set(self.adapters))
        self.assertEqual(len(self.names()), 15)
        self.assertIn("sync_emails", self.names())
        server.configure_tools(self.mcp, self.adapters)
        self.assertEqual(set(self.names()), server.STANDARD_TOOLS)
        with self.assertRaises(ValueError):
            server.configure_tools(self.mcp, self.adapters, "typo")
        self.assertEqual(set(self.names()), server.STANDARD_TOOLS)

    def test_sync_routes_and_offline_controls(self):
        with patch.object(jobs, "start", return_value={"job_id": "fixture"}) as start:
            result = self.adapters["sync_index"]("start", folders=["auto"], since_days=0)
            self.assertEqual(result["job_id"], "fixture")
            start.assert_called_once_with(self.backend.config.return_value, ["auto"], 0, False)
        self.backend.config.reset_mock()
        with (patch.object(features, "local_config", return_value={"data_dir": "fixture"}),
              patch.object(features, "progress", return_value={"folders": [{"folder": "PST"}]}),
              patch.object(jobs, "status", return_value={"state": "not_started"}) as status,
              patch.object(jobs, "cancel", return_value={"cancel_requested": True}) as cancel):
            self.assertEqual(self.adapters["sync_index"]()["folder_progress"], [{"folder": "PST"}])
            status.assert_called_once_with({"data_dir": "fixture"}, None)
            self.assertTrue(self.adapters["sync_index"]("cancel", "a" * 32)["cancel_requested"])
            cancel.assert_called_once_with({"data_dir": "fixture"}, "a" * 32)
        self.backend.config.assert_not_called()
        for kwargs in ({"action": "bad"}, {"action": "start", "job_id": "a" * 32},
                       {"action": "status", "folders": ["auto"]}, {"action": "cancel", "since_days": 0}):
            with self.subTest(kwargs=kwargs), self.assertRaises(ValueError):
                self.adapters["sync_index"](**kwargs)


class IndexPlacement(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.temp = self.stack.enter_context(tempfile.TemporaryDirectory())
        self.root = Path(self.temp).resolve()
        self.stack.enter_context(patch.dict(os.environ, isolated_environment(dict(LOCALAPPDATA=str(self.root / "local"),
            OUTLOOK_RAG_MODEL="fixture", OUTLOOK_RAG_DIMENSIONS="4")), clear=True))
        self.stack.enter_context(patch.object(Path, "home", return_value=self.root / "home"))
        self.real_discovery = paths.connected_pst_directories
        self.discovery = self.stack.enter_context(patch.object(paths, "connected_pst_directories", return_value=[self.root / "メール"] ))

    def test_reads_do_not_discover_outlook_and_sync_remembers_pst_root(self):
        cfg = app.config()
        self.discovery.assert_not_called()
        selected = paths.prepare(cfg)
        self.assertEqual(Path(selected["data_dir"]).parent, self.root / "メール" / "outlook-rag")
        self.assertTrue(Path(selected["_index_location_file"]).is_file())
        self.discovery.assert_called_once()
        self.discovery.side_effect = AssertionError("Disconnected Outlook must not be accessed")
        self.assertEqual(app.config()["data_dir"], selected["data_dir"])
        self.assertEqual(paths.prepare(app.config())["data_dir"], selected["data_dir"])
        with patch.dict(os.environ, {"OUTLOOK_RAG_MODEL": "another-model"}):
            other = paths.prepare(app.config())
            self.assertEqual(Path(other["data_dir"]).parent, Path(selected["data_dir"]).parent)
            self.assertNotEqual(other["data_dir"], selected["data_dir"])

    def test_existing_index_wins_without_pst_discovery(self):
        cfg = app.config()
        existing = Path(cfg["data_dir"])
        existing.mkdir(parents=True)
        (existing / "metadata.sqlite").touch()
        self.discovery.return_value = [self.root / "one", self.root / "two"]
        selected = paths.prepare(cfg)
        self.assertEqual(selected["data_dir"], cfg["data_dir"])
        self.discovery.assert_not_called()
        self.assertEqual(app.config()["data_dir"], cfg["data_dir"])

    def test_status_created_empty_job_directory_does_not_select_location(self):
        cfg = app.config()
        self.assertEqual(jobs.status(cfg)["state"], "not_started")
        selected = paths.prepare(cfg)
        self.assertEqual(Path(selected["data_dir"]).parent, self.root / "メール" / "outlook-rag")

    def test_multiple_directories_require_explicit_root_without_creating_db(self):
        cfg = app.config()
        self.discovery.return_value = [self.root / "one", self.root / "two"]
        with self.assertRaisesRegex(ValueError, "OUTLOOK_RAG_DATA_ROOT"):
            paths.prepare(cfg)
        self.assertFalse(Path(cfg["_index_location_file"]).exists())
        self.assertFalse(Path(cfg["data_dir"]).exists())
        with patch.dict(os.environ, {"OUTLOOK_RAG_DATA_ROOT": str(self.root / "two" / "outlook-rag")}):
            selected = paths.prepare(app.config())
            self.assertEqual(Path(selected["data_dir"]).parent, self.root / "two" / "outlook-rag")
        self.assertEqual(self.discovery.call_count, 1)

    def test_explicit_directory_and_relative_root_resolve_without_discovery(self):
        settings = self.root / "settings.json"
        settings.write_text(json.dumps(dict(model="fixture", dimensions=4, data_root="archives/outlook-rag")), encoding="utf-8")
        with patch.dict(os.environ, {"OUTLOOK_RAG_CONFIG": str(settings)}):
            cfg = paths.prepare(app.config())
            self.assertEqual(Path(cfg["data_dir"]).parent, self.root / "archives" / "outlook-rag")
            with patch.dict(os.environ, {"OUTLOOK_RAG_DATA_DIR": str(self.root / "explicit")}):
                self.assertEqual(paths.prepare(app.config())["data_dir"], str(self.root / "explicit"))
        self.discovery.assert_not_called()

    def test_offline_auto_dimensions_and_job_snapshot_follow_remembered_root(self):
        cfg = paths.prepare(app.config())
        with patch.object(jobs, "launch_worker"):
            record = jobs.start(cfg)
        with patch.dict(os.environ, {"OUTLOOK_RAG_DIMENSIONS": "auto"}), patch.object(app, "detect_dimensions", side_effect=AssertionError("API not allowed")):
            offline = features.local_config()
            self.assertEqual(offline["data_dir"], cfg["data_dir"])
            self.assertEqual(offline["dimensions"], 4)
            self.assertEqual(jobs.status(offline)["job_id"], record["job_id"])
            self.assertTrue(jobs.cancel(offline)["cancel_requested"])
        self.assertEqual(self.discovery.call_count, 1)

    def test_job_selects_location_before_publishing_snapshot(self):
        original = app.config()
        with patch.object(jobs, "launch_worker"):
            record = jobs.start(original)
        selected = app.config()
        snapshot = jobs.root(selected) / (record["job_id"] + ".config.json")
        self.assertEqual(json.loads(snapshot.read_text(encoding="utf-8"))["data_dir"], selected["data_dir"])
        self.assertEqual(Path(selected["data_dir"]).parent, self.root / "メール" / "outlook-rag")
        self.assertFalse((Path(original["data_dir"]) / "jobs").exists())

    def test_saved_root_never_redirects_to_an_old_default_index(self):
        original = app.config()
        selected = paths.prepare(original)
        legacy = Path(original["data_dir"])
        legacy.mkdir(parents=True)
        (legacy / "metadata.sqlite").touch()
        self.assertEqual(app.config()["data_dir"], selected["data_dir"])
        # Simulate a saved parent that is now inaccessible/not a directory.
        parent = Path(selected["data_dir"]).parent
        parent.rmdir()
        parent.write_text("unavailable", encoding="utf-8")
        with self.assertRaises(OSError):
            paths.prepare(app.config())
        self.assertEqual(app.config()["data_dir"], selected["data_dir"])

    def test_corrupt_location_and_unwritable_root_never_fall_back(self):
        cfg = app.config()
        location = Path(cfg["_index_location_file"])
        location.parent.mkdir(parents=True)
        location.write_text("invalid", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "saved index location"):
            app.config()
        location.unlink()
        blocked = self.root / "blocked"
        blocked.write_text("file", encoding="utf-8")
        self.discovery.return_value = [blocked]
        with self.assertRaises(OSError):
            paths.prepare(app.config())
        self.assertFalse(location.exists())
        self.assertFalse(Path(cfg["data_dir"]).exists())

    def test_store_path_inventory_uses_only_pst_paths(self):
        parent = self.root / "日本語 メール"
        namespace = SimpleNamespace(Stores=[SimpleNamespace(FilePath=str(parent / "a.PST")),
            SimpleNamespace(FilePath=str(parent / "b.pst")), SimpleNamespace(FilePath=str(self.root / "cache.ost"))])
        self.assertEqual(paths.pst_directories(namespace), [parent])
        with self.assertRaisesRegex(ValueError, "relative PST"):
            paths.pst_directories(SimpleNamespace(Stores=[SimpleNamespace(FilePath="relative.pst")]))

    def test_windows_discovery_balances_com_and_reads_only_store_paths(self):
        # Run the deferred Windows adapter against synthetic COM on all hosts.
        with synthetic_com(), patch.object(paths.sys, "platform", "win32"):
            import pythoncom
            import win32com.client
            namespace = SimpleNamespace(Stores=[SimpleNamespace(FilePath=str(self.root / "archive.pst"))])
            with (patch.object(pythoncom, "CoInitialize") as initialize,
                  patch.object(pythoncom, "CoUninitialize") as uninitialize,
                  patch.object(win32com.client, "Dispatch", return_value=SimpleNamespace(GetNamespace=lambda _: namespace))):
                self.assertEqual(self.real_discovery(), [self.root])
                initialize.assert_called_once()
                uninitialize.assert_called_once()


if __name__ == "__main__":
    unittest.main()
