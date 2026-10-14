"""COM lookup retries and bulk-job recovery after the second bounded cycle."""

import hashlib
import json
from collections import Counter
from contextlib import ExitStack
from types import SimpleNamespace
import threading
import unittest
from unittest.mock import call, patch

from outlook_rag import app, coverage, jobs, outlook
from outlook_test_support import synthetic_com
import test_improvements as fixtures


class FlakyAttributes:
    def __init__(self, wrapped, failures=None, error=None):
        self.wrapped = wrapped
        self.failures = dict(failures or {})
        self.error = error or AttributeError("private COM lookup detail")
        self.calls = Counter()
        self.threads = []

    def __getattr__(self, name):
        self.calls[name] += 1
        self.threads.append(threading.get_ident())
        if self.failures.get(name, 0):
            self.failures[name] -= 1
            raise self.error
        return getattr(self.wrapped, name)


class ComAttributes(unittest.TestCase):
    def test_retry_succeeds_on_any_of_six_attempts_on_the_calling_thread(self):
        for failures in range(6):
            with self.subTest(failures=failures):
                source = FlakyAttributes(SimpleNamespace(StoreID="store"), {"StoreID": failures})
                with patch.object(outlook.time, "sleep") as sleep:
                    self.assertEqual(outlook.com_attr(source, "StoreID"), "store")
                self.assertEqual(source.calls["StoreID"], failures + 1)
                self.assertEqual(sleep.call_args_list, [call(.2)] * failures)
                self.assertEqual(set(source.threads), {threading.get_ident()})

    def test_exhaustion_raises_the_last_error_without_a_default_id(self):
        error = AttributeError("private COM lookup detail")
        source = FlakyAttributes(SimpleNamespace(StoreID="store"), {"StoreID": 6}, error)
        with patch.object(outlook.time, "sleep") as sleep:
            with self.assertRaises(AttributeError) as raised:
                outlook.com_attr(source, "StoreID")
        self.assertIs(raised.exception, error)
        self.assertEqual(source.calls["StoreID"], 6)
        self.assertEqual(sleep.call_args_list, [call(.2)] * 5)

    def test_native_com_errors_are_also_retried(self):
        with synthetic_com():
            import pywintypes
            error = pywintypes.com_error(-2147418111, "private COM detail", None, None)
            source = FlakyAttributes(SimpleNamespace(EntryID="entry"), {"EntryID": 2}, error)
            with patch.object(outlook.time, "sleep") as sleep:
                self.assertEqual(outlook.com_attr(source, "EntryID"), "entry")
            self.assertEqual(source.calls["EntryID"], 3)
            self.assertEqual(sleep.call_args_list, [call(.2)] * 2)

    def test_unrelated_errors_are_not_retried(self):
        source = FlakyAttributes(SimpleNamespace(StoreID="store"), {"StoreID": 6}, ValueError("bug"))
        with patch.object(outlook.time, "sleep") as sleep, self.assertRaises(ValueError):
            outlook.com_attr(source, "StoreID")
        self.assertEqual(source.calls["StoreID"], 1)
        sleep.assert_not_called()

    def test_folder_identity_and_coverage_recover_without_changing_keys(self):
        for read in (
            lambda folder: outlook.folder_identity(folder),
            lambda folder: coverage.state_key(folder, 30),
            lambda folder: coverage.inventory(None, [("inbox", folder)], ["inbox"], 30),
            lambda folder: coverage.live_coverage({"stores": [], "errors": []},
                                                  [("inbox", folder)], None, {}),
        ):
            with self.subTest(reader=read):
                plain = SimpleNamespace(StoreID="store", EntryID="folder")
                flaky = FlakyAttributes(plain, {"StoreID": 1, "EntryID": 1})
                with patch.object(outlook.time, "sleep"):
                    self.assertEqual(read(flaky), read(plain))

    def test_source_discovery_retries_root_folder_and_search_folder_ids(self):
        child = FlakyAttributes(SimpleNamespace(
            Name="Inbox", EntryID="inbox", StoreID="store", DefaultItemType=0,
            Items=SimpleNamespace(Count=6), Folders=[]), {"EntryID": 1})
        root = FlakyAttributes(SimpleNamespace(
            Name="Archive", EntryID="root", StoreID="store", DefaultItemType=1,
            Items=SimpleNamespace(Count=0), Folders=[child]), {"StoreID": 1, "EntryID": 1})
        search = FlakyAttributes(SimpleNamespace(EntryID="search"), {"EntryID": 1})
        store = SimpleNamespace(DisplayName="Archive", StoreID="", FilePath="fixture.pst",
                                GetRootFolder=lambda: root, GetSearchFolders=lambda: [search])
        with patch.object(outlook.time, "sleep"):
            report, targets = outlook.discover_sources(SimpleNamespace(Stores=[store]),
                                                       {"exclude_system_folders": False})
        self.assertEqual(report["errors"], [])
        self.assertEqual(report["stores"][0]["store_id"], "store")
        self.assertEqual([name for name, _ in targets], ["Archive/Inbox"])
        self.assertEqual(root.calls["StoreID"], 2)
        self.assertEqual(root.calls["EntryID"], 2)
        self.assertEqual(child.calls["EntryID"], 2)
        self.assertEqual(search.calls["EntryID"], 2)


class ComJobRecovery(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        self.cfg = self.fixture.cfg
        self.cfg.update(sync_max_total_emails=2, sync_prefetch_emails=0,
                        embedding_concurrent_requests=1)
        self.fixture.folder("inbox", 6)
        self.folder = FlakyAttributes(self.fixture.folders["inbox"])
        self.fixture.folders["inbox"] = self.folder
        self.context = ExitStack()
        self.addCleanup(self.context.close)
        self.context.enter_context(patch.object(app, "config", return_value=self.cfg))
        self.context.enter_context(patch.object(app, "api_key", return_value="fixture"))
        self.context.enter_context(patch.object(app, "resolve_folder", return_value=self.folder))
        self.context.enter_context(patch("win32com.client.Dispatch", return_value=self.fixture.dispatch))
        self.context.enter_context(patch.object(outlook.time, "sleep"))

    def start(self):
        with patch.object(jobs, "launch_worker"):
            return jobs.start(self.cfg)["job_id"]

    def run_with_fault_after_second_cycle(self, failures):
        cycles = []
        original = app.sync

        def sync(*args, **kwargs):
            if len(cycles) == 2:
                self.folder.failures = {"StoreID": failures, "EntryID": failures}
            result = original(*args, **kwargs)
            cycles.append(result)
            return result

        job = self.start()
        with patch.object(app, "sync", side_effect=sync):
            jobs.run(job, self.cfg)
        return job, cycles

    def assert_complete(self, job):
        status = jobs.status(self.cfg, job)
        self.assertEqual(status["state"], "completed")
        self.assertEqual(status["coverage"]["emails"], 6)
        self.assertEqual(status["coverage"]["chunks"], 6)
        self.assertEqual(len(self.fixture.reads), 6)
        ids = {row[0] for row in self.fixture.sql.execute("SELECT id FROM emails")}
        expected = {hashlib.sha256(f"store:inbox{i}".encode()).hexdigest() for i in range(6)}
        self.assertEqual(ids, expected)

    def test_transient_lookup_after_second_cycle_continues_to_completion(self):
        job, cycles = self.run_with_fault_after_second_cycle(5)
        self.assertEqual(len(cycles), 3)
        self.assert_complete(job)
        folder_ids = {result["target_folders"][0]["folder_id"] for result in cycles}
        self.assertEqual(len(folder_ids), 1)
        inventory = json.loads(self.fixture.sql.execute(
            "SELECT value FROM state WHERE key='folder_inventory'").fetchone()[0])
        self.assertEqual(inventory["folders"][0]["state_key"], "sync:store:inbox:0")

    def test_persistent_lookup_failure_retains_committed_mail_and_can_restart(self):
        job, cycles = self.run_with_fault_after_second_cycle(6)
        self.assertEqual(len(cycles), 2)
        status = jobs.status(self.cfg, job)
        self.assertEqual(status["state"], "failed")
        self.assertEqual(status["error_type"], "AttributeError")
        self.assertEqual(status["coverage"]["emails"], 4)
        self.assertNotIn("private COM", json.dumps(status))
        self.folder.failures.clear()
        resumed = self.start()
        jobs.run(resumed, self.cfg)
        self.assert_complete(resumed)
        self.assertNotEqual(resumed, job)
        self.assertEqual(jobs.status(self.cfg, job)["state"], "failed")

    def test_lookup_failure_after_batch_commit_does_not_fail_the_cycle(self):
        original = app.upsert_mails

        def upsert(*args, **kwargs):
            result = original(*args, **kwargs)
            self.folder.failures = {"StoreID": 1, "EntryID": 1}
            return result

        job = self.start()
        with patch.object(app, "upsert_mails", side_effect=upsert):
            jobs.run(job, self.cfg)
        self.assert_complete(job)


if __name__ == "__main__":
    unittest.main()
