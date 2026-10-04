"""Exercise bulk job cancellation/restart through real sync and synthetic Outlook."""
import threading
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch

import httpx
from filelock import FileLock
from outlook_rag import app, jobs
import test_improvements as fixtures


class JobRecovery(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.cfg = self.fixture.cfg
        self.cfg.update(sync_max_total_emails=2, sync_prefetch_emails=0,
                        embedding_concurrent_requests=1)
        self.fixture.folder('inbox', 6)
        self.context = ExitStack()
        self.context.enter_context(patch.object(app, 'config', return_value=self.cfg))
        self.context.enter_context(patch.object(app, 'api_key', return_value='fixture-key'))
        self.context.enter_context(patch.object(app, 'resolve_folder',
            side_effect=lambda namespace, name: self.fixture.folders[name]))
        self.context.enter_context(patch('win32com.client.Dispatch', return_value=self.fixture.dispatch))
        self.context.enter_context(patch.object(app, 'detect_dimensions',
            side_effect=AssertionError('API probe forbidden')))

    def tearDown(self):
        self.context.close()
        self.fixture.tearDown()

    def start(self):
        with patch.object(jobs, 'launch_worker'):
            return jobs.start(self.cfg)['job_id']

    def test_provider_outage_blocks_and_restart_after_recovery_indexes_remaining_mail(self):
        job = self.start()
        requests = []
        def provider(texts, *_args, **_kwargs):
            requests.append(list(texts))
            if len(requests) > 1:
                raise httpx.ConnectError('Synthetic provider outage')
            return [[1., 0., 0., 0.] for _ in texts]
        with patch.object(app, 'embed', side_effect=provider), patch.object(jobs.time, 'sleep'):
            jobs.run(job, self.cfg)
        state = app.sync_job_status(job)
        self.assertEqual(state['state'], 'blocked')
        self.assertEqual(state['reason'], 'embedding_unavailable')
        self.assertEqual(state['coverage']['emails'], 2)
        self.assertEqual(state['coverage']['cached_embeddings'], 2)
        self.assertEqual(len(requests), 4)
        # Reopen the durable store, as a worker after a process restart would.
        self.fixture.sql.close()
        self.fixture.sql, self.fixture.table = app.open_store(self.cfg)
        resumed = self.start()
        self.assertNotEqual(resumed, job)
        jobs.run(resumed, self.cfg)
        finished = app.sync_job_status(resumed)
        self.assertEqual(finished['state'], 'completed')
        self.assertEqual(finished['coverage']['emails'], 6)
        self.assertEqual(finished['coverage']['chunks'], 6)
        self.assertEqual(finished['indexed_total'], 4)
        self.assertEqual(app.sync_job_status(job)['state'], 'blocked')
        found = app.search('body inbox')
        self.assertEqual(found['count'], 6)
        self.assertTrue(app.search('body inbox')['query_cache_hit'])

    def test_concurrent_cancel_during_outage_preserves_published_mail_and_new_job_resumes(self):
        job = self.start()
        entered, release = threading.Event(), threading.Event()
        requests = []
        def provider(texts, *_args, **_kwargs):
            requests.append(list(texts))
            if len(requests) > 1:
                entered.set()
                if not release.wait(5):
                    raise AssertionError('Cancellation test did not release the provider')
                raise httpx.ConnectError('Synthetic provider outage')
            return [[1., 0., 0., 0.] for _ in texts]
        with patch.object(app, 'embed', side_effect=provider):
            worker = threading.Thread(target=jobs.run, args=(job, self.cfg), daemon=True)
            worker.start()
            try:
                self.assertTrue(entered.wait(5))
                before = app.sync_job_status(job)
                self.assertEqual(before['state'], 'running')
                self.assertEqual(before['coverage']['emails'], 2)
                for _ in range(2):
                    cancelled = app.cancel_sync_job(job)
                    self.assertTrue(cancelled['cancel_requested'])
            finally:
                release.set()
                worker.join(5)
            self.assertFalse(worker.is_alive())
        state = app.sync_job_status(job)
        self.assertEqual(state['state'], 'cancelled')
        self.assertEqual(state['coverage']['emails'], 2)
        self.assertEqual(state['coverage']['failed_emails'], 0)
        resumed = self.start()
        self.assertFalse(app.sync_job_status(resumed)['cancel_requested'])
        jobs.run(resumed, self.cfg)
        self.assertEqual(app.sync_job_status(resumed)['coverage']['emails'], 6)
        self.assertEqual(app.sync_job_status(resumed)['state'], 'completed')

    def test_cancel_while_index_busy_stops_without_retrying_or_touching_mail(self):
        job = self.start()
        seen = []
        def cancel_wait(_delay):
            seen.append(app.sync_job_status(job)['state'])
            app.cancel_sync_job(job)
        lock_path = Path(self.cfg['data_dir'])/'mutation.lock'
        with FileLock(lock_path, timeout=0), patch.object(
                app, 'mutation_lock', return_value=FileLock(lock_path, timeout=0)), \
                patch.object(jobs.time, 'sleep', side_effect=cancel_wait):
            jobs.run(job, self.cfg)
        state = app.sync_job_status(job)
        self.assertEqual(seen, ['waiting'])
        self.assertEqual(state['state'], 'cancelled')
        self.assertEqual(state['cycles'], 0)
        self.assertEqual(state['coverage']['emails'], 0)
        self.assertEqual(self.fixture.reads, [])


if __name__ == '__main__':
    unittest.main()
