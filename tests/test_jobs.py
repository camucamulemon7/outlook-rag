"""Bulk jobs continue across bounded cycles, serialize workers, and cancel safely."""
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from filelock import FileLock
from outlook_rag import app, jobs


class BulkJobs(unittest.TestCase):
    def setUp(self):
        Path('work').mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir='work', ignore_cleanup_errors=True)
        self.cfg = dict(model='fixture', dimensions=4, data_dir=str(Path(self.temp.name).resolve()),
                        embedding_url='http://fixture.invalid')
        self.sql, self.table = app.open_store(self.cfg)
        self.auth = patch.object(app, 'api_key', return_value='fixture-key')
        self.auth.start()

    def tearDown(self):
        self.auth.stop()
        self.sql.close()
        self.temp.cleanup()

    def start(self):
        with patch.object(jobs, 'launch_worker') as launch:
            record = jobs.start(self.cfg)
        self.assertEqual(launch.call_count, 1)
        return record['job_id']

    @staticmethod
    def cycle(folder, done, pending=None, scanned=1, reason=None):
        return dict(folders=[dict(folder=folder, window_complete=done)], pending_folders=pending or [],
                    indexed_total=1 if scanned else 0, scanned_total=scanned, retried=0,
                    stop_reason=reason, elapsed_ms=10)

    def test_continuous_cycles_and_duplicate_start(self):
        job = self.start()
        with patch.object(jobs, 'launch_worker') as launch:
            duplicate = jobs.start(self.cfg)
        self.assertEqual(duplicate['job_id'], job)
        self.assertTrue(duplicate['reused_existing_job'])
        launch.assert_not_called()
        with patch.object(app, 'sync', side_effect=[self.cycle('inbox', False, ['archive']),
                                                   self.cycle('archive', True), self.cycle('inbox', True)]) as sync:
            jobs.run(job, self.cfg)
        state = jobs.status(self.cfg, job)
        self.assertEqual(state['state'], 'completed')
        self.assertEqual(state['cycles'], 3)
        self.assertEqual(state['completed_folders'], 2)
        self.assertEqual(sync.call_count, 3)
        self.assertEqual(state['indexed_total'], 3)

    def test_cancel_between_cycles_retains_index(self):
        job = self.start()
        def cycle(**kwargs):
            self.assertFalse(kwargs['_cancel_check']())
            jobs.cancel(self.cfg, job)
            self.assertTrue(kwargs['_cancel_check']())
            return self.cycle('inbox', False, reason='cancelled')
        with patch.object(app, 'sync', side_effect=cycle) as sync:
            jobs.run(job, self.cfg)
        self.assertEqual(jobs.status(self.cfg, job)['state'], 'cancelled')
        self.assertEqual(sync.call_count, 1)

    def test_exited_worker_and_resume_use_new_job(self):
        job = self.start()
        path, record = jobs.load(self.cfg, job)
        record['state'] = 'running'
        jobs.write(path, record)
        self.assertEqual(jobs.status(self.cfg, job)['state'], 'interrupted')
        next_job = self.start()
        self.assertNotEqual(job, next_job)

    def test_no_progress_stops_instead_of_spinning(self):
        job = self.start()
        with patch.object(app, 'sync', return_value=self.cycle('inbox', False, scanned=0)) as sync:
            jobs.run(job, self.cfg)
        self.assertEqual(jobs.status(self.cfg, job)['state'], 'blocked')
        self.assertEqual(sync.call_count, 3)

    def test_folder_errors_are_not_reported_as_complete(self):
        job = self.start()
        result = self.cycle('inbox', False)
        result['folder_errors'] = [{'error_type':'fixture'}]
        with patch.object(app, 'sync', return_value=result):
            jobs.run(job, self.cfg)
        state = jobs.status(self.cfg, job)
        self.assertEqual(state['state'], 'completed_with_errors')
        self.assertEqual(state['completed_folders'], 0)

    def test_cancelled_embedding_never_calls_provider(self):
        with patch.object(app.httpx, 'Client') as client:
            with self.assertRaises(app.SyncCancelled):
                app.embed(['private text'], dict(self.cfg, _cancel_check=lambda: True))
        client.return_value.post.assert_not_called()

    def test_other_worker_does_not_mask_an_interrupted_job(self):
        job = self.start()
        path, record = jobs.load(self.cfg, job)
        record['state'] = 'running'
        jobs.write(path, record)
        jobs.write(jobs.root(self.cfg)/'worker.owner.json', {'job_id':'0'*32})
        with FileLock(jobs.root(self.cfg)/'worker.lock', timeout=0):
            self.assertEqual(jobs.status(self.cfg, job)['state'], 'interrupted')

    def test_snapshot_does_not_store_api_key(self):
        with patch.dict(jobs.os.environ, {'OUTLOOK_RAG_API_KEY':'fixture-secret'}), patch.object(jobs,'launch_worker') as launch:
            state = jobs.start(dict(self.cfg, api_key='fixture-secret'))
        snapshot = jobs.root(self.cfg)/(state['job_id']+'.config.json')
        self.assertNotIn('fixture-secret', snapshot.read_text(encoding='utf-8'))
        self.assertEqual(launch.call_args.kwargs['env']['OUTLOOK_RAG_API_KEY'], 'fixture-secret')
        with self.assertRaises(ValueError):
            jobs.status(self.cfg, '../escape')


if __name__ == '__main__':
    unittest.main()
