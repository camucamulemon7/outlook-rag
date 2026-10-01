"""Overlap, bounded memory, and checkpoint recovery with COM kept on its owner thread."""
import threading
import time
import unittest
from unittest.mock import patch
import httpx
import test_improvements as fixtures
from outlook_rag import app


class Pipeline(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements()
        self.fixture.setUp()
        self.fixture.cfg['sync_prefetch_emails'] = 2
        self.fixture.folder('inbox', 6)

    def tearDown(self):
        self.fixture.tearDown()

    def wait_for_prefetch(self):
        until = time.monotonic() + 2
        while len(self.fixture.reads) < 4 and time.monotonic() < until:
            time.sleep(.001)
        self.assertGreaterEqual(len(self.fixture.reads), 4, 'Outlook did not read ahead during embedding')

    def test_reads_overlap_embedding_on_owner_thread_with_bounded_buffer(self):
        owner = threading.get_ident()
        read_threads = []
        original = app.read_outlook_mail
        def read(*args):
            read_threads.append(threading.get_ident())
            return original(*args)
        def embed(texts, cfg, **kwargs):
            self.assertNotEqual(threading.get_ident(), owner)
            self.wait_for_prefetch()
            return [[1.,0.,0.,0.] for _ in texts]
        with patch.object(app, 'embed', side_effect=embed), patch.object(app, 'read_outlook_mail', side_effect=read):
            result = self.fixture.sync()
        stats = result['folders'][0]['batch_stats']
        self.assertGreater(stats['prefetched_rows'], 0)
        self.assertLessEqual(stats['prefetch_peak_rows'], 2)
        self.assertEqual(set(read_threads), {owner})
        self.assertEqual(len(self.fixture.reads), 6)
        self.assertEqual(result['indexed_total'], 6)

    def test_provider_failure_does_not_checkpoint_prefetched_mail(self):
        def failure(*args, **kwargs):
            self.wait_for_prefetch()
            raise httpx.ConnectError('fixture provider failure')
        with patch.object(app, 'embed', side_effect=failure):
            result = self.fixture.sync()
        self.assertEqual(result['stop_reason'], 'embedding_unavailable')
        self.assertEqual(self.fixture.sql.execute('SELECT count(*) FROM emails').fetchone()[0], 0)
        self.fixture.sync()
        self.assertEqual(self.fixture.sql.execute('SELECT count(*) FROM emails').fetchone()[0], 6)

    def test_cancel_keeps_published_batch_and_resumes_unpublished_prefetch(self):
        cancelled = threading.Event()
        original = app.read_outlook_mail
        def read(*args):
            mail = original(*args)
            if len(self.fixture.reads) == 3:
                cancelled.set()
            return mail
        def embed(texts, cfg, **kwargs):
            self.assertTrue(cancelled.wait(2))
            return [[1.,0.,0.,0.] for _ in texts]
        with patch.object(app,'read_outlook_mail',side_effect=read), patch.object(app,'embed',side_effect=embed):
            result = self.fixture.sync(_cancel_check=cancelled.is_set)
        self.assertEqual(result['stop_reason'], 'cancelled')
        self.assertEqual(result['indexed_total'], 2)
        self.assertEqual(self.fixture.sql.execute('SELECT count(*) FROM emails').fetchone()[0], 2)
        self.fixture.sync()
        self.assertEqual(self.fixture.sql.execute('SELECT count(*) FROM emails').fetchone()[0], 6)

    def test_prefetched_body_error_is_recorded_without_losing_following_mail(self):
        self.fixture.folder('inbox', 6, broken='inbox2')
        def embed(texts, cfg, **kwargs):
            self.wait_for_prefetch()
            return [[1.,0.,0.,0.] for _ in texts]
        with patch.object(app,'embed',side_effect=embed):
            result = self.fixture.sync()
        self.assertEqual(result['failed_total'], 1)
        self.assertEqual(result['indexed_total'], 5)
        self.assertTrue(result['folders'][0]['window_complete'])

    def test_read_ahead_preserves_order_before_scan_failure(self):
        def rows():
            yield 1
            yield 2
            raise ValueError('fixture scan failure')
        ahead = app.ReadAhead(rows(), lambda value:value, 3, lambda queued:True)
        for _ in range(3):
            ahead.fill_one()
        self.assertEqual(next(ahead), 1)
        self.assertEqual(next(ahead), 2)
        with self.assertRaises(ValueError):
            next(ahead)
