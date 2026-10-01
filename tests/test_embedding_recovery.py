"""Completed embedding requests survive interruption and process restarts."""
import hashlib
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import httpx
from outlook_rag import app


class EmbeddingRecovery(unittest.TestCase):
    def setUp(self):
        Path('work').mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir='work', ignore_cleanup_errors=True)
        self.cfg = dict(model='fixture', dimensions=4, data_dir=self.temp.name,
                        embedding_url='http://fixture.invalid', embedding_batch_size=8,
                        embedding_concurrent_requests=1)
        self.sql, self.table = app.open_store(self.cfg)
        self.parts = [f'Unique chunk {i}' for i in range(24)]
        self.mail = dict(id=hashlib.sha256(b'fixture').hexdigest(), subject='fixture', body='complete body')

    def tearDown(self):
        self.sql.close()
        self.temp.cleanup()

    def count(self, name):
        return self.sql.execute('SELECT count(*) FROM '+name).fetchone()[0]

    @staticmethod
    def vectors(texts, *_args, **_kwargs):
        return [[1., 0., 0., 0.] for _ in texts]

    def test_partial_success_survives_restart_and_maintenance(self):
        calls = []
        def interrupt(texts, *_args, **_kwargs):
            calls.append(texts)
            if len(calls) == 2:
                raise httpx.ReadTimeout('fixture deadline')
            return self.vectors(texts)
        with patch.object(app, 'chunks', return_value=self.parts), patch.object(app, 'embed', side_effect=interrupt):
            with self.assertRaises(httpx.ReadTimeout):
                app.upsert_mails([self.mail], self.sql, self.table, self.cfg)
        self.assertEqual(self.count('embedding_cache'), 8)
        self.assertEqual(self.count('emails'), 0)
        self.assertEqual(self.count('chunks'), 0)
        self.assertEqual(self.table.count_rows(), 0)
        self.sql.close()
        self.sql, self.table = app.open_store(self.cfg)
        with patch.object(app, 'config', return_value=self.cfg):
            self.assertEqual(app.maintain()['removed_cached_embeddings'], 0)
        calls.clear()
        def resume(texts, *_args, **_kwargs):
            calls.extend(texts)
            return self.vectors(texts)
        with patch.object(app, 'chunks', return_value=self.parts), patch.object(app, 'embed', side_effect=resume):
            stats = app.upsert_mails([self.mail], self.sql, self.table, self.cfg)
        self.assertEqual(set(calls), set(self.parts[8:]))
        self.assertEqual(stats['cache_hits'], 8)
        self.assertEqual(stats['embedded_chunks'], 16)
        self.assertEqual(self.count('emails'), 1)
        self.assertEqual(self.count('chunks'), 24)
        self.assertEqual(self.table.count_rows(), 24)
        self.assertEqual(self.count('embedding_pending'), 0)

    def test_successful_inflight_request_survives_earlier_failure(self):
        started = threading.Event()
        calls = []
        def parallel(texts, *_args, **_kwargs):
            calls.append(texts)
            if texts[0] == self.parts[0]:
                self.assertTrue(started.wait(2))
                raise httpx.ReadTimeout('fixture failed request')
            started.set()
            time.sleep(0.05)
            return self.vectors(texts)
        cfg = dict(self.cfg, embedding_concurrent_requests=2)
        with patch.object(app, 'chunks', return_value=self.parts), patch.object(app, 'embed', side_effect=parallel):
            with self.assertRaises(httpx.ReadTimeout):
                app.upsert_mails([self.mail], self.sql, self.table, cfg)
        self.assertEqual(len(calls), 2)
        self.assertEqual(self.count('embedding_cache'), 8)
        self.assertEqual(self.count('emails'), 0)
        cached = {row[0] for row in self.sql.execute('SELECT hash FROM embedding_cache')}
        self.assertEqual(cached, {hashlib.sha256(text.encode()).hexdigest() for text in self.parts[8:16]})

    def test_interrupted_update_preserves_published_mail(self):
        with patch.object(app, 'chunks', return_value=['old chunk']), patch.object(app, 'embed', side_effect=self.vectors):
            app.upsert_mails([dict(self.mail, body='old body')], self.sql, self.table, self.cfg)
        calls = []
        def interrupt(texts, *_args, **_kwargs):
            calls.append(texts)
            if len(calls) == 2:
                raise httpx.ReadTimeout('fixture deadline')
            return self.vectors(texts)
        with patch.object(app, 'chunks', return_value=self.parts), patch.object(app, 'embed', side_effect=interrupt):
            with self.assertRaises(httpx.ReadTimeout):
                app.upsert_mails([self.mail], self.sql, self.table, self.cfg)
        self.assertEqual(self.sql.execute('SELECT body FROM emails').fetchone()[0], 'old body')
        self.assertEqual(self.sql.execute('SELECT text FROM chunks').fetchone()[0], 'old chunk')
        self.assertEqual(self.table.count_rows(), 1)
        self.assertEqual(self.count('embedding_cache'), 9)
        # Replacing a pending version releases its obsolete cache references.
        with patch.object(app, 'chunks', return_value=['replacement']), patch.object(app, 'embed', side_effect=self.vectors):
            app.upsert_mails([self.mail], self.sql, self.table, self.cfg)
        self.assertEqual(self.count('embedding_pending'), 0)
        with patch.object(app, 'config', return_value=self.cfg):
            self.assertEqual(app.maintain()['removed_cached_embeddings'], 9)


if __name__ == '__main__':
    unittest.main()
