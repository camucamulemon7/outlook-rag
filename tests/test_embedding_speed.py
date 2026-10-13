"""Tests verifying embedding speed optimizations:
- Thread-local Sudachi tokenization and concurrent safety
- Vectorized normalization (NumPy with pure-Python fallback parity)
- Batch cache lookup across large chunk collections
- Token pre-computation during embedding wait pipeline
"""
from __future__ import annotations

import hashlib
import math
import tempfile
import threading
from array import array
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
import unittest
from unittest.mock import Mock, patch

from outlook_rag import app
from outlook_rag import text as text_module
from outlook_rag.text import tokens


class EmbeddingSpeedOptimizations(unittest.TestCase):
    @contextmanager
    def store(self, **options):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as root:
            cfg = dict(model="fixture", dimensions=4, data_dir=root,
                       embedding_url="http://fixture.invalid", **options)
            sql, table = app.open_store(cfg)
            try:
                yield cfg, sql, table
            finally:
                sql.close()

    def test_concurrent_tokens_thread_safety_and_consistency(self):
        """Cold concurrent startup shares one dictionary and isolates tokenizer engines."""
        from sudachipy import dictionary
        texts = ["本日のプロジェクト会議議事録を送付します。", "進捗と課題を確認してください。", "Outlook RAG １２３"]
        expected = [tokens(text) for text in texts]
        barrier = threading.Barrier(4)
        def worker():
            barrier.wait(10)
            engine, _ = text_module._get_tokenizer()
            for _ in range(5):
                self.assertEqual([tokens(text) for text in texts], expected)
                self.assertIs(text_module._get_tokenizer()[0], engine)
            return engine
        with patch.object(text_module, "_DICTIONARY", None), patch.object(text_module, "_LOCAL", threading.local()), \
             patch.object(dictionary, "Dictionary", wraps=dictionary.Dictionary) as factory:
            with ThreadPoolExecutor(max_workers=4) as pool:
                engines = [future.result() for future in [pool.submit(worker) for _ in range(4)]]
            self.assertEqual(factory.call_count, 1)
            self.assertEqual(len({id(engine) for engine in engines}), 4)

    def test_vector_normalization_numpy_and_fallback_consistency(self):
        """Vector normalization returns identical outputs in both NumPy and fallback modes."""
        cfg = {"dimensions": 4, "storage_dimensions": 4, "model": "fixture",
               "embedding_url": "http://fixture.invalid/embeddings"}
        raw_rows = [{"index": 0, "embedding": [3.0, 4.0, 0.0, 0.0]},
                    {"index": 1, "embedding": [1.0, 1.0, 1.0, 1.0]}]

        mock_resp = Mock()
        mock_resp.status_code = 200
        mock_resp.is_success = True
        mock_resp.json.return_value = {"data": raw_rows}
        mock_resp.raise_for_status = Mock()

        mock_client = Mock()
        mock_client.post.return_value = mock_resp

        with patch.object(app, "api_key", return_value="fixture-key"):
            # Normal call (uses NumPy if available)
            res_numpy = app.embed(["a", "b"], cfg, client=mock_client)

            # Force ImportError to test fallback path
            with patch.dict("sys.modules", {"numpy": None}):
                res_fallback = app.embed(["a", "b"], cfg, client=mock_client)

        self.assertEqual(len(res_numpy), 2)
        self.assertEqual(len(res_fallback), 2)
        for vec_np, vec_py in zip(res_numpy, res_fallback):
            self.assertEqual(len(vec_np), 4)
            for v_np, v_py in zip(vec_np, vec_py):
                self.assertAlmostEqual(v_np, v_py, places=6)

        # Check normalization norm == 1.0
        for vec in res_numpy:
            norm = math.sqrt(sum(v * v for v in vec))
            self.assertAlmostEqual(norm, 1.0, places=6)

    def test_vector_normalization_rejects_zero_vector(self):
        """Zero vectors are properly rejected with ValueError."""
        cfg = {"dimensions": 4, "storage_dimensions": 4, "model": "fixture",
               "embedding_url": "http://fixture.invalid/embeddings"}
        mock_resp = Mock()
        mock_resp.status_code = 200
        mock_resp.is_success = True
        mock_resp.json.return_value = {"data": [{"index": 0, "embedding": [0.0, 0.0, 0.0, 0.0]}]}
        mock_resp.raise_for_status = Mock()

        mock_client = Mock()
        mock_client.post.return_value = mock_resp

        with patch.object(app, "api_key", return_value="fixture-key"):
            with self.assertRaisesRegex(ValueError, "Embedding API returned a zero vector"):
                app.embed(["a"], cfg, client=mock_client)

    def test_reduction_rejects_nonfinite_values_in_discarded_dimensions(self):
        cfg = dict(dimensions=4, storage_dimensions=2, model="fixture",
                   embedding_url="http://fixture.invalid/embeddings")
        for bad_value in (float("nan"), float("inf"), -float("inf")):
            for fallback in (False, True):
                with self.subTest(value=bad_value, fallback=fallback):
                    response = Mock(status_code=200)
                    response.json.return_value = {"data": [
                        {"index": 0, "embedding": [3., 4., bad_value, 0.]}
                    ]}
                    client = Mock()
                    client.post.return_value = response
                    modules = {"numpy": None} if fallback else {}
                    with patch.object(app, "api_key", return_value="fixture-key"), patch.dict("sys.modules", modules):
                        with self.assertRaisesRegex(ValueError, "dimensions or values"):
                            app.embed(["fixture"], cfg, client=client)

    def test_mixed_native_and_requested_dimensions_keep_response_order(self):
        cfg = dict(dimensions=4, storage_dimensions=2, embedding_request_dimensions=2,
                   model="mixed-dimensions-fixture", embedding_url="http://fixture.invalid/embeddings")
        response = Mock(status_code=200)
        response.json.return_value = {"data": [
            {"index": 1, "embedding": [0., 2., 3., 4.]},
            {"index": 0, "embedding": [3., 4.]},
        ]}
        client = Mock()
        client.post.return_value = response
        for fallback in (False, True):
            with self.subTest(fallback=fallback):
                app._API_DIMENSIONS.clear()
                modules = {"numpy": None} if fallback else {}
                with patch.object(app, "api_key", return_value="fixture-key"), patch.dict("sys.modules", modules):
                    self.assertEqual(app.embed(["a", "b"], cfg, client=client), [[.6, .8], [0., 1.]])

    def test_invalid_norm_never_returns_nonunit_vectors(self):
        cfg = dict(dimensions=4, storage_dimensions=2, model="fixture",
                   embedding_url="http://fixture.invalid/embeddings")
        for vector in ([0., 0., 3., 4.], [1e308, 1e308, 0., 0.]):
            for fallback in (False, True):
                with self.subTest(vector=vector, fallback=fallback):
                    response = Mock(status_code=200)
                    response.json.return_value = {"data": [{"index": 0, "embedding": vector}]}
                    client = Mock()
                    client.post.return_value = response
                    modules = {"numpy": None} if fallback else {}
                    with patch.object(app, "api_key", return_value="fixture-key"), patch.dict("sys.modules", modules):
                        with self.assertRaises(ValueError):
                            app.embed(["fixture"], cfg, client=client)

    def test_batch_cache_lookup_with_large_chunk_set(self):
        """Exercise actual upsert: >500 keys, duplicates, cache hits and misses."""
        with self.store(embedding_batch_size=8) as (cfg, sql, table):
            parts = [f"unique chunk {i}" for i in range(650)]
            digest = lambda text: hashlib.sha256(text.encode()).hexdigest()
            with sql:
                sql.executemany("INSERT INTO embedding_cache VALUES (?, ?, ?)", [
                    (digest(text), array("f", [1., 0., 0., 0.]).tobytes(), f"cached terms {i}")
                    for i, text in enumerate(parts[:600])
                ])
            queries = []
            sql.set_trace_callback(queries.append)
            mail = dict(id=digest("fixture mail"), subject="fixture", body="fixture")
            embedded, tokenized = [], []
            def embed(texts, *args, **kwargs):
                embedded.extend(texts)
                return [[0., 1., 0., 0.] for _ in texts]
            def tokenize(text):
                tokenized.append(text)
                return "new terms " + text
            with patch.object(app, "chunks", return_value=parts + [parts[0], parts[-1]]), \
                 patch.object(app, "embed", side_effect=embed), patch.object(app, "tokens", side_effect=tokenize):
                stats = app.upsert_mails([mail], sql, table, cfg)
            self.assertEqual(stats["cache_hits"], 601)
            self.assertEqual(stats["embedded_chunks"], 50)
            self.assertEqual(set(embedded), set(parts[600:]))
            self.assertCountEqual(tokenized, parts[600:])
            self.assertEqual(sql.execute("SELECT count(*) FROM embedding_cache").fetchone()[0], 650)
            self.assertEqual(table.count_rows(), 652)
            cache_queries = [query for query in queries if query.startswith("SELECT hash, vector, terms FROM embedding_cache")]
            self.assertEqual(len(cache_queries), 2)
            self.assertEqual(sql.execute("SELECT terms FROM lexical WHERE chunk_id=?", (mail["id"] + ":0",)).fetchone()[0], "cached terms 0")
            self.assertEqual(sql.execute("SELECT terms FROM lexical WHERE chunk_id=?", (mail["id"] + ":651",)).fetchone()[0], "new terms " + parts[-1])

    def test_tokenization_overlaps_http_on_caller_thread(self):
        with self.store(embedding_concurrent_requests=1) as (cfg, sql, table):
            owner = threading.get_ident()
            started, tokenized = threading.Event(), threading.Event()
            calls = []
            def embed(texts, *args, **kwargs):
                self.assertNotEqual(threading.get_ident(), owner)
                started.set()
                self.assertTrue(tokenized.wait(3), "Tokenization did not overlap HTTP")
                return [[1., 0., 0., 0.] for _ in texts]
            def tokenize(text):
                self.assertTrue(started.is_set())
                self.assertEqual(threading.get_ident(), owner)
                calls.append(text)
                tokenized.set()
                return "fixture terms"
            mail = dict(id=hashlib.sha256(b"fixture mail").hexdigest(), subject="fixture", body="fixture")
            with patch.object(app, "chunks", return_value=["fixture chunk"]), \
                 patch.object(app, "embed", side_effect=embed), patch.object(app, "tokens", side_effect=tokenize):
                stats = app.upsert_mails([mail], sql, table, cfg)
            self.assertEqual(calls, ["fixture chunk"])
            self.assertEqual(stats["embedded_chunks"], 1)
            self.assertEqual(sql.execute("SELECT terms FROM embedding_cache").fetchone()[0], "fixture terms")

    def test_lookahead_tokenization_error_preserves_earlier_embedding_cache(self):
        self.tokenization_error_recovery(workers=1)

    def test_tokenization_error_preserves_successful_inflight_embedding(self):
        self.tokenization_error_recovery(workers=2)

    def tokenization_error_recovery(self, workers):
        with self.store(embedding_concurrent_requests=workers, embedding_batch_size=1) as (cfg, sql, table):
            bad_token = threading.Event()
            def embed(texts, *args, **kwargs):
                if texts == ["good chunk"]:
                    self.assertTrue(bad_token.wait(3))
                return [[1., 0., 0., 0.] for _ in texts]
            def tokenize(text):
                if text == "bad chunk":
                    bad_token.set()
                    raise ValueError("fixture tokenizer failure")
                return "good terms"
            mail = dict(id=hashlib.sha256(b"fixture mail").hexdigest(), subject="fixture", body="fixture")
            with patch.object(app, "chunks", return_value=["good chunk", "bad chunk"]), \
                 patch.object(app, "embed", side_effect=embed), patch.object(app, "tokens", side_effect=tokenize):
                with self.assertRaisesRegex(ValueError, "fixture tokenizer failure"):
                    app.upsert_mails([mail], sql, table, cfg)
            self.assertEqual(sql.execute("SELECT count(*) FROM embedding_cache").fetchone()[0], 1)
            self.assertEqual(sql.execute("SELECT terms FROM embedding_cache").fetchone()[0], "good terms")
            self.assertEqual(sql.execute("SELECT count(*) FROM chunks").fetchone()[0], 0)
            self.assertEqual(table.count_rows(), 0)
            calls = []
            def resume(texts, *args, **kwargs):
                calls.extend(texts)
                return [[1., 0., 0., 0.] for _ in texts]
            with patch.object(app, "chunks", return_value=["good chunk", "bad chunk"]), \
                 patch.object(app, "embed", side_effect=resume), patch.object(app, "tokens", return_value="fixed terms"):
                stats = app.upsert_mails([mail], sql, table, cfg)
            self.assertEqual(calls, ["bad chunk"])
            self.assertEqual(stats["cache_hits"], 1)
            self.assertEqual(table.count_rows(), 2)
            self.assertEqual(sql.execute("SELECT count(*) FROM embedding_pending").fetchone()[0], 0)


if __name__ == "__main__":
    unittest.main()
