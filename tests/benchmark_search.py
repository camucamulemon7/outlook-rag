"""Opt-in local scalability/ANN benchmark with seeded synthetic unit vectors.

Default: 100k mail IDs, five chunks each, 1024 dimensions. Temporary storage
may exceed 2 GB. No API, Outlook account or production index is accessed.
"""

import argparse
import json
from pathlib import Path
import statistics
import tempfile
import time

import lancedb
from lancedb.index import BTree, IvfHnswFlat
import numpy as np
import pyarrow as pa

from outlook_rag import storage
from outlook_rag.search import vector_candidates


def measure(call, queries):
    samples, results = [], []
    for query in queries:
        started = time.perf_counter()
        results.append(call(query))
        samples.append((time.perf_counter() - started) * 1000)
    return dict(
        median_ms=round(statistics.median(samples), 2),
        p95_ms=round(float(np.percentile(samples, 95)), 2),
    ), results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--emails", type=int, default=100000)
    parser.add_argument("--chunks-per-mail", type=int, default=5)
    parser.add_argument("--dimensions", type=int, default=1024)
    parser.add_argument("--queries", type=int, default=5)
    args = parser.parse_args()
    if min(args.emails, args.chunks_per_mail, args.dimensions, args.queries) < 1:
        raise ValueError("Benchmark sizes must be positive")
    count = args.emails * args.chunks_per_mail
    rng = np.random.default_rng(20261013)
    queries = rng.standard_normal((args.queries, args.dimensions), dtype=np.float32)
    queries /= np.linalg.norm(queries, axis=1, keepdims=True)
    generation = np.random.default_rng(8080)

    def batches():
        for offset in range(0, count, 1000):
            size = min(1000, count - offset)
            vector = generation.standard_normal(
                (size, args.dimensions), dtype=np.float32
            )
            vector /= np.linalg.norm(vector, axis=1, keepdims=True)
            rows = [
                dict(
                    id=f"{i // args.chunks_per_mail + 1:064x}:{i % args.chunks_per_mail}",
                    email_id=f"{i // args.chunks_per_mail + 1:064x}",
                    store_id="synthetic",
                    folder="inbox" if i // args.chunks_per_mail % 2 == 0 else "archive",
                    sender_lc="synthetic@example.invalid",
                    received_date="2026-10-13",
                    vector=values.tolist(),
                )
                for i, values in zip(range(offset, offset + size), vector)
            ]
            yield pa.RecordBatch.from_pylist(
                rows, schema=storage.schema(args.dimensions)
            )

    with tempfile.TemporaryDirectory(prefix="outlook-rag-benchmark-") as directory:
        table = lancedb.connect(str(Path(directory) / "vectors")).create_table(
            "mail_chunks", batches()
        )
        table.create_index("folder", config=BTree())
        print(
            json.dumps(dict(stage="loaded", emails=args.emails, chunks=count)),
            flush=True,
        )
        allowed = [f"{i + 1:064x}" for i in range(0, args.emails, 2)]
        sharded = lambda q: vector_candidates(
            table, q.tolist(), 100, allowed, cfg=dict(search_exact=True)
        )
        native = lambda q: vector_candidates(
            table,
            q.tolist(),
            100,
            None,
            predicate="folder='inbox'",
            cfg=dict(search_exact=True),
        )
        native(queries[0])
        native_time, expected = measure(native, queries)
        shard_time, shards = measure(sharded, queries)
        assert all(
            [r["id"] for r in a] == [r["id"] for r in b]
            for a, b in zip(expected, shards)
        )
        report = dict(
            data="seeded_random_unit_vectors_only",
            emails=args.emails,
            chunks=count,
            dimensions=args.dimensions,
            queries=args.queries,
            lancedb_version=lancedb.__version__,
            cache="warm; no OS cache eviction",
            native_prefilter=native_time,
            legacy_id_shards=shard_time,
            eligible_emails=len(allowed),
            filtered_top100_agreement=True,
            vector_requests_per_round=dict(
                legacy=(len(allowed) + 999) // 1000, native=1
            ),
        )
        partitions = min(64, max(1, count // 256))
        table.create_index(
            "vector",
            config=IvfHnswFlat(distance_type="cosine", num_partitions=partitions),
        )
        print(json.dumps(dict(stage="ann_built", partitions=partitions)), flush=True)
        report["ann"] = {}
        for name, cfg in (
            ("sdk_default", {}),
            ("all_partitions_ef200", dict(search_nprobes=partitions, search_ef=200)),
        ):
            vector_candidates(
                table,
                queries[0].tolist(),
                100,
                None,
                predicate="folder='inbox'",
                cfg=cfg,
            )
            timing, rows = measure(
                lambda q: vector_candidates(
                    table, q.tolist(), 100, None, predicate="folder='inbox'", cfg=cfg
                ),
                queries,
            )
            timing["chunk_recall_at_100"] = statistics.mean(
                len({r["id"] for r in a} & {r["id"] for r in b}) / len(a)
                for a, b in zip(expected, rows)
            )
            timing["search_nprobes"] = cfg.get("search_nprobes")
            timing["search_ef"] = cfg.get("search_ef")
            report["ann"][name] = timing
        # Newly added rows remain searchable even before rebuilding ANN.
        extra = dict(
            id=f"{args.emails + 1:064x}:0",
            email_id=f"{args.emails + 1:064x}",
            store_id="synthetic",
            folder="inbox",
            sender_lc="synthetic@example.invalid",
            received_date="2026-10-13",
            vector=queries[0].tolist(),
        )
        table.add([extra])
        latest = vector_candidates(
            table,
            queries[0].tolist(),
            10,
            None,
            predicate="folder='inbox'",
            cfg=dict(search_nprobes=partitions, search_ef=200),
        )
        assert latest[0]["id"] == extra["id"]
        report["unindexed_new_row_retrieved"] = True
        report["scope"] = (
            "Local artificial vectors; no real-mail relevance or Windows performance claim."
        )
        print(json.dumps(report), flush=True)


if __name__ == "__main__":
    main()
