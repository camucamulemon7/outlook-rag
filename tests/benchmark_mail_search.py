"""Opt-in end-to-end scale benchmark; all mail and document vectors are synthetic.

Use --emails 1000000 --chunks-per-mail 5 for five million 1024D chunks.
Allow >50 GiB temporary disk space and time to build ANN. No Outlook/profile
or production index is accessed. Optional query embeddings use a real API;
credentials must be supplied in OUTLOOK_RAG_API_KEY, never command arguments.
"""

import argparse
from contextlib import nullcontext
import json
from pathlib import Path
import platform
import statistics
import subprocess
import tempfile
import time
from unittest.mock import patch

import lancedb
from lancedb.index import BTree, IvfHnswFlat
import numpy as np
import pyarrow as pa

from outlook_rag import app, search, storage, threads

QUERIES = ["project meeting", "rare", "delivery budget"]
TEXT = "project meeting report budget invoice delivery"


def load(sql, cfg, emails, chunks_per_mail, thread_comparison=False):
    terms, rare = app.tokens(TEXT), app.tokens("rare")
    for offset in range(0, emails, 10000):
        mails = [
            (
                f"{i + 1:064x}",
                str(i),
                "primary-synthetic" if i % 2 == 0 else "pst-synthetic",
                "inbox" if i % 2 == 0 else "archive",
                "Synthetic mail",
                "synthetic@example.invalid",
                f"{2007 + i % 20}-10-13",
                TEXT,
                f"conversation-{i // 20}" if thread_comparison else "",
            )
            for i in range(offset, min(emails, offset + 10000))
        ]
        chunks = [
            (f"{mail[0]}:{position}", mail[0], position, TEXT)
            for mail in mails
            for position in range(chunks_per_mail)
        ]
        with sql:
            sql.executemany(
                "INSERT INTO emails(id,entry_id,store_id,folder,subject,sender,received,body,conversation_id) VALUES (?,?,?,?,?,?,?,?,?)",
                mails,
            )
            sql.executemany("INSERT INTO chunks VALUES (?,?,?,?)", chunks)
            sql.executemany(
                "INSERT INTO lexical VALUES (?,?)",
                [
                    (cid, terms + (" " + rare if mid.endswith("01") else ""))
                    for cid, mid, _, text in chunks
                ],
            )
        if (offset + 10000) % 100000 == 0:
            print(
                json.dumps(
                    dict(stage="sql_loaded", emails=min(offset + 10000, emails))
                ),
                flush=True,
            )
    with sql:
        sql.execute("INSERT INTO lexical(lexical) VALUES ('optimize')")
        # Direct synthetic loading bypasses the normal derived-thread updater.
        sql.execute("UPDATE state SET value='0' WHERE key='thread_index_complete'")
    sql.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    generator = np.random.default_rng(8080)
    schema = storage.schema(cfg["storage_dimensions"])
    count = emails * chunks_per_mail

    def batches():
        for offset in range(0, count, 10000):
            size = min(10000, count - offset)
            indices = range(offset, offset + size)
            matrix = generator.standard_normal(
                (size, cfg["storage_dimensions"]), dtype=np.float32
            )
            matrix /= np.linalg.norm(matrix, axis=1, keepdims=True)
            mail_numbers = [i // chunks_per_mail for i in indices]
            values = dict(
                id=pa.array(
                    [
                        f"{i // chunks_per_mail + 1:064x}:{i % chunks_per_mail}"
                        for i in indices
                    ]
                ),
                email_id=pa.array([f"{i + 1:064x}" for i in mail_numbers]),
                vector=pa.FixedSizeListArray.from_arrays(
                    pa.array(matrix.reshape(-1)), cfg["storage_dimensions"]
                ),
                folder=pa.array(
                    ["inbox" if i % 2 == 0 else "archive" for i in mail_numbers]
                ),
                store_id=pa.array(
                    [
                        "primary-synthetic" if i % 2 == 0 else "pst-synthetic"
                        for i in mail_numbers
                    ]
                ),
                sender_lc=pa.array(["synthetic@example.invalid"] * size),
                received_date=pa.array(
                    [f"{2007 + i % 20}-10-13" for i in mail_numbers]
                ),
            )
            yield pa.RecordBatch.from_arrays(
                [values[field.name] for field in schema], schema=schema
            )
            if (offset + size) % 500000 == 0:
                print(
                    json.dumps(dict(stage="vectors_loaded", rows=offset + size)),
                    flush=True,
                )

    table = lancedb.connect(str(Path(cfg["data_dir"]) / "vectors")).create_table(
        "mail_chunks", batches(), mode="overwrite"
    )
    for column in ("folder", "received_date"):
        table.create_index(column, config=BTree())
    print(json.dumps(dict(stage="loaded", emails=emails, chunks=count)), flush=True)
    if count >= 256:
        table.create_index(
            "vector",
            config=IvfHnswFlat(
                distance_type="cosine", num_partitions=min(64, max(1, count // 256))
            ),
        )
    with sql:
        app.publish_vectors(sql, table)
    print(
        json.dumps(dict(stage="ann_ready", **storage.vector_status(table))), flush=True
    )


def run(args, cfg, sql):
    # Prime only query embeddings. Full searches reopen SQL/Lance handles.
    embedding_ms = []
    for query in QUERIES:
        started = time.perf_counter()
        app.cached_query_vector(query, sql, cfg)
        embedding_ms.append(round((time.perf_counter() - started) * 1000, 2))
    report = dict(
        emails=args.emails,
        chunks=args.emails * args.chunks_per_mail,
        dimensions=args.dimensions,
        platform=platform.platform(),
        lancedb_version=lancedb.__version__,
        query_embedding_ms=embedding_ms,
        query_embedding_source="real API" if args.embedding_url else "synthetic",
        data="random unit vectors and synthetic text; no relevance claim",
        cache="query embeddings primed; no OS cache eviction; handles reopened each search",
        pipeline={},
    )
    with patch.object(app, "local_config", return_value=cfg):
        for mode in ("semantic", "hybrid", "lexical"):
            for scope, filters in (
                ("all_dates", {}),
                ("recent", dict(since="2026-01-01")),
            ):
                samples = []
                for query in QUERIES:
                    result = app.search(query, mode=mode, **filters)
                    if args.emails >= 20 and query != "rare":
                        assert result["count"] > 0, (mode, scope, query)
                    samples.append(
                        dict(
                            query=query,
                            elapsed_ms=result["elapsed_ms"],
                            embedding_ms=result["embedding_ms"],
                            count=result["count"],
                            candidate_limit=result["candidate_limit"],
                        )
                    )
                    print(
                        json.dumps(
                            dict(stage="search", mode=mode, scope=scope, **samples[-1])
                        ),
                        flush=True,
                    )
                report["pipeline"][mode + "_" + scope] = dict(
                    median_ms=statistics.median([s["elapsed_ms"] for s in samples]),
                    max_ms=max(s["elapsed_ms"] for s in samples),
                    samples=samples,
                )
        with patch.object(
            app, "local_config", return_value=dict(cfg, query_cache_size=0)
        ):
            result = app.search(QUERIES[0], mode="hybrid")
            report["uncached_hybrid"] = dict(
                elapsed_ms=result["elapsed_ms"],
                embedding_ms=result["embedding_ms"],
                count=result["count"],
            )
    print(json.dumps(report), flush=True)


def compare_threads(args, cfg, sql, table, built):
    """Same scale index, actual old search source, interleaved measured runs."""
    source = subprocess.check_output(
        ["git", "show", "427ee3d:src/outlook_rag/search.py"], text=True
    )
    namespace = dict(__name__="outlook_rag.benchmark_before", __package__="outlook_rag")
    exec(compile(source, "<baseline search>", "exec"), namespace)
    for query in QUERIES:
        app.cached_query_vector(query, sql, cfg)
    report = dict(
        emails=args.emails,
        chunks=args.emails * args.chunks_per_mail,
        data="synthetic text and random document vectors; latency only, no relevance claim",
        baseline="427ee3d",
        thread_documents=sql.execute(
            "SELECT count(*) FROM thread_documents"
        ).fetchone()[0],
        vector_coverage=storage.vector_status(table),
        thread_build=built,
        cache="query embeddings primed; OS cache retained; SQL/Lance handles reopened",
        pipeline={},
    )
    for scope, filters in [("all", {}), ("recent", dict(since="2026-01-01"))]:
        samples = {
            name: []
            for name in (
                "baseline_hybrid",
                "current_hybrid",
                "current_thread",
                "current_metadata",
            )
        }
        for repetition in range(2):
            for number, query in enumerate(QUERIES):
                order = list(samples)
                if (repetition + number) % 2:
                    order.reverse()
                for name in order:
                    started = time.perf_counter()
                    with patch.object(app, "local_config", return_value=cfg):
                        if name == "baseline_hybrid":
                            result = namespace["search"](
                                app, query, mode="hybrid", **filters
                            )
                        else:
                            result = app.search(
                                "" if name == "current_metadata" else query,
                                mode=name.removeprefix("current_"),
                                **filters,
                            )
                    sample = dict(
                        query=query,
                        ms=round((time.perf_counter() - started) * 1000, 2),
                        count=result["count"],
                    )
                    assert result["count"] > 0
                    if name == "current_thread":
                        assert result["thread_context"]["body_chars"] <= 12000
                    samples[name].append(sample)
                    print(
                        json.dumps(
                            dict(
                                stage="thread_comparison",
                                scope=scope,
                                mode=name,
                                **sample,
                            )
                        ),
                        flush=True,
                    )
        report["pipeline"][scope] = {
            name: dict(
                median_ms=statistics.median(s["ms"] for s in values),
                max_ms=max(s["ms"] for s in values),
                samples=values,
            )
            for name, values in samples.items()
        }
    if args.output:
        Path(args.output).write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report), flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--emails", type=int, default=100000)
    parser.add_argument("--chunks-per-mail", type=int, default=5)
    parser.add_argument("--dimensions", type=int, default=1024)
    parser.add_argument("--embedding-url")
    parser.add_argument("--model")
    parser.add_argument("--output")
    parser.add_argument(
        "--thread-comparison",
        action="store_true",
        help="Build derived thread documents and compare old/current hybrid, thread and SQL metadata on the same corpus",
    )
    args = parser.parse_args()
    if min(args.emails, args.chunks_per_mail, args.dimensions) < 1:
        parser.error("Benchmark sizes must be positive")
    if bool(args.embedding_url) != bool(args.model):
        parser.error("Supply both --embedding-url and --model")
    with tempfile.TemporaryDirectory(prefix="outlook-rag-mail-benchmark-") as root:
        cfg = dict(
            model=args.model or "synthetic",
            dimensions=args.dimensions,
            storage_dimensions=args.dimensions,
            data_dir=root,
            folders="auto",
            embedding_url=args.embedding_url or "http://synthetic.invalid",
        )
        if args.embedding_url:
            cfg["dimensions"] = app.detect_dimensions(cfg)
            app.vector_dimensions(cfg)
        generator = np.random.default_rng(20261013)

        def synthetic_embedding(texts, *unused, **options):
            matrix = generator.standard_normal(
                (len(texts), args.dimensions), dtype=np.float32
            )
            matrix /= np.linalg.norm(matrix, axis=1, keepdims=True)
            return matrix.tolist()

        with (
            nullcontext()
            if args.embedding_url
            else patch.object(app, "embed", side_effect=synthetic_embedding)
        ):
            sql, table = app.open_store(cfg)
            try:
                load(
                    sql, cfg, args.emails, args.chunks_per_mail, args.thread_comparison
                )
                if args.thread_comparison:
                    print(json.dumps(dict(stage="building_threads")), flush=True)
                    with patch.object(
                        app,
                        "embed",
                        side_effect=AssertionError("thread build called provider"),
                    ):
                        built = threads.rebuild(cfg)
                    print(json.dumps(dict(stage="threads_ready", **built)), flush=True)
                    compare_threads(args, cfg, sql, table, built)
                else:
                    run(args, cfg, sql)
            finally:
                sql.close()


if __name__ == "__main__":
    main()
