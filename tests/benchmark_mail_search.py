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
import tempfile
import time
from unittest.mock import patch

import lancedb
from lancedb.index import BTree, IvfHnswFlat
import numpy as np
import pyarrow as pa

from outlook_rag import app, storage

QUERIES = ["project meeting", "rare", "delivery budget"]
TEXT = "project meeting report budget invoice delivery"


def load(sql, cfg, emails, chunks_per_mail):
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
                "INSERT INTO emails(id,entry_id,store_id,folder,subject,sender,received,body) VALUES (?,?,?,?,?,?,?,?)",
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
    with sql:
        sql.execute("INSERT INTO lexical(lexical) VALUES ('optimize')")
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--emails", type=int, default=100000)
    parser.add_argument("--chunks-per-mail", type=int, default=5)
    parser.add_argument("--dimensions", type=int, default=1024)
    parser.add_argument("--embedding-url")
    parser.add_argument("--model")
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
                load(sql, cfg, args.emails, args.chunks_per_mail)
                run(args, cfg, sql)
            finally:
                sql.close()


if __name__ == "__main__":
    main()
