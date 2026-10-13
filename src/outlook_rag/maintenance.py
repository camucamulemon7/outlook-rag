"""Local vector optimization and cache/version maintenance.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

import hashlib
import re
import time
from datetime import timedelta
from pathlib import Path
from typing import Any

import lancedb
from lancedb.index import IvfHnswFlat

from . import jobs, storage

def optimize(backend: Any) -> dict:
    cfg = backend.local_config()
    with backend.mutation_lock(cfg):
        sql, table = storage.open_index(cfg)
        try:
            count = table.count_rows()
            if count < 256:
                return {"indexed": False, "rows": count, "note": "Small index uses exact vector search; ANN enabled from 256 chunks."}
            with backend.vector_publication(sql, table):
                table.create_index("vector", config=IvfHnswFlat(distance_type="cosine", num_partitions=max(1, min(64, count // 256))))
                with sql:
                    backend.publish_vectors(sql, table)
            return {"indexed": True, "rows": count, "indexes": [str(index) for index in table.list_indices()]}
        finally:
            sql.close()


def maintain(backend: Any) -> dict:
    """Reclaim unreferenced caches and compact Lance versions, retaining seven days."""
    cfg = backend.local_config()
    with backend.mutation_lock(cfg):
        sql, table = storage.open_index(cfg)
        try:
            sql.execute("CREATE TEMP TABLE active_hashes(hash TEXT PRIMARY KEY)")
            cursor = sql.execute("SELECT text FROM chunks")
            while True:
                rows = cursor.fetchmany(1000)
                if not rows:
                    break
                sql.executemany("INSERT OR IGNORE INTO active_hashes VALUES (?)",
                    [(hashlib.sha256(row[0].encode()).hexdigest(),) for row in rows])
            with sql:
                sql.execute("INSERT OR IGNORE INTO active_hashes SELECT hash FROM embedding_pending")
                removed = sql.execute("DELETE FROM embedding_cache WHERE hash NOT IN (SELECT hash FROM active_hashes)").rowcount
                removed_aliases = sql.execute("DELETE FROM entry_aliases WHERE last_used<?", (time.time() - 7 * 86400,)).rowcount
                sql.execute("DELETE FROM query_cache WHERE hash NOT IN (SELECT hash FROM query_cache ORDER BY last_used DESC LIMIT ?)", (cfg.get("query_cache_size", 256),))
            sql.execute("PRAGMA wal_checkpoint(PASSIVE)")
            # Never delete recent/unverified versions used by concurrent readers.
            with backend.reader_lock(cfg):
                # Protect the rollback version even if it was published more
                # than seven days ago. Keep the tag through failed maintenance.
                published = backend.pin_vectors(sql, table)
                tag = "outlook-rag-published"
                if tag in table.tags.list():
                    table.tags.update(tag, published)
                else:
                    table.tags.create(tag, published)
                with backend.vector_publication(sql, table):
                    result = table.optimize(cleanup_older_than=timedelta(days=7), delete_unverified=False)
                    with sql:
                        backend.publish_vectors(sql, table)
                vectors = lancedb.connect(str(Path(cfg["data_dir"]) / "vectors"))
                current = storage.table_name(sql)
                for name in vectors.list_tables().tables:
                    if name != current and re.fullmatch(r"mail_chunks(?:_[0-9a-f]{32})?", name):
                        vectors.drop_table(name)
            removed_logs = jobs.prune_logs(cfg)
            return {"removed_worker_logs": removed_logs, "removed_cached_embeddings": removed, "removed_entry_aliases": removed_aliases, "vector_maintenance": str(result), "retention_days": 7}
        finally:
            sql.close()

