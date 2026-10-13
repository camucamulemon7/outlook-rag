"""Canonical schema, locks and paired SQLite/vector publication.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

from contextlib import contextmanager
import sqlite3
from pathlib import Path
from typing import Any, Iterator

import lancedb
from filelock import FileLock

from . import storage, threads, mail_metadata, attachments

def mutation_lock(backend: Any, cfg: dict) -> FileLock:
    root = Path(cfg["data_dir"])
    root.mkdir(parents=True, exist_ok=True)
    return FileLock(root / "mutation.lock", timeout=1)


def reader_lock(backend: Any, cfg: dict) -> FileLock:
    # Maintenance must not retire a version between a reader's SQL snapshot
    # and its Lance checkout. Ordinary publication does not take this lock.
    return FileLock(Path(cfg["data_dir"]) / "reader.lock", timeout=30)


def pin_vectors(backend: Any, sql: sqlite3.Connection, table: Any) -> int:
    row = sql.execute("SELECT value FROM state WHERE key='vector_version'").fetchone()
    if not row:
        raise ValueError("Index publication marker missing. Run rebuild-vectors with the matching settings.")
    version = int(row[0])
    table.checkout(version)
    return version


def publish_vectors(backend: Any, sql: sqlite3.Connection, table: Any) -> None:
    # Call inside the same SQL transaction that publishes mail/chunks.
    sql.execute("INSERT OR REPLACE INTO state VALUES ('vector_version', ?)", (str(table.version),))


@contextmanager
def indexed_snapshot(backend: Any, sql: sqlite3.Connection, table: Any, cfg: dict) -> Iterator[Any]:
    with backend.reader_lock(cfg):
        sql.execute("BEGIN")
        try:
            # A rebuild may have published a new table while embedding was in flight.
            if table.name != storage.table_name(sql):
                table = storage.current_table(sql, cfg)
            backend.pin_vectors(sql, table)
            yield table
        finally:
            sql.rollback()


@contextmanager
def vector_publication(backend: Any, sql: sqlite3.Connection, table: Any) -> Iterator[None]:
    """Serialized writer: recover staged changes, then publish with SQLite."""
    published = backend.pin_vectors(sql, table)
    table.checkout_latest()
    if table.version != published:
        table.restore(published)
    try:
        yield
    except BaseException:
        table.checkout(published)
        raise
    else:
        backend.pin_vectors(sql, table)


def open_store(backend: Any, cfg: dict) -> tuple[sqlite3.Connection, Any]:
    root = Path(cfg["data_dir"])
    root.mkdir(parents=True, exist_ok=True)
    sql = sqlite3.connect(root / "metadata.sqlite", timeout=30)
    sql.row_factory = sqlite3.Row
    sql.execute("PRAGMA journal_mode=WAL")
    sql.executescript("""
    CREATE TABLE IF NOT EXISTS emails (
      id TEXT PRIMARY KEY, entry_id TEXT NOT NULL, store_id TEXT NOT NULL,
      folder TEXT NOT NULL, subject TEXT, sender TEXT, recipients TEXT,
      received TEXT, modified TEXT, body TEXT, conversation_id TEXT, indexed_at TEXT);
    CREATE INDEX IF NOT EXISTS emails_folder ON emails(folder);
    CREATE TABLE IF NOT EXISTS chunks (id TEXT PRIMARY KEY, email_id TEXT NOT NULL, position INTEGER, text TEXT);
    CREATE INDEX IF NOT EXISTS chunks_email_position ON chunks(email_id, position);
    CREATE TABLE IF NOT EXISTS identifiers(identifier TEXT,chunk_id TEXT,PRIMARY KEY(identifier,chunk_id));
    CREATE INDEX IF NOT EXISTS identifiers_chunk ON identifiers(chunk_id);
    CREATE INDEX IF NOT EXISTS emails_received_date ON emails(substr(received,1,10));
    CREATE VIRTUAL TABLE IF NOT EXISTS lexical USING fts5(chunk_id UNINDEXED, terms);
    CREATE TABLE IF NOT EXISTS state (key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS embedding_cache (hash TEXT PRIMARY KEY, vector BLOB NOT NULL, terms TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS embedding_pending (email_id TEXT, hash TEXT, PRIMARY KEY(email_id, hash));
    CREATE TABLE IF NOT EXISTS query_cache (hash TEXT PRIMARY KEY, vector BLOB NOT NULL, last_used REAL NOT NULL);
    CREATE TABLE IF NOT EXISTS sync_seen (scan_key TEXT, mail_id TEXT, PRIMARY KEY(scan_key, mail_id));
    CREATE TABLE IF NOT EXISTS sync_failures (mail_id TEXT PRIMARY KEY, entry_id TEXT, store_id TEXT,
      folder TEXT, modified TEXT, attempts INTEGER NOT NULL, error_type TEXT, next_retry REAL, last_failed REAL);
    CREATE TABLE IF NOT EXISTS entry_aliases (session TEXT, store_id TEXT, short_id TEXT, received TEXT,
      canonical_id TEXT, last_used REAL, PRIMARY KEY(session,store_id,short_id,received));
    """)
    # Keep the original failure schema compatible with older background workers.
    sql.execute("CREATE TABLE IF NOT EXISTS sync_failure_details (mail_id TEXT PRIMARY KEY,error_detail TEXT,last_failed REAL)")
    sql.execute("""CREATE TRIGGER IF NOT EXISTS remove_sync_failure_details AFTER DELETE ON sync_failures
        BEGIN DELETE FROM sync_failure_details WHERE mail_id=OLD.mail_id; END""")
    sql.execute("CREATE INDEX IF NOT EXISTS emails_conversation_order ON emails(store_id,conversation_id,received,id)")
    threads.initialize(sql)
    mail_metadata.initialize(sql)
    attachments.initialize(sql)
    date_schema = sql.execute("SELECT value FROM state WHERE key='outlook_date_schema'").fetchone()
    if date_schema is None:
        # Legacy COM wall-clock values cannot be corrected safely from SQLite.
        # Reset cursors once and reread source dates on the next bounded sync.
        if sql.execute("SELECT 1 FROM emails LIMIT 1").fetchone():
            sql.execute("DELETE FROM state WHERE key LIKE 'sync:%'")
            sql.execute("DELETE FROM sync_seen")
            sql.execute("INSERT OR REPLACE INTO state VALUES ('outlook_dates_need_rescan','1')")
        sql.execute("INSERT INTO state VALUES ('outlook_date_schema','2')")
    # Refuse a model/configuration mismatch rather than mixing embedding spaces.
    identity = backend.index_identity(cfg)
    previous = sql.execute("SELECT value FROM state WHERE key='embedding_identity'").fetchone()
    if previous and previous[0] != identity:
        sql.close()
        raise ValueError("Index configuration changed. Use a new data_dir and re-index.")
    sql.execute("INSERT OR IGNORE INTO state VALUES ('embedding_identity', ?)", (identity,))
    sql.commit()
    vectors = lancedb.connect(str(root / "vectors"))
    try:
        table = vectors.open_table(storage.table_name(sql))
    except (FileNotFoundError, ValueError) as exc:
        if isinstance(exc, ValueError) and "was not found" not in str(exc):
            raise
        table = vectors.create_table(storage.table_name(sql), schema=storage.schema(backend.vector_dimensions(cfg)))
        if not sql.execute("SELECT 1 FROM chunks LIMIT 1").fetchone():
            with sql:
                sql.executemany("INSERT OR REPLACE INTO state VALUES (?,?)", [("vector_metadata", "1"), ("identifiers_ready", "1")])
    with backend.reader_lock(cfg):
        # Adopt existing indexes without changing their chunk IDs. A concurrent
        # initializer may win; always reread the marker before checking out.
        with sql:
            sql.execute("INSERT OR IGNORE INTO state VALUES ('vector_version', ?)", (str(table.version),))
        backend.pin_vectors(sql, table)
    return sql, table

