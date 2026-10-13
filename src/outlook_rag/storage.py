"""Local index access and offline recovery; SQLite owns publication state."""

from __future__ import annotations

from array import array
from contextlib import closing
import hashlib
import json
import math
from pathlib import Path
import re
import shutil
import sqlite3
from typing import Any
import uuid

import lancedb
from lancedb.index import BTree, IvfHnswFlat
import pyarrow as pa

METADATA_COLUMNS = {"store_id", "folder", "sender_lc", "received_date"}
IDENTIFIER = re.compile(
    r"(?<![A-Za-z0-9_-])[A-Za-z][A-Za-z0-9]{1,15}[-_][A-Za-z0-9]{2,16}(?![A-Za-z0-9_-])"
)


def identifiers(text: str) -> set[str]:
    return {match.group().lower() for match in IDENTIFIER.finditer(text)}


def ascii_lower(text: str) -> str:
    # Match SQLite's built-in lower(), which only folds ASCII characters.
    return text.translate(
        str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")
    )


def metadata(mail: Any) -> dict:
    values = dict(mail)
    received = values.get("received", "")
    return dict(
        store_id=values.get("store_id", ""),
        folder=values.get("folder", ""),
        sender_lc=ascii_lower(values.get("sender") or ""),
        received_date=received[:10] if received is not None else None,
    )


def schema(dimensions: int) -> pa.Schema:
    return pa.schema(
        [
            pa.field("id", pa.string()),
            pa.field("email_id", pa.string()),
            pa.field("vector", pa.list_(pa.float32(), dimensions)),
            *[pa.field(name, pa.string()) for name in sorted(METADATA_COLUMNS)],
        ]
    )


def table_name(sql: sqlite3.Connection) -> str:
    row = sql.execute("SELECT value FROM state WHERE key='vector_table'").fetchone()
    name = row[0] if row else "mail_chunks"
    if name != "mail_chunks" and not re.fullmatch(r"mail_chunks_[0-9a-f]{32}", name):
        raise ValueError("Invalid published vector table")
    return name


def current_table(sql: sqlite3.Connection, cfg: dict) -> Any:
    path = Path(cfg["data_dir"]) / "vectors"
    if not path.is_dir():
        raise ValueError(
            "Vector index missing; run rebuild-vectors from the embedding cache."
        )
    return lancedb.connect(str(path)).open_table(table_name(sql))


def open_index(cfg: dict, *, writable: bool = True) -> tuple[sqlite3.Connection, Any]:
    """Open an initialized index without DDL, WAL changes or identity writes."""
    from . import app

    path = Path(cfg["data_dir"]) / "metadata.sqlite"
    if not path.is_file():
        raise ValueError("Index not initialized. Run sync_emails first.")
    mode = "rw" if writable else "ro"
    sql = sqlite3.connect(
        path.resolve().as_uri() + "?mode=" + mode, uri=True, timeout=5
    )
    sql.row_factory = sqlite3.Row
    try:
        identity = sql.execute(
            "SELECT value FROM state WHERE key='embedding_identity'"
        ).fetchone()
        if not identity or identity[0] != app.index_identity(cfg):
            raise ValueError(
                "Index configuration changed. Select the matching data_dir and settings."
            )
        with app.reader_lock(cfg):
            sql.execute("BEGIN")
            try:
                table = current_table(sql, cfg)
                app.pin_vectors(sql, table)
            finally:
                sql.rollback()
        return sql, table
    except BaseException:
        sql.close()
        raise


def metadata_ready(sql: sqlite3.Connection, table: Any) -> bool:
    row = sql.execute("SELECT value FROM state WHERE key='vector_metadata'").fetchone()
    return bool(row and row[0] == "1" and METADATA_COLUMNS <= set(table.schema.names))


def vector_status(table: Any) -> dict:
    indexes = [index for index in table.list_indices() if "vector" in index.columns]
    stats = table.index_stats(indexes[0].name) if indexes else None
    return dict(
        search="ann" if stats else "exact",
        indexed_rows=stats.num_indexed_rows if stats else 0,
        unindexed_rows=stats.num_unindexed_rows if stats else table.count_rows(),
        index_type=stats.index_type if stats else None,
    )


def update_folder(
    sql: sqlite3.Connection,
    table: Any,
    folder: str,
    *,
    mail_id: str | None = None,
    store_id: str | None = None,
    old_folder: str | None = None,
) -> None:
    """Publish folder-only changes with the matching vector metadata version."""
    from . import app

    where, args = (
        ("id=?", [mail_id])
        if mail_id
        else ("store_id=? AND folder=?", [store_id, old_folder])
    )
    if not sql.execute(
        f"SELECT 1 FROM emails WHERE {where} AND folder<>? LIMIT 1", [*args, folder]
    ).fetchone():
        return
    with app.vector_publication(sql, table):
        if metadata_ready(sql, table):
            from .search import literal

            predicate = (
                "email_id=" + literal(mail_id)
                if mail_id
                else "store_id="
                + literal(store_id or "")
                + " AND folder="
                + literal(old_folder or "")
            )
            table.update(where=predicate, values={"folder": folder})
        with sql:
            sql.execute(f"UPDATE emails SET folder=? WHERE {where}", [folder, *args])
            app.publish_vectors(sql, table)


def _cached_rows(sql: sqlite3.Connection, dimensions: int) -> Any:
    """Bounded batches; every published chunk must have a valid cached vector."""
    cursor = sql.execute("""SELECT c.id,c.email_id,c.text,e.store_id,e.folder,e.sender,e.received
        FROM chunks c JOIN emails e ON e.id=c.email_id ORDER BY c.id""")
    while rows := cursor.fetchmany(500):
        digests = {hashlib.sha256(row["text"].encode()).hexdigest() for row in rows}
        values = {
            r[0]: r[1]
            for r in sql.execute(
                "SELECT hash,vector FROM embedding_cache WHERE hash IN ("
                + ",".join("?" for _ in digests)
                + ")",
                list(digests),
            )
        }
        decoded = {}
        for digest in digests:
            blob = values.get(digest)
            if blob is None or len(blob) != dimensions * 4:
                raise ValueError(
                    "Embedding cache incomplete; published vectors were not changed. Re-sync missing content."
                )
            vector = array("f")
            vector.frombytes(blob)
            if not all(math.isfinite(v) for v in vector) or not any(vector):
                raise ValueError(
                    "Embedding cache contains an invalid vector; published vectors were not changed."
                )
            decoded[digest] = list(vector)
        output = []
        for row in rows:
            digest = hashlib.sha256(row["text"].encode()).hexdigest()
            output.append(
                dict(
                    id=row["id"],
                    email_id=row["email_id"],
                    vector=decoded[digest],
                    **metadata(row),
                )
            )
        yield output


def rebuild_vectors(cfg: dict) -> dict:
    """Build a new generation from cached embeddings, then publish atomically."""
    from . import app
    from .features import stored_sql

    with app.mutation_lock(cfg), stored_sql(cfg) as source:
        if source is None:
            raise ValueError("Index not initialized")
        root = Path(cfg["data_dir"])
        db = lancedb.connect(str(root / "vectors"))
        name = "mail_chunks_" + uuid.uuid4().hex
        table = db.create_table(name, schema=schema(app.vector_dimensions(cfg)))
        published = False
        try:
            expected = source.execute("SELECT count(*) FROM chunks").fetchone()[0]
            for rows in _cached_rows(source, app.vector_dimensions(cfg)):
                table.add(rows)
            if table.count_rows() != expected:
                raise ValueError(
                    "Orphan chunks found; vector generation was not published."
                )
            for column in ("folder", "received_date"):
                table.create_index(column, config=BTree())
            if expected >= 256:
                table.create_index(
                    "vector",
                    config=IvfHnswFlat(
                        distance_type="cosine",
                        num_partitions=max(1, min(64, expected // 256)),
                    ),
                )
            with (
                app.reader_lock(cfg),
                closing(sqlite3.connect(root / "metadata.sqlite", timeout=5)) as sql,
            ):
                # No partial schema/index migration is visible to readers.
                with sql:
                    sql.execute(
                        "CREATE TABLE IF NOT EXISTS identifiers(identifier TEXT,chunk_id TEXT,PRIMARY KEY(identifier,chunk_id))"
                    )
                    sql.execute(
                        "CREATE INDEX IF NOT EXISTS identifiers_chunk ON identifiers(chunk_id)"
                    )
                    sql.execute("DELETE FROM identifiers")
                    sql.execute("DELETE FROM lexical")
                    cursor = source.execute("SELECT id,text FROM chunks")
                    while rows := cursor.fetchmany(500):
                        sql.executemany(
                            "INSERT OR IGNORE INTO identifiers VALUES (?,?)",
                            [
                                (identifier, row[0])
                                for row in rows
                                for identifier in identifiers(row[1])
                            ],
                        )
                        # Tokens were saved with each successful embedding request.
                        digests = {
                            hashlib.sha256(row[1].encode()).hexdigest() for row in rows
                        }
                        terms = {
                            r[0]: r[1]
                            for r in source.execute(
                                "SELECT hash,terms FROM embedding_cache WHERE hash IN ("
                                + ",".join("?" for _ in digests)
                                + ")",
                                list(digests),
                            )
                        }
                        sql.executemany(
                            "INSERT INTO lexical VALUES (?,?)",
                            [
                                (
                                    row[0],
                                    terms[hashlib.sha256(row[1].encode()).hexdigest()],
                                )
                                for row in rows
                            ],
                        )
                    sql.executemany(
                        "INSERT OR REPLACE INTO state VALUES (?,?)",
                        [
                            ("vector_table", name),
                            ("vector_version", str(table.version)),
                            ("vector_metadata", "1"),
                            ("identifiers_ready", "1"),
                        ],
                    )
                published = True
            return dict(
                rebuilt=True,
                rows=expected,
                embedding_requests=0,
                metadata_filters=True,
                note="Previous generations remain until maintain_index; keep a backup of SQLite and vectors together.",
            )
        finally:
            if not published:
                # An interruption can arrive after SQLite commits but before the
                # Python flag changes. Never delete a committed generation.
                if table_name(source) != name:
                    db.drop_table(name)


def doctor(cfg: dict) -> dict:
    """Inspect a consistent publication without changing either database."""
    from . import app

    try:
        sql, table = open_index(cfg, writable=False)
    except (ValueError, FileNotFoundError):
        from .features import stored_sql

        # A missing/broken vector publication must not hide cache recovery.
        # stored_sql still refuses a mismatched or unreadable canonical DB.
        with stored_sql(cfg) as source:
            if source is None:
                return dict(
                    healthy=False,
                    problems=["index_missing"],
                    cached_rebuild_available=False,
                    embedding_requests=0,
                )
            complete = True
            try:
                for _ in _cached_rows(source, app.vector_dimensions(cfg)):
                    pass
            except ValueError:
                complete = False
            return dict(
                healthy=False,
                problems=["vector_publication_unavailable"],
                chunks=source.execute("SELECT count(*) FROM chunks").fetchone()[0],
                cached_rebuild_available=complete,
                embedding_requests=0,
            )
    try:
        with app.indexed_snapshot(sql, table, cfg) as table:
            problems = []
            integrity = sql.execute("PRAGMA quick_check").fetchone()[0]
            if integrity != "ok":
                problems.append("sqlite_integrity")
            chunks = sql.execute("SELECT count(*) FROM chunks").fetchone()[0]
            rows = table.count_rows()
            if chunks != rows:
                problems.append("vector_count")
            if table.schema.field("vector").type.list_size != app.vector_dimensions(
                cfg
            ):
                problems.append("vector_dimensions")
            # Store IDs in a temporary table, not all vectors in Python memory.
            sql.execute(
                "CREATE TEMP TABLE checked_vectors(id TEXT PRIMARY KEY,email_id TEXT)"
            )
            native = metadata_ready(sql, table)
            columns = (
                ["id", "email_id", *sorted(METADATA_COLUMNS)]
                if native
                else ["id", "email_id"]
            )
            metadata_mismatch = False
            for batch in (
                table.search().select(columns).limit(None).to_batches(batch_size=1000)
            ):
                values = batch.to_pylist()
                sql.executemany(
                    "INSERT OR IGNORE INTO checked_vectors VALUES (?,?)",
                    [(r["id"], r["email_id"]) for r in values],
                )
                if native and values:
                    ids = list({r["email_id"] for r in values})
                    mails = {
                        r["id"]: metadata(r)
                        for r in sql.execute(
                            "SELECT id,store_id,folder,sender,received FROM emails WHERE id IN ("
                            + ",".join("?" for _ in ids)
                            + ")",
                            ids,
                        )
                    }
                    metadata_mismatch |= any(
                        {key: r[key] for key in METADATA_COLUMNS}
                        != mails.get(r["email_id"])
                        for r in values
                    )
            if metadata_mismatch:
                problems.append("vector_metadata")
            if sql.execute("""SELECT 1 FROM chunks c LEFT JOIN checked_vectors v ON c.id=v.id
                WHERE v.id IS NULL OR c.email_id<>v.email_id LIMIT 1""").fetchone():
                problems.append("vector_ids")
            if sql.execute(
                "SELECT 1 FROM chunks c LEFT JOIN emails e ON c.email_id=e.id WHERE e.id IS NULL LIMIT 1"
            ).fetchone():
                problems.append("orphan_chunks")
            lexical = sql.execute("SELECT count(*) FROM lexical").fetchone()[0]
            # chunk_id is UNINDEXED in FTS5. Joining chunks directly to lexical
            # would rescan the virtual table once per chunk on a large index.
            sql.execute("CREATE TEMP TABLE checked_lexical(id TEXT PRIMARY KEY)")
            cursor = sql.execute("SELECT chunk_id FROM lexical")
            while ids := cursor.fetchmany(10000):
                sql.executemany(
                    "INSERT OR IGNORE INTO checked_lexical VALUES (?)",
                    [(row[0],) for row in ids],
                )
            if (
                lexical != chunks
                or sql.execute(
                    "SELECT 1 FROM chunks c LEFT JOIN checked_lexical l ON c.id=l.id WHERE l.id IS NULL LIMIT 1"
                ).fetchone()
            ):
                problems.append("lexical_index")
            cache_complete = "orphan_chunks" not in problems
            try:
                for _ in _cached_rows(sql, app.vector_dimensions(cfg)):
                    pass
            except ValueError:
                cache_complete = False
                problems.append("embedding_cache")
            return dict(
                healthy=not problems,
                problems=problems,
                emails=sql.execute("SELECT count(*) FROM emails").fetchone()[0],
                chunks=chunks,
                vector_rows=rows,
                cached_rebuild_available=cache_complete,
                metadata_filters=native,
                live_outlook_access=False,
                embedding_requests=0,
            )
    finally:
        sql.close()


def backup_index(cfg: dict, destination: str) -> dict:
    """Pair SQLite's backup API with its pinned vector generation."""
    from . import app

    target = Path(destination).expanduser().resolve()
    root = Path(cfg["data_dir"]).resolve()
    if target == root or root in target.parents:
        raise ValueError("Backup destination must be outside data_dir")
    with app.mutation_lock(cfg):
        sql, table = open_index(cfg, writable=False)
        try:
            target.mkdir(parents=True, exist_ok=False)
            try:
                with app.indexed_snapshot(sql, table, cfg):
                    with closing(sqlite3.connect(target / "metadata.sqlite")) as copy:
                        sql.backup(copy)
                    name = table_name(sql)
                    shutil.copytree(
                        root / "vectors" / (name + ".lance"),
                        target / "vectors" / (name + ".lance"),
                    )
                return dict(
                    backed_up=True, destination=str(target), embedding_requests=0
                )
            except BaseException:
                shutil.rmtree(target)
                raise
        finally:
            sql.close()


def restore_index(cfg: dict, source: str, destination: str) -> dict:
    """Restore only to a new directory; verify before presenting it as usable."""
    origin, target = (
        Path(source).expanduser().resolve(),
        Path(destination).expanduser().resolve(),
    )
    if target == origin or origin in target.parents:
        raise ValueError("Restore destination must be outside the backup directory")
    if target.exists():
        raise ValueError("Restore destination already exists")
    if not (origin / "metadata.sqlite").is_file() or not (origin / "vectors").is_dir():
        raise ValueError("Backup requires metadata.sqlite and vectors")
    if not cfg.get("dimensions"):
        with closing(
            sqlite3.connect(
                (origin / "metadata.sqlite").as_uri() + "?mode=ro", uri=True
            )
        ) as sql:
            identity = json.loads(
                sql.execute(
                    "SELECT value FROM state WHERE key='embedding_identity'"
                ).fetchone()[0]
            )
        cfg = dict(cfg, dimensions=identity["dimensions"])
        if "qwen3-embedding" in cfg["model"].lower():
            cfg.setdefault("storage_dimensions", min(1024, cfg["dimensions"]))
    target.mkdir(parents=True, exist_ok=False)
    try:
        # Copy only index files, never worker jobs, credentials or locks.
        shutil.copy2(origin / "metadata.sqlite", target / "metadata.sqlite")
        shutil.copytree(origin / "vectors", target / "vectors")
        result = doctor(dict(cfg, data_dir=str(target)))
        if not result["healthy"]:
            raise ValueError("Restored index failed integrity checks")
        return dict(
            restored=True,
            data_dir=str(target),
            chunks=result["chunks"],
            embedding_requests=0,
        )
    except BaseException:
        shutil.rmtree(target)
        raise
