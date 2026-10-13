from __future__ import annotations

from array import array
from collections import deque
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from contextlib import contextmanager
import hashlib
import json
import logging
import math
import os
import re
import sqlite3
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterator

import httpx
import lancedb
from lancedb.index import IvfHnswFlat
from filelock import FileLock
from mcp.server.fastmcp import FastMCP

from . import cli, coverage, jobs, search as search_engine, server, storage, threads
from .features import local_config, stored_sql, folder_progress, failure_reason
from .outlook import (
    safe_attr as safe_attr,
    iso as iso,
    same_modified as same_modified,
    resolve_folder as resolve_folder,
    folder_identity as folder_identity,
    discover_sources as discover_sources,
    mail_metadata as mail_metadata,
    outlook_session_key as outlook_session_key,
    canonical_metadata as canonical_metadata,
    read_outlook_mail as read_outlook_mail,
)
from .text import chunks as chunks, clean_body as clean_body, safe_text as safe_text, tokens as tokens

QUERY_INSTRUCTION = "Retrieve email passages relevant to the user's question, including paraphrases and related concepts."
logging.getLogger("httpx").setLevel(logging.WARNING)
_LOCK = threading.RLock()
_DIMENSIONS: dict[tuple, int] = {}
_API_DIMENSIONS: dict[tuple, bool] = {}


class SyncCancelled(httpx.RequestError):
    """Cooperative cancellation; completed embedding requests remain cached."""


class ReadAhead:
    """Bounded caller-thread prefetch; COM and SQLite never cross threads."""
    def __init__(self, source: Any, prepare: Any, limit: int, allowed: Any) -> None:
        self.source, self.prepare = iter(source), prepare
        self.limit, self.allowed = limit, allowed
        self.buffer: deque[Any] = deque()
        self.finished = False
        self.error: Exception | None = None
        self.read_ms = 0.0
        self.rows = 0
        self.peak = 0

    def __iter__(self) -> ReadAhead:
        return self

    def __next__(self) -> Any:
        if self.buffer:
            return self.buffer.popleft()
        if self.error is not None:
            raise self.error
        if self.finished:
            raise StopIteration
        return next(self.source)

    def can_fill(self) -> bool:
        return not self.finished and len(self.buffer) < self.limit and self.allowed(len(self.buffer))

    def fill_one(self) -> None:
        if not self.can_fill():
            return
        started = time.perf_counter()
        try:
            value = self.prepare(next(self.source))
        except StopIteration:
            self.finished = True
        except Exception as exc:
            # Surface metadata scan errors only after earlier buffered rows.
            self.finished, self.error = True, exc
        else:
            self.buffer.append(value)
            self.rows += 1
            self.peak = max(self.peak, len(self.buffer))
        finally:
            self.read_ms += (time.perf_counter() - started) * 1000

def config(*, resolve_dimensions: bool = True) -> dict:
    path = os.environ.get("OUTLOOK_RAG_CONFIG")
    config_path = Path(path).expanduser().resolve() if path else Path.cwd() / "config.json"
    cfg = json.loads(config_path.read_text(encoding="utf-8-sig")) if path else {}
    converters: dict[str, Callable[[str], Any]] = {"dimensions": lambda value: None if value.lower() == "auto" else int(value), "storage_dimensions": int, "embedding_batch_size": int,
                  "embedding_concurrent_requests": int, "embedding_max_batch_chars": int,
                  "sync_prefetch_emails": int, "sync_batch_emails": int, "sync_max_emails": int, "sync_max_total_emails": int, "text_cleaning_version": int,
                  "since_days": lambda value: 0 if value.lower() == "all" else int(value),
                  "folders": lambda value: "auto" if value.lower() == "auto" else json.loads(value),
                  "excluded_folders": json.loads, "exclude_system_folders": json.loads,
                  "chunk_size": int, "chunk_overlap": int, "warmup_on_start": json.loads,
                  "sync_max_seconds": float, "sync_max_scanned": int, "sync_retry_limit": int,
                  "embedding_request_dimensions": int, "query_cache_size": int,
                  "search_candidates": int, "rrf_k": int, "semantic_weight": float, "lexical_weight": float,
                  "search_nprobes": int, "search_ef": int, "search_exact": json.loads}
    fields = ("embedding_url", "model", "data_dir", "key_file", "api_key_env", "query_instruction", "query_prefix", "document_prefix", "model_revision", "search_mode", *converters)
    for field in fields:
        value = os.environ.get("OUTLOOK_RAG_" + field.upper())
        if value is not None:
            cfg[field] = converters.get(field, str)(value)
    if not cfg.get("model"):
        raise ValueError("Set OUTLOOK_RAG_MODEL to your embedding model ID.")
    cfg.setdefault("embedding_url", "http://localhost:8080/api/v1/embeddings")
    automatic_data_dir = not cfg.get("data_dir")
    cfg["_automatic_data_dir"] = automatic_data_dir
    local_root = Path(os.environ["LOCALAPPDATA"]) if os.environ.get("LOCALAPPDATA") else Path.home() / "AppData" / "Local"
    if automatic_data_dir:
        index_root = Path.home() / "Documents" / "Outlook\u30d5\u30a1\u30a4\u30eb" / "outlook-rag"
        cfg["data_dir"] = str(index_root / "pending")
    cfg.setdefault("key_file", str(local_root / "outlook-rag" / "api-key.dpapi"))
    cfg.setdefault("folders", "auto")
    cfg["folders"] = coverage.folder_selection(cfg["folders"])
    cfg.setdefault("since_days", 0)
    cfg.setdefault("sync_max_emails", 200)
    cfg.setdefault("sync_max_total_emails", 200)
    cfg.setdefault("sync_max_seconds", 30)
    cfg.setdefault("sync_max_scanned", 2000)
    cfg.setdefault("sync_retry_limit", 10)
    cfg.setdefault("sync_prefetch_emails", 16)
    if type(cfg["sync_prefetch_emails"]) is not int or not 0 <= cfg["sync_prefetch_emails"] <= 256:
        raise ValueError("sync_prefetch_emails must be 0..256")
    cfg.setdefault("query_cache_size", 256)
    cfg.setdefault("search_candidates", 100)
    cfg.setdefault("rrf_k", 60)
    cfg.setdefault("semantic_weight", 1.0)
    cfg.setdefault("lexical_weight", 1.0)
    cfg.setdefault("search_mode", "hybrid")
    cfg.setdefault("search_exact", False)
    if cfg["search_mode"] not in {"semantic", "hybrid", "lexical", "auto", "identifier", "metadata", "thread"}:
        raise ValueError("Invalid search_mode")
    if type(cfg["search_exact"]) is not bool:
        raise ValueError("search_exact must be a boolean")
    for name in ("search_nprobes", "search_ef"):
        if cfg.get(name) is not None and (type(cfg[name]) is not int or not 1 <= cfg[name] <= 65536):
            raise ValueError(f"{name} must be 1..65536")
    if not 10 <= cfg["search_candidates"] <= 5000 or not 1 <= cfg["rrf_k"] <= 1000:
        raise ValueError("Invalid search candidate count or RRF constant")
    if not all(math.isfinite(cfg[name]) and 0 <= cfg[name] <= 10 for name in ("semantic_weight", "lexical_weight")) or not cfg["semantic_weight"] + cfg["lexical_weight"]:
        raise ValueError("Search weights must be finite, nonnegative, at most 10, and not both zero")
    if not 1 <= cfg["sync_max_seconds"] <= 3600 or not 1 <= cfg["sync_max_scanned"] <= 100000:
        raise ValueError("sync_max_seconds must be 1..3600; sync_max_scanned must be 1..100000")
    if not 0 <= cfg["sync_retry_limit"] <= 200 or not 0 <= cfg["query_cache_size"] <= 10000:
        raise ValueError("sync_retry_limit must be 0..200; query_cache_size must be 0..10000")
    cfg.setdefault("exclude_system_folders", True)
    cfg.setdefault("excluded_folders", [])
    cfg.setdefault("chunk_size", 1600)
    cfg.setdefault("chunk_overlap", 200)
    if not 100 <= cfg["chunk_size"] <= 16000 or not 0 <= cfg["chunk_overlap"] < cfg["chunk_size"]:
        raise ValueError("chunk_size must be 100..16000 and chunk_overlap must be smaller than chunk_size")
    if cfg["folders"] != "auto" and (not isinstance(cfg["folders"], list) or not all(isinstance(item,str) for item in cfg["folders"])):
        raise ValueError("folders must be 'auto' or a JSON array of folder paths")
    if not isinstance(cfg["excluded_folders"], list):
        raise ValueError("excluded_folders must be a JSON array of wildcard paths")
    for field in ("data_dir", "key_file"):
        if cfg.get(field):
            field_path = Path(os.path.expandvars(cfg[field])).expanduser()
            cfg[field] = str(field_path if field_path.is_absolute() else config_path.parent / field_path)
    if resolve_dimensions and not cfg.get("dimensions"):
        cfg["dimensions"] = detect_dimensions(cfg)
    if cfg.get("dimensions") and "qwen3-embedding" in cfg["model"].lower():
        cfg.setdefault("storage_dimensions", min(1024, cfg["dimensions"]))
        cfg.setdefault("embedding_request_dimensions", cfg["storage_dimensions"])
    requested = cfg.get("embedding_request_dimensions", 0)
    if requested and cfg.get("dimensions") and not vector_dimensions(cfg) <= requested <= cfg["dimensions"]:
        raise ValueError("embedding_request_dimensions must cover storage_dimensions and not exceed dimensions")
    for field in ("embedding_url", "model", "dimensions", "data_dir"):
        if field == "dimensions" and not resolve_dimensions:
            continue
        if not cfg.get(field):
            raise ValueError(f"Configuration requires {field}.")
    cfg.setdefault("text_cleaning_version", 2)
    if cfg["text_cleaning_version"] not in (1, 2, 3):
        raise ValueError("text_cleaning_version must be 1, 2 or 3")
    for field, default, maximum in (("embedding_batch_size", 8, 128), ("embedding_concurrent_requests", 2, 8), ("sync_batch_emails", 16, 256)):
        cfg.setdefault(field, default)
        if not isinstance(cfg[field], int) or not 1 <= cfg[field] <= maximum:
            raise ValueError(f"{field} must be 1..{maximum}")
    cfg.setdefault("embedding_max_batch_chars", 24000)
    if not 1600 <= cfg["embedding_max_batch_chars"] <= 262144:
        raise ValueError("embedding_max_batch_chars must be 1600..262144")
    if automatic_data_dir:
        digest = hashlib.sha256(index_identity(cfg).encode()).hexdigest()[:12]
        model_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", cfg["model"]).strip(" .")[:80].rstrip(" .") or "model"
        if model_name.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
            model_name = "model_" + model_name
        stored_dimensions = vector_dimensions(cfg) if cfg.get("dimensions") else "auto"
        preferred = index_root / f"{model_name}-{stored_dimensions}d-{digest}"
        # Prefer a readable name for new indexes; never move an active database.
        previous = [preferred, index_root / digest, local_root / "outlook-rag" / digest]
        cfg["data_dir"] = str(next((folder for folder in previous if (folder / "metadata.sqlite").is_file()), preferred))
    return cfg

def detect_dimensions(cfg: dict) -> int:
    key = (cfg["embedding_url"], cfg["model"], cfg.get("model_revision"))
    with _LOCK:
        if key not in _DIMENSIONS:
            with httpx.Client(timeout=120, trust_env=False) as client:
                response = client.post(cfg["embedding_url"], headers={"Authorization": f"Bearer {api_key(cfg)}"},
                    json={"model": cfg["model"], "input": ["Embedding dimension check"]})
                response.raise_for_status()
            rows = response.json().get("data", [])
            if len(rows) != 1 or not rows[0].get("embedding") or not all(math.isfinite(value) for value in rows[0]["embedding"]):
                raise ValueError("Embedding API did not return a valid dimension probe. Set OUTLOOK_RAG_DIMENSIONS explicitly.")
            _DIMENSIONS[key] = len(rows[0]["embedding"])
        return _DIMENSIONS[key]

def query_instruction(cfg: dict) -> str:
    # Most OpenAI-compatible embedding models accept plain query text.
    return cfg.get("query_instruction", QUERY_INSTRUCTION if "qwen3-embedding" in cfg["model"].lower() else "")

def index_identity(cfg: dict) -> str:
    return json.dumps({"model": cfg["model"], "dimensions": cfg.get("dimensions"),
        "instruction": query_instruction(cfg), "chunk_version": cfg.get("text_cleaning_version", 2),
        **({"storage_dimensions": vector_dimensions(cfg)} if cfg.get("dimensions") and vector_dimensions(cfg) != cfg["dimensions"] else {}),
        **({field: cfg[field] for field in ("query_prefix", "document_prefix") if field in cfg}),
        **({field: cfg[field] for field, default in (("chunk_size", 1600), ("chunk_overlap", 200)) if cfg.get(field,default) != default}),
        **({"model_revision": cfg["model_revision"]} if cfg.get("model_revision") else {})},sort_keys=True)

def api_key(cfg: dict) -> str:
    key = os.environ.get(cfg.get("api_key_env", "OUTLOOK_RAG_API_KEY"))
    if key:
        return key
    if not cfg.get("key_file") or not Path(cfg["key_file"]).is_file():
        raise ValueError("Set OUTLOOK_RAG_API_KEY or key_file. The set-key command saves a Windows-encrypted key file.")
    import win32crypt
    return win32crypt.CryptUnprotectData(Path(cfg["key_file"]).read_bytes(), None, None, None, 0)[1].decode("utf-8")

def vector_dimensions(cfg: dict) -> int:
    dimensions = cfg.get("storage_dimensions", cfg["dimensions"])
    if type(dimensions) is not int or not 1 <= dimensions <= cfg["dimensions"]:
        raise ValueError("storage_dimensions must be 1..dimensions. Reduction requires an MRL-compatible model.")
    return dimensions

def embed(texts: list[str], cfg: dict, *, query: bool = False, client: httpx.Client | None = None) -> list[list[float]]:
    if client is None:
        limits = httpx.Limits(max_connections=32, max_keepalive_connections=16, keepalive_expiry=60.0)
        with httpx.Client(timeout=httpx.Timeout(120, connect=10), limits=limits, trust_env=False) as owned:
            return embed(texts, cfg, query=query, client=owned)
    texts = [safe_text(text) for text in texts]
    if query and "query_prefix" in cfg:
        texts = [cfg["query_prefix"] + text for text in texts]
    elif query and query_instruction(cfg):
        texts = [f"Instruct: {query_instruction(cfg)}\nQuery: {text}" for text in texts]
    elif not query and cfg.get("document_prefix"):
        texts = [cfg["document_prefix"] + text for text in texts]
    if not texts:
        return []
    capability_key = (cfg["embedding_url"], cfg["model"], cfg.get("embedding_request_dimensions", 0))
    requested = cfg.get("embedding_request_dimensions", 0) if _API_DIMENSIONS.get(capability_key) is not False else 0
    payload = {"model": cfg["model"], "input": texts}
    if requested:
        payload["dimensions"] = requested
    def send(body: Any) -> Any:
        options: dict[str, Any] = {}
        if cfg.get("_cancel_check", lambda: False)():
            raise SyncCancelled("Sync cancelled")
        if "_embedding_deadline" in cfg:
            remaining = cfg["_embedding_deadline"] - time.perf_counter()
            if remaining <= 0:
                raise httpx.ReadTimeout("Sync embedding time budget exhausted")
            options["timeout"] = httpx.Timeout(remaining, connect=min(10, remaining))
        return client.post(cfg["embedding_url"], headers={"Authorization": f"Bearer {api_key(cfg)}"}, json=body, **options)
    for attempt in range(3):
        response = send(payload)
        if requested and response.status_code in (400, 422):
            # Some proxy/model routes reject dimensions. Retry once without it.
            response = send({"model": cfg["model"], "input": texts})
            if response.is_success:
                _API_DIMENSIONS[capability_key] = False
                requested = 0
                payload.pop("dimensions", None)
        if response.status_code in (429, 502, 503, 504) and attempt < 2:
            time.sleep(attempt + 1)
            continue
        response.raise_for_status()
        break
    rows = sorted(response.json()["data"], key=lambda row: row["index"])
    if len(rows) != len(texts) or [row["index"] for row in rows] != list(range(len(texts))):
        raise ValueError("Embedding API returned an unexpected number or ordering of vectors.")
    target_dim = vector_dimensions(cfg)
    expected_dims = {cfg["dimensions"], requested} if requested else {cfg["dimensions"]}
    dimension_groups: dict[int, list[int]] = {}
    for position, row in enumerate(rows):
        dim = len(row["embedding"])
        if dim not in expected_dims:
            raise ValueError("Embedding dimensions or values do not match the index configuration.")
        dimension_groups.setdefault(dim, []).append(position)
    try:
        import numpy as np
    except ImportError:
        result: list[list[float]] = []
        for row in rows:
            vector = row["embedding"]
            if not all(math.isfinite(v) for v in vector):
                raise ValueError("Embedding dimensions or values do not match the index configuration.")
            vector = vector[:target_dim]
            norm = math.sqrt(sum(v * v for v in vector))
            if norm == 0:
                raise ValueError("Embedding API returned a zero vector.")
            if not math.isfinite(norm):
                raise ValueError("Embedding dimensions or values do not match the index configuration.")
            result.append([v / norm for v in vector])
    else:
        result = [[] for _ in rows]
        # Providers may return either native or requested dimensions per row.
        # Validate every value before truncation, then restore response order.
        for positions in dimension_groups.values():
            matrix = np.asarray([rows[position]["embedding"] for position in positions], dtype=float)
            if not np.isfinite(matrix).all():
                raise ValueError("Embedding dimensions or values do not match the index configuration.")
            matrix = matrix[:, :target_dim]
            with np.errstate(over="ignore"):
                norms = np.linalg.norm(matrix, axis=1, keepdims=True)
            if (norms == 0).any():
                raise ValueError("Embedding API returned a zero vector.")
            if not np.isfinite(norms).all():
                raise ValueError("Embedding dimensions or values do not match the index configuration.")
            matrix /= norms
            for position, vector in zip(positions, matrix.tolist()):
                result[position] = vector
    if requested:
        _API_DIMENSIONS[capability_key] = set(dimension_groups) == {requested}
    return result


def mutation_lock(cfg: dict) -> FileLock:
    root = Path(cfg["data_dir"])
    root.mkdir(parents=True, exist_ok=True)
    return FileLock(root / "mutation.lock", timeout=1)

def reader_lock(cfg: dict) -> FileLock:
    # Maintenance must not retire a version between a reader's SQL snapshot
    # and its Lance checkout. Ordinary publication does not take this lock.
    return FileLock(Path(cfg["data_dir"]) / "reader.lock", timeout=30)

def pin_vectors(sql: sqlite3.Connection, table: Any) -> int:
    row = sql.execute("SELECT value FROM state WHERE key='vector_version'").fetchone()
    if not row:
        raise ValueError("Index publication marker missing. Run rebuild-vectors with the matching settings.")
    version = int(row[0])
    table.checkout(version)
    return version

def publish_vectors(sql: sqlite3.Connection, table: Any) -> None:
    # Call inside the same SQL transaction that publishes mail/chunks.
    sql.execute("INSERT OR REPLACE INTO state VALUES ('vector_version', ?)", (str(table.version),))


@contextmanager
def indexed_snapshot(sql: sqlite3.Connection, table: Any, cfg: dict) -> Iterator[Any]:
    with reader_lock(cfg):
        sql.execute("BEGIN")
        try:
            # A rebuild may have published a new table while embedding was in flight.
            if table.name != storage.table_name(sql):
                table = storage.current_table(sql, cfg)
            pin_vectors(sql, table)
            yield table
        finally:
            sql.rollback()


@contextmanager
def vector_publication(sql: sqlite3.Connection, table: Any) -> Iterator[None]:
    """Serialized writer: recover staged changes, then publish with SQLite."""
    published = pin_vectors(sql, table)
    table.checkout_latest()
    if table.version != published:
        table.restore(published)
    try:
        yield
    except BaseException:
        table.checkout(published)
        raise
    else:
        pin_vectors(sql, table)


def open_store(cfg: dict) -> tuple[sqlite3.Connection, Any]:
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
    # Refuse a model/configuration mismatch rather than mixing embedding spaces.
    identity = index_identity(cfg)
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
        table = vectors.create_table(storage.table_name(sql), schema=storage.schema(vector_dimensions(cfg)))
        if not sql.execute("SELECT 1 FROM chunks LIMIT 1").fetchone():
            with sql:
                sql.executemany("INSERT OR REPLACE INTO state VALUES (?,?)", [("vector_metadata", "1"), ("identifiers_ready", "1")])
    with reader_lock(cfg):
        # Adopt existing indexes without changing their chunk IDs. A concurrent
        # initializer may win; always reread the marker before checking out.
        with sql:
            sql.execute("INSERT OR IGNORE INTO state VALUES ('vector_version', ?)", (str(table.version),))
        pin_vectors(sql, table)
    return sql, table

def upsert_mail(mail: dict, sql: sqlite3.Connection, table: Any, cfg: dict) -> dict:
    return upsert_mails([mail], sql, table, cfg)

def upsert_mails(mails: list[dict], sql: sqlite3.Connection, table: Any, cfg: dict, client: httpx.Client | None = None) -> dict:
    """Batch across mail boundaries and reuse embeddings by exact chunk content.

    SQLite publishes the matching Lance version with the mail/chunk transaction.
    A failed batch is retried idempotently, and sync watermarks do not advance.
    Successful embedding requests are committed separately for reuse on retry.
    """
    stats = dict(embedded_chunks=0, cache_hits=0, unchanged_content=0, embedding_requests=0, sanitized_emails=0,
                 prepare_ms=0, embedding_ms=0, tokenize_ms=0, cache_write_ms=0, vector_write_ms=0, sql_write_ms=0)
    normalized = []
    for mail in mails:
        cleaned = {key: safe_text(value) if isinstance(value,str) else value for key,value in mail.items()}
        stats["sanitized_emails"] += int(cleaned != mail)
        normalized.append(cleaned)
    mails = normalized
    pin_vectors(sql, table)
    stage = time.perf_counter()
    changed, records = [], []
    metadata_changes = []
    unique: dict[str, str] = {}
    for mail in mails:
        if not re.fullmatch(r"[0-9a-f]{64}", mail["id"]):
            raise ValueError("Mail IDs must be SHA256 hex digests.")
        parts = chunks(mail["subject"], mail["body"], cfg.get("chunk_size",1600), cfg.get("chunk_overlap",200), cleaning_version=cfg.get("text_cleaning_version", 2))
        old_chunks = sql.execute("SELECT id,text FROM chunks WHERE email_id=? ORDER BY position", (mail["id"],)).fetchall()
        previous = [row["text"] for row in old_chunks]
        if parts == previous:
            present = {row["id"] for row in table.search().where(f"email_id = '{mail['id']}'")
                       .select(["id"]).limit(len(old_chunks) + 1).to_list()}
            if present == {row["id"] for row in old_chunks}:
                stats["unchanged_content"] += 1
                old = sql.execute("SELECT store_id,folder,sender,received FROM emails WHERE id=?", (mail["id"],)).fetchone()
                if old and storage.metadata(old) != storage.metadata(mail):
                    metadata_changes.append(mail)
                continue
        changed.append(mail["id"])
        for position, text in enumerate(parts):
            digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
            records.append((f"{mail['id']}:{position}", mail["id"], position, text, digest))
            unique.setdefault(digest, text)
    cached: dict[str, tuple[list[float], str]] = {}
    missing: list[tuple[str, str]] = []
    digests = list(unique.keys())
    for i in range(0, len(digests), 500):
        batch_digests = digests[i : i + 500]
        placeholders = ",".join("?" for _ in batch_digests)
        for row in sql.execute(
            f"SELECT hash, vector, terms FROM embedding_cache WHERE hash IN ({placeholders})",
            batch_digests,
        ).fetchall():
            vector = array("f")
            vector.frombytes(row["vector"])
            cached[row["hash"]] = (list(vector), row["terms"])
    for digest, text in unique.items():
        if digest not in cached:
            missing.append((digest, text))
    stats["cache_hits"] = sum(digest in cached for *_, digest in records)
    stats["prepare_ms"] = round((time.perf_counter() - stage) * 1000)
    # Protect completed embeddings from maintenance until their mail is published.
    with sql:
        sql.executemany("DELETE FROM embedding_pending WHERE email_id=?", [(mid,) for mid in changed])
        sql.executemany("INSERT OR IGNORE INTO embedding_pending VALUES (?,?)", [(mid, digest) for _, mid, _, _, digest in records])
    if client is None:
        limits = httpx.Limits(max_connections=32, max_keepalive_connections=16, keepalive_expiry=60.0)
        client = httpx.Client(timeout=httpx.Timeout(120, connect=10), limits=limits, trust_env=False)
        owned = True
    else:
        owned = False
    try:
        batches = []
        pending: list[tuple[str, str]] = []
        chars = 0
        for digest, text in missing:
            if pending and (len(pending) >= cfg.get("embedding_batch_size", 8) or chars + len(text) > cfg.get("embedding_max_batch_chars", 24000)):
                batches.append(pending)
                pending, chars = [], 0
            pending.append((digest, text))
            chars += len(text)
        if pending:
            batches.append(pending)
        stage = time.perf_counter()
        tokenize_seconds = cache_seconds = 0.0
        precomputed_terms: dict[str, str] = {}
        token_errors: dict[str, Exception] = {}
        missing_iter = iter(missing)

        def precompute_one_token() -> bool:
            try:
                d, t = next(missing_iter)
                if d not in precomputed_terms:
                    try:
                        precomputed_terms[d] = tokens(t)
                    except Exception as exc:
                        # Surface this error only when its embedding completes;
                        # earlier successful requests must still reach the cache.
                        token_errors[d] = exc
                return True
            except StopIteration:
                return False

        def persist_batch(batch: Any, vectors: Any) -> None:
            nonlocal tokenize_seconds, cache_seconds
            if len(vectors) != len(batch):
                raise ValueError("Embedding API returned an unexpected number of vectors.")
            mark = time.perf_counter()
            entries = []
            for (digest, text), vector in zip(batch, vectors):
                if digest in token_errors:
                    raise token_errors[digest]
                terms = precomputed_terms.get(digest)
                if terms is None:
                    terms = tokens(text)
                    precomputed_terms[digest] = terms
                cached[digest] = (vector, terms)
                entries.append((digest, array("f", vector).tobytes(), terms))
            tokenize_seconds += time.perf_counter() - mark
            mark = time.perf_counter()
            with sql:
                sql.executemany("INSERT OR IGNORE INTO embedding_cache VALUES (?, ?, ?)", entries)
            cache_seconds += time.perf_counter() - mark
            stats["embedding_requests"] += 1
            stats["embedded_chunks"] += len(batch)
        workers = cfg.get("embedding_concurrent_requests", 2)
        with ThreadPoolExecutor(max_workers=workers) as pool:
            iterator = iter(batches)
            active = {}
            def submit_next() -> None:
                batch = next(iterator, None)
                if batch is not None:
                    active[pool.submit(embed, [text for _, text in batch], cfg, client=client)] = batch
            for _ in range(workers):
                submit_next()
            try:
                while active:
                    idle = cfg.get("_embedding_idle")
                    ready = idle and cfg.get("_embedding_idle_ready", lambda: True)()
                    has_more_tokens = len(precomputed_terms) + len(token_errors) < len(missing)
                    completed, _ = wait(active, timeout=0.01 if (ready or has_more_tokens) else None, return_when=FIRST_COMPLETED)
                    if not completed:
                        if idle and ready:
                            idle()
                        elif has_more_tokens:
                            mark = time.perf_counter()
                            precompute_one_token()
                            tokenize_seconds += time.perf_counter() - mark
                        continue
                    failure = None
                    for future in completed:
                        batch = active.pop(future)
                        try:
                            vectors = future.result()
                            persist_batch(batch, vectors)
                        except Exception as exc:
                            failure = failure or exc
                    if failure is not None:
                        # Do not launch more requests. Retain successes already in flight,
                        # even if the failed future completed before a successful one.
                        for future in active:
                            future.cancel()
                        for future, batch in active.items():
                            if future.cancelled():
                                continue
                            try:
                                vectors = future.result()
                                persist_batch(batch, vectors)
                            except Exception:
                                continue
                        raise failure
                    for _ in completed:
                        submit_next()
            except Exception:
                for future in active:
                    future.cancel()
                raise
        stats["embedding_ms"] = round(max(0, time.perf_counter() - stage - tokenize_seconds - cache_seconds) * 1000)
        stats["tokenize_ms"] = round(tokenize_seconds * 1000)
        stats["cache_write_ms"] = round(cache_seconds * 1000)
    finally:
        if owned:
            client.close()
    with vector_publication(sql, table):
        stage = time.perf_counter()
        native_metadata = storage.metadata_ready(sql, table)
        if native_metadata:
            for mail in metadata_changes:
                table.update(where=f"email_id = '{mail['id']}'", values=storage.metadata(mail))
        if changed:
            table.delete("email_id IN (" + ",".join(f"'{mid}'" for mid in changed) + ")")
            mail_by_id = {mail["id"]: mail for mail in mails}
            table.add([{"id": cid, "email_id": mid, "vector": cached[digest][0],
                        **(storage.metadata(mail_by_id[mid]) if native_metadata else {})} for cid, mid, _, _, digest in records])
        stats["vector_write_ms"] = round((time.perf_counter() - stage) * 1000)
        stage = time.perf_counter()
        with sql:
            affected_threads = threads.previous(sql, [mail["id"] for mail in mails])
            for mail_id in changed:
                old = sql.execute("SELECT id FROM chunks WHERE email_id=?", (mail_id,)).fetchall()
                sql.executemany("DELETE FROM lexical WHERE chunk_id=?", [(row[0],) for row in old])
                sql.executemany("DELETE FROM identifiers WHERE chunk_id=?", [(row[0],) for row in old])
                sql.execute("DELETE FROM chunks WHERE email_id=?", (mail_id,))
            sql.executemany("INSERT INTO chunks VALUES (?, ?, ?, ?)", [row[:4] for row in records])
            sql.executemany("INSERT INTO lexical VALUES (?, ?)", [(cid, cached[digest][1]) for cid, _, _, _, digest in records])
            sql.executemany("INSERT OR IGNORE INTO identifiers VALUES (?,?)", [(identifier, cid) for cid, _, _, text, _ in records for identifier in storage.identifiers(text)])
            columns = ("id", "entry_id", "store_id", "folder", "subject", "sender", "recipients", "received", "modified", "body", "conversation_id")
            sql.executemany("INSERT OR REPLACE INTO emails VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                            [tuple(mail.get(column, "") for column in columns) + (datetime.now(timezone.utc).isoformat(),) for mail in mails])
            sql.executemany("DELETE FROM embedding_pending WHERE email_id=?", [(mail["id"],) for mail in mails])
            affected_threads.update(threads.identity(mail) for mail in mails)
            thread_started = time.perf_counter()
            stats["threads_updated"] = threads.refresh(sql, affected_threads, cfg)
            stats["thread_update_ms"] = round((time.perf_counter()-thread_started)*1000)
            publish_vectors(sql, table)
    stats["sql_write_ms"] = round((time.perf_counter() - stage) * 1000)
    return stats


def sources() -> dict:
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cfg = local_config()
        report, targets = discover_sources(namespace, cfg)
        with stored_sql(cfg) as sql:
            return coverage.live_coverage(report, targets, sql, cfg)
    finally:
        pythoncom.CoUninitialize()


def record_sync_failure(sql: sqlite3.Connection, metadata: dict, store_id: str, folder: str | None, exc: Exception) -> None:
    mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
    row = sql.execute("SELECT attempts FROM sync_failures WHERE mail_id=?", (mail_id,)).fetchone()
    attempts = (row[0] if row else 0) + 1
    failed_at = time.time()
    with sql:
        sql.execute("""INSERT OR REPLACE INTO sync_failures
            (mail_id,entry_id,store_id,folder,modified,attempts,error_type,next_retry,last_failed)
            VALUES (?,?,?,?,?,?,?,?,?)""",
            (mail_id, metadata["entry_id"], store_id, folder, metadata["modified"], attempts,
             type(exc).__name__, failed_at + min(3600, 30 * 2 ** min(attempts - 1, 7)), failed_at))
        sql.execute("INSERT OR REPLACE INTO sync_failure_details VALUES (?,?,?)", (mail_id,failure_reason(exc),failed_at))


def permanent_mail_error(exc: Exception) -> bool:
    return isinstance(exc, ValueError) or (isinstance(exc, httpx.HTTPStatusError)
        and exc.response.status_code in (400, 413, 422))

def sync(folders: list[str] | None = None, max_emails: int | None = None, since_days: int | None = None,
         reconcile: bool = False, max_total_emails: int | None = None,
         max_seconds: float | None = None, max_scanned: int | None = None, *, _cancel_check: Any=None) -> dict:
    cfg = config()
    selection = coverage.folder_selection(folders if folders is not None else cfg.get("folders", "auto"))
    max_emails = max_emails if max_emails is not None else cfg.get("sync_max_emails", 200)
    max_total_emails = max_total_emails if max_total_emails is not None else cfg.get("sync_max_total_emails", 200)
    max_seconds = max_seconds if max_seconds is not None else cfg.get("sync_max_seconds", 30)
    max_scanned = max_scanned if max_scanned is not None else cfg.get("sync_max_scanned", 2000)
    since_days = since_days if since_days is not None else cfg.get("since_days", 0)
    if since_days == "all":
        since_days = 0
    if not 1 <= max_emails <= 2000 or not 1 <= max_total_emails <= 20000:
        raise ValueError("Invalid email limits")
    if not 1 <= max_seconds <= 3600 or not 1 <= max_scanned <= 100000:
        raise ValueError("Invalid time or scan limits")
    if not isinstance(since_days, int) or not 0 <= since_days <= 36500:
        raise ValueError("since_days must be 0 (all time), 'all', or 1..36500")
    started = time.perf_counter()
    deadline = started + max_seconds
    cancelled = _cancel_check or (lambda: False)
    cfg = dict(cfg, _embedding_deadline=deadline, _cancel_check=cancelled)
    import pythoncom
    import pywintypes
    import win32com.client
    reports, pending_folders, discovery_errors, folder_errors, target_folders = [], [], [], [], []
    indexed_total = scanned_total = attempted_total = retried = failed_total = 0
    stop_reason = None
    limits = httpx.Limits(max_connections=32, max_keepalive_connections=16, keepalive_expiry=60.0)
    with mutation_lock(cfg), httpx.Client(timeout=httpx.Timeout(120, connect=10), limits=limits, trust_env=False) as client:
        pythoncom.CoInitialize()
        sql = None
        try:
            namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            sql, table = open_store(cfg)
            alias_session = outlook_session_key(namespace)
            discovery = None
            if selection == "auto":
                discovery, targets = discover_sources(namespace, cfg)
                discovery_errors = discovery["errors"]
            else:
                targets = [(name, resolve_folder(namespace, name)) for name in selection]
            target_folders = [dict(folder=name, folder_id=folder_identity(folder)) for name, folder in targets]
            with sql:
                sql.execute("INSERT OR REPLACE INTO state VALUES ('folder_inventory',?)",
                            (json.dumps(coverage.inventory(discovery, targets, selection, since_days)),))
            rotation_key = "rotation:" + hashlib.sha256(json.dumps([selection, since_days, reconcile], sort_keys=True).encode()).hexdigest()
            rotation = sql.execute("SELECT value FROM state WHERE key=?", (rotation_key,)).fetchone()
            if rotation:
                position = next((i for i, (_, folder) in enumerate(targets)
                    if f"{folder.StoreID}:{folder.EntryID}" == rotation[0]), 0)
                targets = targets[position:] + targets[:position]
            allowed_folders = {name for name, _ in targets}

            # Retry a bounded number of due failures before continuing the scan.
            due = sql.execute("SELECT * FROM sync_failures WHERE next_retry<=? AND folder IN (" +
                ",".join("?" for _ in allowed_folders) + ") ORDER BY next_retry LIMIT ?",
                [time.time(), *sorted(allowed_folders), cfg.get("sync_retry_limit", 10)]).fetchall() if allowed_folders else []
            for failure in due:
                if cancelled():
                    stop_reason = "cancelled"
                    break
                if failure["folder"] not in allowed_folders:
                    continue
                if time.perf_counter() >= deadline or attempted_total >= max_total_emails:
                    break
                metadata = {"entry_id": failure["entry_id"], "modified": failure["modified"], "received": ""}
                attempted_total += 1
                retried += 1
                try:
                    item = namespace.GetItemFromID(failure["entry_id"], failure["store_id"])
                    metadata.update(item=item, modified=iso(item.LastModificationTime), received=iso(item.ReceivedTime))
                    mail = read_outlook_mail(namespace, metadata, failure["store_id"], failure["folder"])
                except Exception as exc:
                    record_sync_failure(sql, metadata, failure["store_id"], failure["folder"], exc)
                    failed_total += 1
                    continue
                try:
                    upsert_mails([mail], sql, table, cfg, client)
                    with sql:
                        sql.execute("DELETE FROM sync_failures WHERE mail_id IN (?,?)", (mail["id"], failure["mail_id"]))
                    indexed_total += 1
                except Exception as exc:
                    if isinstance(exc, httpx.HTTPError) and not permanent_mail_error(exc):
                        stop_reason = "cancelled" if isinstance(exc, SyncCancelled) else "time_limit" if time.perf_counter() >= deadline else "embedding_unavailable"
                        break
                    if not permanent_mail_error(exc):
                        raise
                    record_sync_failure(sql, metadata, failure["store_id"], failure["folder"], exc)
                    failed_total += 1

            for target_position, (folder_name, folder) in enumerate(targets):
                if cancelled():
                    stop_reason = "cancelled"
                if stop_reason or attempted_total >= max_total_emails or scanned_total >= max_scanned or time.perf_counter() >= deadline:
                    stop_reason = stop_reason or ("email_limit" if attempted_total >= max_total_emails else "scan_limit" if scanned_total >= max_scanned else "time_limit")
                    pending_folders = [name for name, _ in targets[target_position:]]
                    break
                store_id = folder.StoreID
                state_key = f"sync:{store_id}:{folder.EntryID}:{since_days}"
                row = sql.execute("SELECT value FROM state WHERE key=?", (state_key,)).fetchone()
                checkpoint = json.loads(row[0]) if row else {}
                old_name = checkpoint.get("folder")
                if old_name and old_name != folder_name:
                    # Stable folder IDs survive path/store display-name changes.
                    # Update rows already scanned before resuming their cursor.
                    storage.update_folder(sql, table, folder_name, store_id=store_id, old_folder=old_name)
                    with sql:
                        sql.execute("UPDATE sync_failures SET folder=? WHERE store_id=? AND folder=?", (folder_name, store_id, old_name))
                continuing = not checkpoint.get("complete", True) and "scan_mode" in checkpoint
                mode = checkpoint.get("scan_mode") if continuing else ("reconcile" if reconcile else "incremental" if checkpoint.get("complete") else "backfill")
                if continuing and reconcile != (mode == "reconcile"):
                    continuing = False
                    mode = "reconcile" if reconcile else "backfill"
                now = datetime.now().astimezone()
                scan_started = checkpoint.get("scan_started", now.isoformat()) if continuing else (
                    checkpoint.get("backfill_started", now.isoformat()) if mode == "backfill" else now.isoformat())
                scan_from = checkpoint.get("scan_from") if continuing else checkpoint.get("watermark") if mode == "incremental" else None
                cursor = checkpoint.get("cursor_received") if continuing or (mode == "backfill" and not reconcile) else None
                cursor_ids = set(checkpoint.get("cursor_ids", [])) if cursor else set()
                cutoff = datetime.now() - timedelta(days=since_days) if since_days else None
                filters = [f"[ReceivedTime] >= '{cutoff.strftime('%m/%d/%Y %I:%M %p')}'"] if cutoff else []
                if mode == "incremental" and scan_from:
                    modified = datetime.fromisoformat(scan_from).astimezone() - timedelta(minutes=5)
                    filters.append(f"[LastModificationTime] >= '{modified.strftime('%m/%d/%Y %I:%M %p')}'")
                if cursor:
                    boundary = datetime.fromisoformat(cursor) + timedelta(minutes=1)
                    filters.append(f"[ReceivedTime] <= '{boundary.strftime('%m/%d/%Y %I:%M %p')}'")
                if mode == "reconcile" and not continuing:
                    with sql:
                        sql.execute("DELETE FROM sync_seen WHERE scan_key=?", (state_key,))
                indexed = skipped = scanned = failed = attempted = overlap = 0
                reconcile_uncertain = checkpoint.get("reconcile_uncertain", False) if continuing else False
                complete = True
                batch_stats: dict[str, Any] = {}
                pending: list[tuple[dict, dict]] = []
                metadata_source = None
                progress_started = time.perf_counter()
                processed_base = checkpoint.get("processed_rows", 0) if continuing else 0
                processing_base = checkpoint.get("processing_seconds", 0.0) if continuing else 0.0
                progress_known = checkpoint.get("progress_known", "processed_rows" in checkpoint) if continuing and cursor else not cursor
                advanced_rows = 0
                source_count = safe_attr(safe_attr(folder, "Items", None), "Count", None)
                source_count = source_count if type(source_count) is int and source_count >= 0 else None

                def save_state(done: Any=False) -> None:
                    value = dict(complete=done, watermark=scan_started if done else checkpoint.get("watermark"),
                        last_run=now.isoformat(), folder=folder_name, since_days=since_days,
                        scan_mode=mode, scan_started=scan_started, scan_from=scan_from,
                        cursor_received=None if done else cursor, cursor_ids=[] if done else sorted(cursor_ids), reconcile_uncertain=reconcile_uncertain,
                        processed_rows=processed_base + advanced_rows, processing_seconds=processing_base + time.perf_counter() - progress_started,
                        progress_known=progress_known, source_items_estimate=source_count if since_days == 0 and mode in ("backfill", "reconcile") else None,
                        progress_updated_at=datetime.now().astimezone().isoformat())
                    with sql:
                        sql.execute("INSERT OR REPLACE INTO state VALUES (?,?)", (state_key, json.dumps(value)))

                def advance(metadata: dict) -> None:
                    nonlocal cursor, cursor_ids, advanced_rows
                    advanced_rows += 1
                    if metadata["received"] != cursor:
                        cursor = metadata["received"]
                        cursor_ids = set()
                    cursor_ids.add(metadata.get("cursor_id", metadata["entry_id"]))
                    if mode == "reconcile":
                        mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
                        sql.execute("INSERT OR IGNORE INTO sync_seen VALUES (?,?)", (state_key, mail_id))

                def flush() -> Any:
                    nonlocal indexed, indexed_total, failed, failed_total
                    if not pending:
                        return
                    try:
                        stats = upsert_mails([mail for mail, _ in pending], sql, table, cfg, client)
                        indexed += len(pending)
                        indexed_total += len(pending)
                        for key, value in stats.items():
                            batch_stats[key] = batch_stats.get(key, 0) + value
                        for mail, metadata in pending:
                            sql.execute("DELETE FROM sync_failures WHERE mail_id=?", (mail["id"],))
                            advance(metadata)
                    except Exception as exc:
                        if not permanent_mail_error(exc):
                            raise
                        # Isolate malformed input without hiding provider outages or storage errors.
                        for mail, metadata in pending:
                            try:
                                stats = upsert_mails([mail], sql, table, cfg, client)
                                indexed += 1
                                indexed_total += 1
                                for key, value in stats.items():
                                    batch_stats[key] = batch_stats.get(key, 0) + value
                                sql.execute("DELETE FROM sync_failures WHERE mail_id=?", (mail["id"],))
                            except Exception as single_exc:
                                if not permanent_mail_error(single_exc):
                                    raise
                                record_sync_failure(sql, metadata, store_id, folder_name, single_exc)
                                failed += 1
                                failed_total += 1
                            advance(metadata)
                    pending.clear()
                    save_state()

                def prepare_ahead(value: Any) -> Any:
                    metadata, source = value
                    if not metadata["mail"]:
                        return value
                    try:
                        metadata = canonical_metadata(namespace, metadata, store_id, sql, alias_session)
                    except Exception as exc:
                        return (dict(metadata, _canonical_error=exc), source)
                    metadata = dict(metadata, _canonical=True)
                    mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
                    previous = sql.execute("SELECT modified FROM emails WHERE id=?", (mail_id,)).fetchone()
                    failure = sql.execute("SELECT modified FROM sync_failures WHERE mail_id=?", (mail_id,)).fetchone()
                    if not ((previous and same_modified(previous["modified"], metadata["modified"])) or
                            (failure and same_modified(failure["modified"], metadata["modified"]))):
                        try:
                            metadata["_prefetched_mail"] = read_outlook_mail(namespace, metadata, store_id, folder_name)
                        except Exception as exc:
                            metadata["_read_error"] = exc
                    return metadata, source

                ahead = ReadAhead(mail_metadata(folder, filters), prepare_ahead,
                    cfg.get("sync_prefetch_emails", 16),
                    lambda queued: not cancelled() and time.perf_counter() < deadline and
                        queued < min(max_emails-attempted, max_total_emails-attempted_total, max_scanned-scanned_total))
                cfg = dict(cfg, _embedding_idle=ahead.fill_one if ahead.limit else None,
                           _embedding_idle_ready=ahead.can_fill)
                try:
                    for metadata, metadata_source in ahead:
                        if cancelled():
                            complete, stop_reason = False, "cancelled"
                            break
                        if time.perf_counter() >= deadline:
                            complete, stop_reason = False, "time_limit"
                            break
                        if cursor and metadata["received"]:
                            current_time = datetime.fromisoformat(metadata["received"]).timestamp()
                            cursor_time = datetime.fromisoformat(cursor).timestamp()
                            if current_time > cursor_time or (current_time == cursor_time and metadata["entry_id"] in cursor_ids):
                                overlap += 1
                                continue
                        if scanned_total >= max_scanned or attempted_total >= max_total_emails or attempted >= max_emails:
                            complete = False
                            if scanned_total >= max_scanned:
                                stop_reason = "scan_limit"
                            elif attempted_total >= max_total_emails:
                                stop_reason = "email_limit"
                            break
                        scanned += 1
                        scanned_total += 1
                        if not metadata["mail"]:
                            flush()
                            advance(metadata)
                            continue
                        try:
                            if "_canonical_error" in metadata:
                                raise metadata["_canonical_error"]
                            if not metadata.get("_canonical"):
                                metadata = canonical_metadata(namespace, metadata, store_id, sql, alias_session)
                        except Exception as exc:
                            flush()
                            attempted += 1
                            attempted_total += 1
                            record_sync_failure(sql, metadata, store_id, folder_name, exc)
                            failed += 1
                            failed_total += 1
                            reconcile_uncertain = True
                            advance(metadata)
                            continue
                        mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
                        previous = sql.execute("SELECT modified FROM emails WHERE id=?", (mail_id,)).fetchone()
                        failure = sql.execute("SELECT modified FROM sync_failures WHERE mail_id=?", (mail_id,)).fetchone()
                        if (previous and same_modified(previous["modified"], metadata["modified"])) or (failure and same_modified(failure["modified"], metadata["modified"])):
                            flush()
                            if previous:
                                storage.update_folder(sql, table, folder_name, mail_id=mail_id)
                            skipped += 1
                            advance(metadata)
                            continue
                        attempted += 1
                        attempted_total += 1
                        try:
                            if "_read_error" in metadata:
                                raise metadata["_read_error"]
                            prepared_mail = metadata.get("_prefetched_mail")
                            if prepared_mail is None:
                                prepared_mail = read_outlook_mail(namespace, metadata, store_id, folder_name)
                        except Exception as exc:
                            flush()
                            record_sync_failure(sql, metadata, store_id, folder_name, exc)
                            failed += 1
                            failed_total += 1
                            advance(metadata)
                            continue
                        pending.append((prepared_mail, metadata))
                        if len(pending) >= cfg.get("sync_batch_emails", 16):
                            flush()
                    flush()
                except httpx.HTTPError as exc:
                    complete, stop_reason = False, "cancelled" if isinstance(exc, SyncCancelled) else "time_limit" if time.perf_counter() >= deadline else "embedding_unavailable"
                    pending.clear()
                except pywintypes.com_error as exc:
                    complete = False
                    flush()
                    folder_errors.append({"folder": folder_name, "error_type": type(exc).__name__})
                batch_stats.update(prefetched_rows=ahead.rows, prefetch_peak_rows=ahead.peak,
                                   prefetch_read_ms=round(ahead.read_ms))
                cfg.pop("_embedding_idle", None)
                cfg.pop("_embedding_idle_ready", None)
                removed = 0
                if complete and mode == "reconcile" and not reconcile_uncertain:
                    candidates = sql.execute("SELECT id,received FROM emails WHERE store_id=? AND folder=? AND id NOT IN (SELECT mail_id FROM sync_seen WHERE scan_key=?)",
                        (store_id, folder_name, state_key)).fetchall()
                    for candidate in candidates:
                        if cutoff and datetime.fromisoformat(candidate["received"]).replace(tzinfo=None) < cutoff:
                            continue
                        remove_mail(candidate["id"], sql, table)
                        removed += 1
                    if cutoff is None:
                        sql.execute("DELETE FROM sync_failures WHERE store_id=? AND folder=? AND mail_id NOT IN (SELECT mail_id FROM sync_seen WHERE scan_key=?)",
                            (store_id, folder_name, state_key))
                    sql.execute("DELETE FROM sync_seen WHERE scan_key=?", (state_key,))
                save_state(complete)
                reports.append(dict(folder=folder_name, folder_id=folder_identity(folder), indexed=indexed, unchanged=skipped, scanned=scanned,
                    cursor_overlap_rows=overlap, failed=failed, window_complete=complete, incremental=mode == "incremental",
                    removed_local=removed, reconcile_skipped=complete and mode == "reconcile" and reconcile_uncertain,
                    metadata_source=metadata_source, batch_stats=batch_stats))
                if targets:
                    next_folder = targets[(target_position + 1) % len(targets)][1]
                    with sql:
                        sql.execute("INSERT OR REPLACE INTO state VALUES (?,?)", (rotation_key, f"{next_folder.StoreID}:{next_folder.EntryID}"))
            if not pending_folders and targets and stop_reason:
                pending_folders = [name for name, _ in targets[len(reports):]]
        finally:
            if sql:
                sql.close()
            pythoncom.CoUninitialize()
    return dict(folders=reports, indexed_total=indexed_total, attempted_total=attempted_total,
        scanned_total=scanned_total, failed_total=failed_total, retried=retried,
        pending_folders=pending_folders, target_folders=target_folders,
        discovery_errors=discovery_errors, folder_errors=folder_errors, stop_reason=stop_reason,
        elapsed_ms=round((time.perf_counter() - started) * 1000),
        note="Repeat incomplete windows and pending folders. Time limits are cooperative; in-flight Outlook/API calls may exceed the budget. Failed emails are retained for retry.")

def remove_mail(mail_id: str, sql: sqlite3.Connection, table: Any) -> None:
    if not re.fullmatch(r"[0-9a-f]{64}", mail_id):
        raise ValueError("Invalid mail_id")
    with vector_publication(sql, table):
        table.delete(f"email_id = '{mail_id}'")
        with sql:
            affected_threads = threads.previous(sql, [mail_id]) if threads.available(sql) else set()
            ids = sql.execute("SELECT id FROM chunks WHERE email_id=?", (mail_id,)).fetchall()
            sql.executemany("DELETE FROM lexical WHERE chunk_id=?", [(row[0],) for row in ids])
            sql.executemany("DELETE FROM identifiers WHERE chunk_id=?", [(row[0],) for row in ids])
            sql.execute("DELETE FROM chunks WHERE email_id=?", (mail_id,))
            sql.execute("DELETE FROM emails WHERE id=?", (mail_id,))
            sql.execute("DELETE FROM embedding_pending WHERE email_id=?", (mail_id,))
            if affected_threads:
                identity = json.loads(sql.execute("SELECT value FROM state WHERE key='embedding_identity'").fetchone()[0])
                threads.refresh(sql, affected_threads, {"text_cleaning_version": identity.get("chunk_version", 2)})
            publish_vectors(sql, table)

def cached_query_vector(query: str, sql: sqlite3.Connection, cfg: dict) -> tuple[list[float], bool]:
    key = hashlib.sha256((index_identity(cfg) + "\n" + safe_text(query)).encode()).hexdigest()
    size = cfg.get("query_cache_size", 256)
    row = sql.execute("SELECT vector FROM query_cache WHERE hash=?", (key,)).fetchone() if size else None
    if row:
        values = array("f")
        values.frombytes(row[0])
        cache_query(sql, "UPDATE query_cache SET last_used=? WHERE hash=?", (time.time(), key))
        return list(values), True
    vector = embed([query], cfg, query=True)[0]
    if size:
        cache_query(sql, "INSERT OR REPLACE INTO query_cache VALUES (?,?,?)", (key, array("f", vector).tobytes(), time.time()), size)
    return vector, False

def cache_query(sql: sqlite3.Connection, statement: str, values: tuple, size: int | None = None) -> None:
    """Cache writes are optional; never wait for the sync writer on a search."""
    timeout = sql.execute("PRAGMA busy_timeout").fetchone()[0]
    sql.execute("PRAGMA busy_timeout=0")
    try:
        with sql:
            sql.execute(statement, values)
            if size is not None:
                sql.execute("DELETE FROM query_cache WHERE hash NOT IN (SELECT hash FROM query_cache ORDER BY last_used DESC LIMIT ?)", (size,))
    except sqlite3.OperationalError as exc:
        if "locked" not in str(exc).lower() and "busy" not in str(exc).lower():
            raise
    finally:
        sql.execute(f"PRAGMA busy_timeout={timeout}")

def search(query: str, limit: int = 10, folder: str | None = None, sender: str | None = None,
           since: str | None = None, until: str | None = None, hybrid: bool = True, group_by_thread: bool = False,
           min_similarity: float | None = None, mode: str | None = None, *,
           subject: str | None = None, recipient: str | None = None, cursor: str | None = None) -> dict:
    return search_engine.search(sys.modules[__name__], query, limit, folder, sender, since, until,
                                hybrid, group_by_thread, min_similarity, mode,
                                subject=subject, recipient=recipient, cursor=cursor)

def status() -> dict:
    cfg = local_config()
    if not cfg.get('dimensions'):
        raise ValueError('Index not initialized. Run sync_emails with the embedding API available first.')
    with _LOCK:
        sql, table = storage.open_index(cfg, writable=False)
        try:
            with indexed_snapshot(sql, table, cfg) as table:
                return {"folder_progress": folder_progress(sql, dict(cfg, folders="auto")), "search_coverage": coverage.index_coverage(sql, cfg), "model": cfg["model"], "dimensions": cfg["dimensions"],
                    "storage_dimensions": vector_dimensions(cfg),
                    "text_cleaning_version": cfg["text_cleaning_version"],
                    "emails": sql.execute("SELECT count(*) FROM emails").fetchone()[0],
                    "chunks": sql.execute("SELECT count(*) FROM chunks").fetchone()[0],
                    "cached_embeddings": sql.execute("SELECT count(*) FROM embedding_cache").fetchone()[0],
                    "cached_queries": sql.execute("SELECT count(*) FROM query_cache").fetchone()[0],
                    "failed_emails": sql.execute("SELECT count(*) FROM sync_failures").fetchone()[0],
                    "sync_max_seconds": cfg.get("sync_max_seconds", 30), "sync_max_scanned": cfg.get("sync_max_scanned", 2000),
                    "embedding_request_dimensions": cfg.get("embedding_request_dimensions", 0),
                    "api_dimension_reduction": _API_DIMENSIONS.get((cfg["embedding_url"], cfg["model"], cfg.get("embedding_request_dimensions", 0)), "unverified"),
                    "embedding_batch_size": cfg["embedding_batch_size"], "embedding_concurrent_requests": cfg["embedding_concurrent_requests"], "sync_batch_emails": cfg["sync_batch_emails"], "sync_prefetch_emails": cfg.get("sync_prefetch_emails", 16),
                    "vector_rows": table.count_rows(), "vector_indexes": [str(index) for index in table.list_indices()],
                    "vector_coverage": storage.vector_status(table), "vector_metadata_filters": storage.metadata_ready(sql, table),
                    "thread_index": threads.status(sql),
                    "folders": [dict(row) for row in sql.execute("SELECT folder, count(*) AS emails, max(indexed_at) AS last_indexed FROM emails GROUP BY folder")],
                    "sync_windows": [window for row in sql.execute("SELECT value FROM state WHERE key LIKE 'sync:%'")
                                     if (window := json.loads(row[0])) and (cfg.get("folders") == "auto" or window.get("folder") in cfg["folders"])],
                    "data_dir": cfg["data_dir"]}
        finally:
            sql.close()


def get_mail(mail_id: str, max_body_chars: int = 10000) -> dict:
    if not 0 <= max_body_chars <= 100000:
        raise ValueError("max_body_chars must be 0..100000")
    with _LOCK, stored_sql(local_config()) as sql:
        if sql is None:
            raise ValueError('Indexed mail not found.')
        row = sql.execute("SELECT * FROM emails WHERE id=?", (mail_id,)).fetchone()
        if not row:
            raise ValueError("Indexed mail not found.")
        result = dict(row)
        result["body_truncated"] = len(result["body"]) > max_body_chars
        result["body"] = result["body"][:max_body_chars]
        result["source"] = "local_index"
        return result

def optimize() -> dict:
    cfg = local_config()
    with mutation_lock(cfg):
        sql, table = storage.open_index(cfg)
        try:
            count = table.count_rows()
            if count < 256:
                return {"indexed": False, "rows": count, "note": "Small index uses exact vector search; ANN enabled from 256 chunks."}
            with vector_publication(sql, table):
                table.create_index("vector", config=IvfHnswFlat(distance_type="cosine", num_partitions=max(1, min(64, count // 256))))
                with sql:
                    publish_vectors(sql, table)
            return {"indexed": True, "rows": count, "indexes": [str(index) for index in table.list_indices()]}
        finally:
            sql.close()

def maintain() -> dict:
    """Reclaim unreferenced caches and compact Lance versions, retaining seven days."""
    cfg = local_config()
    with mutation_lock(cfg):
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
            with reader_lock(cfg):
                # Protect the rollback version even if it was published more
                # than seven days ago. Keep the tag through failed maintenance.
                published = pin_vectors(sql, table)
                tag = "outlook-rag-published"
                if tag in table.tags.list():
                    table.tags.update(tag, published)
                else:
                    table.tags.create(tag, published)
                with vector_publication(sql, table):
                    result = table.optimize(cleanup_older_than=timedelta(days=7), delete_unverified=False)
                    with sql:
                        publish_vectors(sql, table)
                vectors = lancedb.connect(str(Path(cfg["data_dir"]) / "vectors"))
                current = storage.table_name(sql)
                for name in vectors.list_tables().tables:
                    if name != current and re.fullmatch(r"mail_chunks(?:_[0-9a-f]{32})?", name):
                        vectors.drop_table(name)
            removed_logs = jobs.prune_logs(cfg)
            return {"removed_worker_logs": removed_logs, "removed_cached_embeddings": removed, "removed_entry_aliases": removed_aliases, "vector_maintenance": str(result), "retention_days": 7}
        finally:
            sql.close()


mcp = FastMCP("outlook-rag", instructions=server.MAIL_SEARCH_INSTRUCTIONS)
logging.getLogger("mcp.server.lowlevel.server").setLevel(logging.WARNING)

def warmup_embedding_api() -> None:
    try:
        embed(["email retrieval warmup"], config(), query=True)
    except Exception as exc:
        logging.getLogger(__name__).warning("Embedding warmup failed (%s); search will retry on demand.", type(exc).__name__)

_tools = server.register_tools(mcp, sys.modules[__name__])
search_emails = _tools['search_emails']
index_status = _tools['index_status']
list_outlook_sources = _tools['list_outlook_sources']
get_indexed_mail = _tools['get_indexed_mail']
sync_emails = _tools['sync_emails']
optimize_index = _tools['optimize_index']
maintain_index = _tools['maintain_index']
start_sync_job = _tools['start_sync_job']
sync_job_status = _tools['sync_job_status']
cancel_sync_job = _tools['cancel_sync_job']
list_sync_failures = _tools['list_sync_failures']
retry_failed_emails = _tools['retry_failed_emails']
get_sync_progress = _tools['get_sync_progress']
get_mail_thread = _tools['get_mail_thread']

def main() -> None:
    cli.main(sys.modules[__name__])


if __name__ == "__main__":
    main()
