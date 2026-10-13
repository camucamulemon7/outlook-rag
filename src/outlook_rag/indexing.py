"""Idempotent mail/chunk publication, embedding reuse and failure records.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

from array import array
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
import hashlib
import json
import re
import sqlite3
import time
from datetime import datetime, timezone
from typing import Any

import httpx

from . import storage, threads

def upsert_mail(backend: Any, mail: dict, sql: sqlite3.Connection, table: Any, cfg: dict) -> dict:
    return backend.upsert_mails([mail], sql, table, cfg)


def upsert_mails(backend: Any, mails: list[dict], sql: sqlite3.Connection, table: Any, cfg: dict, client: httpx.Client | None = None) -> dict:
    """Batch across mail boundaries and reuse embeddings by exact chunk content.

    SQLite publishes the matching Lance version with the mail/chunk transaction.
    A failed batch is retried idempotently, and sync watermarks do not advance.
    Successful embedding requests are committed separately for reuse on retry.
    """
    stats = dict(embedded_chunks=0, cache_hits=0, unchanged_content=0, embedding_requests=0, sanitized_emails=0,
                 prepare_ms=0, embedding_ms=0, tokenize_ms=0, cache_write_ms=0, vector_write_ms=0, sql_write_ms=0)
    normalized = []
    for mail in mails:
        cleaned = {key: backend.safe_text(value) if isinstance(value,str) else value for key,value in mail.items()}
        stats["sanitized_emails"] += int(cleaned != mail)
        normalized.append(cleaned)
    mails = normalized
    backend.pin_vectors(sql, table)
    stage = time.perf_counter()
    changed, records = [], []
    metadata_changes = []
    unique: dict[str, str] = {}
    for mail in mails:
        if not re.fullmatch(r"[0-9a-f]{64}", mail["id"]):
            raise ValueError("Mail IDs must be SHA256 hex digests.")
        parts = backend.chunks(mail["subject"], mail["body"], cfg.get("chunk_size",1600), cfg.get("chunk_overlap",200), cleaning_version=cfg.get("text_cleaning_version", 2))
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
                        precomputed_terms[d] = backend.tokens(t)
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
                    terms = backend.tokens(text)
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
                    active[pool.submit(backend.embed, [text for _, text in batch], cfg, client=client)] = batch
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
    with backend.vector_publication(sql, table):
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
            backend.publish_vectors(sql, table)
    stats["sql_write_ms"] = round((time.perf_counter() - stage) * 1000)
    return stats


def record_sync_failure(backend: Any, sql: sqlite3.Connection, metadata: dict, store_id: str, folder: str | None, exc: Exception) -> None:
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
        sql.execute("INSERT OR REPLACE INTO sync_failure_details VALUES (?,?,?)", (mail_id,backend.failure_reason(exc),failed_at))


def permanent_mail_error(backend: Any, exc: Exception) -> bool:
    return isinstance(exc, ValueError) or (isinstance(exc, httpx.HTTPStatusError)
        and exc.response.status_code in (400, 413, 422))


def remove_mail(backend: Any, mail_id: str, sql: sqlite3.Connection, table: Any) -> None:
    if not re.fullmatch(r"[0-9a-f]{64}", mail_id):
        raise ValueError("Invalid mail_id")
    with backend.vector_publication(sql, table):
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
            backend.publish_vectors(sql, table)

