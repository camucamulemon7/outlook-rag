"""Embedding HTTP requests, validation and durable query caching.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

from array import array
import hashlib
import math
import sqlite3
import time
from typing import Any

import httpx


class SyncCancelled(httpx.RequestError):
    """Cooperative cancellation; completed embedding requests remain cached."""


def detect_dimensions(backend: Any, cfg: dict) -> int:
    key = (cfg["embedding_url"], cfg["model"], cfg.get("model_revision"))
    with backend._LOCK:
        if key not in backend._DIMENSIONS:
            with httpx.Client(timeout=120, trust_env=False) as client:
                response = client.post(cfg["embedding_url"], headers={"Authorization": f"Bearer {backend.api_key(cfg)}"},
                    json={"model": cfg["model"], "input": ["Embedding dimension check"]})
                response.raise_for_status()
            rows = response.json().get("data", [])
            if len(rows) != 1 or not rows[0].get("embedding") or not all(math.isfinite(value) for value in rows[0]["embedding"]):
                raise ValueError("Embedding API did not return a valid dimension probe. Set OUTLOOK_RAG_DIMENSIONS explicitly.")
            backend._DIMENSIONS[key] = len(rows[0]["embedding"])
        return backend._DIMENSIONS[key]


def embed(backend: Any, texts: list[str], cfg: dict, *, query: bool = False, client: httpx.Client | None = None) -> list[list[float]]:
    if client is None:
        limits = httpx.Limits(max_connections=32, max_keepalive_connections=16, keepalive_expiry=60.0)
        with httpx.Client(timeout=httpx.Timeout(120, connect=10), limits=limits, trust_env=False) as owned:
            return backend.embed(texts, cfg, query=query, client=owned)
    texts = [backend.safe_text(text) for text in texts]
    if query and "query_prefix" in cfg:
        texts = [cfg["query_prefix"] + text for text in texts]
    elif query and backend.query_instruction(cfg):
        texts = [f"Instruct: {backend.query_instruction(cfg)}\nQuery: {text}" for text in texts]
    elif not query and cfg.get("document_prefix"):
        texts = [cfg["document_prefix"] + text for text in texts]
    if not texts:
        return []
    capability_key = (cfg["embedding_url"], cfg["model"], cfg.get("embedding_request_dimensions", 0))
    requested = cfg.get("embedding_request_dimensions", 0) if backend._API_DIMENSIONS.get(capability_key) is not False else 0
    payload = {"model": cfg["model"], "input": texts}
    if requested:
        payload["dimensions"] = requested
    def send(body: Any) -> Any:
        options: dict[str, Any] = {}
        if cfg.get("_cancel_check", lambda: False)():
            raise backend.SyncCancelled("Sync cancelled")
        if "_embedding_deadline" in cfg:
            remaining = cfg["_embedding_deadline"] - time.perf_counter()
            if remaining <= 0:
                raise httpx.ReadTimeout("Sync embedding time budget exhausted")
            options["timeout"] = httpx.Timeout(remaining, connect=min(10, remaining))
        return client.post(cfg["embedding_url"], headers={"Authorization": f"Bearer {backend.api_key(cfg)}"}, json=body, **options)
    for attempt in range(3):
        response = send(payload)
        if requested and response.status_code in (400, 422):
            # Some proxy/model routes reject dimensions. Retry once without it.
            response = send({"model": cfg["model"], "input": texts})
            if response.is_success:
                backend._API_DIMENSIONS[capability_key] = False
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
    target_dim = backend.vector_dimensions(cfg)
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
        backend._API_DIMENSIONS[capability_key] = set(dimension_groups) == {requested}
    return result


def cached_query_vector(backend: Any, query: str, sql: sqlite3.Connection, cfg: dict) -> tuple[list[float], bool]:
    key = hashlib.sha256((backend.index_identity(cfg) + "\n" + backend.safe_text(query)).encode()).hexdigest()
    size = cfg.get("query_cache_size", 256)
    row = sql.execute("SELECT vector FROM query_cache WHERE hash=?", (key,)).fetchone() if size else None
    if row:
        values = array("f")
        values.frombytes(row[0])
        backend.cache_query(sql, "UPDATE query_cache SET last_used=? WHERE hash=?", (time.time(), key))
        return list(values), True
    vector = backend.embed([query], cfg, query=True)[0]
    if size:
        backend.cache_query(sql, "INSERT OR REPLACE INTO query_cache VALUES (?,?,?)", (key, array("f", vector).tobytes(), time.time()), size)
    return vector, False


def cache_query(backend: Any, sql: sqlite3.Connection, statement: str, values: tuple, size: int | None = None) -> None:
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

