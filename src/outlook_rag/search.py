"""Filtered vector retrieval, RRF ranking and query-centered passages."""

from __future__ import annotations

from datetime import datetime
import heapq
import math
import re
from typing import Any

FILTER_BATCH = 1000


def id_filter(column: str, ids: list[str]) -> str:
    # Only generated SHA256 IDs may be interpolated into Lance expressions.
    pattern = r"[0-9a-f]{64}:[0-9]+" if column == "id" else r"[0-9a-f]{64}"
    if column not in {"id", "email_id"} or not all(
        re.fullmatch(pattern, value) for value in ids
    ):
        raise ValueError("Invalid index identifier")
    return column + " IN (" + ",".join(f"'{value}'" for value in ids) + ")"


def vector_candidates(
    table: Any, vector: list[float], count: int, allowed: list[str] | None
) -> list[dict]:
    """Merge shard top-Ks to preserve filtered recall with bounded expressions."""
    groups = (
        [None]
        if allowed is None
        else [
            allowed[i : i + FILTER_BATCH] for i in range(0, len(allowed), FILTER_BATCH)
        ]
    )
    best: list[dict] = []
    for group in groups:
        request = (
            table.search(vector)
            .distance_type("cosine")
            .limit(count)
            .select(["id", "_distance"])
        )
        if group is not None:
            request = request.where(id_filter("email_id", group), prefilter=True)
        rows = request.to_list()
        # Retain at most K between shards; vector payloads are never materialized.
        best = heapq.nsmallest(
            count, best + rows, key=lambda row: (float(row["_distance"]), row["id"])
        )
    return best


def contextual_snippet(
    text: str, query: str, terms: list[str], size: int = 1800
) -> str:
    """Plain text window; never insert HTML into untrusted cached mail."""
    if len(text) <= size:
        return text
    phrases = [query.strip(), *terms]
    # Unicode lowercasing can expand a character (e.g. İ -> i + dot).
    # Match the original string so offsets remain valid for its slices.
    positions = [
        match.start()
        for word in phrases
        if word.strip() and (match := re.search(re.escape(word), text, re.IGNORECASE))
    ]
    if not positions:
        return text[:size]
    position = min(positions)
    start = max(0, position - size // 3)
    start = min(start, len(text) - size)
    return (
        ("…" if start else "")
        + text[start : start + size]
        + ("…" if start + size < len(text) else "")
    )


def search(
    backend: Any,
    query: str,
    limit: int = 10,
    folder: str | None = None,
    sender: str | None = None,
    since: str | None = None,
    until: str | None = None,
    hybrid: bool = True,
    group_by_thread: bool = False,
    min_similarity: float | None = None,
) -> dict:
    if not query.strip() or len(query) > 4000 or not 1 <= limit <= 50:
        raise ValueError("Provide a nonempty query (max 4000 characters), limit 1..50.")
    if min_similarity is not None and (
        isinstance(min_similarity, bool)
        or not math.isfinite(min_similarity)
        or not -1 <= min_similarity <= 1
    ):
        raise ValueError("min_similarity must be finite and in -1..1")
    time = backend.time
    started = time.perf_counter()
    cfg = backend.config()
    with backend._LOCK:
        sql, table = backend.open_store(cfg)
        try:
            if not sql.execute("SELECT 1 FROM emails LIMIT 1").fetchone():
                return {
                    "items": [],
                    "count": 0,
                    "note": "Index empty: run sync_emails first.",
                }
            clauses, params = [], []
            if folder:
                clauses.append("e.folder=?")
                params.append(folder)
            if sender:
                clauses.append("instr(lower(e.sender), lower(?)) > 0")
                params.append(sender)
            if since:
                bound = datetime.fromisoformat(since).date().isoformat()
                clauses.append("substr(e.received,1,10)>=?")
                params.append(bound)
            if until:
                bound = datetime.fromisoformat(until).date().isoformat()
                clauses.append("substr(e.received,1,10)<=?")
                params.append(bound)
            condition = " AND ".join(clauses) or "1=1"
            if (
                clauses
                and not sql.execute(
                    f"SELECT 1 FROM emails e WHERE {condition} LIMIT 1", params
                ).fetchone()
            ):
                # An empty SQL result is already a consistent answer at this
                # point. Nonempty filters are rechecked after embedding below.
                return {
                    "items": [],
                    "count": 0,
                    "elapsed_ms": round((time.perf_counter() - started) * 1000),
                }
            semantic_weight = cfg.get("semantic_weight", 1.0) if hybrid else 1.0
            lexical_weight = cfg.get("lexical_weight", 1.0) if hybrid else 0.0
            t = time.perf_counter()
            vector, query_cache_hit = [], False
            if semantic_weight or min_similarity is not None:
                vector, query_cache_hit = backend.cached_query_vector(query, sql, cfg)
            embedding_ms = round((time.perf_counter() - t) * 1000)
            with backend.indexed_snapshot(sql, table, cfg):
                allowed = None
                if clauses:
                    allowed = [
                        row[0]
                        for row in sql.execute(
                            f"SELECT e.id FROM emails e WHERE {condition}", params
                        )
                    ]
                    if not allowed:
                        return {
                            "items": [],
                            "count": 0,
                            "elapsed_ms": round((time.perf_counter() - started) * 1000),
                        }
                terms = (
                    list(dict.fromkeys(backend.tokens(query).split()))[:24]
                    if lexical_weight
                    else []
                )
                expression = " OR ".join(
                    '"' + term.replace('"', '""') + '"' for term in terms
                )
                candidate_count = max(cfg.get("search_candidates", 100), limit * 10)
                candidate_cap = 5000
                search_deadline = time.perf_counter() + 5
                mail_rows = {}
                while True:
                    vector_rows = (
                        vector_candidates(table, vector, candidate_count, allowed)
                        if semantic_weight
                        else []
                    )
                    scores, similarities = {}, {}
                    for rank, row in enumerate(vector_rows, 1):
                        scores[row["id"]] = semantic_weight / (
                            cfg.get("rrf_k", 60) + rank
                        )
                        similarities[row["id"]] = 1 - float(row["_distance"])
                    lexical_rows = []
                    if expression:
                        lexical_rows = sql.execute(
                            f"""SELECT l.chunk_id FROM lexical l JOIN chunks c ON c.id=l.chunk_id
                            JOIN emails e ON e.id=c.email_id WHERE lexical MATCH ? AND {condition}
                            ORDER BY bm25(lexical),l.chunk_id LIMIT ?""",
                            [expression] + params + [candidate_count],
                        ).fetchall()
                        for rank, row in enumerate(lexical_rows, 1):
                            scores[row[0]] = scores.get(row[0], 0) + lexical_weight / (
                                cfg.get("rrf_k", 60) + rank
                            )
                    if min_similarity is not None:
                        missing = [cid for cid in scores if cid not in similarities]
                        for offset in range(0, len(missing), FILTER_BATCH):
                            ids = missing[offset : offset + FILTER_BATCH]
                            condition_ids = id_filter("id", ids)
                            for hit in (
                                table.search()
                                .where(condition_ids)
                                .select(["id", "vector"])
                                .limit(len(ids))
                                .to_list()
                            ):
                                values = hit["vector"]
                                norm = math.sqrt(sum(float(v) ** 2 for v in values))
                                query_norm = math.sqrt(
                                    sum(float(v) ** 2 for v in vector)
                                )
                                if norm and query_norm:
                                    similarities[hit["id"]] = sum(
                                        float(a) * float(b)
                                        for a, b in zip(values, vector)
                                    ) / (norm * query_norm)
                    output, seen, threads = [], set(), set()
                    for cid, score in sorted(
                        scores.items(), key=lambda pair: pair[1], reverse=True
                    ):
                        if (
                            min_similarity is not None
                            and similarities.get(cid, -2) < min_similarity
                        ):
                            continue
                        if cid not in mail_rows:
                            mail_rows[cid] = sql.execute(
                                """SELECT e.*, c.text AS snippet FROM chunks c
                                JOIN emails e ON e.id=c.email_id WHERE c.id=?""",
                                (cid,),
                            ).fetchone()
                        row = mail_rows[cid]
                        if not row or row["id"] in seen:
                            continue
                        seen.add(row["id"])
                        thread = (row["store_id"], row["conversation_id"] or row["id"])
                        if group_by_thread and thread in threads:
                            continue
                        threads.add(thread)
                        output.append(
                            {
                                "mail_id": row["id"],
                                "entry_id": row["entry_id"],
                                "store_id": row["store_id"],
                                "subject": row["subject"],
                                "sender": row["sender"],
                                "received": row["received"],
                                "folder": row["folder"],
                                "conversation_id": row["conversation_id"],
                                "snippet": contextual_snippet(
                                    row["snippet"], query, terms
                                ),
                                "rank_score": round(score, 6),
                                "semantic_similarity": round(similarities[cid], 4)
                                if cid in similarities
                                else None,
                            }
                        )
                        if group_by_thread:
                            output[-1]["thread_mail_count"] = (
                                sql.execute(
                                    "SELECT count(*) FROM emails WHERE store_id=? AND conversation_id=?",
                                    (row["store_id"], row["conversation_id"]),
                                ).fetchone()[0]
                                if row["conversation_id"]
                                else 1
                            )
                        if len(output) >= limit:
                            break
                    exhausted = (
                        len(vector_rows) < candidate_count
                        and len(lexical_rows) < candidate_count
                    )
                    time_limited = time.perf_counter() >= search_deadline
                    if (
                        len(output) >= limit
                        or exhausted
                        or candidate_count >= candidate_cap
                        or time_limited
                    ):
                        break
                    # More chunks from the same mail/thread must not hide the
                    # next distinct result. Recompute RRF using the larger pool.
                    candidate_count = min(candidate_cap, candidate_count * 2)
                return {
                    "query": query,
                    "count": len(output),
                    "items": output,
                    "embedding_ms": embedding_ms,
                    "elapsed_ms": round((time.perf_counter() - started) * 1000),
                    "hybrid": hybrid,
                    "query_cache_hit": query_cache_hit,
                    "min_similarity": min_similarity,
                    "candidate_limit": candidate_count,
                    "candidate_cap": candidate_cap,
                    "candidate_limit_reached": len(output) < limit
                    and not exhausted
                    and candidate_count >= candidate_cap,
                    "candidate_time_limit_reached": len(output) < limit
                    and not exhausted
                    and time_limited,
                    "grouped_by_thread": group_by_thread,
                    "source": "local_index",
                    "freshness": "As of last sync; run sync_emails to refresh.",
                }
        finally:
            sql.close()
