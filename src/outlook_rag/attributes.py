"""Exact attribute lookup and keyset pages without embeddings or LanceDB."""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import math
from typing import Any

from . import coverage, features

INSTANT = "coalesce(julianday(e.received),0)"


def search(
    backend: Any,
    cfg: dict,
    query: str,
    limit: int,
    folder: str | None,
    sender: str | None,
    since: str | None,
    until: str | None,
    subject: str | None,
    recipient: str | None,
    cursor: str | None,
) -> dict:
    from .search import sql_filter

    started = backend.time.perf_counter()
    condition, params = sql_filter(folder, sender, since, until)
    for column, value in (
        ("subject", subject),
        ("recipients", recipient),
        ("subject", query),
    ):
        if value:
            condition += f" AND instr(lower(e.{column}),lower(?))>0"
            params.append(value)
    signature = hashlib.sha256(
        json.dumps(
            [query, folder, sender, since, until, subject, recipient],
            ensure_ascii=False,
        ).encode()
    ).hexdigest()
    if cursor is not None:
        try:
            if not isinstance(cursor, str) or len(cursor) > 1000:
                raise ValueError("Invalid cursor")
            page = json.loads(base64.urlsafe_b64decode(cursor.encode()))
            if (
                page["scope"] != signature
                or not isinstance(page["id"], str)
                or len(page["id"]) != 64
                or any(c not in "0123456789abcdef" for c in page["id"])
                or type(page["instant"]) not in (int, float)
                or not math.isfinite(page["instant"])
                or not 0 <= page["instant"] <= 5373485
            ):
                raise ValueError("Invalid cursor")
        except (ValueError, TypeError, KeyError, binascii.Error, UnicodeError) as exc:
            raise ValueError(
                "Invalid metadata cursor or changed search filters"
            ) from exc
        condition += f" AND ({INSTANT}<? OR ({INSTANT}=? AND e.id<?))"
        params.extend([page["instant"], page["instant"], page["id"]])
    with backend._LOCK, features.stored_sql(cfg) as sql:
        rows = (
            []
            if sql is None
            else sql.execute(
                f"""SELECT e.id,e.entry_id,e.store_id,e.subject,e.sender,e.recipients,
            e.received,e.folder,e.conversation_id,substr(e.body,1,1800) snippet,
            {INSTANT} instant FROM emails e WHERE {condition}
            ORDER BY {INSTANT} DESC,e.id DESC LIMIT ?""",
                [*params, limit + 1],
            ).fetchall()
        )
        items = [
            dict(
                mail_id=row["id"],
                **{
                    name: row[name]
                    for name in (
                        "entry_id",
                        "store_id",
                        "subject",
                        "sender",
                        "recipients",
                        "received",
                        "folder",
                        "conversation_id",
                        "snippet",
                    )
                },
            )
            for row in rows[:limit]
        ]
        next_cursor = None
        if len(rows) > limit:
            last = rows[limit - 1]
            next_cursor = base64.urlsafe_b64encode(
                json.dumps(
                    dict(
                        scope=signature,
                        id=last["id"],
                        instant=last["instant"],
                    )
                ).encode()
            ).decode()
        result = coverage.search_result(
            dict(
                query=query,
                items=items,
                count=len(items),
                next_cursor=next_cursor,
                search_mode="metadata",
                hybrid=False,
                retrieval="sql_attributes",
                embedding_ms=0,
                query_cache_hit=False,
                source="local_index",
                elapsed_ms=round((backend.time.perf_counter() - started) * 1000),
                order="received_instant_desc_then_mail_id",
                metadata_prefilter=True,
                freshness="As of last sync; run sync_emails to refresh.",
            ),
            sql,
            cfg,
            folder,
            since,
            until,
        )
        result["search_scope"].update(
            sender=sender,
            subject=subject,
            recipient=recipient,
            literal_subject=query or None,
        )
        return result
