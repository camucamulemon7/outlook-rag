"""Exact attribute lookup and keyset pages without embeddings or LanceDB."""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import math
from typing import Any

from . import coverage, features, mail_metadata

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
    has_attachments: bool | None = None,
    deduplicate: bool = True,
) -> dict:
    from .search import sql_filter

    started = backend.time.perf_counter()
    signature = hashlib.sha256(
        json.dumps(
            [query, folder, sender, since, until, subject, recipient, has_attachments, deduplicate],
            ensure_ascii=False,
        ).encode()
    ).hexdigest()
    page = None
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
    with features.stored_sql(cfg) as sql:
        extra = sql is not None and mail_metadata.available(sql)
        condition, params = sql_filter(folder, sender, since, until, extra=extra, recipient=recipient, has_attachments=has_attachments)
        for value in (subject, query):
            if value:
                condition += " AND instr(lower(e.subject),lower(?))>0"
                params.append(value)
        if deduplicate and extra:
            eligible, eligible_params = condition.replace("e.", "d."), list(params)
            condition += f""" AND NOT EXISTS(SELECT 1 FROM mail_metadata m1
                JOIN mail_metadata m2 ON m2.internet_message_id=m1.internet_message_id
                JOIN emails d ON d.id=m2.email_id WHERE m1.email_id=e.id
                AND m1.internet_message_id<>'' AND {eligible}
                AND (coalesce(julianday(d.received),0)>{INSTANT}
                  OR (coalesce(julianday(d.received),0)={INSTANT} AND d.id>e.id)))"""
            params.extend(eligible_params)
        if page is not None:
            condition += f" AND ({INSTANT}<? OR ({INSTANT}=? AND e.id<?))"
            params.extend([page["instant"], page["instant"], page["id"]])
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
        if sql is not None:
            extras = mail_metadata.lookup(sql, [row["mail_id"] for row in items])
            for row in items:
                row.update(extras.get(row["mail_id"], {}))
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
                deduplicate=deduplicate,
                freshness="As of last sync; use sync_index(action='start') to refresh.",
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
