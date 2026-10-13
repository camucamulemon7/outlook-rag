"""Bounded extractive thread documents and original-mail evidence; no model API."""

from __future__ import annotations

from contextlib import closing
import hashlib
import json
from pathlib import Path
import sqlite3
import time
from typing import Any

from . import text

INSTANT = "coalesce(julianday(received),0)"
SUMMARY_ROWS = 12
SUMMARY_CHARS = 6000
CONTEXT_CHARS = 12000


def initialize(sql: sqlite3.Connection) -> None:
    sql.executescript(f"""
    CREATE TABLE IF NOT EXISTS thread_documents(
      key TEXT PRIMARY KEY,store_id TEXT,conversation_id TEXT,anchor_mail_id TEXT,
      member_count INTEGER,search_text TEXT,sources TEXT,fingerprint TEXT);
    CREATE VIRTUAL TABLE IF NOT EXISTS thread_lexical USING fts5(thread_key UNINDEXED,terms);
    CREATE INDEX IF NOT EXISTS emails_received_instant ON emails({INSTANT},id);
    CREATE INDEX IF NOT EXISTS emails_conversation_instant ON emails(store_id,conversation_id,{INSTANT},id);
    """)
    with sql:
        complete = (
            "0" if sql.execute("SELECT 1 FROM emails LIMIT 1").fetchone() else "1"
        )
        sql.execute(
            "INSERT OR IGNORE INTO state VALUES ('thread_index_complete',?)", [complete]
        )


def available(sql: sqlite3.Connection) -> bool:
    return (
        sql.execute(
            "SELECT count(*) FROM sqlite_master WHERE name IN ('thread_documents','thread_lexical')"
        ).fetchone()[0]
        == 2
    )


def identity(mail: Any) -> tuple[str, str, str]:
    mail = dict(mail)
    conversation = mail.get("conversation_id") or ""
    return (
        mail.get("store_id", ""),
        conversation,
        mail["id"] if not conversation else "",
    )


def key(identity: tuple[str, str, str]) -> str:
    return hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()


def condition(identity: tuple[str, str, str]) -> tuple[str, list]:
    store, conversation, fallback = identity
    return (
        ("store_id=? AND conversation_id=?", [store, conversation])
        if conversation
        else ("id=?", [fallback])
    )


def previous(sql: sqlite3.Connection, mail_ids: list[str]) -> set[tuple[str, str, str]]:
    if not mail_ids:
        return set()
    return {
        identity(row)
        for row in sql.execute(
            "SELECT id,store_id,conversation_id FROM emails WHERE id IN ("
            + ",".join("?" for _ in mail_ids)
            + ")",
            mail_ids,
        )
    }


def refresh(
    sql: sqlite3.Connection, affected: set[tuple[str, str, str]], cfg: dict
) -> int:
    updated = 0
    for owner in sorted(affected):
        predicate, params = condition(owner)
        thread_key = key(owner)
        count = sql.execute(
            f"SELECT count(*) FROM emails WHERE {predicate}", params
        ).fetchone()[0]
        old = sql.execute(
            "SELECT fingerprint,rowid FROM thread_documents WHERE key=?", [thread_key]
        ).fetchone()
        if not count:
            if old:
                sql.execute("DELETE FROM thread_lexical WHERE rowid=?", [old[1]])
                sql.execute("DELETE FROM thread_documents WHERE key=?", [thread_key])
                updated += 1
            continue
        columns = "id,subject,sender,received,substr(body,1,4000) body"
        first = sql.execute(
            f"SELECT {columns} FROM emails WHERE {predicate} ORDER BY {INSTANT},id LIMIT 2",
            params,
        ).fetchall()
        recent = sql.execute(
            f"SELECT {columns} FROM emails WHERE {predicate} ORDER BY {INSTANT} DESC,id DESC LIMIT ?",
            [*params, SUMMARY_ROWS],
        ).fetchall()
        # Reserve search-text space for the latest source even when unusually
        # large headers would consume the entire excerpt budget.
        selected = {row["id"]: row for row in [recent[0], *first, *reversed(recent)]}
        excerpts: list[str] = []
        sources: list[str] = []
        budget = SUMMARY_CHARS
        for row in selected.values():
            separator = 2 if excerpts else 0
            if budget <= separator:
                break
            excerpt = text.clean_body(
                row["body"] or "", cfg.get("text_cleaning_version", 2)
            )[:300]
            part = f"[{row['id']}] {row['received']} {row['sender']} {row['subject']}\n{excerpt}"
            part = part[: budget - separator]
            if part:
                excerpts.append(part)
                sources.append(row["id"])
                budget -= len(part) + separator
        document = "\n\n".join(excerpts)
        anchor = recent[0]["id"]
        fingerprint = hashlib.sha256(
            json.dumps([count, anchor, document, sources]).encode()
        ).hexdigest()
        if old and old[0] == fingerprint:
            continue
        if old:
            sql.execute("DELETE FROM thread_lexical WHERE rowid=?", [old[1]])
        sql.execute(
            "INSERT OR REPLACE INTO thread_documents VALUES (?,?,?,?,?,?,?,?)",
            [
                thread_key,
                owner[0],
                owner[1],
                anchor,
                count,
                document,
                json.dumps(sources),
                fingerprint,
            ],
        )
        sql.execute(
            "INSERT INTO thread_lexical(rowid,thread_key,terms) VALUES (?,?,?)",
            [
                sql.execute(
                    "SELECT rowid FROM thread_documents WHERE key=?", [thread_key]
                ).fetchone()[0],
                thread_key,
                text.tokens(document),
            ],
        )
        updated += 1
    return updated


def candidates(sql: sqlite3.Connection, expression: str, count: int) -> list[Any]:
    if not expression or not available(sql):
        return []
    return sql.execute(
        """SELECT c.id FROM thread_lexical l
        JOIN thread_documents d ON d.rowid=l.rowid
        JOIN chunks c ON c.email_id=d.anchor_mail_id AND c.position=0
        WHERE thread_lexical MATCH ? AND l.rank MATCH 'bm25()'
        ORDER BY l.rank LIMIT ?""",
        [expression, count],
    ).fetchall()


def status(sql: sqlite3.Connection) -> dict:
    if not available(sql):
        return dict(
            state="missing",
            representation="bounded_original_excerpts",
            model_api_required=False,
        )
    flag = sql.execute(
        "SELECT value FROM state WHERE key='thread_index_complete'"
    ).fetchone()
    return dict(
        state="complete" if flag and flag[0] == "1" else "partial",
        representation="bounded_original_excerpts",
        model_api_required=False,
    )


def contexts(
    sql: sqlite3.Connection, items: list[dict], scope: str, params: list
) -> dict:
    """Attach focus, one neighbor each side, and latest within the requested scope."""
    budget, total_sources = CONTEXT_CHARS, 0
    for item in items:
        cid, store = item.get("conversation_id"), item["store_id"]
        predicate, owner_params = (
            ("e.store_id=? AND e.conversation_id=?", [store, cid])
            if cid
            else ("e.id=?", [item["mail_id"]])
        )
        predicate += " AND " + scope
        columns = "e.id,e.entry_id,e.store_id,e.subject,e.sender,e.received,e.folder,e.conversation_id,substr(e.body,1,1200) body,length(e.body) body_length,coalesce(julianday(e.received),0) instant"
        focus = sql.execute(
            f"SELECT {columns} FROM emails e WHERE e.id=? AND {scope}",
            [item["mail_id"], *params],
        ).fetchone()
        if focus is None:
            item["thread_context"] = dict(items=[], count=0, truncated=True)
            continue
        recent = sql.execute(
            f"SELECT {columns} FROM emails e WHERE {predicate} ORDER BY coalesce(julianday(e.received),0) DESC,e.id DESC LIMIT 1",
            [*owner_params, *params],
        ).fetchone()
        selected = {focus["id"]: focus}
        if recent:
            selected[recent["id"]] = recent
        earliest = sql.execute(
            f"SELECT {columns} FROM emails e WHERE {predicate} ORDER BY coalesce(julianday(e.received),0),e.id LIMIT 1",
            [*owner_params, *params],
        ).fetchone()
        if earliest:
            selected[earliest["id"]] = earliest
        for operator, order in (("<", "DESC"), (">", "ASC")):
            neighbor = sql.execute(
                f"""SELECT {columns} FROM emails e WHERE {predicate}
                AND (coalesce(julianday(e.received),0),e.id){operator}(?,?)
                ORDER BY coalesce(julianday(e.received),0) {order},e.id {order} LIMIT 1""",
                [*owner_params, *params, focus["instant"], focus["id"]],
            ).fetchone()
            if neighbor:
                selected[neighbor["id"]] = neighbor
        # Spend the global budget on focus/latest before other context.
        evidence = []
        for row in selected.values():
            value = dict(row)
            keep = min(1200, budget)
            if keep <= 0:
                continue
            value["mail_id"] = value.pop("id")
            value["body"] = (value["body"] or "")[:keep]
            value["body_truncated"] = (value.pop("body_length") or 0) > keep
            value["is_latest_scoped"] = bool(recent and row["id"] == recent["id"])
            value["role"] = (
                "focus"
                if row["id"] == item["mail_id"]
                else "latest"
                if recent and row["id"] == recent["id"]
                else "first"
                if earliest and row["id"] == earliest["id"]
                else "neighbor"
            )
            evidence.append(value)
            budget -= len(value["body"])
        evidence.sort(key=lambda row: (row.pop("instant"), row["mail_id"]))
        total_sources += len(evidence)
        total = sql.execute(
            f"SELECT count(*) FROM emails e WHERE {predicate}", [*owner_params, *params]
        ).fetchone()[0]
        item["thread_context"] = dict(
            items=evidence,
            count=len(evidence),
            scoped_thread_mail_count=total,
            truncated=total > len(evidence)
            or any(row["body_truncated"] for row in evidence),
            chronological=True,
            source="original_cached_mail",
            summary_is_generated=False,
        )
        item["thread_mail_count"] = total
    return dict(
        added_sources=total_sources,
        body_chars=CONTEXT_CHARS - budget,
        body_limit=CONTEXT_CHARS,
    )


def rebuild(cfg: dict) -> dict:
    from . import app
    from .features import stored_sql

    started = time.perf_counter()
    with app.mutation_lock(cfg), stored_sql(cfg) as source:
        if source is None:
            raise ValueError("Index not initialized")
        with closing(
            sqlite3.connect(Path(cfg["data_dir"]) / "metadata.sqlite", timeout=5)
        ) as sql:
            sql.row_factory = sqlite3.Row
            initialize(sql)
            with sql:
                sql.execute("DELETE FROM thread_documents")
                sql.execute("DELETE FROM thread_lexical")
                updated = 0
                cursor = source.execute(
                    "SELECT store_id,coalesce(conversation_id,'') conversation_id,CASE WHEN coalesce(conversation_id,'')='' THEN id ELSE '' END id FROM emails GROUP BY store_id,conversation_id,CASE WHEN coalesce(conversation_id,'')='' THEN id ELSE '' END"
                )
                while rows := cursor.fetchmany(200):
                    updated += refresh(sql, {identity(row) for row in rows}, cfg)
                sql.execute(
                    "INSERT OR REPLACE INTO state VALUES ('thread_index_complete','1')"
                )
    return dict(
        threads=updated,
        embedding_requests=0,
        summary_model_requests=0,
        elapsed_ms=round((time.perf_counter() - started) * 1000),
    )
