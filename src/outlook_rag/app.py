from __future__ import annotations

from contextlib import contextmanager
import logging
import sqlite3
import sys
import threading
import time as time
from typing import Any, Iterator

import httpx
import lancedb as lancedb
from filelock import FileLock
from mcp.server.fastmcp import FastMCP

from outlook_rag import cli, search as search_engine, server
from outlook_rag.features import local_config as local_config, stored_sql as stored_sql, folder_progress as folder_progress, failure_reason as failure_reason
from outlook_rag.outlook import (
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
from outlook_rag.text import chunks as chunks, clean_body as clean_body, safe_text as safe_text, tokens as tokens

from outlook_rag import features as features_engine
from outlook_rag import settings as settings_engine
from outlook_rag import embeddings as embeddings_engine
from outlook_rag import indexing as indexing_engine
from outlook_rag import sync as sync_engine
from outlook_rag import maintenance as maintenance_engine
from outlook_rag import database as database_engine
from outlook_rag.embeddings import SyncCancelled as SyncCancelled
from outlook_rag.sync import ReadAhead as ReadAhead

QUERY_INSTRUCTION = "Retrieve email passages relevant to the user's question, including paraphrases and related concepts."
logging.getLogger("httpx").setLevel(logging.WARNING)
_LOCK = threading.RLock()
_DIMENSIONS: dict[tuple, int] = {}
_API_DIMENSIONS: dict[tuple, bool] = {}


def config(*, resolve_dimensions: bool = True) -> dict:
    return settings_engine.config(sys.modules[__name__], resolve_dimensions=resolve_dimensions)

def detect_dimensions(cfg: dict) -> int:
    return embeddings_engine.detect_dimensions(sys.modules[__name__], cfg)

def query_instruction(cfg: dict) -> str:
    return settings_engine.query_instruction(sys.modules[__name__], cfg)

def index_identity(cfg: dict) -> str:
    return settings_engine.index_identity(sys.modules[__name__], cfg)

def api_key(cfg: dict) -> str:
    return settings_engine.api_key(sys.modules[__name__], cfg)

def vector_dimensions(cfg: dict) -> int:
    return settings_engine.vector_dimensions(sys.modules[__name__], cfg)

def embed(texts: list[str], cfg: dict, *, query: bool = False, client: httpx.Client | None = None) -> list[list[float]]:
    return embeddings_engine.embed(sys.modules[__name__], texts, cfg, query=query, client=client)


def mutation_lock(cfg: dict) -> FileLock:
    return database_engine.mutation_lock(sys.modules[__name__], cfg)

def reader_lock(cfg: dict) -> FileLock:
    return database_engine.reader_lock(sys.modules[__name__], cfg)

def pin_vectors(sql: sqlite3.Connection, table: Any) -> int:
    return database_engine.pin_vectors(sys.modules[__name__], sql, table)

def publish_vectors(sql: sqlite3.Connection, table: Any) -> None:
    return database_engine.publish_vectors(sys.modules[__name__], sql, table)


@contextmanager
def indexed_snapshot(sql: sqlite3.Connection, table: Any, cfg: dict) -> Iterator[Any]:
    with database_engine.indexed_snapshot(sys.modules[__name__], sql, table, cfg) as table:
        yield table


@contextmanager
def vector_publication(sql: sqlite3.Connection, table: Any) -> Iterator[None]:
    with database_engine.vector_publication(sys.modules[__name__], sql, table):
        yield


def open_store(cfg: dict) -> tuple[sqlite3.Connection, Any]:
    return database_engine.open_store(sys.modules[__name__], cfg)

def upsert_mail(mail: dict, sql: sqlite3.Connection, table: Any, cfg: dict) -> dict:
    return indexing_engine.upsert_mail(sys.modules[__name__], mail, sql, table, cfg)

def upsert_mails(mails: list[dict], sql: sqlite3.Connection, table: Any, cfg: dict, client: httpx.Client | None = None) -> dict:
    return indexing_engine.upsert_mails(sys.modules[__name__], mails, sql, table, cfg, client)


def sources() -> dict:
    return features_engine.sources(sys.modules[__name__])


def record_sync_failure(sql: sqlite3.Connection, metadata: dict, store_id: str, folder: str | None, exc: Exception) -> None:
    return indexing_engine.record_sync_failure(sys.modules[__name__], sql, metadata, store_id, folder, exc)


def permanent_mail_error(exc: Exception) -> bool:
    return indexing_engine.permanent_mail_error(sys.modules[__name__], exc)

def sync(folders: list[str] | None = None, max_emails: int | None = None, since_days: int | None = None,
         reconcile: bool = False, max_total_emails: int | None = None,
         max_seconds: float | None = None, max_scanned: int | None = None, *, _cancel_check: Any=None) -> dict:
    return sync_engine.sync(sys.modules[__name__], folders, max_emails, since_days, reconcile, max_total_emails, max_seconds, max_scanned, _cancel_check=_cancel_check)

def remove_mail(mail_id: str, sql: sqlite3.Connection, table: Any) -> None:
    return indexing_engine.remove_mail(sys.modules[__name__], mail_id, sql, table)

def cached_query_vector(query: str, sql: sqlite3.Connection, cfg: dict) -> tuple[list[float], bool]:
    return embeddings_engine.cached_query_vector(sys.modules[__name__], query, sql, cfg)

def cache_query(sql: sqlite3.Connection, statement: str, values: tuple, size: int | None = None) -> None:
    return embeddings_engine.cache_query(sys.modules[__name__], sql, statement, values, size)

def search(query: str, limit: int = 10, folder: str | None = None, sender: str | None = None,
           since: str | None = None, until: str | None = None, hybrid: bool = True, group_by_thread: bool = False,
           min_similarity: float | None = None, mode: str | None = None, *,
           subject: str | None = None, recipient: str | None = None, cursor: str | None = None,
           has_attachments: bool | None = None, deduplicate: bool = True) -> dict:
    return search_engine.search(sys.modules[__name__], query, limit, folder, sender, since, until,
                                hybrid, group_by_thread, min_similarity, mode,
                                subject=subject, recipient=recipient, cursor=cursor, has_attachments=has_attachments, deduplicate=deduplicate)

def status() -> dict:
    return features_engine.status(sys.modules[__name__])


def get_mail(mail_id: str, max_body_chars: int = 10000) -> dict:
    return features_engine.get_mail(sys.modules[__name__], mail_id, max_body_chars)

def optimize() -> dict:
    return maintenance_engine.optimize(sys.modules[__name__])

def maintain() -> dict:
    return maintenance_engine.maintain(sys.modules[__name__])


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
sync_index = _tools['sync_index']

def main() -> None:
    cli.main(sys.modules[__name__])


if __name__ == "__main__":
    main()
