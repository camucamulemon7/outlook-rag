"""MCP adapters for index operations and detached job control."""

from __future__ import annotations

from typing import Any, Callable
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from . import features, jobs


def register_tools(mcp: FastMCP, backend: Any) -> dict[str, Callable[..., dict]]:
    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def search_emails(
        query: str,
        limit: int = 10,
        folder: str | None = None,
        sender: str | None = None,
        since: str | None = None,
        until: str | None = None,
        hybrid: bool = True,
        group_by_thread: bool = False,
        min_similarity: float | None = None,
        mode: str | None = None,
    ) -> dict:
        """Search indexed mail. mode: semantic (paraphrases), hybrid (semantic + BM25), lexical (keywords), identifier (one exact ID such as INV-1047), auto (semantic or unambiguous single ID). Default preserves hybrid behavior. Auto is opt-in; comparisons/negation use semantic. since/until are inclusive dates. min_similarity (-1..1) is a model-calibrated cosine filter, not confidence. group_by_thread keeps the strongest match per store/conversation."""
        options = {"mode": mode} if mode is not None else {}
        return backend.search(
            query,
            limit,
            folder,
            sender,
            since,
            until,
            hybrid,
            group_by_thread,
            min_similarity,
            **options,
        )

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def index_status() -> dict:
        """Index coverage, last indexing times, model, vector index and incomplete sync windows."""
        return backend.status()

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def list_outlook_sources() -> dict:
        """List Outlook-connected mail stores, PST/OST file paths when available, and folders selected by automatic discovery. No email bodies are read."""
        return backend.sources()

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def get_indexed_mail(mail_id: str, max_body_chars: int = 10000) -> dict:
        """Read a cached original email body using mail_id returned by search_emails. No live Outlook access."""
        return backend.get_mail(mail_id, max_body_chars)

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def sync_emails(
        folders: list[str] | None = None,
        max_emails: int | None = None,
        since_days: int | None = None,
        reconcile: bool = False,
        max_total_emails: int | None = None,
        max_seconds: float | None = None,
        max_scanned: int | None = None,
    ) -> dict:
        """Update the local index without Outlook writes. Defaults: all dates, 200 mail attempts, 2000 new metadata rows, 30-second cooperative budget. In-flight calls may overrun. Repeat incomplete windows/pending folders; folders rotate fairly. Failed emails are retained for retry. Reconciliation only deletes local stale entries after a complete scan."""
        return backend.sync(
            folders,
            max_emails,
            since_days,
            reconcile,
            max_total_emails,
            max_seconds,
            max_scanned,
        )

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def optimize_index() -> dict:
        """Build a local approximate vector index after bulk sync. Use from 256 chunks onward."""
        return backend.optimize()

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def maintain_index() -> dict:
        """Remove unused embedding caches and compact local vector storage; retain recent versions for seven days. No Outlook writes."""
        return backend.maintain()

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def start_sync_job(
        folders: list[str] | None = None,
        since_days: int | None = None,
        reconcile: bool = False,
    ) -> dict:
        """Start one detached bulk sync worker per database. Returns promptly with a job_id. Continues across MCP disconnects; uses existing sync limits and checkpoints. Inspect with sync_job_status; cancel cooperatively with cancel_sync_job."""
        return jobs.start(backend.config(), folders, since_days, reconcile)

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def sync_job_status(job_id: str | None = None) -> dict:
        """Read durable bulk job progress, counters and current index coverage. Omit job_id for the latest job. Completed_with_errors means some mail or folders need retry."""
        cfg = features.local_config()
        result = jobs.status(cfg, job_id)
        result["folder_progress"] = features.progress(cfg)["folders"]
        return result

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def cancel_sync_job(job_id: str | None = None) -> dict:
        """Request cancellation of the latest or specified bulk sync job. Stops cooperatively at safe boundaries; in-flight API/Outlook calls may finish. Keeps indexed mail and completed embedding caches."""
        return jobs.cancel(features.local_config(), job_id)

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def list_sync_failures(
        limit: int = 50, offset: int = 0, folder: str | None = None
    ) -> dict:
        """List retained failed mail IDs, sanitized causes, retry counts and scheduled retry times. No mail bodies or provider calls."""
        return features.failures(features.local_config(), limit, offset, folder)

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))
    def retry_failed_emails(mail_ids: list[str], max_seconds: float = 30) -> dict:
        """Retry only 1..100 specified failure IDs immediately, bypassing backoff. Read-only Outlook access; preserve the index on failure. Time budget is cooperative."""
        return features.retry(backend.config(), mail_ids, max_seconds)

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def get_sync_progress() -> dict:
        """Read per-folder scan progress, indexed/failed counts and estimated remaining items, processing rate and ETA when known. No live Outlook calls."""
        return features.progress(features.local_config())

    @mcp.tool(annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False))
    def get_mail_thread(
        mail_id: str, before: int = 3, after: int = 3, max_body_chars: int = 2000
    ) -> dict:
        """Return the selected cached mail and preceding/following mail in the same Outlook store/conversation, chronologically. No live Outlook or embedding calls."""
        return features.thread(
            features.local_config(), mail_id, before, after, max_body_chars
        )

    return {
        "search_emails": search_emails,
        "index_status": index_status,
        "list_outlook_sources": list_outlook_sources,
        "get_indexed_mail": get_indexed_mail,
        "sync_emails": sync_emails,
        "optimize_index": optimize_index,
        "maintain_index": maintain_index,
        "start_sync_job": start_sync_job,
        "sync_job_status": sync_job_status,
        "cancel_sync_job": cancel_sync_job,
        "list_sync_failures": list_sync_failures,
        "retry_failed_emails": retry_failed_emails,
        "get_sync_progress": get_sync_progress,
        "get_mail_thread": get_mail_thread,
    }
