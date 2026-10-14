"""CLI adapters share the same operations and offline behavior as MCP tools."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
from importlib.metadata import version
from typing import Any

from . import features, jobs, storage, threads


def main(backend: Any) -> None:
    parser = argparse.ArgumentParser(
        description="Search and maintain the local Outlook index; no command starts MCP stdio."
    )
    parser.add_argument("--config")
    parser.add_argument("--version", action="version", version=version("outlook-rag"))
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("status")
    sub.add_parser("sources")
    sub.add_parser(
        "set-key",
        help="Prompt for an API key and save it encrypted with Windows DPAPI.",
    )
    sync_parser = sub.add_parser("sync")
    sync_parser.add_argument("--folders", nargs="+")
    sync_parser.add_argument("--max-emails", type=int)
    sync_parser.add_argument("--max-total-emails", type=int)
    sync_parser.add_argument("--since-days", type=int)
    sync_parser.add_argument("--reconcile", action="store_true")
    sync_parser.add_argument("--max-seconds", type=float)
    sync_parser.add_argument("--max-scanned", type=int)
    search_parser = sub.add_parser("search")
    search_parser.add_argument("query", nargs="?", default="", help="Required except for metadata lists")
    search_parser.add_argument("--limit", type=int, default=10)
    for name in ("folder", "sender", "since", "until"):
        search_parser.add_argument("--" + name)
    search_parser.add_argument("--min-similarity", type=float)
    search_parser.add_argument("--semantic-only", action="store_true")
    search_parser.add_argument(
        "--mode", choices=("semantic", "hybrid", "lexical", "auto", "identifier", "metadata", "thread")
    )
    for name in ("subject", "recipient", "cursor"):
        search_parser.add_argument("--" + name)
    search_parser.add_argument("--group-by-thread", action="store_true")
    search_parser.add_argument("--has-attachments", action=argparse.BooleanOptionalAction, default=None)
    search_parser.add_argument("--no-deduplicate", action="store_true")
    job = sub.add_parser("job", help="Start, inspect or cancel a detached bulk sync.")
    job_sub = job.add_subparsers(dest="action", required=True)
    job_start = job_sub.add_parser("start", help="Start or reuse a bulk job.")
    job_start.add_argument("--folders", nargs="+")
    job_start.add_argument("--since-days", type=int)
    job_start.add_argument("--reconcile", action="store_true")
    for action in ("status", "cancel"):
        action_parser = job_sub.add_parser(action)
        action_parser.add_argument("job_id", nargs="?")
    failure = sub.add_parser(
        "failures", help="Inspect and retry retained failed emails."
    )
    failure_sub = failure.add_subparsers(dest="action", required=True)
    failure_list = failure_sub.add_parser("list")
    failure_list.add_argument("--limit", type=int, default=50)
    failure_list.add_argument("--offset", type=int, default=0)
    failure_list.add_argument("--folder")
    failure_retry = failure_sub.add_parser("retry")
    failure_retry.add_argument("mail_ids", nargs="+")
    failure_retry.add_argument("--max-seconds", type=float, default=30)
    sub.add_parser("progress", help="Read cached per-folder progress and estimates.")
    thread = sub.add_parser("thread", help="Read cached mail in the same conversation.")
    thread.add_argument("mail_id")
    thread.add_argument("--before", type=int, default=3)
    thread.add_argument("--after", type=int, default=3)
    thread.add_argument("--max-body-chars", type=int, default=2000)
    sub.add_parser("optimize")
    sub.add_parser("maintain")
    sub.add_parser("doctor", help="Check the saved index without the embedding API.")
    sub.add_parser("rebuild-threads", help="Rebuild extractive thread documents from cached original mail; no API.")
    sub.add_parser(
        "rebuild-vectors",
        help="Rebuild search indexes using cached embeddings; no API calls.",
    )
    backup = sub.add_parser(
        "backup-index",
        help="Back up the paired SQLite and vector index to a new directory.",
    )
    backup.add_argument("destination")
    restore = sub.add_parser(
        "restore-index", help="Restore a backup to a new directory and verify it."
    )
    restore.add_argument("source")
    restore.add_argument("destination")
    worker_parser = sub.add_parser("_job-worker", help=argparse.SUPPRESS)
    worker_parser.add_argument("job_id")
    args = parser.parse_args()
    if args.config:
        os.environ["OUTLOOK_RAG_CONFIG"] = args.config
    if args.command is None:
        from . import entry
        entry.serve(backend.mcp, backend._tools)
        return
    offline = (
        args.command
        in (
            "status",
            "progress",
            "thread",
            "doctor",
            "rebuild-vectors",
            "rebuild-threads",
            "backup-index",
            "restore-index",
            "optimize",
            "maintain",
            "search",
        )
        or (args.command == "job" and args.action != "start")
        or (args.command == "failures" and args.action == "list")
    )
    if offline:
        cfg = features.local_config()
    else:
        cfg = backend.config(resolve_dimensions=args.command != "set-key")
    if args.command == "_job-worker":
        jobs.run(args.job_id, cfg)
        return
    if args.command != "set-key" and not offline:
        backend.api_key(cfg)
    if args.command == "set-key":
        import getpass
        import win32crypt

        target = cfg.get("key_file")
        if not target:
            raise ValueError(
                "Set key_file in the configuration before running set-key."
            )
        secret = getpass.getpass("Embedding API key (hidden): ").strip()
        if not secret:
            raise ValueError("API key cannot be empty.")
        path = Path(target)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(
            win32crypt.CryptProtectData(
                secret.encode(), "Outlook RAG embedding API", None, None, None, 0
            )
        )
        result = {
            "saved": True,
            "key_file": str(path),
            "encryption": "Windows DPAPI CurrentUser",
        }
    elif args.command == "status":
        result = backend.status()
    elif args.command == "sources":
        result = backend.sources()
    elif args.command == "sync":
        result = backend.sync(
            args.folders,
            args.max_emails,
            args.since_days,
            args.reconcile,
            args.max_total_emails,
            args.max_seconds,
            args.max_scanned,
        )
    elif args.command == "search":
        options = {"mode": args.mode} if args.mode is not None else {}
        options.update({name: getattr(args, name) for name in ("subject", "recipient", "cursor") if getattr(args, name) is not None})
        if args.has_attachments is not None:
            options["has_attachments"] = args.has_attachments
        if args.no_deduplicate:
            options["deduplicate"] = False
        if args.semantic_only and args.mode is not None:
            parser.error("Choose --mode or --semantic-only")
        result = backend.search(
            args.query,
            args.limit,
            args.folder,
            args.sender,
            args.since,
            args.until,
            not args.semantic_only,
            args.group_by_thread,
            args.min_similarity,
            **options,
        )
    elif args.command == "job":
        if args.action == "start":
            days = args.since_days if args.since_days is not None else backend.config().get("since_days", 0)
            result = backend.start_sync_job(
                args.folders, 0 if days == "all" else days, args.reconcile, True
            )
        elif args.action == "status":
            result = backend.sync_job_status(args.job_id)
        else:
            result = backend.cancel_sync_job(args.job_id)
    elif args.command == "failures":
        if args.action == "list":
            result = backend.list_sync_failures(args.limit, args.offset, args.folder)
        else:
            result = backend.retry_failed_emails(args.mail_ids, args.max_seconds)
    elif args.command == "progress":
        result = backend.get_sync_progress()
    elif args.command == "thread":
        result = backend.get_mail_thread(
            args.mail_id, args.before, args.after, args.max_body_chars
        )
    elif args.command == "optimize":
        result = backend.optimize()
    elif args.command == "maintain":
        result = backend.maintain()
    elif args.command == "doctor":
        result = storage.doctor(cfg)
    elif args.command == "rebuild-vectors":
        result = storage.rebuild_vectors(cfg)
    elif args.command == "rebuild-threads":
        result = threads.rebuild(cfg)
    elif args.command == "backup-index":
        result = storage.backup_index(cfg, args.destination)
    elif args.command == "restore-index":
        result = storage.restore_index(cfg, args.source, args.destination)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
