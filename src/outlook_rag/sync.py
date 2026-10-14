"""Bounded Outlook synchronization and caller-thread read-ahead.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

from collections import deque
import hashlib
import json
import time
from datetime import datetime, timedelta
from typing import Any

import httpx

from . import coverage, storage, paths, outlook, mail_metadata

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

    def close(self) -> None:
        """Drop queued COM items/exception frames before leaving the apartment."""
        self.buffer.clear()
        self.error = None
        close = getattr(self.source, "close", None)
        if close:
            close()


def sync(backend: Any, folders: list[str] | None = None, max_emails: int | None = None, since_days: int | None = None,
         reconcile: bool = False, max_total_emails: int | None = None,
         max_seconds: float | None = None, max_scanned: int | None = None, *, _cancel_check: Any=None) -> dict:
    cfg = backend.config()
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
    cfg = paths.prepare(cfg)
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
    with backend.mutation_lock(cfg), httpx.Client(timeout=httpx.Timeout(120, connect=10), limits=limits, trust_env=False) as client:
        pythoncom.CoInitialize()
        sql = None
        namespace: Any = None
        folder: Any = None
        next_folder: Any = None
        item: Any = None
        metadata: Any = None
        ahead: Any = None
        targets: list = []
        pending: list = []
        try:
            namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            sql, table = backend.open_store(cfg)
            alias_session = backend.outlook_session_key(namespace)
            discovery = None
            if selection == "auto":
                discovery, targets = backend.discover_sources(namespace, cfg)
                discovery_errors = discovery["errors"]
            else:
                targets = [(name, backend.resolve_folder(namespace, name)) for name in selection]
            target_folders = [dict(folder=name, folder_id=backend.folder_identity(folder)) for name, folder in targets]
            with sql:
                sql.execute("INSERT OR REPLACE INTO state VALUES ('folder_inventory',?)",
                            (json.dumps(coverage.inventory(discovery, targets, selection, since_days)),))
            rotation_key = "rotation:" + hashlib.sha256(json.dumps([selection, since_days, reconcile], sort_keys=True).encode()).hexdigest()
            rotation = sql.execute("SELECT value FROM state WHERE key=?", (rotation_key,)).fetchone()
            if rotation:
                position = next((i for i, (_, folder) in enumerate(targets)
                    if f"{outlook.com_attr(folder, 'StoreID')}:{outlook.com_attr(folder, 'EntryID')}" == rotation[0]), 0)
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
                    metadata.update(item=item, modified=outlook.item_date(item, "LastModificationTime", outlook.MODIFIED), received=outlook.item_date(item, "ReceivedTime", outlook.RECEIVED), _attachment_cfg=cfg)
                    mail = backend.read_outlook_mail(namespace, metadata, failure["store_id"], failure["folder"])
                except Exception as exc:
                    backend.record_sync_failure(sql, metadata, failure["store_id"], failure["folder"], exc)
                    failed_total += 1
                    continue
                try:
                    backend.upsert_mails([mail], sql, table, cfg, client)
                    with sql:
                        sql.execute("DELETE FROM sync_failures WHERE mail_id IN (?,?)", (mail["id"], failure["mail_id"]))
                    indexed_total += 1
                except Exception as exc:
                    if isinstance(exc, httpx.HTTPError) and not backend.permanent_mail_error(exc):
                        stop_reason = "cancelled" if isinstance(exc, backend.SyncCancelled) else "time_limit" if time.perf_counter() >= deadline else "embedding_unavailable"
                        break
                    if not backend.permanent_mail_error(exc):
                        raise
                    backend.record_sync_failure(sql, metadata, failure["store_id"], failure["folder"], exc)
                    failed_total += 1

            for target_position, (folder_name, folder) in enumerate(targets):
                if cancelled():
                    stop_reason = "cancelled"
                if stop_reason or attempted_total >= max_total_emails or scanned_total >= max_scanned or time.perf_counter() >= deadline:
                    stop_reason = stop_reason or ("email_limit" if attempted_total >= max_total_emails else "scan_limit" if scanned_total >= max_scanned else "time_limit")
                    pending_folders = [name for name, _ in targets[target_position:]]
                    break
                store_id = outlook.com_attr(folder, "StoreID")
                state_key = f"sync:{store_id}:{outlook.com_attr(folder, 'EntryID')}:{since_days}"
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
                filters = [f"[ReceivedTime] >= '{outlook.date_literal(cutoff)}'"] if cutoff else []
                if mode == "incremental" and scan_from:
                    modified = datetime.fromisoformat(scan_from).astimezone() - timedelta(minutes=5)
                    filters.append(f"[LastModificationTime] >= '{outlook.date_literal(modified)}'")
                if cursor:
                    boundary = datetime.fromisoformat(cursor).astimezone() + timedelta(minutes=1)
                    filters.append(f"[ReceivedTime] <= '{outlook.date_literal(boundary)}'")
                if mode == "reconcile" and not continuing:
                    with sql:
                        sql.execute("DELETE FROM sync_seen WHERE scan_key=?", (state_key,))
                indexed = skipped = scanned = failed = attempted = overlap = 0
                reconcile_uncertain = checkpoint.get("reconcile_uncertain", False) if continuing else False
                complete = True
                batch_stats: dict[str, Any] = {}
                pending = []
                metadata_source = None
                progress_started = time.perf_counter()
                processed_base = checkpoint.get("processed_rows", 0) if continuing else 0
                processing_base = checkpoint.get("processing_seconds", 0.0) if continuing else 0.0
                progress_known = checkpoint.get("progress_known", "processed_rows" in checkpoint) if continuing and cursor else not cursor
                advanced_rows = 0
                source_count = backend.safe_attr(backend.safe_attr(folder, "Items", None), "Count", None)
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
                        stats = backend.upsert_mails([mail for mail, _ in pending], sql, table, cfg, client)
                        indexed += len(pending)
                        indexed_total += len(pending)
                        for key, value in stats.items():
                            batch_stats[key] = batch_stats.get(key, 0) + value
                        for mail, metadata in pending:
                            sql.execute("DELETE FROM sync_failures WHERE mail_id=?", (mail["id"],))
                            advance(metadata)
                    except Exception as exc:
                        if not backend.permanent_mail_error(exc):
                            raise
                        # Isolate malformed input without hiding provider outages or storage errors.
                        for mail, metadata in pending:
                            try:
                                stats = backend.upsert_mails([mail], sql, table, cfg, client)
                                indexed += 1
                                indexed_total += 1
                                for key, value in stats.items():
                                    batch_stats[key] = batch_stats.get(key, 0) + value
                                sql.execute("DELETE FROM sync_failures WHERE mail_id=?", (mail["id"],))
                            except Exception as single_exc:
                                if not backend.permanent_mail_error(single_exc):
                                    raise
                                backend.record_sync_failure(sql, metadata, store_id, folder_name, single_exc)
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
                        metadata = backend.canonical_metadata(namespace, metadata, store_id, sql, alias_session)
                    except Exception as exc:
                        return (dict(metadata, _canonical_error=exc), source)
                    metadata = dict(metadata, _canonical=True, _attachment_cfg=cfg)
                    mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
                    previous = sql.execute("SELECT modified FROM emails WHERE id=?", (mail_id,)).fetchone()
                    failure = sql.execute("SELECT modified FROM sync_failures WHERE mail_id=?", (mail_id,)).fetchone()
                    if not ((previous and backend.same_modified(previous["modified"], metadata["modified"])) or
                            (failure and backend.same_modified(failure["modified"], metadata["modified"]))):
                        try:
                            metadata["_prefetched_mail"] = backend.read_outlook_mail(namespace, metadata, store_id, folder_name)
                        except Exception as exc:
                            metadata["_read_error"] = exc
                    return metadata, source

                ahead = backend.ReadAhead(backend.mail_metadata(folder, filters), prepare_ahead,
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
                                metadata = backend.canonical_metadata(namespace, metadata, store_id, sql, alias_session)
                        except Exception as exc:
                            flush()
                            attempted += 1
                            attempted_total += 1
                            backend.record_sync_failure(sql, metadata, store_id, folder_name, exc)
                            failed += 1
                            failed_total += 1
                            reconcile_uncertain = True
                            advance(metadata)
                            continue
                        mail_id = hashlib.sha256(f"{store_id}:{metadata['entry_id']}".encode()).hexdigest()
                        previous = sql.execute("SELECT modified FROM emails WHERE id=?", (mail_id,)).fetchone()
                        failure = sql.execute("SELECT modified FROM sync_failures WHERE mail_id=?", (mail_id,)).fetchone()
                        if (previous and backend.same_modified(previous["modified"], metadata["modified"])) or (failure and backend.same_modified(failure["modified"], metadata["modified"])):
                            flush()
                            if previous:
                                storage.update_folder(sql, table, folder_name, mail_id=mail_id)
                                storage.update_dates(sql, table, mail_id, metadata["received"], metadata["modified"], cfg)
                                if not sql.execute("SELECT 1 FROM mail_metadata WHERE email_id=?", (mail_id,)).fetchone():
                                    item = metadata.get("item")
                                    if item is None:
                                        item = namespace.GetItemFromID(metadata["entry_id"], store_id)
                                    with sql:
                                        mail_metadata.save(sql, dict(id=mail_id, **mail_metadata.read(item)))
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
                                metadata["_attachment_cfg"] = cfg
                                prepared_mail = backend.read_outlook_mail(namespace, metadata, store_id, folder_name)
                        except Exception as exc:
                            flush()
                            backend.record_sync_failure(sql, metadata, store_id, folder_name, exc)
                            failed += 1
                            failed_total += 1
                            advance(metadata)
                            continue
                        pending.append((prepared_mail, metadata))
                        if len(pending) >= cfg.get("sync_batch_emails", 16):
                            flush()
                    flush()
                except httpx.HTTPError as exc:
                    complete, stop_reason = False, "cancelled" if isinstance(exc, backend.SyncCancelled) else "time_limit" if time.perf_counter() >= deadline else "embedding_unavailable"
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
                        backend.remove_mail(candidate["id"], sql, table)
                        removed += 1
                    if cutoff is None:
                        sql.execute("DELETE FROM sync_failures WHERE store_id=? AND folder=? AND mail_id NOT IN (SELECT mail_id FROM sync_seen WHERE scan_key=?)",
                            (store_id, folder_name, state_key))
                    sql.execute("DELETE FROM sync_seen WHERE scan_key=?", (state_key,))
                save_state(complete)
                reports.append(dict(folder=folder_name, folder_id=backend.folder_identity(folder), indexed=indexed, unchanged=skipped, scanned=scanned,
                    cursor_overlap_rows=overlap, failed=failed, window_complete=complete, incremental=mode == "incremental",
                    removed_local=removed, reconcile_skipped=complete and mode == "reconcile" and reconcile_uncertain,
                    metadata_source=metadata_source, batch_stats=batch_stats))
                if targets:
                    next_folder = targets[(target_position + 1) % len(targets)][1]
                    with sql:
                        sql.execute("INSERT OR REPLACE INTO state VALUES (?,?)", (rotation_key, f"{outlook.com_attr(next_folder, 'StoreID')}:{outlook.com_attr(next_folder, 'EntryID')}"))
            if not pending_folders and targets and stop_reason:
                pending_folders = [name for name, _ in targets[len(reports):]]
            if selection == "auto" and since_days == 0 and not stop_reason and not discovery_errors and not folder_errors and not failed_total and len(reports) == len(targets) and all(r["window_complete"] for r in reports):
                connected = {(str(outlook.com_attr(folder, "StoreID")), name) for name, folder in targets}
                cached = {(str(r[0]), r[1]) for r in sql.execute("SELECT DISTINCT store_id,folder FROM emails")}
                if cached <= connected:
                    with sql:
                        sql.execute("INSERT OR REPLACE INTO state VALUES ('outlook_dates_need_rescan','0')")
        finally:
            try:
                cfg.pop("_embedding_idle", None)
                cfg.pop("_embedding_idle_ready", None)
                try:
                    if ahead is not None:
                        ahead.close()
                finally:
                    pending.clear()
                    metadata = item = folder = next_folder = namespace = None
                    targets.clear()
            finally:
                try:
                    if sql:
                        sql.close()
                finally:
                    pythoncom.CoUninitialize()
    return dict(folders=reports, indexed_total=indexed_total, attempted_total=attempted_total,
        scanned_total=scanned_total, failed_total=failed_total, retried=retried,
        pending_folders=pending_folders, target_folders=target_folders,
        discovery_errors=discovery_errors, folder_errors=folder_errors, stop_reason=stop_reason,
        elapsed_ms=round((time.perf_counter() - started) * 1000),
        note="Repeat incomplete windows and pending folders. Time limits are cooperative; in-flight Outlook/API calls may exceed the budget. Failed emails are retained for retry.")

