"""Failure recovery, durable scan progress and cached thread context."""
from __future__ import annotations

import json
import hashlib
import math
import os
import sqlite3
from contextlib import contextmanager, closing
from pathlib import Path
import re
import time
from datetime import datetime, timezone

import httpx
from filelock import Timeout


def local_config():
    """Resolve a cached index without probing an unavailable embedding API."""
    from . import app
    cfg = app.config(resolve_dimensions=False)
    if cfg.get("dimensions"):
        return cfg
    if cfg.get("_automatic_data_dir"):
        roots = [Path(cfg["data_dir"]).parent,
                 (Path(os.environ["LOCALAPPDATA"]) if os.environ.get("LOCALAPPDATA") else Path.home()/"AppData"/"Local")/"outlook-rag"]
        candidates = []
        for root in roots:
            if root.is_dir():
                candidates.extend(sorted((p for p in root.iterdir() if p.is_dir() and (p/"metadata.sqlite").is_file()), key=lambda p: (bool(re.fullmatch(r"[0-9a-f]{12}",p.name)),p.name)))
    else:
        candidates = [Path(cfg["data_dir"])]
    matches = []
    for folder in candidates:
        path = folder/"metadata.sqlite"
        if not path.is_file():
            continue
        with closing(sqlite3.connect(path.resolve().as_uri()+"?mode=ro",uri=True,timeout=5)) as sql:
            row = sql.execute("SELECT value FROM state WHERE key='embedding_identity'").fetchone()
        if not row:
            continue
        identity = json.loads(row[0])
        dimensions = identity.get("dimensions")
        if type(dimensions) is not int or not 1 <= dimensions <= 65536:
            continue
        candidate = dict(cfg,dimensions=dimensions,data_dir=str(folder))
        if "qwen3-embedding" in cfg["model"].lower():
            candidate.setdefault("storage_dimensions",min(1024,dimensions))
            candidate.setdefault("embedding_request_dimensions",candidate["storage_dimensions"])
        if json.loads(app.index_identity(candidate)) == identity:
            matches.append(candidate)
    identities = {app.index_identity(c) for c in matches}
    if len(identities)>1:
        raise ValueError("Multiple cached indexes match; set OUTLOOK_RAG_DIMENSIONS or OUTLOOK_RAG_DATA_DIR to select one.")
    if matches:
        if cfg.get("_automatic_data_dir"):
            root = Path(cfg["data_dir"]).parent
            prefix = Path(cfg["data_dir"]).name.rsplit("-",2)[0]
            def preference(candidate):
                digest = hashlib.sha256(app.index_identity(candidate).encode()).hexdigest()[:12]
                folder = Path(candidate["data_dir"])
                named = root/f"{prefix}-{app.vector_dimensions(candidate)}d-{digest}"
                return 0 if folder==named else 1 if folder==root/digest else 2 if folder.name==digest else 3
            matches.sort(key=preference)
        return matches[0]
    if not cfg.get("_automatic_data_dir") and (Path(cfg["data_dir"])/"metadata.sqlite").is_file():
        raise ValueError("Cached index settings do not match the configured model or chunk settings.")
    return cfg


@contextmanager
def stored_sql(cfg):
    path = Path(cfg["data_dir"])/"metadata.sqlite"
    if not path.is_file():
        yield None
        return
    sql = sqlite3.connect(path.resolve().as_uri()+"?mode=ro",uri=True,timeout=5)
    sql.row_factory = sqlite3.Row
    try:
        identity = sql.execute("SELECT value FROM state WHERE key='embedding_identity'").fetchone()
        from . import app
        if not identity or json.loads(identity[0]) != json.loads(app.index_identity(cfg)):
            raise ValueError("Cached index settings do not match; select the matching model, dimensions and chunk settings.")
        yield sql
    finally:
        sql.close()


def failure_reason(exc):
    # Never retain response bodies, request text, URLs, headers or credentials.
    if isinstance(exc, httpx.HTTPStatusError):
        code = exc.response.status_code
        hint = {400: "Input or API parameters rejected", 401: "API authentication failed",
                403: "API access denied", 413: "Embedding input too large",
                422: "Invalid embedding input", 429: "API rate limit reached"}.get(code, "Embedding provider returned an error")
        return f"HTTP {code}: {hint}."
    if isinstance(exc, httpx.TimeoutException):
        return "Embedding request timed out."
    if isinstance(exc, httpx.ConnectError):
        return "Cannot connect to the embedding provider."
    if isinstance(exc, UnicodeError):
        return "Mail text contains invalid Unicode."
    if isinstance(exc, ValueError):
        return "Invalid mail text, embedding vector, or index configuration."
    if isinstance(getattr(exc, "hresult", None), int):
        return f"Outlook COM access failed (HRESULT {exc.hresult})."
    if getattr(exc, "sqlite_errorname", None):
        return f"Local database error: {exc.sqlite_errorname}."
    if isinstance(exc, PermissionError):
        return "Access denied to mail or local index files."
    return f"Mail read or index update failed ({type(exc).__name__})."


def failures(cfg, limit=50, offset=0, folder=None):
    from . import app
    if type(limit) is not int or not 1 <= limit <= 200 or type(offset) is not int or not 0 <= offset <= 100000:
        raise ValueError("limit must be 1..200 and offset 0..100000")
    with app._LOCK, stored_sql(cfg) as sql:
        if sql is None:
            return dict(items=[],count=0,total=0,offset=offset,next_offset=None)
        if sql is not None:
            where, args = ("WHERE f.folder=?", [folder]) if folder is not None else ("", [])
            total = sql.execute(f"SELECT count(*) FROM sync_failures f {where}", args).fetchone()[0]
            detailed = sql.execute("SELECT 1 FROM sqlite_master WHERE name='sync_failure_details' AND type='table'").fetchone()
            details_join = "LEFT JOIN sync_failure_details d ON d.mail_id=f.mail_id AND d.last_failed=f.last_failed" if detailed else ""
            details_column = "d.error_detail" if detailed else "NULL AS error_detail"
            rows = sql.execute(f"""SELECT f.*, e.subject, {details_column} FROM sync_failures f LEFT JOIN emails e ON e.id=f.mail_id
                {details_join}
                {where} ORDER BY f.last_failed DESC,f.mail_id LIMIT ? OFFSET ?""", [*args, limit, offset]).fetchall()
            items = []
            for row in rows:
                value = dict(row)
                value['retry_due'] = value['next_retry'] <= time.time()
                for field in ('next_retry', 'last_failed'):
                    value[field] = datetime.fromtimestamp(value[field], timezone.utc).isoformat()
                value['error_detail'] = value['error_detail'] or 'No detailed cause recorded; retry to refresh the cause.'
                items.append(value)
            return dict(items=items, count=len(items), total=total, offset=offset,
                        next_offset=offset+len(items) if offset+len(items)<total else None)


def retry(cfg, mail_ids, max_seconds=30):
    from . import app
    if not isinstance(mail_ids, list) or not 1 <= len(mail_ids) <= 100 or any(not isinstance(v,str) or not re.fullmatch(r'[0-9a-f]{64}',v) for v in mail_ids):
        raise ValueError('Provide 1..100 SHA256 mail IDs from list_sync_failures.')
    if type(max_seconds) not in (int,float) or not math.isfinite(max_seconds) or not 1 <= max_seconds <= 3600:
        raise ValueError('max_seconds must be 1..3600')
    requested = list(dict.fromkeys(mail_ids))
    started = time.perf_counter()
    deadline = started + max_seconds
    local = dict(cfg, _embedding_deadline=deadline)
    results, pending, missing = [], [], []
    try:
        with app.mutation_lock(cfg):
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            sql = None
            try:
                sql, table = app.open_store(cfg)
                rows = {row['mail_id']:dict(row) for row in sql.execute('SELECT * FROM sync_failures WHERE mail_id IN ('+','.join('?' for _ in requested)+')',requested)}
                missing = [mid for mid in requested if mid not in rows]
                selected = [mid for mid in requested if mid in rows]
                namespace = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI') if selected else None
                with httpx.Client(timeout=httpx.Timeout(120,connect=10),trust_env=False) as client:
                    for position, mid in enumerate(selected):
                        if time.perf_counter() >= deadline:
                            pending = selected[position:]
                            break
                        old = rows[mid]
                        metadata = dict(entry_id=old['entry_id'], modified=old['modified'], received='')
                        phase = 'outlook'
                        try:
                            item = namespace.GetItemFromID(old['entry_id'],old['store_id'])
                            metadata.update(item=item,modified=app.iso(item.LastModificationTime),received=app.iso(item.ReceivedTime))
                            mail = app.read_outlook_mail(namespace,metadata,old['store_id'],old['folder'])
                            phase = 'index'
                            app.upsert_mails([mail],sql,table,local,client)
                        except Exception as exc:
                            app.record_sync_failure(sql,metadata,old['store_id'],old['folder'],exc)
                            results.append(dict(mail_id=mid,state='failed',error_type=type(exc).__name__,error_detail=failure_reason(exc)))
                            # Do not hammer an unavailable provider or damaged local store.
                            if phase == 'index' and not app.permanent_mail_error(exc):
                                pending = selected[position+1:]
                                break
                        else:
                            with sql:
                                sql.execute('DELETE FROM sync_failures WHERE mail_id IN (?,?)',(mid,mail['id']))
                            results.append(dict(mail_id=mid,indexed_mail_id=mail['id'],state='recovered'))
            finally:
                if sql is not None:
                    sql.close()
                pythoncom.CoUninitialize()
    except Timeout:
        return dict(state='busy',items=[],pending_mail_ids=requested,note='Another sync owns this index; no retry state was changed.')
    return dict(state='incomplete' if pending else 'completed_with_errors' if any(r['state']=='failed' for r in results) else 'completed',
                items=results,recovered=sum(r['state']=='recovered' for r in results),
                failed=sum(r['state']=='failed' for r in results),pending_mail_ids=pending,
                not_found_mail_ids=missing,elapsed_ms=round((time.perf_counter()-started)*1000))


def folder_progress(sql, cfg):
    counts = {r['folder']:r['n'] for r in sql.execute('SELECT folder,count(*) n FROM emails GROUP BY folder')}
    failed = {r['folder']:r['n'] for r in sql.execute('SELECT folder,count(*) n FROM sync_failures GROUP BY folder')}
    latest = {}
    for row in sql.execute("SELECT value FROM state WHERE key LIKE 'sync:%'"):
        try:
            value = json.loads(row[0])
        except (ValueError,TypeError):
            continue
        if not isinstance(value,dict) or not value.get('folder'):
            continue
        name = value['folder']
        timestamp = value.get('progress_updated_at') or value.get('last_run') or ''
        previous = latest.get(name,{})
        if timestamp >= (previous.get('progress_updated_at') or previous.get('last_run') or ''):
            latest[name] = value
    inventory = sql.execute("SELECT value FROM state WHERE key='folder_inventory'").fetchone()
    try:
        inventory = json.loads(inventory[0]) if inventory else {}
    except (ValueError,TypeError):
        inventory = {}
    totals = {f['folder']:f.get('source_items_estimate') for f in inventory.get('folders',[])}
    names = set(counts) | set(failed) | set(latest) | set(totals)
    if cfg.get('folders','auto') != 'auto':
        names &= set(cfg['folders'])
        names |= set(cfg['folders'])
    output = []
    for name in sorted(names):
        value = latest.get(name,{})
        complete = bool(value.get('complete'))
        known = bool(value.get('progress_known',False))
        processed = value.get('processed_rows',0) if known else None
        total = value.get('source_items_estimate') if value else totals.get(name)
        seconds = value.get('processing_seconds',0)
        rate = processed/seconds if processed is not None and seconds > 0 and math.isfinite(seconds) else None
        remaining = 0 if complete else max(0,total-processed) if total is not None and processed is not None else total if not value else None
        eta = 0.0 if complete else remaining/rate if remaining is not None and rate else None
        percent = 100.0 if complete else min(100.0,100*processed/total) if processed is not None and total else None
        output.append(dict(folder=name,since_days=value.get('since_days',inventory.get('since_days')),
            state='scan_complete' if complete else 'scanning' if value else 'not_started',
            indexed_emails=counts.get(name,0),failed_emails=failed.get(name,0),
            source_items_estimate=total,processed_items=processed,remaining_items_estimate=remaining,
            scan_progress_percent=round(percent,2) if percent is not None else None,
            items_per_second=round(rate,3) if rate is not None else None,
            eta_seconds_estimate=round(eta,1) if eta is not None else None,
            updated_at=value.get('progress_updated_at') or value.get('last_run'),
            estimate_basis='Committed metadata rows; source Items.Count may include non-mail. Completion means scan complete, not all failures recovered.'))
    return output


def progress(cfg):
    from . import app
    with app._LOCK, stored_sql(cfg) as sql:
        return dict(folders=folder_progress(sql,cfg) if sql is not None else [],source='saved_sync_checkpoints',live_outlook_access=False)


def thread(cfg, mail_id, before=3, after=3, max_body_chars=2000):
    from . import app
    if not isinstance(mail_id,str) or not re.fullmatch(r'[0-9a-f]{64}',mail_id):
        raise ValueError('Invalid mail_id')
    if any(type(v) is not int or not 0 <= v <= 50 for v in (before,after)) or type(max_body_chars) is not int or not 0 <= max_body_chars <= 10000:
        raise ValueError('before/after must be 0..50 and max_body_chars 0..10000')
    with app._LOCK, stored_sql(cfg) as sql:
        if sql is None:
            raise ValueError('Indexed mail not found.')
        if sql is not None:
            columns = "id,entry_id,store_id,folder,subject,sender,recipients,received,modified,conversation_id,substr(body,1,?) AS body,length(body) AS body_length"
            anchor = sql.execute(f'SELECT {columns} FROM emails WHERE id=?',(max_body_chars,mail_id)).fetchone()
            if anchor is None:
                raise ValueError('Indexed mail not found.')
            if anchor['conversation_id']:
                args = [anchor['store_id'],anchor['conversation_id'],anchor['received'],anchor['received'],mail_id]
                previous_where = 'store_id=? AND conversation_id=? AND (received<? OR (received=? AND id<?))'
                following_where = 'store_id=? AND conversation_id=? AND (received>? OR (received=? AND id>?))'
                previous = sql.execute(f'SELECT {columns} FROM emails WHERE {previous_where} ORDER BY received DESC,id DESC LIMIT ?', [max_body_chars,*args,before]).fetchall()
                following = sql.execute(f'SELECT {columns} FROM emails WHERE {following_where} ORDER BY received,id LIMIT ?', [max_body_chars,*args,after]).fetchall()
                previous_total = sql.execute('SELECT count(*) FROM emails WHERE '+previous_where,args).fetchone()[0]
                following_total = sql.execute('SELECT count(*) FROM emails WHERE '+following_where,args).fetchone()[0]
            else:
                previous, following, previous_total, following_total = [], [], 0, 0
            # Reserve space for the selected mail before spending the context budget.
            focus_budget = min(max_body_chars,anchor['body_length'])
            budget, items = 60000-focus_budget, []
            for row in [*reversed(previous),anchor,*following]:
                value = dict(row)
                keep = focus_budget if row['id']==mail_id else min(max_body_chars,budget)
                value.update(body=value['body'][:keep],body_truncated=value.pop('body_length')>keep,
                             position='focus' if row['id']==mail_id else 'previous' if (row['received'],row['id'])<(anchor['received'],mail_id) else 'next')
                if row['id']!=mail_id:
                    budget -= len(value['body'])
                value['mail_id'] = value.pop('id')
                items.append(value)
            return dict(mail_id=mail_id,conversation_id=anchor['conversation_id'],store_id=anchor['store_id'],
                        items=items,count=len(items),thread_mail_count=previous_total+1+following_total,source='local_index',
                        missing_conversation_id=not bool(anchor['conversation_id']),total_body_limit=60000,
                        has_more_previous=previous_total>len(previous),has_more_next=following_total>len(following))
