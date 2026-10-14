"""Read-only sync scope previews and explicitly qualified time estimates."""
from __future__ import annotations

from datetime import datetime, timedelta
import math
import time
from typing import Any

from . import coverage, features, outlook


def validate(folders: list[str] | None, since_days: int | None) -> None:
    if folders is not None:
        coverage.folder_selection(folders)
    if since_days is not None and (type(since_days) is not int or not 0 <= since_days <= 36500):
        raise ValueError("since_days must be 0 (all time) or 1..36500")


def estimate(rows: list[dict], *, total: int | None = None, unknown: int | None = None) -> dict:
    """Estimate scanning time; never turn missing counts/rates into a zero ETA."""
    measured: list[tuple[float, float]] = []
    for row in rows:
        n, s = row.get("processed_items"), row.get("processing_seconds")
        if isinstance(n, (int, float)) and isinstance(s, (int, float)) and not isinstance(n, bool) and not isinstance(s, bool) and n > 0 and s > 0 and math.isfinite(n) and math.isfinite(s):
            measured.append((float(n), float(s)))
    processed = sum(n for n, _ in measured)
    seconds = sum(s for _, s in measured)
    rate = processed / seconds if processed >= 20 and seconds >= 1 else None
    if total is None:
        amounts = [r.get("remaining_items_estimate") for r in rows]
        unknown = (sum(n is None for n in amounts) if rows else 1) if unknown is None else unknown
        remaining = sum(float(n) for n in amounts if isinstance(n, (int, float)))
    else:
        remaining = float(total)
    unknown = unknown or 0
    point = remaining / rate if rate and not unknown else None
    low = remaining / (rate * 2 if rate else 10) if not unknown else None
    high = remaining / (rate / 2 if rate else 1) if not unknown else None
    if remaining == 0 and not unknown:
        point = low = high = 0.0
    return dict(remaining_items_estimate=int(remaining) if not unknown else None,
                known_remaining_items=int(remaining), unknown_folder_count=unknown,
                items_per_second=round(rate, 3) if rate else None,
                estimated_seconds=round(point, 1) if point is not None else None,
                range_seconds=[round(low, 1), round(high, 1)] if low is not None and high is not None else None,
                basis="same_index_measured_scan_rate" if rate else "assumed_1_to_10_items_per_second",
                confidence="approximate" if rate and not unknown else "low",
                note="Scan time estimate, not a deadline or guaranteed completion. Initial range assumes 1..10 items/s; API speed, mail size, attachments and cached mail change throughput. Unknown counts leave ETA unavailable. Scan completion does not recover retained failures.")


def count_window(folder: Any, days: int, now: datetime) -> int:
    expression = f"[ReceivedTime] >= '{outlook.date_literal(now - timedelta(days=days))}'"
    table = items = None
    try:
        try:
            table = folder.GetTable(expression)
            value = table.GetRowCount()
        except Exception:
            table = None
            items = folder.Items.Restrict(expression)
            value = items.Count
        if type(value) is not int or value < 0:
            raise ValueError("Outlook returned an invalid row count")
        return value
    finally:
        table = items = None


def plan(backend: Any, folders: list[str] | None = None, since_days: int | None = None) -> dict:
    """Only metadata counts: no mail bodies, embeddings, credentials or writes."""
    validate(folders, since_days)
    cfg = backend.local_config()
    selection = coverage.folder_selection(folders if folders is not None else cfg.get("folders", "auto"))
    rates = features.progress(dict(cfg, folders="auto"))["folders"]
    days_list = [since_days] if since_days is not None else [30, 90, 365, 0]
    now = datetime.now().astimezone()
    deadline = time.perf_counter() + 15
    import pythoncom
    import win32com.client
    namespace = targets = folder = None
    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        report, discovered = backend.discover_sources(namespace, cfg)
        counts = {f["path"]: f.get("item_count") for s in report["stores"] for f in s["folders"]}
        targets = discovered if selection == "auto" else [(name, backend.resolve_folder(namespace, name)) for name in selection]
        previews = []
        for days in days_list:
            rows: list[dict[str, Any]] = []
            for name, folder in targets:
                error = None
                if days == 0:
                    count = counts.get(name)
                    if selection != "auto":
                        count = outlook.safe_attr(folder.Items, "Count", None)
                elif time.perf_counter() >= deadline:
                    count, error = None, "planning_time_budget"
                else:
                    try:
                        count = count_window(folder, days, now)
                    except Exception as exc:
                        count, error = None, type(exc).__name__
                count = count if type(count) is int and count >= 0 else None
                rows.append(dict(folder=name, source_items_estimate=count, error_type=error))
            unknown = sum(r["source_items_estimate"] is None for r in rows) + len(report["errors"])
            total = sum(int(r["source_items_estimate"]) for r in rows if r["source_items_estimate"] is not None)
            previews.append(dict(since_days=days, period="all_time" if days == 0 else f"last_{days}_days",
                                 candidate_items=total if not unknown else None, known_candidate_items=total,
                                 folder_count=len(rows), folder_counts=rows,
                                 time_estimate=estimate(rates, total=total, unknown=unknown)))
        return dict(state="confirmation_required", started=False, selection=selection,
                    model=cfg["model"], previews=previews, recommended_since_days=30 if since_days is None else since_days,
                    count_basis="Outlook metadata rows; includes indexed and non-mail items. Unchanged cached mail may scan faster.",
                    discovery_errors=report["errors"], live_outlook_access=True, embedding_api_called=False,
                    next_action="Show scope, candidate counts and estimated time to the user and ask which period to index. Do not select a period yourself. Only after the user chooses/approves a scope, call sync_index(action='start', since_days=<chosen>, confirmed=true). An explicit 'all/everything/全部/全期間' request already approves folders=['auto'], since_days=0; no second confirmation is needed.")
    finally:
        folder = targets = namespace = discovered = None
        pythoncom.CoUninitialize()
