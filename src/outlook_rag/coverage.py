"""Search scope and archive coverage without live Outlook calls on search."""

from __future__ import annotations

import json
import sqlite3
from typing import Any

from .outlook import com_attr


def folder_selection(value: Any) -> str | list[str]:
    """MCP accepts list[str]; ['auto'] explicitly selects every connected store."""
    if value == "auto" or value == ["auto"]:
        return "auto"
    if not isinstance(value, list) or not all(isinstance(path, str) for path in value):
        raise ValueError("folders must be 'auto', ['auto'], or a list of folder paths")
    if "auto" in value:
        raise ValueError(
            "Use ['auto'] alone to include all connected stores and PST archives"
        )
    return value


def state_key(folder: Any, since_days: int = 0) -> str:
    return f"sync:{com_attr(folder, 'StoreID')}:{com_attr(folder, 'EntryID')}:{since_days}"


def inventory(
    discovery: dict | None, targets: list, selection: Any, since_days: int
) -> dict:
    """Retain the full discovered scope even when a bounded sync stops early."""
    stores = discovery.get("stores", []) if discovery else []
    counts = {
        info["path"]: info.get("item_count")
        for store in stores
        for info in store.get("folders", [])
    }
    return dict(
        selection=selection,
        since_days=since_days,
        stores=[
            dict(
                store_id=store.get("store_id", ""),
                name=store.get("name", ""),
                is_pst=store.get("is_pst", False),
            )
            for store in stores
        ],
        discovery_errors=discovery.get("errors", []) if discovery else [],
        folders=[
            dict(
                folder=name,
                store_id=str(com_attr(folder, "StoreID")),
                state_key=state_key(folder, since_days),
                source_items_estimate=counts.get(name) if since_days == 0 else None,
            )
            for name, folder in targets
        ],
    )


def completed(sql: sqlite3.Connection | None, folders: list[dict]) -> set[str]:
    if sql is None:
        return set()
    keys = [folder["state_key"] for folder in folders if folder.get("state_key")]
    done = set()
    for offset in range(0, len(keys), 500):
        batch = keys[offset : offset + 500]
        for key, raw in sql.execute(
            "SELECT key,value FROM state WHERE key IN ("
            + ",".join("?" for _ in batch)
            + ")",
            batch,
        ):
            try:
                checkpoint = json.loads(raw)
            except (ValueError, TypeError):
                continue
            if (
                isinstance(checkpoint, dict)
                and checkpoint.get("complete") is True
                and not checkpoint.get("reconcile_uncertain")
            ):
                done.add(key)
    return done


def index_coverage(sql: sqlite3.Connection | None, cfg: dict) -> dict:
    """Describe the last discovered scope, never claim current live completeness."""
    row = (
        sql.execute("SELECT value FROM state WHERE key='folder_inventory'").fetchone()
        if sql is not None
        else None
    )
    try:
        saved = json.loads(row[0]) if row else {}
        if not isinstance(saved, dict):
            saved = {}
    except (ValueError, TypeError):
        saved = {}
    folders = saved.get("folders", [])
    folders = (
        [f for f in folders if isinstance(f, dict) and f.get("folder")]
        if isinstance(folders, list)
        else []
    )
    done = completed(sql, folders)
    pending = [f["folder"] for f in folders if f.get("state_key") not in done]
    failed = bool(
        sql is not None
        and sql.execute("SELECT 1 FROM sync_failures LIMIT 1").fetchone()
    )
    known = bool(
        saved.get("selection") == "auto" and all(f.get("state_key") for f in folders)
    )
    complete = bool(
        known
        and not pending
        and not failed
        and not saved.get("discovery_errors")
        and saved.get("since_days") == 0
    )
    stores = saved.get("stores", [])
    pst = (
        [s.get("name", "") for s in stores if isinstance(s, dict) and s.get("is_pst")]
        if isinstance(stores, list)
        else []
    )
    return dict(
        basis="last_sync_inventory" if known else "unknown",
        live_sources_checked=False,
        known_inventory_complete=complete if known else None,
        known_pst_stores=pst[:20],
        known_pst_store_count=len(pst),
        configured_sync_folders=cfg.get("folders", "auto"),
        last_sync_folders=saved.get("selection"),
        last_sync_since_days=saved.get("since_days"),
        pending_folder_count=len(pending),
        pending_folders=pending[:20],
        retained_failures=failed,
        next_action="Verify connected stores with list_outlook_sources. To index all connected PST archives, use sync_index(action='start') with folders=['auto'], since_days=0.",
    )


def search_result(
    result: dict,
    sql: sqlite3.Connection | None,
    cfg: dict,
    folder: str | None,
    since: str | None,
    until: str | None,
) -> dict:
    result["search_scope"] = dict(
        stores="all_indexed_stores" if not folder else "selected_folder",
        folder=folder,
        since=since,
        until=until,
        archives="All indexed PST/archive stores are included unless explicitly restricted by filters.",
    )
    result["coverage"] = index_coverage(sql, cfg)
    if not result["count"]:
        result["absence_confirmed"] = False
        result.setdefault(
            "note",
            "No match in the indexed scope. This does not prove that mail is absent from Outlook or unindexed PST archives; check source coverage before answering.",
        )
    return result


def live_coverage(
    report: dict, targets: list, sql: sqlite3.Connection | None, cfg: dict
) -> dict:
    """Attach offline scan readiness to the stores actually observed by Outlook."""
    folders = [
        dict(folder=name, state_key=state_key(folder)) for name, folder in targets
    ]
    done = completed(sql, folders)
    pending = [f["folder"] for f in folders if f["state_key"] not in done]
    failed = bool(
        sql is not None
        and sql.execute("SELECT 1 FROM sync_failures LIMIT 1").fetchone()
    )
    counts = (
        {
            r[0]: r[1]
            for r in sql.execute(
                "SELECT store_id,count(*) FROM emails GROUP BY store_id"
            )
        }
        if sql is not None
        else {}
    )
    for store in report["stores"]:
        store["indexed_emails"] = counts.get(store.get("store_id", ""), 0)
    report["coverage"] = dict(
        basis="connected_outlook_stores",
        live_sources_checked=True,
        all_date_scan_complete=not pending and not failed and not report["errors"],
        pending_folder_count=len(pending),
        pending_folders=pending[:20],
        configured_sync_folders=cfg.get("folders", "auto"),
        retained_failures=failed,
    )
    report["search_guidance"] = (
        "For ordinary mail lookups, use search_emails without folder/since/until unless the user explicitly restricts them. Include PST archives. If folders remain unindexed, use folders=['auto'], since_days=0 for sync; an empty result is not proof of absence."
    )
    return report
