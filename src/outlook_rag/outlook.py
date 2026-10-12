"""Outlook read-only COM adapters and stable metadata identities."""

from __future__ import annotations

from datetime import datetime
import fnmatch
import hashlib
import json
import sqlite3
import time
from typing import Any, Iterator
import uuid


def safe_attr(item: Any, name: str, default: Any = "") -> Any:
    try:
        return getattr(item, name)
    except Exception:
        return default


def iso(value: Any) -> str:
    return value.isoformat() if hasattr(value, "isoformat") else str(value or "")


def same_modified(left: str, right: str) -> bool:
    if left == right:
        return True
    try:
        return (
            datetime.fromisoformat(left).timestamp()
            == datetime.fromisoformat(right).timestamp()
        )
    except (ValueError, TypeError):
        return False


def resolve_folder(namespace: Any, name: str) -> Any:
    known = {"inbox": 6, "sent": 5, "drafts": 16, "deleted": 3, "junk": 23}
    if name.lower() in known:
        return namespace.GetDefaultFolder(known[name.lower()])
    segments = name.replace("\\", "/").split("/")
    root = namespace.GetDefaultFolder(6).Parent
    for store in namespace.Stores:
        if store.DisplayName.lower() == segments[0].lower():
            root = store.GetRootFolder()
            segments = segments[1:]
            if len(segments) == 1 and segments[0].lower() in known:
                return store.GetDefaultFolder(known[segments[0].lower()])
            break
    for part in segments:
        found = None
        for child in root.Folders:
            if child.Name.lower() == part.lower():
                found = child
                break
        if found is None:
            raise ValueError(f"Outlook folder not found: {name}")
        root = found
    return root


def folder_identity(folder: Any) -> str:
    return hashlib.sha256(
        json.dumps([str(folder.StoreID), str(folder.EntryID)]).encode()
    ).hexdigest()


def discover_sources(namespace: Any, cfg: dict) -> tuple[dict, list]:
    stores, targets, errors = [], [], []
    for store in namespace.Stores:
        name = str(safe_attr(store, "DisplayName", "Outlook store"))
        excluded = set()
        if cfg.get("exclude_system_folders", True):
            for kind in (3, 20, 23):  # Deleted Items, Sync Issues and Junk Email.
                try:
                    excluded.add(str(store.GetDefaultFolder(kind).EntryID))
                except Exception:
                    pass
        try:
            excluded.update(str(folder.EntryID) for folder in store.GetSearchFolders())
        except Exception:
            pass
        store_info: dict[str, Any] = {
            "name": name,
            "file_path": str(safe_attr(store, "FilePath", "")),
            "folders": [],
        }
        stores.append(store_info)
        try:
            stack = [(store.GetRootFolder(), name, 0)]
            visited = set()
            while stack:
                folder, path, depth = stack.pop()
                entry = str(folder.EntryID)
                if entry in visited:
                    continue
                visited.add(entry)
                if depth > 64 or len(visited) > 5000:
                    raise ValueError("Outlook folder hierarchy exceeds traversal limit")
                system_name = str(safe_attr(folder, "Name", "")).casefold() in {
                    "junk e-mail",
                    "junk",
                    "spam",
                    "trash",
                    "deleted items",
                    "迷惑メール",
                    "ゴミ箱",
                    "削除済みアイテム",
                }
                excluded_path = (
                    entry in excluded
                    or (cfg.get("exclude_system_folders", True) and system_name)
                    or any(
                        fnmatch.fnmatchcase(path.casefold(), pattern.casefold())
                        for pattern in cfg.get("excluded_folders", [])
                    )
                )
                item_count = safe_attr(safe_attr(folder, "Items", None), "Count", 0)
                mail_folder = safe_attr(folder, "DefaultItemType", -1) == 0
                store_info["folders"].append(
                    {
                        "path": path,
                        "item_count": item_count,
                        "mail_folder": mail_folder,
                        "included": mail_folder and not excluded_path,
                    }
                )
                if excluded_path:
                    continue
                if mail_folder:
                    targets.append((path, folder))
                for child in folder.Folders:
                    stack.append((child, path + "/" + str(child.Name), depth + 1))
        except Exception as exc:
            errors.append({"store": name, "error": type(exc).__name__})
    return {"stores": stores, "errors": errors, "target_folders": len(targets)}, targets


def mail_metadata(folder: Any, filters: list[str]) -> Iterator[tuple[dict, str]]:
    """Fetch lightweight table rows. Use Items only when GetTable is unavailable."""
    expression = " AND ".join(filters)
    try:
        table = folder.GetTable(expression)
        table.Columns.RemoveAll()
        for name in ("EntryID", "MessageClass", "ReceivedTime", "LastModificationTime"):
            table.Columns.Add(name)
        table.Sort("ReceivedTime", True)
    except Exception:
        # Some providers do not implement GetTable or its mail-property columns.
        table = None
    if table is not None:
        while not table.EndOfTable:
            columns = table.GetArray(64)
            if not columns:
                break
            # Outlook SAFEARRAY dimensions are (column, row), preserved by
            # pywin32. Validate before transposing, including square batches.
            if len(columns) != 4 or len({len(column) for column in columns}) != 1:
                raise ValueError("Unexpected Outlook metadata table shape")
            if not columns[0]:
                break
            for row in zip(*columns):
                yield (
                    {
                        "entry_id": str(row[0]),
                        "cursor_id": str(row[0]),
                        "mail": str(row[1]).startswith("IPM.Note"),
                        "received": iso(row[2]),
                        "modified": iso(row[3]),
                        "item": None,
                    },
                    "GetTable",
                )
    else:
        items = folder.Items.Restrict(expression) if filters else folder.Items
        items.Sort("[ReceivedTime]", True)
        for item in items:
            yield (
                {
                    "entry_id": str(item.EntryID),
                    "mail": safe_attr(item, "Class", 0) == 43,
                    "received": iso(safe_attr(item, "ReceivedTime")),
                    "modified": iso(safe_attr(item, "LastModificationTime")),
                    "item": item,
                },
                "Items",
            )


def outlook_session_key(namespace: Any) -> str:
    # Short-term GetTable IDs are not portable across Outlook sessions.
    # Scope aliases to the profile and Outlook process creation times.
    try:
        profile = namespace.CurrentProfileName
        import win32com.client

        processes = win32com.client.GetObject("winmgmts:").ExecQuery(
            "SELECT ProcessId,CreationDate FROM Win32_Process WHERE Name='OUTLOOK.EXE'"
        )
        identities = sorted(
            f"{process.ProcessId}:{process.CreationDate}" for process in processes
        )
        if identities:
            return hashlib.sha256(
                json.dumps([profile, identities]).encode()
            ).hexdigest()
    except Exception:
        pass
    # If session identity cannot be verified, do not reuse stored aliases.
    return str(uuid.uuid4())


def canonical_metadata(
    namespace: Any, metadata: dict, store_id: str, sql: sqlite3.Connection, session: str
) -> dict:
    if metadata.get("item") is not None:
        return metadata
    short_id = metadata["entry_id"]
    key = (session, store_id, short_id, metadata["received"])
    row = sql.execute(
        "SELECT canonical_id,last_used FROM entry_aliases WHERE session=? AND store_id=? AND short_id=? AND received=?",
        key,
    ).fetchone()
    if row:
        if row[1] < time.time() - 86400:
            sql.execute(
                "UPDATE entry_aliases SET last_used=? WHERE session=? AND store_id=? AND short_id=? AND received=?",
                (time.time(), *key),
            )
        return dict(metadata, entry_id=row[0])
    item = namespace.GetItemFromID(short_id, store_id)
    canonical_id = str(item.EntryID)
    sql.execute(
        "INSERT OR REPLACE INTO entry_aliases VALUES (?,?,?,?,?,?)",
        (*key, canonical_id, time.time()),
    )
    return dict(metadata, entry_id=canonical_id, item=item)


def read_outlook_mail(
    namespace: Any, metadata: dict, store_id: str, folder_name: str
) -> dict:
    item = metadata.get("item")
    if item is None:
        item = namespace.GetItemFromID(metadata["entry_id"], store_id)
    entry_id = str(item.EntryID)
    # Fail explicitly if required text cannot be read; never overwrite it with empty text.
    subject, body = str(item.Subject), str(item.Body)
    sender = str(safe_attr(item, "SenderEmailAddress"))
    try:
        sender = (
            item.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
            )
            or sender
        )
    except Exception:
        pass
    return dict(
        id=hashlib.sha256(f"{store_id}:{entry_id}".encode()).hexdigest(),
        entry_id=entry_id,
        store_id=store_id,
        folder=folder_name,
        subject=subject,
        body=body,
        sender=sender,
        recipients=str(safe_attr(item, "To")),
        received=metadata["received"],
        modified=metadata["modified"],
        conversation_id=str(safe_attr(item, "ConversationID")),
    )
