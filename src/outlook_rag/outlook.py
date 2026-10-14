"""Outlook read-only COM adapters and stable metadata identities."""

from __future__ import annotations

from datetime import datetime, timezone
import fnmatch
import hashlib
import json
import sqlite3
import sys
import time
from typing import Any, Iterator
import uuid


def com_attr(item: Any, name: str) -> Any:
    """Read a required COM property in at most six attempts on its owning thread.

    Dynamic pywin32 dispatch can turn a failed GetIDsOfNames lookup into an
    AttributeError. Never use safe_attr's empty default for persistent IDs.
    """
    errors: tuple[type[Exception], ...] = (AttributeError,)
    try:
        from pywintypes import com_error
    except ImportError:
        pass  # Keep offline metadata readers usable without Windows extensions.
    else:
        errors += (com_error,)
    for _ in range(5):
        try:
            return getattr(item, name)
        except errors:
            time.sleep(0.2)
    # The final attempt propagates the original exception and traceback.
    return getattr(item, name)


def safe_attr(item: Any, name: str, default: Any = "") -> Any:
    try:
        return getattr(item, name)
    except Exception:
        return default


RECEIVED = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
MODIFIED = "http://schemas.microsoft.com/mapi/proptag/0x30080040"


def iso(value: Any, *, utc: bool = False) -> str:
    if isinstance(value, datetime):
        # COM DATE tzinfo is not evidence of a timezone. The property namespace
        # determines UTC/local semantics. astimezone uses the OS rules for this date.
        value = value.replace(tzinfo=timezone.utc if utc else None).astimezone()
        return value.isoformat()
    return str(value or "")


def item_date(item: Any, name: str, property_name: str) -> str:
    try:
        value = item.PropertyAccessor.GetProperty(property_name)
        if isinstance(value, datetime):
            return iso(value, utc=True)
    except Exception:
        pass
    return iso(safe_attr(item, name))


def date_literal(value: datetime) -> str:
    """Jet dates use the Windows user's short-date/time formats, without seconds."""
    value = value.astimezone().replace(tzinfo=None) if value.tzinfo is not None else value
    if sys.platform != "win32":
        return value.strftime("%x %H:%M")  # Offline test hosts have no Outlook.
    import ctypes
    from ctypes import wintypes
    class SystemTime(ctypes.Structure):
        _fields_ = [(name, wintypes.WORD) for name in ("year", "month", "weekday", "day", "hour", "minute", "second", "milliseconds")]
    instant = SystemTime(value.year, value.month, 0, value.day, value.hour, value.minute, 0, 0)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    parts = []
    for function, flags, arguments in ((kernel.GetDateFormatEx, 1, (None,)), (kernel.GetTimeFormatEx, 2, ())):
        output = ctypes.create_unicode_buffer(256)
        function.restype = ctypes.c_int
        if not function(None, flags, ctypes.byref(instant), None, output, len(output), *arguments):
            raise OSError("Cannot format Outlook date using Windows regional settings")
        parts.append(output.value)
    return " ".join(parts).replace("'", "''")


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
        json.dumps([str(com_attr(folder, "StoreID")), str(com_attr(folder, "EntryID"))]).encode()
    ).hexdigest()


def discover_sources(namespace: Any, cfg: dict) -> tuple[dict, list]:
    stores, targets, errors = [], [], []
    for store in namespace.Stores:
        name = str(safe_attr(store, "DisplayName", "Outlook store"))
        excluded = set()
        if cfg.get("exclude_system_folders", True):
            for kind in (3, 20, 23):  # Deleted Items, Sync Issues and Junk Email.
                try:
                    excluded.add(str(com_attr(store.GetDefaultFolder(kind), "EntryID")))
                except Exception:
                    pass
        try:
            excluded.update(str(com_attr(folder, "EntryID")) for folder in store.GetSearchFolders())
        except Exception:
            pass
        file_path = str(safe_attr(store, "FilePath", ""))
        store_info: dict[str, Any] = {
            "name": name,
            "store_id": str(safe_attr(store, "StoreID", "")),
            "file_path": file_path,
            "is_pst": file_path.casefold().endswith(".pst"),
            "folders": [],
        }
        stores.append(store_info)
        try:
            root = store.GetRootFolder()
            store_info["store_id"] = store_info["store_id"] or str(com_attr(root, "StoreID"))
            stack = [(root, name, 0)]
            visited = set()
            while stack:
                folder, path, depth = stack.pop()
                entry = str(com_attr(folder, "EntryID"))
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
                item_count = safe_attr(safe_attr(folder, "Items", None), "Count", None)
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


def _metadata_rows(values: Any) -> tuple[tuple[Any, ...], ...]:
    """Normalize Python COM arrays without guessing the direction of a 4x4 batch."""
    error = "Unexpected Outlook metadata table shape or values"
    if values is None or isinstance(values, (tuple, list)) and not values:
        return ()
    if not isinstance(values, (tuple, list)) or any(
        not isinstance(group, (tuple, list)) for group in values
    ):
        raise ValueError(error)
    groups = tuple(tuple(group) for group in values)
    if len({len(group) for group in groups}) != 1:
        raise ValueError(error)
    if len(groups) == 4 and not groups[0]:
        return ()

    def date_value(value: Any) -> bool:
        if value is None or isinstance(value, datetime):
            return True
        if isinstance(value, str):
            if not value:
                return True
            try:
                datetime.fromisoformat(value)
                return True
            except ValueError:
                pass
        return False

    candidates = []
    if len(groups[0]) == 4:
        candidates.append(groups)
    if len(groups) == 4:
        candidates.append(tuple(zip(*groups)))
    matches = [rows for rows in candidates if all(
        isinstance(row[0], str) and bool(row[0]) and isinstance(row[1], str)
        and date_value(row[2]) and date_value(row[3]) for row in rows
    )]
    # Validate the entire batch before yielding any IDs or updating sync cursors.
    # If both orientations match, refuse it instead of silently mixing fields.
    if len(matches) != 1:
        raise ValueError(error)
    return matches[0]


def mail_metadata(folder: Any, filters: list[str]) -> Iterator[tuple[dict, str]]:
    """Fetch lightweight table rows. Use Items only when GetTable is unavailable."""
    expression = " AND ".join(filters)
    try:
        table = folder.GetTable(expression)
        table.Columns.RemoveAll()
        for name in ("EntryID", "MessageClass", RECEIVED, MODIFIED):
            table.Columns.Add(name)
        table.Sort("ReceivedTime", True)
    except Exception:
        # Some providers do not implement GetTable or its mail-property columns.
        table = None
    if table is not None:
        while not table.EndOfTable:
            # Real Outlook/pywin32 returns row-major tuples on reported hosts;
            # also accept column-major arrays without relying on VBA dimensions.
            rows = _metadata_rows(table.GetArray(64))
            if not rows:
                break
            for row in rows:
                yield (
                    {
                        "entry_id": str(row[0]),
                        "cursor_id": str(row[0]),
                        "mail": str(row[1]).startswith("IPM.Note"),
                        "received": iso(row[2], utc=True),
                        "modified": iso(row[3], utc=True),
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
                    "received": item_date(item, "ReceivedTime", RECEIVED),
                    "modified": item_date(item, "LastModificationTime", MODIFIED),
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
    from .mail_metadata import read
    result = dict(
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
        **read(item),
    )
    cfg = metadata.get("_attachment_cfg")
    if cfg and cfg.get("index_attachments"):
        from .attachments import read as read_attachments
        try:
            result["attachments"], result["attachment_errors"] = read_attachments(item, cfg)
        except Exception as exc:
            # Extraction/capture errors must never erase an otherwise readable body.
            result["attachments"] = []
            result["attachment_errors"] = [dict(error_type=type(exc).__name__)]
    return result
