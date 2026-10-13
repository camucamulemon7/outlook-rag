"""Extended message attributes, compatible with the original email row schema."""
from __future__ import annotations
import sqlite3
import re
from typing import Any

FIELDS = ("sender_name", "cc", "internet_message_id", "has_attachments", "categories", "importance")


def initialize(sql: sqlite3.Connection) -> None:
    sql.executescript("""
    CREATE TABLE IF NOT EXISTS mail_metadata(email_id TEXT PRIMARY KEY,
      sender_name TEXT, cc TEXT, internet_message_id TEXT, has_attachments INTEGER,
      categories TEXT, importance INTEGER);
    CREATE INDEX IF NOT EXISTS mail_metadata_message_id ON mail_metadata(internet_message_id);
    CREATE TRIGGER IF NOT EXISTS remove_mail_metadata AFTER DELETE ON emails
      BEGIN DELETE FROM mail_metadata WHERE email_id=OLD.id; END;
    INSERT OR REPLACE INTO state VALUES ('mail_metadata_schema','1');
    """)


def available(sql: sqlite3.Connection) -> bool:
    return bool(sql.execute("SELECT 1 FROM sqlite_master WHERE name='mail_metadata'").fetchone())


def read(item: Any) -> dict:
    from .outlook import safe_attr
    message_id = ""
    for tag in ("0x1035001F", "0x1035001E"):
        try:
            message_id = str(item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/" + tag) or "").strip()
            if re.fullmatch(r"<[^<>\s]+@[^<>\s]+>", message_id):
                break
            message_id = ""
        except Exception:
            pass
    return dict(sender_name=str(safe_attr(item, "SenderName")), cc=str(safe_attr(item, "CC")),
                internet_message_id=message_id, has_attachments=bool(safe_attr(safe_attr(item, "Attachments", None), "Count", 0)),
                categories=str(safe_attr(item, "Categories")), importance=safe_attr(item, "Importance", 1))


def save(sql: sqlite3.Connection, mail: dict) -> None:
    sql.execute("INSERT OR REPLACE INTO mail_metadata VALUES (?,?,?,?,?,?,?)",
                [mail["id"], *[mail.get(field, 0 if field == "has_attachments" else "") for field in FIELDS]])


def lookup(sql: sqlite3.Connection, ids: list[str]) -> dict[str, dict]:
    if not ids or not available(sql):
        return {}
    result = {}
    for offset in range(0, len(ids), 500):
        batch = ids[offset:offset + 500]
        for row in sql.execute("SELECT * FROM mail_metadata WHERE email_id IN (" + ",".join("?" for _ in batch) + ")", batch):
            values = dict(row)
            result[values.pop("email_id")] = values
    return result
