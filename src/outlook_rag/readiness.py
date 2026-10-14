"""Read-only first-use checks; never create files or load the vector engine."""
from __future__ import annotations

from contextlib import closing
import json
from pathlib import Path
import sqlite3


def metadata_initialized(sql: sqlite3.Connection) -> bool:
    tables = {row[0] for row in sql.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    identity = sql.execute("SELECT value FROM state WHERE key='embedding_identity'").fetchone() if "state" in tables else None
    if identity and not isinstance(json.loads(identity[0]), dict):
        raise ValueError("Invalid cached embedding identity; repair the selected index.")
    if {"state", "emails", "chunks"} <= tables and identity:
        return True
    if "emails" in tables and sql.execute("SELECT 1 FROM emails LIMIT 1").fetchone():
        raise ValueError("Populated index metadata is incomplete; repair or restore the selected index.")
    return False


def inspect(cfg: dict) -> dict:
    path = Path(cfg["data_dir"]) / "metadata.sqlite"
    state = "not_initialized"
    if path.is_file():
        with closing(sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True, timeout=1)) as sql:
            if not metadata_initialized(sql):
                state = "initializing"
            elif sql.execute("SELECT 1 FROM emails LIMIT 1").fetchone():
                return {"state": "ready"}
            else:
                state = "empty" if (path.parent / "vectors").is_dir() else "initializing"
    note = ("Index is being initialized; inspect sync_index(action='status') and retry after initialization."
            if state == "initializing" else "Index empty: use sync_index(action='start') with the embedding API available first.")
    return {"state": state, "emails": 0, "chunks": 0, "note": note, "data_dir": cfg["data_dir"]}
