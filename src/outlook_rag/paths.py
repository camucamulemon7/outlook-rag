"""Stable index placement; Outlook is consulted only before a new sync."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import sys
from typing import Any

from filelock import FileLock


def local_root() -> Path:
    return Path(os.environ["LOCALAPPDATA"]) if os.environ.get("LOCALAPPDATA") else Path.home() / "AppData" / "Local"


def default_root() -> Path:
    return Path.home() / "Documents" / "Outlookファイル" / "outlook-rag"


def remembered_root(path: Path) -> Path | None:
    if not path.exists():
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        root = Path(value["index_root"])
        if value["version"] != 1 or not root.is_absolute():
            raise ValueError("Invalid location")
        return root
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise ValueError("Cannot read saved index location; repair the location file or set OUTLOOK_RAG_DATA_ROOT.") from exc


def configure(cfg: dict, config_path: Path, *, from_file: bool) -> list[Path]:
    """Pure path resolution, including previous default roots for compatibility."""
    if cfg.get("data_root"):
        cfg["_explicit_data_root"] = True
        return [Path(cfg["data_root"])]
    scope = os.path.normcase(str(config_path)) if from_file else "environment"
    digest = hashlib.sha256(scope.encode()).hexdigest()[:16]
    location = local_root() / "outlook-rag" / "locations" / f"{digest}.json"
    cfg["_index_location_file"] = str(location)
    root = remembered_root(location)
    # Once chosen, a disconnected drive must not redirect reads to another DB.
    return [root] if root is not None else [default_root(), local_root() / "outlook-rag"]


def index_name(cfg: dict) -> tuple[str, str]:
    from . import app
    digest = hashlib.sha256(app.index_identity(cfg).encode()).hexdigest()[:12]
    model = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", cfg["model"]).strip(" .")[:80].rstrip(" .") or "model"
    if model.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        model = "model_" + model
    dimensions = app.vector_dimensions(cfg) if cfg.get("dimensions") else "auto"
    return f"{model}-{dimensions}d-{digest}", digest


def has_index(folder: Path) -> bool:
    # Empty job status directories must not fix placement before the first sync.
    return (folder / "metadata.sqlite").is_file() or any((folder / "jobs").glob("*.config.json"))


def select(cfg: dict, roots: list[Path]) -> dict:
    name, digest = index_name(cfg)
    candidates = [folder for root in roots for folder in (root / name, root / digest)]
    directory = str(next((p for p in candidates if has_index(p)), candidates[0]))
    return dict(cfg, data_dir=directory, _selected_data_dir=directory,
                _index_roots=[str(root) for root in roots])


def pst_directories(namespace: Any) -> list[Path]:
    """Read store file paths only; never inspect messages or folders."""
    from .outlook import safe_attr
    directories = []
    for store in namespace.Stores:
        path = str(safe_attr(store, "FilePath", ""))
        if path.casefold().endswith(".pst"):
            value = Path(path)
            if not value.is_absolute():
                raise ValueError("Outlook returned a relative PST path; set OUTLOOK_RAG_DATA_ROOT.")
            directories.append(value.parent)
    return list(dict.fromkeys(directories))


def connected_pst_directories() -> list[Path]:
    if sys.platform != "win32":
        return []
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    namespace = None
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        return pst_directories(namespace)
    finally:
        namespace = None
        pythoncom.CoUninitialize()


def prepare(cfg: dict) -> dict:
    """Choose once, before worker files/writer locks; never move an existing DB."""
    if not cfg.get("_automatic_data_dir") or cfg.get("_explicit_data_root"):
        return cfg
    if cfg.get("_selected_data_dir") != cfg["data_dir"]:
        # Callers may explicitly select a restored directory in a resolved cfg.
        return cfg
    from . import jobs
    location = Path(cfg["_index_location_file"])
    location.parent.mkdir(parents=True, exist_ok=True)
    with FileLock(str(location) + ".lock", timeout=30):
        saved = remembered_root(location)
        roots = [saved] if saved is not None else [default_root(), local_root() / "outlook-rag"]
        selected = select(cfg, roots)
        if has_index(Path(selected["data_dir"])):
            root = Path(selected["data_dir"]).parent
        elif saved is not None:
            root = saved
        else:
            parents = connected_pst_directories()
            if len(parents) > 1:
                raise ValueError("Connected PSTs are in multiple directories. Set OUTLOOK_RAG_DATA_ROOT to one directory (for example D:\\Mail\\outlook-rag), then retry. No index was created.")
            root = parents[0] / "outlook-rag" if parents else default_root()
        # An inaccessible remembered root must fail, never select another DB.
        root.mkdir(parents=True, exist_ok=True)
        if saved != root:
            jobs.write(location, {"version": 1, "index_root": str(root)})
        return select(cfg, [root])
