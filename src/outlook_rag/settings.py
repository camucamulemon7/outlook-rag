"""Configuration parsing, credentials and embedding identity.

The shared runtime is supplied by app; this module never imports app.
"""
from __future__ import annotations

import json
import math
import os
from pathlib import Path
from typing import Any, Callable
import threading


from . import coverage, paths


_KEY_CACHE: dict[tuple, str] = {}
_KEY_LOCK = threading.Lock()


def boolean(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (str, int)):
        token = str(value).strip().casefold()
        if token in {"true", "yes", "1", "on"}:
            return True
        if token in {"false", "no", "0", "off"}:
            return False
    raise ValueError("Expected a boolean (true/false/yes/no/1/0)")


def validate_values(cfg: dict) -> None:
    """Coerce known settings before comparisons or provider calls; omit inputs from errors."""
    bounds = {
        "dimensions": (1, 65536), "storage_dimensions": (1, 65536),
        "embedding_batch_size": (1, 128), "embedding_concurrent_requests": (1, 8),
        "embedding_max_batch_chars": (1600, 262144), "embedding_request_dimensions": (0, 65536),
        "sync_batch_emails": (1, 256), "sync_prefetch_emails": (0, 256),
        "sync_max_emails": (1, 2000), "sync_max_total_emails": (1, 20000),
        "since_days": (0, 36500), "sync_max_scanned": (1, 100000),
        "sync_retry_limit": (0, 200), "query_cache_size": (0, 10000),
        "chunk_size": (100, 16000), "chunk_overlap": (0, 15999),
        "search_candidates": (10, 5000), "rrf_k": (1, 1000),
        "search_nprobes": (1, 65536), "search_ef": (1, 65536),
        "ann_num_partitions": (1, 4096),
        "text_cleaning_version": (1, 3),
    }
    for name, (low, high) in bounds.items():
        if name not in cfg:
            continue
        value = cfg[name]
        if name in {"dimensions", "search_nprobes", "search_ef", "ann_num_partitions"} and (value is None or str(value).casefold() == "auto"):
            cfg[name] = None
            continue
        if name == "since_days" and str(value).casefold() == "all":
            value = 0
        try:
            if isinstance(value, bool) or not isinstance(value, (str, int)):
                raise ValueError()
            number = int(value)
            if not low <= number <= high:
                raise ValueError()
        except (ValueError, TypeError, OverflowError) as exc:
            raise ValueError(f"{name} must be an integer in {low}..{high}") from exc
        cfg[name] = number
    for name, low, high in (("sync_max_seconds", 1, 3600), ("mcp_tool_timeout_seconds", 1, 300), ("semantic_weight", 0, 10), ("lexical_weight", 0, 10)):
        if name in cfg:
            try:
                if isinstance(cfg[name], bool):
                    raise ValueError()
                floating = float(cfg[name])
                if not math.isfinite(floating) or not low <= floating <= high:
                    raise ValueError()
            except (ValueError, TypeError, OverflowError) as exc:
                raise ValueError(f"{name} must be a finite number in {low}..{high}") from exc
            cfg[name] = floating
    for name in ("warmup_on_start", "exclude_system_folders", "search_exact", "index_attachments"):
        if name in cfg:
            try:
                cfg[name] = boolean(cfg[name])
            except ValueError as exc:
                raise ValueError(f"{name} must be a boolean") from exc
    for name in ("excluded_folders", "folders"):
        if name in cfg and not (name == "folders" and cfg[name] == "auto"):
            if not isinstance(cfg[name], list) or not all(isinstance(v, str) and v for v in cfg[name]):
                raise ValueError(f"{name} must be an array of nonempty strings")
    for name in ("model", "embedding_url", "data_dir", "data_root", "key_file", "api_key_env", "query_prefix", "document_prefix", "query_instruction", "model_revision", "mcp_tool_profile", "search_mode"):
        if name in cfg and cfg[name] is not None and not isinstance(cfg[name], str):
            raise ValueError(f"{name} must be a string")

def config(backend: Any, *, resolve_dimensions: bool = True) -> dict:
    path = os.environ.get("OUTLOOK_RAG_CONFIG")
    config_path = Path(path).expanduser().resolve() if path else Path.cwd() / "config.json"
    cfg = json.loads(config_path.read_text(encoding="utf-8-sig")) if path else {}
    if not isinstance(cfg, dict):
        raise ValueError("Configuration must be a JSON object")
    converters: dict[str, Callable[[str], Any]] = {"dimensions": lambda value: None if value.lower() == "auto" else int(value), "storage_dimensions": int, "embedding_batch_size": int,
                  "embedding_concurrent_requests": int, "embedding_max_batch_chars": int,
                  "sync_prefetch_emails": int, "sync_batch_emails": int, "sync_max_emails": int, "sync_max_total_emails": int, "text_cleaning_version": int,
                  "since_days": lambda value: 0 if value.lower() == "all" else int(value),
                  "folders": lambda value: "auto" if value.lower() == "auto" else json.loads(value),
                  "excluded_folders": json.loads, "exclude_system_folders": boolean,
                  "chunk_size": int, "chunk_overlap": int, "warmup_on_start": boolean,
                  "sync_max_seconds": float, "mcp_tool_timeout_seconds": float, "sync_max_scanned": int, "sync_retry_limit": int,
                  "embedding_request_dimensions": int, "query_cache_size": int,
                  "search_candidates": int, "rrf_k": int, "semantic_weight": float, "lexical_weight": float,
                  "search_nprobes": int, "search_ef": int, "ann_num_partitions": int, "search_exact": boolean, "index_attachments": boolean}
    fields = ("embedding_url", "model", "data_dir", "data_root", "mcp_tool_profile", "key_file", "api_key_env", "query_instruction", "query_prefix", "document_prefix", "model_revision", "search_mode", "embedding_encoding_format", *converters)
    for field in fields:
        value = os.environ.get("OUTLOOK_RAG_" + field.upper())
        if value is not None:
            try:
                cfg[field] = converters.get(field, str)(value)
            except (ValueError, TypeError, OverflowError) as exc:
                raise ValueError(f"Invalid setting: {field}") from exc
    if not isinstance(cfg, dict):
        raise ValueError("Configuration must be a JSON object")
    validate_values(cfg)
    cfg.setdefault("embedding_encoding_format", "float")
    cfg.setdefault("index_attachments", False)
    if cfg["embedding_encoding_format"] not in ("float", "base64"):
        raise ValueError("embedding_encoding_format must be float or base64")
    if not cfg.get("model"):
        raise ValueError("Set OUTLOOK_RAG_MODEL to your embedding model ID.")
    cfg.setdefault("embedding_url", "http://localhost:8080/api/v1/embeddings")
    automatic_data_dir = not cfg.get("data_dir")
    cfg["_automatic_data_dir"] = automatic_data_dir
    local_root = paths.local_root()
    if automatic_data_dir:
        cfg["data_dir"] = str(paths.default_root() / "pending")
    cfg.setdefault("mcp_tool_profile", "standard")
    if cfg["mcp_tool_profile"] not in {"standard", "full"}:
        raise ValueError("mcp_tool_profile must be standard or full")
    cfg.setdefault("key_file", str(local_root / "outlook-rag" / "api-key.dpapi"))
    cfg.setdefault("folders", "auto")
    cfg["folders"] = coverage.folder_selection(cfg["folders"])
    cfg.setdefault("since_days", 0)
    cfg.setdefault("sync_max_emails", 200)
    cfg.setdefault("sync_max_total_emails", 200)
    cfg.setdefault("sync_max_seconds", 30)
    cfg.setdefault("sync_max_scanned", 2000)
    cfg.setdefault("sync_retry_limit", 10)
    cfg.setdefault("sync_prefetch_emails", 16)
    if type(cfg["sync_prefetch_emails"]) is not int or not 0 <= cfg["sync_prefetch_emails"] <= 256:
        raise ValueError("sync_prefetch_emails must be 0..256")
    cfg.setdefault("query_cache_size", 256)
    cfg.setdefault("search_candidates", 100)
    cfg.setdefault("rrf_k", 60)
    cfg.setdefault("semantic_weight", 1.0)
    cfg.setdefault("lexical_weight", 1.0)
    cfg.setdefault("search_mode", "hybrid")
    cfg.setdefault("search_exact", False)
    if cfg["search_mode"] not in {"semantic", "hybrid", "lexical", "auto", "identifier", "metadata", "thread"}:
        raise ValueError("Invalid search_mode")
    if type(cfg["search_exact"]) is not bool:
        raise ValueError("search_exact must be a boolean")
    for name in ("search_nprobes", "search_ef"):
        if cfg.get(name) is not None and (type(cfg[name]) is not int or not 1 <= cfg[name] <= 65536):
            raise ValueError(f"{name} must be 1..65536")
    if not 10 <= cfg["search_candidates"] <= 5000 or not 1 <= cfg["rrf_k"] <= 1000:
        raise ValueError("Invalid search candidate count or RRF constant")
    if not all(math.isfinite(cfg[name]) and 0 <= cfg[name] <= 10 for name in ("semantic_weight", "lexical_weight")) or not cfg["semantic_weight"] + cfg["lexical_weight"]:
        raise ValueError("Search weights must be finite, nonnegative, at most 10, and not both zero")
    if not 1 <= cfg["sync_max_seconds"] <= 3600 or not 1 <= cfg["sync_max_scanned"] <= 100000:
        raise ValueError("sync_max_seconds must be 1..3600; sync_max_scanned must be 1..100000")
    if not 0 <= cfg["sync_retry_limit"] <= 200 or not 0 <= cfg["query_cache_size"] <= 10000:
        raise ValueError("sync_retry_limit must be 0..200; query_cache_size must be 0..10000")
    cfg.setdefault("exclude_system_folders", True)
    cfg.setdefault("excluded_folders", [])
    cfg.setdefault("chunk_size", 1600)
    cfg.setdefault("chunk_overlap", 200)
    if not 100 <= cfg["chunk_size"] <= 16000 or not 0 <= cfg["chunk_overlap"] < cfg["chunk_size"]:
        raise ValueError("chunk_size must be 100..16000 and chunk_overlap must be smaller than chunk_size")
    if cfg["folders"] != "auto" and (not isinstance(cfg["folders"], list) or not all(isinstance(item,str) for item in cfg["folders"])):
        raise ValueError("folders must be 'auto' or a JSON array of folder paths")
    if not isinstance(cfg["excluded_folders"], list):
        raise ValueError("excluded_folders must be a JSON array of wildcard paths")
    for field in ("data_dir", "data_root", "key_file"):
        if cfg.get(field):
            field_path = Path(os.path.expandvars(cfg[field])).expanduser()
            cfg[field] = str(field_path if field_path.is_absolute() else config_path.parent / field_path)
    if resolve_dimensions and not cfg.get("dimensions"):
        cfg["dimensions"] = backend.detect_dimensions(cfg)
    if cfg.get("dimensions") and "qwen3-embedding" in cfg["model"].lower():
        cfg.setdefault("storage_dimensions", min(1024, cfg["dimensions"]))
        cfg.setdefault("embedding_request_dimensions", cfg["storage_dimensions"])
    requested = cfg.get("embedding_request_dimensions", 0)
    if requested and cfg.get("dimensions") and not backend.vector_dimensions(cfg) <= requested <= cfg["dimensions"]:
        raise ValueError("embedding_request_dimensions must cover storage_dimensions and not exceed dimensions")
    for field in ("embedding_url", "model", "dimensions", "data_dir"):
        if field == "dimensions" and not resolve_dimensions:
            continue
        if not cfg.get(field):
            raise ValueError(f"Configuration requires {field}.")
    cfg.setdefault("text_cleaning_version", 2)
    if cfg["text_cleaning_version"] not in (1, 2, 3):
        raise ValueError("text_cleaning_version must be 1, 2 or 3")
    for field, default, maximum in (("embedding_batch_size", 8, 128), ("embedding_concurrent_requests", 2, 8), ("sync_batch_emails", 16, 256)):
        cfg.setdefault(field, default)
        if not isinstance(cfg[field], int) or not 1 <= cfg[field] <= maximum:
            raise ValueError(f"{field} must be 1..{maximum}")
    cfg.setdefault("embedding_max_batch_chars", 24000)
    if not 1600 <= cfg["embedding_max_batch_chars"] <= 262144:
        raise ValueError("embedding_max_batch_chars must be 1600..262144")
    if automatic_data_dir:
        cfg = paths.select(cfg, paths.configure(cfg, config_path, from_file=bool(path)))
    return cfg


def query_instruction(backend: Any, cfg: dict) -> str:
    # Most OpenAI-compatible embedding models accept plain query text.
    return cfg.get("query_instruction", backend.QUERY_INSTRUCTION if "qwen3-embedding" in cfg["model"].lower() else "")


def index_identity(backend: Any, cfg: dict) -> str:
    return json.dumps({"model": cfg["model"], "dimensions": cfg.get("dimensions"),
        "instruction": backend.query_instruction(cfg), "chunk_version": cfg.get("text_cleaning_version", 2),
        **({"storage_dimensions": backend.vector_dimensions(cfg)} if cfg.get("dimensions") and backend.vector_dimensions(cfg) != cfg["dimensions"] else {}),
        **({field: cfg[field] for field in ("query_prefix", "document_prefix") if field in cfg}),
        **({field: cfg[field] for field, default in (("chunk_size", 1600), ("chunk_overlap", 200)) if cfg.get(field,default) != default}),
        **({"model_revision": cfg["model_revision"]} if cfg.get("model_revision") else {}),
        **({"attachment_extraction_version": 1} if cfg.get("index_attachments") else {})},sort_keys=True)


def api_key(backend: Any, cfg: dict) -> str:
    key = os.environ.get(cfg.get("api_key_env", "OUTLOOK_RAG_API_KEY"))
    if key:
        return key
    if not cfg.get("key_file") or not Path(cfg["key_file"]).is_file():
        raise ValueError("Set OUTLOOK_RAG_API_KEY or key_file. The set-key command saves a Windows-encrypted key file.")
    path = Path(cfg["key_file"]).resolve()
    stat = path.stat()
    identity = (str(path), stat.st_mtime_ns, stat.st_size, stat.st_ino)
    with _KEY_LOCK:
        if identity not in _KEY_CACHE:
            import win32crypt
            key = win32crypt.CryptUnprotectData(path.read_bytes(), None, None, None, 0)[1].decode("utf-8")
            _KEY_CACHE.clear()
            _KEY_CACHE[identity] = key
        return _KEY_CACHE[identity]


def vector_dimensions(backend: Any, cfg: dict) -> int:
    dimensions = cfg.get("storage_dimensions", cfg["dimensions"])
    if type(dimensions) is not int or not 1 <= dimensions <= cfg["dimensions"]:
        raise ValueError("storage_dimensions must be 1..dimensions. Reduction requires an MRL-compatible model.")
    return dimensions

