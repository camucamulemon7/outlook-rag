# Configuration

All settings can be supplied through the MCP server's `environment` object. Only the embedding model and API credentials are required; see [opencode.example.json](opencode.example.json).

```json
{
  "OUTLOOK_RAG_MODEL": "Qwen/Qwen3-Embedding-8B",
  "OUTLOOK_RAG_API_KEY": "YOUR_API_KEY"
}
```

Set `OUTLOOK_RAG_EMBEDDING_URL` if your endpoint differs from the default. As an alternative to a plaintext environment variable, `OUTLOOK_RAG_KEY_FILE` can point to a Windows DPAPI-encrypted key file. Do not publish configurations containing real credentials.

## Defaults

| Environment variable | Default or behavior |
| --- | --- |
| OUTLOOK_RAG_EMBEDDING_URL | http://localhost:8080/api/v1/embeddings |
| OUTLOOK_RAG_DIMENSIONS | Detect API output dimensions on the first request; reuse within the process |
| OUTLOOK_RAG_STORAGE_DIMENSIONS | Up to 1024 for Qwen3-Embedding; full API output dimensions for other models |
| OUTLOOK_RAG_DATA_DIR | %LOCALAPPDATA%\outlook-rag\<index-settings-hash> |
| OUTLOOK_RAG_KEY_FILE | %LOCALAPPDATA%\outlook-rag\api-key.dpapi |
| OUTLOOK_RAG_FOLDERS | auto: recursively discover connected Outlook mail folders |
| OUTLOOK_RAG_SINCE_DAYS | 0: all dates; "all" is also accepted |
| OUTLOOK_RAG_EXCLUDE_SYSTEM_FOLDERS | true: exclude Deleted Items, Junk, and synchronization-error folders |
| OUTLOOK_RAG_EXCLUDED_FOLDERS | []: no additional exclusions |
| OUTLOOK_RAG_EMBEDDING_BATCH_SIZE | 8 chunks |
| OUTLOOK_RAG_EMBEDDING_CONCURRENT_REQUESTS | Up to 2 requests |
| OUTLOOK_RAG_EMBEDDING_MAX_BATCH_CHARS | 24000 characters |
| OUTLOOK_RAG_SYNC_BATCH_EMAILS | Save batches of 16 emails |
| OUTLOOK_RAG_SYNC_MAX_EMAILS | Up to 200 changed emails per folder per call |
| OUTLOOK_RAG_SYNC_MAX_TOTAL_EMAILS | Up to 200 changed emails across the entire call |
| OUTLOOK_RAG_CHUNK_SIZE | 1600 characters |
| OUTLOOK_RAG_CHUNK_OVERLAP | 200 characters |
| OUTLOOK_RAG_TEXT_CLEANING_VERSION | 2: enhanced text cleanup |
| OUTLOOK_RAG_WARMUP_ON_START | true: send a small embedding warmup request |
| OUTLOOK_RAG_QUERY_INSTRUCTION | Default retrieval instruction for Qwen3-Embedding; empty for other models |
| OUTLOOK_RAG_QUERY_PREFIX | Empty; set if required by your model |
| OUTLOOK_RAG_DOCUMENT_PREFIX | Empty; set if required by your model |
| OUTLOOK_RAG_MODEL_REVISION | Empty; use to distinguish weight revisions under the same model ID |
| OUTLOOK_RAG_API_KEY_ENV | OUTLOOK_RAG_API_KEY |

`OUTLOOK_RAG_FOLDERS` and `OUTLOOK_RAG_EXCLUDED_FOLDERS` accept JSON array strings. For example, set `OUTLOOK_RAG_EXCLUDED_FOLDERS` to `["*/Advertisements","*/RSS Feeds"]`. Exclusions support wildcard paths. Search folders are always excluded from automatic discovery to avoid duplicate indexing.

Startup does not perform full indexing or schedule synchronization. Call `sync_emails` to update the database. For a small batch, pass `{"max_total_emails": 10}`. Repeat until folder windows are complete and no folders remain pending.

## Archives and file locations

`list_outlook_sources` returns connected store names, available PST/OST paths, and folders without reading email bodies. The equivalent CLI command is `outlook-rag sources`.

Open archive PSTs are included in discovery. Disconnected PSTs are not searched for on disk. Stores without a local file, such as some server-side archives, may return an empty file path.

Folder counts come from Outlook's `Items.Count` and may include non-mail items. Synchronization indexes mail items only.

## Database changes and authentication

Changing the model, vector dimensions, chunking, or other index identity settings requires re-indexing. The default database path selects a separate index automatically. If `OUTLOOK_RAG_DATA_DIR` is fixed, choose a new directory for the new settings.

The CLI also accepts `--config <path>`; MCP environment variables override values in that file. Relative database and key-file paths in a config file are resolved against its directory.

To store an encrypted API key interactively, run `outlook-rag set-key` with the desired settings. Input is hidden. DPAPI keys can be decrypted only by the Windows user who saved them; save the key again on another computer or account.
