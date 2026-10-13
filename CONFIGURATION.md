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
| OUTLOOK_RAG_DATA_DIR | %USERPROFILE%\Documents\Outlookファイル\outlook-rag\<model>-<dimensions>d-<settings-id> |
| OUTLOOK_RAG_KEY_FILE | %LOCALAPPDATA%\outlook-rag\api-key.dpapi |
| OUTLOOK_RAG_FOLDERS | auto: recursively discover connected Outlook mail folders |
| OUTLOOK_RAG_SINCE_DAYS | 0: all dates; "all" is also accepted |
| OUTLOOK_RAG_EXCLUDE_SYSTEM_FOLDERS | true: exclude Deleted Items, Junk, and synchronization-error folders |
| OUTLOOK_RAG_EXCLUDED_FOLDERS | []: no additional exclusions |
| OUTLOOK_RAG_EMBEDDING_BATCH_SIZE | 8 chunks |
| OUTLOOK_RAG_EMBEDDING_CONCURRENT_REQUESTS | Up to 2 requests |
| OUTLOOK_RAG_EMBEDDING_MAX_BATCH_CHARS | 24000 characters |
| OUTLOOK_RAG_SYNC_BATCH_EMAILS | Save batches of 16 emails |
| OUTLOOK_RAG_SYNC_PREFETCH_EMAILS | Read up to 16 additional rows during embedding requests; 0 disables prefetch (range 0..256) |
| OUTLOOK_RAG_SYNC_MAX_EMAILS | Up to 200 mail attempts per folder per call |
| OUTLOOK_RAG_SYNC_MAX_TOTAL_EMAILS | Up to 200 mail attempts across the entire call |
| OUTLOOK_RAG_SYNC_MAX_SECONDS | 30-second cooperative budget; in-flight calls may overrun |
| OUTLOOK_RAG_SYNC_MAX_SCANNED | 2000 new metadata rows, including unchanged/non-mail rows; cursor overlap is skipped separately |
| OUTLOOK_RAG_SYNC_RETRY_LIMIT | Retry up to 10 due failures per call, within the overall mail/time budget |
| OUTLOOK_RAG_QUERY_CACHE_SIZE | 256 query embeddings, evicted by least recent use; 0 disables caching |
| OUTLOOK_RAG_EMBEDDING_REQUEST_DIMENSIONS | Qwen3-Embedding: stored dimensions. Other models: unset. 0 disables server-side reduction |
| OUTLOOK_RAG_SEARCH_CANDIDATES | Initial 100 candidates per search source (also at least 10 times the result limit); expand when deduplication leaves too few results |
| OUTLOOK_RAG_RRF_K | 60 |
| OUTLOOK_RAG_SEMANTIC_WEIGHT | 1.0 |
| OUTLOOK_RAG_LEXICAL_WEIGHT | 1.0 |
| OUTLOOK_RAG_SEARCH_MODE | hybrid; semantic, lexical, identifier, auto, metadata and thread are also supported |
| OUTLOOK_RAG_SEARCH_NPROBES | Unset: SDK default. 1..65536 partitions probed by ANN |
| OUTLOOK_RAG_SEARCH_EF | Unset: SDK default. 1..65536 HNSW effort; raised to at least the current candidate count |
| OUTLOOK_RAG_SEARCH_EXACT | false; true bypasses ANN for an exact cosine baseline |
| OUTLOOK_RAG_CHUNK_SIZE | 1600 characters |
| OUTLOOK_RAG_CHUNK_OVERLAP | 200 characters |
| OUTLOOK_RAG_TEXT_CLEANING_VERSION | 2: enhanced text cleanup; opt-in 3 packs paragraphs and preserves recognized forwards/quote-only bodies |
| OUTLOOK_RAG_WARMUP_ON_START | true: send a small embedding warmup request |
| OUTLOOK_RAG_QUERY_INSTRUCTION | Default retrieval instruction for Qwen3-Embedding; empty for other models |
| OUTLOOK_RAG_QUERY_PREFIX | Empty; set if required by your model |
| OUTLOOK_RAG_DOCUMENT_PREFIX | Empty; set if required by your model |
| OUTLOOK_RAG_MODEL_REVISION | Empty; use to distinguish weight revisions under the same model ID |
| OUTLOOK_RAG_API_KEY_ENV | OUTLOOK_RAG_API_KEY |

`OUTLOOK_RAG_FOLDERS` and `OUTLOOK_RAG_EXCLUDED_FOLDERS` accept JSON array strings. For example, set `OUTLOOK_RAG_EXCLUDED_FOLDERS` to `["*/Advertisements","*/RSS Feeds"]`. Exclusions support wildcard paths. Search folders are always excluded from automatic discovery to avoid duplicate indexing.

For ordinary email searches, omit `folder`, `since` and `until` unless the user explicitly narrows the scope. Searches include all indexed stores and PST archives even if configured synchronization is limited to `["inbox"]`. Configuration controls which folders get indexed; it does not silently filter existing search results. Use `folders=["auto"]`, `since_days=0` with `sync_emails` or `start_sync_job` to index all connected PST stores despite an inbox-only setting; CLI equivalent: `job start --folders auto --since-days 0`. Use `["auto"]` alone, not mixed with paths. Named exclusions still apply.

`search_scope` reports the applied folder/date filters. `coverage` describes the last saved sync inventory, known PST stores and pending scans; it does not claim to have checked live Outlook during search. `index_status.search_coverage` exposes the same summary. Older inventories without complete scope metadata are reported as unknown. `list_outlook_sources` reads currently connected store/folder metadata and compares it with all-date checkpoints, reporting unindexed archives, discovery errors and retained failures. Zero results include `absence_confirmed=false`; check coverage before concluding that a message does not exist. Connected PSTs must be indexed; disconnected PSTs cannot be read from Outlook. MCP tool descriptions and initialization instructions request all-store search by default, but clients must load the updated server rather than an older pinned Git tag.

Startup does not perform full indexing or schedule synchronization. Call `sync_emails` to update the database. For a small batch, pass `{"max_total_emails": 10, "max_seconds": 5, "max_scanned": 100}`. Repeat until folder windows are complete and no folders remain pending. Folders rotate between calls so large folders do not starve smaller ones.

Mail attempts include retries and failed reads, not just successful updates. Metadata uses Outlook GetTable/GetArray where available; unsupported providers fall back to Items. Checkpoints are saved after successful batches. Incremental filters and reconciliation progress survive bounded calls.

GetTable short-term EntryIDs are mapped to canonical MailItem IDs before indexing. Aliases are scoped to the Outlook profile and process lifetime; if that identity cannot be verified, mappings are resolved again. Existing databases retain their mail IDs. Reconciliation skips deletion when any mail ID could not be resolved during the scan.

Unreadable mail and permanent input errors are retained for retry with increasing backoff. Provider/network outages stop the current call without advancing over uncommitted input. `index_status` exposes the failed-mail count; a completed scan can still have failures awaiting retry.

Completed embedding requests are cached immediately, including successful requests already in flight when another request fails. A retry reuses those chunks after a process restart. Interrupted mail updates keep the previous searchable version until all replacement chunks are ready. Maintenance retains caches referenced by unfinished mail batches; references are released when a mail is published or its pending content is replaced.

For bulk indexing, call `start_sync_job` once, then inspect `sync_job_status`. The detached worker repeats bounded sync cycles without agent/tool invocation gaps. It survives MCP disconnects and keeps one worker per database. `cancel_sync_job` requests cooperative cancellation while retaining indexed mail, checkpoints and completed embeddings. Existing sync settings apply; no additional settings are required. Startup alone does not launch a bulk job.

Job records and resolved settings are stored under `<data_dir>/jobs`. API key environment values are inherited by the worker and are not written into job snapshots. Encrypted key-file paths are retained. Worker crashes/computer restarts are reported as interrupted; start another job to resume the index. Jobs with failed mail/folders end as completed_with_errors; repeated provider outages or lack of progress end as blocked. No automatic ANN rebuild or continuous mail monitoring is performed by a bulk job.

Job diagnostics are stored in `<data_dir>/jobs/<job_id>.log` and exposed as `log_path` / `recent_logs` by status. Failed jobs include a sanitized `error_detail`. Progress and exception stack locations are recorded without raw messages, source lines, stdout/stderr or API credentials. A stdlib bootstrap records import/configuration and Windows environment-transfer errors. Each log is bounded to about 1 MiB; inactive logs older than 14 days are pruned by maintenance.

`maintain_index` prunes unused document embedding caches, enforces the query cache limit, and compacts vector storage while retaining versions for seven days. It does not delete mail or current chunks.

Server-side dimension requests do not change index identity: full output is still detected as the model's native dimension, and stored dimensions remain fixed. Unsupported dimension requests fall back once per process. Only request reduced dimensions from models that support MRL.

Candidate count, search mode, ANN search parameters and RRF weights do not change embedding identity or require re-indexing. Hybrid weights apply in hybrid and thread modes (parent excerpts use the lexical weight); semantic and lexical modes enable their single source with weight 1. Use a labeled query set to compare changes before selecting new defaults. `mode` in an individual search overrides `SEARCH_MODE`, and `hybrid=false` selects semantic mode when `mode` is omitted.

`mode="metadata"` requires no embedding or LanceDB access. It supports `query=""`, sender, inclusive dates, exact folder, literal subject and recipient substrings, and `cursor` paging. A nonempty metadata query is a literal subject substring, not a natural-language request. SQL uses ASCII case folding for substrings. Results sort by received UTC instant and mail ID descending; unknown dates sort last. Reuse `next_cursor` with identical query/filters. Pages reflect the current sync state, not a frozen multi-page snapshot; changed timestamps during sync can affect pagination. `subject`, `recipient`, `cursor` cannot be used in other modes. Similarity thresholds and thread grouping cannot be used in metadata mode.

CLI metadata lists can omit the positional query: `outlook-rag search --mode metadata --sender tanaka@ --subject PCIe`. This avoids empty-string argument handling differences between PowerShell versions. Other modes require a nonempty query.

`mode="thread"` combines message hybrid search with BM25 of derived thread excerpts, adds their ranks through RRF, groups by store/conversation and attaches bounded chronological original mail evidence. Filters are applied to every context query, and whole-thread excerpt retrieval is disabled whenever sender/folder/date filters are supplied. The earliest and latest eligible messages and neighbors around the matched message are included when the global 12,000-character body budget permits (1,200 characters per body, at most five sources per result). Body/selection truncation is explicit. No generated summary or fact/decision inference runs. Cached originals are evidence, so agents must verify chronological changes before asserting a final decision. `get_mail_thread` itself returns a whole-conversation window and has no search filters: check original dates/senders/folders when using it after a restricted search.

Thread documents contain at most 6,000 characters: source IDs and excerpts from the latest message, first two and last twelve messages, with deduplication. They do not represent every intermediate decision in long threads. On a new index all conversations are updated incrementally with canonical SQL publication; metadata/conversation changes, moves and deletions refresh only affected owners. Unchanged fingerprints avoid rewriting FTS rows. Direct row-ID deletes avoid scanning all thread documents per update. No additional model call is required.

Existing indexes start with a partial derived thread index after the first writer initializes its schema. Read-only searches do not initialize or rebuild it. Run `outlook-rag rebuild-threads` for an explicit full build from cached originals, without API credentials. The rebuild holds the existing writer lock and publishes derived rows atomically in SQLite; another writer reports busy, and failures keep the previous rows. UTC sort indexes are created once during writer initialization or this explicit rebuild. `index_status.thread_index` reports state and representation; `complete` means all cached conversations are represented, not that their entire history fits the excerpt budget. A missing index still permits message retrieval and original context. Backups include these derived tables.

Text-cleaning version 3 preserves short emails as single chunks, packs whole paragraphs, and splits oversized paragraphs at sentence/space boundaries. Overlap reuses whole paragraphs only when they fit its character budget; oversized-paragraph splits do not use a hard character overlap. Recognized `Fwd:`, `Fw:` or `転送:` subjects retain forwarded originals, and a quote-only body retains its sole content. Authored replies still use conservative quote removal. Keep version 2 when compatibility is needed: v1/v2 behavior is unchanged. Version 3 changes the embedding identity and needs a full all-date/all-store sync into a new `data_dir` (or the automatic new directory). It does not migrate existing chunks in place. `rebuild-vectors` and `rebuild-threads` rebuild existing derived data, not v2-to-v3 chunking. Stop older servers before writing the upgraded index, and retain a backup.

When chunk candidates collapse into too few distinct emails or threads, search doubles the pool up to 5,000 candidates per source. A five-second cooperative expansion budget is checked between rounds; a single database operation may overrun it. Results include `candidate_limit`, `candidate_cap`, `candidate_limit_reached` and `candidate_time_limit_reached`.

Large-index search costs depend on chunk count, not just mail count: one million mails averaging five chunks means five million vectors. After bulk sync, use `optimize_index` (CLI `optimize`) and check `index_status.vector_coverage` for ANN and unindexed rows; newly added chunks remain searchable but are scanned until indexed. Hybrid/lexical search uses FTS5's native BM25 ranked cursor to avoid joining and sorting every match before returning the requested candidates. Equal BM25 scores follow FTS5's tie order; date/folder/sender filters are still applied before the candidate limit.

Specify inclusive `since` / `until` dates when you want a narrower period, for example `search_emails(query="出張の予約", since="2025-01-01", until="2025-12-31")`. Searching a period includes indexed archives in that period. The server does not silently narrow dates to speed up searches. The expansion budget is not a wall-clock timeout for query embedding, locks or a single database operation; API and disk performance still affect total latency. Compare representative queries on the target PC before selecting a restricted scope.

## Archives and file locations

`list_outlook_sources` returns connected store names, available PST/OST paths, and folders without reading email bodies. The equivalent CLI command is `outlook-rag sources`.

Open archive PSTs are included in discovery. Disconnected PSTs are not searched for on disk. Stores without a local file, such as some server-side archives, may return an empty file path.

Folder counts come from Outlook's `Items.Count` and may include non-mail items. Synchronization indexes mail items only.

## Database changes and authentication

New indexes default to `%USERPROFILE%\Documents\Outlookファイル\outlook-rag\<model>-<dimensions>d-<settings-id>`. The model name is sanitized for Windows and stored dimensions are visible; the settings ID prevents mixing embedding/chunk configurations. Existing hash-only default indexes in Documents or Local AppData remain in use, with the new named location taking precedence if it exists. Explicit `OUTLOOK_RAG_DATA_DIR` values are unchanged. API key files remain under Local AppData. To relocate an existing index, stop its MCP and background workers, move the complete index directory, and set `OUTLOOK_RAG_DATA_DIR` to the new path.

Changing the model, vector dimensions, chunking, or other index identity settings requires re-indexing. The default database path selects a separate index automatically. If `OUTLOOK_RAG_DATA_DIR` is fixed, choose a new directory for the new settings.

SQLite now records the LanceDB version published with each mail/chunk transaction. Searches read matching snapshots; a failed or interrupted vector write leaves the previous version searchable and is restored before another write. Existing indexes are adopted without changing IDs or re-indexing. Stop MCP and bulk workers before upgrading, and do not run older server versions against an index using this publication marker. Maintenance protects the published rollback version while compacting storage.

SQLite is the canonical local database; LanceDB is a rebuildable vector index. No persistent DB service is needed. `doctor`, `backup-index`, `restore-index`, `rebuild-vectors`, `optimize` and `maintain` use saved dimensions and require no API credentials. Back up both stores together with `backup-index <new-directory>`; copying only a live SQLite file misses its WAL and the matching vector generation. Restore with `restore-index <backup> <new-directory>` and update `data_dir` afterward. The destinations must not exist; backup cannot be inside `data_dir`. These commands exclude credentials and worker jobs.

`rebuild-vectors` validates every cached chunk vector, builds a separate vector table with metadata and indexes, repairs lexical/identifier indexes, then publishes the table name/version in one SQLite transaction. It also upgrades legacy tables without re-embedding. Failure keeps the old publication. Missing or invalid cached vectors require re-syncing the affected content; doctor reports whether an offline rebuild is possible. Old generations consume disk until `maintain_index` removes them under the reader lock. Stop older application versions before using the new generation marker.

Offline automatic index discovery logs unreadable SQLite/JSON/snapshot candidates and continues to a valid matching index. An explicitly selected damaged directory, or no usable match with damaged candidates present, produces a recovery error. Model/settings matching and ambiguous-dimension checks still apply.

The CLI also accepts `--config <path>`; MCP environment variables override values in that file. Relative database and key-file paths in a config file are resolved against its directory.

To store an encrypted API key interactively, run `outlook-rag set-key` with the desired settings. Input is hidden. DPAPI keys can be decrypted only by the Windows user who saved them; save the key again on another computer or account.

During embedding requests, sync reads ahead on the Outlook/SQLite owner thread. HTTP requests run in the existing worker pool. At most `SYNC_PREFETCH_EMAILS` extra rows are buffered; already-indexed mail is checked before reading its body. Prefetched rows do not advance checkpoints until processed and saved. Cancellation, time limits and provider failures leave unpublished rows for the next sync. API batch size/concurrency and indexed content are unchanged. Folder `batch_stats` reports `prefetched_rows`, `prefetch_peak_rows` and `prefetch_read_ms`; these timings overlap embedding time and should not be added to it.

Search filters are evaluated against matching SQL/vector snapshots. New and rebuilt indexes use native folder, ASCII-case-insensitive sender substring, and inclusive received-date prefilters, without collecting every eligible mail ID. Candidate metadata is rechecked in SQLite. Legacy tables use at most 1,000 eligible email IDs per expression, merging shard top candidates by cosine distance; run `rebuild-vectors` to upgrade. Search batches candidate metadata and thread counts without loading full mail bodies or writing schema state. Query-cache writes are best effort when a sync transaction owns SQLite; a cached query can still be searched during that write.

`search_emails` and CLI `search --min-similarity` accept an optional cosine lower bound from -1 to 1. The default is unset. BM25 candidates also receive a measured cosine similarity when filtering is enabled. Calibrate thresholds with your model; scores are not probabilities. Snippets prefer query context within the matched chunk and remain plain text. Filtering occurs before grouping and can legitimately return zero results.

CLI `job start/status/cancel`, `failures list/retry`, `progress` and `thread` reuse MCP operations and return JSON. Cached reads and cancellation recover dimensions locally and need no API key. Job start/retry use the configured provider credentials. `--config` precedes the subcommand; `--help` is available for each command group.

CLI `search --mode lexical` and `search --mode identifier` need no provider unless `--min-similarity` requests a cosine check. Identifier search accepts exactly one complete ASCII letter/digit ID with `-` or `_`, such as `HT-5639`, and orders exact matches by received date. Unknown IDs return zero results. `auto` chooses this path for one unambiguous positive ID and otherwise semantic search; multiple IDs and common negation phrases are left to semantic retrieval. Auto is opt-in and is a heuristic, not a Boolean parser. Without a calibrated threshold, natural-language searches can return irrelevant relative neighbors.

A search source with weight `0` is disabled and cannot add zero-score candidates. With only BM25 enabled and no similarity threshold, search does not request a query embedding. A requested threshold still needs a query vector, which can come from the cache. An empty metadata-filter result returns before query embedding; nonempty conditions are checked again in the matching SQL/vector snapshot. Snippet offsets are measured in the original Unicode text. WMI startup failures retain their numeric process-creation return code in sanitized diagnostics.
