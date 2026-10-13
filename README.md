# outlook-rag

![outlook-rag banner](assets/header.png)

Semantic and keyword search for Outlook email through MCP. Index your mail locally and search it using an OpenAI-compatible embedding API.

## Features

- Semantic search combined with Japanese keyword search (BM25 + vector search).
- All-date indexing of connected mail folders, including open archive PSTs.
- Background bulk indexing with progress and cancellation, resumable updates, and failed-mail retries.
- Persistent document and query embedding caches, plus local database maintenance.
- Support for Qwen and other OpenAI-compatible embedding models.
- Local storage with read-only access to Outlook.

## Requirements

- Windows and Outlook Classic with a configured mail profile.
- [uv](https://docs.astral.sh/uv/getting-started/installation/) on your PATH.
- An embedding API and its model name and API key.

Python 3.11-3.13 is supported. The example uses Python 3.12; uv can download it on the first run.

## Installation

Add the following to your OpenCode configuration. Replace the model and API key with your own values.

```json
{
  "mcp": {
    "servers": {
      "outlook_rag": {
        "type": "local",
        "command": [
          "uvx", "--python", "3.12", "--from",
          "git+https://github.com/camucamulemon7/outlook-rag.git@v0.8.0",
          "outlook-rag"
        ],
        "environment": {
          "OUTLOOK_RAG_MODEL": "Qwen/Qwen3-Embedding-8B",
          "OUTLOOK_RAG_API_KEY": "YOUR_API_KEY"
        }
      }
    }
  }
}
```

uvx installs the server and its dependencies on first connection. No clone or separate application configuration file is required. Reload OpenCode to connect; use an absolute path to uvx if it is not on your PATH.

The default embedding endpoint is `http://localhost:8080/api/v1/embeddings`. Add `OUTLOOK_RAG_EMBEDDING_URL` to `environment` for another endpoint.

For offline startup, replace `@v0.8.0` with the full commit SHA shown on GitHub, run that command online once, then add `--offline` after `uvx`. Keep the Python version and uv caches available. Your embedding API must still be running.

## Usage

1. Call `start_sync_job` for bulk indexing. It returns a `job_id` promptly and continues through bounded sync cycles in a separate process. Call `sync_job_status` to inspect progress, or `cancel_sync_job` to request a safe stop.
2. Call `search_emails` with a natural-language query, then `get_indexed_mail` to read a result.

Ordinary mail lookups should leave `folder`, `since` and `until` unset so search spans all indexed stores, including older mail in connected PST archives. An inbox-only live lookup can miss archived mail. Search responses report `search_scope` and `coverage`; zero hits do not establish absence from unindexed archives. Before answering that mail was not found, compare `index_status` with `list_outlook_sources`, which checks currently connected stores and pending PST folders without reading bodies or requesting embeddings.

If archives have not been indexed, explicitly run `start_sync_job(folders=["auto"], since_days=0)` or repeat bounded `sync_emails` with the same arguments. `["auto"]` overrides an inbox-only `OUTLOOK_RAG_FOLDERS` setting and discovers all connected stores; explicit excluded/system/search folders remain excluded. `search_emails` searches existing indexed mail immediately and does not start a bulk job. An existing job can be reused, so check its `options` to confirm it covers all stores. Restart the MCP server after upgrading to reload the all-store search guidance. A client pinned to an older Git tag continues to use that tag's code.
3. Call `sync_emails` again when you want to include new or changed mail. Sync does not run automatically.

For a smaller sync, call `sync_emails` with `{"max_total_emails": 10}`. After bulk indexing, call `optimize_index` to build the vector search index.

Background jobs survive MCP disconnects. Only one bulk worker runs per database; another start returns the existing job. Cancellation is cooperative, so an in-flight Outlook/API call may finish first. After a computer restart or an interrupted worker, start another job to resume saved checkpoints and cached embeddings. A `completed_with_errors` result means some mail or folders still need attention.

Jobs track folders by store/folder ID. A renamed folder retains its progress and updates cached paths when scanned. If an unfinished folder disappears or its PST is disconnected, the job ends as `completed_with_errors` with `reason=folder_unavailable` and `missing_folders` identifying the affected scope.

`sync_emails` remains available for small foreground updates. Its default limits are 200 mail attempts, 2,000 new metadata rows and a 30-second cooperative budget. Bulk jobs use the same settings and immediately start the next cycle when needed. While embedding requests run, sync reads ahead up to 16 additional rows on the Outlook thread. Set `OUTLOOK_RAG_SYNC_PREFETCH_EMAILS=0` to disable it.

Worker status includes a local `log_path` and the last 20 diagnostic events. Logs include timestamps, progress counters, safe error codes and stack frame locations. Raw exception messages, stdout/stderr, provider bodies and credentials are not persisted. Import/configuration failures are recorded by a stdlib bootstrap before the application loads; a hard kill is still reported as interrupted. Logs are bounded to about 1 MiB per job, and maintenance removes inactive logs older than 14 days.

Failures are retained for later retry and counted by `index_status`. Use `maintain_index` when you want to remove unused embedding caches and compact vector storage; recent versions are retained for seven days.

With an existing matching index, cached operations and maintenance recover dimensions locally when the embedding API is unavailable. Semantic searches need the API only for an uncached query. Lexical and identifier searches need no embeddings unless `min_similarity` is set. Initial indexing needs the API; MCP startup still requires configured API credentials. A saved job snapshot also allows status and cancellation before database initialization. Omitting a job ID selects the most recently created job.

| Tool | Purpose |
| --- | --- |
| `sync_emails` | Update the local mail index |
| `list_sync_failures` | Inspect failed mail IDs, causes and retry schedules |
| `retry_failed_emails` | Retry only selected failed emails |
| `get_sync_progress` | Inspect folder progress, rate and estimated remaining time |
| `start_sync_job` | Start continuous background bulk indexing |
| `sync_job_status` | Inspect the latest or specified bulk job |
| `cancel_sync_job` | Request a safe stop while retaining completed work |
| `search_emails` | Search by meaning and keywords, with metadata filters |
| `get_indexed_mail` | Read a cached email body |
| `get_mail_thread` | Read cached preceding/following mail in the same conversation |
| `index_status` | Check indexed counts and sync progress |
| `list_outlook_sources` | Inspect connected stores, folders, and PST/OST paths |
| `optimize_index` | Build a vector search index from at least 256 chunks |
| `maintain_index` | Prune unused caches and compact local vector storage |

## Recovery, progress and threads

Use `list_sync_failures` to inspect causes and retry schedules, then pass its `mail_id` values to `retry_failed_emails(mail_ids=[...])`. The retry bypasses backoff for only those IDs; another sync owning the DB returns `busy` without changing retry state. Old failure records gain detailed causes after another attempt. Failure lists, saved progress and cached thread context remain available without the embedding API; automatic dimensions are recovered from the matching index. Provider response bodies and credentials are never stored as error details.

`get_sync_progress`, `index_status` and `sync_job_status` expose per-folder indexed/failed counts, scan progress, processing rate and ETA. Remaining counts use Outlook item-count estimates, which can include non-mail items. ETA is `null` when the scope or rate is unknown, including incomplete checkpoints from older versions. A completed scan may still have failed mail to recover. These status calls use saved snapshots and make no live Outlook calls.

Use `search_emails(min_similarity=0.6)` to exclude chunks below a cosine similarity threshold, including BM25-only candidates. The default is `None`, preserving relative ranking; calibrate a threshold for your model rather than interpreting the score as a confidence probability. Long chunks return a plain-text window around a query match when available.

`mode` selects `semantic`, `hybrid`, `lexical`, `identifier` or `auto`. The default remains `hybrid`; `hybrid=false` still selects semantic search. For paraphrases, select `semantic`. `identifier` requires one complete ASCII ID such as `HT-5639` and returns only exact ID matches, with newest mail first. Prefixes such as `HN-5639` and suffixes such as `HT-56390` do not match. `auto` is opt-in: a single unambiguous ID uses identifier search, otherwise it uses semantic search. Multiple IDs and common negation phrases use semantic search; this is a heuristic and does not implement Boolean queries. An unknown exact ID returns zero matches. Natural-language no-answer handling still requires a calibrated `min_similarity`.

Fresh indexes keep folder, sender and received date in the vector table for native prefiltering. Legacy indexes retain bounded ID-shard filtering until `rebuild-vectors` upgrades them without an API call. SQLite remains authoritative and the matching vector version is published with each metadata/content change. Search batches metadata and thread counts, omits full bodies, and does not initialize schemas on the read path.

Set `search_emails(group_by_thread=true)` to keep the strongest match per store/conversation. `thread_mail_count` includes all cached mail in that conversation, even outside search filters. Call `get_mail_thread(mail_id=..., before=3, after=3)` for chronological context around a result. Empty conversation IDs remain separate emails. Context is limited to indexed mail, with bounded body output; no additional embedding request is needed.

## Configuration

Only the model and API credentials are required. Other settings have defaults:

| Setting | Default |
| --- | --- |
| Mail scope | All dates and connected mail folders; system/search folders excluded |
| Database | `%USERPROFILE%\Documents\Outlookファイル\outlook-rag\<model>-<dimensions>d-<settings-id>` |
| Embedding requests | Up to 8 chunks per batch, 2 requests concurrently |
| Sync limits | 200 mail attempts, 2,000 new metadata rows, 30 seconds |
| Vector dimensions | Up to 1,024 for Qwen3-Embedding; full dimensions for other models |

Qwen3-Embedding requests reduced dimensions from the API when possible, falling back to local truncation if the route rejects or ignores the request. Stored vector dimensions do not change. Backend setup is described in [VLLM-TUNING.md](VLLM-TUNING.md).

New folders include a Windows-safe model name and stored dimensions; the trailing settings ID separates different embedding/chunk settings. Existing hash-only default folders in Documents or Local AppData remain in use until explicitly moved. `OUTLOOK_RAG_DATA_DIR` always overrides the default. The database location does not determine where Outlook reads PST/OST files; connected stores are discovered through Outlook.

Changing the embedding model requires re-indexing. With the default database location, a separate index is selected automatically. If you set `OUTLOOK_RAG_DATA_DIR`, use a new directory for the new model.

See [Configuration](CONFIGURATION.md) for optional settings and [opencode.example.json](opencode.example.json) for a ready-to-edit configuration.

## Limitations and data handling

The server does not modify Outlook mail. Email text and vectors are stored locally; cleaned text and queries are sent to your configured embedding API.

Only indexed mail is searchable. Attachments, New Outlook, and disconnected PST files are not supported. Deleted or moved mail is removed from the local index by `sync_emails` with `reconcile=true` after a complete scan.

## Development

```powershell
git clone https://github.com/camucamulemon7/outlook-rag.git
cd outlook-rag
uv run --frozen python -m unittest discover -s tests -p "test_*.py" -v
```

Evaluate private labeled queries with `uv run --frozen python tests/evaluate_search.py config.json labels.json --ann`. Keep real-mail labels outside Git. It reports cold runs with query caching disabled, explicitly primed warm runs, Hit@1/5/10, MRR, labeled recall and no-answer rate. Empty target lists label no-answer queries; `--min-similarity` tests a threshold. `--ann` separately measures chunk Recall@10 against exact cosine search, excluding embedding time. Configure `search_nprobes` and `search_ef` for the latency/recall tradeoff, or `search_exact=true` for a ground truth baseline. ANN defaults remain the SDK defaults; measure representative queries before tuning.

Search expands its initial chunk pool when mail/thread deduplication leaves too few results, up to 5,000 candidates per source and a five-second cooperative budget checked between rounds. Results report the final `candidate_limit` and whether a count/time limit stopped expansion.

The CLI exposes the same maintenance operations as MCP. Read-only cached commands and job cancellation also work without provider credentials or API access:

```powershell
uv run --frozen outlook-rag job start --folders inbox sent --since-days 30
uv run --frozen outlook-rag job start --folders auto --since-days 0
uv run --frozen outlook-rag job status
uv run --frozen outlook-rag job cancel <job_id>
uv run --frozen outlook-rag failures list --folder inbox --limit 50 --offset 0
uv run --frozen outlook-rag failures retry <mail_id> --max-seconds 30
uv run --frozen outlook-rag progress
uv run --frozen outlook-rag thread <mail_id> --before 3 --after 3
uv run --frozen outlook-rag search "shipping changes" --folder inbox --min-similarity 0.6
uv run --frozen outlook-rag search "HT-5639" --mode identifier
uv run --frozen outlook-rag doctor
uv run --frozen outlook-rag backup-index C:\backups\outlook-rag-20261013
uv run --frozen outlook-rag rebuild-vectors
uv run --frozen outlook-rag restore-index C:\backups\outlook-rag-20261013 C:\indexes\outlook-rag-restored
```

SQLite (`metadata.sqlite`) holds mail, chunks, lexical indexes, embedding caches and publication state. LanceDB (`vectors/`) is a derived index. Both are embedded local databases; no DB service is required. `doctor` checks SQLite, chunk/vector IDs, lexical coverage and cached embeddings. `backup-index` captures the SQLite backup and its vector generation under the writer/reader locks. Backups contain private mail; choose an appropriately protected location outside `data_dir`. Backup and restore destinations must be new directories. Restore verifies the pair and never replaces an existing index; set `data_dir` to the reported restored directory to use it.

`rebuild-vectors` reconstructs vector metadata, ANN (from 256 chunks), lexical and identifier indexes from cached chunks/embeddings, then atomically publishes the new table. It requires no API key or Outlook connection and leaves the existing publication intact on failure. Missing or invalid cached embeddings require a re-sync; it does not guess replacements. Keep a paired backup before maintenance. Old vector generations remain until `maintain_index` runs; rebuilding temporarily requires space for another vector index. Maintenance removes unused generations while protecting readers. Rebuild/backup/restore/doctor are CLI commands; the MCP tool count remains fourteen.

Use `outlook-rag --help`, `outlook-rag job --help` and `outlook-rag failures --help` for options. Commands return JSON; no command starts MCP stdio.

The offline unit tests also run on macOS and Linux using synthetic COM imports. They do not validate a real Outlook profile or Windows worker launch; those integration checks require Windows.

On Windows with Outlook Classic, run `uv run --frozen python tests/verify_outlook_metadata.py --folder inbox --max-rows 65` to compare GetArray metadata with GetNextRow on the same table. Choose a populated folder; other folder paths are accepted. This opt-in check reads only IDs/classes/timestamps and outputs counts, without email bodies, identifying values, API calls or index writes.

GitHub Actions runs the complete offline unit suite on every push and pull request, on Ubuntu 24.04 and Windows Server 2025 with Python 3.11, 3.12 and 3.13. Dependencies come from `uv.lock` via `uv sync --frozen`; the test command then runs with uv offline. A separate quality job runs locked Ruff and Mypy. CI needs no API credentials, OpenWebUI service or user mail. Windows installs the real pywin32 package, but Outlook dispatch and background worker launch remain mocked; Linux substitutes COM imports. No tests are intentionally skipped for either OS. These checks do not validate a real Outlook profile, PST/GetTable provider behavior, WMI worker survival or screen interaction.

To test an already configured embedding API with synthetic mail, provide its existing credentials through `OUTLOOK_RAG_API_KEY`, then run `uv run --frozen python tests/verify_live_api.py --url <embedding-endpoint> --model <model-id>`. Add `--uv <uv-executable>` to include the API/ANN and MCP integration scripts. The test uses temporary indexes and a local fault proxy to verify outage recovery without stopping your provider. No real Outlook mail is read; COM/WMI remain untested by this check.

`tests/verify_live_search.py --url <embedding-endpoint> --model <model-id>` checks identifier/paraphrase regressions, eight additional natural-language queries, threshold tradeoffs, offline recovery and ANN Recall@10/100 with real API vectors. `tests/benchmark_search.py` separately measures native filtering against legacy shards on 100k synthetic mail IDs / 500k random-vector chunks by default; it needs temporary disk space above 2 GB and no API. Both use temporary indexes and print aggregate results.

## License

[MIT](LICENSE). Dependencies and embedding models have their own licenses.

Code is organized into text, Outlook adapters, search, MCP adapters, CLI and worker diagnostics modules. `app.py` retains existing public entry points and coordinates configuration, storage and sync. Windows-only imports remain deferred so cached operations work on other platforms. Source functions have annotations; COM/backend boundaries use `Any`. Ruff checks syntax, imports and correctness rules, and Mypy checks annotated source including function bodies. NumPy stubs are excluded from dependency traversal because newer stubs use Python 3.12 syntax while the source targets Python 3.11.
