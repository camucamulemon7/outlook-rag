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
3. Call `sync_emails` again when you want to include new or changed mail. Sync does not run automatically.

For a smaller sync, call `sync_emails` with `{"max_total_emails": 10}`. After bulk indexing, call `optimize_index` to build the vector search index.

Background jobs survive MCP disconnects. Only one bulk worker runs per database; another start returns the existing job. Cancellation is cooperative, so an in-flight Outlook/API call may finish first. After a computer restart or an interrupted worker, start another job to resume saved checkpoints and cached embeddings. A `completed_with_errors` result means some mail or folders still need attention.

`sync_emails` remains available for small foreground updates. Its default limits are 200 mail attempts, 2,000 new metadata rows and a 30-second cooperative budget. Bulk jobs use the same settings and immediately start the next cycle when needed. While embedding requests run, sync reads ahead up to 16 additional rows on the Outlook thread. Set `OUTLOOK_RAG_SYNC_PREFETCH_EMAILS=0` to disable it.

Failures are retained for later retry and counted by `index_status`. Use `maintain_index` when you want to remove unused embedding caches and compact vector storage; recent versions are retained for seven days.

With an existing matching index, MCP startup, cached mail reads, index/job status and job cancellation recover dimensions locally when the embedding API is unavailable. A saved job snapshot also allows status and cancellation before the worker initializes its database. Initial setup, search and indexing still require the API; startup still requires configured API credentials. Omitting a job ID selects the most recently created job, even if an older worker updates its status later.

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

Evaluate your own labeled search queries with `tests/evaluate_search.py`; keep real-mail labels outside the repository. Candidate count and RRF weights are configurable without rebuilding the index.

The offline unit tests also run on macOS and Linux using synthetic COM imports. They do not validate a real Outlook profile or Windows worker launch; those integration checks require Windows.

To test an already configured embedding API with synthetic mail, provide its existing credentials through `OUTLOOK_RAG_API_KEY`, then run `uv run --frozen python tests/verify_live_api.py --url <embedding-endpoint> --model <model-id>`. Add `--uv <uv-executable>` to include the API/ANN and MCP integration scripts. The test uses temporary indexes and a local fault proxy to verify outage recovery without stopping your provider. No real Outlook mail is read; COM/WMI remain untested by this check.

## License

[MIT](LICENSE). Dependencies and embedding models have their own licenses.
