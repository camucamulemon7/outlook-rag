# outlook-rag

![outlook-rag banner](assets/header.png)

Semantic and keyword search for Outlook email through MCP. Index your mail locally and search it using an OpenAI-compatible embedding API.

## Features

- Semantic search combined with Japanese keyword search (BM25 + vector search).
- All-date indexing of connected mail folders, including open archive PSTs.
- Resumable updates with time/scan limits, fair folder rotation, and failed-mail retries.
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
          "git+https://github.com/camucamulemon7/outlook-rag.git@v0.6.1",
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

For offline startup, replace `@v0.6.1` with the full commit SHA shown on GitHub, run that command online once, then add `--offline` after `uvx`. Keep the Python version and uv caches available. Your embedding API must still be running.

## Usage

1. Call `sync_emails` to index mail. Defaults are 200 mail attempts, 2,000 new metadata rows, and a 30-second cooperative time budget; repeat until folder windows are complete and no folders remain pending. In-flight API/Outlook calls can overrun the time budget.
2. Call `search_emails` with a natural-language query, then `get_indexed_mail` to read a result.
3. Call `sync_emails` again when you want to include new or changed mail. Sync does not run automatically.

For a smaller sync, pass `{"max_total_emails": 10}`. After bulk indexing, call `optimize_index` to build the vector search index.

Failures are retained for later retry and counted by `index_status`. Use `maintain_index` when you want to remove unused embedding caches and compact vector storage; recent versions are retained for seven days.

| Tool | Purpose |
| --- | --- |
| `sync_emails` | Update the local mail index |
| `search_emails` | Search by meaning and keywords, with metadata filters |
| `get_indexed_mail` | Read a cached email body |
| `index_status` | Check indexed counts and sync progress |
| `list_outlook_sources` | Inspect connected stores, folders, and PST/OST paths |
| `optimize_index` | Build a vector search index from at least 256 chunks |
| `maintain_index` | Prune unused caches and compact local vector storage |

## Configuration

Only the model and API credentials are required. Other settings have defaults:

| Setting | Default |
| --- | --- |
| Mail scope | All dates and connected mail folders; system/search folders excluded |
| Database | `%LOCALAPPDATA%\outlook-rag\<index-settings-hash>` |
| Embedding requests | Up to 8 chunks per batch, 2 requests concurrently |
| Sync limits | 200 mail attempts, 2,000 new metadata rows, 30 seconds |
| Vector dimensions | Up to 1,024 for Qwen3-Embedding; full dimensions for other models |

Qwen3-Embedding requests reduced dimensions from the API when possible, falling back to local truncation if the route rejects or ignores the request. Stored vector dimensions do not change. Backend setup is described in [VLLM-TUNING.md](VLLM-TUNING.md).

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

## License

[MIT](LICENSE). Dependencies and embedding models have their own licenses.
