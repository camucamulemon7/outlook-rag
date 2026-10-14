# Optional operations in 0.9.3

## Scheduled synchronization

Default startup and search never start a sync. For periodic upkeep, install the
CLI once, complete an interactive `outlook-rag --config C:\Mail\config.json job
start --folders auto --since-days 0`, and wait for `job status` to finish. Confirm
that `status` shows the intended index path. Outlook must be running under the
same signed-in Windows user; store the API key with `set-key` for that user.

Create a Windows Task Scheduler task explicitly:

- Trigger: repeat every 15 minutes indefinitely.
- User: the same Outlook user; choose **Run only when user is logged on**.
- Program: `powershell.exe`.
- Arguments: `-NoProfile -File "C:\Tools\outlook-rag\scripts\sync-scheduled.ps1" -Config "C:\Mail\config.json" -Executable "C:\Tools\outlook-rag.exe"`.
- If the task is already running: **Do not start a new instance**.

Use the absolute installed executable path and script path. The script reuses
active jobs, starts an all-store job when idle, requests reconcile every 24 hours,
absorbs ANN tail rows after a successful job when at least 256 rows and 10% of
the index are unindexed, and runs maintenance weekly. Override the intervals
with `-ReconcileEveryHours` / `-MaintainEveryDays`. The first idle cycle after
installation reconciles. Failed/partial jobs never count as a successful
reconcile. `schedule-state.json` stores timestamps and the processed job ID inside
the index directory. A detached worker continues after the task exits.

Remove/disable the task to stop future automatic starts, and explicitly cancel an
active worker with `outlook-rag --config C:\Mail\config.json job cancel`. Choose
one scheduled task per database. There is no additional service or MCP tool.

## Attachment indexing

Install optional parsers before enabling the feature:

```powershell
uv tool install --python 3.12 "outlook-rag[attachments] @ git+https://github.com/camucamulemon7/outlook-rag.git@v0.9.4"
```

For a uvx MCP command, add `--with pypdf --with defusedxml --with psutil` before
`--from`. Set `index_attachments: true` in JSON, or
`OUTLOOK_RAG_INDEX_ATTACHMENTS=true`. It changes the index identity; automatic
placement selects a separate directory. An explicit existing DATA_DIR rejects
the mismatch, so choose a new directory and run a full sync. Default false leaves
existing text/vector identities compatible.

Supported text formats are TXT/CSV/MD (UTF-8, BOM-marked UTF-16, CP932), PDF text,
and DOCX/XLSX/PPTX XML text and stored cell values. No OCR, macro execution,
password unlocking, Excel recalculation or binary Office formats are supported.
Limits: first 10 attachments per mail, 10 MiB per file, 20 MiB expanded Office
XML, 1,000 ZIP entries, first 50 PDF pages, 100,000 extracted characters per file,
10 seconds per parser (also bounded by remaining sync time). Parser children
have a 1 GiB memory limit/monitor and are killed on timeout. Temporary files are
deleted. This is process isolation and resource limiting, not a filesystem or
network sandbox. Live Outlook SaveAsFile calls retain COM's cooperative timing.

Search hits include `attachment_name` when the selected matching chunk came
from an attachment. `get_indexed_mail` preserves the original body and returns
`attachment_indexing` with parsed-file counts and sanitized failures. Failed
attachments do not stop indexing the body. Extracted attachment text is stored
locally and sent to the configured embedding API, under the same data policy as
email text. Disabling the feature selects the original index identity.

## Legacy dates and metadata

The original 12-column `emails` table remains compatible; extra attributes use
versioned auxiliary tables created on the next writer initialization. Unchanged
messages missing those attributes are backfilled from Outlook without reading
their body or re-embedding. `sender` matches display name/address, `recipient`
matches To/CC in every search mode, `has_attachments` filters a known attachment
flag, and `folder` supports case-insensitive `*`, `?`, bracket patterns and a
trailing slash prefix. Newly populated attributes are not available from an
unconnected archive until its next scan. Message-ID deduplication defaults on;
`deduplicate=false` or CLI `--no-deduplicate` returns all copies. Metadata pages
deduplicate inside the requested scope before paging, choosing the latest
received copy. Ranked retrieval chooses the strongest matching copy. Canonical
originals retain every copy and ID.

Outlook UTC property columns are normalized with the OS timezone for the message
date. Returned timestamps include the actual offset; since/until compare the
stored local calendar date. Windows regional short-date/time formatting is used
for Jet filters. A legacy database's date cursors reset once rather than shifting
all cached times by an assumed offset. Run all-store `since_days=0` sync to reread
dates; cached content embeddings are reused. `index_status.date_validation` stays
`rescan_required` until connected coverage finishes, or while cached folders are
unavailable. Windows CI tests formatting and synthetic COM sources; a real
Outlook/PST profile still needs its own acceptance test.

## Embedding and history

Transport failures and 429/502/503/504 retry up to three attempts with jitter and
Retry-After, within one 120-second call budget or the shorter sync budget. Waits
check cancellation every 100 ms. Optional `embedding_encoding_format=base64`
decodes little-endian float32 responses and falls back to float once if the route
rejects the format. Float remains the default. Decrypted key files are cached in
memory and reread when file identity/mtime/size changes; environment keys take
precedence. No keys enter job snapshots or parser environments.

Job starts and maintenance retain the latest 20 terminal records up to 14 days,
plus the latest snapshot and active jobs; associated config/cancel/log files are
removed together. Invalid records are ignored with a content-free diagnostic.
Active jobs and the latest snapshot are preserved for recovery. Older interrupted
jobs follow terminal retention. OneDrive paths produce
a status notice; explicit/PST-selected paths are never silently moved. New
indexes without a PST use Local AppData, and existing legacy indexes are reused.

`search_nprobes`, `search_ef`, `search_exact` remain available without
re-embedding. `ann_num_partitions` optionally configures index construction
(1..4096, at least 256 chunks per partition). The default remains capped at 64;
benchmark against exact search before changing it for a large corpus. Tail rows
remain searchable until the next optimization.
