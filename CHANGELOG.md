# 0.9.4

- Preview initial sync periods (30/90/365 days/all time), metadata counts and estimated duration before asking the user to choose. MCP bulk start requires an approved explicit period; an explicit all/everything request remains supported without a second question.
- Estimate remaining scan time from same-index measured throughput, with a qualified initial range when unmeasured; keep unknown counts unknown and scope progress to the selected period.

- Reap Windows MCP helpers and parser descendants after abrupt transport death,
  using parent-owned Job Objects registered by the actual interpreter.
- Supervise detached native sync workers separately. Bound stalled cycles and
  cancellation, release writer locks and retain committed mail/checkpoints.
  Forced stops report `worker_timeout` / `cancel_timeout` and `forced_stop`.
- Fix Windows test environments to retain TEMP/TMP/SystemRoot and restore
  configuration patches even if fixture setup fails.
- Add real Windows COM/WMI probes, nested-job/forced-exit checks, native worker
  timeout/cancellation regressions and Windows CI coverage for helper ownership.
- Keep attachment-worker type checking portable on Windows and POSIX.

# 0.9.3

- Start MCP without resolving embedding dimensions, reading API credentials,
  warming the model, or importing native vector/Outlook libraries (#57).
- Return explicit not_initialized/initializing/empty states before the first
  vector table exists; keep actual corruption and configuration errors visible.
- Isolate synchronous tools in a reusable process pool (maximum four helpers).
  A configurable 30-second deadline, request cancellation and transport EOF
  terminate blocked helpers while preserving the MCP connection for other calls.
- Contain COM errors, SystemExit, native process exits and accidental stdout;
  return sanitized diagnostics without serializing COM objects/exceptions.
- Own helper processes from OS spawn through reaping, including cancellation
  before initialization; use a private JSON channel and redirect native handles.
- Release read-ahead COM references before CoUninitialize. Read-only job status
  no longer creates a data directory on first use.
- Close SQLite explicitly if schema/vector initialization fails, preserving a
  retryable initial state without leaking the database handle.
- Require MCP SDK 1.30+ for cancellation of in-flight handlers on transport EOF.
  Existing CLI commands and the six standard MCP tool names remain available.

Existing databases and embeddings remain compatible. MCP no longer performs
startup warmup even if an older configuration contains warmup_on_start.

# 0.9.2

- Retry required Outlook folder ID properties up to six times, with 0.2 seconds
  between attempts, when dynamic COM lookup raises AttributeError or a COM error.
- Apply retries during folder discovery, coverage, sync checkpoints, rotation and
  completion reports so a transient lookup failure does not abort bulk indexing.
- Keep the original exception after exhaustion and preserve committed mail and
  checkpoint identities for a later restart; never substitute an empty ID.
- Add regressions for failures after two sync cycles and immediately after a batch
  commit, including persistent failure followed by successful restart.

Existing databases and cached embeddings remain compatible. Upgrade and restart
the MCP client before starting a new sync job to resume saved progress.

# 0.9.1

- Fix Outlook sync failing with `Unexpected Outlook metadata table shape` when
  GetArray returns row-major Python tuples, including OST and PST stores (#53).
- Preserve compatibility with column-major arrays and validate metadata fields
  before accepting a batch, including ambiguous 4-by-4 batches and 64+4 tails.
- Exercise the reported row-major layout throughout synthetic sync/recovery tests
  and compare both layouts with the read-only Outlook metadata probe.

Existing indexes remain compatible. Upgrade the installed package and restart the
MCP client before retrying sync; a new database or embedding model is not required.

# 0.9.0

- Six standard MCP tools, with the complete compatibility profile available explicitly.
- PST-inclusive archive coverage, durable index placement and functional modules.
- Metadata lists, scoped thread history, hybrid retrieval and ANN evaluation controls.
- Responsive asynchronous MCP adapters and bounded embedding transport retries.
- UTC-aware Outlook date sources and Windows regional date-filter formatting.
- Sender display names, CC, attachment attributes and scoped Message-ID deduplication.
- Bounded job history, opt-in scheduled upkeep and isolated attachment indexing.
- Python 3.14, wheel installation tests, coverage, dependency audits and release-pin checks.

New installations use a connected PST's parent directory when there is one unique
parent; otherwise an explicit root is needed for multiple PST directories. With no
PST, new indexes use Local AppData. Existing indexes are preferred.

Existing canonical email rows and embedding identities remain compatible by default.
Legacy Outlook date cursors reset once so subsequent bounded sync rereads dates from
the source. Attachment indexing is off by default and changes index identity when
enabled, requiring a separate/rebuilt index. Text cleanup version 3 remains opt-in.
