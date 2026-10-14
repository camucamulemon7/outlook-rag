# MCP startup, recovery and verification

The `outlook-rag` executable starts a lightweight stdio transport before loading
LanceDB/Arrow, Outlook/pywin32, API credentials or embedding dimensions. Search
does not start indexing. Older `python -m outlook_rag.app` invocations still work;
the installed executable and `python -m outlook_rag.entry` avoid native imports
in the transport process.

| Index state | Search and status behavior |
| --- | --- |
| No SQLite file | `not_initialized`, zero results, explicit sync guidance |
| Empty SQLite or incomplete initial schema | `initializing`, retry/status guidance |
| Identity committed, vector creation still pending | `initializing`, no vector open or API probe |
| Initialized empty store | `empty`, zero results |
| Populated store | Normal retrieval; existing identity and publication checks apply |
| Corrupt metadata, wrong identity, missing vectors with populated mail | An explicit tool error; never interpreted as proof that mail is absent |

All zero-result searches retain coverage information and `absence_confirmed:
false`. Metadata retrieval can use valid populated SQLite even if the vector
store needs repair. Reading an index never creates a database. Status of a job
that has never started also leaves the data directory untouched.

Each synchronous MCP operation runs in a reusable helper process. At most four
helpers execute concurrently. Native crashes become tool errors. Helpers use a
private JSON channel; Python output, CRT and native Windows standard handles
are redirected away from that channel. COM stays on the
helper's calling thread; database handles and COM proxies never cross the process
boundary. Existing six-attempt COM property retries are retained.

`mcp_tool_timeout_seconds` (or `OUTLOOK_RAG_MCP_TOOL_TIMEOUT_SECONDS`) sets a hard
per-call deadline of 1–300 seconds, default 30, including queue and helper startup
time. A timed-out helper is terminated. The MCP connection can still handle ping,
status and subsequent tools. Cancellation and stdio EOF also terminate active
helpers. MCP SDK 1.30+ is required for EOF cancellation. Configuration errors are
returned by tools, allowing an edited config file to be used in the same session;
changing the tool profile or deadline requires a server restart.

Ownership is tracked immediately after OS spawn, before any cancellable startup
read or write. A 20-call startup cancellation regression verifies that no helpers
remain. Cleanup covers initialization as well as running operations.
Windows helpers register their actual interpreter in a parent-owned Job Object
with `KILL_ON_JOB_CLOSE` before loading native engines. Abrupt transport death
therefore also terminates helpers and parser descendants. Registering in the
interpreter avoids a race when a virtualenv launcher starts a child before job
assignment. POSIX helper timeouts/cancellation terminate the process group.

Use `sync_index(action="plan")` before bulk indexing, then `sync_index(action="start", since_days=<approved>, confirmed=true)` for the approved scope. Planning does not call the embedding API or create an index. Foreground sync in the full
profile receives a shorter cooperative budget before the hard deadline. Long
optimization/maintenance operations can use the CLI or an increased deadline.
A forcibly stopped write retains committed data; check job/index status before
retrying. Existing detached sync jobs intentionally survive MCP disconnects;
their cancellation first checks Outlook/API boundaries. A separate stdlib
supervisor owns each detached native worker. It terminates a worker after no
cycle/status update for `max(60, sync_max_seconds + 30)` seconds, configurable
with `worker_timeout_seconds` / `OUTLOOK_RAG_WORKER_TIMEOUT_SECONDS` (1..7200).
Cancellation grants at most 10 seconds, bounded by this timeout, before forcing
a stop. Status records `forced_stop` and `reason=cancel_timeout` or
`worker_timeout`; committed mail/checkpoints are retained for a new job.
Supervisors are independent of MCP; closing a search helper never stops a
detached job, Outlook or OpenCode.

Startup no longer warms the embedding model. `warmup_on_start` in older configs
is accepted for compatibility and has no MCP startup effect. The first tool call
pays the native-library import cost once per helper; subsequent calls reuse it.

## Reproducible checks

```shell
uv run --frozen --offline python -m unittest discover -s tests -p 'test_*.py'
uv run --frozen --offline python tests/benchmark_mcp.py --baseline-root /path/to/v0.9.2
uv run --frozen --offline python tests/verify_windows_resilience.py
uv run --frozen --offline python tests/verify_windows_resilience.py --real-outlook
```

The stdio tests use real MCP clients and helper processes. They exercise no DB,
partial initialization, unavailable API/no key, invalid config repaired in the
same session, SystemExit, COM/AttributeError, native exit, Python/native stdout,
blocked calls, startup/running cancellation, four-helper saturation, concurrent status, and EOF
during a blocked call. PID checks verify helper cleanup. Windows CI uses real
pywin32 for its COM error type and dynamic property lookup regression; it does
not have a real Outlook mailbox or OpenCode Desktop session.
The opt-in Windows probe checks abrupt transport death, native descendants and
nested Job Objects under an outer watchdog. `--real-outlook` also checks actual
GetArray/GetNextRow metadata, bounded real-mail sync into a temporary index,
real WMI/named-pipe startup, continuation after MCP EOF, and cancellation from a
new session. Embeddings go only to its local synthetic HTTP endpoint; output
contains aggregate counts. Mail identifiers, bodies and credentials are omitted.
See [the Windows 11 report](benchmarks/windows11-resilience-2026-10-14.md).

On macOS/Python 3.12, a 40-mail synthetic lexical benchmark (31 calls per version,
no API) measured the following. These numbers measure transport overhead, not
million-message scalability or general retrieval accuracy.

| Measurement | v0.9.2 | Isolated MCP |
| --- | ---: | ---: |
| Handshake | 574.5 ms | 184.3 ms |
| First search after handshake | 23.7 ms | 607.8 ms |
| Warm search P50 (30 calls) | 2.8 ms | 3.6 ms |
| Warm search P95 (30 calls) | 3.1 ms | 4.0 ms |

Top-10 mail IDs were identical. First-use cost moves from transport startup to
the first helper operation; warm median overhead in this run was 0.8 ms. Actual
embedding latency, cold Windows imports and large-index search time vary.
