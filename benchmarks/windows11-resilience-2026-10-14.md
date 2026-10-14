# Windows 11 resilience verification — 2026-10-14

Verified on a physical Windows 11 Pro machine, build 26200, with Outlook
Classic, three connected stores and real pywin32 312. The starting source was
the #57 fix merged in PR #58 (`8c76ee2`). Follow-up changes are version 0.9.4.
No Outlook mail was modified. Initial probes created real-mail vectors in
temporary indexes using a synthetic HTTP embedding endpoint on loopback; the
production follow-up below uses the user's regular service and database.
The report and probe output contain aggregate counts, never mail bodies/IDs
or credentials.

## Findings and fixes

1. The #57 cold-start, timeout, cancellation and EOF paths passed on Windows.
   Abruptly killing the transport exposed an additional fault: its blocked
   native helper survived. Helpers now join a parent-owned Windows Job Object
   before native imports; closing its final handle kills native descendants.
   Registering the actual interpreter also avoids nested-job ordering races
   with virtualenv launchers.
2. Detached sync had only cooperative cancellation. A COM/native call that
   never returned could retain the worker and writer locks. A separate stdlib
   supervisor now enforces the cycle/status watchdog and cancellation grace.
   Native wait fault injection verified terminal status and lock release.
3. Clearing the entire fixture environment removed Windows TEMP/TMP/SystemRoot.
   As a standard user, Lance tried to create spill files in `C:\Windows` and
   panicked; subsequent fixture setup failures leaked patches and caused
   Winsock initialization failures. Tests now preserve OS essentials and
   register cleanup before fallible setup. This was a fixture failure, not
   evidence that an ordinary configured index crashed.
4. Windows Mypy rejected POSIX-only attachment resource APIs. Platform narrowing
   now passes Windows and preserves POSIX limits.

## Checks

| Check | Result |
| --- | --- |
| Complete offline unit suite, Python 3.12.13 | 242 passed, 74.4 seconds |
| Complete offline unit suite from installed wheel, Python 3.13.3 | 242 passed, 99.2 seconds |
| Actual stdio cold start | No DB, dimensions, API or credentials required |
| All seven search modes before first DB | Explicit `not_initialized`; no implicit DB write |
| SQLite-only initialization window | `initializing`; connection remains usable |
| Bad config repaired in the same session | Passed |
| COM/AttributeError, SystemExit, native exit, stdout/native handles | Sanitized tool errors/output; same-session ping/status recover |
| Blocked tool cancellation, timeout, pool saturation, startup cancellation, EOF | Passed; helper PIDs are reaped |
| Forced transport death, blocked native helper | No survivors |
| Forced transport death, native helper plus child interpreter | No survivors |
| Forced transport death inside a nested Windows Job Object | No survivors |
| Parent-death native process cleanup | Approximately 5–6 ms in the recorded probes |
| Native sync worker stuck in Win32 wait | Watchdog stops it and releases writer/worker locks |
| Native sync worker cancellation during Win32 wait | Bounded forced stop; no raw output or credentials persisted |
| Status polling while watchdog terminates native worker | Terminal timeout/cancel state wins over transient `interrupted` classification |
| Real Outlook metadata | 65 rows: GetArray/GetNextRow, item UTC and local wall time agree |
| Real Outlook bounded foreground sync | First temporary DB created; three original mails searchable |
| Real WMI + authenticated named pipe | Empty-inbox job completes |
| Real WMI job blocked at loopback API, MCP EOF | Job survives; a fresh session reads its status |
| Cancel real WMI job from fresh session | `cancelled`; committed temporary index remains readable (four mails) |
| Installed wheel through uv tool, then offline MCP startup | Seven cold modes and 12 tool calls pass; config repair and failed job start preserve session |
| Running OpenCode Desktop backend 2.0.29 | Three add/disconnect/reconnect/remove cycles; connects in 544–594 ms; no fixture processes remain; backend stays alive |
| Ruff, Windows Mypy, scheduled-sync PowerShell fixtures | Passed |
| Wheel/source build, release pins and lock consistency | Passed; dependency versions retained |

OpenCode initially had a separate startup failure on 2.0.26. Its default port
49374 fell in Windows' excluded range 49354–49453; direct binding failed with
WinError 10013, although no listener owned that port. Port 52490 bound normally.
The user restored Desktop operation and its running backend was 2.0.29 during
the successful lifecycle check. These checks distinguish OpenCode's process
lifecycle from Outlook-RAG's tool execution.

## Reproduction

```powershell
uv run --frozen --offline python -m unittest discover -s tests -p 'test_*.py' -v
uv run --frozen --offline python tests/verify_windows_resilience.py
uv run --frozen --offline python tests/verify_windows_resilience.py --real-outlook
uv build
uvx --python 3.12 --from .\dist\outlook_rag-0.9.4-py3-none-any.whl outlook-rag --version
uv run --frozen python tests/verify_installed_mcp.py -- uvx --offline --python 3.12 --from .\dist\outlook_rag-0.9.4-py3-none-any.whl outlook-rag
```

The optional real Outlook probe reads a bounded mail sample into a temporary
index and starts/cancels real WMI workers. An outer watchdog tracks only this
probe's Python processes, including detached worker PIDs, and leaves Outlook
itself running. Windows CI runs the forced-exit probe without a real mailbox;
the native sync watchdog regressions are part of the ordinary offline suite.
Both Python 3.12.13 and 3.13.3 passed the real Outlook/65-row metadata/WMI
probe and all three abrupt-exit checks. The status-polling race regression was
also rechecked after the final supervisor adjustment.

## Scope

The initial checks above verify Windows process cleanup, COM/metadata and index
lifecycle. Their loopback vectors do not establish semantic relevance with the
user's production embedding model. The production follow-up below adds the
actual embedding service and model-generated Desktop tool calls. Hardware failure,
every possible Outlook provider and arbitrary million-mail workloads are not
covered by this sample.

At the start, OpenCode's config pinned an older Outlook-RAG Git commit. After
the user restored OpenCode, the current config and running backend had no
Outlook-RAG registration. Tests added and removed a temporary runtime server;
they did not install it into the user's permanent config during the initial
checks. The subsequent production follow-up installed the tested Git commit.
Version 0.9.4 needs the follow-up commit/PR or its tag after publication.

## Production settings follow-up, 2026-10-15

The user's regular OpenAI-compatible service and
`text-embedding-qwen3-embedding-4b` model were tested on the same Windows 11
machine. Authentication and model discovery succeeded; real embeddings have
2,560 native dimensions and use the default 1,024-dimensional local storage.
The existing 8B database was retained separately: its 10,310 emails and 22,114
chunks were unchanged. A fresh 4B database was populated from real Outlook.

The permanent OpenCode config now pins this PR's tested executable commit
`a34cba3c9ca5c126afc6f0d9804828e3137ab634`, with a prewarmed offline uv tool
launch. The key was entered into a masked local dialog and saved with Windows
DPAPI for the current user. A WMI-launched worker could not see the original
LocalAppData key path in this environment. Copying the encrypted file to a
shared Documents path, and using that path in the config and future key-entry
dialog, resolved the failure. No plaintext key, private service URL, or mail
contents are included in this report.

| Production check | Result |
| --- | --- |
| First real API-backed sync | Created a fresh searchable 4B database; native 2,560 / stored 1,024 dimensions |
| Japanese query through actual MCP stdio | Semantic, hybrid, lexical, metadata and thread modes each returned three items; same-session ping passed |
| Real API query latency during sync | Semantic 1,766 ms; subsequent cached hybrid 94 ms, lexical 31 ms, metadata below 1 ms, thread 47 ms in this sample |
| Actual OpenCode 2.0.29 Desktop conversation | Usual MiMo-V2.6-Flash Free model called index_status, semantic search and sync_index status; search returned one item, with no tool error |
| Cancel active production sync | Request returned in 16 ms; worker reached cancelled; all 2,419 committed emails remained searchable |
| Restart production sync | Request returned in 609 ms; new detached job resumed and continued after MCP client EOF |
| Cold DB search from actual Desktop conversation | All seven modes returned not_initialized / count 0 / no error using the production API/model config with an empty separate data root |
| Same Desktop conversation after cold searches | Cold index_status still responded; regular index_status returned ready with 3,387 emails; the empty data root had not been created |
| Remove temporary Desktop MCP | All eight fixture processes exited; regular MCP stayed connected |
| Read a production original after restart | Cached original body returned; same-session ping passed |
| OpenCode process continuity | Backend PID unchanged throughout production search, sync cancellation/restart and cold search; no backend restart or manual MCP reconnection required |

The Desktop cold fixture was registered at the conversation's actual Documents
location. Registering a runtime server only at OpenCode's default home location
does not expose its tools to an existing Documents conversation. This is a
location-scoping behavior of OpenCode, separate from database initialization.

The final production observation had 3,823 emails, 6,455 chunks and matching
6,455 vector rows, with zero persisted mail, discovery or folder errors. The
restarted worker was running after 13 cycles. Full mailbox indexing
was still running; these checks establish first-database creation, bounded
control operations and concurrent search, not complete mailbox coverage or a
semantic relevance benchmark.

## Exact Desktop creation prompt after resetting the database

The user requested a fresh test using only
`outlook-ragでdbを作成してください` in OpenCode Desktop. The active production
sync was cancelled and all three loaded OpenCode locations disconnected their
MCP processes to release database handles; no matching process remained.
Recursive deletion was rejected by the execution environment's approval
policy, including a literal absolute-path command. Instead, the existing 4B
database directory was moved outside the configured data root to a separate
backup directory. The key and older 8B database were retained. This is a
recreation from an absent active database, not a claim that the backup was
physically deleted.

Before sending the prompt, the original database path and every metadata.sqlite
under the configured root were absent, and actual MCP index_status returned
not_initialized / zero emails. The regular MCP configuration, model, API and
credentials were unchanged. The same configuration was reconnected for the
test; no temporary cold-server configuration was used.

A fresh Desktop conversation with the user's usual MiMo-V2.6-Flash Free model
received exactly that one sentence, without extra instructions. Its observed
sequence was index_status, list_outlook_sources, tool discovery, then
sync_index(action="start", folders=["auto"], since_days=0), followed by progress
checks. It did not issue an email search before starting the job. All six
Desktop tool executions completed without an assistant/tool error, and the
Desktop response finished normally.

The new database reached ready with 440 real emails / 733 chunks, native 2,560
dimensions and 1,024-dimensional storage, with zero persisted mail failures.
The new detached worker remained running after three cycles. An uncached
Japanese semantic query returned one item in 1,662 ms (embedding 1,635 ms), and
the same-session MCP ping passed. The OpenCode backend PID stayed unchanged.
Indexing continued in the background; complete mailbox coverage was still
pending.
