# Windows 11 resilience verification — 2026-10-14

Verified on a physical Windows 11 Pro machine, build 26200, with Outlook
Classic, three connected stores and real pywin32 312. The starting source was
the #57 fix merged in PR #58 (`8c76ee2`). Follow-up changes are version 0.9.4.
No Outlook mail was modified. Real-mail vectors were created only in temporary
indexes, using a synthetic HTTP embedding endpoint on loopback. The report and
probe output contain aggregate counts, never mail bodies/IDs or credentials.

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

This verifies Windows process cleanup, COM/metadata and index lifecycle. The
loopback vectors do not establish semantic relevance with the user's production
embedding model. OpenCode's runtime API verifies its actual Desktop backend
connection lifecycle; the search/error/cancellation assertions use actual MCP
stdio clients, not a model-generated Desktop conversation. Hardware failure,
every possible Outlook provider and arbitrary million-mail workloads are not
covered by this sample.

At the start, OpenCode's config pinned an older Outlook-RAG Git commit. After
the user restored OpenCode, the current config and running backend had no
Outlook-RAG registration. Tests added and removed a temporary runtime server;
they did not install it into the user's permanent config. Version 0.9.4 needs
the follow-up commit/PR or its tag after publication.
