# Local thread-aware retrieval measurements

Adopted changes are tracked in #37 (SQL attribute lists), #38 (incremental thread excerpts and original evidence) and #39 (opt-in paragraph/forward handling). The existing SQLite + FTS5 + LanceDB design is retained to meet the requirement for local operation without extra persistent services. BM25/vector/RRF already existed. PostgreSQL, OpenSearch, generated summaries/facts, GraphRAG, Graph delta sync and a new reranker are deferred: they add deployment/model costs or apply to a different source API than the current Outlook COM/PST reader.

## Method

On 2026-10-14, the benchmark used the real local OpenWebUI 0.12.0 embedding API, `text-embedding-qwen3-embedding-4b`, native 2,560 / stored 1,024 dimensions. Host: macOS, 64 GiB RAM, SSD, Python 3.12.14, LanceDB 0.39.0. No real Outlook account or email was accessed. Credentials stayed in process memory/environment.

The same 220 synthetic mails and 100 labeled questions cover 20 cases each of paraphrases, approval/reversal history, forwarded originals, long paragraphs and attribute lists. Labels are constructed before retrieval. Fifty regression and fifty held-out topic variants share templates, so this is a diagnostic comparison rather than a diverse production evaluation. The corpus deliberately includes the known v2 forwarding-loss failure. The corpus hash and source hashes are in [aggregate data](architecture-2026-10-14.json); detailed synthetic IDs/results are reproducible with `tests/benchmark_architecture.py`.

The genuine old source at `427ee3d3793bf8ac32d446b973e136f71366526a` runs from a Git archive in a separate process, with the same frozen runtime. Current source is copied once before both ablations so edits cannot change it during a measurement. Each revision builds a fresh temporary index. The small index uses exact vector search, not ANN. The baseline uses ordinary hybrid search. The routed variant selects metadata for attribute labels and thread for history labels; other labels use hybrid. Routing is explicitly supplied by labels: accuracy of OpenCode's LLM/tool selection is not measured.

The measured current package is preserved in source commit `03c2374`. A subsequent CLI-only change allows the empty metadata query to be omitted on PowerShell; it does not change the measured retrieval functions.

Cold runs bypass the query embedding cache. Warm runs prime the query and measure three searches, including database opening, ranking and context. P95 uses linear interpolation. OS caches are retained. Aggregate cold timing has 100 observations, category timing only 20; warm timing has three repeats per question. Revision runs are sequential, so small latency differences may reflect system/API variation, not a statistically established performance change.

## Accuracy and latency

Mail Recall@10 counts the labeled target IDs among the ten primary results. MRR@10 uses the first labeled target's position. Thread evidence recall counts labeled history sources in primary results or their attached originals, separately from message rank. The two have different retrieval units and budgets.

| Configuration | Mail Recall@10 | MRR@10 | Cold P50 / P95 | Warm P50 / P95 |
| --- | ---: | ---: | ---: | ---: |
| Old hybrid, cleaning v2 | 62% | 0.497 | 69.77 / 74.63 ms | 5.37 / 6.97 ms |
| Current hybrid, cleaning v2 | 62% | 0.497 | 71.21 / 76.92 ms | 6.07 / 8.22 ms |
| Current routed, cleaning v2 | 77% | 0.531 | 67.65 / 75.63 ms | 4.80 / 6.83 ms |
| Current hybrid, cleaning v3 | 81% | 0.652 | 69.52 / 74.38 ms | 5.17 / 6.83 ms |
| Current routed, cleaning v3 | 97% | 0.688 | 67.94 / 72.47 ms | 4.57 / 6.06 ms |

Routed v2 adds 15 percentage points without re-embedding existing chunks; opt-in v3 adds 35 points versus baseline in this constructed set. These are retrieval improvements, not generated-answer accuracy. Held-out topic variants improve from 60% to 74% (v2 routed) / 94% (v3 routed). The unchanged hybrid/v2 accuracy is a compatibility control, not evidence that ordinary searches automatically gain the routed results.

For the twenty history questions, latest-original evidence hits improve from 10% to 85%, and history-source recall from 70% to 85%. The primary final-mail Hit@1 remains **0%**: candidates/context improve, but direct ranking of the final mail is still weak. Three held-out history cases (`thread-14`, `thread-16`, `thread-17`) still miss the expected latest source. The highest ranked result alone must not be treated as a final decision. All returned context bodies match original prefixes and stay in the same store/conversation, with zero observed chronology errors. This verifies retrieved provenance, not an LLM's citations or understanding of reversals.

Forwarded-original Recall@10 improves from 0% to 100% only with v3; it retains facts that v2 removed. Paragraph-case Recall@10 stays 100%, while MRR drops from 0.468 to 0.452. Paragraph packing alone has no demonstrated ranking improvement here, so version 2 remains the default. Version 3 requires an explicit new index/full sync.

Metadata lists retain 100% labeled recall, improve exact-list precision from 50% to 100% by omitting unrelated-subject mails, and use zero embedding calls. Their v3 cold P50/P95 is 0.40/0.52 ms. Thread originals increase response size: the v3 history category averages 28,066 bytes versus 10,261 bytes for old history responses, and context bodies are capped at 12,000 characters per response. This additional evidence is a material payload/token cost, even when local retrieval time stays small. Category timings and payload sizes are retained in the JSON.

## Updating and adoption

| Operation | Old v2 | Current v2 | Current v3 |
| --- | ---: | ---: | ---: |
| Initial indexing, 220 mails | 19,025 ms | 19,025 ms | 19,976 ms |
| Unique embedded chunks | 123 | 123 | 142 |
| Embedding batches | 16 | 16 | 18 |
| Stored chunks | 240 | 240 | 240 |
| New reply, complete publication | 67.34 ms | 74.62 ms | 76.97 ms |
| New reply, thread documents updated | — | 1 | 1 |
| Thread-only full rebuild, 180 conversations | — | 34 ms | 32 ms |

The new reply's derived thread update itself takes approximately 1 ms and no extra model call; all revisions embed one new chunk. An unchanged mail makes zero embedding requests and changes zero thread fingerprints. Initial v3 indexing is about 5% slower here because retained forwarding content adds embedding work. This is not an estimate of a million-mail re-embedding duration. The derived full rebuild uses zero embedding/summary requests and is explicit; search never starts it.

To adopt on an existing matching v2 index, run `outlook-rag rebuild-threads`, then use metadata/thread modes. To adopt v3, set `OUTLOOK_RAG_TEXT_CLEANING_VERSION=3`, choose a new `data_dir` if fixed, and sync all connected stores/dates (`folders=['auto']`, `since_days=0`). Offline vector rebuilding cannot re-chunk old data into v3. Keep the old index/backup until the new sync finishes. The MCP client must load this updated server to see its routing instructions and arguments.

## Reproduce

```powershell
# API credentials must already be in OUTLOOK_RAG_API_KEY.
uv run --frozen python tests/benchmark_architecture.py --url <embedding-endpoint> --model <model> --output comparison.json
uv run --frozen python tests/benchmark_mail_search.py --emails 1000000 --chunks-per-mail 5 --dimensions 1024 --embedding-url <embedding-endpoint> --model <model> --thread-comparison --output scale.json
```

The second command uses synthetic mail and random document vectors, genuine provider query embeddings and temporary databases. It compares the old/current search algorithms on one index with primed query embeddings and interleaves modes. It requires substantial temporary disk/RAM and index-build time. It measures latency and derived thread build cost; its random vectors do not measure email relevance. Source-body matching in the first benchmark does not measure answer/citation correctness; those require a separate answer-generation evaluation with real representative questions.

## One million mails / five million chunks

An additional run built one million synthetic mails, five million 1,024D random unit vectors and 100,000 thread documents on the same host. Twenty consecutive generated mails share a conversation name; alternating stores split each into two ten-mail conversations. The vector index is IVF_HNSW_FLAT with 64 partitions, five million indexed rows and zero unindexed rows. Query embeddings come from OpenWebUI and are primed before retrieval. Three queries (two common keyword combinations and one rarer term), each repeated twice, are interleaved across old hybrid, current hybrid, current thread and metadata modes on the same database. All searches reopen SQL/Lance handles; OS caches are retained.

| Scope / mode | Median | Maximum of six samples |
| --- | ---: | ---: |
| All dates, old hybrid | 2.955 s | 7.508 s |
| All dates, current hybrid | 2.951 s | 4.384 s |
| All dates, current thread | 3.260 s | 5.001 s |
| All dates, current metadata | 1.915 ms | 8.700 ms |
| Since 2026-01-01, old hybrid | 6.100 s | 6.325 s |
| Since 2026-01-01, current hybrid | 3.523 s | 4.091 s |
| Since 2026-01-01, current thread | 2.678 s | 4.552 s |
| Since 2026-01-01, current metadata | 1.055 ms | 1.870 ms |

Across the all-date samples, thread retrieval adds approximately 0.31 s / 10% to current hybrid's median while ordinary hybrid stays approximately unchanged. These six observations are insufficient for a reliable P95 or a claim that every mailbox/query fits five seconds. The different period-query timings show substantial variability; the ordinary hybrid algorithm is essentially unchanged, so its measured timing differences must not be interpreted as a new algorithmic speedup. A date restriction does not universally reduce latency. Metadata samples list the newest ten messages with no sender/subject/recipient restrictions; rare substring filters can require much more SQL scanning.

The explicit thread-only rebuild takes 59.038 s for 100,000 conversations, with zero embedding/summary requests. This is a one-time/manual derived-index build, not work done during a search. ANN construction itself took much longer than querying. Temporary files occupied about 50 GiB and were removed after measurement. Generated texts and random document vectors make this a latency test only: it does not validate ANN relevance, Windows Outlook speed, or complete OpenCode/LLM response time. The existing ANN recall concern in #33 remains unresolved.

[Full scale samples](million-thread-2026-10-14.json) include the recorded source of an auxiliary-statistics correction: the first summary used a stale empty Lance handle, although searches reopened the published populated generation. Vector coverage was corrected from the fresh `ann_ready` statistics in the same run; timing samples are unchanged. The benchmark helper now opens the published generation for its statistics, verified on a fresh smaller ANN index.
