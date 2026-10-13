# Remaining issue validation (2026-10-14)

Local OpenWebUI `/api/v1/embeddings`, text-embedding-qwen3-embedding-4b,
native 2560 / stored 1024 dimensions. Synthetic data only; no user mail.
Reproduce with `tests/verify_live_search.py --url <endpoint> --model <model>`
and credentials supplied in the API-key environment variable.

| Retrieval check | Measurement |
| --- | --- |
| Hybrid labeled Recall@10, 14 relevant queries | 1.0000 |
| Hybrid MRR@10 | 0.9167 |
| Hybrid cold P50/P95 (includes query API) | 65.04 / 72.98 ms |
| Hybrid warm P50/P95 | 4.65 / 6.12 ms |
| ANN SDK default Recall@10/100, 20 queries | 1.0000 / 1.0000 |
| ANN SDK default P50, k=10/100 | 1.67 / 1.87 ms |
| Exact baseline P50, k=10/100 | 1.37 / 1.44 ms |
| All 8 partitions + ef=200 P50, k=10/100 | 1.70 / 1.85 ms |

The separate ANN corpus contains 273 chunks: 272 indexed into 8 IVF_HNSW_FLAT
partitions and one appended row. Ten filtered and ten unfiltered queries are
reported separately; both have full neighbor recall at k=10/100. The fresh row
is retrieved without dropping unindexed rows. This small corpus is not evidence
of recall/latency at one million messages, and no ANN exploration default is
changed. The earlier 100k random-vector experiment in #33 remains relevant:
SDK defaults are not a universal quality guarantee. Partition count is now
configurable, and scheduled maintenance absorbs a sufficiently large tail.

These are neighbor-recall checks, distinct from labeled mail relevance and
answer correctness. No-answer queries still return relative neighbors without a
calibrated similarity threshold. Eight additional natural-language queries are
included separately from the six original regression queries; the corpus remains
synthetic and small. Real-mail evaluation is still needed before model-specific
thresholds or default exploration values are selected.
