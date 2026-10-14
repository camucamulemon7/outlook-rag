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
