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
