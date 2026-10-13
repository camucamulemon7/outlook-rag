"""Evaluate labeled queries privately; emit aggregate metrics only.

Labels: {"queries": [{"query": "...", "expected_mail_ids": ["..."], "filters": {}}]}
Empty target lists label no-answer queries. Keep real-mail labels outside Git.
Cold runs bypass the query cache for every mode. Warm runs explicitly prime
each query before timing. ANN recall measures chunk retrieval against exact
cosine neighbors independently of labeled mail relevance and RRF ranking.
"""

import argparse
import json
import os
from pathlib import Path
import random
import statistics
import time
from unittest.mock import patch

from outlook_rag import app, search, storage


def validate(cases):
    if not cases:
        raise ValueError("Query set is empty")
    for case in cases:
        if not isinstance(case["query"], str) or not isinstance(
            case["expected_mail_ids"], list
        ):
            raise ValueError("Each query requires a string and a target ID list")
        if set(case.get("filters", {})) - {
            "sender",
            "folder",
            "since",
            "until",
            "subject",
            "recipient",
            "cursor",
        }:
            raise ValueError("Unsupported evaluation filter")


def metrics(cases, results, elapsed):
    ranks, recalls, no_answers, evidence_recalls = [], [], [], []
    for case, result in zip(cases, results):
        expected = set(case["expected_mail_ids"])
        retrieved = [row["mail_id"] for row in result["items"]]
        if expected:
            ranks.append(
                next((i for i, mid in enumerate(retrieved, 1) if mid in expected), None)
            )
            recalls.append(len(expected.intersection(retrieved)) / len(expected))
        else:
            no_answers.append(not retrieved)
        if case.get("expected_evidence_ids"):
            evidence = set(retrieved) | {
                source["mail_id"]
                for item in result["items"]
                for source in item.get("thread_context", {}).get("items", [])
            }
            truth = set(case["expected_evidence_ids"])
            evidence_recalls.append(len(truth & evidence) / len(truth))
    return dict(
        queries=len(cases),
        relevant_queries=len(ranks),
        no_answer_queries=len(no_answers),
        hit_at_1=sum(rank == 1 for rank in ranks) / len(ranks) if ranks else None,
        hit_at_5=sum(rank is not None and rank <= 5 for rank in ranks) / len(ranks)
        if ranks
        else None,
        hit_at_10=sum(rank is not None for rank in ranks) / len(ranks)
        if ranks
        else None,
        mrr_at_10=sum(1 / rank if rank else 0 for rank in ranks) / len(ranks)
        if ranks
        else None,
        labeled_recall_at_10=statistics.mean(recalls) if recalls else None,
        no_answer_rate=statistics.mean(no_answers) if no_answers else None,
        median_ms=round(statistics.median(elapsed), 2),
        p50_ms=round(statistics.median(elapsed), 2),
        p95_ms=round(statistics.quantiles(elapsed, n=100, method="inclusive")[94], 2)
        if len(elapsed) > 1
        else round(elapsed[0], 2),
        thread_evidence_recall=statistics.mean(evidence_recalls)
        if evidence_recalls
        else None,
        query_cache_hits=sum(bool(r.get("query_cache_hit")) for r in results),
        candidate_limit_warnings=sum(
            bool(
                r.get("candidate_limit_reached")
                or r.get("candidate_time_limit_reached")
            )
            for r in results
        ),
    )


def evaluate(cases, mode="hybrid", cache="warm", min_similarity=None):
    validate(cases)
    cfg = dict(app.local_config())
    cfg["query_cache_size"] = (
        0 if cache == "cold" else max(1, cfg.get("query_cache_size", 256))
    )
    results, elapsed = [], []
    with patch.object(app, "local_config", return_value=cfg):
        for case in cases:
            options = dict(
                case.get("filters", {}), mode=mode, min_similarity=min_similarity
            )
            if cache == "warm":
                app.search(case["query"], limit=10, **options)
            started = time.perf_counter()
            result = app.search(case["query"], limit=10, **options)
            elapsed.append((time.perf_counter() - started) * 1000)
            results.append(result)
    return metrics(cases, results, elapsed)


def evaluate_ann(cases, k=10):
    validate(cases)
    if any(
        not case["query"].strip()
        or set(case.get("filters", {})) - {"sender", "folder", "since", "until"}
        for case in cases
    ):
        raise ValueError("ANN evaluation requires text and vector-supported filters")
    cfg = app.local_config()
    sql, table = storage.open_index(cfg)
    recalls, timings, exact_timings = [], [], []
    grouped = {"filtered": [], "unfiltered": []}
    state = {}
    try:
        for number, case in enumerate(cases):
            vector, _ = app.cached_query_vector(case["query"], sql, cfg)
            filters = {
                name: case.get("filters", {}).get(name)
                for name in ("folder", "sender", "since", "until")
            }
            with app.indexed_snapshot(sql, table, cfg) as table:
                state = dict(
                    table=storage.table_name(sql),
                    version=table.version,
                    coverage=storage.vector_status(table),
                    metadata_prefilter=storage.metadata_ready(sql, table),
                )
                condition, params = search.sql_filter(**filters)
                native = storage.metadata_ready(sql, table)
                predicate = search.metadata_predicate(**filters) if native else None
                allowed = (
                    [
                        r[0]
                        for r in sql.execute(
                            f"SELECT e.id FROM emails e WHERE {condition}", params
                        )
                    ]
                    if condition != "1=1" and not native
                    else None
                )
                found = {}
                for exact in [False, True] if number % 2 == 0 else [True, False]:
                    started = time.perf_counter()
                    rows = search.vector_candidates(
                        table,
                        vector,
                        k,
                        allowed,
                        predicate=predicate,
                        cfg=dict(cfg, search_exact=exact),
                    )
                    (exact_timings if exact else timings).append(
                        (time.perf_counter() - started) * 1000
                    )
                    found[exact] = {row["id"] for row in rows}
                if found[True]:
                    recall = len(found[False] & found[True]) / len(found[True])
                    recalls.append(recall)
                    grouped["filtered" if condition != "1=1" else "unfiltered"].append(
                        recall
                    )
        return dict(
            queries=len(cases),
            queries_with_neighbors=len(recalls),
            k=k,
            chunk_recall=statistics.mean(recalls) if recalls else None,
            ann_median_ms=round(statistics.median(timings), 2),
            exact_median_ms=round(statistics.median(exact_timings), 2),
            search_nprobes=cfg.get("search_nprobes"),
            search_ef=cfg.get("search_ef"),
            publication=state,
            by_filter={
                name: dict(
                    queries=len(values),
                    chunk_recall=statistics.mean(values) if values else None,
                )
                for name, values in grouped.items()
            },
        )
    finally:
        sql.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("config")
    parser.add_argument("labels")
    parser.add_argument(
        "--modes",
        nargs="+",
        choices=("semantic", "hybrid", "lexical", "auto", "thread", "metadata"),
        default=["semantic", "hybrid", "lexical", "auto"],
    )
    parser.add_argument("--min-similarity", type=float)
    parser.add_argument(
        "--ann",
        action="store_true",
        help="Compare ANN chunk neighbors to exact cosine search",
    )
    args = parser.parse_args()
    os.environ["OUTLOOK_RAG_CONFIG"] = str(Path(args.config).resolve())
    cases = json.loads(Path(args.labels).read_text(encoding="utf-8-sig"))["queries"]
    validate(cases)
    runs = [(mode, cache) for mode in args.modes for cache in ("cold", "warm")]
    random.Random(0).shuffle(runs)
    report = {mode: {} for mode in args.modes}
    for mode, cache in runs:
        report[mode][cache] = evaluate(cases, mode, cache, args.min_similarity)
    if args.ann:
        report["ann_vs_exact"] = evaluate_ann(cases)
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
