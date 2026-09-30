"""Evaluate a private labeled query set; print aggregate metrics only.

Usage: python tests/evaluate_search.py config.json labels.json
Labels: {"queries": [{"query": "...", "expected_mail_ids": ["..."], "filters": {}}]}
Keep label files outside the repository. Labels identify known relevant emails;
they need not be exhaustive relevance judgments. No mail content is printed.
"""
import argparse
import json
import os
import statistics
import time
from pathlib import Path

from outlook_rag import app


def evaluate(cases, hybrid=True):
    ranks, elapsed, cache_hits = [], [], 0
    for case in cases:
        expected = set(case['expected_mail_ids'])
        if not expected or not case['query'].strip():
            raise ValueError('Each query requires nonempty text and known target IDs')
        filters = case.get('filters', {})
        if set(filters) - {'sender', 'folder', 'since', 'until'}:
            raise ValueError('Unsupported evaluation filter')
        started = time.perf_counter()
        result = app.search(case['query'], limit=10, hybrid=hybrid, **filters)
        elapsed.append((time.perf_counter()-started)*1000)
        cache_hits += bool(result.get('query_cache_hit'))
        ranks.append(next((i for i,row in enumerate(result['items'],1) if row['mail_id'] in expected),None))
    return dict(queries=len(cases), hit_at_1=sum(rank==1 for rank in ranks)/len(cases),
        hit_at_5=sum(rank is not None and rank<=5 for rank in ranks)/len(cases),
        mrr_at_10=sum(1/rank if rank else 0 for rank in ranks)/len(cases),
        median_ms=round(statistics.median(elapsed)), query_cache_hits=cache_hits)


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('config')
    parser.add_argument('labels')
    args=parser.parse_args()
    os.environ['OUTLOOK_RAG_CONFIG']=str(Path(args.config).resolve())
    cases=json.loads(Path(args.labels).read_text(encoding='utf-8-sig'))['queries']
    if not cases:raise ValueError('Query set is empty')
    print(json.dumps({'semantic':evaluate(cases,False),'hybrid':evaluate(cases,True)},indent=2))


if __name__=='__main__':main()
