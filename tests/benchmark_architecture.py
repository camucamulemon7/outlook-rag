"""Reproducible synthetic before/after retrieval benchmark using a real API.

No Outlook access. Credentials come only from OUTLOOK_RAG_API_KEY. The old
revision runs in a separate process from a Git archive, with the same runtime,
corpus and labels. Metadata routing is supplied by labels, not an evaluated LLM.
Keep private labels separate; this deliberately constructed corpus is diagnostic.
"""

import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import random
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from unittest.mock import patch
import zipfile


TOPICS = [
    ("PCIe", "次世代インターフェースの帯域幅"),
    ("冷却装置", "サーバーの発熱を抑える方法"),
    ("バックアップ", "消失した資料を復旧する手順"),
    ("返金", "二重入金の払い戻し"),
    ("航空券", "飛行機の予約を変更する条件"),
    ("ホテル", "出張時の宿泊手配"),
    ("会場", "参加人数に合わせた会議室の変更"),
    ("配送", "荷物の送付先の訂正"),
    ("修理保証", "故障した機器の無償修理"),
    ("旅行保険", "海外出張中のけがへの補償"),
    ("育児休業", "子供の世話のための休暇"),
    ("ネットワーク", "社内からウェブが見られない障害"),
    ("パスワード", "ログイン情報を忘れた場合の再設定"),
    ("請求書", "仕入先への支払期限の変更"),
    ("採用面接", "候補者と会う日程の調整"),
    ("備品", "事務所で使う消耗品の発注"),
    ("翻訳", "製品マニュアルの多言語化"),
    ("蓄電池", "停電時の予備電源の稼働時間"),
    ("セキュリティ", "社外からの不正アクセスへの対応"),
    ("契約", "取引先との更新条件"),
]


def synthetic_corpus():
    mails, cases = [], []

    def mail(key, **values):
        value = dict(
            id=hashlib.sha256(key.encode()).hexdigest(),
            entry_id=key,
            store_id="pst-synthetic",
            folder="Archive PST/技術",
            subject="検討結果",
            sender="synthetic@example.invalid",
            recipients="team@example.invalid",
            received="2026-10-01T12:00:00+09:00",
            modified="2026-10-01T12:00:00+09:00",
            conversation_id=key,
            body="合成メールです。",
        )
        value.update(values)
        mails.append(value)
        return value

    def case(category, number, query, targets, **values):
        cases.append(
            dict(
                case_id=f"{category}-{number}",
                category=category,
                split="regression" if number < 10 else "held_out",
                query=query,
                expected_mail_ids=[m["id"] for m in targets],
                **values,
            )
        )

    for n, (topic, paraphrase) in enumerate(TOPICS):
        sender = f"担当{n:02d}<person{n:02d}@example.invalid>"
        owner = f"thread-{n}"
        code = f"RQX-{1000 + n}"
        approval = mail(
            f"approval-{n}",
            subject=f"{topic} の条件確認",
            sender=sender,
            conversation_id=owner,
            received="2008-10-01T12:00:00+09:00",
            body=f"{topic}、{paraphrase}の件です。管理番号 {code}。当初案を承認しました。期限は11月1日です。",
        )
        withdrawal = mail(
            f"withdrawal-{n}",
            subject=f"Re: {topic} の条件確認",
            sender=sender,
            conversation_id=owner,
            received="2008-10-05T12:00:00+09:00",
            body=f"{topic}の前回承認を撤回します。再評価を実施します。期限は未定です。",
        )
        latest = mail(
            f"latest-{n}",
            subject="Re: 検討結果",
            sender=sender,
            conversation_id=owner,
            received="2008-10-08T12:00:00+09:00",
            body="再評価に合格したので再承認しました。最終期限は11月20日に変更します。これが現在有効な決定です。",
        )
        case("semantic", n, f"{paraphrase}についての最初の承認の連絡", [approval])
        case(
            "thread",
            n,
            f"{topic}は最終的にどう決まった？承認と撤回の経緯も確認したい",
            [latest],
            expected_evidence_ids=[approval["id"], withdrawal["id"], latest["id"]],
            expected_latest_id=latest["id"],
        )
        # Distinct keywords exist only in the forwarded original, never the
        # outer subject or authored cover note. v2 is known to discard them.
        forwarded_code = f"FWX-{5000 + n}"
        forward = mail(
            f"forward-{n}",
            subject="Fwd: ご参考",
            sender=sender,
            body=f"ご確認ください。\n\n-----Original Message-----\nFrom: supplier@example.invalid\nSent: 2026-09-30\nTo: team@example.invalid\nSubject: 確定仕様\n\n型番 {forwarded_code} の{topic}に関する確定仕様です。処理能力は{n + 10}単位、保証期間は{n + 1}年です。",
        )
        case("forward", n, f"{forwarded_code} の確定仕様と保証期間", [forward])
        paragraph_code = f"DPX-{7000 + n}"
        paragraphs = [
            (f"{topic}の実施計画を共有します。関連部署で準備状況を確認します。" * 16),
            (
                "検討事項の担当者と実施時期を確認してください。調整会議を開催します。"
                * 16
            ),
            f"{paragraph_code} の結論：{paraphrase}に対応するための設計を変更します。担当は開発部、締切は12月15日です。",
            (
                "実施後の評価結果は別途報告します。問題があれば運用部へ連絡してください。"
                * 16
            ),
            (
                "検証手順と記録様式は共有ドライブに保存します。定期的に進捗を更新します。"
                * 16
            ),
        ]
        long = mail(
            f"paragraph-{n}",
            subject=f"{topic} 計画書",
            sender=sender,
            body="\n\n".join(paragraphs),
        )
        case("paragraph", n, f"{paragraph_code} の変更の担当部署と締切", [long])
        attributes = []
        for k in range(4):
            row = mail(
                f"attribute-{n}-{k}",
                subject=f"{topic} 月次報告" if k < 2 else "別件のお知らせ",
                folder="Archive PST/月次",
                sender=sender,
                received=f"2026-09-{10 + k:02d}T12:00:00+09:00",
                body=f"{topic}の報告を共有します。月次確認資料です。"
                if k < 2
                else "一般のお知らせを共有します。",
            )
            if k < 2:
                attributes.append(row)
        case(
            "metadata",
            n,
            f"{topic}の月次報告を一覧にして",
            attributes,
            filters=dict(
                sender=f"person{n:02d}@",
                folder="Archive PST/月次",
                since="2026-09-01",
                until="2026-09-30",
            ),
            metadata_subject=f"{topic} 月次報告",
        )
        # Relevant topic noise in another store is intentionally not the final
        # decision and must never be added to that thread's original evidence.
        for k in range(2):
            mail(
                f"newsletter-{n}-{k}",
                store_id="inbox-synthetic",
                folder="Inbox",
                subject=f"{topic} ニュース",
                body=f"{topic}、{paraphrase}の業界動向です。一般のニュースであり個別案件の決定ではありません。",
            )
    assert len(cases) == 100
    return dict(mails=mails, queries=cases, labels_defined_before_search=True)


def percentile(values, p):
    ordered = sorted(values)
    position = (len(ordered) - 1) * p / 100
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    return round(ordered[low] + (ordered[high] - ordered[low]) * (position - low), 2)


def summarize(records):
    ranks, recall, evidence, latest = [], [], [], []
    list_precision = []
    for row in records:
        target = set(row["expected_mail_ids"])
        ids = row["retrieved_ids"]
        rank = next((i for i, mid in enumerate(ids, 1) if mid in target), None)
        ranks.append(rank)
        recall.append(len(target.intersection(ids)) / len(target))
        if row["category"] == "metadata":
            list_precision.append(
                len(target.intersection(ids)) / len(ids) if ids else 0
            )
        if "expected_evidence_ids" in row:
            truth = set(row["expected_evidence_ids"])
            evidence.append(len(truth.intersection(row["evidence_ids"])) / len(truth))
            latest.append(row["expected_latest_id"] in row["evidence_ids"])
    return dict(
        queries=len(records),
        hit_at_1=statistics.mean(r == 1 for r in ranks),
        hit_at_10=statistics.mean(r is not None for r in ranks),
        mail_recall_at_10=statistics.mean(recall),
        mrr_at_10=statistics.mean(1 / r if r else 0 for r in ranks),
        metadata_list_precision=statistics.mean(list_precision)
        if list_precision
        else None,
        thread_evidence_recall=statistics.mean(evidence) if evidence else None,
        latest_evidence_hit=statistics.mean(latest) if latest else None,
        cold_p50_ms=percentile([r["cold_ms"] for r in records], 50),
        cold_p95_ms=percentile([r["cold_ms"] for r in records], 95),
        warm_p50_ms=percentile([v for r in records for v in r["warm_ms"]], 50),
        warm_p95_ms=percentile([v for r in records for v in r["warm_ms"]], 95),
        mean_result_bytes=round(statistics.mean(r["result_bytes"] for r in records)),
        original_body_mismatches=sum(r["original_body_mismatches"] for r in records),
        chronology_errors=sum(r["chronology_errors"] for r in records),
        query_embedding_calls=sum(r["embedding_calls"] for r in records),
        cold_embedding_calls=sum(r["cold_embedding_calls"] for r in records),
        warm_embedding_calls=sum(r["warm_embedding_calls"] for r in records),
        misses=[
            r["case_id"]
            for r in records
            if not set(r["expected_mail_ids"]).intersection(r["retrieved_ids"])
        ],
    )


def run_stage(args):
    from datetime import datetime
    from outlook_rag import app

    corpus = json.loads(Path(args.corpus).read_text(encoding="utf-8"))
    root = Path(args.stage_dir)
    root.mkdir()
    cfg = dict(
        model=args.model,
        embedding_url=args.url,
        data_dir=str(root / "index"),
        dimensions=args.dimensions,
        storage_dimensions=min(1024, args.dimensions),
        text_cleaning_version=3 if args.stage == "v3" else 2,
        warmup_on_start=False,
        folders=["inbox"],
        query_cache_size=256,
    )
    config_path = root / "config.json"
    config_path.write_text(json.dumps(cfg), encoding="utf-8")
    os.environ["OUTLOOK_RAG_CONFIG"] = str(config_path)
    cfg = app.config()
    source_root = Path(app.__file__).resolve().parent
    source_digest = hashlib.sha256()
    for file in sorted(source_root.glob("*.py")):
        source_digest.update(file.name.encode())
        source_digest.update(file.read_bytes())
    sql, table = app.open_store(cfg)
    counts = dict(requests=0, inputs=0)
    embed = app.embed

    def measured_embed(texts, *positional, **options):
        # embed() recursively calls itself once to create an owned client.
        # Count only the inner invocation that actually submits provider work.
        if options.get("client") is not None:
            counts["requests"] += 1
            counts["inputs"] += len(texts)
        return embed(texts, *positional, **options)

    try:
        with patch.object(app, "embed", side_effect=measured_embed):
            started = time.perf_counter()
            indexing = app.upsert_mails(corpus["mails"], sql, table, cfg)
            indexing.update(
                wall_ms=round((time.perf_counter() - started) * 1000, 2),
                **counts,
                chunks=sql.execute("SELECT count(*) FROM chunks").fetchone()[0],
            )
            operations = (
                ["hybrid"] if args.stage == "baseline" else ["hybrid", "routed"]
            )
            reports = {}
            originals = {m["id"]: m for m in corpus["mails"]}
            for operation in operations:
                rows = []
                cases = list(corpus["queries"])
                random.Random(41).shuffle(cases)
                for number, case in enumerate(cases):
                    category = case["category"]
                    mode = (
                        (
                            "metadata"
                            if category == "metadata"
                            else "thread"
                            if category == "thread"
                            else "hybrid"
                        )
                        if operation == "routed"
                        else "hybrid"
                    )
                    query = "" if mode == "metadata" else case["query"]
                    options = dict(case.get("filters", {}), mode=mode, limit=10)
                    if mode == "metadata":
                        options["subject"] = case["metadata_subject"]
                    count_before = counts["requests"]
                    with patch.object(
                        app, "local_config", return_value=dict(cfg, query_cache_size=0)
                    ):
                        started = time.perf_counter()
                        cold = app.search(query, **options)
                        cold_ms = (time.perf_counter() - started) * 1000
                    cold_calls = counts["requests"] - count_before
                    # Prime then measure three cached runs. Timings include all
                    # retrieval/context work, not only the vector-cache lookup.
                    app.search(query, **options)
                    warm_start_calls = counts["requests"]
                    warm = []
                    for _ in range(3):
                        started = time.perf_counter()
                        result = app.search(query, **options)
                        warm.append((time.perf_counter() - started) * 1000)
                        assert mode == "metadata" or result["query_cache_hit"]
                    ids = [item["mail_id"] for item in result["items"]]
                    evidence_ids, mismatch, order_errors = set(ids), 0, 0
                    for item in result["items"]:
                        evidence = item.get("thread_context", {}).get("items", [])
                        instants = [
                            (
                                datetime.fromisoformat(e["received"]).timestamp(),
                                e["mail_id"],
                            )
                            for e in evidence
                        ]
                        order_errors += instants != sorted(instants)
                        for source in evidence:
                            evidence_ids.add(source["mail_id"])
                            original = originals[source["mail_id"]]
                            mismatch += not original["body"].startswith(source["body"])
                            mismatch += (
                                original["store_id"] != item["store_id"]
                                or original["conversation_id"]
                                != item["conversation_id"]
                            )
                    rows.append(
                        dict(
                            case,
                            mode=mode,
                            retrieved_ids=ids,
                            evidence_ids=sorted(evidence_ids),
                            cold_ms=round(cold_ms, 2),
                            warm_ms=[round(v, 2) for v in warm],
                            result_bytes=len(
                                json.dumps(result, ensure_ascii=False).encode()
                            ),
                            embedding_calls=counts["requests"] - count_before,
                            cold_embedding_calls=cold_calls,
                            warm_embedding_calls=counts["requests"] - warm_start_calls,
                            original_body_mismatches=mismatch,
                            chronology_errors=order_errors,
                        )
                    )
                    if number % 20 == 19:
                        print(
                            f"{args.stage}/{operation}: {number + 1}/100",
                            file=sys.stderr,
                            flush=True,
                        )
                reports[operation] = dict(
                    aggregate=summarize(rows),
                    by_category={
                        c: summarize([r for r in rows if r["category"] == c])
                        for c in sorted({r["category"] for r in rows})
                    },
                    by_split={
                        s: summarize([r for r in rows if r["split"] == s])
                        for s in ("regression", "held_out")
                    },
                    cases=rows,
                )
            # Compare unchanged, new reply and full derived rebuild costs. New
            # reply contains a unique text, so embedding work is counted too.
            update_cost = {}
            for name, batch in (
                ("unchanged", corpus["mails"][:1]),
                (
                    "new_reply",
                    [
                        dict(
                            corpus["mails"][2],
                            id=hashlib.sha256(b"new-reply").hexdigest(),
                            entry_id="new-reply",
                            received="2008-10-09T12:00:00+09:00",
                            body="追加の確認です。再承認は引き続き有効です。期限変更はありません。",
                        )
                    ],
                ),
            ):
                before = dict(counts)
                started = time.perf_counter()
                stats = app.upsert_mails(batch, sql, table, cfg)
                update_cost[name] = dict(
                    stats,
                    wall_ms=round((time.perf_counter() - started) * 1000, 2),
                    embedding_requests=counts["requests"] - before["requests"],
                )
            if args.stage != "baseline":
                from outlook_rag import threads

                with patch.object(
                    app,
                    "embed",
                    side_effect=AssertionError("thread rebuild used model"),
                ):
                    update_cost["thread_rebuild"] = threads.rebuild(cfg)
            value = dict(
                stage=args.stage,
                source_sha256=source_digest.hexdigest(),
                indexing=indexing,
                search=reports,
                update_cost=update_cost,
            )
            Path(args.output).write_text(
                json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
            )
    finally:
        sql.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", required=True)
    parser.add_argument("--model", required=True)
    parser.add_argument(
        "--baseline-ref", default="427ee3d3793bf8ac32d446b973e136f71366526a"
    )
    parser.add_argument("--output", required=True)
    parser.add_argument("--stage", choices=["baseline", "v2", "v3"])
    parser.add_argument("--corpus")
    parser.add_argument("--stage-dir")
    parser.add_argument("--dimensions", type=int)
    args = parser.parse_args()
    if args.stage:
        run_stage(args)
        return
    if not os.environ.get("OUTLOOK_RAG_API_KEY"):
        raise ValueError("Set OUTLOOK_RAG_API_KEY in the environment")
    from outlook_rag import app

    # One untimed warmup/probe; all revisions use the actual provider dimension.
    dimensions = app.detect_dimensions(dict(model=args.model, embedding_url=args.url))
    repo = Path(__file__).resolve().parents[1]
    archive = subprocess.check_output(
        ["git", "archive", "--format=zip", args.baseline_ref], cwd=repo
    )
    corpus = synthetic_corpus()
    encoded = json.dumps(corpus, ensure_ascii=False, sort_keys=True).encode()
    with tempfile.TemporaryDirectory(prefix="outlook-rag-architecture-") as temporary:
        root = Path(temporary)
        baseline = root / "baseline-source"
        current = root / "current-source"
        shutil.copytree(
            repo / "src", current / "src", ignore=shutil.ignore_patterns("__pycache__")
        )
        with zipfile.ZipFile(io.BytesIO(archive)) as files:
            files.extractall(baseline)
        corpus_path = root / "corpus.json"
        corpus_path.write_bytes(encoded)
        report = dict(
            baseline_ref=args.baseline_ref,
            data="synthetic_only",
            emails=len(corpus["mails"]),
            queries=len(corpus["queries"]),
            corpus_sha256=hashlib.sha256(encoded).hexdigest(),
            native_dimensions=dimensions,
            stored_dimensions=min(1024, dimensions),
            labels_defined_before_search=True,
            metadata_route="label_supplied_not_LLM_classifier",
            answer_generation_evaluated=False,
            warm_repeats=3,
            note="Diagnostic templates, 50 regression + 50 held-out topic variants; small corpus uses exact vectors. Not a production or million-mail accuracy estimate.",
            revisions={},
        )
        for stage in ("baseline", "v2", "v3"):
            env = {
                k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RAG_")
            }
            env["OUTLOOK_RAG_API_KEY"] = os.environ["OUTLOOK_RAG_API_KEY"]
            env["PYTHONPATH"] = str(
                (baseline if stage == "baseline" else current) / "src"
            )
            result_path = root / f"{stage}.json"
            subprocess.run(
                [
                    sys.executable,
                    str(Path(__file__).resolve()),
                    "--url",
                    args.url,
                    "--model",
                    args.model,
                    "--stage",
                    stage,
                    "--corpus",
                    str(corpus_path),
                    "--stage-dir",
                    str(root / stage),
                    "--output",
                    str(result_path),
                    "--dimensions",
                    str(dimensions),
                ],
                env=env,
                check=True,
            )
            report["revisions"][stage] = json.loads(
                result_path.read_text(encoding="utf-8")
            )
        Path(args.output).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(
            json.dumps(
                {k: v for k, v in report.items() if k != "revisions"},
                ensure_ascii=False,
            )
        )


if __name__ == "__main__":
    main()
