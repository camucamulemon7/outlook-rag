"""Synthetic search/maintenance check against a real embedding provider.

Run with OUTLOOK_RAG_API_KEY in the environment, never command arguments.
All indexes and backups are temporary; no Outlook account is accessed.
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import tempfile
from unittest.mock import patch

from lancedb.index import IvfHnswFlat

from outlook_rag import app, search, storage
from array import array
from evaluate_search import evaluate, evaluate_ann

DOCUMENTS = [
    (
        "backup",
        "バックアップからの復元完了",
        "誤って消したファイルをバックアップから復旧しました。共有ドライブで参照できます。",
    ),
    (
        "refund",
        "二重支払いの返金",
        "代金が二重に入金されたため、重複分をお客様に払い戻します。返金先口座を確認してください。",
    ),
    (
        "flight",
        "航空券予約番号 HN-5639",
        "札幌への出張航空券を予約しました。予約番号は HN-5639、出発時刻は午前9時です。",
    ),
    (
        "hotel",
        "出張のホテル予約",
        "札幌駅近くの宿泊施設を手配しました。チェックインは10月20日、予約番号は HT-5639 です。",
    ),
    ("invoice", "請求書 INV-7391", "請求番号 INV-7391 の支払期限は11月20日です。"),
    (
        "invoice_other",
        "請求書 INV-7392",
        "請求番号 INV-7392 の支払期限は12月10日です。",
    ),
    (
        "product",
        "型番 ZX-82 の見積",
        "型番 ZX-82 の販売価格は9万円です。納期は2週間です。",
    ),
    (
        "product_other",
        "型番 ZX-820 の見積",
        "型番 ZX-820 の販売価格は20万円です。納期は4週間です。",
    ),
    (
        "venue",
        "セミナー会場の変更",
        "11月の勉強会は参加者増加のため、第一会議室から別館の大ホールへ移します。",
    ),
    (
        "delivery",
        "配送先住所の訂正",
        "注文した荷物の宛先に誤りがありました。発送前に届け先を新しい住所へ変更してください。",
    ),
    (
        "warranty",
        "修理保証のご案内",
        "購入したプリンターが故障しました。保証期間中なので修理費用は無料です。",
    ),
    (
        "insurance",
        "出張保険の加入手続き",
        "海外出張に備え、旅行中の病気やけがを補償する保険へ加入してください。",
    ),
    (
        "leave",
        "育児休業の申請書類",
        "子供が生まれた社員は、育児休業の申請書を人事部へ提出してください。",
    ),
    (
        "internet",
        "社内インターネット障害",
        "社内から外部のウェブサイトを開けません。通信回線の設備故障が原因で、午後に復旧予定です。",
    ),
    (
        "meeting",
        "定例会議の開始時刻変更",
        "水曜日の定例会議を午前10時から午後2時へ変更します。参加者は予定表を更新してください。",
    ),
    (
        "password",
        "アカウントのパスワード再設定",
        "ログイン用パスワードを忘れた場合は、本人確認後に新しいパスワードを設定してください。",
    ),
]
CASES = [
    ("削除した資料を元に戻せた連絡", "backup"),
    ("余分に受け取った代金を戻す案内", "refund"),
    ("HT-5639 の予約", "hotel"),
    ("HN-5639 の出発時刻", "flight"),
    ("INV-7391 の支払い", "invoice"),
    ("zx-82 の価格", "product"),
    ("勉強会の場所が変わるお知らせ", "venue"),
    ("送付先が間違っていた荷物の対応", "delivery"),
    ("壊れた印刷機を無償で直してもらえる連絡", "warranty"),
    ("外国で働く期間のけがに備える手続き", "insurance"),
    ("赤ちゃんの世話のために休む際の書類", "leave"),
    ("会社からウェブが見られない原因と復旧予定", "internet"),
    ("水曜の打合せが午後に移動する案内", "meeting"),
    ("ログインの秘密の文字列を忘れたときの対応", "password"),
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", required=True)
    parser.add_argument("--model", required=True)
    args = parser.parse_args()
    key = os.environ.get("OUTLOOK_RAG_API_KEY")
    if not key:
        raise ValueError("Set OUTLOOK_RAG_API_KEY")
    with tempfile.TemporaryDirectory(prefix="outlook-rag-live-search-") as temporary:
        root = Path(temporary)
        cfg = dict(
            model=args.model,
            embedding_url=args.url,
            data_dir=str(root / "index"),
            warmup_on_start=False,
        )
        cfg["dimensions"] = app.detect_dimensions(cfg)
        path = root / "config.json"
        path.write_text(json.dumps(cfg), encoding="utf-8")
        env = {
            name: value
            for name, value in os.environ.items()
            if not name.startswith("OUTLOOK_RAG_")
        }
        env.update(OUTLOOK_RAG_CONFIG=str(path), OUTLOOK_RAG_API_KEY=key)
        with patch.dict(os.environ, env, clear=True):
            cfg = app.config()
            sql, table = app.open_store(cfg)
            try:
                ids = {
                    key: hashlib.sha256(key.encode()).hexdigest()
                    for key, _, _ in DOCUMENTS
                }
                mails = [
                    dict(
                        id=ids[key],
                        entry_id=key,
                        store_id="synthetic",
                        folder="test",
                        subject=subject,
                        body=body,
                        sender="synthetic@example.invalid",
                        received="2026-10-13T12:00:00+09:00",
                    )
                    for key, subject, body in DOCUMENTS
                ]
                app.upsert_mails(mails, sql, table, cfg)
                labels = [
                    dict(query=query, expected_mail_ids=[ids[key]])
                    for query, key in CASES
                ]
                labels.append(dict(query="HT-9999 の予約", expected_mail_ids=[]))
                report = dict(
                    provider="configured_real_api",
                    data="synthetic_only",
                    emails=len(mails),
                    queries=len(labels),
                    regression_queries=6,
                    held_out_queries=8,
                    native_dimensions=cfg["dimensions"],
                    stored_dimensions=app.vector_dimensions(cfg),
                )
                report["search"] = {
                    mode: {
                        cache: evaluate(labels, mode, cache)
                        for cache in ("cold", "warm")
                    }
                    for mode in ("hybrid", "semantic", "lexical", "auto")
                }
                assert report["search"]["auto"]["warm"]["hit_at_5"] == 1
                assert report["search"]["auto"]["warm"]["no_answer_rate"] == 1
                assert (
                    app.search("HT-5639 の予約", mode="auto")["items"][0]["mail_id"]
                    == ids["hotel"]
                )
                assert app.search("ZX-82", mode="identifier")["count"] == 1
                assert (
                    app.search("削除した資料を元に戻せた連絡", mode="semantic")[
                        "items"
                    ][0]["mail_id"]
                    == ids["backup"]
                )
                # Threshold tradeoffs are reported for this small corpus, never
                # selected as global defaults from these synthetic observations.
                no_answer = [
                    dict(query="木星の衛星探査機の燃料補給計画", expected_mail_ids=[])
                ]
                report["thresholds"] = {
                    str(threshold): evaluate(
                        labels[:-1] + no_answer, "auto", "warm", threshold
                    )
                    for threshold in (None, 0.5, 0.6)
                }
                before = app.search("勉強会の場所が変わるお知らせ", mode="semantic")
                offline_rebuild = storage.rebuild_vectors(cfg)
                backup, restored = root / "backup", root / "restored"
                with (
                    patch.object(
                        app,
                        "embed",
                        side_effect=AssertionError("offline operation used API"),
                    ),
                    patch.object(
                        app,
                        "api_key",
                        side_effect=AssertionError("offline operation requested key"),
                    ),
                ):
                    report["doctor"] = storage.doctor(cfg)
                    report["backup"] = storage.backup_index(cfg, str(backup))[
                        "backed_up"
                    ]
                    report["restore"] = storage.restore_index(
                        cfg, str(backup), str(restored)
                    )["restored"]
                    assert app.search("HT-5639", mode="identifier")["count"] == 1
                    assert (
                        app.search("勉強会の場所が変わるお知らせ", mode="semantic")[
                            "items"
                        ][0]["mail_id"]
                        == before["items"][0]["mail_id"]
                    )
                report["offline_rebuild_embedding_requests"] = offline_rebuild[
                    "embedding_requests"
                ]
                report["neighbors_vs_exact"] = evaluate_ann(labels)
                report["neighbors_vs_exact"]["note"] = (
                    "Small corpus uses exact search; ANN tuning is measured separately on a larger index."
                )
                # A second, larger corpus uses genuine provider vectors and
                # several IVF partitions rather than duplicated toy vectors.
                sql2, table2 = app.open_store(cfg)
                try:
                    extra = [
                        dict(
                            id=hashlib.sha256(f"ann-{i}".encode()).hexdigest(),
                            entry_id=f"ann-{i}",
                            store_id="synthetic",
                            folder="test" if i % 2 else "other",
                            subject=f"部門{i}の月次報告",
                            body=f"部門{i}の業務報告を共有します。{DOCUMENTS[i % len(DOCUMENTS)][2]}",
                            sender="synthetic@example.invalid",
                            received="2026-10-12T12:00:00+09:00",
                        )
                        for i in range(256)
                    ]
                    app.upsert_mails(extra, sql2, table2, cfg)
                    with app.vector_publication(sql2, table2):
                        table2.create_index(
                            "vector",
                            config=IvfHnswFlat(
                                distance_type="cosine", num_partitions=8
                            ),
                        )
                        with sql2:
                            app.publish_vectors(sql2, table2)
                    # A fresh row is added after ANN construction and must be read.
                    fresh = dict(
                        extra[0],
                        id=hashlib.sha256(b"fresh-ann").hexdigest(),
                        entry_id="fresh-ann",
                        body="会議の予定番号 MT-8872 は明日の午後3時です。",
                        folder="test",
                    )
                    app.upsert_mails([fresh], sql2, table2, cfg)
                    ann_labels = [dict(case) for case in labels[:10]]
                    ann_labels += [
                        dict(case, filters=dict(folder="test")) for case in labels[:10]
                    ]
                    report["real_api_ann"] = dict(
                        rows=table2.count_rows(), partitions=8, results={}
                    )
                    for setting, controls in (
                        ("sdk_default", {}),
                        ("all_partitions_ef200", dict(search_nprobes=8, search_ef=200)),
                    ):
                        with patch.object(
                            app, "local_config", return_value=dict(cfg, **controls)
                        ):
                            report["real_api_ann"]["results"][setting] = {
                                str(k): evaluate_ann(ann_labels, k) for k in (10, 100)
                            }
                    assert app.search("MT-8872", mode="identifier")["count"] == 1
                    text = sql2.execute(
                        "SELECT text FROM chunks WHERE email_id=?", (fresh["id"],)
                    ).fetchone()[0]
                    values = array("f")
                    values.frombytes(
                        sql2.execute(
                            "SELECT vector FROM embedding_cache WHERE hash=?",
                            (hashlib.sha256(text.encode()).hexdigest(),),
                        ).fetchone()[0]
                    )
                    with app.indexed_snapshot(sql2, table2, cfg) as table2:
                        assert search.vector_candidates(
                            table2, list(values), 10, None, cfg=cfg
                        )[0]["id"].startswith(fresh["id"] + ":")
                    report["real_api_ann"]["new_rows_included"] = True
                finally:
                    sql2.close()
                assert report["doctor"]["healthy"]
                report["passed"] = True
                print(json.dumps(report, ensure_ascii=False))
            finally:
                sql.close()


if __name__ == "__main__":
    main()
