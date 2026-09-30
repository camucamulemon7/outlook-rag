"""Integration checks against the configured real embedding API, isolated mail index.

Run: uv run --project ... python tests/verify.py <runtime-config.json>
No real mail is written to the fixture index or printed.
"""
import json
import os
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from outlook_rag import app


def main():
    cfg = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8-sig"))
    scratch_root = Path(cfg["data_dir"]).parent
    scratch_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="outlook-rag-test-", dir=scratch_root, ignore_cleanup_errors=True) as directory:
        cfg["data_dir"] = str(Path(directory) / "data")
        config_path = Path(directory) / "config.json"
        config_path.write_text(json.dumps(cfg), encoding="utf-8")
        os.environ["OUTLOOK_RAG_CONFIG"] = str(config_path)
        sql, table = app.open_store(cfg)
        subjects = ["製品出荷の予定変更", "社内懇親会のご案内", "請求書の送付"]
        bodies = ["部材の到着が遅れているため、当初予定していた出荷日を来月へ延期します。",
                  "金曜日に歓迎会を開催します。参加希望者は総務までご連絡ください。",
                  "今月分の請求書を添付しました。振込期限は月末です。"]
        for index, (subject, body) in enumerate(zip(subjects, bodies)):
            app.upsert_mail(dict(id=str(index) * 64, entry_id=f"fixture-{index}", store_id="fixture-store",
                                 folder="test", subject=subject, body=body,
                                 sender=f"sender{index}@example.invalid", recipients="fixture@example.invalid",
                                 received=f"2026-09-{20+index}T12:00:00+09:00", modified="1", conversation_id=str(index)), sql, table, cfg)
        query = "納期が遅れる可能性がある案件"
        result = app.search(query, hybrid=False)
        assert result["items"][0]["entry_id"] == "fixture-0", result
        hybrid = app.search("出荷の延期", hybrid=True)
        assert hybrid["items"][0]["entry_id"] == "fixture-0"
        filtered = app.search(query, sender="sender1", since="2026-09-21", until="2026-09-21")
        assert len(filtered["items"]) == 1 and filtered["items"][0]["entry_id"] == "fixture-1"
        assert app.search(query, folder="missing")["count"] == 0
        assert app.get_mail("0" * 64)["body"] == bodies[0]
        assert app.get_mail("0" * 64, 5)["body_truncated"] is True
        # Bounded Outlook backfill, unchanged bodies, incremental state and reconciliation.
        reads = []
        class Mail:
            Class = 43
            Subject = "納期のお知らせ"
            SenderEmailAddress = "fixture@example.invalid"
            To = "fixture@example.invalid"
            ConversationID = "fixture"
            ReceivedTime = datetime.now(timezone.utc)
            LastModificationTime = datetime.now(timezone.utc)
            def __init__(self, entry):
                self.EntryID = entry
            @property
            def Body(self):
                reads.append(self.EntryID)
                return "部材の到着が遅れています。"
        class Items(list):
            def Restrict(self, expression):
                return self
            def Sort(self, *args):
                pass
        folder = SimpleNamespace(EntryID="test-folder", StoreID="test-store", Items=Items([Mail("one"), Mail("two")]))
        namespace = SimpleNamespace(GetDefaultFolder=lambda _: folder)
        outlook = SimpleNamespace(GetNamespace=lambda _: namespace)
        with patch("win32com.client.Dispatch", return_value=outlook):
            first = app.sync(["inbox"], max_emails=1)
            assert not first["folders"][0]["window_complete"]
            second = app.sync(["inbox"], max_emails=1)
            assert second["folders"][0]["window_complete"] and len(reads) == 2
            third = app.sync(["inbox"], max_emails=1)
            assert third["folders"][0]["incremental"] and third["folders"][0]["indexed"] == 0
            assert len(reads) == 2
            folder.Items.pop(0)
            reconciled = app.sync(["inbox"], reconcile=True)
            assert reconciled["folders"][0]["removed_local"] == 1
        # Remove the remaining COM fixture from the isolated index.
        fresh_sql, table = app.open_store(cfg)
        fresh_sql.close()
        remaining = sql.execute("SELECT id FROM emails WHERE folder='inbox'").fetchall()
        for row in remaining:
            app.remove_mail(row[0], sql, table)
        # LanceDB handles pin a read version; refresh after separate sync connections.
        fresh_sql, table = app.open_store(cfg)
        fresh_sql.close()
        long_body = "".join(f"unique{i:05d}。\n" for i in range(700))
        parts = app.chunks("長文", long_body)
        assert len(parts) > 1 and "unique00699" in parts[-1]
        assert all(f"unique{i:05d}" in "".join(parts) for i in range(700))
        # Verify vector index construction on real 4096-dimensional vectors.
        sample = table.search().limit(1).to_list()[0]["vector"]
        table.add([{"id": f"ann-{i}", "email_id": "0" * 64, "vector": sample} for i in range(300)])
        optimized = app.optimize()
        assert optimized["indexed"] and optimized["indexes"]
        assert table.search(sample).distance_type("cosine").limit(1).to_list()
        app.remove_mail("0" * 64, sql, table)
        assert not sql.execute("SELECT 1 FROM emails WHERE id=?", ("0" * 64,)).fetchone()
        assert table.count_rows() == 2, [(row["id"], row["email_id"]) for row in table.search().to_list()]
        sql.close()
        changed = dict(cfg, dimensions=1024)
        try:
            app.open_store(changed)
        except ValueError:
            pass
        else:
            raise AssertionError("Mixed embedding dimensions were accepted")
        print(json.dumps({"passed": True, "checks": ["real API semantic ranking", "Japanese hybrid search",
              "sender/date/folder prefilters", "original body and truncation", "long-email coverage",
              "bounded Outlook backfill", "unchanged bodies skipped", "incremental sync", "deleted-mail reconciliation",
              f"{app.vector_dimensions(cfg)}-dimensional ANN index", "local deletion consistency", "embedding identity guard"],
              "semantic_search_ms": result["elapsed_ms"], "hybrid_search_ms": hybrid["elapsed_ms"]}))


if __name__ == "__main__":
    main()
