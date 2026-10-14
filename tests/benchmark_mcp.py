"""Compare real stdio transport latency on a small, API-free synthetic index.

python tests/benchmark_mcp.py --baseline-root /path/to/previous-checkout
This measures process isolation overhead, not million-message scalability.
"""
import argparse
import asyncio
import hashlib
import json
import os
from pathlib import Path
import statistics
import sys
import tempfile
import time
from unittest.mock import patch

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def measure(root, config, module, env):
    child_env = dict(env, PYTHONPATH=str(root / "src"))
    params = StdioServerParameters(command=sys.executable, args=["-m", module, "--config", str(config)], env=child_env)
    started = time.perf_counter()
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            handshake = (time.perf_counter() - started) * 1000
            times = []
            ids = None
            for _ in range(31):
                started = time.perf_counter()
                result = await session.call_tool("search_emails", {"query": "PCIe", "mode": "lexical"})
                elapsed = (time.perf_counter() - started) * 1000
                if result.isError:
                    raise ValueError("Synthetic benchmark tool failed")
                data = result.structuredContent or json.loads(result.content[0].text)
                current = [item["mail_id"] for item in data["items"]]
                assert len(current) == 10 and (ids is None or ids == current)
                ids = current
                times.append(elapsed)
    return dict(handshake_ms=round(handshake, 1), first_search_ms=round(times[0], 1),
                warm_p50_ms=round(statistics.median(times[1:]), 1),
                warm_p95_ms=round(sorted(times[1:])[28], 1), calls=len(times), ids=ids)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline-root", type=Path, required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    env = {k: v for k, v in os.environ.items() if not k.startswith("OUTLOOK_RAG_")}
    env["OUTLOOK_RAG_API_KEY"] = "synthetic-only"
    with tempfile.TemporaryDirectory() as folder:
        cfg = dict(model="fixture", dimensions=4, data_dir=str(Path(folder) / "index"),
                   embedding_url="http://127.0.0.1:1/embeddings", warmup_on_start=False)
        config = Path(folder) / "config.json"
        config.write_text(json.dumps(cfg), encoding="utf-8")
        from outlook_rag import app
        with patch.dict(os.environ, dict(env, OUTLOOK_RAG_CONFIG=str(config)), clear=True):
            cfg = app.config()
            sql, table = app.open_store(cfg)
            try:
                with patch.object(app, "embed", side_effect=lambda texts, *a, **kw: [[1., 0., 0., 0.] for _ in texts]):
                    for i in range(40):
                        mail = dict(id=hashlib.sha256(str(i).encode()).hexdigest(), entry_id=str(i), store_id="synthetic", folder="archive",
                                    subject=f"PCIe {i}", body=f"PCIe timing review reference {i}", sender="sender@example.invalid", recipients="team@example.invalid",
                                    received="2026-10-01T00:00:00+00:00", modified="1", conversation_id=str(i))
                        app.upsert_mail(mail, sql, table, cfg)
            finally:
                sql.close()
        before = asyncio.run(measure(args.baseline_root, config, "outlook_rag.app", env))
        after = asyncio.run(measure(root, config, "outlook_rag.entry", env))
        identical = before.pop("ids") == after.pop("ids")
        print(json.dumps(dict(synthetic_emails=40, queries="lexical; API-free", platform=sys.platform,
                              baseline=before, isolated=after, identical_top10=identical), ensure_ascii=False))


if __name__ == "__main__":
    main()
