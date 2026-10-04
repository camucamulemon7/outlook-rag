"""MCP stdio handshake and read-tool checks using the registered command."""
import asyncio
import json
import os
import sys
import time
from pathlib import Path
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def payload(result):
    if result.isError:
        raise AssertionError(str(result.content))
    if result.structuredContent:
        return result.structuredContent
    return json.loads(result.content[0].text)


def server_environment(config_path):
    key_name = json.loads(Path(config_path).read_text(encoding='utf-8-sig')).get('api_key_env', 'OUTLOOK_RAG_API_KEY')
    return {key: value for key, value in os.environ.items()
            if key.startswith(('OUTLOOK_RAG_', 'UV_')) or key == key_name}


async def main():
    root = Path(__file__).resolve().parents[1]
    config_path = Path(sys.argv[1]).resolve()
    uv = sys.argv[2]
    server = StdioServerParameters(command=uv, args=["run", "--frozen", "--project", str(root),
                                  "python", "-m", "outlook_rag.app", "--config", str(config_path)],
                                  env=server_environment(config_path))
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            names = [tool.name for tool in tools.tools]
            assert set(names) == {"search_emails", "index_status", "get_indexed_mail", "sync_emails", "optimize_index", "list_outlook_sources", "maintain_index", "start_sync_job", "sync_job_status", "cancel_sync_job", "list_sync_failures", "retry_failed_emails", "get_sync_progress", "get_mail_thread"}
            status = payload(await session.call_tool("index_status", {}))
            assert status["emails"] > 0
            timings = []
            for query in ("請求書と支払い期限について", "商品やサービスの購入に関するメール", "旅行の予約や移動について"):
                start = time.perf_counter()
                result = payload(await session.call_tool("search_emails", {"query": query, "limit": 3}))
                timings.append({"tool_ms": round((time.perf_counter()-start)*1000),
                                "search_ms": result["elapsed_ms"], "embedding_ms": result["embedding_ms"], "count": result["count"]})
                assert result["count"] > 0
                assert all("vector" not in item for item in result["items"])
            original = payload(await session.call_tool("get_indexed_mail", {"mail_id": result["items"][0]["mail_id"], "max_body_chars": 100}))
            assert original["source"] == "local_index" and len(original["body"]) <= 100
            report = {"mcp_handshake": "passed", "tools": names, "indexed_emails": status["emails"],
                              "indexed_chunks": status["chunks"], "queries": timings, "original_body": "passed"}
            print(json.dumps(report))
            if len(sys.argv) > 3:
                Path(sys.argv[3]).write_text(json.dumps(report, indent=2), encoding="utf-8")


if __name__ == "__main__":
    asyncio.run(main())
