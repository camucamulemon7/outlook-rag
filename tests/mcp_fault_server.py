"""Real stdio test server with the production isolation boundary."""
import anyio
import sys
from pathlib import Path
from mcp.server.fastmcp import FastMCP
from outlook_rag import tool_runtime


def main():
    mcp = FastMCP("fault-test")
    runtime = tool_runtime.Pool([sys.executable, '-u', str(Path(__file__).with_name('mcp_fault_worker.py'))])
    timeout = float(sys.argv[1]) if len(sys.argv) > 1 else 2

    @mcp.tool()
    async def search_emails(query: str, marker: str | None = None) -> dict:
        return await runtime.call("search_emails", {"query": query, "marker": marker}, timeout)

    @mcp.tool()
    async def index_status() -> dict:
        return await runtime.call("index_status", {}, timeout)

    async def run():
        try:
            await mcp.run_stdio_async()
        finally:
            await runtime.aclose()
    anyio.run(run)


if __name__ == "__main__":
    main()
