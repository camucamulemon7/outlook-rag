"""Real stdio test server with the production isolation boundary."""
import anyio
import sys
from mcp.server.fastmcp import FastMCP
from outlook_rag import tool_runtime
from mcp_faults import execute


def main():
    mcp = FastMCP("fault-test")
    limiter = anyio.CapacityLimiter(4)
    timeout = float(sys.argv[1]) if len(sys.argv) > 1 else 2

    @mcp.tool()
    async def search_emails(query: str, marker: str | None = None) -> dict:
        return await tool_runtime.call("search_emails", {"query": query, "marker": marker}, timeout, limiter, execute)

    @mcp.tool()
    async def index_status() -> dict:
        return await tool_runtime.call("index_status", {}, timeout, limiter, execute)

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
