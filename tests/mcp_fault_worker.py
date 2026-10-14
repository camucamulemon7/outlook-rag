"""Worker entry with test faults; no fault hooks in the production protocol."""
from mcp_faults import execute
from outlook_rag.tool_worker import run

if __name__ == '__main__':
    run(execute)
