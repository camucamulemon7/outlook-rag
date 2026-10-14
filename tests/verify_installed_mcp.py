"""Verify an installed wheel/uvx executable through actual MCP stdio.

Example: python tests/verify_installed_mcp.py -- C:/Python313/python.exe -m uv
tool run --offline --python 3.12 --from <wheel> outlook-rag
"""
import asyncio
from contextlib import closing
import json
import os
from pathlib import Path
import sqlite3
import sys
import tempfile

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def payload(result):
    assert not result.isError
    return result.structuredContent or json.loads(result.content[0].text)


async def verify(command):
    with tempfile.TemporaryDirectory(prefix='outlook-rag-installed-cold-') as temp:
        root = Path(temp)
        index = root / 'index'
        config = root / 'config.json'
        cfg = dict(model='installed-fixture', data_dir=str(index), key_file=str(root / 'missing-key'),
                   embedding_url='http://127.0.0.1:1/embeddings')
        config.write_text(json.dumps(cfg), encoding='utf-8')
        env = {k: v for k, v in os.environ.items() if not k.startswith(('OUTLOOK_RAG_', '_OUTLOOK_RAG_'))}
        params = StdioServerParameters(command=command[0], args=[*command[1:], '--config', str(config)], env=env)
        calls = 0
        with anyio.fail_after(60):
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    assert len((await session.list_tools()).tools) == 6
                    for mode in ('semantic', 'hybrid', 'lexical', 'identifier', 'auto', 'metadata', 'thread'):
                        result = await session.call_tool('search_emails', {'query': 'AB-1234', 'mode': mode})
                        assert payload(result)['state'] == 'not_initialized'
                        calls += 1
                    assert not index.exists()
                    index.mkdir()
                    with closing(sqlite3.connect(index / 'metadata.sqlite')):
                        pass
                    result = await session.call_tool('search_emails', {'query': 'test'})
                    assert payload(result)['state'] == 'initializing'
                    calls += 1
                    config.write_text('{broken', encoding='utf-8')
                    assert (await session.call_tool('index_status', {})).isError
                    calls += 1
                    config.write_text(json.dumps(cfg), encoding='utf-8')
                    assert not (await session.call_tool('index_status', {})).isError
                    calls += 1
                    assert (await session.call_tool('sync_index', {'action': 'start', 'since_days': 0, 'confirmed': True})).isError
                    calls += 1
                    await session.send_ping()
                    assert not (await session.call_tool('index_status', {})).isError
                    calls += 1
        return dict(installed_executable=True, cold_search_modes=7, tool_calls=calls, partial_sqlite=True,
                    same_session_config_repair=True, missing_credentials_job_error_without_disconnect=True)


if __name__ == '__main__':
    command = sys.argv[1:]
    if command and command[0] == '--':
        command.pop(0)
    if not command:
        raise SystemExit('Supply the installed MCP command after --')
    print(json.dumps(asyncio.run(verify(command))))
