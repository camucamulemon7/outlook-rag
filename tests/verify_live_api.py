"""Synthetic mail integration against an already configured embedding endpoint.

Run with OUTLOOK_RAG_API_KEY in the process environment (never in arguments):
    python tests/verify_live_api.py --url URL --model MODEL
The local fault proxy simulates an outage without stopping the real provider.
No real Outlook/COM/WMI calls or account data are used.
"""
import argparse
import asyncio
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

import httpx
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from outlook_rag import app, jobs


class IntegrationFailure(Exception):
    def __init__(self, details):
        self.details = details


class FaultProxy:
    def __init__(self, upstream):
        self.available = True
        self.requests = 0
        self.upstream = upstream
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_args):
                pass

            def do_POST(self):
                owner.requests += 1
                body = self.rfile.read(int(self.headers['Content-Length']))
                if owner.available:
                    try:
                        response = httpx.post(owner.upstream, content=body, headers={
                            'Content-Type': 'application/json',
                            'Authorization': self.headers.get('Authorization', '')},
                            trust_env=False, timeout=120)
                        status, payload = response.status_code, response.content
                    except httpx.HTTPError:
                        status, payload = 502, b'{"error":"Synthetic proxy upstream unavailable"}'
                else:
                    status, payload = 503, b'{"error":"Synthetic test outage"}'
                self.send_response(status)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)

        self.server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.url = f'http://127.0.0.1:{self.server.server_port}/embeddings'

    def __enter__(self):
        self.thread.start()
        return self

    def __exit__(self, *_args):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(5)


async def offline_mcp(config_path, mail_id, job_id):
    from verify_mcp import payload, server_environment
    server = StdioServerParameters(command=sys.executable, args=[
        '-m', 'outlook_rag.app', '--config', str(config_path)], env=server_environment(config_path))
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            assert len(tools.tools) == 14
            status = payload(await session.call_tool('index_status', {}))
            assert status['emails'] == 3
            original = payload(await session.call_tool('get_indexed_mail', {'mail_id': mail_id}))
            assert original['source'] == 'local_index'
            for tool, arguments in (
                ('list_sync_failures', {}), ('get_sync_progress', {}),
                ('get_mail_thread', {'mail_id': mail_id}),
                ('sync_job_status', {'job_id': job_id}),
            ):
                payload(await session.call_tool(tool, arguments))
            cancelled = payload(await session.call_tool('cancel_sync_job', {'job_id': job_id}))
            assert cancelled['cancel_requested']
    return 'passed'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--url', required=True)
    parser.add_argument('--model', required=True)
    parser.add_argument('--uv', help='Also execute the existing API and MCP integration scripts')
    args = parser.parse_args()
    if not os.environ.get('OUTLOOK_RAG_API_KEY'):
        raise ValueError('Provide existing API credentials through OUTLOOK_RAG_API_KEY')
    report = dict(endpoint=args.url, model=args.model, data='synthetic_only',
                  platform=sys.platform, python='.'.join(map(str, sys.version_info[:3])),
                  real_outlook_com_wmi='not_tested', outage='local_fault_proxy_only')
    with tempfile.TemporaryDirectory(prefix='outlook-rag-live-api-') as temporary, FaultProxy(args.url) as proxy:
        config_path = Path(temporary)/'config.json'
        cfg = dict(model=args.model, embedding_url=proxy.url,
                   data_dir=str(Path(temporary)/'data'), warmup_on_start=False)
        cfg['dimensions'] = app.detect_dimensions(cfg)
        config_path.write_text(json.dumps(cfg), encoding='utf-8')
        # Caller settings must never redirect this synthetic check into an
        # existing mail index. Inherit only the requested credential value.
        test_environment = {key: value for key, value in os.environ.items()
                            if not key.startswith('OUTLOOK_RAG_')}
        test_environment.update(OUTLOOK_RAG_CONFIG=str(config_path),
                                OUTLOOK_RAG_API_KEY=os.environ['OUTLOOK_RAG_API_KEY'])
        with patch.dict(os.environ, test_environment, clear=True):
            cfg = app.config()
            config_path.write_text(json.dumps({key: value for key, value in cfg.items()
                                               if not key.startswith('_')}), encoding='utf-8')
            report.update(native_dimensions=cfg['dimensions'], stored_dimensions=app.vector_dimensions(cfg))
            sql, table = app.open_store(cfg)
            try:
                examples = [
                    ('製品出荷の予定変更', '部材の到着が遅れているため、出荷日を来月へ延期します。'),
                    ('社内懇親会のご案内', '金曜日に歓迎会を開催します。参加希望者は総務までご連絡ください。'),
                    ('請求書の送付', '今月分の請求書を添付しました。振込期限は月末です。'),
                ]
                mails = [dict(id=hashlib.sha256(f'synthetic:{i}'.encode()).hexdigest(),
                    entry_id=f'fixture-{i}', store_id='fixture-store', folder='test',
                    subject=subject, body=body, sender=f'sender{i}@example.invalid',
                    recipients='recipient@example.invalid', received=f'2026-09-{20+i}T12:00:00+09:00',
                    modified='1', conversation_id=f'fixture-{i}') for i, (subject, body) in enumerate(examples)]
                stats = app.upsert_mails(mails, sql, table, cfg)
                assert stats['embedded_chunks'] == 3
                ranks = []
                queries = ['納期が遅れる可能性がある案件', '歓迎会への参加方法', '請求書の支払い期限']
                for hybrid in (False, True):
                    for i, query in enumerate(queries):
                        result = app.search(query, hybrid=hybrid)
                        rank = next(n for n, item in enumerate(result['items'], 1) if item['mail_id'] == mails[i]['id'])
                        assert rank == 1
                        ranks.append(rank)
                assert app.search(queries[0], sender='sender1', since='2026-09-21', until='2026-09-21')['count'] == 1
                assert app.search(queries[0], folder='missing')['count'] == 0
                report.update(indexed_emails=3, semantic_and_hybrid_queries=6, top1_matches=sum(v == 1 for v in ranks))
                with patch.object(jobs, 'launch_worker'):
                    job_id = jobs.start(cfg)['job_id']
                proxy.available = False
                before = proxy.requests
                # Explicitly remove configured dimensions to exercise restart recovery.
                with patch.dict(os.environ, {'OUTLOOK_RAG_DIMENSIONS': 'auto'}):
                    report['offline_mcp_restart_and_controls'] = asyncio.run(offline_mcp(config_path, mails[0]['id'], job_id))
                    assert proxy.requests == before
                assert app.search(queries[0])['query_cache_hit']
                try:
                    app.search('A new uncached synthetic query during provider outage')
                except httpx.HTTPStatusError as exc:
                    assert exc.response.status_code == 503
                else:
                    raise AssertionError('Uncached query unexpectedly succeeded during outage')
                changed = dict(mails[0], modified='2', body='部材の手配が完了し、製品の出荷予定は11月2日に確定しました。')
                try:
                    app.upsert_mails([changed], sql, table, cfg)
                except httpx.HTTPStatusError as exc:
                    assert exc.response.status_code == 503
                else:
                    raise AssertionError('Mail update unexpectedly succeeded during outage')
                assert app.get_mail(mails[0]['id'])['body'] == mails[0]['body']
                assert sql.execute('SELECT count(*) FROM emails').fetchone()[0] == 3
                assert table.count_rows() == 3
                proxy.available = True
                app.upsert_mails([changed], sql, table, cfg)
                assert app.get_mail(changed['id'])['body'] == changed['body']
                recovered = app.search('確定した製品出荷日')
                assert recovered['items'][0]['mail_id'] == changed['id']
                report.update(outage_preserved_published_mail='passed', recovery_update_and_search='passed',
                              api_dimension_reduction=app.status()['api_dimension_reduction'])
            finally:
                sql.close()
            if args.uv:
                root = Path(__file__).resolve().parent
                for script in ('verify.py', 'verify_mcp.py'):
                    command = [sys.executable, str(root/script), str(config_path)]
                    if script == 'verify_mcp.py':
                        command.append(args.uv)
                    run = subprocess.run(command, capture_output=True, text=True, timeout=240)
                    if run.returncode:
                        # Do not print provider error bodies or process environment.
                        frames = re.findall(r'File "([^"]+)", line (\d+)', run.stderr)
                        details = dict(stage=script, return_code=run.returncode)
                        if frames:
                            details.update(child_failure_file=Path(frames[-1][0]).name,
                                           child_failure_line=int(frames[-1][1]))
                        raise IntegrationFailure(details)
                    report[script] = json.loads(run.stdout.splitlines()[-1])
    report['passed'] = True
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        last = traceback.extract_tb(exc.__traceback__)[-1]
        print(json.dumps({'passed': False, 'error_type': type(exc).__name__,
                          'failure_file': Path(last.filename).name, 'failure_line': last.lineno,
                          **(exc.details if isinstance(exc, IntegrationFailure) else {})}))
        raise SystemExit(1)
