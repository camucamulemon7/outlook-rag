"""Behavior tests for targeted recovery, resumable progress and cached threads."""
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import httpx
from filelock import FileLock
from outlook_rag import app, features, jobs
import test_improvements as fixtures


class FunctionalFeatures(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.Improvements('test_metadata_table_and_tied_dates_resume_without_loss')
        self.fixture.setUp()
        self.cfg, self.sql, self.table = self.fixture.cfg, self.fixture.sql, self.fixture.table

    def tearDown(self):
        self.fixture.tearDown()

    def add(self, i, conversation='thread', store='store', body='cached body'):
        mail = dict(id=hashlib.sha256(f'{store}:{i}'.encode()).hexdigest(), entry_id=str(i),store_id=store,
                    folder='inbox',subject=f'mail {i}',body=body,sender='sender',recipients='recipient',
                    received=f'2026-09-{i+1:02d}T00:00:00+00:00',modified='1',conversation_id=conversation)
        app.upsert_mail(mail,self.sql,self.table,self.cfg)
        return mail['id']

    def record_failure(self, entry, folder='inbox', error=None):
        app.record_sync_failure(self.sql,dict(entry_id=entry,modified='1'),'store',folder,error or ValueError('private mail text'))
        return hashlib.sha256(f'store:{entry}'.encode()).hexdigest()

    def test_offline_diagnostics_resolve_explicit_index_without_api_probe(self):
        self.record_failure('a')
        env=dict(OUTLOOK_RAG_MODEL='fixture',OUTLOOK_RAG_DATA_DIR=str(Path(self.cfg['data_dir']).resolve()),LOCALAPPDATA=str(Path(self.fixture.temp.name).resolve()))
        with patch.dict(os.environ,env,clear=True),patch.object(app,'detect_dimensions',side_effect=AssertionError('API probe forbidden')):
            cfg=features.local_config()
            self.assertEqual(cfg['dimensions'],4)
            self.assertEqual(app.list_sync_failures()['total'],1)
            self.assertEqual(app.get_sync_progress()['folders'][0]['failed_emails'],1)
        with self.assertRaises(ValueError):features.failures(dict(self.cfg,model='different'))

    def test_cached_mail_status_and_job_controls_do_not_probe_provider(self):
        mid = self.add(0)
        with patch.object(jobs, 'launch_worker'):
            job = jobs.start(self.cfg)['job_id']
        env = dict(OUTLOOK_RAG_MODEL='fixture',
                   OUTLOOK_RAG_DATA_DIR=str(Path(self.cfg['data_dir']).resolve()),
                   LOCALAPPDATA=str(Path(self.fixture.temp.name).resolve()))
        with patch.dict(os.environ, env, clear=True), patch.object(
                app, 'detect_dimensions', side_effect=AssertionError('API probe forbidden')):
            for name, invoke in (
                ('cached mail', lambda: app.get_indexed_mail(mid)),
                ('index status', app.index_status),
                ('job status', lambda: app.sync_job_status(job)),
                ('job cancel', lambda: app.cancel_sync_job(job)),
            ):
                with self.subTest(tool=name):
                    result = invoke()
                    if name == 'cached mail':
                        self.assertEqual(result['body'], 'cached body')
                    elif name == 'index status':
                        self.assertEqual(result['emails'], 1)
                        self.assertEqual(result['dimensions'], 4)
                    else:
                        self.assertEqual(result['job_id'], job)
                        if name == 'job cancel':
                            self.assertTrue(result['cancel_requested'])

    def test_mcp_restart_uses_cached_dimensions_before_starting_transport(self):
        self.add(0)
        env = dict(OUTLOOK_RAG_MODEL='fixture', OUTLOOK_RAG_API_KEY='fixture-key',
                   OUTLOOK_RAG_DATA_DIR=str(Path(self.cfg['data_dir']).resolve()),
                   LOCALAPPDATA=str(Path(self.fixture.temp.name).resolve()))
        with patch.dict(os.environ, env, clear=True), patch.object(
                app, 'detect_dimensions', side_effect=AssertionError('API probe forbidden')), \
                patch.object(app.sys, 'argv', ['outlook-rag']), \
                patch.object(app.threading, 'Thread') as warmup, patch.object(app.mcp, 'run') as run:
            app.main()
        run.assert_called_once_with(transport='stdio')
        warmup.return_value.start.assert_called_once()

    def test_cli_status_reads_cached_index_when_provider_is_offline(self):
        self.add(0)
        env = {key: value for key, value in os.environ.items() if not key.startswith('OUTLOOK_RAG_')}
        env.update(OUTLOOK_RAG_MODEL='fixture', OUTLOOK_RAG_API_KEY='fixture-key',
                   OUTLOOK_RAG_DATA_DIR=str(Path(self.cfg['data_dir']).resolve()),
                   OUTLOOK_RAG_EMBEDDING_URL='http://127.0.0.1:1/embeddings')
        result = subprocess.run([sys.executable, '-m', 'outlook_rag.app', 'status'],
                                env=env, capture_output=True, text=True, timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['emails'], 1)

    def test_cached_mail_does_not_open_vector_store_and_preserves_identity_guard(self):
        mid = self.add(0)
        with patch.object(app, 'config', return_value=self.cfg), patch.object(
                app.lancedb, 'connect', side_effect=AssertionError('Vector store not required')):
            result = app.get_indexed_mail(mid, 3)
            self.assertEqual(result['body'], 'cac')
            self.assertTrue(result['body_truncated'])
        with patch.object(app, 'config', return_value=dict(self.cfg, model='different')):
            with self.assertRaisesRegex(ValueError, 'settings do not match'):
                app.get_indexed_mail(mid)

    def test_failure_list_paging_and_safe_causes(self):
        request=httpx.Request('POST','http://fixture.invalid',headers={'Authorization':'Bearer secret'})
        response=httpx.Response(401,json={'error':'private body and secret'},request=request)
        first=self.record_failure('a',error=httpx.HTTPStatusError('private body and secret',request=request,response=response))
        self.record_failure('b','other')
        result=features.failures(self.cfg,1,0,'inbox')
        self.assertEqual(result['total'],1)
        self.assertEqual(result['items'][0]['mail_id'],first)
        self.assertFalse(result['items'][0]['retry_due'])
        self.assertIn('401',result['items'][0]['error_detail'])
        self.assertNotIn('secret',json.dumps(result))
        self.assertNotIn('private body',json.dumps(result))
        self.assertEqual(features.failures(self.cfg,1)['next_offset'],1)

    def test_targeted_retry_bypasses_backoff_and_leaves_others(self):
        self.fixture.folder('inbox',2)
        ids=[self.record_failure('inbox'+str(i)) for i in range(2)]
        with patch('win32com.client.Dispatch',return_value=self.fixture.dispatch):
            result=features.retry(self.cfg,[ids[0],ids[0]])
        self.assertEqual(result['recovered'],1)
        self.assertEqual(self.fixture.reads,['inbox0'])
        self.assertEqual(self.sql.execute('SELECT mail_id FROM sync_failures').fetchone()[0],ids[1])
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],1)

    def test_retry_provider_outage_retains_failure_and_other_pending(self):
        self.fixture.folder('inbox',2)
        ids=[self.record_failure('inbox'+str(i)) for i in range(2)]
        with patch('win32com.client.Dispatch',return_value=self.fixture.dispatch), patch.object(app,'embed',side_effect=httpx.ConnectError('unavailable')):
            result=features.retry(self.cfg,ids)
        self.assertEqual(result['state'],'incomplete')
        self.assertEqual(result['pending_mail_ids'],[ids[1]])
        self.assertEqual(self.sql.execute('SELECT count(*) FROM sync_failures').fetchone()[0],2)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],0)

    def test_outlook_read_failure_does_not_prevent_next_selected_retry(self):
        self.fixture.folder('inbox',2)
        ids=[self.record_failure('inbox'+str(i)) for i in range(2)]
        original=app.read_outlook_mail
        def read(ns,metadata,store,folder):
            if metadata['entry_id']=='inbox0':raise PermissionError('private description')
            return original(ns,metadata,store,folder)
        with patch('win32com.client.Dispatch',return_value=self.fixture.dispatch),patch.object(app,'read_outlook_mail',side_effect=read):
            result=features.retry(self.cfg,ids)
        self.assertEqual(result['failed'],1)
        self.assertEqual(result['recovered'],1)
        self.assertEqual(result['pending_mail_ids'],[])
        self.assertEqual(self.sql.execute('SELECT mail_id FROM sync_failures').fetchone()[0],ids[0])

    def test_retry_busy_does_not_change_backoff(self):
        mid=self.record_failure('a')
        before=tuple(self.sql.execute('SELECT * FROM sync_failures').fetchone())
        with FileLock(Path(self.cfg['data_dir'])/'mutation.lock'):
            result=features.retry(self.cfg,[mid])
        self.assertEqual(result['state'],'busy')
        self.assertEqual(tuple(self.sql.execute('SELECT * FROM sync_failures').fetchone()),before)

    def test_progress_continues_tied_date_window_without_double_count(self):
        self.fixture.folder('inbox',5,same_time=True)
        self.fixture.folders['inbox'].Items=SimpleNamespace(Count=5)
        self.fixture.sync(max_total_emails=2)
        value=features.progress(self.cfg)['folders'][0]
        self.assertEqual(value['processed_items'],2)
        self.assertEqual(value['remaining_items_estimate'],3)
        self.assertIsNotNone(value['eta_seconds_estimate'])
        self.fixture.sync(max_total_emails=2)
        self.fixture.sync(max_total_emails=2)
        value=features.progress(self.cfg)['folders'][0]
        self.assertEqual(value['processed_items'],5)
        self.assertEqual(value['eta_seconds_estimate'],0)
        self.assertEqual(value['indexed_emails'],5)
        # Incremental scope has no known total; do not reuse full-folder counts.
        self.fixture.sync()
        value=features.progress(self.cfg)['folders'][0]
        self.assertIsNone(value['source_items_estimate'])

    def test_unpublished_batch_does_not_advance_progress(self):
        self.fixture.folder('inbox',3)
        self.fixture.folders['inbox'].Items=SimpleNamespace(Count=3)
        with patch.object(app,'embed',side_effect=httpx.ConnectError('down')):
            self.fixture.sync()
        value=features.progress(self.cfg)['folders'][0]
        self.assertEqual(value['processed_items'],0)
        self.assertIsNone(value['eta_seconds_estimate'])
        self.fixture.sync()
        self.assertEqual(features.progress(self.cfg)['folders'][0]['processed_items'],3)

    def test_partial_publication_counts_only_durable_rows_after_outage(self):
        self.fixture.folder('inbox',5)
        self.fixture.folders['inbox'].Items=SimpleNamespace(Count=5)
        with patch.object(app,'embed',side_effect=[[[1.,0.,0.,0.]]*2,httpx.ConnectError('down')]):
            self.fixture.sync()
        value=features.progress(self.cfg)['folders'][0]
        self.assertEqual(value['processed_items'],2)
        self.assertEqual(value['remaining_items_estimate'],3)
        self.fixture.sync()
        self.assertEqual(features.progress(self.cfg)['folders'][0]['processed_items'],5)

    def test_old_checkpoint_does_not_invent_remaining_count(self):
        value=dict(folder='inbox',complete=False,scan_mode='backfill',cursor_received='2026-09-01',last_run='2026-09-01')
        with self.sql:self.sql.execute('INSERT OR REPLACE INTO state VALUES (?,?)',('sync:old',json.dumps(value)))
        result=features.folder_progress(self.sql,self.cfg)[0]
        self.assertIsNone(result['processed_items'])
        self.assertIsNone(result['remaining_items_estimate'])
        self.assertIsNone(result['eta_seconds_estimate'])

    def test_thread_context_is_ordered_bounded_and_store_scoped(self):
        ids=[self.add(i,body='body '+str(i)) for i in range(5)]
        self.add(2,store='other')
        result=features.thread(self.cfg,ids[2],before=1,after=1,max_body_chars=3)
        self.assertEqual([v['mail_id'] for v in result['items']],ids[1:4])
        self.assertEqual(result['thread_mail_count'],5)
        self.assertTrue(result['has_more_previous'])
        self.assertTrue(result['has_more_next'])
        self.assertTrue(all(v['body_truncated'] for v in result['items']))
        boundary=features.thread(self.cfg,ids[0],before=0,after=0)
        self.assertFalse(boundary['has_more_previous'])
        self.assertTrue(boundary['has_more_next'])

    def test_missing_conversation_is_not_joined_or_collapsed(self):
        first=self.add(0,conversation='')
        self.add(1,conversation='')
        self.assertEqual(features.thread(self.cfg,first)['count'],1)
        with patch.object(app,'config',return_value=self.cfg):
            self.assertEqual(app.search('cached',group_by_thread=True)['count'],2)

    def test_grouped_search_retains_best_match_and_distinct_store(self):
        self.add(0);self.add(1);self.add(2,store='other')
        with patch.object(app,'config',return_value=self.cfg):
            raw=app.search('cached',group_by_thread=False)
            grouped=app.search('cached',group_by_thread=True)
        self.assertEqual(raw['count'],3)
        self.assertEqual(grouped['count'],2)
        self.assertEqual(sorted(v['thread_mail_count'] for v in grouped['items']),[1,2])
        self.assertEqual(grouped['items'][0]['mail_id'],raw['items'][0]['mail_id'])

    def test_thread_body_budget_preserves_focus(self):
        ids=[self.add(i,body='x'*10000) for i in range(10)]
        result=features.thread(self.cfg,ids[9],before=9,max_body_chars=10000)
        self.assertLessEqual(sum(len(v['body']) for v in result['items']),60000)
        self.assertEqual(len(next(v['body'] for v in result['items'] if v['position']=='focus')),10000)

    def test_old_worker_can_write_original_failure_schema_and_stale_causes_are_hidden(self):
        mid=self.record_failure('a')
        row=list(self.sql.execute('SELECT * FROM sync_failures').fetchone())
        self.assertEqual(len(row),9)
        row[5]+=1;row[8]+=1
        with self.sql:self.sql.execute('INSERT OR REPLACE INTO sync_failures VALUES (?,?,?,?,?,?,?,?,?)',row)
        self.assertIn('No detailed cause',features.failures(self.cfg)['items'][0]['error_detail'])
        with self.sql:self.sql.execute('DELETE FROM sync_failures WHERE mail_id=?',(mid,))
        self.assertEqual(self.sql.execute('SELECT count(*) FROM sync_failure_details').fetchone()[0],0)

    def test_legacy_failure_schema_migrates_without_losing_rows(self):
        mid=self.record_failure('a')
        with self.sql:
            self.sql.execute('ALTER TABLE sync_failures RENAME TO old_failures')
            self.sql.execute('CREATE TABLE sync_failures AS SELECT mail_id,entry_id,store_id,folder,modified,attempts,error_type,next_retry,last_failed FROM old_failures')
            self.sql.execute('DROP TABLE old_failures')
        sql,_=app.open_store(self.cfg)
        try:self.assertEqual(sql.execute('SELECT mail_id FROM sync_failures').fetchone()[0],mid)
        finally:sql.close()


class ReadablePaths(unittest.TestCase):
    def test_job_snapshot_recovery_preserves_settings_and_ambiguity_guards(self):
        with tempfile.TemporaryDirectory() as root:
            env = dict(LOCALAPPDATA=str(Path(root)/'local'), OUTLOOK_RAG_MODEL='snapshot-model')
            with patch.dict(os.environ, env, clear=True), patch.object(Path, 'home', return_value=Path(root)/'home'):
                configs = []
                for dimension in ('16', '32'):
                    with patch.dict(os.environ, {'OUTLOOK_RAG_DIMENSIONS':dimension}):
                        configs.append(app.config())
                        with patch.object(jobs, 'launch_worker'):
                            jobs.start(configs[-1])
                with patch.object(app, 'detect_dimensions', side_effect=AssertionError('API probe forbidden')):
                    with self.assertRaisesRegex(ValueError, 'Multiple cached indexes'):
                        features.local_config()
                    with patch.dict(os.environ, {'OUTLOOK_RAG_DATA_DIR':configs[0]['data_dir']}):
                        self.assertEqual(features.local_config()['dimensions'], 16)
                    for changed in ({'OUTLOOK_RAG_MODEL':'different'}, {'OUTLOOK_RAG_CHUNK_SIZE':'800'}):
                        with patch.dict(os.environ, changed):
                            self.assertFalse(features.local_config().get('dimensions'))

    def test_offline_job_controls_resolve_snapshot_before_metadata_is_initialized(self):
        with tempfile.TemporaryDirectory() as root:
            home = Path(root)/'profile'
            env = dict(LOCALAPPDATA=str(Path(root)/'local'), OUTLOOK_RAG_MODEL='starting-model',
                       OUTLOOK_RAG_DIMENSIONS='16')
            with patch.dict(os.environ, env, clear=True), patch.object(Path, 'home', return_value=home):
                cfg = app.config()
                with patch.object(jobs, 'launch_worker'):
                    job = jobs.start(cfg)['job_id']
                self.assertFalse((Path(cfg['data_dir'])/'metadata.sqlite').exists())
                with patch.dict(os.environ, {'OUTLOOK_RAG_DIMENSIONS':'auto'}), patch.object(
                        app, 'detect_dimensions', side_effect=AssertionError('API probe forbidden')):
                    self.assertEqual(app.sync_job_status(job)['state'], 'starting')
                    self.assertTrue(app.cancel_sync_job(job)['cancel_requested'])
                    self.assertEqual(features.local_config()['data_dir'], cfg['data_dir'])

    def test_auto_index_diagnostics_find_cached_dimensions_without_api(self):
        with tempfile.TemporaryDirectory() as root:
            home=Path(root)/'profile';local=Path(root)/'local'
            env=dict(LOCALAPPDATA=str(local),TEMP=root,TMP=root,OUTLOOK_RAG_MODEL='cached-model',OUTLOOK_RAG_DIMENSIONS='16')
            with patch.dict(os.environ,env,clear=True),patch.object(Path,'home',return_value=home):
                cfg=app.config();sql,_=app.open_store(cfg);sql.close()
                with patch.dict(os.environ,{'OUTLOOK_RAG_DIMENSIONS':'auto'}),patch.object(app,'detect_dimensions',side_effect=AssertionError('No probe')):
                    resolved=features.local_config()
                    self.assertEqual(resolved['data_dir'],cfg['data_dir'])
                    self.assertEqual(resolved['dimensions'],16)
                    self.assertEqual(app.list_sync_failures()['total'],0)
                    self.assertEqual(app.get_sync_progress()['folders'],[])
                with patch.dict(os.environ,{'OUTLOOK_RAG_MODEL':'no-existing-index','OUTLOOK_RAG_DIMENSIONS':'auto'}),patch.object(app,'detect_dimensions',side_effect=AssertionError('No probe')):
                    self.assertEqual(app.list_sync_failures()['items'],[])
                    self.assertEqual(app.get_sync_progress()['folders'],[])
                    with self.assertRaisesRegex(ValueError, 'Index not initialized'):
                        app.index_status()
                    with self.assertRaisesRegex(ValueError, 'Indexed mail not found'):
                        app.get_indexed_mail('0'*64)

    def test_model_name_dimensions_and_settings_isolation_with_legacy_lookup(self):
        with tempfile.TemporaryDirectory() as root:
            home=Path(root)/'profile';local=Path(root)/'local'
            env=dict(LOCALAPPDATA=str(local),OUTLOOK_RAG_MODEL='org/model',OUTLOOK_RAG_DIMENSIONS='16')
            with patch.dict(os.environ,env,clear=True),patch.object(Path,'home',return_value=home):
                cfg=app.config();preferred=Path(cfg['data_dir'])
                self.assertTrue(preferred.name.startswith('org_model-16d-'))
                with patch.dict(os.environ,{'OUTLOOK_RAG_CHUNK_SIZE':'800'}):
                    self.assertNotEqual(app.config()['data_dir'],str(preferred))
                digest=hashlib.sha256(app.index_identity(cfg).encode()).hexdigest()[:12]
                old=preferred.parent/digest;old.mkdir(parents=True);(old/'metadata.sqlite').touch()
                self.assertEqual(Path(app.config()['data_dir']),old)
                preferred.mkdir();(preferred/'metadata.sqlite').touch()
                self.assertEqual(Path(app.config()['data_dir']),preferred)
                with patch.dict(os.environ,{'OUTLOOK_RAG_MODEL':'CON'}):
                    self.assertTrue(Path(app.config()['data_dir']).name.startswith('model_CON-'))
                with patch.dict(os.environ,{'OUTLOOK_RAG_MODEL':'org/x?<>|:*'}):
                    self.assertFalse(any(c in Path(app.config()['data_dir']).name for c in '<>:"/\\|?*'))

if __name__=='__main__':unittest.main()
