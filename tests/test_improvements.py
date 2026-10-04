"""Regression tests for bounded sync, recovery, caches, and API dimensions."""
import hashlib
import json
import math
import tempfile
import time
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import httpx
from outlook_rag import app
from outlook_test_support import synthetic_com


class Improvements(unittest.TestCase):
    def setUp(self):
        com = synthetic_com()
        com.__enter__()
        self.addCleanup(com.__exit__, None, None, None)
        Path('work').mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir='work', ignore_cleanup_errors=True)
        self.cfg = dict(model='fixture', dimensions=4, embedding_url='http://fixture.invalid',
            data_dir=str(Path(self.temp.name)/'index'), folders=['inbox'], since_days=0,
            sync_max_seconds=30, sync_max_scanned=2000, sync_retry_limit=10, sync_batch_emails=2)
        self.sql, self.table = app.open_store(self.cfg)
        self.reads = []
        self.filters = []
        self.mails = {}
        self.folders = {}
        self.embed_calls = []
        def embed(texts, cfg, **kwargs):
            self.embed_calls.extend(texts)
            return [[1., 0., 0., 0.] for _ in texts]
        self.embed_patch = patch.object(app, 'embed', side_effect=embed)
        self.embed_patch.start()
        self.namespace = SimpleNamespace(GetItemFromID=lambda entry,store:self.mails[entry])
        self.dispatch = SimpleNamespace(GetNamespace=lambda _:self.namespace)

    def tearDown(self):
        self.embed_patch.stop()
        self.sql.close()
        self.temp.cleanup()
        # Pipeline/recovery fixtures also call setUp/tearDown directly.
        self.doCleanups()

    def folder(self, name, count, same_time=False, broken=None):
        metadata=[]
        owner=self
        now=datetime(2026,9,1,12,tzinfo=timezone.utc)
        for i in range(count):
            entry=name+str(i)
            received=now if same_time else now-timedelta(minutes=i)
            class Mail:
                Class=43
                Subject='subject'
                SenderEmailAddress='fixture@example.invalid'
                To='fixture@example.invalid'
                ConversationID='fixture'
                def __init__(self,entry,received):
                    self.EntryID=entry
                    self.ReceivedTime=received
                    self.LastModificationTime=now
                @property
                def Body(self):
                    owner.reads.append(self.EntryID)
                    if self.EntryID==broken:
                        raise ValueError('fixture unreadable mail')
                    return 'body '+self.EntryID
            self.mails[entry]=Mail(entry,received)
            metadata.append((entry,'IPM.Note',received,now))
        class Table:
            def __init__(self):
                self.position=0
                self.Columns=SimpleNamespace(RemoveAll=lambda:None,Add=lambda _:None)
            def Sort(self,*args): pass
            @property
            def EndOfTable(self): return self.position>=len(metadata)
            def GetArray(self,n):
                rows=metadata[self.position:self.position+n]
                self.position+=len(rows)
                return rows
        def get_table(expression):
            owner.filters.append(expression)
            return Table()
        folder=SimpleNamespace(EntryID=name,StoreID='store',GetTable=get_table)
        self.folders[name]=folder
        return metadata

    def sync(self, **kwargs):
        cfg=dict(self.cfg,folders=list(self.folders))
        with patch.object(app,'config',return_value=cfg),patch.object(app,'resolve_folder',side_effect=lambda ns,name:self.folders[name]),patch('win32com.client.Dispatch',return_value=self.dispatch):
            return app.sync(**kwargs)

    def test_metadata_table_and_tied_dates_resume_without_loss(self):
        self.folder('inbox',5,same_time=True)
        for _ in range(3):
            result=self.sync(max_total_emails=2)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],5)
        self.assertTrue(result['folders'][0]['window_complete'])
        self.assertEqual(len(self.reads),5)
        self.assertEqual(result['folders'][0]['metadata_source'],'GetTable')

    def test_scan_limit_covers_unchanged_metadata_without_body_reads(self):
        self.folder('inbox',5)
        self.sync()
        self.reads.clear()
        result=self.sync(max_scanned=2)
        self.assertEqual(result['scanned_total'],2)
        self.assertEqual(result['stop_reason'],'scan_limit')
        self.assertEqual(self.reads,[])
        self.sync(max_scanned=2)
        result=self.sync(max_scanned=2)
        self.assertTrue(result['folders'][0]['window_complete'])
        self.assertTrue(result['folders'][0]['incremental'])
        self.assertTrue(all('LastModificationTime' in f for f in self.filters[1:]))

    def test_folder_rotation_prevents_starvation(self):
        self.folder('first',5)
        self.folder('second',5)
        first=self.sync(max_total_emails=1)
        second=self.sync(max_total_emails=1)
        self.assertEqual(first['folders'][0]['folder'],'first')
        self.assertEqual(second['folders'][0]['folder'],'second')

    def test_short_term_table_ids_preserve_original_mail_ids(self):
        metadata=self.folder('inbox',3)
        for i, row in enumerate(metadata):
            short='short'+str(i)
            self.mails[short]=self.mails[row[0]]
            metadata[i]=(short,*row[1:])
        with patch.object(app,'outlook_session_key',return_value='fixture-session'):
            self.sync(max_total_emails=2)
            self.sync(max_total_emails=2)
            self.reads.clear()
            with patch.object(self.namespace,'GetItemFromID',side_effect=AssertionError('Aliases must avoid reopening unchanged mail')):
                self.sync()
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],3)
        self.assertEqual({row[0] for row in self.sql.execute('SELECT entry_id FROM emails')},{'inbox0','inbox1','inbox2'})
        self.assertEqual(self.reads,[])

    def test_unreadable_mail_is_retained_for_retry_and_other_mail_progresses(self):
        self.folder('inbox',3,broken='inbox1')
        result=self.sync()
        self.assertEqual(result['indexed_total'],2)
        self.assertEqual(result['failed_total'],1)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM sync_failures').fetchone()[0],1)
        self.folder('inbox',3)
        with self.sql:self.sql.execute('UPDATE sync_failures SET next_retry=0')
        result=self.sync()
        self.assertEqual(result['retried'],1)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM sync_failures').fetchone()[0],0)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],3)

    def test_reconcile_does_not_remove_mail_when_table_id_cannot_be_resolved(self):
        self.folder('inbox',3)
        self.sync()
        resolve=self.namespace.GetItemFromID
        def broken(entry,store):
            if entry=='inbox1':raise RuntimeError('fixture unresolved ID')
            return resolve(entry,store)
        with patch.object(self.namespace,'GetItemFromID',side_effect=broken):
            for _ in range(3):
                result=self.sync(reconcile=True,max_scanned=1)
                if result['folders'][0]['window_complete']:break
        self.assertTrue(result['folders'][0]['reconcile_skipped'])
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],3)

    def test_reconcile_remembers_seen_rows_across_bounded_calls(self):
        metadata=self.folder('inbox',5)
        self.sync()
        metadata.pop(2)
        for _ in range(3):
            result=self.sync(reconcile=True,max_scanned=2)
            if result['folders'][0]['window_complete']:
                break
            self.assertEqual(result['folders'][0]['removed_local'],0)
        self.assertTrue(result['folders'][0]['window_complete'])
        self.assertEqual(result['folders'][0]['removed_local'],1)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],4)

    def test_outage_does_not_advance_cursor_over_uncommitted_mail(self):
        self.folder('inbox',3)
        with patch.object(app,'embed',side_effect=httpx.ReadTimeout('fixture outage')):
            result=self.sync()
        self.assertEqual(result['stop_reason'],'embedding_unavailable')
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],0)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM sync_failures').fetchone()[0],0)
        self.sync()
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],3)

    def test_cooperative_time_limit_checkpoints_completed_batch(self):
        self.folder('inbox',5)
        original=app.embed.side_effect
        def slow(texts,cfg,**kwargs):
            time.sleep(1.05)
            return original(texts,cfg,**kwargs)
        with patch.object(app,'embed',side_effect=slow):result=self.sync(max_seconds=1)
        self.assertEqual(result['stop_reason'],'time_limit')
        self.assertEqual(result['indexed_total'],2)
        self.sync()
        self.assertEqual(self.sql.execute('SELECT count(*) FROM emails').fetchone()[0],5)

    def test_query_cache_reuses_embedding_but_not_search_results(self):
        self.folder('inbox',1)
        self.sync()
        with patch.object(app,'config',return_value=self.cfg):
            a=app.search('first query')
            calls=len(self.embed_calls)
            b=app.search('first query')
            self.assertFalse(a['query_cache_hit'])
            self.assertTrue(b['query_cache_hit'])
            self.assertEqual(calls,len(self.embed_calls))
            self.table.delete('id IS NOT NULL')
            with self.sql:
                self.sql.execute('DELETE FROM chunks')
                self.sql.execute('DELETE FROM lexical')
            self.assertEqual(app.search('first query')['count'],0)

    def test_cache_maintenance_keeps_referenced_embeddings(self):
        self.folder('inbox',1)
        self.sync()
        with self.sql:
            self.sql.execute('INSERT INTO embedding_cache VALUES (?,?,?)',('unused',b'1234','unused'))
        with patch.object(app,'config',return_value=self.cfg):result=app.maintain()
        self.assertEqual(result['removed_cached_embeddings'],1)
        self.assertEqual(self.sql.execute('SELECT count(*) FROM embedding_cache').fetchone()[0],1)


class ApiDimensions(unittest.TestCase):
    def setUp(self): app._API_DIMENSIONS.clear()
    def test_server_side_reduction_matches_stored_prefix(self):
        cfg=dict(model='fixture',dimensions=4,storage_dimensions=2,embedding_request_dimensions=2,embedding_url='http://fixture.invalid')
        requests=[]
        def respond(request):
            requests.append(json.loads(request.content))
            return httpx.Response(200,json={'data':[{'index':0,'embedding':[3.,4.]}]})
        with httpx.Client(transport=httpx.MockTransport(respond)) as client,patch.object(app,'api_key',return_value='fixture'):
            self.assertEqual(app.embed(['text'],cfg,client=client),[[.6,.8]])
        self.assertEqual(requests[0]['dimensions'],2)

    def test_rejected_dimensions_fall_back_once_and_remain_cached(self):
        cfg=dict(model='fixture',dimensions=4,storage_dimensions=2,embedding_request_dimensions=2,embedding_url='http://fixture.invalid')
        requests=[]
        def respond(request):
            payload=json.loads(request.content)
            requests.append(payload)
            if 'dimensions' in payload:return httpx.Response(400,json={'error':'unsupported dimensions'})
            return httpx.Response(200,json={'data':[{'index':0,'embedding':[3.,4.,8.,9.]}]})
        with httpx.Client(transport=httpx.MockTransport(respond)) as client,patch.object(app,'api_key',return_value='fixture'):
            self.assertEqual(app.embed(['text'],cfg,client=client),[[.6,.8]])
            app.embed(['other'],cfg,client=client)
        self.assertEqual(len(requests),3)
        self.assertNotIn('dimensions',requests[-1])

    def test_expired_sync_deadline_does_not_send_another_request(self):
        cfg=dict(model='fixture',dimensions=4,embedding_url='http://fixture.invalid',_embedding_deadline=time.perf_counter()-1)
        with httpx.Client(transport=httpx.MockTransport(lambda request:self.fail('Request after deadline'))) as client,patch.object(app,'api_key',return_value='fixture'):
            with self.assertRaises(httpx.ReadTimeout):app.embed(['text'],cfg,client=client)
