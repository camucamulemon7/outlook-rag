"""Scope consent, read-only previews, bounded COM counts and qualified ETA."""
from contextlib import ExitStack
from datetime import datetime
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from mcp.server.fastmcp import FastMCP
from outlook_rag import features, jobs, server, sync_planning
from outlook_test_support import synthetic_com


class Planning(unittest.TestCase):
    def setUp(self):
        stack = ExitStack()
        self.addCleanup(stack.close)
        stack.enter_context(synthetic_com())
        self.folder = SimpleNamespace(GetTable=Mock(return_value=SimpleNamespace(GetRowCount=lambda: 123)),
                                      Items=SimpleNamespace(Count=100000, Restrict=Mock(return_value=SimpleNamespace(Count=7))))
        self.backend = SimpleNamespace(local_config=Mock(return_value=dict(model='fixture', data_dir='unused', folders='auto')),
            config=Mock(side_effect=AssertionError('Preview called API-resolving config')),
            resolve_folder=Mock(return_value=self.folder),
            discover_sources=Mock(return_value=(dict(stores=[dict(folders=[dict(path='inbox', item_count=100000)])], errors=[]), [('inbox', self.folder)])))
        self.dispatch = stack.enter_context(patch('win32com.client.Dispatch', return_value=SimpleNamespace(GetNamespace=lambda _: object())))
        stack.enter_context(patch.object(features, 'progress', return_value={'folders': []}))
        self.initialize = stack.enter_context(patch('pythoncom.CoInitialize'))
        self.uninitialize = stack.enter_context(patch('pythoncom.CoUninitialize'))

    def test_unspecified_scope_has_choices_and_never_resolves_api_or_starts(self):
        with patch.object(jobs, 'start') as start:
            tools = server.register_tools(FastMCP('fixture'), self.backend)
            result = tools['sync_index'](action='start')
        self.assertEqual(result['state'], 'confirmation_required')
        self.assertFalse(result['started'])
        self.assertEqual([p['since_days'] for p in result['previews']], [30, 90, 365, 0])
        self.assertEqual(result['previews'][-1]['candidate_items'], 100000)
        self.assertEqual(result['previews'][0]['candidate_items'], 123)
        self.assertEqual(result['previews'][-1]['time_estimate']['range_seconds'], [10000., 100000.])
        self.assertEqual(result['previews'][-1]['time_estimate']['basis'], 'assumed_1_to_10_items_per_second')
        start.assert_not_called()
        self.backend.config.assert_not_called()
        self.initialize.assert_called_once()
        self.uninitialize.assert_called_once()

    def test_recent_preview_counts_filtered_rows_without_reading_bodies(self):
        result = sync_planning.plan(self.backend, ['inbox'], 30)
        self.assertEqual(result['previews'][0]['candidate_items'], 123)
        expression = self.folder.GetTable.call_args.args[0]
        self.assertIn('[ReceivedTime] >=', expression)
        self.backend.resolve_folder.assert_called_once()
        self.assertFalse(result['embedding_api_called'])

    def test_provider_without_table_counts_uses_restricted_items(self):
        self.folder.GetTable.side_effect = AttributeError('unsupported')
        self.assertEqual(sync_planning.count_window(self.folder, 30, datetime.now().astimezone()), 7)
        self.folder.Items.Restrict.assert_called_once()

    def test_unknown_count_is_not_empty_or_zero_duration(self):
        self.backend.discover_sources.return_value[0]['stores'][0]['folders'][0]['item_count'] = None
        result = sync_planning.plan(self.backend, since_days=0)['previews'][0]
        self.assertIsNone(result['candidate_items'])
        self.assertIsNone(result['time_estimate']['range_seconds'])
        self.assertEqual(result['time_estimate']['unknown_folder_count'], 1)

    def test_failed_restricted_count_keeps_other_choices_and_releases_com(self):
        self.folder.GetTable.side_effect = OSError('secret provider failure')
        self.folder.Items.Restrict.side_effect = OSError('private provider details')
        result = sync_planning.plan(self.backend)
        self.assertIsNone(result['previews'][0]['candidate_items'])
        self.assertNotIn('private', str(result))
        self.assertEqual(result['previews'][-1]['candidate_items'], 100000)
        self.uninitialize.assert_called_once()

    def test_cooperative_preview_budget_skips_further_com_counts(self):
        with patch.object(sync_planning.time, 'perf_counter', side_effect=[0, 20, 20, 20]):
            result = sync_planning.plan(self.backend)
        self.folder.GetTable.assert_not_called()
        self.assertIsNone(result['previews'][0]['candidate_items'])
        self.assertEqual(result['previews'][-1]['candidate_items'], 100000)

    def test_invalid_period_fails_before_outlook(self):
        for value in [-1, 36501, True, 'all']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                sync_planning.plan(self.backend, since_days=value)
        self.dispatch.assert_not_called()


class ConsentAndEta(unittest.TestCase):
    def setUp(self):
        self.backend = SimpleNamespace(config=Mock(return_value={'data_dir': 'fixture'}))
        self.tools = server.register_tools(FastMCP('fixture'), self.backend)

    def test_confirmation_and_explicit_period_are_both_required(self):
        with (patch.object(sync_planning, 'plan', return_value={'state': 'confirmation_required', 'previews': [{}]}),
              patch.object(jobs, 'start') as start):
            for options in [{}, {'since_days': 0}, {'confirmed': True}]:
                self.assertEqual(self.tools['sync_index'](action='start', **options)['state'], 'confirmation_required')
        self.backend.config.assert_not_called()
        start.assert_not_called()

    def test_approved_all_or_recent_scope_reaches_worker_without_second_question(self):
        with (patch.object(sync_planning, 'plan', return_value={'selection': 'auto', 'previews': [{'candidate_items': 100}]}),
              patch.object(jobs, 'start', return_value={'state': 'starting'}) as start):
            for days in [0, 30]:
                self.assertEqual(self.tools['sync_index'](action='start', folders=['auto'], since_days=days, confirmed=True)['state'], 'starting')
                self.assertEqual(start.call_args.args[2], days)
                self.assertEqual(start.call_args.kwargs['preflight']['candidate_items'], 100)

    def test_eta_uses_measured_rate_and_qualifies_assumptions_unknown_and_zero(self):
        rows = [dict(processed_items=100, processing_seconds=20, remaining_items_estimate=1000)]
        value = sync_planning.estimate(rows)
        self.assertEqual(value['estimated_seconds'], 200)
        self.assertEqual(value['range_seconds'], [100., 400.])
        self.assertEqual(value['basis'], 'same_index_measured_scan_rate')
        self.assertIsNone(sync_planning.estimate([])['estimated_seconds'])
        self.assertIsNone(sync_planning.estimate([dict(remaining_items_estimate=None)])['range_seconds'])
        self.assertEqual(sync_planning.estimate([dict(remaining_items_estimate=0)])['estimated_seconds'], 0)
        self.assertIsNone(sync_planning.estimate([dict(processed_items=float('nan'), processing_seconds=5)], total=10)['items_per_second'])

    def test_reused_all_time_job_does_not_claim_selected_recent_scope(self):
        old = dict(state='running', reused_existing_job=True, options=dict(folders=['auto'], since_days=0, reconcile=False))
        with (patch.object(sync_planning, 'plan', return_value={'selection': 'auto', 'previews': [{}]}),
              patch.object(jobs, 'start', return_value=old)):
            result = self.tools['sync_index'](action='start', folders=['auto'], since_days=30, confirmed=True)
        self.assertFalse(result['scope_matches_request'])
        self.assertEqual(result['options']['since_days'], 0)
        self.assertIn('NOT applied', result['scope_warning'])

    def test_offline_job_eta_respects_period_and_preflight_counts(self):
        record = dict(state='running', options=dict(folders=['auto'], since_days=30), preflight=dict(folder_counts=[dict(folder='inbox', source_items_estimate=1000)]))
        rows = [dict(processed_items=100, processing_seconds=20, remaining_items_estimate=900)]
        with (patch.object(features, 'local_config', return_value={'data_dir': 'fixture', 'folders': 'auto'}),
              patch.object(jobs, 'status', return_value=record),
              patch.object(features, 'progress', return_value={'folders': rows}) as progress):
            result = self.tools['sync_index']()
        cfg = progress.call_args.args[0]
        self.assertEqual(cfg['_progress_since_days'], 30)
        self.assertEqual(cfg['_progress_totals'], {'inbox': 1000})
        self.assertEqual(result['time_estimate']['estimated_seconds'], 180)
        self.backend.config.assert_not_called()
