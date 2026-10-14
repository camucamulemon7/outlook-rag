"""Actual detached bootstrap watchdog with an unresponsive native worker."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest

from filelock import FileLock
import psutil
from outlook_rag import worker, jobs as job_control


class WorkerWatchdog(unittest.TestCase):
    def exercise(self, cancel):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            package = root / 'outlook_rag'
            package.mkdir()
            (package / '__init__.py').write_text('', encoding='utf-8')
            (package / 'app.py').write_text(
                "import os,json\ndef config(): return json.load(open(os.environ['OUTLOOK_RAG_CONFIG']))\n",
                encoding='utf-8')
            (package / 'jobs.py').write_text('''
import json,os,time
from pathlib import Path
from filelock import FileLock
def run(job_id,cfg):
    root=Path(cfg['data_dir'])/'jobs'
    record=root/(job_id+'.job.json')
    (root/'worker.owner.json').write_text(json.dumps({'job_id':job_id}))
    with FileLock(root/'worker.lock'):
        with FileLock(root/'mutation.lock'):
            value=json.loads(record.read_text())
            value.update(state='running',pid=os.getpid(),updated_at=time.time())
            temporary=record.with_suffix('.tmp')
            temporary.write_text(json.dumps(value));temporary.replace(record)
            (root/'native.pid').write_text(str(os.getpid()))
            print('fixture-secret native output',flush=True)
            if os.name=='nt':
                import win32event
                event=win32event.CreateEvent(None,False,False,None)
                win32event.WaitForSingleObject(event,win32event.INFINITE)
            else:
                time.sleep(60)
''', encoding='utf-8')
            jobs = root / 'index/jobs'
            jobs.mkdir(parents=True)
            job_id = 'b' * 32
            cfg = dict(data_dir=str(jobs.parent), sync_max_seconds=1, worker_timeout_seconds=2)
            snapshot = jobs / (job_id + '.config.json')
            snapshot.write_text(json.dumps(cfg))
            record = jobs / (job_id + '.job.json')
            record.write_text(json.dumps(dict(job_id=job_id, state='starting', created_at=time.time(), updated_at=time.time())))
            env = {k: v for k, v in os.environ.items() if not k.startswith(('OUTLOOK_RAG_', '_OUTLOOK_RAG_'))}
            env.update(PYTHONPATH=str(root), OUTLOOK_RAG_API_KEY='fixture-secret', PYTHONUTF8='1')
            process = subprocess.Popen([sys.executable, str(Path(worker.__file__).resolve()), '--config',
                                        str(snapshot), '_job-worker', job_id], stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
            native = None
            try:
                marker = jobs / 'native.pid'
                deadline = time.monotonic() + 10
                while not marker.exists():
                    log = jobs / (job_id + '.log')
                    self.assertIsNone(process.poll(), 'Watchdog worker exited before native fault: ' + (log.read_text() if log.exists() else record.read_text()))
                    self.assertLess(time.monotonic(), deadline)
                    time.sleep(.02)
                native = psutil.Process(int(marker.read_text()))
                if cancel:
                    (jobs / (job_id + '.cancel')).touch()
                while process.poll() is None:
                    job_control.status(cfg, job_id)
                    self.assertLess(time.monotonic(), deadline)
                    time.sleep(.005)
                output, errors = process.communicate(timeout=10)
                self.assertEqual(process.returncode, 0 if cancel else 1)
                state = json.loads(record.read_text())
                self.assertEqual(state['state'], 'cancelled' if cancel else 'failed')
                self.assertEqual(state['reason'], 'cancel_timeout' if cancel else 'worker_timeout')
                self.assertTrue(state['forced_stop'])
                self.assertFalse(native.is_running())
                with FileLock(jobs / 'worker.lock', timeout=0), FileLock(jobs / 'mutation.lock', timeout=0):
                    pass
                self.assertNotIn('fixture-secret', json.dumps(state) + output.decode() + errors.decode())
            finally:
                if native is not None and native.is_running():
                    native.kill()
                    native.wait(timeout=3)
                if process.poll() is None:
                    process.kill()
                process.communicate(timeout=3)

    def test_stuck_native_sync_releases_locks_after_watchdog_timeout(self):
        self.exercise(False)

    def test_stuck_native_sync_cancellation_is_bounded(self):
        self.exercise(True)
