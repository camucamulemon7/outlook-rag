"""Stdlib watchdog outside the detached worker's COM/native interpreter."""
from __future__ import annotations

import json
import math
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
import uuid

if __package__:
    from . import diagnostics, process_guard
else:
    import diagnostics  # type: ignore[no-redef]
    import process_guard  # type: ignore[no-redef]


def supervise(config_path: str, job_id: str) -> int:
    snapshot = Path(config_path).resolve()
    cfg = json.loads(snapshot.read_text(encoding="utf-8"))
    raw_timeout = cfg.get("worker_timeout_seconds", max(60, cfg.get("sync_max_seconds", 30) + 30))
    timeout = float(raw_timeout)
    if isinstance(raw_timeout, bool) or not math.isfinite(timeout) or not 1 <= timeout <= 7200:
        raise ValueError("worker_timeout_seconds must be 1..7200")
    cancel_grace = min(10., timeout)
    record_path = snapshot.parent / (job_id + ".job.json")
    flag = snapshot.parent / (job_id + ".cancel")
    env = dict(os.environ)
    job = None
    process = None
    try:
        if os.name == "nt":
            name, job = process_guard.create_job()
            env[process_guard.JOB_ENV] = name
        process = subprocess.Popen([sys.executable, str(Path(__file__).with_name("worker.py")),
                                    "--execute", "--config", str(snapshot), "_job-worker", job_id],
                                   stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                   env=env, close_fds=True, start_new_session=os.name != "nt")
        heartbeat = time.monotonic()
        last_update = None
        cancelled_at = None
        while process.poll() is None:
            now = time.monotonic()
            try:
                record = json.loads(record_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                record = {}
            if (updated := record.get("updated_at")) != last_update:
                last_update, heartbeat = updated, now
            if flag.exists() and cancelled_at is None:
                cancelled_at = now
            cancel_expired = cancelled_at is not None and now - cancelled_at >= cancel_grace
            if cancel_expired or now - heartbeat >= timeout:
                # The idle deadline can precede the cancellation grace. A
                # pending user stop must still finish as cancelled, not failed.
                cancelled = cancelled_at is not None
                if job is not None:
                    job.Close()
                    job = None
                elif sys.platform != "win32":
                    os.killpg(process.pid, signal.SIGKILL)
                else:
                    process.kill()
                process.wait(timeout=5)
                record = json.loads(record_path.read_text(encoding="utf-8"))
                # Concurrent status can observe the released worker lock and
                # mark this job interrupted between termination and this write.
                if record.get("state") in {"starting", "running", "waiting", "interrupted"}:
                    reason = "cancel_timeout" if cancelled else "worker_timeout"
                    record.update(state="cancelled" if cancelled else "failed", reason=reason, forced_stop=True,
                                  updated_at=time.time(), finished_at=time.time())
                    if not cancelled:
                        record.update(error_type="TimeoutError", error_detail="Sync worker stopped responding; committed mail and checkpoints were retained. Start a new job to resume.")
                    temporary = record_path.with_name(record_path.name + "." + uuid.uuid4().hex + ".tmp")
                    try:
                        temporary.write_text(json.dumps(record, ensure_ascii=False), encoding="utf-8")
                        os.replace(temporary, record_path)
                    finally:
                        temporary.unlink(missing_ok=True)
                    diagnostics.append(cfg, job_id, "worker_watchdog_stopped", state=record["state"], reason=reason)
                return 0 if cancelled else 1
            time.sleep(.1)
        return int(process.returncode or 0)
    finally:
        if job is not None:
            job.Close()
        if process is not None:
            if process.poll() is None:
                if sys.platform == "win32":
                    process.kill()
                else:
                    os.killpg(process.pid, signal.SIGKILL)
            process.wait(timeout=5)
