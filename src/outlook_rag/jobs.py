"""Detached, checkpointed bulk sync jobs using the existing sync implementation."""
from __future__ import annotations

import json
import os
import re
import sqlite3
import subprocess
import sys
import time
import uuid
from contextlib import closing
from pathlib import Path

from filelock import FileLock, Timeout

ACTIVE = {"starting", "running", "waiting"}


def root(cfg):
    path = Path(cfg["data_dir"]) / "jobs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def paths(cfg, job_id):
    if not isinstance(job_id, str) or not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise ValueError("Invalid job_id")
    folder = root(cfg)
    return folder / (job_id + ".job.json"), folder / (job_id + ".cancel")


def write(path, value):
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temporary.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def load(cfg, job_id=None):
    if job_id is None:
        files = list(root(cfg).glob("*.job.json"))
        if not files:
            return None, None
        # Worker updates must not change which job an omitted ID selects.
        records = [(p, json.loads(p.read_text(encoding="utf-8"))) for p in files]
        return max(records, key=lambda pair: (pair[1]["created_at"], pair[0].name))
    else:
        path, _ = paths(cfg, job_id)
    if not path.is_file():
        raise ValueError("Sync job not found")
    return path, json.loads(path.read_text(encoding="utf-8"))


def worker_running(cfg):
    try:
        with FileLock(root(cfg) / "worker.lock", timeout=0):
            return False
    except Timeout:
        return True


def coverage(cfg):
    path = (Path(cfg["data_dir"]) / "metadata.sqlite").resolve()
    counts = dict(emails=0, chunks=0, cached_embeddings=0, failed_emails=0)
    if path.is_file():
        with closing(sqlite3.connect(path.as_uri() + "?mode=ro", uri=True, timeout=5)) as sql:
            for name, table in (("emails", "emails"), ("chunks", "chunks"),
                                ("cached_embeddings", "embedding_cache"), ("failed_emails", "sync_failures")):
                if sql.execute("SELECT 1 FROM sqlite_master WHERE name=?", (table,)).fetchone():
                    counts[name] = sql.execute("SELECT count(*) FROM " + table).fetchone()[0]
    return counts


def launch_worker(command, *, env):
    options = dict(stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   env=env, close_fds=True)
    if os.name == "nt":
        # WMI-created processes do not inherit the MCP client's Windows Job
        # Object. Transfer credentials in memory, never in command lines/files.
        import pythoncom
        import win32com.client
        from multiprocessing.connection import Client

        pipe = r"\\.\pipe\outlook-rag-" + uuid.uuid4().hex
        key = os.urandom(32)
        bootstrap = (
            "import os,sys,runpy,threading; from multiprocessing.connection import Listener; "
            "timer=threading.Timer(15,lambda:os._exit(1)); timer.daemon=True; timer.start(); "
            "listener=Listener(sys.argv[1],family='AF_PIPE',authkey=bytes.fromhex(sys.argv[2])); "
            "connection=listener.accept(); "
            "assert connection.poll(15), 'Worker environment transfer timed out'; "
            "os.environ.clear(); os.environ.update(connection.recv()); "
            "connection.close(); listener.close(); timer.cancel(); "
            "sys.argv=sys.argv[4:]; "
            "sys.stdin=open(os.devnull); sys.stdout=sys.stderr=open(os.devnull,'w'); "
            "runpy.run_module('outlook_rag.app',run_name='__main__')"
        )
        pythoncom.CoInitialize()
        try:
            wmi = win32com.client.GetObject(r"winmgmts:\\.\root\cimv2")
            startup = wmi.Get("Win32_ProcessStartup").SpawnInstance_()
            startup.ShowWindow = 0
            process = wmi.Get("Win32_Process")
            arguments = process.Methods_("Create").InParameters.SpawnInstance_()
            arguments.CommandLine = subprocess.list2cmdline([command[0], "-c", bootstrap, pipe, key.hex(), *command[1:]])
            arguments.ProcessStartupInformation = startup
            result = process.ExecMethod_("Create", arguments)
            if result.ReturnValue:
                raise OSError("Could not start background worker (WMI code %s)" % result.ReturnValue)
            deadline = time.monotonic() + 10
            while True:
                try:
                    connection = Client(pipe, family="AF_PIPE", authkey=key)
                    break
                except (FileNotFoundError, OSError):
                    if time.monotonic() >= deadline:
                        raise TimeoutError("Background worker did not start")
                    time.sleep(0.05)
            with connection:
                connection.send(env)
        finally:
            # Release COM objects before uninitializing this thread's apartment.
            result = arguments = process = startup = wmi = None
            pythoncom.CoUninitialize()
    else:
        subprocess.Popen(command, **options, start_new_session=True)


def status(cfg, job_id=None):
    path, record = load(cfg, job_id)
    if record is None:
        return {"state": "not_started"}
    # A released OS lock distinguishes a crashed/exited worker from a slow COM call.
    running = worker_running(cfg)
    owner_path = root(cfg) / "worker.owner.json"
    owner = json.loads(owner_path.read_text(encoding="utf-8")) if running and owner_path.is_file() else {}
    owns_worker = running and owner.get("job_id") == record["job_id"]
    # The worker can publish its terminal state between the first read and the
    # lock inspection. Do not classify that completed job using a stale record.
    _, record = load(cfg, record["job_id"])
    if record["state"] in ACTIVE and not owns_worker:
        if record["state"] != "starting" or time.time() - record["updated_at"] > 60:
            record = dict(record, state="interrupted", reason="worker_exited")
            # Status is read-only; a later start can resume the index checkpoints.
    _, cancellation = paths(cfg, record["job_id"])
    return {key: value for key, value in record.items() if key not in ("known", "complete")} | {
        "elapsed_ms": round(1000 * ((record.get("finished_at") or time.time()) - record["created_at"])),
        "cancel_requested": cancellation.exists(), "coverage": coverage(cfg)}


def start(cfg, folders=None, since_days=None, reconcile=False):
    if folders is not None and (not isinstance(folders, list) or not all(isinstance(f, str) for f in folders)):
        raise ValueError("folders must be a list of paths")
    if since_days is not None and (type(since_days) is not int or not 0 <= since_days <= 36500):
        raise ValueError("since_days must be 0..36500")
    with FileLock(root(cfg) / "start.lock", timeout=1):
        _, old = load(cfg)
        if worker_running(cfg) or (old and old["state"] == "starting" and time.time() - old["updated_at"] < 60):
            return status(cfg) | {"reused_existing_job": True}
        job_id = uuid.uuid4().hex
        path, _ = paths(cfg, job_id)
        snapshot_path = root(cfg) / (job_id + ".config.json")
        snapshot = {key: value for key, value in cfg.items()
                    if not key.startswith("_") and ("api_key" not in key.lower() or key == "api_key_env")}
        write(snapshot_path, snapshot)
        now = time.time()
        record = dict(job_id=job_id, state="starting", created_at=now, updated_at=now,
                      options=dict(folders=folders, since_days=since_days, reconcile=reconcile),
                      cycles=0, indexed_total=0, scanned_total=0, retried_total=0,
                      known=[], complete=[], known_folders=0, completed_folders=0)
        write(path, record)
        key_name = cfg.get("api_key_env", "OUTLOOK_RAG_API_KEY")
        env = {key: value for key, value in os.environ.items()
               if not key.startswith("OUTLOOK_RAG_") or key == key_name}
        env["OUTLOOK_RAG_CONFIG"] = str(snapshot_path.resolve())
        try:
            launch_worker([sys.executable, "-m", "outlook_rag.app", "--config", str(snapshot_path.resolve()),
                           "_job-worker", job_id], env=env)
        except Exception as exc:
            record.update(state="failed", error_type=type(exc).__name__, finished_at=time.time())
            write(path, record)
            raise
        return status(cfg, job_id) | {"reused_existing_job": False}


def cancel(cfg, job_id=None):
    _, record = load(cfg, job_id)
    if record is None:
        return {"state": "not_started"}
    if record["state"] in ACTIVE:
        _, flag = paths(cfg, record["job_id"])
        flag.touch(exist_ok=True)
    return status(cfg, record["job_id"])


def scoped_failures(cfg, known):
    path = (Path(cfg["data_dir"]) / "metadata.sqlite").resolve()
    if not path.is_file() or not known:
        return 0
    with closing(sqlite3.connect(path.as_uri() + "?mode=ro", uri=True, timeout=5)) as sql:
        return sql.execute("SELECT count(*) FROM sync_failures WHERE folder IN (" +
                           ",".join("?" for _ in known) + ")", sorted(known)).fetchone()[0]


def run(job_id, cfg):
    from . import app
    path, flag = paths(cfg, job_id)
    _, record = load(cfg, job_id)
    claimed_worker = False
    try:
        with FileLock(root(cfg) / "worker.lock", timeout=0):
            claimed_worker = True
            write(root(cfg) / "worker.owner.json", {"job_id": job_id})
            app.api_key(cfg)
            record.update(state="running", pid=os.getpid(), updated_at=time.time())
            write(path, record)
            known, complete = set(record["known"]), set(record["complete"])
            stalls = outages = 0
            while True:
                if flag.exists():
                    record.update(state="cancelled")
                    break
                before = coverage(cfg)
                try:
                    result = app.sync(**record["options"], _cancel_check=flag.exists)
                except Timeout:
                    record.update(state="waiting", reason="index_busy", updated_at=time.time())
                    write(path, record)
                    # Another bounded/manual sync owns the database; retry without spinning.
                    for _ in range(10):
                        if flag.exists():
                            break
                        time.sleep(0.1)
                    continue
                record.update(state="running", updated_at=time.time(), cycles=record["cycles"] + 1)
                record.pop("reason", None)
                for field, source in (("indexed_total", "indexed_total"), ("scanned_total", "scanned_total"), ("retried_total", "retried")):
                    record[field] += result.get(source, 0)
                for folder in result["folders"]:
                    known.add(folder["folder"])
                    if folder["window_complete"]:
                        complete.add(folder["folder"])
                    else:
                        complete.discard(folder["folder"])
                known.update(result.get("pending_folders", []))
                record.update(known=sorted(known), complete=sorted(complete), known_folders=len(known),
                              completed_folders=len(complete), last_cycle_ms=result.get("elapsed_ms"),
                              stop_reason=result.get("stop_reason"),
                              discovery_errors=len(result.get("discovery_errors", [])),
                              folder_errors=len(result.get("folder_errors", [])))
                failed = scoped_failures(cfg, known)
                record["failed_emails"] = failed
                if flag.exists() or result.get("stop_reason") == "cancelled":
                    record["state"] = "cancelled"
                elif known <= complete:
                    record["state"] = "completed_with_errors" if failed or record["discovery_errors"] or record["folder_errors"] else "completed"
                elif result.get("discovery_errors") or result.get("folder_errors"):
                    record.update(state="completed_with_errors", reason="folder_scan_error")
                else:
                    after = coverage(cfg)
                    progressed = result.get("scanned_total", 0) > 0 or after != before
                    stalls = 0 if progressed else stalls + 1
                    outages = outages + 1 if result.get("stop_reason") == "embedding_unavailable" else 0
                    if outages >= 3 or stalls >= 3:
                        record.update(state="blocked", reason="embedding_unavailable" if outages >= 3 else "no_progress")
                write(path, record)
                if record["state"] not in ACTIVE:
                    break
                if result.get("stop_reason") == "embedding_unavailable":
                    for _ in range(10 * outages):
                        if flag.exists():
                            break
                        time.sleep(0.1)
            record.update(updated_at=time.time(), finished_at=time.time())
            write(path, record)
    except Exception as exc:
        if isinstance(exc, Timeout) and not claimed_worker:
            owner_path = root(cfg) / "worker.owner.json"
            owner = json.loads(owner_path.read_text(encoding="utf-8")) if owner_path.is_file() else {}
            if owner.get("job_id") == job_id:
                # A duplicate launch must not overwrite the active owner's job.
                return
        record.update(state="failed", error_type=type(exc).__name__, updated_at=time.time(), finished_at=time.time())
        write(path, record)
