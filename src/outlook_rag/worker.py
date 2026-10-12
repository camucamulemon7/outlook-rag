"""Stdlib-only detached bootstrap: diagnose imports, config and WMI transfer."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
import threading
import time

# Direct execution works even when importing app (and its dependencies) fails.
if __package__:
    from . import diagnostics
else:
    import diagnostics as direct_diagnostics

    diagnostics = direct_diagnostics


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pipe")
    parser.add_argument("--authkey")
    parser.add_argument("--config", required=True)
    parser.add_argument("command", choices=["_job-worker"])
    parser.add_argument("job_id")
    args = parser.parse_args()
    # The snapshot is a credential-free, parent-generated file in the jobs dir.
    cfg = {"data_dir": str(Path(args.config).resolve().parent.parent)}
    diagnostics.log_path(cfg, args.job_id)
    diagnostics.append(cfg, args.job_id, "bootstrap_started")
    original_streams = sys.stdout, sys.stderr
    timer = None
    try:
        with open(os.devnull, "w") as quiet:
            sys.stdout = sys.stderr = quiet
            if args.pipe:
                from multiprocessing.connection import Listener

                def timed_out() -> None:
                    diagnostics.append(
                        cfg,
                        args.job_id,
                        "bootstrap_timeout",
                        error_detail="Worker environment transfer timed out.",
                    )
                    os._exit(1)

                timer = threading.Timer(15, timed_out)
                timer.daemon = True
                timer.start()
                with Listener(
                    args.pipe, family="AF_PIPE", authkey=bytes.fromhex(args.authkey)
                ) as listener:
                    with listener.accept() as connection:
                        if not connection.poll(15):
                            raise TimeoutError("Environment transfer timed out")
                        environment = connection.recv()
                        os.environ.clear()
                        os.environ.update(environment)
                timer.cancel()
            os.environ["OUTLOOK_RAG_CONFIG"] = args.config
            from outlook_rag import app, jobs

            resolved = app.config()
            diagnostics.append(cfg, args.job_id, "bootstrap_ready")
            jobs.run(args.job_id, resolved)
        return 0
    except BaseException as exc:
        info = diagnostics.exception_info(exc)
        diagnostics.append(cfg, args.job_id, "bootstrap_failed", **info)
        path = Path(args.config).parent / (args.job_id + ".job.json")
        if path.is_file():
            record = json.loads(path.read_text(encoding="utf-8"))
            if record["state"] in {"starting", "running", "waiting"}:
                record.update(
                    state="failed",
                    error_type=info["error_type"],
                    error_detail=info["error_detail"],
                    finished_at=time.time(),
                    updated_at=time.time(),
                )
                temporary = path.with_suffix(".bootstrap.tmp")
                temporary.write_text(json.dumps(record), encoding="utf-8")
                os.replace(temporary, path)
        return 1
    finally:
        sys.stdout, sys.stderr = original_streams
        if timer is not None:
            timer.cancel()


if __name__ == "__main__":
    raise SystemExit(main())
