"""Windows helper ownership even when the MCP transport is forcibly killed."""
from __future__ import annotations

import os
import uuid
from typing import Any


JOB_ENV = "_OUTLOOK_RAG_HELPER_JOB"


def create_job() -> tuple[str, Any]:
    import win32job

    name = "outlook-rag-helper-" + uuid.uuid4().hex
    job = win32job.CreateJobObject(None, name)
    try:
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits["BasicLimitInformation"]["LimitFlags"] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
    except BaseException:
        job.Close()
        raise
    return name, job


def attach_process(job: Any, pid: int) -> None:
    import win32api
    import win32con
    import win32job

    process = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_INFORMATION, False, pid)
    try:
        if not win32job.IsProcessInJob(process, job):
            win32job.AssignProcessToJobObject(job, process)
    finally:
        process.Close()


def join_job() -> None:
    """Register the actual interpreter, including a venv launcher's child.

    Close this temporary handle immediately: the transport must be the sole
    owner so Windows kills every helper when its last handle disappears.
    Register here rather than assigning a venv launcher after spawn: its
    interpreter may already have started, making nested job order race.
    """
    if os.name != "nt" or not (name := os.environ.pop(JOB_ENV, None)):
        return
    import win32api
    import win32job

    job = win32job.OpenJobObject(win32job.JOB_OBJECT_ASSIGN_PROCESS | win32job.JOB_OBJECT_QUERY, False, name)
    try:
        if not win32job.IsProcessInJob(win32api.GetCurrentProcess(), job):
            win32job.AssignProcessToJobObject(job, win32api.GetCurrentProcess())
    finally:
        job.Close()
