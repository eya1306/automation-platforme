"""Job execution.

A run is a :class:`Job`: a folder on disk holding its inputs and outputs, plus
an in-memory record of its status and the progress lines the engine emitted.
Runs execute on a small thread pool so the browser never waits on a request,
and the page polls for status.

The engines are ordinary blocking Python and are called exactly as their
authors wrote them. Everything asynchronous lives here.

Some engines report their progress by printing. While a run is executing, the
worker thread's stdout is routed into that run's log so those lines show up in
the run panel instead of on the server console -- see :class:`_StdoutRouter`.
"""

from __future__ import annotations

import shutil
import sys
import threading
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from . import config
from .registry import Artifact, RunResult, ToolSpec

QUEUED = "queued"
RUNNING = "running"
DONE = "done"
FAILED = "failed"


def _now() -> datetime:
    return datetime.now(timezone.utc)



class _StdoutRouter:
    """Sends each worker thread's ``print`` output to the run it belongs to.

    Installed once over ``sys.stdout``. A thread that has claimed a job gets
    its writes appended to that job's log; every other thread -- the Flask
    request threads, the console -- passes straight through to the real
    stdout, so this is safe with several runs going at once.
    """

    def __init__(self, original):
        self._original = original
        self._jobs: Dict[int, "Job"] = {}
        self._buffers: Dict[int, str] = {}
        self._lock = threading.Lock()

    # -- claim / release ---------------------------------------------------- #

    def claim(self, job: "Job") -> None:
        with self._lock:
            self._jobs[threading.get_ident()] = job

    def release(self) -> None:
        ident = threading.get_ident()
        with self._lock:
            job = self._jobs.pop(ident, None)
            leftover = self._buffers.pop(ident, "")
        # A last line the engine printed without a trailing newline.
        if job is not None and leftover.strip():
            job.say(leftover.rstrip())

    # -- file protocol ------------------------------------------------------ #

    def write(self, text: str) -> int:
        ident = threading.get_ident()
        with self._lock:
            job = self._jobs.get(ident)
        if job is None:
            return self._original.write(text)

        # Engines print line by line; buffer until a newline so one log entry
        # is one line of output.
        with self._lock:
            pending = self._buffers.get(ident, "") + text
            lines = pending.split("\n")
            self._buffers[ident] = lines.pop()
        for line in lines:
            if line.strip():
                job.say(line.rstrip())
        return len(text)

    def flush(self) -> None:
        self._original.flush()

    def isatty(self) -> bool:
        return False

    def __getattr__(self, name):
        return getattr(self._original, name)


_stdout_router = _StdoutRouter(sys.stdout)
sys.stdout = _stdout_router


@dataclass
class LogLine:
    at: str
    text: str
    level: str = "info"      # info | warn | error | done

    def as_json(self) -> dict:
        return {"at": self.at, "text": self.text, "level": self.level}


@dataclass
class Job:
    id: str
    tool_id: str
    tool_name: str
    label: str                       # what the user sees in the history rail
    status: str = QUEUED
    progress: float = 0.0
    created_at: datetime = field(default_factory=_now)
    started_at: Optional[datetime] = None
    finished_at: Optional[datetime] = None
    log: List[LogLine] = field(default_factory=list)
    result: Optional[RunResult] = None
    error: Optional[str] = None
    inputs: List[str] = field(default_factory=list)     # display names of uploads
    options: Dict[str, Any] = field(default_factory=dict)

    _lock: threading.Lock = field(default_factory=threading.Lock, repr=False)

    # -- directories ------------------------------------------------------- #

    @property
    def dir(self) -> Path:
        return config.JOBS_DIR / self.id

    @property
    def input_dir(self) -> Path:
        return self.dir / "input"

    @property
    def output_dir(self) -> Path:
        return self.dir / "output"

    # -- progress ---------------------------------------------------------- #

    def say(self, message: str, fraction: Optional[float] = None,
            level: str = "info") -> None:
        text = str(message).strip()
        if not text:
            return
        with self._lock:
            self.log.append(LogLine(_now().strftime("%H:%M:%S"), text, level))
            if fraction is not None:
                self.progress = max(self.progress, min(float(fraction), 1.0))

    @property
    def duration(self) -> Optional[float]:
        if not self.started_at:
            return None
        end = self.finished_at or _now()
        return (end - self.started_at).total_seconds()

    # -- serialisation ----------------------------------------------------- #

    def as_json(self, since: int = 0) -> dict:
        with self._lock:
            lines = [line.as_json() for line in self.log[since:]]
            total_lines = len(self.log)
        data = {
            "id": self.id,
            "tool_id": self.tool_id,
            "tool_name": self.tool_name,
            "label": self.label,
            "status": self.status,
            "progress": round(self.progress, 3),
            "created_at": self.created_at.isoformat(),
            "duration": round(self.duration, 1) if self.duration else None,
            "log": lines,
            "log_count": total_lines,
            "inputs": self.inputs,
            "error": self.error,
            "artifacts": [],
            "stats": [],
            "table": None,
            "notes": [],
        }
        if self.result:
            data["artifacts"] = [
                {
                    "name": a.path.name,
                    "label": a.label,
                    "role": a.role,
                    "size": a.path.stat().st_size if a.path.exists() else 0,
                    "url": f"/api/jobs/{self.id}/files/{a.path.name}",
                }
                for a in self.result.artifacts
            ]
            data["stats"] = self.result.stats
            data["table"] = self.result.table
            data["notes"] = self.result.notes
        return data


class JobStore:
    """Holds every run for the life of the process and executes them."""

    def __init__(self) -> None:
        self._jobs: Dict[str, Job] = {}
        self._order: List[str] = []
        self._lock = threading.Lock()
        self._pool = ThreadPoolExecutor(
            max_workers=config.MAX_WORKERS, thread_name_prefix="run"
        )

    # -- lookup ------------------------------------------------------------ #

    def get(self, job_id: str) -> Optional[Job]:
        return self._jobs.get(job_id)

    def recent(self, limit: int = config.HISTORY_LIMIT) -> List[Job]:
        with self._lock:
            ids = list(reversed(self._order))[:limit]
        return [self._jobs[i] for i in ids if i in self._jobs]

    # -- creation ---------------------------------------------------------- #

    def create(self, tool: ToolSpec, label: str) -> Job:
        job_id = uuid.uuid4().hex[:12]
        job = Job(id=job_id, tool_id=tool.id, tool_name=tool.name, label=label)
        job.input_dir.mkdir(parents=True, exist_ok=True)
        job.output_dir.mkdir(parents=True, exist_ok=True)
        with self._lock:
            self._jobs[job_id] = job
            self._order.append(job_id)
        return job

    def submit(self, job: Job, tool: ToolSpec, values: Dict[str, Any]) -> None:
        self._pool.submit(self._execute, job, tool, values)

    # -- execution --------------------------------------------------------- #

    def _execute(self, job: Job, tool: ToolSpec, values: Dict[str, Any]) -> None:
        job.status = RUNNING
        job.started_at = _now()
        job.say(f"Starting {tool.name}.", 0.02)
        try:
            runner = tool.runner()
            _stdout_router.claim(job)
            try:
                result = runner(values, job.output_dir, job.say)
            finally:
                _stdout_router.release()
            # Keep only artifacts that really landed on disk.
            result.artifacts = [a for a in result.artifacts if a.path.exists()]
            job.result = result
            job.progress = 1.0
            job.status = DONE
            job.say("Finished.", 1.0, level="done")
        except Exception as exc:                      # noqa: BLE001 - surfaced to the user
            job.status = FAILED
            job.error = f"{type(exc).__name__}: {exc}"
            job.say(job.error, level="error")
            detail = traceback.format_exc().strip().splitlines()
            for line in detail[-12:]:
                job.say(line, level="error")
        finally:
            job.finished_at = _now()

    # -- housekeeping ------------------------------------------------------ #

    def sweep(self) -> int:
        """Delete finished runs past the retention window. Returns how many went."""
        cutoff = _now() - timedelta(hours=config.RETENTION_HOURS)
        removed = 0
        with self._lock:
            stale = [
                job_id
                for job_id, job in self._jobs.items()
                if job.status in (DONE, FAILED)
                and job.finished_at
                and job.finished_at < cutoff
            ]
            for job_id in stale:
                job = self._jobs.pop(job_id, None)
                if job_id in self._order:
                    self._order.remove(job_id)
                if job:
                    shutil.rmtree(job.dir, ignore_errors=True)
                removed += 1
        return removed

    def discard(self, job_id: str) -> bool:
        with self._lock:
            job = self._jobs.pop(job_id, None)
            if job_id in self._order:
                self._order.remove(job_id)
        if not job:
            return False
        shutil.rmtree(job.dir, ignore_errors=True)
        return True


store = JobStore()
