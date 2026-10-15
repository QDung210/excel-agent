"""Background runs with an ordered event log, streamed as SSE (sequence ids, Last-Event-ID resume)."""
from __future__ import annotations

import json
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Any, Callable, Iterator

from . import storage
from .errors import NotFound, WorkspaceError

TERMINAL = {"done", "error", "cancelled", "validation_failed"}
_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="run")
_runs: dict[str, "Run"] = {}
_guard = threading.Lock()


@dataclass
class Run:
    id: str
    kind: str
    status: str = "queued"
    events: list[dict] = field(default_factory=list)
    result: Any = None
    cond: threading.Condition = field(default_factory=threading.Condition)
    cancel: bool = False
    created: float = field(default_factory=storage.now)

    def emit(self, event: str, data: dict | None = None) -> None:
        with self.cond:
            ev = {"event": event, "run_id": self.id, "seq": len(self.events) + 1, "data": data or {}}
            self.events.append(ev)
            if event in TERMINAL:
                self.status = event if event != "validation_failed" else "error"
            self.cond.notify_all()

    def public(self) -> dict:
        return {"id": self.id, "kind": self.kind, "status": self.status, "result": self.result,
                "events": len(self.events), "created": self.created}


Emit = Callable[[str, dict | None], None]


def start(kind: str, fn: Callable[[Run], Any], prefix: str = "run") -> Run:
    run = Run(storage.new_id(prefix), kind)
    with _guard:
        _runs[run.id] = run
        if len(_runs) > 500:                       # bound memory: drop oldest finished runs
            for rid in [r.id for r in sorted(_runs.values(), key=lambda r: r.created) if r.status in TERMINAL][:100]:
                _runs.pop(rid, None)
    run.emit("run_started", {"kind": kind})

    def work() -> None:
        run.status = "running"
        try:
            run.result = fn(run)
            if run.status not in TERMINAL:
                run.emit("done", run.result if isinstance(run.result, dict) else {"result": run.result})
        except WorkspaceError as e:
            run.emit("error", e.body())
        except Exception as e:  # noqa: BLE001 - never leak traces to the client
            traceback.print_exc()
            run.emit("error", {"code": "INTERNAL", "message": f"{type(e).__name__}: {str(e)[:300]}"})

    _executor.submit(work)
    return run


def get(run_id: str) -> Run:
    run = _runs.get(run_id)
    if not run:
        raise NotFound("Run not found.")
    return run


def wait(run_id: str, timeout: float = 60) -> Run:
    """Test helper: block until the run reaches a terminal event."""
    run = get(run_id)
    end = time.time() + timeout
    with run.cond:
        while run.status not in TERMINAL and time.time() < end:
            run.cond.wait(0.2)
    return run


def sse(run_id: str, last_event_id: int = 0, heartbeat: float = 15.0) -> Iterator[str]:
    run = get(run_id)
    seq = last_event_id
    yield "retry: 2000\n\n"
    while True:
        with run.cond:
            if len(run.events) <= seq and run.status not in TERMINAL:
                run.cond.wait(heartbeat)
            pending = run.events[seq:]
            finished = run.status in TERMINAL
        if not pending and not finished:
            yield ": keep-alive\n\n"
            continue
        for ev in pending:
            seq = ev["seq"]
            yield f"id: {seq}\nevent: {ev['event']}\ndata: {json.dumps(ev, ensure_ascii=False, default=str)}\n\n"
        if finished and seq >= len(run.events):
            return
