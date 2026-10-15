"""Filesystem persistence for chat threads and edit plans (data/agent/...)."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from .. import config, storage
from ..errors import NotFound

MAX_THREAD_MESSAGES = 20          # history replayed to the model per turn (bounded context)


def _dir(kind: str) -> Path:
    return config.data_dir() / "agent" / kind


def _path(kind: str, prefix: str, item_id: str) -> Path:
    try:
        storage.check_id(item_id, prefix)
    except KeyError:
        raise NotFound(f"{kind[:-1].capitalize()} not found.") from None
    return _dir(kind) / f"{item_id}.json"


# ---------------------------------------------------------------- threads
def get_thread(thread_id: str | None, workbook_id: str) -> dict:
    if thread_id:
        t = storage.read_json(_path("threads", "thr", thread_id))
        if not t:
            raise NotFound("Thread not found.")
        if t["workbook_id"] != workbook_id:
            raise NotFound("Thread belongs to another workbook.")
        return t
    return {"id": storage.new_id("thr"), "workbook_id": workbook_id, "created": storage.now(), "messages": []}


def save_thread(t: dict) -> None:
    t["messages"] = t["messages"][-200:]
    storage.atomic_write_json(_path("threads", "thr", t["id"]), t)


def recent_messages(t: dict) -> list[dict]:
    return t["messages"][-MAX_THREAD_MESSAGES:]


# ---------------------------------------------------------------- plans
def create_plan(**fields: Any) -> dict:
    plan = {"id": storage.new_id("plan"), "created": storage.now(), "status": "pending", **fields}
    storage.atomic_write_json(_path("plans", "plan", plan["id"]), plan)
    return plan


def get_plan(plan_id: str) -> dict:
    p = storage.read_json(_path("plans", "plan", plan_id))
    if not p:
        raise NotFound("Plan not found.")
    return p


def save_plan(plan: dict) -> dict:
    storage.atomic_write_json(_path("plans", "plan", plan["id"]), plan)
    return plan


def log(thread_id: str, entry: dict) -> None:
    """Tool-call diagnostics: names, statuses, revisions and cell addresses only (no cell values)."""
    storage.append_jsonl(_dir("logs") / f"{thread_id}.jsonl", {"t": storage.now(), **entry})
