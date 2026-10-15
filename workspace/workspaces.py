"""Workspace (create-flow session) state: mode, attached files, prompt, clarification, result workbook."""
from __future__ import annotations

import threading
from pathlib import Path

from . import storage, upload
from .errors import InvalidOperation, NotFound

_lock = threading.Lock()


def _path(ws_id: str) -> Path:
    try:
        storage.check_id(ws_id, "ws")
    except KeyError:
        raise NotFound("Workspace not found.") from None
    return storage.workspaces_dir() / f"{ws_id}.json"


def create(mode: str, name: str | None = None) -> dict:
    ws = {"id": storage.new_id("ws"), "mode": mode, "name": name or "", "created": storage.now(),
          "updated": storage.now(), "files": [], "template_id": None, "prompt": "", "clarification": None,
          "workbook_id": None, "runs": []}
    storage.atomic_write_json(_path(ws["id"]), ws)
    return ws


def get(ws_id: str) -> dict:
    ws = storage.read_json(_path(ws_id))
    if not ws:
        raise NotFound("Workspace not found.")
    return ws


def save(ws: dict) -> dict:
    ws["updated"] = storage.now()
    storage.atomic_write_json(_path(ws["id"]), ws)
    return ws


def attach(ws_id: str, data: bytes, filename: str, role: str) -> dict:
    if role not in ("source", "template"):
        raise InvalidOperation("role must be 'source' or 'template'.")
    allowed = upload.XLSX_EXT if role == "template" else upload.SOURCE_EXT
    with _lock:
        ws = get(ws_id)
        if role == "template" and ws["mode"] != "template":
            raise InvalidOperation("Templates can only be attached in template mode.")
        sources = [f for f in ws["files"] if f["role"] == "source"]
        if role == "source" and len(sources) >= 10:
            raise InvalidOperation("At most 10 source files per workspace.")
        f = upload.store(data, filename, allowed)
        entry = {**f.public(), "role": role}
        if role == "template":                    # one template per workspace: replacing detaches the old one
            ws["files"] = [x for x in ws["files"] if x["role"] != "template"]
            ws["template_id"] = None
        ws["files"].append(entry)
        save(ws)
    return entry


def detach(ws_id: str, file_id: str) -> dict:
    """Removes the file from the workspace only; the stored upload is kept (never delete user data)."""
    with _lock:
        ws = get(ws_id)
        before = len(ws["files"])
        ws["files"] = [f for f in ws["files"] if f["id"] != file_id]
        if len(ws["files"]) == before:
            raise NotFound("File is not attached to this workspace.")
        return save(ws)
