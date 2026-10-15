"""WorkbookService: the single write path for workbook revisions (manual edits, agent, template fill).

Layout: data/workbooks/{wb_id}/
    metadata.json          current revision, revision list
    revisions/{rev}.xlsx   immutable canonical files (rev 1 = original upload / generated file)
    previews/{rev}.json    cached FortuneSheet DTO
    ops/{operation_id}.json  idempotency records
    history.jsonl          append-only audit log
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import threading
import warnings
from pathlib import Path
from typing import Any

from . import ooxml_edit, preview, storage, upload, validate_edit
from .errors import InvalidOperation, NotFound, RevisionConflict, UnsupportedFile, ValidationFailed
from .schemas import OperationBatch, RevisionResult, ValidationSummary

_locks: dict[str, threading.Lock] = {}
_locks_guard = threading.Lock()


def _lock(wb_id: str) -> threading.Lock:
    with _locks_guard:
        return _locks.setdefault(wb_id, threading.Lock())


def _dir(wb_id: str) -> Path:
    try:
        storage.check_id(wb_id, "wb")
    except KeyError:
        raise NotFound("Workbook not found.") from None
    d = storage.workbooks_dir() / wb_id
    if not (d / "metadata.json").exists():
        raise NotFound("Workbook not found.")
    return d


def _open_check(path: Path) -> list[str]:
    """L4-style smoke check for a new rev 1: openpyxl must load it (never saved)."""
    from openpyxl import load_workbook
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(path, read_only=True)
            names = list(wb.sheetnames)
            wb.close()
    except Exception as e:  # noqa: BLE001
        raise UnsupportedFile(f"The workbook could not be read ({type(e).__name__}).") from None
    return names


def get_metadata(wb_id: str) -> dict:
    return storage.read_json(_dir(wb_id) / "metadata.json")


def update_metadata(wb_id: str, **fields: Any) -> dict:
    """Adds informational fields (warnings, reports); revision bookkeeping keys cannot be changed here."""
    if set(fields) & {"id", "revisions", "current_revision", "sha256", "created"}:
        raise ValueError("protected metadata field")
    with _lock(wb_id):
        meta = get_metadata(wb_id)
        meta.update(fields)
        storage.atomic_write_json(_dir(wb_id) / "metadata.json", meta)
        return meta


def revision_path(wb_id: str, rev: int | None = None) -> Path:
    meta = get_metadata(wb_id)
    rev = meta["current_revision"] if rev is None else rev
    p = _dir(wb_id) / "revisions" / f"{int(rev)}.xlsx"
    if not p.exists():
        raise NotFound(f"Revision {rev} not found.")
    return p


def _public(meta: dict) -> dict:
    keys = ("id", "name", "filename", "created", "updated", "current_revision", "origin", "workspace_id",
            "sheets", "llm_mode", "warnings")
    out = {k: meta.get(k) for k in keys}
    out["revisions"] = [{k: r.get(k) for k in ("revision", "parent", "created", "origin", "note", "changed_count",
                                                "restored")} for r in meta.get("revisions", [])]
    out["can_undo"] = _undo_target(meta, meta["current_revision"]) is not None
    return out


def public_metadata(wb_id: str) -> dict:
    return _public(get_metadata(wb_id))


def list_workbooks(limit: int = 50) -> list[dict]:
    d = storage.workbooks_dir()
    if not d.exists():
        return []
    metas = []
    for p in d.iterdir():
        m = storage.read_json(p / "metadata.json")
        if m:
            metas.append({k: m.get(k) for k in ("id", "name", "filename", "updated", "current_revision", "origin")})
    return sorted(metas, key=lambda m: m.get("updated") or 0, reverse=True)[:limit]


def create_from_path(path: Path | str, filename: str, origin: str, workspace_id: str | None = None,
                     extra: dict[str, Any] | None = None) -> dict:
    data = Path(path).read_bytes()
    return create_from_bytes(data, filename, origin, workspace_id, extra)


def create_from_bytes(data: bytes, filename: str, origin: str, workspace_id: str | None = None,
                      extra: dict[str, Any] | None = None) -> dict:
    upload.check_xlsx(data)
    wb_id = storage.new_id("wb")
    d = storage.workbooks_dir() / wb_id
    rev_path = d / "revisions" / "1.xlsx"
    storage.atomic_write_bytes(rev_path, data)
    try:
        sheets = _open_check(rev_path)
    except UnsupportedFile:
        shutil.rmtree(d, ignore_errors=True)
        raise
    name = Path(filename).stem[:120] or "Workbook"
    meta = {"id": wb_id, "name": name, "filename": Path(filename).name[:120], "created": storage.now(),
            "updated": storage.now(), "current_revision": 1, "origin": origin, "workspace_id": workspace_id,
            "sheets": sheets, "sha256": storage.sha256(data),
            "revisions": [{"revision": 1, "parent": None, "created": storage.now(), "origin": origin,
                           "note": "original" if origin == "upload" else f"created ({origin})", "changed_count": 0,
                           "restored": None}]}
    meta.update(extra or {})
    storage.atomic_write_json(d / "metadata.json", meta)
    storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "created", "revision": 1, "origin": origin})
    return meta


def get_preview(wb_id: str, rev: int | None = None) -> dict:
    meta = get_metadata(wb_id)
    rev = meta["current_revision"] if rev is None else int(rev)
    cache = _dir(wb_id) / "previews" / f"{rev}.json"
    dto = storage.read_json(cache)
    if dto is None:
        dto = preview.build_preview(str(revision_path(wb_id, rev)))
        storage.atomic_write_json(cache, dto)
    return {"workbook_id": wb_id, "revision": rev, **dto}


def _batch_hash(batch: OperationBatch) -> str:
    body = json.dumps([a.model_dump() for a in batch.actions], sort_keys=True, default=str)
    return hashlib.sha256(body.encode()).hexdigest()


def _ranges(changed: list[str]) -> list[str]:
    return sorted(set(changed))[:200]


def apply_operations(wb_id: str, batch: OperationBatch) -> RevisionResult:
    d = _dir(wb_id)
    with _lock(wb_id):
        meta = get_metadata(wb_id)
        op_file = d / "ops" / f"{hashlib.sha1(batch.operation_id.encode()).hexdigest()}.json"
        prior = storage.read_json(op_file)
        if prior:
            if prior["hash"] != _batch_hash(batch):
                raise InvalidOperation("operation_id was already used for a different batch.")
            res = RevisionResult(**prior["result"])
            res.idempotent_replay = True
            return res
        cur = meta["current_revision"]
        if batch.base_revision != cur:
            raise RevisionConflict(f"Workbook changed: you edited revision {batch.base_revision}, "
                                   f"current is {cur}. Reload to continue.", current_revision=cur)
        src = revision_path(wb_id, cur)
        fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp.xlsx")
        os.close(fd)
        try:
            stats = ooxml_edit.apply_actions(str(src), tmp, list(batch.actions))
            if not stats.touched_parts:
                raise InvalidOperation("The operations do not change the workbook.")
            report = validate_edit.validate(str(src), tmp, stats)
            if not report["pass"]:
                storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "rejected",
                                                           "base_revision": cur, "operation_id": batch.operation_id,
                                                           "errors": report["errors"][:10]})
                raise ValidationFailed("The edit failed validation and was not saved.", errors=report["errors"][:10])
            new_rev = cur + 1
            os.replace(tmp, d / "revisions" / f"{new_rev}.xlsx")
        finally:
            Path(tmp).unlink(missing_ok=True)
        entry = {"revision": new_rev, "parent": cur, "created": storage.now(), "origin": batch.origin,
                 "note": batch.note, "changed_count": len(set(stats.changed)), "restored": None,
                 "operation_id": batch.operation_id, "touched_parts": sorted(stats.touched_parts)}
        meta["revisions"].append(entry)
        meta["current_revision"] = new_rev
        meta["updated"] = storage.now()
        storage.atomic_write_json(d / "metadata.json", meta)
        storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "revision", **entry,
                                                   "changed": _ranges(stats.changed)[:50],
                                                   "actions": [{"type": a.type, "sheet": a.sheet, "cell": a.cell}
                                                               for a in batch.actions][:200]})   # no cell values
        result = RevisionResult(workbook_id=wb_id, revision=new_rev, parent_revision=cur, status="validated",
                                changed_ranges=_ranges(stats.changed),
                                validation=ValidationSummary(**{"pass": True, "level": report["level"],
                                                                "warnings": report["warnings"],
                                                                "checks": report["checks"]}),
                                preview_url=f"/api/workbooks/{wb_id}/preview?revision={new_rev}")
        storage.atomic_write_json(op_file, {"hash": _batch_hash(batch), "result": result.model_dump(by_alias=True)})
        return result


def apply_transform(wb_id: str, base_revision: int, operation_id: str, fn, note: str = "",
                    origin: str = "system") -> RevisionResult:
    """Structural write path (template restructuring): fn(src, tmp) writes the new file and returns
    (changed refs, validation report). Same revision/lock/idempotency/audit rules as apply_operations."""
    d = _dir(wb_id)
    with _lock(wb_id):
        meta = get_metadata(wb_id)
        op_file = d / "ops" / f"{hashlib.sha1(operation_id.encode()).hexdigest()}.json"
        prior = storage.read_json(op_file)
        if prior:
            res = RevisionResult(**prior["result"])
            res.idempotent_replay = True
            return res
        cur = meta["current_revision"]
        if base_revision != cur:
            raise RevisionConflict(f"Workbook changed: you edited revision {base_revision}, "
                                   f"current is {cur}. Reload to continue.", current_revision=cur)
        src = revision_path(wb_id, cur)
        fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp.xlsx")
        os.close(fd)
        try:
            changed, report = fn(str(src), tmp)
            if not report["pass"]:
                storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "rejected",
                                                           "base_revision": cur, "operation_id": operation_id,
                                                           "errors": report["errors"][:10]})
                raise ValidationFailed("The restructured workbook failed validation and was not saved.",
                                       errors=report["errors"][:10])
            new_rev = cur + 1
            os.replace(tmp, d / "revisions" / f"{new_rev}.xlsx")
        finally:
            Path(tmp).unlink(missing_ok=True)
        entry = {"revision": new_rev, "parent": cur, "created": storage.now(), "origin": origin, "note": note,
                 "changed_count": len(set(changed)), "restored": None, "operation_id": operation_id,
                 "touched_parts": ["structure"]}
        meta["revisions"].append(entry)
        meta["current_revision"] = new_rev
        meta["updated"] = storage.now()
        storage.atomic_write_json(d / "metadata.json", meta)
        storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "revision", **entry,
                                                   "changed": _ranges(changed)[:50]})
        result = RevisionResult(workbook_id=wb_id, revision=new_rev, parent_revision=cur, status="validated",
                                changed_ranges=_ranges(changed),
                                validation=ValidationSummary(**{"pass": True, "level": report["level"],
                                                                "warnings": report["warnings"],
                                                                "checks": report["checks"]}),
                                preview_url=f"/api/workbooks/{wb_id}/preview?revision={new_rev}")
        storage.atomic_write_json(op_file, {"hash": operation_id, "result": result.model_dump(by_alias=True)})
        return result


def _undo_target(meta: dict, rev: int) -> int | None:
    by = {r["revision"]: r for r in meta["revisions"]}
    seen = set()
    while rev in by and rev not in seen:
        seen.add(rev)
        r = by[rev]
        if r.get("restored"):
            rev = r["restored"]          # undoing an undo walks further back through history
            continue
        return r.get("parent")
    return None


def undo(wb_id: str, base_revision: int, target_revision: int | None = None) -> RevisionResult:
    d = _dir(wb_id)
    with _lock(wb_id):
        meta = get_metadata(wb_id)
        cur = meta["current_revision"]
        if base_revision != cur:
            raise RevisionConflict(f"Workbook changed: current revision is {cur}.", current_revision=cur)
        target = target_revision or _undo_target(meta, cur)
        if not target or target >= cur + 1 or target < 1:
            raise InvalidOperation("Nothing to undo.")
        src = revision_path(wb_id, target)
        new_rev = cur + 1
        tmp = d / "revisions" / f".{new_rev}.tmp"
        shutil.copyfile(src, tmp)
        os.replace(tmp, d / "revisions" / f"{new_rev}.xlsx")
        entry = {"revision": new_rev, "parent": cur, "created": storage.now(), "origin": "system",
                 "note": f"undo: restored revision {target}", "changed_count": 0, "restored": target}
        meta["revisions"].append(entry)
        meta["current_revision"] = new_rev
        meta["updated"] = storage.now()
        storage.atomic_write_json(d / "metadata.json", meta)
        storage.append_jsonl(d / "history.jsonl", {"t": storage.now(), "event": "undo", **entry})
        return RevisionResult(workbook_id=wb_id, revision=new_rev, parent_revision=cur, status="validated",
                              changed_ranges=[], validation=ValidationSummary(**{"pass": True, "level": "copy"}),
                              preview_url=f"/api/workbooks/{wb_id}/preview?revision={new_rev}")


def history(wb_id: str, limit: int = 200) -> list[dict]:
    p = _dir(wb_id) / "history.jsonl"
    if not p.exists():
        return []
    lines = p.read_text(encoding="utf-8").splitlines()[-limit:]
    return [json.loads(x) for x in lines if x.strip()]


def export_info(wb_id: str, rev: int | None = None) -> tuple[Path, str]:
    meta = get_metadata(wb_id)
    rev = meta["current_revision"] if rev is None else int(rev)
    path = revision_path(wb_id, rev)
    name = meta.get("name") or "workbook"
    safe = "".join(ch if ch.isalnum() or ch in " -_." else "_" for ch in name).strip() or "workbook"
    return path, f"{safe}.xlsx" if rev == 1 and meta.get("origin") != "upload" else f"{safe}_rev{rev}.xlsx"
