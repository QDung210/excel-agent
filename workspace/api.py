"""HTTP routes (spec 05). Errors follow {code, message, ...}; ids are opaque, never paths."""
from __future__ import annotations

import json
from typing import Any

from fastapi import APIRouter, File, Form, Header, Query, Request, UploadFile
from fastapi.responses import FileResponse, StreamingResponse

from . import clarify, config, extractor_adapter, generate, preview, runs, service, templates, upload, workspaces
from .errors import FileTooLarge, InvalidOperation, NotFound
from .schemas import (AgentTurnRequest, AnalyzeRequest, ClarifyAnswers, CreateWorkspace, ExtractRequest, OperationBatch,
                      PlanApply, UndoRequest)

router = APIRouter(prefix="/api")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


async def _read_upload(file: UploadFile) -> bytes:
    limit = config.max_upload_bytes()
    data = await file.read(limit + 1)
    if len(data) > limit:
        raise FileTooLarge(f"File is larger than {limit // (1024 * 1024)} MB.")
    return data


def _sse(run_id: str, last_event_id: str | None) -> StreamingResponse:
    runs.get(run_id)                                   # 404 before streaming starts
    try:
        start = int(last_event_id or 0)
    except ValueError:
        start = 0
    return StreamingResponse(runs.sse(run_id, start), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ---------------------------------------------------------------- health
@router.get("/health")
def health() -> dict:
    return {"ok": True, "llm_mode": config.llm_mode(), "model": config.active_model(),
            "max_upload_mb": config.max_upload_bytes() // (1024 * 1024),
            "extractor_available": extractor_adapter.available()}


# ---------------------------------------------------------------- workspaces
@router.post("/workspaces")
def create_workspace(body: CreateWorkspace) -> dict:
    return workspaces.create(body.mode, body.name)


@router.get("/workspaces/{ws_id}")
def get_workspace(ws_id: str) -> dict:
    return workspaces.get(ws_id)


@router.post("/workspaces/{ws_id}/files")
async def add_file(ws_id: str, file: UploadFile = File(...), role: str = Form("source")) -> dict:
    workspaces.get(ws_id)
    data = await _read_upload(file)
    return workspaces.attach(ws_id, data, file.filename or "", role)


@router.delete("/workspaces/{ws_id}/files/{file_id}")
def remove_file(ws_id: str, file_id: str) -> dict:
    return workspaces.detach(ws_id, file_id)


@router.get("/workspaces/{ws_id}/files/{file_id}/preview")
def file_preview(ws_id: str, file_id: str) -> dict:
    """Read-only grid preview of an .xlsx attached to this workspace (nothing is written)."""
    f = next((x for x in workspaces.get(ws_id)["files"] if x["id"] == file_id), None)
    if f is None or f.get("kind") != "xlsx":
        raise NotFound("No .xlsx file with that id is attached to this workspace.")
    return preview.build_preview(str(upload.get(file_id).path))


# ---------------------------------------------------------------- runs
@router.post("/workspaces/{ws_id}/analyze")
def analyze_workspace(ws_id: str, body: AnalyzeRequest) -> dict:
    return clarify.analyze(ws_id, body).model_dump()


@router.post("/workspaces/{ws_id}/clarify")
def clarify_workspace(ws_id: str, body: ClarifyAnswers) -> dict:
    return clarify.clarify(ws_id, body).model_dump()


@router.post("/workspaces/{ws_id}/generate")
def generate_workspace(ws_id: str) -> dict:
    return {"run_id": generate.start(ws_id).id}


@router.get("/runs/{run_id}")
def get_run(run_id: str) -> dict:
    return runs.get(run_id).public()


@router.get("/runs/{run_id}/events")
def run_events(run_id: str, last_event_id: str | None = Header(default=None)) -> StreamingResponse:
    return _sse(run_id, last_event_id)


# ---------------------------------------------------------------- workbooks
# ---------------------------------------------------------------- templates gallery
@router.get("/templates")
def list_templates() -> list[dict]:
    return templates.list_all()


@router.post("/templates")
async def add_template(file: UploadFile = File(...), name: str = Form(""), description: str = Form("")) -> dict:
    from . import upload
    data = await _read_upload(file)
    upload.kind_of(file.filename or "", upload.XLSX_EXT)
    meta = templates.add(data, name.strip() or (file.filename or "Template").rsplit(".", 1)[0], description.strip())
    return {k: meta[k] for k in ("id", "name", "description", "origin", "sheets", "created")}


@router.get("/templates/{tpl_id}/preview")
def template_preview(tpl_id: str) -> dict:
    return preview.build_preview(str(templates.path(tpl_id)))


@router.get("/templates/{tpl_id}/slots")
def template_slots(tpl_id: str) -> dict:
    return templates.detect_slots(templates.path(tpl_id)).summary()


@router.post("/templates/extract")
def extract_template(body: ExtractRequest) -> dict:
    return {"run_id": extractor_adapter.start(body.file_id, body.template_id).id}


@router.post("/workbooks/upload")
async def upload_workbook(file: UploadFile = File(...)) -> dict:
    from . import upload
    data = await _read_upload(file)
    upload.kind_of(file.filename or "", upload.XLSX_EXT)
    meta = service.create_from_bytes(data, file.filename or "workbook.xlsx", origin="upload")
    return service.public_metadata(meta["id"])


@router.get("/workbooks")
def list_workbooks() -> list[dict]:
    return service.list_workbooks()


@router.get("/workbooks/{wb_id}/metadata")
def workbook_metadata(wb_id: str) -> dict:
    return service.public_metadata(wb_id)


@router.get("/workbooks/{wb_id}/preview")
def workbook_preview(wb_id: str, revision: int | None = Query(default=None, ge=1)) -> dict:
    return service.get_preview(wb_id, revision)


@router.post("/workbooks/{wb_id}/operations")
def workbook_operations(wb_id: str, batch: OperationBatch) -> dict:
    return service.apply_operations(wb_id, batch).model_dump(by_alias=True)


@router.post("/workbooks/{wb_id}/undo")
def workbook_undo(wb_id: str, body: UndoRequest) -> dict:
    return service.undo(wb_id, body.base_revision, body.target_revision).model_dump(by_alias=True)


@router.get("/workbooks/{wb_id}/history")
def workbook_history(wb_id: str) -> list[dict]:
    return service.history(wb_id)


@router.get("/workbooks/{wb_id}/export")
def workbook_export(wb_id: str, revision: int | None = Query(default=None, ge=1)) -> FileResponse:
    path, name = service.export_info(wb_id, revision)
    return FileResponse(path, filename=name, media_type=XLSX_MIME, headers={"Cache-Control": "no-store"})


# ---------------------------------------------------------------- agent chat (Phase 4)
@router.post("/agent/turns")
def agent_turn(body: AgentTurnRequest) -> dict:
    from .agent import runtime
    return runtime.start_turn(body)


@router.get("/agent/turns/{turn_id}/events")
def agent_turn_events(turn_id: str, last_event_id: str | None = Header(default=None)) -> StreamingResponse:
    return _sse(turn_id, last_event_id)


@router.post("/agent/plans/{plan_id}/apply")
def agent_plan_apply(plan_id: str, body: PlanApply) -> dict:
    from .agent import tools
    return tools.apply_plan(plan_id, body.base_revision).model_dump(by_alias=True)


@router.post("/agent/plans/{plan_id}/reject")
def agent_plan_reject(plan_id: str) -> dict:
    from .agent import tools
    return tools.reject_plan(plan_id)
