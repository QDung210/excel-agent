"""Minimal web UI: upload .xlsx -> progress bar -> download template.

Run:  .venv\\Scripts\\python.exe -m xlsx_template.server   then open http://127.0.0.1:8000
"""
from __future__ import annotations

import os
import re
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from .graph import run

load_dotenv()

ROOT = Path(__file__).resolve().parent
JOBS_DIR = Path(os.environ.get("XLSX_TEMPLATE_JOBS", ROOT.parent / "jobs"))
MAX_UPLOAD_MB = 30

app = FastAPI(title="Excel → Template")
# One job at a time: the Groq free tier quota is per minute, parallel jobs would only fight over it.
_executor = ThreadPoolExecutor(max_workers=1)
_jobs: dict[str, dict] = {}
_lock = threading.Lock()


def _update(job_id: str, **kw) -> None:
    with _lock:
        _jobs[job_id].update(kw)


def _safe_name(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"[^\w\-. ()]+", "_", stem, flags=re.UNICODE).strip() or "workbook"
    return stem[:80]


class Cancelled(Exception):
    pass


def _work(job_id: str) -> None:
    job = _jobs[job_id]
    if job.get("cancel"):
        return
    _update(job_id, status="running", started=time.time(), message="Bắt đầu…", progress=1)

    def progress(pct, msg, cap):
        if job.get("cancel"):          # cooperative: stops at the next checkpoint
            raise Cancelled()
        _update(job_id, progress=pct, message=msg, cap=cap)

    try:
        state = run(job["input"], job["output"], progress=progress)
        v, ws = state["validation"], state["write_stats"]
        if not v["ok"]:
            _update(job_id, status="error", progress=100,
                    message="Template không qua bước kiểm tra, không xuất file.",
                    error="; ".join(v["errors"][:5]))
            return
        _update(job_id, status="done", progress=100, message="Hoàn tất",
                summary={"cleared": ws["cleared"], "formulas": v["checks"].get("formulas_preserved"),
                         "kept": v["checks"].get("constants_kept"),
                         "seconds": round(time.time() - job["started"])})
    except Cancelled:
        _update(job_id, status="cancelled", message="Đã huỷ")
    except Exception as e:  # surface the reason in the UI, full trace in the server log
        traceback.print_exc()
        _update(job_id, status="error", message="Xử lý thất bại", error=f"{type(e).__name__}: {e}")


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (ROOT / "web" / "index.html").read_text(encoding="utf-8")


@app.post("/api/jobs")
async def create_job(file: UploadFile = File(...)) -> dict:
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Chỉ hỗ trợ file .xlsx / .xlsm")
    data = await file.read()
    if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"File lớn hơn {MAX_UPLOAD_MB} MB")
    if not data.startswith(b"PK"):
        raise HTTPException(400, "File không phải là .xlsx hợp lệ")

    job_id = uuid.uuid4().hex[:12]
    d = JOBS_DIR / job_id
    d.mkdir(parents=True, exist_ok=True)
    ext = Path(file.filename).suffix.lower()
    inp = d / f"input{ext}"
    inp.write_bytes(data)
    name = _safe_name(file.filename)
    with _lock:
        queued = sum(1 for j in _jobs.values() if j["status"] in ("queued", "running"))
        _jobs[job_id] = {
            "status": "queued", "progress": 0, "created": time.time(), "started": None,
            "message": "Đang chờ tới lượt…" if queued else "Đang khởi động…",
            "input": str(inp), "output": str(d / "template.xlsx"),
            "download_name": f"{name}_template{ext if ext == '.xlsx' else '.xlsx'}",
            "filename": file.filename, "error": None, "summary": None,
        }
    _executor.submit(_work, job_id)
    return {"id": job_id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict:
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Không tìm thấy job")
    keys = ("status", "progress", "message", "error", "summary", "filename", "download_name")
    out = {k: job[k] for k in keys}
    out["cap"] = job.get("cap", job["progress"])
    if job["status"] == "queued":
        with _lock:
            ahead = sum(1 for j in _jobs.values()
                        if j["status"] in ("queued", "running") and j["created"] < job["created"])
        out["message"] = f"Đang chờ tới lượt: còn {ahead} file đang xử lý phía trước…" if ahead else "Đang khởi động…"
    out["elapsed"] = round(time.time() - job["started"]) if job["started"] else 0
    return out


@app.delete("/api/jobs/{job_id}")
def cancel_job(job_id: str) -> dict:
    """Queued jobs never start; a running job stops at its next progress checkpoint."""
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Không tìm thấy job")
    if job["status"] in ("queued", "running"):
        _update(job_id, cancel=True, message="Đang huỷ…")
        if job["status"] == "queued":
            _update(job_id, status="cancelled", message="Đã huỷ")
    return {"status": _jobs[job_id]["status"]}


@app.get("/api/jobs/{job_id}/download")
def download(job_id: str) -> FileResponse:
    job = _jobs.get(job_id)
    if not job or job["status"] != "done":
        raise HTTPException(404, "Template chưa sẵn sàng")
    return FileResponse(job["output"], filename=job["download_name"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/jobs/{job_id}/report")
def report(job_id: str) -> FileResponse:
    job = _jobs.get(job_id)
    path = Path(job["output"]).with_suffix(".report.md") if job else None
    if not job or job["status"] != "done" or not path.exists():
        raise HTTPException(404, "Report chưa sẵn sàng")
    return FileResponse(path, filename=job["download_name"].replace(".xlsx", ".report.md"),
                        media_type="text/markdown; charset=utf-8")


def main() -> None:
    import uvicorn
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    print(f"Open http://{host}:{port}")
    uvicorn.run(app, host=host, port=port)


if __name__ == "__main__":
    main()
