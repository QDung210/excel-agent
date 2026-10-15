"""ASGI app: workspace API + legacy extractor routes (/api/jobs*, UI at /extractor) + built SPA."""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from xlsx_template import server as legacy

from . import config
from .api import router
from .errors import WorkspaceError

DIST = config.ROOT / "frontend" / "dist"


def create_app() -> FastAPI:
    app = FastAPI(title="AI Excel Workspace")

    @app.exception_handler(WorkspaceError)
    async def _ws_error(_: Request, exc: WorkspaceError) -> JSONResponse:
        return JSONResponse(status_code=exc.status, content=exc.body())

    @app.exception_handler(RequestValidationError)
    async def _validation(_: Request, exc: RequestValidationError) -> JSONResponse:
        errs = [{"loc": [str(x) for x in e.get("loc", [])], "msg": e.get("msg")} for e in exc.errors()][:20]
        return JSONResponse(status_code=422, content={"code": "INVALID_REQUEST", "message": "Request is invalid.",
                                                      "errors": errs})

    app.include_router(router)

    # legacy extractor: same /api/jobs* routes and behaviour, UI moved from / to /extractor
    for route in legacy.app.router.routes:
        if getattr(route, "path", "") == "/":
            continue
        if getattr(route, "path", "").startswith("/api/jobs"):
            app.router.routes.append(route)

    @app.get("/extractor", response_class=HTMLResponse, include_in_schema=False)
    def extractor() -> str:
        return (legacy.ROOT / "web" / "index.html").read_text(encoding="utf-8")

    if (DIST / "assets").is_dir():
        app.mount("/assets", StaticFiles(directory=DIST / "assets"), name="assets")

    @app.get("/{full_path:path}", include_in_schema=False)
    def spa(full_path: str):
        if full_path.startswith("api/"):
            return JSONResponse(status_code=404, content={"code": "NOT_FOUND", "message": "Unknown API route."})
        target = (DIST / full_path).resolve()
        if full_path and target.is_file() and DIST.resolve() in target.parents:
            return FileResponse(target)
        index = DIST / "index.html"
        if index.exists():
            # never cache the shell, so a rebuilt frontend is picked up on the next load
            return FileResponse(index, headers={"Cache-Control": "no-cache"})
        return HTMLResponse("<h1>AI Excel Workspace</h1><p>Frontend not built. Run <code>npm run build</code> in "
                            "<code>frontend/</code>, or use the dev server. Legacy extractor: "
                            "<a href='/extractor'>/extractor</a>.</p>")

    return app


app = create_app()


def main() -> None:
    import uvicorn
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    print(f"AI Excel Workspace: http://{host}:{port}  (LLM mode: {config.llm_mode()})")
    uvicorn.run(app, host=host, port=port)
