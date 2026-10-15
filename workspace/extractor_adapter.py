"""Calls the existing LangGraph template extractor (xlsx_template.graph.run) in-process, unchanged.

Used only when the user asks to derive a blank template from a populated workbook. It needs the Azure OpenAI model
(the decider is LLM-based; the legacy decider is pointed at Azure here, its files stay unchanged), so in mock mode it is refused with MODEL_UNAVAILABLE instead of faking a result. The input
upload is never modified; the output becomes a new user template in the gallery.
"""
from __future__ import annotations

import tempfile
import threading
from contextlib import contextmanager
from pathlib import Path

from . import config, runs, templates, upload
from .errors import InvalidOperation, ModelUnavailable


_patch_lock = threading.Lock()


@contextmanager
def _azure_decider():
    """Runs the legacy graph with its LLM decider calling Azure OpenAI instead of Groq."""
    from xlsx_template import decider as dec

    class AzureDecider(dec.LLMDecider):
        def __init__(self, model=None, reasoning=dec.DEFAULT_REASONING, **kw):
            super().__init__(model=config.azure_deployment(), api_key="azure", fallback_model=None,
                             reasoning=reasoning, **kw)

        def _llm(self, model: str, max_tokens: int):
            return config.azure_chat(max_tokens, 120, model_kwargs={"response_format": {"type": "json_object"}})

    with _patch_lock:
        orig, dec.LLMDecider = dec.LLMDecider, AzureDecider
        try:
            yield
        finally:
            dec.LLMDecider = orig


def available() -> bool:
    return config.llm_mode() == "azure"


def start(file_id: str | None = None, template_id: str | None = None, name: str | None = None) -> runs.Run:
    if not available():
        raise ModelUnavailable("Making a blank template needs the Azure OpenAI model (AZURE_OPENAI_*); it is disabled in "
                               "mock mode.")
    if bool(file_id) == bool(template_id):
        raise InvalidOperation("Give exactly one of file_id or template_id.")
    if file_id:
        f = upload.get(file_id)
        if f.kind != "xlsx":
            raise InvalidOperation("Only .xlsx files can become templates.")
        src, base = f.path, Path(f.filename).stem
    else:
        src, base = templates.path(template_id), templates.get(template_id)["name"]

    def work(run: runs.Run) -> dict:
        from xlsx_template import graph
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "blank.xlsx"
            run.emit("tool_started", {"tool": "run_template_extractor"})

            def progress(pct, msg, cap=None):
                run.emit("progress", {"percent": int(pct), "message": str(msg)[:200]})

            with _azure_decider():
                state = graph.run(str(src), str(out), progress=progress)
            v, ws = state["validation"], state["write_stats"]
            if not v["ok"]:
                run.emit("tool_completed", {"tool": "run_template_extractor", "status": "error",
                                            "summary": "validation failed"})
                run.emit("validation_failed", {"errors": list(v["errors"])[:10]})
                return {"errors": list(v["errors"])[:10]}
            meta = templates.add(out.read_bytes(), f"{base} (blank)"[:120],
                                 f"Made blank by the template extractor: {ws.get('cleared', 0)} cell(s) cleared, "
                                 f"{v['checks'].get('formulas_preserved', 0)} formula(s) preserved.")
            run.emit("tool_completed", {"tool": "run_template_extractor", "status": "ok",
                                        "summary": f"{ws.get('cleared', 0)} cells cleared"})
        return {"template_id": meta["id"], "cleared": ws.get("cleared", 0),
                "formulas_preserved": v["checks"].get("formulas_preserved")}

    return runs.start("extract", work)
