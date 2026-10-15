"""Generate run (scratch mode, and template mode via template_fill): emits tool events, creates revision 1."""
from __future__ import annotations

import re
import tempfile
from pathlib import Path

from . import builder, clarify, formula_eval, runs, service, sources, spec_gen, workspaces
from .errors import InvalidOperation


def _filename(state: dict, title: str) -> str:
    name = (state.get("output_filename") or title or "workbook").strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", name).strip(" .")[:100] or "workbook"
    return name if name.lower().endswith(".xlsx") else name + ".xlsx"


def validate_generated(path: Path, spec_stats: dict) -> tuple[list[str], list[str]]:
    """Re-open the produced file: structure, formula functions, cached error values."""
    from openpyxl import load_workbook
    errors, warnings = [], []
    wb = load_workbook(path, read_only=False)
    try:
        nform, unsupported = 0, set()
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        nform += 1
                        unsupported |= formula_eval.functions_used(c.value) - formula_eval.SUPPORTED_FUNCTIONS
        if nform != spec_stats["formulas"]:
            errors.append(f"formula count mismatch: wrote {spec_stats['formulas']}, found {nform}")
        if unsupported:
            warnings.append(f"Formulas use functions the preview cannot compute: {sorted(unsupported)}")
    finally:
        wb.close()
    return errors, warnings



def _step(run: runs.Run, tool: str, fn, summary):
    run.emit("tool_started", {"tool": tool})
    try:
        out = fn()
    except Exception:
        run.emit("tool_completed", {"tool": tool, "status": "error", "summary": "failed"})
        raise
    run.emit("tool_completed", {"tool": tool, "status": "ok", "summary": summary(out)})
    return out


def start(ws_id: str) -> runs.Run:
    ws = workspaces.get(ws_id)
    state = ws.get("clarification")
    if not state:
        raise InvalidOperation("Run analyze first.")
    if state["status"] != "ready":
        raise InvalidOperation(f"Analysis is '{state['status']}'; answer the questions first."
                               if state["status"] == "needs_clarification" else state.get("reason") or "Blocked.")
    if state.get("structure") and not state["structure"].get("approved"):
        raise InvalidOperation("The proposed structure has not been approved yet (or its preview is invalid); review "
                               "and approve it first.")

    def work(run: runs.Run) -> dict:
        run.emit("analysis_ready", {"clarification_id": state["id"], "plan": state["plan"]})
        if ws["mode"] == "template":
            from . import restructure, template_fill
            if state.get("structure"):
                return restructure.run(run, ws, state)
            return template_fill.run(run, ws, state)
        prompt, answers = state["prompt"], state.get("answers") or {}
        b = _step(run, "read_sources", lambda: sources.load(state["source_file_ids"], prompt),
                  lambda b: f"{len(b.facts)} facts, {len(b.tables)} tables")
        spec, warns, mode = _step(run, "build_spec", lambda: spec_gen.build_spec(b, prompt, answers),
                                  lambda o: f"{len(o[0].sheets)} sheet(s), layout={o[2]}")
        vi = clarify.is_vi(prompt)
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "generated.xlsx"
            stats = _step(run, "build_workbook", lambda: builder.build(spec, out, vi),
                          lambda s: f"{s['cells']} values, {s['formulas']} formulas")
            errors, vwarns = _step(run, "validate_workbook", lambda: validate_generated(out, stats),
                                   lambda o: "pass" if not o[0] else f"{len(o[0])} error(s)")
            if errors:
                run.emit("validation_failed", {"errors": errors})    # terminal event; no revision is created
                return {"errors": errors}
            if stats["uncomputed"]:
                vwarns.append(f"{stats['uncomputed']} formula value(s) are computed by Excel on open.")
            warnings = list(dict.fromkeys(state.get("warnings", []) + b.warnings + warns + vwarns))
            meta = service.create_from_path(out, _filename(state, spec.title), "generated", ws_id,
                                            extra={"warnings": warnings, "llm_mode": mode,
                                                   "spec": spec.model_dump(mode="json")})
        cur = workspaces.get(ws_id)
        cur["workbook_id"] = meta["id"]
        cur["runs"] = (cur.get("runs") or [])[-20:] + [run.id]
        workspaces.save(cur)
        run.emit("workbook_revision", {"workbook_id": meta["id"], "revision": 1, "changed_ranges": []})
        return {"workbook_id": meta["id"], "revision": 1, "warnings": warnings, "llm_mode": mode}

    return runs.start("generate", work)

