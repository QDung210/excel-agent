"""LangGraph pipeline: parse -> decide (LLM or saved plan) -> apply -> validate -> report.

New features (edit / create file, chat edits...) can be added as new nodes or branches
behind a router without touching these stages.
"""
from __future__ import annotations

import json
import os
import time
from collections import defaultdict
from typing import Any, TypedDict

from langchain_core.runnables import RunnableConfig
from langgraph.graph import END, START, StateGraph
from openpyxl.utils import get_column_letter, range_boundaries

from . import decider as dec
from .cache import DecisionCache
from .regions import ParsedWorkbook, build
from .validate import validate
from .writer import write_template


class State(TypedDict, total=False):
    input_path: str
    output_path: str
    plan_in: str | None          # reuse a saved / hand-edited plan instead of calling the LLM
    model: str
    reasoning: str
    use_cache: bool
    dry_run: bool
    parsed: ParsedWorkbook
    decisions: dict[int, dec.Decision]
    llm_info: dict[str, Any]
    clear: dict[str, set[tuple[int, int]]]
    write_stats: dict[str, Any]
    validation: dict[str, Any]
    plan_path: str
    report_path: str
    timings: dict[str, float]


def _progress(config: RunnableConfig | None, pct: float, message: str, cap: float | None = None) -> None:
    """Report progress to an optional callback passed as config['configurable']['progress'].

    cap = where progress will be when the current (long) step finishes, so a UI can creep
    toward it instead of looking frozen.
    """
    cb = ((config or {}).get("configurable") or {}).get("progress")
    if cb:
        cb(round(pct), message, round(cap if cap is not None else pct))


def _timed(name: str, state: State, t0: float) -> dict[str, float]:
    t = dict(state.get("timings") or {})
    t[name] = round(time.perf_counter() - t0, 2)
    return t


# ---------------------------------------------------------------- A1 helpers
def cells_to_ranges(cells: list[tuple[int, int]]) -> list[str]:
    by_col: dict[int, list[int]] = defaultdict(list)
    for r, c in cells:
        by_col[c].append(r)
    out = []
    for c, rows in sorted(by_col.items()):
        rows.sort()
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            col = get_column_letter(c)
            out.append(f"{col}{start}" if start == prev else f"{col}{start}:{col}{prev}")
            if r is not None:
                start = prev = r
    return out


def ranges_to_cells(ranges: list[str]) -> set[tuple[int, int]]:
    out = set()
    for a1 in ranges:
        c1, r1, c2, r2 = range_boundaries(a1)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out.add((r, c))
    return out


# ---------------------------------------------------------------- nodes
def parse_node(state: State, config: RunnableConfig) -> State:
    t0 = time.perf_counter()
    _progress(config, 2, "Đang phân tích cấu trúc file (excel-parser)…", 10)
    pw = build(state["input_path"])
    _progress(config, 10, f"Đã chia {len(pw.regions)} vùng trên {len(pw.sheets)} sheet")
    return {"parsed": pw, "timings": _timed("parse", state, t0)}


def decide_node(state: State, config: RunnableConfig) -> State:
    t0 = time.perf_counter()
    pw = state["parsed"]
    cache = DecisionCache() if state.get("use_cache", True) else None
    hits, pending = cache.lookup(pw) if cache else ({}, pw.regions)
    decisions = {i: dec.Decision(i, a, "", "cache") for i, a in hits.items()}
    info: dict[str, Any] = {"cache_hits": len(hits), "llm_regions": len(pending)}

    if pending:
        _progress(config, 12, f"{len(hits)} vùng lấy từ cache, AI xử lý {len(pending)} vùng…", 80)
        d = dec.LLMDecider(model=state.get("model") or dec.DEFAULT_MODEL,
                           reasoning=state.get("reasoning") or dec.DEFAULT_REASONING)

        def on_progress(done: int, total: int, sheets: str) -> None:
            pct = 12 + 73 * done / max(total, 1)
            nxt = 12 + 73 * min(done + 1, total) / max(total, 1)
            msg = (f"AI đang đọc {len(pending)} vùng ({sheets})" + (f" – lượt {done + 1}/{total}" if total > 1 else "")
                   if sheets else "AI đã quyết định xong")
            _progress(config, pct, msg, nxt)

        res = d.decide(pw, pending, on_progress=on_progress)
        decisions.update(res.decisions)
        if cache:   # only the primary model's answers: a fallback model's mistakes must not stick
            info["cached_new"] = cache.store(pw, {i: x.action for i, x in res.decisions.items()
                                                  if x.by == f"llm:{res.model}"})
        info.update({"model": res.model, "calls": res.calls, "usage": dict(res.usage),
                     "errors": res.errors, "seconds": round(res.seconds, 1)})
    else:
        _progress(config, 80, f"Cấu trúc đã gặp trước đó: dùng lại {len(hits)} quyết định từ cache", 85)
    return {"decisions": decisions, "llm_info": info, "timings": _timed("decide", state, t0)}


def load_plan_node(state: State) -> State:
    """Apply a saved / hand-edited plan: actions per region from the JSON, cells from its ranges.

    The plan is a human-approved answer, so its decisions are also written to the cache:
    the next workbook with the same structure reuses them without calling the LLM.
    """
    with open(state["plan_in"], encoding="utf-8") as f:
        plan = json.load(f)
    clear: dict[str, set[tuple[int, int]]] = defaultdict(set)
    for r in plan["regions"]:
        if r["action"] == "clear":
            clear[r["sheet"]] |= ranges_to_cells(r["cells"])
    info: dict[str, Any] = {"plan": state["plan_in"]}
    pw = state.get("parsed")
    if pw is not None and state.get("use_cache", True):
        by_cells = {(r["sheet"], tuple(r["cells"])): r["action"] for r in plan["regions"]}
        learned = {reg.id: by_cells[(reg.sheet, tuple(cells_to_ranges(reg.cells)))]
                   for reg in pw.regions if (reg.sheet, tuple(cells_to_ranges(reg.cells))) in by_cells}
        info["cache_learned"] = DecisionCache().store(pw, learned)
        info["plan_regions_matched"] = f"{len(learned)}/{len(pw.regions)}"
    return {"clear": dict(clear), "llm_info": info}


def plan_node(state: State) -> State:
    """Decisions -> cells to clear + plan JSON on disk (reviewable, re-appliable with --plan)."""
    pw, decisions = state["parsed"], state["decisions"]
    clear: dict[str, set[tuple[int, int]]] = defaultdict(set)
    regions = []
    for r in pw.regions:
        d = decisions.get(r.id) or dec.Decision(r.id, "review", "missing", "fallback")
        if d.action == "clear":           # keep and review both leave the cells untouched
            clear[r.sheet].update(r.cells)
        regions.append({
            "id": r.id, "sheet": r.sheet, "chunk": r.chunk, "action": d.action, "reason": d.reason,
            "by": d.by, "n_cells": len(r.cells), "cells": cells_to_ranges(r.cells),
            "header_above": r.header_above, "samples": r.samples,
        })
    plan_path = os.path.splitext(state["output_path"])[0] + ".plan.json"
    with open(plan_path, "w", encoding="utf-8") as f:
        json.dump({"source": os.path.basename(state["input_path"]), "llm": state.get("llm_info"),
                   "regions": regions}, f, ensure_ascii=False, indent=1)
    return {"clear": dict(clear), "plan_path": plan_path}


def apply_node(state: State, config: RunnableConfig) -> State:
    t0 = time.perf_counter()
    _progress(config, 86, "Đang xoá dữ liệu trên file gốc…")
    st = write_template(state["input_path"], state["output_path"], state["clear"])
    return {"write_stats": st.__dict__, "timings": _timed("apply", state, t0)}


def validate_node(state: State, config: RunnableConfig) -> State:
    t0 = time.perf_counter()
    _progress(config, 92, "Đang kiểm tra template (công thức, style, dropdown)…")
    rep = validate(state["input_path"], state["output_path"], state["clear"])
    return {"validation": rep.__dict__, "timings": _timed("validate", state, t0)}


def report_node(state: State, config: RunnableConfig) -> State:
    from .report import write_report
    path = os.path.splitext(state["output_path"])[0] + ".report.md"
    write_report(state, path)
    _progress(config, 100, "Hoàn tất")
    return {"report_path": path}


# ---------------------------------------------------------------- graph
def build_graph():
    g = StateGraph(State)
    g.add_node("parse", parse_node)
    g.add_node("decide", decide_node)
    g.add_node("load_plan", load_plan_node)
    g.add_node("plan", plan_node)
    g.add_node("apply", apply_node)
    g.add_node("validate", validate_node)
    g.add_node("report", report_node)

    # parse always runs first: with a plan it is only needed to teach the cache
    g.add_edge(START, "parse")
    g.add_conditional_edges("parse", lambda s: "load_plan" if s.get("plan_in") else "decide",
                            {"load_plan": "load_plan", "decide": "decide"})
    g.add_edge("decide", "plan")
    g.add_conditional_edges("plan", lambda s: "end" if s.get("dry_run") else "apply",
                            {"end": END, "apply": "apply"})
    g.add_edge("load_plan", "apply")
    g.add_edge("apply", "validate")
    g.add_edge("validate", "report")
    g.add_edge("report", END)
    return g.compile()


def run(input_path: str, output_path: str | None = None, model: str | None = None,
        plan_in: str | None = None, dry_run: bool = False, reasoning: str = dec.DEFAULT_REASONING,
        progress=None, use_cache: bool = True) -> State:
    """progress: optional callable(percent: int, message: str, cap: int)."""
    if output_path is None:
        root, _ = os.path.splitext(input_path)
        output_path = root + "_template.xlsx"
    if os.path.abspath(output_path) == os.path.abspath(input_path):
        raise ValueError("output_path must differ from input_path")
    return build_graph().invoke({
        "input_path": input_path, "output_path": output_path, "model": model, "reasoning": reasoning,
        "plan_in": plan_in, "dry_run": dry_run, "use_cache": use_cache, "timings": {},
    }, config={"configurable": {"progress": progress}})
