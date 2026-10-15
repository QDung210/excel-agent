"""Human-readable markdown report of what was kept / cleared and why."""
from __future__ import annotations

import json
import os


def _esc(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def write_report(state: dict, path: str) -> None:
    v = state.get("validation") or {}
    ws = state.get("write_stats") or {}
    llm = state.get("llm_info") or {}
    lines = [
        f"# Template report: {os.path.basename(state['input_path'])}",
        "",
        f"- Output: `{os.path.basename(state['output_path'])}`",
        f"- Validation: **{'PASS' if v.get('ok') else 'FAIL'}**  {json.dumps(v.get('checks', {}))}",
        f"- Cells cleared: {ws.get('cleared')}, formula caches stripped: {ws.get('formula_cache_stripped')}, "
        f"hyperlinks removed: {ws.get('hyperlinks_removed')}, shared strings "
        f"{ws.get('shared_strings_before')} -> {ws.get('shared_strings_after')}",
        f"- Decider: {json.dumps(llm, ensure_ascii=False)}",
        f"- Timings (s): {json.dumps(state.get('timings', {}))}",
    ]
    for e in v.get("errors", []):
        lines.append(f"- ERROR: {e}")
    for w in ws.get("warnings", []):
        lines.append(f"- warning: {w}")

    plan_path = state.get("plan_path") or state.get("plan_in")
    if plan_path and os.path.exists(plan_path):
        with open(plan_path, encoding="utf-8") as f:
            regions = json.load(f)["regions"]
        sheet = None
        for r in regions:
            if r["sheet"] != sheet:
                sheet = r["sheet"]
                lines += ["", f"## {sheet}", "", "| id | action | cells | header | values | reason |",
                          "|---|---|---|---|---|---|"]

            cells = ", ".join(r["cells"][:4]) + (" …" if len(r["cells"]) > 4 else "")
            mark = {"clear": "🧹 clear", "keep": "keep", "review": "⚠ review (kept)"}.get(r["action"], r["action"])
            lines.append(f"| {r['id']} | {mark} | {_esc(cells)} | {_esc(r.get('header_above', ''))} | "
                         f"{_esc(' / '.join(r.get('samples', [])[:3]))} | {_esc(r.get('reason', ''))} |")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
