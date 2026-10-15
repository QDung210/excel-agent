"""CLI:  python -m xlsx_template INPUT.xlsx [-o OUT.xlsx] [--model M] [--plan PLAN.json] [--dry-run]"""
from __future__ import annotations

import argparse
import sys

from dotenv import load_dotenv

from .decider import DEFAULT_MODEL, DEFAULT_REASONING
from .graph import run


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="xlsx_template", description="Turn a filled-in .xlsx into a blank template.")
    ap.add_argument("input", help="input .xlsx")
    ap.add_argument("-o", "--output", help="output .xlsx (default: <input>_template.xlsx)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Groq model id (default {DEFAULT_MODEL})")
    ap.add_argument("--reasoning", choices=["default", "none"], default=DEFAULT_REASONING,
                    help=f"'none' (default) = no model reasoning: faster, fewer tokens; 'default' = let it think")
    ap.add_argument("--no-cache", action="store_true", help="ignore and do not update the decision cache")
    ap.add_argument("--plan", help="apply a saved/edited <output>.plan.json instead of calling the LLM")
    ap.add_argument("--dry-run", action="store_true", help="only parse + decide, write the plan JSON")
    args = ap.parse_args(argv)

    load_dotenv()
    state = run(args.input, args.output, model=args.model, plan_in=args.plan, dry_run=args.dry_run,
                reasoning=args.reasoning, use_cache=not args.no_cache)

    if args.dry_run:
        print(f"plan: {state['plan_path']}")
        return 0
    v = state["validation"]
    ws = state["write_stats"]
    print(f"template: {state['output_path']}")
    print(f"report:   {state['report_path']}")
    if state.get("plan_path"):
        print(f"plan:     {state['plan_path']}")
    print(f"cleared {ws['cleared']} cells | validation {'PASS' if v['ok'] else 'FAIL'} {v['checks']}")
    for e in v["errors"][:10]:
        print("  ERROR", e)
    return 0 if v["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
