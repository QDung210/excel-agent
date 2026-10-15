"""Stage 3: LLM (Groq, OpenAI-compatible) labels each region keep / clear / review.

The LLM only returns a one-letter action per region id. Anything missing, malformed or
"review" is treated as KEEP by the writer: deleting structure is worse than leaving data.

Speed on the Groq free tier (8k tokens/min, 1k OUTPUT tokens/min, 200k tokens/day):
  - region table in TOON (official encoder) -> compact input
  - all sheets packed into as few calls as fit the per-minute budget (usually 1)
  - output is {"d": {"<id>": "k|c|r"}} -> ~5 tokens per region, no reasons
  - reasoning off by default (hidden reasoning tokens count as output)
  - decisions are cached per region structure (cache.py) -> only new regions go to the LLM
"""
from __future__ import annotations

import json
import os
import re
import time
from collections import defaultdict
from dataclasses import dataclass, field

import toon_format as toon
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from .regions import ParsedWorkbook, Region

GROQ_BASE_URL = "https://api.groq.com/openai/v1"
DEFAULT_MODEL = "qwen/qwen3.8-27b"
FALLBACK_MODEL = "openai/gpt-oss-120b"
DEFAULT_REASONING = "none"
# Input token budget per call (cl100k estimate; Groq's own count runs ~10-20% higher).
# Free tier allows 8k tokens/min including the expected output.
CALL_INPUT_BUDGET = 6000

SYSTEM_PROMPT = """You turn a filled-in Excel workbook into a reusable BLANK template.
You get the workbook's chunks (from a parser) and a table of regions. A region is a group of
CONSTANT cells (never formulas: formulas are always kept) in one column (or one row run) of a chunk.

The region table is in TOON format: `regions[N]{fields}:` declares the columns once, then each
indented line is one region, values comma-separated in the same order ("" = empty).
Columns: id; sheet; chunk; cells (A1 range, or Col[rows]); n = cell count; type; annotation
(parser: label/data); style (bold/filled); header_band (parser thinks it is in the column-header
rows); header_above / label_left (nearest label above / on the left); read_by_formulas and
cross_sheet (how many formulas read these cells); dropdown_source (feeds a dropdown list);
sequence (consecutive series like 1..13 or S1..S7); values (up to 5 distinct sample values).

For each region choose:
- k = keep: template structure: titles of generic sections, column headers, field labels,
  legends, notes/instructions to the user, fixed option lists used by dropdowns (status,
  priority, sprint codes...), generic configuration constants (hours per day, working days
  per week...), unit descriptions.
- c = clear: data specific to this one project/instance that the next user re-enters:
  project names/titles, task names, descriptions, IDs/codes of tasks, people's names,
  emails, roles of specific people, dates, estimates, progress, statuses of tasks,
  notes about specific tasks, per-person numbers, out-of-scope items, leave records.
- r = review: genuinely unsure.

Hints:
- annotation=label from the parser and bold/filled header rows usually mean keep.
- header_band=true: usually keep, but the band can wrongly include a data row
  (e.g. a person's name) -> judge by values.
- A region under a header (header_above) holding many distinct values is usually data -> clear.
- dropdown_source=true: the region feeds a dropdown. If it is a generic option list -> keep;
  if it lists this project's people/items -> clear.
- Regions in hidden helper sheets are usually lookup lists -> keep, unless clearly project data.
- sequence=1..N or S1..SN: a consecutive series. Timeline axis labels (week/sprint/period numbers),
  fixed codes of a lookup table and row numbering are structure -> keep. Hierarchical IDs of
  project items (1.1.1, 1.2.3) are data -> clear.
- A small lookup table read by many formulas (VLOOKUP/INDEX/LOOKUP) whose values are generic
  codes or settings -> keep. Attributes of specific people (seniority, allocation, tags) -> clear.
- A title naming this specific project is project data -> clear.
Be careful: wrongly clearing structure is worse than leaving some data.

Reply ONLY with JSON, one entry for EVERY region id in the table:
{"d":{"<id>":"k|c|r", ...}}"""

ACTIONS = {"k": "keep", "c": "clear", "r": "review", "keep": "keep", "clear": "clear", "review": "review"}

try:
    import tiktoken
    _ENC = tiktoken.get_encoding("cl100k_base")

    def count_tokens(s: str) -> int:
        return len(_ENC.encode(s))
except Exception:  # tiktoken unavailable or its encoding cannot be loaded offline
    def count_tokens(s: str) -> int:
        return len(s) // 3


@dataclass
class Decision:
    id: int
    action: str            # keep | clear | review
    reason: str = ""
    by: str = "llm"        # llm:<model> | cache | fallback


@dataclass
class DecideResult:
    decisions: dict[int, Decision]
    model: str
    calls: int = 0
    usage: dict[str, int] = field(default_factory=lambda: defaultdict(int))
    errors: list[str] = field(default_factory=list)
    seconds: float = 0.0


def region_record(r: Region) -> dict:
    """One row of the region table. Field names are explained in SYSTEM_PROMPT."""
    return {
        "id": r.id, "sheet": r.sheet, "chunk": r.chunk, "cells": r.range_text, "n": len(r.cells),
        "type": r.vtype, "annotation": r.annotation,
        "style": "+".join(s for s, on in (("bold", r.bold), ("filled", bool(r.fill))) if on),
        "header_band": r.in_header, "header_above": r.header_above, "label_left": r.label_left,
        "read_by_formulas": r.refs, "cross_sheet": r.refs_cross_sheet,
        "dropdown_source": r.dropdown_source, "sequence": r.sequence,
        "values": " | ".join(r.samples),
    }


def regions_toon(regions: list[Region]) -> str:
    """Region table in TOON (official toon-format encoder): field names once, then one row
    per region. ~55% fewer tokens than JSON on this kind of uniform table."""
    return toon.encode({"regions": [region_record(r) for r in regions]})


def chunk_block(pw: ParsedWorkbook, idx: int) -> str:
    ch = pw.chunks[idx]
    hidden = " HIDDEN sheet" if ch.sheet in pw.hidden_sheets else ""
    return f"[chunk {ch.idx}] {ch.sheet}!{ch.a1} ({ch.block_type}{hidden})\n{ch.preview}"


def batch_prompt(pw: ParsedWorkbook, regions: list[Region]) -> str:
    hidden = f" (hidden: {', '.join(sorted(pw.hidden_sheets))})" if pw.hidden_sheets else ""
    out = [f"Workbook sheets: {', '.join(pw.sheets)}{hidden}", "", "Chunks (parser preview):"]
    out += [chunk_block(pw, i) for i in sorted({r.chunk for r in regions if r.chunk >= 0})]
    out += ["", "Regions (TOON table):", regions_toon(regions)]
    return "\n".join(out)


def pack_batches(pw: ParsedWorkbook, regions: list[Region], budget: int = CALL_INPUT_BUDGET) -> list[list[Region]]:
    """Fewest calls that each fit the per-minute token budget. Regions stay in sheet order;
    a sheet is only split when it alone does not fit."""
    fixed = count_tokens(SYSTEM_PROMPT)

    def cost(batch: list[Region]) -> int:
        return fixed + count_tokens(batch_prompt(pw, batch))

    if cost(regions) <= budget:
        return [regions]
    batches: list[list[Region]] = []
    cur: list[Region] = []
    for r in regions:        # greedy with exact token counts (a few hundred regions at most)
        if cur and cost(cur + [r]) > budget:
            batches.append(cur)
            cur = []
        cur.append(r)
    if cur:
        batches.append(cur)
    return batches


class LLMDecider:
    def __init__(self, model: str = DEFAULT_MODEL, api_key: str | None = None,
                 base_url: str = GROQ_BASE_URL, max_retries: int = 3,
                 fallback_model: str | None = FALLBACK_MODEL, reasoning: str = DEFAULT_REASONING,
                 input_budget: int = CALL_INPUT_BUDGET):
        self.model = model
        self.reasoning = reasoning   # "none" = fastest; "default" lets the model think (more tokens)
        self.fallback_model = fallback_model if fallback_model != model else None
        self.api_key = api_key or os.environ.get("GROQ_API_KEY")
        if not self.api_key:
            raise RuntimeError("GROQ_API_KEY is not set (put it in .env)")
        self.base_url = base_url
        self.max_retries = max_retries
        self.input_budget = input_budget

    def _llm(self, model: str, max_tokens: int) -> ChatOpenAI:
        kwargs = {"response_format": {"type": "json_object"}}
        extra = {}
        if model.startswith("qwen/"):
            if self.reasoning == "none":
                extra["reasoning_effort"] = "none"
            else:
                extra["reasoning_format"] = "hidden"
        elif model.startswith("openai/gpt-oss"):
            extra["reasoning_effort"] = "low" if self.reasoning == "none" else "medium"
        # Explicit max_tokens: without it Groq counts the model's full output window
        # against the per-minute quota and rejects the request as "too large".
        return ChatOpenAI(model=model, api_key=self.api_key, base_url=self.base_url,
                          temperature=0, max_retries=2, timeout=120, max_tokens=max_tokens,
                          model_kwargs=kwargs, extra_body=extra or None)

    def _call(self, model: str, prompt: str, n_regions: int, result: DecideResult) -> dict:
        # ~6 output tokens per region ("123":"c",). gpt-oss always reasons a little.
        budget = 8 * n_regions + 200
        if self.reasoning != "none":
            budget += 2500
        elif model.startswith("openai/gpt-oss"):
            budget += 1200
        msg = self._llm(model, min(budget, 16000)).invoke([SystemMessage(SYSTEM_PROMPT), HumanMessage(prompt)])
        result.calls += 1
        meta = getattr(msg, "usage_metadata", None) or {}
        for k in ("input_tokens", "output_tokens"):
            result.usage[k] += meta.get(k, 0) or 0
        text = msg.content if isinstance(msg.content, str) else str(msg.content)
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
        m = re.search(r"\{.*\}", text, re.S)
        return json.loads(m.group(0) if m else text)

    def decide(self, pw: ParsedWorkbook, regions: list[Region] | None = None, on_progress=None) -> DecideResult:
        """Label `regions` (default: all). on_progress(done, total, label) before/after each call."""
        t0 = time.perf_counter()
        result = DecideResult(decisions={}, model=self.model)
        regions = pw.regions if regions is None else regions
        batches = pack_batches(pw, regions, self.input_budget) if regions else []

        for n, batch in enumerate(batches):
            if on_progress:
                sheets = list(dict.fromkeys(r.sheet for r in batch))
                on_progress(n, len(batches), ", ".join(sheets))
            ids = {r.id for r in batch}
            got = self._decide_batch(pw, batch, ids, result)
            for rid in ids:
                result.decisions[rid] = got.get(rid) or Decision(rid, "review", "no answer from LLM", "fallback")
        if on_progress:
            on_progress(len(batches), len(batches), "")
        result.seconds = time.perf_counter() - t0
        return result

    def _decide_batch(self, pw, batch, ids, result) -> dict[int, Decision]:
        prompt = batch_prompt(pw, batch)
        label = f"batch of {len(batch)} regions"
        models = [self.model] + ([self.fallback_model] if self.fallback_model else [])
        best: dict[int, Decision] = {}
        for model in models:
            for attempt in range(self.max_retries):
                try:
                    data = self._call(model, prompt, len(batch), result)
                    out: dict[int, Decision] = {}
                    answers = data.get("d", data)
                    if isinstance(answers, list):   # tolerate [[id, "c"], ...]
                        answers = {str(a[0]): a[1] for a in answers if isinstance(a, (list, tuple)) and len(a) >= 2}
                    for k, v in (answers.items() if isinstance(answers, dict) else []):
                        try:
                            rid = int(k)
                        except (TypeError, ValueError):
                            continue
                        action = ACTIONS.get(str(v).lower().strip())
                        if rid in ids and action:
                            out[rid] = Decision(rid, action, "", f"llm:{model}")
                    if len(out) > len(best):
                        best = out
                    if len(out) >= 0.8 * len(ids):
                        return out
                    result.errors.append(f"{label}: {model} answered {len(out)}/{len(ids)} ids (attempt {attempt + 1})")
                except Exception as e:  # network, rate limit, bad JSON
                    msg = str(e)
                    result.errors.append(f"{label}: {model} attempt {attempt + 1}: {type(e).__name__}: {msg[:400]}")
                    if "per day" in msg:          # daily quota: retrying this model is pointless
                        break
                    m = re.search(r"try again in (?:(\d+)m)?([\d.]+)s", msg)
                    if m:
                        wait = 60 * int(m.group(1) or 0) + float(m.group(2)) + 1
                    elif "per minute" in msg:     # Groq free tier: per-minute token window
                        wait = 30 * (attempt + 1)
                    else:
                        wait = 5 * (attempt + 1)
                    time.sleep(min(wait, 65))
        return best
