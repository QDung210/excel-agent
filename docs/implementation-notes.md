# Implementation notes — AI Excel Workspace

Audit date: 2026-10-13. These are facts read from the code and checked by running it, not taken from the prompt pack. Where the prompt pack and the code differ, the code wins and the difference is written below.

## 1. Repo audit (answers to `12_REPO_AUDIT_CHECKLIST.md`)

### Q1. How is the current app launched? Which route serves the old screen?
- CLI: `python -m xlsx_template <input.xlsx> [-o out] [--model] [--reasoning] [--no-cache] [--plan plan.json] [--dry-run]`.
- Web: `python -m xlsx_template.server` starts the FastAPI app `xlsx_template.server:app` with uvicorn on `HOST`/`PORT`, default 127.0.0.1:8000.
  - `GET /` serves `xlsx_template/web/index.html`, a single static page with no build step.
  - API: `POST /api/jobs`, `GET /api/jobs/{id}`, `DELETE /api/jobs/{id}`, `GET /api/jobs/{id}/download`, `GET /api/jobs/{id}/report`.
  - Jobs live in memory, with files under `jobs/<id>/`. One worker thread runs at a time because of the Groq per-minute quota.
- The old page calls absolute `/api/jobs...` paths. The new combined server therefore keeps those exact routes, by mounting the legacy routes, and serves the old page at **`/extractor`**.
- `python -m xlsx_template.server` is not modified and still serves the old page at `/`.

### Q2. Which tests exist? Baseline?
- The repo had **no tests** for `xlsx_template` before this work. `excel-parser-main/` has its own upstream test suite; it is not ours and was not run here.
- The baseline was recorded with the CLI (see `docs/phase-status.md`):
  - Applying the saved plan `output/Report_WBS_TEMPLATE_final.plan.json` to `Report_WBS (5) (1) 1.xlsx` gives validation PASS, 1722 cells cleared, 1511 formulas preserved, 22 parts identical and 8 changed.
  - A dry run with the real Groq key wrote the plan: 126 regions in 2 calls, 161 s.
- Added `tests/test_legacy_extractor.py`, 12 tests covering plan apply, x14 preservation, untouched parts, dry run, the fail-safe KEEP, cache, writer and validator. They use a fake decider and never call Groq or touch `cache/decisions.json`.

### Q3. How do we call the Template Extractor as a Python function?
- `xlsx_template.graph.run(input_path, output_path=None, model=None, plan_in=None, dry_run=False, reasoning="none", progress=None, use_cache=True) -> State`.
- `progress(pct, msg, cap)` is optional; raising inside it cancels the run cooperatively.
- `run` refuses `output_path == input_path`.
- The workspace uses it through `workspace/extractor_adapter.py`, in-process and without a shell.

### Q4. What can the OOXML writer do?
`xlsx_template.writer.write_template(src, dst, clear: {sheet: {(row, col)}}) -> WriteStats` does **CLEAR only**:
- removes the value of constant cells and keeps the `s` style attribute;
- strips cached `<v>` from **all** formula cells;
- drops shared strings that are no longer used and renumbers them;
- removes `<hyperlink>` entries of cleared single cells, plus their relationship;
- sets `<calcPr fullCalcOnLoad="1">`;
- copies every other zip member byte for byte, with the same order, ZipInfo and compression type.

It **does not** do set_value, set_formula, styles, or row/column insert/delete. `set_value`/`clear_value` for the workspace are implemented separately in `workspace/ooxml_edit.py` and leave `writer.py` unchanged.

### Q5. Which source types does the parser handle? Which structures stay untouched?
- `excel_parser.parse_workbook` (ks-xlsx-parser 0.2.1, MIT, installed editable from `excel-parser-main/`) reads `.xlsx`/`.xlsm` through openpyxl, which is read-only for us. It produces chunks, header spans and cells with annotation, style and formula information.
- `regions.build()` also reads the raw XML for dropdown sources and hidden sheets.
- The CLEAR writer leaves untouched:
  - every part except worksheets, worksheet rels, `sharedStrings.xml` and `workbook.xml`;
  - inside a worksheet, everything outside `<sheetData>`/`<hyperlinks>`: merges, column widths, sheet views, conditional formats, x14 `extLst` validations, and so on.

### Q6. How much is preserved byte for byte?
- **Per-member content** of untouched parts is byte-identical (`writestr(info, raw)`).
- The **archive container** is rewritten: zip headers, timestamps from the original ZipInfo and recompressed data, so the `.xlsx` file hash changes.
- The validator compares member contents, not archive bytes. We never claim the archive bytes are identical.

### Q7. Which Groq modes really work? What happens without a key?
These were verified on 2026-10-13 with the user's key (probe script, 1 call per mode):

| model | JSON mode (`json_object`) | tool calling (`bind_tools`) | `json_schema` structured output |
|---|---|---|---|
| `qwen/qwen3.8-27b` | works (used by decider) | **works** (0.5 s) | **works** |
| `openai/gpt-oss-120b` | works (fallback) | **works** (1.9 s) | **works** |

Additional findings:
- **Limit:** in this org, qwen's real per-request/minute token limit is about 7000 tokens. The decider's tiktoken estimate (budget 6000) is under Groq's own count: one batch counted as 7082 and got HTTP **413 "Request too large"**. The decider retries the same model 3× on a 413, which can never succeed, before it falls back. This is **not changed**, so legacy behavior is preserved; it is documented as a known issue (see §4).
- **Without a key:** `LLMDecider()` raises `RuntimeError("GROQ_API_KEY is not set")`. The legacy job shows the error.
- **Workspace without a key:** set `WORKSPACE_LLM=mock` and the workspace uses deterministic rule-based clarify, spec and agent logic. The UI shows "AI: mock mode".

### Q8. Existing web stack. Can the Vite frontend coexist?
- The legacy UI is one static HTML file served by FastAPI, with no npm.
- The new app lives in `frontend/` (Vite). `npm run build` outputs `frontend/dist`, which the combined server `python -m workspace` serves at `/`. The legacy page is at `/extractor` and the legacy API stays at `/api/jobs*`. Dev mode uses the Vite proxy `/api → :8000`.

### Q9. Git status / uncommitted user work
- The project folder is **not its own git repository**. `git rev-parse --show-toplevel` gives `C:/Users/dung`, a home-directory repo where this folder is untracked.
- Nothing was committed and no user files were moved or deleted.
- User artifacts that are kept untouched:
  - `output/*` (including a LibreOffice lock file)
  - `jobs/*`
  - `cache/decisions.json`
  - `.env`
  - the old `.venv`, which is broken because its interpreter path points to `C:\Users\DungDQ23\...`
  - `xlcompress_llm_template_test.ipynb`
  - `.deepeval/`

### Q10. Python/Node versions, dependencies
- The old `.venv` is broken, and system Python 3.13 lacks `excel_parser` and `toon_format`. A **new** environment `.venv-ws` was created with uv (CPython 3.12.12). The old `.venv` was left in place.
- Installed versions:

| package | version |
|---|---|
| langgraph | 1.2.15 |
| langchain-core | 1.6.10 |
| langchain-openai | 1.7.1 |
| deepagents | 0.7.24 |
| fastapi | 0.143.0 |
| starlette | 1.8.0 |
| uvicorn | 0.54.0 |
| pydantic | 2.14.1 |
| openpyxl | 3.1.5 |
| lxml | 6.1.3 |
| xlsxwriter | 3.2.9 |
| toon-format | 1.1.0 |
| tiktoken | 0.14.0 |
| pytest | 9.1.1 |

- No conflicts were found.
- `requirements.txt` was missing `fastapi`, `uvicorn` and `tiktoken`, although the code imports them. They are now added, together with the new dependencies, in `requirements.txt`.
- Node v24.13.0, npm 11.6.2.

## 2. Chosen architecture (what was built)

```
workspace/                 new package; xlsx_template/ is NOT modified
  config.py                env + paths (WORKSPACE_DATA_DIR, limits, LLM mode)
  storage.py               opaque ids, atomic writes, data/ layout
  upload.py                extension/signature/zip-safety checks
  schemas.py               Pydantic contracts (08_SCHEMAS_AND_EXAMPLES)
  sources.py               txt/csv -> SourceFacts with line/row provenance
  builder.py               WorkbookSpec -> XlsxWriter
  ooxml_edit.py            set_value / clear_value OOXML patcher (new writer track)
  validate_edit.py         L0-L4 validation of an edit revision
  preview.py               XLSX -> FortuneSheet DTO + compatibility warnings
  service.py               WorkbookService: the only write boundary (revisions, 409, undo)
  extractor_adapter.py     calls xlsx_template.graph.run in-process (Groq only; 503 in mock)
  clarify.py               analyze -> questions (<=5, all at once) -> ready/blocked
  spec_gen.py              WorkbookSpec: mock rules or Groq layout-only + grounding guard
  generate.py              generate run (scratch: spec_gen+builder; template: template_fill)
  formula_eval.py          cached values for allowlisted formulas (preview only)
  templates.py             gallery (2 built-ins + uploads) + slot detection (openpyxl read-only)
  template_fill.py         slots <- SourceFacts (label/header match, LLM pairs ids only) -> set_value ops
  workspaces.py            workspace state (mode, files, clarification)
  agent/store.py           threads / plans / tool logs under data/agent/ (logs hold addresses, never values)
  agent/tools.py           8 workbook tools, grounding guard, plans (<=25 cells auto, else confirm), apply/reject
  agent/runtime.py         Deep Agents turn (FS/shell/todo/task excluded), Groq primary->fallback, claim guard
  agent/mock_model.py      deterministic regex command model for WORKSPACE_LLM=mock (labelled as mock)
  runs.py                  in-process run registry + SSE event log (seq ids)
  api.py / app.py          FastAPI routers; app mounts legacy routes + SPA
frontend/                  Vite + React + TS
```

Data layout (`WORKSPACE_DATA_DIR`, default `./data`, git-ignored):
```
data/uploads/{file_id}/original.<ext>  + meta.json      (immutable originals)
data/workspaces/{ws_id}.json                              (mode, files, prompt, clarify state)
data/workbooks/{wb_id}/revisions/{rev}.xlsx               (canonical revisions, never overwritten)
data/workbooks/{wb_id}/metadata.json                      (revisions, hashes, validation)
data/workbooks/{wb_id}/history.jsonl                      (operation log, values not logged)
data/templates/{tpl_id}/template.xlsx + meta.json         (gallery)
data/agent/{threads,plans,logs}/                          (chat threads, edit plans, tool logs)
```

## 3. Decisions taken without asking (small missing inputs)

| Topic | Default chosen | Why |
|---|---|---|
| Python env | new `.venv-ws` (uv, 3.12) | the old `.venv` is broken; deleting it is not ours to do |
| Data dir | `./data` | filesystem only, as the brief says; no DB/queue |
| Server | `python -m workspace` on :8000 with legacy routes inside | one origin, so the legacy page keeps working |
| Formula literals | `set_value` with a string starting with `=` is stored as **text** | prevents formula injection |
| Edit op allowlist | `set_value`, `clear_value` only | `set_formula`/`set_style` are disabled until a writer and tests exist |
| Gallery | templates the user uploads with "Save to gallery", plus 2 built-in templates generated by XlsxWriter at first start | no fake third-party templates |
| Agent context | Deep Agents with built-in filesystem tools and the `task` subagent **excluded** (`HarnessProfile`) | keeps the 7k-token Groq limit and the no-host-FS rule |
| LLM mode in tests | `WORKSPACE_LLM=mock` | CI never calls the network |
| Undo | creates a **new** revision whose bytes copy the target revision | history is never rewritten |
| Clarify limit | up to 5 questions, required ones first | per PRD |
| Language | English UI labels, Vietnamese-friendly text (UTF-8 everywhere) | per spec 04 |

## 4. Known issues found in the existing code (documented, not changed)
1. Decider 413 handling: on `413 Request too large` the decider retries the same model (max_retries=3) before falling back, which wastes about 2 minutes. Suggested fix: break on 413 and/or lower `CALL_INPUT_BUDGET` to about 5000. It was not applied, because changing legacy behavior needs the owner's OK.
2. `requirements.txt` did not list `fastapi`, `uvicorn` or `tiktoken`, which the code imports. This is fixed by adding them, with no behavior change.
3. `server.py` accepts `.xlsm` uploads, but the writer strips nothing VBA-specific. It works, but macro files are passed through as is. The workspace upload refuses `.xlsm`, per security spec 11.
4. `xlsx_template/writer.py:46` strips the formula cache with `<v>.*?</v>`, which misses `<v xml:space="preserve">`. That leaves text inside a numeric cell, and Excel or openpyxl may reject it. The same bug in the workspace writer was fixed in Phase 4. The legacy file is left unchanged, as the rules require.
