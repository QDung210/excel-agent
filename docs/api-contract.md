# API contract (backend `workspace/`, FastAPI)

Python sources: `workspace/schemas.py`. TypeScript mirror: `frontend/src/lib/types.ts`.
Errors always return `{code, message, ...extra}` with one of these codes:

| Status | Code |
|---|---|
| 400 | `UNSUPPORTED_FILE` |
| 404 | `NOT_FOUND` |
| 409 | `REVISION_CONFLICT` (also carries `current_revision`) |
| 413 | `FILE_TOO_LARGE` |
| 422 | `INVALID_OPERATION`, `VALIDATION_FAILED`, `INVALID_REQUEST` |
| 503 | `MODEL_UNAVAILABLE` |

IDs are opaque: `ws_…`, `file_…`, `wb_…`, `run_…`, `turn_…`, `plan_…`, `cl_…`, each followed by 16 hex characters.

The Status column below records whether a route is implemented and tested.

## Health
| Method | Path | Response | Status |
|---|---|---|---|
| GET | `/api/health` | `{ok, llm_mode: "mock"|"azure"\|"groq", model, max_upload_mb, extractor_available}` | done |

## Workspaces (creation flow)
| Method | Path | Body | Response | Status |
|---|---|---|---|---|
| POST | `/api/workspaces` | `{mode: "scratch"\|"template", name?}` | Workspace | done |
| GET | `/api/workspaces/{id}` | — | Workspace | done |
| POST | `/api/workspaces/{id}/files` | multipart `file`, form `role=source\|template` | `{id, filename, kind: txt\|csv\|xlsx, size, role}` | done |
| DELETE | `/api/workspaces/{id}/files/{file_id}` | — | Workspace (only detaches; the stored file is kept) | done |
| POST | `/api/workspaces/{id}/analyze` | `{prompt, source_file_ids?, template_file_id?, template_id?, output_filename?}` | AnalyzeResponse | Phase 2 |
| POST | `/api/workspaces/{id}/clarify` | `{clarification_id, answers: {question_id: value}}` | AnalyzeResponse | Phase 2 |
| POST | `/api/workspaces/{id}/generate` | `{}` (reuses the stored analysis and answers) | `{run_id}` | Phase 2 |

Fields of the main shapes:
- **Workspace:** `{id, mode, name, files: [{id, filename, kind, size, role}], template_id, prompt, clarification, workbook_id, runs}`
- **AnalyzeResponse:** `{status: "ready"|"needs_clarification"|"blocked", clarification_id, questions: ClarifyQuestion[], reason, plan: string[], facts: number, warnings: string[]}`
- **ClarifyQuestion:** `{id, type: select|boolean|text|number, question, options?, required, default?, explanation?}`. At most 5 questions are returned, and all are asked at once.

## Templates gallery (Phase 3)
| Method | Path | Body | Response | Status |
|---|---|---|---|---|
| GET | `/api/templates` | — | `[{id, name, description, origin: "builtin"\|"user", sheets: string[], created}]` (built-ins first; the 2 built-ins are created on first call) | done |
| POST | `/api/templates` | multipart `file` (.xlsx), form `name?`, `description?` | template entry | done |
| GET | `/api/templates/{id}/preview` | — | FortuneSheet DTO (same shape as the workbook preview) | done |
| GET | `/api/templates/{id}/slots` | — | `{kv: [{sheet, cell, label}], tables: [{sheet, header_row, columns, input_rows, prefilled_rows}], hidden_sheets}` | done |
| POST | `/api/templates/extract` | `{file_id}` or `{template_id}` (exactly one) | `{run_id}`; run result `{template_id, cleared, formulas_preserved}` | done; **503 `MODEL_UNAVAILABLE` in mock mode** |

Template mode uses the normal flow: `POST /api/workspaces {mode:"template"}`. Then either upload an xlsx with `role=template` or pass `template_id` to `/analyze`, followed by `/generate`. The run result is `{workbook_id, revision, warnings, llm_mode, filled}`. Revision 1 is the untouched template and revision 2 is the filled one. `metadata.fill_report[]` lists `{sheet, cell, label, value, fact_id, locator, via: label|header|llm|sequence}` for every filled cell.

## Runs (SSE)
| Method | Path | Notes |
|---|---|---|
| GET | `/api/runs/{id}` | `{id, kind, status: running\|done\|error\|cancelled, result, events, created}` |
| GET | `/api/runs/{id}/events` | `text/event-stream`. Every message has `id: <seq>`, `event: <name>` and `data: {event, run_id, seq, data}`. Reconnect with `Last-Event-ID`; `retry: 2000` |

Generate run events:
- `run_started`
- `analysis_ready`
- `tool_started {tool}` and `tool_completed {tool, status, summary}`
- `workbook_revision {workbook_id, revision, changed_ranges}`
- `validation_failed {errors}`
- `done {workbook_id, revision, warnings}`
- `error {code, message}`

No workbook JSON is ever sent over SSE. The client fetches `/preview` instead.

## Workbooks (canonical revisions)
| Method | Path | Body / params | Response | Status |
|---|---|---|---|---|
| POST | `/api/workbooks/upload` | multipart `file` (.xlsx) | WorkbookMeta | done |
| GET | `/api/workbooks` | — | `[{id, name, filename, updated, current_revision, origin}]` | done |
| GET | `/api/workbooks/{id}/metadata` | — | WorkbookMeta | done |
| GET | `/api/workbooks/{id}/preview` | `?revision=` | Preview | done |
| POST | `/api/workbooks/{id}/operations` | OperationBatch | RevisionResult | done |
| POST | `/api/workbooks/{id}/undo` | `{base_revision, target_revision?}` | RevisionResult | done |
| GET | `/api/workbooks/{id}/history` | — | `[{event, revision, …}]` | done |
| GET | `/api/workbooks/{id}/export` | `?revision=` | canonical `.xlsx` bytes (attachment, no-store) | done |

Fields of the workbook shapes:
- **WorkbookMeta:** `{id, name, filename, created, updated, current_revision, origin, workspace_id, sheets, llm_mode, warnings, revisions: [{revision, parent, created, origin, note, changed_count, restored}], can_undo}`
- **Preview:** `{workbook_id, revision, sheets: FortuneSheetSheet[], warnings: string[], stale_formulas: number}`. Each sheet is `{name, index, order, status, hide?, celldata: [{r, c, v}], config: {merge, borderInfo, columnlen, rowlen, rowhidden, colhidden}, frozen?, row, column}`, and `v` is FortuneSheet `Cell` (`v, m, f, ct, bl, it, un, cl, fs, ff, fc, bg, ht, vt, tb, mc`).
- **OperationBatch:** `{base_revision, operation_id (6–80 chars [A-Za-z0-9_-:.]), origin: manual|agent|system, note?, actions: Action[] (1–1000)}`
- **Action** is one of:
  - `{type: "set_value", sheet, cell: "B2", value: string|number|boolean|null, value_type?: auto|string|number|bool|date, source_fact_id?}`
  - `{type: "clear_value", sheet, cell}`
  - Text starting with `=` is stored as **literal text**; there is no `set_formula` yet.
  - Rejections (422): formula cells, non-anchor cells of a merge, unknown sheets, and any other `type`.
- **RevisionResult:** `{workbook_id, revision, parent_revision, status, changed_ranges, validation: {pass, level, warnings, errors, checks}, preview_url, idempotent_replay}`
  - Re-posting the same `operation_id` with the same content returns the stored result with `idempotent_replay: true`.
  - Re-posting it with different content returns 422.

## Agent chat (Phase 4)
| Method | Path | Body | Response |
|---|---|---|---|
| POST | `/api/agent/turns` | `{workbook_id, revision, message, thread_id?, selection?, active_sheet?}` | `{turn_id, run_id, thread_id}` |
| GET | `/api/agent/turns/{turn_id}/events` | SSE, same envelope as runs | — |
| POST | `/api/agent/plans/{plan_id}/apply` | `{base_revision}` | RevisionResult |
| POST | `/api/agent/plans/{plan_id}/reject` | `{}` | `{ok: true, status: "rejected"}` |

Errors:
- `/turns` returns 409 `{current_revision}` when `revision` is stale, and 404 for an unknown thread or a thread that belongs to another workbook.
- `/apply` returns 404 for an unknown plan, 422 if the plan is not pending, and 409 if the workbook has moved past the plan's `base_revision`.

Agent turn events (in order of appearance):
- `run_started`
- `tool_started {tool, args}`. `args` carries addresses and counts only, never cell values.
- `workbook_revision {workbook_id, revision, changed_ranges}`. Emitted before the matching `tool_completed`.
- `tool_completed {tool, status: ok|error, summary}`
- `confirmation_required {plan_id, base_revision, actions, diff: [{cell, before, after}], cells, reason}`
- `workbook_created {workbook_id}`. Only from `build_workbook`, which makes a separate new workbook.
- `model_fallback {from, error, detail}`. The Azure OpenAI model failed before anything was written.
- `text_delta {text}`. This includes the appended warning when a reply claims an edit that no tool made.
- `done {text, workbook_id, revision, thread_id, model, new_workbook_id?}`
- `error {code, message}`

Tool allowlist given to the model:
- `inspect_workbook`
- `read_range` (at most 400 cells)
- `read_source`
- `propose_workbook_edits`
- `apply_workbook_edits`
- `validate_workbook`
- `build_workbook`
- `run_template_extractor`

The Deep Agents tools are not available to the model: no filesystem, shell, todo or sub-agent tools.

Agent edits of up to 25 cells are auto-applied. Larger edits (up to 1000 cells) emit `confirmation_required` and wait for one of these:
- `/apply`
- an explicit approval message ("đồng ý", "ok", "yes", "apply"…) in the same thread

## Legacy extractor (unchanged)
- `POST /api/jobs`, `GET/DELETE /api/jobs/{job_id}`, `GET /api/jobs/{job_id}/download` and `/report`.
- The legacy HTML page is served at `/extractor`.
