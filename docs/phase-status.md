# Phase status

Only statuses from commands that were actually run are listed. Python commands run from the repo root with `.venv-ws/Scripts/python.exe`, shortened to `py` below.

## Phase 0 — Repo audit & baseline — **PASS** (2026-10-13)

| Check | Command | Result |
|---|---|---|
| env | `uv venv .venv-ws --python 3.12` + `uv pip install -r requirements.txt deepagents XlsxWriter fastapi[standard] python-multipart pytest httpx` | OK, `import xlsx_template.graph, deepagents` works |
| legacy plan apply | `py -m xlsx_template "Report_WBS (5) (1) 1.xlsx" -o <scratch>/wbs.xlsx --plan output/Report_WBS_TEMPLATE_final.plan.json --no-cache` | validation **PASS** in 3.5 s, cells_cleared 1722, formulas_preserved 1511, constants_kept 173, parts identical/changed 22/8, opens_in_openpyxl 1 |
| legacy dry run (real Groq) | `py -m xlsx_template "Report_WBS (5) (1) 1.xlsx" -o <scratch>/dry.xlsx --dry-run --no-cache` | plan written, 126 regions, 2 calls, 11 680 in / 1 371 out tokens, 161 s; one batch hit HTTP 413 on qwen (7082 > 7000) and fell back to gpt-oss-120b |
| Groq capabilities | probe script (bind_tools + json_schema) | qwen3.8-27b and gpt-oss-120b both **work** for tool calling and json_schema |
| legacy regression tests | `py -m pytest tests/test_legacy_extractor.py` | **12 passed** in 4.9 s |

Files written:
- `docs/implementation-notes.md`, `docs/compatibility-matrix.md`, `docs/phase-status.md`
- `.env.example`
- `pytest.ini`
- `tests/` (conftest, fixtures, legacy tests)
- `requirements.txt` (added lines; existing lines unchanged)

No code was deleted.

## Phase 1 — Workbook service, OOXML edit, preview, API — **PASS** (2026-10-13)

| Check | Command | Result |
|---|---|---|
| upload safety (ext/signature/zip bomb/xlsm refused) | `py -m pytest tests/test_upload.py` | 14 passed |
| OOXML `set_value`/`clear_value`, untouched parts byte-equal, shared formulas kept | `py -m pytest tests/test_ooxml_edit.py` | 10 passed |
| FortuneSheet preview DTO + compat warnings | `py -m pytest tests/test_preview.py` | 4 passed |
| revisions, 409 on stale base, undo = new revision | `py -m pytest tests/test_service.py tests/test_api.py` | 7 passed |
| full suite at end of phase | `py -m pytest -q` | 47 passed |

Not claimed: full XLSX fidelity. Unsupported features are listed in `compatibility-matrix.md` and shown as preview warnings.

## Phase 2 — Analyze → clarify → generate (scratch) — **PASS** (2026-10-13)

| Check | Command | Result |
|---|---|---|
| txt/csv facts with provenance, VN/EN number parsing | `py -m pytest tests/test_generate.py -k parse_number` | 9 passed |
| generate from sources: only source numbers in the file, live formulas `=B5*C5`, `=SUM`, cached values, `fullCalcOnLoad`, metadata `origin=generated` | `test_generate_from_sources_uses_only_source_numbers` | pass |
| no data → questions (currency, missing_numbers…) → blank inputs, **0 numeric constants** | `test_no_numbers_asks_then_leaves_inputs_blank` | pass |
| user figures used with `answer:figures` provenance | `test_user_figures_are_used_with_provenance` | pass |
| formula guard (allowlist, no invented constants, no direct refs) + LLM layout cannot type data + mock_fallback | `test_formula_guard`, `test_llm_layout_never_types_data`, `test_ground_check_blanks_unknown_numbers` | pass |
| group-by summary sheet with SUMIF (mock and LLM path; unit price never summed) | `test_group_sheet_sumif_mock_and_layout` | pass |
| full suite | `py -m pytest -q` | **65 passed** in 8.4 s |
| live Groq (manual smoke, not in CI) | build_spec on a sales CSV + notes, 2 prompts | `groq:qwen/qwen3.8-27b`, 2.7 s and 1.3 s; detail + SUMIF group sheet + facts sheet; 18/18 and 25/25 formulas computed |

Design: the LLM proposes **layout only** (sheets, columns mapped to source headers, formula templates, totals, group-by). Data rows are copied deterministically from sources and every numeric value is grounded against source facts. Deep Agents comes in Phase 4 (chat), not here.
Limitations: no charts (warning shown); only txt/csv sources; the formula preview evaluator covers the allowlisted functions only (Excel recalculates everything on open).

## Phase 3 — Template mode: gallery, slot detection, fill — **PASS** (2026-10-13)

| Check | Command | Result |
|---|---|---|
| Gallery: 2 built-ins (Hóa đơn bán hàng, Ngân sách chi phí) created once; user upload; preview/slots; 400 on non-xlsx; 404 on unknown id | `test_gallery_builtins_preview_and_upload` | pass |
| Invoice fill from CSV+TXT. Revision 2 has KV B3/B5/B6 (date serial), item rows, derived STT 1..3, formulas `=IF(C9="","",C9*D9)` and `=SUM(E9:E20)` intact, styles kept, VAT left blank with a warning, `fill_report` with fact_id/locator | `test_fill_builtin_invoice_from_sources` | pass |
| Uploaded x14 template. Dropdown and `x14:dataValidation` kept, formulas intact, untouched ZIP members byte-equal, original upload sha256 unchanged | `test_fill_uploaded_x14_template_preserves_parts` | pass |
| No sources → the single `figures` question is asked; a text answer fills a KV slot; an empty answer leaves revision 1 with an "unchanged template" warning; unknown template_id gives 404 | `test_template_without_sources_asks_and_uses_answers` | pass |
| Row overflow and an unmatched source table are warned | `test_overflow_and_unmatched_are_warned` | pass |
| LLM pairs ids only; unknown or invented ids are dropped; LLM failure → deterministic with a warning | `test_llm_mapping_pairs_ids_only` (monkeypatched) | pass |
| Populated template (WBS) detected as prefilled; `extractor_available=false` and `/templates/extract` returns 503 in mock mode | `test_populated_template_is_detected_and_extractor_gated` | pass |
| Legacy extractor unchanged | `tests/test_legacy_extractor.py` | pass |
| Full suite | `py -m pytest -q` | **75 passed** in 8.1 s |
| Live Groq (manual smoke, not in CI) | Invoice with differently worded labels (Bên mua, MST/Số HĐ, Ngày lập, Địa chỉ giao; CSV Tên hàng/SL/Giá) | `groq:qwen/qwen3.8-27b`, 1.2 s, 12 cells filled, B3/B4/B5/B6 via `llm`, all values traced to `line:N`/`row:N`; only "VAT (%)" warned |

Bugs found and fixed in this phase:
1. `parse_number` read "15/09/2026" as 15092026 and "Quý 3/2026" as 32026, which corrupted dates in the TXT and answer paths. Dates and multi-group digit text are now rejected, with regression cases added to `test_parse_number`.
2. The template-mapping LLM client was missing `api_key`. It was caught by the live smoke test, where it fell back honestly to deterministic mode.

Limitations:
- No row insertion when the source has more rows than the template; the extra rows are warned about.
- The Make-blank extractor needs Groq and is not exercised live in CI.
- KV detection is heuristic, so it gives noisy extra slots on dense sheets. Unmatched slots are only listed and never filled.

## Phase 4 — Interactive editing (manual + agent chat) — **PASS with limitations** (2026-10-13)

| Check | Command | Result |
|---|---|---|
| Chat edit via the Deep Agents loop (mock model): tool_started → workbook_revision → tool_completed, revision 2, formula `D3` kept, thread saved, log has addresses only | `tests/test_agent.py::test_chat_sets_cell_and_streams_events` | pass |
| No filesystem/shell/todo/sub-agent tools reach the model: exactly the 8 workbook tools are bound | `test_no_filesystem_or_shell_tools_are_exposed` | pass |
| read_range / inspect / unknown request are read-only; mock mode says it is mock | `test_read_and_inspect_are_read_only` | pass |
| > 25 cells → `confirmation_required` with a diff, nothing written; "đồng ý" in the same thread applies; re-apply → 422 | `test_large_edit_needs_confirmation_then_applies` | pass |
| Plan apply/reject over HTTP; unknown plan 404; stale plan 409 | `test_plan_apply_and_reject_over_http` | pass |
| Ungrounded numbers refused, formulas refused, numbers read this turn are grounded, read cap 400 cells (`MAX_READ_CELLS`) | `test_ungrounded_numbers_and_formulas_are_refused` | pass |
| Stale turn revision 409; unknown thread 404 | `test_stale_revision_is_409` | pass |
| A reply that claims an edit without any write tool is flagged in the reply, the stream and the thread | `test_unbacked_edit_claim_is_flagged` | pass |
| **Gate:** manual edit → agent edit → export holds both, preview 200, origins upload/manual/agent → undo gives revision 4 with the manual state | `test_gate_manual_then_agent_then_export_and_undo` | pass |
| Cache-strip regression (a formula cell keeping `<v xml:space>` text made the export unreadable) | `tests/test_ooxml_edit.py::test_stripped_cache_leaves_no_text_in_numeric_cells` | pass (fails on the old regex) |
| Full backend suite | `.venv-ws/Scripts/python.exe -m pytest -q` | **85 passed** in 13.8 s (3 reruns: all green) |
| Frontend | `npm run build` / `npm run lint` | build OK; oxlint exit 0, 0 errors, warnings only (react refs / set-state-in-effect in `SpreadsheetPanel.tsx`) |

Live checks (manual, Groq `qwen/qwen3.8-27b`, not in CI):
- "Số lượng Mouse thực tế là 40": inspect → read → propose. Revision 2 with B4 = 40 in 12 s.
- "Điền doanh thu tháng 10 vào E3" (no figure given): the model asked for the number instead of inventing one.
- "Thêm công thức tổng": the model said formulas are not supported.
- Browser E2E on the real WBS workbook (`docs/screenshots/07-chat-confirm.png`, `08-chat-plan-applied.png`):
  - Chat edit of WBS!E8 → r6, and the grid reloaded.
  - "Xóa toàn bộ nội dung vùng WBS!H300:K310" → a ConfirmCard with 44 rows → Apply → r7. Preview and export are 200, and the revision log shows origins `agent`, `agent`.

Bugs found and fixed in this phase:
1. **Export/preview broke after a manual edit.** `_strip_formula_cache` removed `t="str"` but kept `<v xml:space="preserve">text</v>`, so a numeric cell held text (Daily Report!L26). Fixed the regex in `ooxml_edit.py` and `validate_edit.py`, and added a regression test. The legacy `xlsx_template/writer.py:46` has the same pattern; it is left unchanged and reported.
2. **Model claimed an action it never performed.** In the first browser run, the model replied "Đã xóa … (revision 6)" without calling any tool. Fixed in two ways:
   - The runtime now appends a visible warning ("lượt này KHÔNG có thay đổi nào…") when a reply claims a change but no write tool ran.
   - The system prompt forbids claiming edits without a tool result.
   On the retry the model called the tool and the confirmation flow worked.
3. Flaky test: the log assertion matched "25" inside random ids and timestamps. It is now a structural check.

Limitations (honest):
- The agent only sets or clears values. Formula, style and structural ops stay disabled, per roadmap item 5, because there is no tested OOXML writer for them.
- Derived numbers are refused. A total that the model computes itself is not grounded, so it is not written; the user must give the number, or the cell must keep a formula.
- Mock mode is a regex command parser, not an LLM. It is labelled as mock in the UI and in its replies.
- Model fallback only happens if nothing was written yet in that turn.
- The claim detector is a VI/EN regex. It can miss unusual wording; the revision number in the header and the tool cards remain the source of truth.
- Plans may include no-op rows (clearing cells that are already empty). They are harmless but noisy.

## Follow-up (2026-10-13) — clarification gaps found with `sample_data/`

None of the sample sets triggered a clarification question. The causes and fixes:

- **Template mode only asked when there were no sources at all.** Now `analyze` runs the deterministic slot matcher (no LLM call). If any labelled template cells still have no source value, it lists them in one optional `figures` question. The answers become `user_confirmed` facts. A blank answer leaves the cells blank, with a warning as before.
- **Mock chat answered vague write requests with help text.** It now asks which cells and which values. A request for a formula, formatting or row insert gets an "unsupported" reply.
- **Groq chat, in a fresh thread, turned "điền doanh thu tháng 10" into renaming the A1 title.** The system prompt now requires one clarifying question, and no write tool, when the cells or values are not given. After the change:
  - 4 vague prompts (fresh workbook each) all asked back, at revision 1.
  - In 3 repeats of the hardest prompt, 1 got a reply with no write; the other 2 were rate-limited and changed nothing.
- **Leading zeros were lost.** `parse_number` turned `0100000000` (tax code) and `0901 234 567` (phone) into numbers. Values that start with `0` followed by a digit now stay text.
- Full suite: `.venv-ws/Scripts/python.exe -m pytest -q` → **92 passed**.

## Follow-up (2026-10-14) — AI field mapping with confirmation (template mode)

Goal: arbitrary content against an arbitrary template. The agent decides, for every field, where its value comes from, shows that decision, and asks only about uncertain or missing fields.

- **`analyze` builds a proposal** (`template_fill.propose`):
  - Exact label/header match comes first.
  - In Groq mode the model then pairs fields by id. It may only point at a fact id, a source column, or a `quote:` copied verbatim from a source text (whole words, one line, not from CSV files). Targets that cannot be found are dropped.
  - Fuzzy matches become unconfident guesses.
  - A leftover source table whose headers differ is guessed for an unmatched template table. Its columns are asked, and the table is not filled unless at least one column is chosen.
- **Response `mapping`** (one row per input cell or column: target, source, value, via, status `auto|confirm|missing`) is rendered by `MappingTable` in the create chat (see next follow-up).
- **Questions:**
  - A `map_k*` / `map_t*c*` select question per guess. The default is the AI's suggestion, plus a `(để trống)` option.
  - At most 7 such questions. Guesses that were not asked are not filled.
  - Missing fields still go to the `figures` question.
- **`generate` reuses the confirmed proposal** (`fills_from`) with no second model call. Values are always read from the sources.
- **`sample_data/_verify.py`** prints the mapping and accepts default answers.
  - Set 02 in mock mode: 2 label guesses, plus 5 column questions for the guessed table.
  - Set 02 in Groq mode: every field and column is mapped, and only VAT is asked.
  - A Groq run first quoted a lone `0` from a CSV row for VAT. Quotes are now whole-word and limited to non-CSV text files.
- Full suite: **95 passed**. `npx tsc -b --noEmit` is clean.

## Follow-up (2026-10-14) — chat-driven create page (wizard removed)

The step-by-step wizard (`pages/Wizard.tsx`) is gone. `/new/template` and `/new/scratch` now render `pages/Studio.tsx`, with a read-only grid on the left and the agent chat on the right. On narrow screens the two are stacked.

- **Choosing a template:**
  - The chat shows one template dropdown (`components/workspace/TemplatePicker.tsx`). Each gallery entry has a thumbnail drawn on a canvas from its real read-only preview (`TemplateThumb.tsx`: fills, text and merges of the top-left of the active sheet), so new templates get one automatically. The last entry is "Upload your own template (.xlsx)". A dropped .xlsx file also becomes the template.
  - The selected template is previewed on the left (`components/workspace/PreviewGrid.tsx`, FortuneSheet `allowEdit={false}`). The preview's compatibility-note count is shown.
  - Nothing typed in the preview is saved. The real workbook only exists after generate.
- **New route:** `GET /api/workspaces/{ws}/files/{file}/preview`.
  - It returns a read-only preview of an .xlsx attached to that workspace and writes nothing.
  - A non-xlsx file or a file from another workspace returns 404.
  - The test is `test_workspace_template_preview`.
- **Chat flow:**
  - The user attaches .txt/.csv sources with the clip button and sends a prompt.
  - The analyze result is rendered inline. Depending on the state, it shows the plan, the `MappingTable` with `ClarifyForm`, or a blocked message.
  - Answers are posted to `clarify`.
  - "Generate workbook" streams the `RunTimeline` inline and opens `/w/:id` when done.
  - Changing the template or the sources after a plan marks the old plan stale.
- **Verified in the browser (mock mode, set 02):**
  - The uploaded template previews correctly.
  - Clarify, then plan, then generate opens `/w/…` at r2.
  - Switching to a built-in template updates the preview.
  - Scratch mode works.
- **Known limitations:**
  - The FortuneSheet name box shows `A1:NaN` (cosmetic).
  - A backend started before this change lacks the new route. Restart it for uploaded-template previews; built-in previews still work.
- Full suite: **96 passed**. `tsc` is clean and `vite build` passes.

## Follow-up (2026-10-14) — AI-proposed template structure from prose documents

Problem: when the sources are prose (e.g. a change request plus a monthly report), there are no source columns to map, so the old flow asked which column each field came from. Now the AI reads the documents and proposes the template table's final column structure. The user confirms it, and generate rewrites the template copy and fills it.

- **When it runs** (`clarify._template_structure`):
  - Template mode with Groq on and at least one source.
  - The sources have no tables, or the prompt/documents ask for column/formula changes (`EDIT_WORDS`).
  - On any model error the old field mapping runs, with a warning. A Groq per-minute limit gives a plain "wait 1-2 minutes and send again" warning.
- **Proposal** (`workspace/restructure.py`, `propose` → `normalize`):
  - Final ordered columns: keep / rename / new / input / formula. Formulas use `{Header}` keys only; they are checked (known headers, no cell or sheet refs, supported functions, constants in a safe list or the documents).
  - Percent columns are scaled exactly once according to their format (`fix_percent`).
  - A column with a date number format in the template stays a date, whatever the model guesses.
  - A template lookup column (VLOOKUP etc.) becomes an input column when the documents state its values. Otherwise stated names and prices would be replaced by "unknown code" or 0.
  - Header fields and rows only take values backed by a verbatim quote from one source line. Anything else stays blank.
- **Review in chat** (`components/wizard/StructurePlan.tsx`):
  - Columns with letter, old header, change badges, formula and format.
  - Removed or hidden columns with what still uses them.
  - Header values with their source line, plus the first 20 rows.
- **Questions:**
  - `structure_apply`: apply the proposal, or keep the columns and only fill.
  - `drop_<col>`: hide or keep each removed column that other formulas still use.
  - A follow-up message re-analyzes with all prompts since the inputs last changed, so the user can adjust the proposal in chat.
- **Generate** (`restructure.run`):
  - Revision 1 is the untouched template.
  - Revision 2 is produced by `service.apply_transform`, a pure OOXML rewrite. It moves/inserts/renames columns and remaps formulas on every sheet, definedNames, charts, autoFilter, widths, merges, conditional formats and validations.
  - The new revision is validated: XML parses, `styles.xml` is unchanged, no new `#REF!`, headers are correct. A failed check keeps revision 1 and reports the error.
- **Limits:**
  - Not supported: templates with Excel table objects (ListObject) or shared formulas (values only), and reordering existing columns.
  - A removed column still referenced elsewhere is hidden, not deleted.
  - Groq free tier allows about one proposal per 1-2 minutes on the default model.
  - Cached values are not written; Excel recalculates on open.
  - Visual fidelity is close to the template, not guaranteed identical.
- **Verified live** (Groq, `bao_cao_kinh_doanh_thang_09_2026.txt` + `yeu_cau_chinh_sua_so_chi_tiet_ban_hang.txt` on `template_so_ban_hang.xlsx`):
  - 17 columns A..Q: Mã KH and Tháng are hidden because 'Tổng hợp' uses them.
  - The requested formulas are `J*K/100`, `(J-L)*M/100` and `J-L+N`.
  - The 3 rows are filled from the request and the report adds nothing.
  - 'Tổng hợp' references follow the moved columns.
- Full suite: **106 passed**. `tsc` is clean and `vite build` passes.

### Report adaptation: understanding, review and approval (2026-10-14)
- **Pipeline:** parse sources -> AI understanding (template purpose, document summary, fit, entities, derivable,
  missing, assumptions) -> proposed columns/labels/header fields/extra sheets -> **dry run** of the plan on a temporary
  copy (OOXML writer + validator, read back with the safe evaluator) -> review -> approve -> generate -> validate.
- **Approval is enforced in the backend:**
  - `analyze` always returns `structure_apply` (and `structure_sheets` when sheets are proposed) for a structure plan.
  - `clarify` re-runs the dry run on the user's choices; `structure.approved` is true only when that dry run is valid.
  - `generate` refuses an unapproved plan (422). `AnalyzeRequest.auto_approve=true` applies the defaults and approves
    in one step, still subject to a valid dry run.
- **Review UI** (`StructurePlan.tsx`): A. what the documents say; B. template vs proposed columns with reasons and
  label edits; C. data preview from the dry run with provenance (explicit / derived / assumed / missing), formulas and
  quotes as tooltips, totals, extra-sheet grids and the missing-info list. Actions: Approve plan / Request changes /
  Cancel; Generate stays disabled until approved.
- **Data rules:** every value needs a quote from a source line that contains it (whitespace- and case-insensitive);
  only 0 may be implied (marked *assumed*); otherwise the cell is left blank and listed as missing. Row formulas (main
  table and extra sheets) are blank-guarded on every same-row input, so a missing input never becomes a computed 0.
- **LLM providers:** Groq (`GROQ_API_KEY`) or Azure OpenAI (`AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT`,
  `AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_VERSION`, optional `AZURE_OPENAI_REASONING`). Configure them in `.env`
  (see `.env.example`); keys are never logged.
- **Verified live** (Azure gpt-5, `bao_cao_kinh_doanh_thang_09_2026.txt` + builtin "Hóa đơn bán hàng"):
  - 8 rows (2 branches x 4 products), net qty = shipped - returned, discount on post-return revenue, no VAT.
  - Totals: revenue after discount 781,376,800 (HN 467,953,200; HP 313,423,600).
  - Extra sheets: branch summary + receivables, costs (108,400,000), inventory (closing 17/28/43/83/14/16/26/62).
  - Generate before approval -> 422; after approval `matches_preview=true`. Excel recalculation: 0 error cells.
- **Limits:** data rows are limited to the template's table rows (overflow is reported, not inserted); the template's
  `N()`-based VAT formula shows as pending in the preview (Excel computes it); `docProps/app.xml` does not list new
  sheets; the AI may hide unused template columns (a hidden column can still hold the template's VAT input); values
  whose quote cannot be found stay blank, so some runs need a follow-up "Request changes".
- Full suite: **115 passed**. `tsc` is clean and `vite build` passes.

## Phase 5 — Polish & expand safe surface — **NOT STARTED**
This was outside the master prompt's scope (Phase 0 → 4). No row/column insert, styles, formula writing, perf profiling, dark mode or auth hardening has been built. Nothing in the UI claims them.

### Agent editor: deterministic conversions (2026-10-14)
- Problem seen live: "Cột này thành tiền đô" + "1USD = 25000VND" looped — the model computed 636/138/... itself,
  `propose_workbook_edits` refused them as ungrounded each time, and the turn ended with "Model không phản hồi".
- New tool `transform_range(range, op, factor, round_digits=2)`: the factor (or its reciprocal) must be typed by the
  user/sources; the new values are computed in code from the cells' stored numbers; formula and text cells are
  skipped (formulas recalculate from their inputs); same auto-apply (<=25 cells) / confirmation path as other edits.
  Warns when the number format is a currency or hides decimals (the editor cannot change formats).
- Ungrounded-number errors now point the model at `transform_range`; prompt: "cột này/ô này" = the Selection, don't
  re-ask a target already named, "xác nhận" means do it; a turn that dies after tool refusals reports the refusal
  instead of "Model không phản hồi".
- Live (Azure gpt-5, Selection G9): turn 1 asks the rate, turn 2 "1USD = 25000VND" -> revision 2, G9:G16 =
  636/138/50/15.6, H-K formulas untouched. Tests: tests/test_agent.py (transform + loop message).

### Restructure: STT, paraphrased quotes, unused inputs (2026-10-14)
- Live 'Hóa đơn' bug (user screenshot): blank STT column and blank Số lượng/Thành tiền on the Hải Phòng LT14 row
  although the source says "12 chiếc LT14". Causes were in the pipeline, not the data:
  - a template STT/TT/No/# column kept with no values stayed blank -> such columns now become an auto `sequence`
    (1..n, not counted as a user-visible change when the template already had it);
  - the model quoted the figure with extra words not in the source, so grounding blanked it -> for numeric cells the
    longest verbatim span of the quote that still contains the number (and the word after it) is used as the quote;
    implied-0 cells keep the longest verbatim span the same way. A value absent from the source is still refused.
- One self-review round: if a numeric input column (e.g. returns) is used by no formula — including kept template
  formulas — the model gets its answer back with the problem list; the fix is adopted only if it has fewer issues.
  Otherwise a warning names the unused column. Tests in tests/test_report_adapt.py (120 passed).

### LLM provider: Azure OpenAI only (2026-10-14)
- Groq removed from the workspace: `config.chat_models()` returns only Azure OpenAI; without it the app is in mock mode
  (`llm_mode` is now `"azure"` | `"mock"`, the health badge says "Azure OpenAI · gpt-5").
- Variables: `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT` (default gpt-5); the docker
  style `OPENAI_API_KEY` + `API_BASE=https://<res>.openai.azure.com/openai/deployments/<dep>` + `LANGUAGE_MODEL` also work.
- The legacy `xlsx_template/` package (unchanged, still Groq-only as a standalone CLI) runs from the workspace through
  `extractor_adapter._azure_decider()`, which swaps its decider for one that calls Azure.
- Live: invoice template + bao_cao_kinh_doanh_thang_09_2026.txt on azure:gpt-5 -> STT 1..8, HP LT14 = 12, returns
  column feeds "Bán ròng" and every money formula uses net qty. 123 tests pass.
