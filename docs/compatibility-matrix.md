# XLSX compatibility matrix

The source of truth is the code plus the tests listed here. The editor does not reach full XLSX fidelity, and no part of this doc claims it does. "Preserved" means the bytes or XML for that feature come through an edit unchanged, and validator levels L0–L4 check this on every revision (`workspace/validate_edit.py`).

## Input files
| Input | Status | Where enforced / tested |
|---|---|---|
| `.xlsx` | Supported | `upload.check_xlsx`, `tests/test_upload.py` |
| `.xlsm` / VBA inside `.xlsx` | **Rejected**. Macros are never executed or kept | `test_refused_extensions`, `test_vba_rejected_even_with_xlsx_extension` |
| `.xlsb`, `.xls`, encrypted (OLE) | **Rejected** with a clear message | `test_refused_extensions`, `test_encrypted_ole_rejected` |
| Zip bomb, path traversal, duplicate members | **Rejected** | `test_zip_bomb_rejected`, `test_path_traversal_rejected` |
| `.txt` / `.csv` sources | Supported (UTF-8 with or without BOM, then cp1258/cp1252 fallback; UTF-16 and binary content are rejected) | `test_text_sources_decoded` |

Note: the legacy extractor (`/api/jobs`) still accepts `.xlsm`. That behaviour was left unchanged on purpose.

## Editing existing workbooks (`workspace/ooxml_edit.py`)
| Operation | Status | Notes |
|---|---|---|
| `set_value` text / number / bool / date / blank | **Supported** | Text goes to sharedStrings (append-only). New cells inherit the row style, otherwise the column style; existing cells keep their style id |
| `clear_value` | **Supported** | Keeps the cell style |
| Text starting with `=` | Stored as **literal text** | Formula-injection guard (`test_equals_text_stored_as_literal_not_formula`) |
| Writing to a formula cell | **Refused (422)** | Formulas are never overwritten (`test_formula_target_rejected`) |
| Writing to a non-anchor merged cell | **Refused (422)** | `test_merged_non_anchor_rejected` |
| `set_formula`, `set_style`, insert/delete rows or cols, merge, sheet add/rename | **Unsupported**, rejected (422) | Disabled in the UI with a message |
| Formula results after an edit | Cached `<v>` is stripped in the edited sheet and every sheet that references it, with `fullCalcOnLoad="1"` set | Excel/LibreOffice recalculate on open. The preview shows such formulas without values (counted in `stale_formulas`) |

## Preserved through edits (byte-equal or XML-equal outside the edited cells)
| Feature | Status | Evidence |
|---|---|---|
| Untouched ZIP members (styles, theme, drawings, comments, tables, pivots, charts, printer settings…) | Byte-equal | L1 check; `test_set_values_validate_and_preserve` |
| Worksheet XML outside `<sheetData>` (merges, conditional formatting, data validation, `extLst`/x14 dropdowns, page setup) | Byte-equal except `<dimension ref>` | L2 check; `test_x14_extension_survives_edit` |
| Formulas (text) and cell style ids | Unchanged | L2 check; 1511 formulas on the WBS sample (`test_real_sample_edit`) |
| workbook.xml | Only `calcPr` may change | L2 check |
| Defined names, external links | Preserved. Formulas that depend on an external link get a warning | `ooxml_edit` |

## Preview (`workspace/preview.py` → FortuneSheet)
| Feature | Preview | Exported file |
|---|---|---|
| Values, number formats (common), dates, booleans | Shown | Preserved |
| Formulas (`f`) with cached values | Shown (FortuneSheet recalculates its own copy for display only) | Preserved |
| Font bold/italic/underline/strike/size/name/color, fill, alignment, wrap | Shown | Preserved |
| Borders (thin/medium/thick/dashed/dotted/double) | Shown | Preserved |
| Merges, column widths, row heights, hidden rows/cols/sheets, frozen panes | Shown | Preserved |
| Theme/indexed colours | Approximated (theme + tint, indexed palette) | Preserved |
| Data validation (incl. x14), conditional formatting, comments, hyperlinks | **Warning only** (not rendered) | Preserved |
| Charts, images/drawings, pivot tables, Excel tables, sheet protection, chart sheets | **Warning only** | Preserved |
| Very large sheets | Truncated to `MAX_PREVIEW_CELLS`, with a warning | Full file preserved |

## Generated workbooks (scratch mode, Phase 2: XlsxWriter)
These are produced from a validated `WorkbookSpec`. Supported features: styled title, headers, typed columns (text, number, integer, currency, percent, date), row formulas from an allowlisted template, SUM totals, freeze header, column widths, notes, and three style presets. Charts are not generated.
Group-by summary sheets use `SUMIF` over the detail sheet with absolute ranges. Unit-price columns are never summed. Numeric constants in the file come only from source facts or user answers, and numeric literals in formulas must be grounded (0–12, 100, 1000 or a source number).

## Template fill (template mode, Phase 3)
| Feature | Support |
|---|---|
| Slot detection | Two kinds. KV slots: a label cell (≤60 chars, not a total) whose right-hand neighbour is empty, where the label ends with ":" or the neighbour is styled. Table slots: a header row of 2 or more text cells followed by empty or formula input rows, up to the total row (at most 1000 rows). Hidden sheets are skipped and reported. Merged headers are handled. |
| Writing | Only through `set_value` operations (OOXML patcher). Revision 1 holds the original bytes; revision 2 holds the fill. Formulas, styles, data validation (including x14), conditional formatting, defined names and untouched parts are preserved. This is covered by a test comparing untouched ZIP members byte for byte. |
| Mapping | Deterministic normalised-label similarity (accents stripped). In Azure mode the model only pairs slot ids with fact or column ids. Invented ids are dropped, and values always come from sources. A sequence column (STT/No) gets the derived row number. |
| Not supported (warned) | Inserting rows when the source has more rows than the template ("N row(s) did not fit"). Templates that already contain data, where 0 input rows are found; for these, a Make-blank suggestion runs the legacy extractor, which needs Azure OpenAI (its decider is pointed at Azure by the adapter). Charts and pivot tables are not filled. Labels without a source value are left blank and listed. |
| Dates | ISO and dd/mm/yyyy are written as date serials. The parser no longer reads "15/09/2026" or "Quý 3/2026" as numbers. |
