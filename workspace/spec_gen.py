"""Source facts + prompt + answers -> validated WorkbookSpec.

Grounding rule: data cells are copied from source tables/facts deterministically. The LLM (azure mode) only designs
the layout (titles, sheet names, which columns, types, derived formulas, totals, notes); it never types data values.
Formula templates are checked (allowlisted functions, column keys only, numeric literals must be grounded).
If the LLM fails, the deterministic (mock) layout is used and a warning says so.
"""
from __future__ import annotations

import json
import re
import unicodedata
from typing import Any, Literal

from pydantic import BaseModel, Field, ValidationError

from . import config, formula_eval, sources
from .builder import KEY_RE
from .clarify import is_vi
from .schemas import ColumnSpec, SheetSpec, WorkbookSpec

SAFE_LITERALS = {float(i) for i in range(13)} | {100.0, 1000.0}   # ROUND digits, months, %
LEADING_REQUEST = re.compile(r"^(hãy\s+)?(tạo|lập|làm|build|create|make|generate)\s+(cho\s+tôi\s+|me\s+)?(một\s+|a\s+|an\s+)?",
                             re.I)


def slug(text: str, used: set[str]) -> str:
    s = unicodedata.normalize("NFKD", text.replace("đ", "d").replace("Đ", "D"))
    s = re.sub(r"[^A-Za-z0-9]+", "_", s.encode("ascii", "ignore").decode()).strip("_").lower()[:30] or "col"
    if s[0].isdigit():
        s = "c_" + s
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def sheet_name(text: str, used: set[str]) -> str:
    s = re.sub(r"[\[\]:*?/\\']", " ", text).strip()[:28] or "Sheet"
    base, i = s, 2
    while s.lower() in used:
        s, i = f"{base[:25]} {i}", i + 1
    used.add(s.lower())
    return s


def title_from_prompt(prompt: str, vi: bool) -> str:
    first = (prompt or "").strip().splitlines()[0] if (prompt or "").strip() else ""
    first = LEADING_REQUEST.sub("", first)
    first = re.sub(r"\s*(?:,|kèm theo|kèm|có|và|with|and|including)?\s*(?:một\s+|a\s+)?(?:biểu\s*đồ|chart|graph|đồ\s*thị)(?:\s.*)?$",
                   "", first, flags=re.I).strip(" .:,")
    first = first[:1].upper() + first[1:] if first else ("Bảng tính" if vi else "Workbook")
    return first[:80]


def col_type(header: str, values: list[Any]) -> str:
    vals = [v for v in values if v not in (None, "")]
    if vals and all(isinstance(v, str) and builder_date(v) for v in vals):
        return "date"
    if not vals or not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
        return "text"
    if sources.PERCENT_WORDS.search(header) and all(abs(v) <= 1.5 for v in vals):
        return "percent"
    if sources.QTY_WORDS.search(header):
        return "integer" if all(float(v).is_integer() for v in vals) else "number"
    if sources.MONEY_WORDS.search(header):
        return "currency"
    return "integer" if all(float(v).is_integer() for v in vals) else "number"


def builder_date(v: str) -> bool:
    from .builder import to_date
    return to_date(v) is not None


def _summable(col: ColumnSpec) -> bool:
    if col.type not in ("currency", "number", "integer"):
        return False
    h = col.header
    if sources.UNIT_PRICE_WORDS.search(h) and not sources.AMOUNT_WORDS.search(h):
        return False                                   # summing unit prices is meaningless
    return not re.search(r"\b(id|mã|code|năm|year|stt|no\.?)\b", h, re.I)


# ---------------------------------------------------------------- deterministic layout
def table_sheet(t: sources.Table, vi: bool, used_sheets: set[str], cols_override: list[dict] | None = None,
                title: str | None = None) -> SheetSpec:
    used: set[str] = set()
    cols, rows_src = [], []
    for j, h in enumerate(t.headers):
        cols.append(ColumnSpec(key=slug(h, used), header=h[:120], type=col_type(h, [r[j] for r in t.rows])))
    qty = next((c for c in cols if sources.QTY_WORDS.search(c.header) and c.type in ("integer", "number")), None)
    price = next((c for c in cols if sources.UNIT_PRICE_WORDS.search(c.header) and c.type in ("currency", "number",
                                                                                             "integer")), None)
    has_amount = any(sources.AMOUNT_WORDS.search(c.header) for c in cols)
    if qty and price and qty is not price and not has_amount:
        cols.append(ColumnSpec(key=slug("amount", used), header="Thành tiền" if vi else "Amount", type="currency",
                               formula=f"={{{qty.key}}}*{{{price.key}}}"))
    for r in t.rows:
        rows_src.append({c.key: v for c, v in zip(cols, r) if v not in (None, "")})
    return SheetSpec(name=sheet_name(re.sub(r"\.(csv|txt)$", "", t.name, flags=re.I), used_sheets), title=title,
                     subtitle=(f"Nguồn: {t.name}" if vi else f"Source: {t.name}"), columns=cols, rows=rows_src,
                     totals=[c.key for c in cols if _summable(c)][:10])


def facts_sheet(facts: list[sources.SourceFact], vi: bool, used_sheets: set[str], names: dict[str, str]) -> SheetSpec:
    cols = [ColumnSpec(key="item", header="Hạng mục" if vi else "Item", type="text", width=36),
            ColumnSpec(key="value", header="Giá trị" if vi else "Value", width=18,
                       type="integer" if all(float(f.value).is_integer() for f in facts
                                             if isinstance(f.value, (int, float))) else "number"),
            ColumnSpec(key="unit", header="Đơn vị" if vi else "Unit", type="text", width=10),
            ColumnSpec(key="source", header="Nguồn" if vi else "Source", type="text", width=40)]
    rows = []
    for f in facts:
        src = names.get(f.provenance.file_id or "", "prompt" if f.provenance.locator == "prompt" else "answer")
        rows.append({"item": f.field, "value": f.value, "unit": f.unit or "",
                     "source": f"{src} ({f.provenance.locator})"})
    return SheetSpec(name=sheet_name("Tổng hợp" if vi else "Summary", used_sheets), columns=cols, rows=rows,
                     subtitle="Mỗi giá trị kèm nguồn gốc" if vi else "Every value with its source")


GROUP_RE = re.compile(r"(?:theo|by|per|từng|mỗi)\s+([^\n,.;]{1,40})", re.I)


def group_sheet(detail: SheetSpec, group_key: str, value_keys: list[str], name: str, vi: bool,
                used_sheets: set[str], title: str | None = None) -> SheetSpec:
    """One row per distinct group value (copied text) and SUMIF formulas into the detail sheet: no typed numbers."""
    from xlsxwriter.utility import xl_col_to_name

    from .builder import _Layout
    L, keys = _Layout(detail), [c.key for c in detail.columns]
    first, last = L.first + 1, L.first + max(L.n, 1)
    gl = xl_col_to_name(keys.index(group_key))
    groups = list(dict.fromkeys(str(r[group_key]).strip() for r in detail.rows if r.get(group_key) not in (None, "")))
    gcol = next(c for c in detail.columns if c.key == group_key)
    cols = [ColumnSpec(key=group_key, header=gcol.header, type="text")]
    cols += [ColumnSpec(key=k, header=next(c.header for c in detail.columns if c.key == k),
                        type=next(c.type for c in detail.columns if c.key == k)) for k in value_keys]
    q = "'" + detail.name.replace("'", "''") + "'"
    formulas = {}
    for i in range(len(groups)):
        r = i + 2                                       # spec convention: header = row 1
        for j, k in enumerate(value_keys, start=1):
            vl = xl_col_to_name(keys.index(k))
            formulas[f"{xl_col_to_name(j)}{r}"] = f"=SUMIF({q}!${gl}${first}:${gl}${last},A{r},{q}!${vl}${first}:${vl}${last})"
    return SheetSpec(name=sheet_name(name, used_sheets), columns=cols, rows=[{group_key: g} for g in groups],
                     formulas=formulas, totals=value_keys,
                     title=title or (f"Tổng hợp theo {gcol.header}" if vi else f"Totals by {gcol.header}"),
                     subtitle=(f"Công thức SUMIF từ sheet '{detail.name}'" if vi else
                               f"SUMIF formulas over sheet '{detail.name}'"))


def default_columns(prompt: str, answers: dict, vi: bool) -> list[ColumnSpec]:
    used: set[str] = set()
    wanted = [c.strip() for c in re.split(r"[,;\n]", str(answers.get("columns") or "")) if c.strip()][:30]
    if wanted:
        cols = []
        for h in wanted:
            t = "currency" if sources.MONEY_WORDS.search(h) else "integer" if sources.QTY_WORDS.search(h) else \
                "percent" if sources.PERCENT_WORDS.search(h) else "date" if re.search(r"ngày|date", h, re.I) else "text"
            cols.append(ColumnSpec(key=slug(h, used), header=h[:120], type=t))
        return cols
    if sources.QTY_WORDS.search(prompt) or re.search(r"hóa\s*đơn|invoice|báo\s*giá|quotation|đơn\s*hàng|order",
                                                       prompt, re.I):
        return [ColumnSpec(key="item", header="Mô tả" if vi else "Description", type="text", width=36),
                ColumnSpec(key="qty", header="Số lượng" if vi else "Quantity", type="integer"),
                ColumnSpec(key="unit_price", header="Đơn giá" if vi else "Unit price", type="currency"),
                ColumnSpec(key="amount", header="Thành tiền" if vi else "Amount", type="currency",
                           formula="={qty}*{unit_price}")]
    if re.search(r"chi\s*phí|expense|ngân\s*sách|budget", prompt, re.I):
        return [ColumnSpec(key="item", header="Khoản mục" if vi else "Item", type="text", width=32),
                ColumnSpec(key="planned", header="Kế hoạch" if vi else "Planned", type="currency"),
                ColumnSpec(key="actual", header="Thực tế" if vi else "Actual", type="currency"),
                ColumnSpec(key="variance", header="Chênh lệch" if vi else "Variance", type="currency",
                           formula="={actual}-{planned}"),
                ColumnSpec(key="note", header="Ghi chú" if vi else "Note", type="text", width=28)]
    return [ColumnSpec(key="item", header="Hạng mục" if vi else "Item", type="text", width=32),
            ColumnSpec(key="value", header="Giá trị" if vi else "Value",
                       type="currency" if sources.MONEY_WORDS.search(prompt) else "number"),
            ColumnSpec(key="note", header="Ghi chú" if vi else "Note", type="text", width=28)]


def mock_spec(b: sources.SourceBundle, prompt: str, answers: dict) -> tuple[WorkbookSpec, list[str]]:
    vi = is_vi(prompt) or any(is_vi(t) for t in b.texts.values())
    warnings: list[str] = []
    title = title_from_prompt(prompt, vi)
    period = answers.get("period") or b.period
    used: set[str] = set()
    sheets: list[SheetSpec] = []
    for i, t in enumerate(b.tables[:config.MAX_SHEETS - 1]):
        sheets.append(table_sheet(t, vi, used, title=title if i == 0 else None))
    if len(b.tables) > config.MAX_SHEETS - 1:
        warnings.append(f"Only the first {config.MAX_SHEETS - 1} source tables were used.")
    for phrase in GROUP_RE.findall(prompt or ""):         # "tổng theo sản phẩm" / "totals by region"
        for detail in list(sheets):
            g = next((c for c in detail.columns if c.type == "text" and c.header.strip().lower() in phrase.lower()),
                     None)
            vals = [c.key for c in detail.columns if _summable(c)]
            if g and vals and len(sheets) < config.MAX_SHEETS - 1:
                sheets.append(group_sheet(detail, g.key, vals, f"Theo {g.header}" if vi else f"By {g.header}",
                                          vi, used))
                break
    in_tables = {fid for t in b.tables for row in t.fact_ids for fid in row if fid}
    loose = [f for f in b.facts if f.id not in in_tables]
    if loose:
        sheets.append(facts_sheet(loose[:config.MAX_ROWS_PER_SHEET], vi, used, b.names))
    if not sheets:
        cols = default_columns(prompt, answers, vi)
        sheets.append(SheetSpec(
            name=sheet_name("Dữ liệu" if vi else "Data", used), title=title, columns=cols, blank_rows=15,
            totals=[c.key for c in cols if _summable(c)],
            notes=["Chưa có số liệu nguồn: các ô nhập (nền vàng nhạt) được để trống, không tự điền."
                   if vi else "No source figures: input cells (light yellow) are left blank on purpose."]))
        warnings.append("No source figures were provided; input cells are blank." if not vi else
                        "Không có số liệu nguồn; các ô nhập được để trống.")
    if period and sheets[0].subtitle is None:
        sheets[0].subtitle = (f"Kỳ: {period}" if vi else f"Period: {period}")
    elif period:
        sheets[0].subtitle = f"{sheets[0].subtitle} · {'Kỳ' if vi else 'Period'}: {period}"
    currency = answers.get("currency") or b.currency
    spec = WorkbookSpec(title=title, sheets=sheets, currency=currency)
    return spec, warnings


# ---------------------------------------------------------------- LLM layout (azure)
class LlmColumn(BaseModel):
    header: str = Field(max_length=120)
    source_header: str | None = Field(default=None, description="exact header of the source table column to copy, "
                                      "or null for a derived/blank column")
    type: Literal["text", "number", "integer", "currency", "percent", "date"]
    formula: str | None = Field(default=None, description="row template using other columns' headers in braces, "
                                "e.g. ={Quantity}*{Unit price}; null if not derived")


class LlmSheet(BaseModel):
    name: str = Field(max_length=31)
    title: str | None = Field(default=None, max_length=200)
    subtitle: str | None = Field(default=None, max_length=300)
    source_table: int | None = Field(default=None, description="index of the source table to copy rows from, or "
                                     "null for an input sheet with blank rows")
    include_facts: bool = Field(default=False, description="true for a summary sheet listing the key-value facts")
    columns: list[LlmColumn] = Field(default_factory=list, max_length=30)
    blank_rows: int = Field(default=0, ge=0, le=200)
    total_headers: list[str] = Field(default_factory=list, description="headers of columns to SUM in a totals row")
    group_by_sheet: str | None = Field(default=None, description="for a per-group summary: name of an EARLIER sheet "
                                       "in this layout; code writes one row per distinct value with SUMIF formulas")
    group_by_header: str | None = Field(default=None, description="header (in that earlier sheet) to group by")
    sum_headers: list[str] = Field(default_factory=list, description="headers (in that earlier sheet) to sum per "
                                   "group")
    notes: list[str] = Field(default_factory=list, max_length=5)


class LlmLayout(BaseModel):
    title: str = Field(max_length=120)
    style_preset: Literal["clean_finance", "minimal", "bold_header"] = "clean_finance"
    sheets: list[LlmSheet] = Field(min_length=1, max_length=8)


SYSTEM = """You design Excel workbook LAYOUTS. You never write data values: rows are copied by code from the source
tables you reference. Rules:
- Use source_table to point at a source table; map columns with source_header (exact header text).
- Derived columns use formula templates with other column headers in braces, e.g. ={Số lượng}*{Đơn giá}. Only
  + - * / and SUM, AVERAGE, MIN, MAX, ROUND, IF, IFERROR, ABS. No constants other than integers 0-12, 100, 1000 unless the number
  appears in the facts.
- Never invent numbers, rates, targets or examples. If data is missing, make an input sheet with blank_rows.
- Use include_facts=true for a sheet listing the key-value facts (code fills it).
- For totals per product/region/etc. add a sheet with group_by_sheet + group_by_header + sum_headers (headers of
  the earlier sheet, including derived ones); code builds SUMIF formulas. Do not copy rows for that.
- Write titles/headers in the user's language. Keep sheet names <= 31 chars, no []:*?/\\.
- Charts are not supported; mention that in notes if asked.
Return only the JSON object."""


def _context(b: sources.SourceBundle, prompt: str, answers: dict) -> str:
    tables = []
    for i, t in enumerate(b.tables[:8]):
        tables.append({"index": i, "name": t.name, "headers": t.headers, "row_count": len(t.rows),
                       "sample_rows": t.rows[:3]})
    in_tables = {fid for t in b.tables for row in t.fact_ids for fid in row if fid}
    facts = [{"field": f.field, "value": f.value, "unit": f.unit} for f in b.facts if f.id not in in_tables][:60]
    return json.dumps({"request": prompt[:4000], "answers": answers, "detected_currency": b.currency,
                       "detected_period": b.period, "source_tables": tables, "facts": facts}, ensure_ascii=False,
                      default=str)


def llm_layout(b: sources.SourceBundle, prompt: str, answers: dict) -> tuple[LlmLayout, str]:
    last: Exception | None = None
    for model, llm in config.chat_models(max_tokens=3000, timeout=90):
        try:
            out = llm.with_structured_output(LlmLayout, method="json_schema").invoke(
                [("system", SYSTEM), ("user", _context(b, prompt, answers))])
            if isinstance(out, LlmLayout):
                return out, model
            last = ValueError("empty structured output")
        except Exception as e:  # noqa: BLE001 - fall through to the next model / mock
            last = e
    raise RuntimeError(f"{type(last).__name__}: {str(last)[:200]}")


def check_formula(template: str, keys: set[str], allowed: set[float]) -> str | None:
    """Returns an error string, or None when the row-template formula is safe."""
    body = template.strip().lstrip("=")
    if not body or len(body) > 400:
        return "empty or too long"
    for k in KEY_RE.findall(body):
        if k not in keys:
            return f"unknown column {k}"
    probe = KEY_RE.sub("A1", body)
    if re.search(r"[A-Za-z]{1,3}\$?\d+|!|\[|'", KEY_RE.sub("", body)):
        return "direct cell/sheet references are not allowed"
    funcs = formula_eval.functions_used("=" + probe)
    if "?" in funcs or not funcs <= formula_eval.SUPPORTED_FUNCTIONS:
        return f"unsupported function(s) {sorted(funcs - formula_eval.SUPPORTED_FUNCTIONS)}"
    for lit in re.findall(r"(?<![A-Za-z_\d.])\d+(?:\.\d+)?", KEY_RE.sub("", body)):
        if float(lit) not in SAFE_LITERALS and round(float(lit), 9) not in allowed:
            return f"ungrounded constant {lit}"
    return None


def from_layout(layout: LlmLayout, b: sources.SourceBundle, prompt: str, answers: dict
                ) -> tuple[WorkbookSpec, list[str]]:
    vi = is_vi(prompt) or any(is_vi(t) for t in b.texts.values())
    warnings: list[str] = []
    allowed = b.allowed_numbers() | SAFE_LITERALS
    used_sheets: set[str] = set()
    sheets: list[SheetSpec] = []
    in_tables = {fid for t in b.tables for row in t.fact_ids for fid in row if fid}
    built: dict[str, SheetSpec] = {}
    for ls in layout.sheets:
        if ls.group_by_sheet and ls.group_by_header:
            ls.source_table, ls.columns = None, ls.columns or [LlmColumn(header=ls.group_by_header, type="text")]
        if ls.include_facts:
            loose = [f for f in b.facts if f.id not in in_tables]
            if loose:
                fs = facts_sheet(loose, vi, used_sheets, b.names)
                fs.title, fs.notes = ls.title, ls.notes[:5]
                sheets.append(fs)
            continue
        t = b.tables[ls.source_table] if ls.source_table is not None and 0 <= ls.source_table < len(b.tables) else None
        used: set[str] = set()
        cols: list[ColumnSpec] = []
        by_header: dict[str, str] = {}
        src_idx: dict[str, int] = {}
        for lc in ls.columns:
            key = slug(lc.header, used)
            by_header[lc.header.strip().lower()] = key
            if t is not None and lc.source_header:
                j = next((i for i, h in enumerate(t.headers) if h.strip().lower() == lc.source_header.strip().lower()),
                         None)
                if j is None:
                    warnings.append(f"'{lc.header}': source column '{lc.source_header}' not found; left blank.")
                else:
                    src_idx[key] = j
            cols.append(ColumnSpec(key=key, header=lc.header, type=lc.type))
        if not cols:
            continue
        keys = {c.key for c in cols}
        for lc, col in zip(ls.columns, cols):
            if lc.formula:
                tmpl = re.sub(r"\{([^{}]{1,120})\}", lambda m: "{" + by_header.get(m[1].strip().lower(), "?") + "}",
                              lc.formula)
                err = check_formula(tmpl, keys - {col.key}, allowed) if "{?}" not in tmpl else "unknown column"
                if err:
                    warnings.append(f"Formula for '{lc.header}' dropped ({err}).")
                else:
                    col.formula = tmpl if tmpl.startswith("=") else "=" + tmpl
        rows = []
        if t is not None:
            for r in t.rows:
                rows.append({k: r[j] for k, j in src_idx.items() if r[j] not in (None, "")})
            for col in cols:                           # mapped columns: type follows the data, LLM may only say currency
                if col.key in src_idx:
                    inferred = col_type(t.headers[src_idx[col.key]], [r[src_idx[col.key]] for r in t.rows])
                    money = col.type == "currency" or sources.MONEY_WORDS.search(col.header)
                    col.type = "currency" if inferred in ("number", "integer") and money and not \
                        sources.QTY_WORDS.search(col.header) else inferred
        by_key = {c.key: c for c in cols}
        for col in cols:                               # derived money columns
            if col.formula and col.type in ("number", "integer") and any(
                    by_key[k].type == "currency" for k in KEY_RE.findall(col.formula) if k in by_key):
                col.type = "currency"
        totals = [by_header[h.strip().lower()] for h in ls.total_headers
                  if h.strip().lower() in by_header and by_key[by_header[h.strip().lower()]].type in
                  ("number", "integer", "currency")]
        if not totals and (rows or ls.blank_rows):
            totals = [c.key for c in cols if _summable(c)]
        if ls.group_by_sheet and ls.group_by_header:
            detail = built.get(ls.group_by_sheet.strip().lower())
            gkey = next((c.key for c in (detail.columns if detail else [])
                         if c.header.strip().lower() == ls.group_by_header.strip().lower() and c.type == "text"), None)
            if detail is None or gkey is None:
                warnings.append(f"Group sheet '{ls.name}' skipped: '{ls.group_by_header}' not found.")
                continue
            vals = [c.key for c in detail.columns for h in ls.sum_headers
                    if c.header.strip().lower() == h.strip().lower() and _summable(c)]
            gs = group_sheet(detail, gkey, vals or [c.key for c in detail.columns if _summable(c)], ls.name, vi,
                             used_sheets, ls.title)
            gs.notes = ls.notes[:5]
            built[ls.name.strip().lower()] = gs
            sheets.append(gs)
            continue
        sheets.append(SheetSpec(name=sheet_name(ls.name, used_sheets), title=ls.title, subtitle=ls.subtitle,
                                columns=cols, rows=rows, totals=list(dict.fromkeys(totals)),
                                blank_rows=0 if rows else max(ls.blank_rows, 10), notes=ls.notes[:5]))
        built[ls.name.strip().lower()] = sheets[-1]
    if not sheets:
        raise ValueError("layout produced no sheets")
    spec = WorkbookSpec(title=layout.title, sheets=sheets[:config.MAX_SHEETS], style_preset=layout.style_preset,
                        currency=answers.get("currency") or b.currency)
    return spec, warnings


def ground_check(spec: WorkbookSpec, b: sources.SourceBundle) -> list[str]:
    """Final guard: every numeric data cell must equal a known fact. Ungrounded values are blanked."""
    allowed = b.allowed_numbers()
    removed = 0
    for s in spec.sheets:
        for row in s.rows:
            for k, v in list(row.items()):
                if isinstance(v, (int, float)) and not isinstance(v, bool) and round(float(v), 9) not in allowed:
                    row[k] = None
                    removed += 1
    return [f"{removed} value(s) without a source were removed."] if removed else []


def build_spec(b: sources.SourceBundle, prompt: str, answers: dict) -> tuple[WorkbookSpec, list[str], str]:
    """Returns (spec, warnings, mode) where mode is 'azure:<deployment>' | 'mock' | 'mock_fallback'."""
    if answers.get("figures"):
        sources.answer_figures(b, str(answers["figures"]))
    if config.llm_mode() == "azure":
        try:
            layout, model = llm_layout(b, prompt, answers)
            spec, warns = from_layout(layout, b, prompt, answers)
            return spec, warns + ground_check(spec, b), model
        except (RuntimeError, ValueError, ValidationError) as e:
            spec, warns = mock_spec(b, prompt, answers)
            return spec, [f"AI layout unavailable ({str(e)[:120]}); used the deterministic layout."] + warns + \
                ground_check(spec, b), "mock_fallback"
    spec, warns = mock_spec(b, prompt, answers)
    return spec, warns + ground_check(spec, b), "mock"
