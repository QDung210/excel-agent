"""WorkbookSpec -> styled .xlsx via XlsxWriter (new files only; existing workbooks are edited by ooxml_edit).

Layout per sheet (0-based rows): [title] [subtitle] [blank] header, data rows, [totals], [blank, notes].
Formulas get a cached value from formula_eval when it can compute one, otherwise an empty cache; XlsxWriter sets
fullCalcOnLoad so Excel recalculates everything on open.
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from . import formula_eval
from .schemas import ColumnSpec, SheetSpec, WorkbookSpec

KEY_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]{0,40})\}")
REF_RE = re.compile(r"(?<![A-Za-z0-9_$!'\"])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
CURRENCY_FMT = {"VND": '#,##0 "₫"', "USD": '"$"#,##0.00', "EUR": '#,##0.00 "€"'}
PRESETS = {
    "clean_finance": {"header": {"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1,
                                 "border_color": "#BFBFBF", "valign": "vcenter", "text_wrap": True},
                      "cell": {"border": 1, "border_color": "#D9D9D9"}, "title_color": "#1F4E78",
                      "total": {"bold": True, "top": 2, "bg_color": "#DDEBF7", "border": 1, "border_color": "#BFBFBF"}},
    "minimal": {"header": {"bold": True, "bottom": 2, "valign": "vcenter", "text_wrap": True},
                "cell": {}, "title_color": "#000000", "total": {"bold": True, "top": 1}},
    "bold_header": {"header": {"bold": True, "font_color": "#FFFFFF", "bg_color": "#C65911", "border": 1,
                               "valign": "vcenter", "text_wrap": True, "font_size": 12},
                    "cell": {"border": 1, "border_color": "#D9D9D9"}, "title_color": "#C65911",
                    "total": {"bold": True, "top": 2, "bg_color": "#FCE4D6", "border": 1}},
}
INPUT_FILL = "#FFF9E6"


def num_format(col: ColumnSpec, currency: str | None) -> str | None:
    return {"currency": CURRENCY_FMT.get(currency or "", "#,##0.00"), "integer": "#,##0", "number": "#,##0.00",
            "percent": "0.0%", "date": "dd/mm/yyyy"}.get(col.type)


def to_date(v: Any) -> dt.datetime | None:
    if isinstance(v, (dt.date, dt.datetime)):
        return v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(v.strip(), fmt)
            except ValueError:
                pass
    return None


def shift_refs(formula: str, shift: int) -> str:
    """Moves relative A1 refs down by `shift` rows (strings untouched)."""
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    for i in range(0, len(parts), 2):
        parts[i] = REF_RE.sub(lambda m: f"{m[1]}{m[2]}{m[3]}{int(m[4]) + (0 if m[3] else shift)}", parts[i])
    return "".join(parts)


class _Layout:
    def __init__(self, s: SheetSpec):
        r = 0
        self.title = r if s.title else None
        r += 1 if s.title else 0
        self.subtitle = r if s.subtitle else None
        r += 1 if s.subtitle else 0
        r += 1 if (s.title or s.subtitle) else 0
        self.header = r
        self.first = r + 1
        self.n = len(s.rows) + s.blank_rows
        self.last = self.first + max(self.n, 1) - 1
        self.totals = self.first + self.n if s.totals else None
        self.notes = (self.totals if self.totals is not None else self.first + self.n - 1) + 2


def row_formula(template: str, keys: dict[str, int], excel_row: int) -> str:
    def sub(m: re.Match) -> str:
        if m[1] not in keys:
            raise ValueError(f"formula references unknown column {{{m[1]}}}")
        return f"{xl_col_to_name(keys[m[1]])}{excel_row}"
    f = KEY_RE.sub(sub, template.strip())
    return f if f.startswith("=") else "=" + f


def plan_cells(spec: WorkbookSpec) -> dict[str, dict[tuple[int, int], Any]]:
    """{sheet: {(r0, c0): value | '=formula'}} for data/formula/total cells (0-based) - shared with the evaluator."""
    out: dict[str, dict[tuple[int, int], Any]] = {}
    for s in spec.sheets:
        L, keys = _Layout(s), {c.key: i for i, c in enumerate(s.columns)}
        grid: dict[tuple[int, int], Any] = {}
        for i in range(L.n):
            r = L.first + i
            row = s.rows[i] if i < len(s.rows) else {}
            for c, col in enumerate(s.columns):
                if col.formula:
                    grid[(r, c)] = row_formula(col.formula, keys, r + 1)
                elif row.get(col.key) not in (None, ""):
                    v = row[col.key]
                    if col.type == "date" and (d := to_date(v)):
                        v = (d - dt.datetime(1899, 12, 30)).days
                    grid[(r, c)] = v
        if L.totals is not None:
            for k in s.totals:
                c = keys[k]
                rng = f"{xl_col_to_name(c)}{L.first + 1}:{xl_col_to_name(c)}{L.first + max(L.n, 1)}"
                grid[(L.totals, c)] = f"=SUM({rng})"
        for ref, f in s.formulas.items():
            m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", ref.upper())
            col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m[1]))) - 1
            grid[(int(m[2]) - 1 + L.header, col)] = shift_refs(f if f.startswith("=") else "=" + f, L.header)
        out[s.name] = grid
    return out


def compute(spec: WorkbookSpec, cells: dict[str, dict[tuple[int, int], Any]]) -> formula_eval.Workbook:
    plain = {sh: {(r + 1, c + 1): v for (r, c), v in g.items() if not (isinstance(v, str) and v.startswith("="))}
             for sh, g in cells.items()}
    forms = {sh: {(r + 1, c + 1): v for (r, c), v in g.items() if isinstance(v, str) and v.startswith("=")}
             for sh, g in cells.items()}
    return formula_eval.Workbook(plain, forms)


def build(spec: WorkbookSpec, path: str | Path, lang_vi: bool = False) -> dict:
    """Writes the workbook; returns stats {formulas, computed, uncomputed, cells}."""
    preset = PRESETS[spec.style_preset]
    cells = plan_cells(spec)
    ev = compute(spec, cells)
    stats = {"formulas": 0, "computed": 0, "uncomputed": 0, "cells": 0}
    wb = xlsxwriter.Workbook(str(path), {"strings_to_numbers": False, "strings_to_formulas": False,
                                         "strings_to_urls": False})
    wb.set_properties({"title": spec.title[:200], "comments": "Generated by AI Excel Workspace"})
    fmt_cache: dict[tuple, Any] = {}

    def fmt(**kw: Any) -> Any:
        key = tuple(sorted(kw.items()))
        if key not in fmt_cache:
            fmt_cache[key] = wb.add_format(kw)
        return fmt_cache[key]

    for s in spec.sheets:
        ws = wb.add_worksheet(s.name)
        L, grid = _Layout(s), cells[s.name]
        ncol = len(s.columns)
        if L.title is not None:
            ws.write_string(L.title, 0, s.title, fmt(bold=True, font_size=15, font_color=preset["title_color"]))
            ws.set_row(L.title, 22)
        if L.subtitle is not None:
            ws.write_string(L.subtitle, 0, s.subtitle, fmt(italic=True, font_color="#595959"))
        for c, col in enumerate(s.columns):
            ws.write_string(L.header, c, col.header, fmt(**preset["header"]))
        ws.set_row(L.header, 30)
        widths = [max(8, min(50, len(col.header) + 2)) for col in s.columns]
        for i in range(max(L.n, 0)):
            r = L.first + i
            is_blank_input = i >= len(s.rows)
            for c, col in enumerate(s.columns):
                nf = num_format(col, spec.currency)
                base = dict(preset["cell"], **({"num_format": nf} if nf else {}))
                v = grid.get((r, c))
                if isinstance(v, str) and v.startswith("="):
                    _write_formula(ws, r, c, v, ev, s.name, fmt(**base), stats)
                elif v is None:
                    style = dict(base, bg_color=INPUT_FILL) if is_blank_input else base
                    ws.write_blank(r, c, None, fmt(**style))
                elif col.type == "date" and isinstance(v, (int, float)):
                    ws.write_number(r, c, v, fmt(**base))
                    stats["cells"] += 1
                elif isinstance(v, bool):
                    ws.write_boolean(r, c, v, fmt(**base))
                    stats["cells"] += 1
                elif isinstance(v, (int, float)):
                    ws.write_number(r, c, v, fmt(**base))
                    widths[c] = max(widths[c], min(50, len(f"{v:,.2f}") + 3))
                    stats["cells"] += 1
                else:
                    ws.write_string(r, c, str(v), fmt(**dict(base, text_wrap=len(str(v)) > 50)))
                    widths[c] = max(widths[c], min(50, len(str(v)) + 2))
                    stats["cells"] += 1
        if L.totals is not None:
            label_col = next((i for i, col in enumerate(s.columns) if col.key not in s.totals), None)
            for c in range(ncol):
                col = s.columns[c]
                nf = num_format(col, spec.currency)
                style = dict(preset["total"], **({"num_format": nf} if nf else {}))
                v = grid.get((L.totals, c))
                if isinstance(v, str) and v.startswith("="):
                    _write_formula(ws, L.totals, c, v, ev, s.name, fmt(**style), stats)
                elif c == label_col:
                    ws.write_string(L.totals, c, "Tổng cộng" if lang_vi else "Total", fmt(**style))
                else:
                    ws.write_blank(L.totals, c, None, fmt(**style))
        body_rows = set(range(L.first, L.first + L.n)) | ({L.totals} if L.totals is not None else set())
        for (r, c), v in grid.items():                 # extra absolute formulas outside the table body
            if (r not in body_rows or c >= ncol) and isinstance(v, str) and v.startswith("="):
                _write_formula(ws, r, c, v, ev, s.name, None, stats)
        for i, note in enumerate(s.notes):
            ws.write_string(L.notes + i, 0, note, fmt(italic=True, font_color="#595959"))
        for c, col in enumerate(s.columns):
            ws.set_column(c, c, col.width or widths[c])
        if s.freeze_header:
            ws.freeze_panes(L.header + 1, 0)
        if L.n:
            ws.autofilter(L.header, 0, L.last, ncol - 1)
        ws.hide_gridlines(2 if spec.style_preset != "minimal" else 0)
    wb.close()
    return stats


def _write_formula(ws: Any, r: int, c: int, f: str, ev: formula_eval.Workbook, sheet: str, f_: Any,
                   stats: dict) -> None:
    v = ev.value(sheet, r + 1, c + 1)
    stats["formulas"] += 1
    if v is formula_eval.UNSUPPORTED or v is None:
        v = ""
        stats["uncomputed"] += 1
    else:
        stats["computed"] += 1
        if isinstance(v, float) and v.is_integer():
            v = int(v)
    ws.write_formula(r, c, f, f_, v)


def cell_name(r: int, c: int) -> str:
    return xl_rowcol_to_cell(r, c)
