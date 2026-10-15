"""Template mode, restructure: the AI reads prose sources and proposes the template table's column structure
(keep / rename / add / remove columns, row formulas, number formats), the user confirms, then the template copy is
restructured at the OOXML level and filled with values quoted from the sources.

The template's look is kept: new columns reuse an existing cell style of the same kind (input/formula) and number
format (styles.xml is never changed), header styles and widths are copied, formulas on every sheet, defined names,
charts, the autofilter and merges follow the moved columns. A column that is still used by a formula elsewhere is
hidden instead of deleted unless the user decides otherwise. Values never come from the model: every value must be
backed by a verbatim quote from a source (or the prompt) that contains that value; anything else stays blank and is
reported. Revision 1 stays the untouched template; the restructured copy is revision 2 (service.apply_transform).
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import time
import xml.etree.ElementTree as ET
import zipfile
from html import unescape
from pathlib import Path
from typing import Any, Callable, Literal

from openpyxl.utils import column_index_from_string, get_column_letter
from pydantic import BaseModel, Field

from . import builder, config, formula_eval, sources, spec_gen, templates
from .errors import InvalidOperation

# ---------------------------------------------------------------- A1 references in formulas
REF_RE = re.compile(
    r"(?<![\w.$'])(?P<sheet>(?:'(?:[^']|'')+'|[^\W\d][\w.]*)!)?"
    r"(?P<a>\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}(?=:\$?[A-Z]{1,3}(?![\w$(])))"
    r"(?::(?P<b>\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}))?(?![\w(!])")
PART_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d*)$")
STR_SPLIT = re.compile(r'("(?:[^"]|"")*")')
MAX_COL = 16384


def _unq(name: str) -> str:
    return name[1:-1].replace("''", "'") if name.startswith("'") else name


def _split(p: str) -> tuple[str, int, str, str]:
    m = PART_RE.match(p)
    return m[1], column_index_from_string(m[2]), m[3], m[4]


class ColMap:
    """old column -> new column (None = deleted) for one sheet; columns right of the table shift by `delta`."""

    def __init__(self, inside: dict[int, int | None], start: int, end: int, delta: int, top: int | None = None):
        self.inside, self.start, self.end, self.delta, self.top = inside, start, end, delta, top

    def __call__(self, c: int, r: int | None = None) -> int | None:
        if r is not None and self.top and r < self.top:  # rows above the table header (title, fields) stay put
            return c
        if self.start <= c <= self.end:
            return self.inside.get(c)
        return c + self.delta if c > self.end else c


def _map_ref(a: str, b: str | None, cm: ColMap) -> str | None:
    d1, c1, e1, r1 = _split(a)
    if b is None:
        n = cm(c1, int(r1) if r1 else None)
        return None if n is None or n > MAX_COL else f"{d1}{get_column_letter(n)}{e1}{r1}"
    d2, c2, e2, r2 = _split(b)
    lo, hi = min(c1, c2), max(c1, c2)
    row = max(int(r1), int(r2)) if r1 and r2 else None
    news = [n for c in range(lo, hi + 1) if (n := cm(c, row)) is not None and n <= MAX_COL]
    if not news:
        return None
    return f"{d1}{get_column_letter(min(news))}{e1}{r1}:{d2}{get_column_letter(max(news))}{e2}{r2}"


def remap_formula(f: str, own: str | None, target: str, cm: ColMap) -> str:
    """Rewrites references into sheet `target` (own = the sheet the formula lives on; None for workbook/chart
    formulas, where only qualified references count). Deleted cells become #REF!, ranges shrink like Excel."""
    def sub(m: re.Match) -> str:
        sh = m["sheet"]
        if (_unq(sh[:-1]) if sh else own) != target:
            return m[0]
        new = _map_ref(m["a"], m["b"], cm)
        return (sh or "") + (new or "#REF!")
    parts = STR_SPLIT.split(f)
    for i in range(0, len(parts), 2):
        parts[i] = REF_RE.sub(sub, parts[i])
    return "".join(parts)


def refs_to(f: str, own: str | None, target: str) -> set[int]:
    """Columns of sheet `target` that formula `f` refers to."""
    out: set[int] = set()
    parts = STR_SPLIT.split(f)
    for i in range(0, len(parts), 2):
        for m in REF_RE.finditer(parts[i]):
            sh = m["sheet"]
            if (_unq(sh[:-1]) if sh else own) != target:
                continue
            c1 = _split(m["a"])[1]
            c2 = _split(m["b"])[1] if m["b"] else c1
            if abs(c2 - c1) < 1000:
                out |= set(range(min(c1, c2), max(c1, c2) + 1))
    return out


def _x(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _xa(s: str) -> str:
    return _x(s).replace('"', "&quot;")


def _qname(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


# ---------------------------------------------------------------- workbook inspection
ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
ATTR_RE = re.compile(r'([\w:]+)="([^"]*)"')
CELLREF_RE = re.compile(r"([A-Z]{1,3})(\d+)$")


def _attrs(s: str) -> dict[str, str]:
    return dict(ATTR_RE.findall(s))


def _attr_str(a: dict[str, str]) -> str:
    return "".join(f' {k}="{v}"' for k, v in a.items())


def _sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    from xlsx_template import ooxml
    return ooxml.sheet_parts(z)


def _sst(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    return ["".join(unescape(t) for t in re.findall(r"<t\b[^>]*>(.*?)</t>", si, re.S))
            for si in re.findall(r"<si>(.*?)</si>", x, re.S)]


def _cell_text(a: dict, inner: str | None, sst: list[str]) -> str | None:
    inner = inner or ""
    if a.get("t") == "s":
        m = re.search(r"<v>(\d+)</v>", inner)
        return sst[int(m[1])] if m and int(m[1]) < len(sst) else None
    if a.get("t") == "inlineStr":
        return "".join(unescape(t) for t in re.findall(r"<t\b[^>]*>(.*?)</t>", inner, re.S))
    m = re.search(r"<v>(.*?)</v>", inner, re.S)
    return unescape(m[1]) if m and m[1] != "" else None


def _formula(inner: str | None) -> str | None:
    m = re.search(r"<f\b[^>]*>(.*?)</f>", inner or "", re.S)
    return unescape(m[1]) if m else None


def _cat(num_fmt: int, code: str) -> str:
    c = re.sub(r'"[^"]*"|\[[^\]]*\]', "", code.lower())
    if num_fmt in (9, 10) or "%" in c:
        return "percent"
    if 14 <= num_fmt <= 22 or 45 <= num_fmt <= 47 or (code and re.search(r"[dy]", c)) or re.search(r"m{1,2}/", c):
        return "date"
    if num_fmt in (0, 49) or c in ("general", "@"):
        return "text"
    return "number"


def _xf_cats(z: zipfile.ZipFile) -> list[str]:
    st = z.read("xl/styles.xml").decode("utf-8")
    fmts = {}
    for tag in re.findall(r"<numFmt\b[^>]*/?>", st):
        a = _attrs(tag)
        fmts[int(a.get("numFmtId", 0))] = unescape(a.get("formatCode", ""))
    m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", st, re.S)
    out = []
    for tag in re.findall(r"<xf\b[^>]*>", m[1] if m else ""):
        nf = int(_attrs(tag).get("numFmtId", 0))
        out.append(_cat(nf, fmts.get(nf, "")))
    return out or ["text"]


def inspect(path: str | Path, sheet: str, header_row: int, cols: list[int], first_row: int,
            last_row: int | None = None) -> dict:
    """Template table facts the planner and the engine need: per column header, style, formula of the first data
    row (with same-row refs written as {Header}) and which other formulas use the column."""
    with zipfile.ZipFile(path) as z:
        parts = _sheet_parts(z)
        sst = _sst(z)
        x = z.read(parts[sheet]).decode("utf-8")
        cells: dict[tuple[int, int], tuple[dict, str | None]] = {}
        for m in CELL_RE.finditer(x):
            a = _attrs(m[1])
            if (r := CELLREF_RE.match(a.get("r", ""))):
                cells[(int(r[2]), column_index_from_string(r[1]))] = (a, m[2])
        info = {}
        heads = {c: (_cell_text(*cells[(header_row, c)], sst) if (header_row, c) in cells else "") or "" for c in cols}
        for c in cols:
            a, inner = cells.get((first_row, c), ({}, None))
            f = _formula(inner)
            info[c] = {"col": c, "letter": get_column_letter(c), "header": heads[c].strip(),
                       "style": int(a.get("s", 0)), "head_style": int(cells.get((header_row, c), ({}, None))[0]
                                                                     .get("s", 0)),
                       "formula": f, "formula_h": _with_headers(f, first_row, heads) if f else None,
                       "used_by": []}
        # who uses each column: other sheets, defined names, charts, and this sheet's formulas outside the column
        own_refs: dict[int, set[str]] = {}
        for (r, c), (a, inner) in cells.items():
            f = _formula(inner)
            if f:
                for rc in refs_to(f, sheet, sheet):
                    if rc != c and rc in info:
                        own_refs.setdefault(rc, set()).add(heads.get(c) or get_column_letter(c))
        for c, who in own_refs.items():
            info[c]["used_by"] += [f"cột '{w}'" for w in sorted(who)]
        for name, part in parts.items():
            if name == sheet:
                continue
            fx = z.read(part).decode("utf-8")
            hit = set()
            for fm in re.findall(r"<f\b[^>]*>(.*?)</f>", fx, re.S):
                hit |= refs_to(unescape(fm), name, sheet)
            for c in hit & set(info):
                info[c]["used_by"].append(f"sheet '{name}'")
        wb = z.read("xl/workbook.xml").decode("utf-8")
        for dn in re.finditer(r"<definedName\b([^>]*)>(.*?)</definedName>", wb, re.S):
            if _attrs(dn[1]).get("name", "").startswith("_xlnm."):
                continue
            for c in refs_to(unescape(dn[2]), None, sheet) & set(info):
                info[c]["used_by"].append(f"tên '{_attrs(dn[1]).get('name')}'")
        for n in z.namelist():
            if n.startswith("xl/charts/chart") and n.endswith(".xml"):
                for fm in re.findall(r"<c:f>(.*?)</c:f>", z.read(n).decode("utf-8"), re.S):
                    for c in refs_to(unescape(fm), None, sheet) & set(info):
                        info[c]["used_by"].append("biểu đồ")
        for c in info:
            info[c]["used_by"] = list(dict.fromkeys(info[c]["used_by"]))
        has_table = bool(re.search(r"<tablePart\b", x))
        shared = bool(re.search(r'<f\b[^>]*t="shared"', x))
        last = last_row or first_row
        # the totals row: an aggregate over the data rows within a few rows below the table
        totals = None
        for r in range(last + 1, last + 5):
            for c in cols:
                f = _formula(cells.get((r, c), ({}, None))[1])
                if f and re.match(r"\s*(SUM|SUBTOTAL)\(", f, re.I) and c in refs_to(f, sheet, sheet):
                    totals = {"row": r, "style": int(cells[(r, c)][0].get("s", 0))}
                    break
            if totals:
                break
        # title / label texts outside the table body (what a report may need to retitle)
        texts = []
        body = set(range(first_row, last + 1)) | {header_row}
        for (r, c), (a, inner) in sorted(cells.items()):
            if r in body and c in info or _formula(inner):
                continue
            t = _cell_text(a, inner, sst)
            if a.get("t") in ("s", "inlineStr", "str") and t and t.strip() and len(texts) < 40:
                texts.append({"cell": f"{get_column_letter(c)}{r}", "row": r, "col": c, "text": t.strip(),
                              "style": int(a.get("s", 0))})
    return {"cols": info, "has_table_part": has_table, "shared_formulas": shared, "totals": totals, "texts": texts,
            "sheet_names": list(parts)}


def _with_headers(f: str, row: int, heads: dict[int, str]) -> str:
    def sub(m: re.Match) -> str:
        if m["sheet"] or m["b"]:
            return m[0]
        d1, c, e1, r = _split(m["a"])
        return "{" + heads[c] + "}" if r == str(row) and not d1 and c in heads and heads[c] else m[0]
    parts = STR_SPLIT.split(f)
    for i in range(0, len(parts), 2):
        parts[i] = REF_RE.sub(sub, parts[i])
    return "".join(parts)


# ---------------------------------------------------------------- the plan (JSON, stored in the clarification)
# plan = {sheet, header_row, rows:[excel rows], start, end,
#         columns:[{src, from, header, action: keep|input|formula, formula (with {Header}), fmt, default,
#                   default_src, changes:[...], reason}],
#         dropped:[{src, header, used_by:[...], mode: hide|delete|keep}],
#         kv:[{sheet, cell, label, value, value_type, quote, locator}],
#         data:[{header: {value, value_type, quote, locator}}], warnings:[...], mode}
FMT_CATS = {"text": ["text"], "number": ["number", "text"], "percent_number": ["text", "number"],
            "percent": ["percent"], "date": ["date"]}
KEY_RE = re.compile(r"\{([^{}]{1,80})\}")


def _layout(plan: dict) -> tuple[list[tuple[str, Any]], ColMap]:
    """Final column order (('col', plan column index) | ('kept', old col)) and the old->new column map."""
    start, end = plan["start"], plan["end"]
    keep_old = {d["src"] for d in plan["dropped"] if d["mode"] in ("hide", "keep")}
    items: list[tuple[str, Any]] = [("col", i) for i in range(len(plan["columns"]))]
    srcs = [c["src"] for c in plan["columns"]]
    for h in sorted(keep_old):
        anchor = max((i for i, s in enumerate(srcs) if s is not None and s < h), default=None)
        pos = 0 if anchor is None else next(k for k, it in enumerate(items) if it == ("col", anchor)) + 1
        while pos < len(items) and items[pos][0] == "kept" and items[pos][1] < h:
            pos += 1
        items.insert(pos, ("kept", h))
    inside: dict[int, int | None] = {c: None for c in range(start, end + 1)}
    for k, (kind, v) in enumerate(items):
        old = v if kind == "kept" else plan["columns"][v]["src"]
        if old is not None:
            inside[old] = start + k
    return items, ColMap(inside, start, end, len(items) - (end - start + 1), plan["header_row"])


def _excel_serial(iso: str) -> float:
    d = dt.date.fromisoformat(iso[:10])
    return float((d - dt.date(1899, 12, 30)).days)


def _value_xml(ref: str, s: int, value: Any, vtype: str) -> str:
    if value is None or value == "":
        return f'<c r="{ref}" s="{s}"/>'
    if vtype == "date":
        return f'<c r="{ref}" s="{s}"><v>{_excel_serial(str(value)):g}</v></c>'
    if vtype == "number":
        return f'<c r="{ref}" s="{s}"><v>{repr(float(value)) if not float(value).is_integer() else int(value)}</v></c>'
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{_x(str(value))}</t></is></c>'


def _formula_xml(ref: str, s: int, f: str) -> str:
    return f'<c r="{ref}" s="{s}"><f>{_x(f.lstrip("="))}</f></c>'


def _unguard(body: str) -> str:
    """`IF(X="","",REST)` -> REST when that IF wraps the whole formula (the AI's own blank-row guard)."""
    m = re.match(r'\s*IF\(\s*[A-Z]+\d+\s*=\s*""\s*,\s*""\s*,', body, re.I)
    if not m or not body.rstrip().endswith(")"):
        return body
    inner, depth, quoted = body[m.end():body.rstrip().rfind(")")], 0, False
    for ch in inner:
        if ch == '"':
            quoted = not quoted
        elif not quoted:
            depth += {"(": 1, ")": -1}.get(ch, 0)
            if depth < 0 or ch == "," and depth == 0:
                return body
    return inner.strip() if depth == 0 and not quoted else body


def blank_guard(body: str, refs: list[str]) -> str:
    """Blank result when any same-row input is blank: a missing price never becomes a computed 0 (or a negative
    margin) that the totals would then include."""
    refs = list(dict.fromkeys(refs))
    if not refs:
        return body
    test = f'{refs[0]}=""' if len(refs) == 1 else "OR(" + ",".join(f'{r}=""' for r in refs) + ")"
    return f'IF({test},"",{_unguard(body)})'


def row_formula(template: str, letters: dict[str, str], row: int) -> str:
    refs = [f"{letters[k]}{row}" for k in KEY_RE.findall(template)]
    return blank_guard(KEY_RE.sub(lambda m: f"{letters[m[1]]}{row}", template.strip().lstrip("=")), refs)


def _pick_style(info: dict, formula: bool, fmt: str, near: int, cats: list[str]) -> int:
    def cat(s: int) -> str:
        return cats[s] if s < len(cats) else "text"
    pool = [(c, d["style"]) for c, d in info.items() if bool(d["formula"]) == formula]
    for want in FMT_CATS.get(fmt, ["text"]):
        hits = sorted((abs(c - near), s) for c, s in pool if cat(s) == want)
        if hits:
            return hits[0][1]
    for want in FMT_CATS.get(fmt, ["text"]):
        hits = sorted((abs(c - near), d["style"]) for c, d in info.items() if cat(d["style"]) == want)
        if hits:
            return hits[0][1]
    return info[min(info, key=lambda c: abs(c - near))]["style"]


def apply_plan(src: str | Path, out: str | Path, plan: dict) -> dict:
    """Writes the restructured + filled copy of `src` to `out`. Returns {changed:[ranges], parts:[...]}."""
    sheet, hr, rows = plan["sheet"], plan["header_row"], plan["rows"]
    items, cm = _layout(plan)
    start = plan["start"]
    with zipfile.ZipFile(src) as zin:
        parts = _sheet_parts(zin)
        cats = _xf_cats(zin)
        info = {int(k): v for k, v in plan["info"].items()}
        names = zin.namelist()
        files = {n: zin.read(n) for n in names}
        infos = {i.filename: i for i in zin.infolist()}
    target_part = parts[sheet]
    x = files[target_part].decode("utf-8")
    if re.search(r"<tablePart\b", x):
        raise InvalidOperation("This sheet contains an Excel table (ListObject); restructuring it is not supported.")
    if re.search(r'<f\b[^>]*t="shared"', x):
        raise InvalidOperation("This sheet uses shared formulas; restructuring it is not supported yet.")
    risks = _reorder_risks(files, parts, plan)
    if risks:
        raise InvalidOperation("Moving existing columns into a different order is blocked: multi-column ranges "
                               f"would no longer cover the same cells ({'; '.join(risks[:3])}).")
    final_col = {k: start + k for k in range(len(items))}
    letters: dict[str, str] = {}
    col_style: dict[int, int] = {}                      # plan column index -> data cell style
    for k, (kind, v) in enumerate(items):
        if kind != "col":
            continue
        c = plan["columns"][v]
        letters[c["header"]] = get_column_letter(final_col[k])
        near = c["src"] if c["src"] is not None else _near_old(plan, v)
        if c["src"] is not None and (c["action"] == "keep" or (c["action"] == "formula") == bool(info[c["src"]]["formula"])):
            s = info[c["src"]]["style"]
            if c["action"] != "keep" and (cats[s] if s < len(cats) else "text") not in FMT_CATS.get(c["fmt"], ["text"]):
                s = _pick_style(info, c["action"] == "formula", c["fmt"], near, cats)
            col_style[v] = s
        else:
            col_style[v] = _pick_style(info, c["action"] == "formula", c["fmt"], near, cats)
    rowset = set(rows)
    data = plan.get("data") or []
    kv_cells = {}
    for d in plan.get("kv") or []:
        if d.get("value") in (None, ""):
            continue
        r = CELLREF_RE.match(d["cell"])
        nc = cm(column_index_from_string(r[1]), int(r[2]))
        if nc is not None and d["sheet"] == sheet:
            kv_cells[(int(r[2]), nc)] = d
    label_cells = {}
    for lb in plan.get("labels") or []:
        nc = cm(lb["col"], lb["row"])
        if nc is not None:
            label_cells[(lb["row"], nc)] = lb
    tot = plan.get("totals")
    changed: list[str] = []

    def move_cell(a: dict, inner: str | None, nc: int, r: int) -> str:
        a = dict(a)
        a["r"] = f"{get_column_letter(nc)}{r}"
        if inner is None:
            return f"<c{_attr_str(a)}/>"
        fm = re.search(r"<f\b([^>]*)>(.*?)</f>|<f\b([^>]*)/>", inner, re.S)
        if fm:
            fa = _attrs(fm[1] or fm[3] or "")
            if "ref" in fa:
                fa["ref"] = _remap_sqref(fa["ref"], cm) or a["r"]
            body = f"<f{_attr_str(fa)}>{_x(remap_formula(unescape(fm[2]), sheet, sheet, cm))}</f>" if fm[2] is not None \
                else f"<f{_attr_str(fa)}/>"
            if a.get("t") in ("str", "e", "b", "n"):
                a.pop("t")
            return f"<c{_attr_str(a)}>{body}</c>"
        return f"<c{_attr_str(a)}>{inner}</c>"

    before, inner, after = _split_sheet_data(x)
    out_rows: dict[int, tuple[dict, dict[int, str]]] = {}
    for rm in ROW_RE.finditer(inner):
        ra = _attrs(rm[1])
        r = int(ra["r"])
        ra.pop("spans", None)
        cells: dict[int, str] = {}
        old_cells: dict[int, tuple[dict, str | None]] = {}
        for m in CELL_RE.finditer(rm[2] or ""):
            a = _attrs(m[1])
            cr = CELLREF_RE.match(a.get("r", ""))
            if not cr:
                continue
            c = column_index_from_string(cr[1])
            old_cells[c] = (a, m[2])
            nc = cm(c, r)
            if nc is not None:
                cells[nc] = move_cell(a, m[2], nc, r)
        out_rows[r] = (ra, cells, old_cells)
    for r in sorted(rowset | {hr} | {k[0] for k in kv_cells} | {k[0] for k in label_cells} |
                    ({tot["row"]} if tot else set())):
        out_rows.setdefault(r, ({"r": str(r)}, {}, {}))
    for r, (ra, cells, old_cells) in out_rows.items():
        if r < hr:
            continue
        for k, (kind, v) in enumerate(items):
            if kind != "col":
                continue
            c = plan["columns"][v]
            nc = final_col[k]
            ref = f"{get_column_letter(nc)}{r}"
            if r == hr:
                if c["src"] is None or c["header"] != info[c["src"]]["header"]:
                    s = info[c["src"]]["head_style"] if c["src"] is not None else info[_near_old(plan, v)]["head_style"]
                    cells[nc] = _value_xml(ref, s, c["header"], "string")
                    changed.append(ref)
                continue
            if tot and r == tot["row"] and c.get("total") and "<f" not in cells.get(nc, ""):
                cells[nc] = _formula_xml(ref, tot["style"], f"SUM({letters[c['header']]}{rows[0]}:"
                                                            f"{letters[c['header']]}{rows[-1]})")
                changed.append(ref)
                continue
            if r not in rowset:
                if c["src"] is None:                     # e.g. a totals row below the table: same look, empty
                    left = next((final_col[j] for j in range(k - 1, -1, -1) if final_col[j] in cells), None)
                    if left is not None:
                        s = _attrs(re.match(r"<c\b([^>]*?)/?>", cells[left])[1]).get("s")
                        if s is not None:
                            cells[nc] = f'<c r="{ref}" s="{s}"/>'
                continue
            i = rows.index(r)
            cell = (data[i].get(c["header"]) if i < len(data) else None) or {}
            s = col_style[v]
            if c["action"] == "formula":
                cells[nc] = _formula_xml(ref, s, row_formula(c["formula"], letters, r))
                changed.append(ref)
            elif c["action"] == "sequence":
                cells[nc] = _value_xml(ref, s, i + 1 if i < len(data) else None, "number")
                changed.append(ref)
            elif c["action"] == "input":
                val = cell.get("value")
                vt = cell.get("value_type", "string")
                if val in (None, "") and c.get("default") not in (None, ""):
                    val, vt = c["default"], c.get("default_type", "number")
                cells[nc] = _value_xml(ref, s, val, vt)
                if val not in (None, "") or (c["src"] is not None and info[c["src"]]["formula"]):
                    changed.append(ref)
            elif cell.get("value") not in (None, "") and not info[c["src"]]["formula"]:
                cells[nc] = _value_xml(ref, s, cell["value"], cell.get("value_type", "string"))
                changed.append(ref)
    for (r, nc), d in kv_cells.items():
        ra, cells, old_cells = out_rows[r]
        old_c = next((c for c in old_cells if cm(c) == nc), None)
        s = int(old_cells[old_c][0].get("s", 0)) if old_c is not None else 0
        ref = f"{get_column_letter(nc)}{r}"
        cells[nc] = _value_xml(ref, s, d["value"], d.get("value_type", "string"))
        changed.append(ref)
    for (r, nc), lb in label_cells.items():              # retitled labels keep their cell's look
        ref = f"{get_column_letter(nc)}{r}"
        out_rows[r][1][nc] = _value_xml(ref, lb["style"], lb["new"], "string")
        changed.append(ref)
    body = []
    maxc, maxr, minc, minr = 1, 1, MAX_COL, None
    for r in sorted(out_rows):
        ra, cells, _ = out_rows[r]
        if cells:
            maxc, minc = max(maxc, max(cells)), min(minc, min(cells))
            maxr, minr = max(maxr, r), r if minr is None else minr
        inner_r = "".join(cells[c] for c in sorted(cells))
        body.append(f"<row{_attr_str(ra)}>{inner_r}</row>" if inner_r else f"<row{_attr_str(ra)}/>")
    dim = f"{get_column_letter(min(minc, maxc))}{minr or 1}:{get_column_letter(maxc)}{maxr}"
    head = _rewrite_sheet_extras(before, sheet, cm, plan, items, final_col, info, dim)
    tail = _rewrite_sheet_extras(after, sheet, cm, plan, items, final_col, info, dim)
    files[target_part] = (head + "".join(body) + tail).encode("utf-8")
    # every other sheet / chart / defined name that refers to the target sheet
    for name, part in parts.items():
        if part == target_part:
            continue
        files[part] = _rewrite_other_sheet(files[part].decode("utf-8"), name, sheet, cm).encode("utf-8")
    for n in names:
        if n.startswith("xl/charts/") and n.endswith(".xml"):
            cx = files[n].decode("utf-8")
            nx = re.sub(r"(<c:f>)(.*?)(</c:f>)", lambda m: m[1] + _x(remap_formula(unescape(m[2]), None, sheet, cm))
                        + m[3], cx, flags=re.S)
            if nx != cx:
                files[n] = nx.encode("utf-8")
    files["xl/workbook.xml"] = _rewrite_workbook(files["xl/workbook.xml"].decode("utf-8"), sheet, cm, plan, items,
                                                 parts).encode("utf-8")
    if "xl/calcChain.xml" in files:                      # stale after moving formulas; Excel rebuilds it
        files.pop("xl/calcChain.xml")
        files["[Content_Types].xml"] = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>', "",
                                              files["[Content_Types].xml"].decode("utf-8")).encode("utf-8")
        rels = "xl/_rels/workbook.xml.rels"
        files[rels] = re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain.xml"[^>]*/>', "",
                             files[rels].decode("utf-8")).encode("utf-8")
    added = _add_sheets(files, plan, letters, info, cats, changed) if plan.get("sheets") else []
    tmp = str(out)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names + added:
            if n in files:
                zi = zipfile.ZipInfo(n, date_time=infos[n].date_time if n in infos else (1980, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                zo.writestr(zi, files[n])
    return {"changed": changed, "items": items}


def _near_old(plan: dict, v: int) -> int:
    """Nearest existing template column to plan column v (for style/width/header look of a new column)."""
    cols = plan["columns"]
    for j in list(range(v - 1, -1, -1)) + list(range(v + 1, len(cols))):
        if cols[j]["src"] is not None:
            return cols[j]["src"]
    return plan["start"]


def _reorder_risks(files: dict, parts: dict, plan: dict) -> list[str]:
    """When kept columns change order, a multi-column range over the table (merge, SUM(A9:C9), a name, a rule)
    cannot be remapped faithfully. Lists such ranges; the whole-table autofilter is rebuilt and is fine."""
    srcs = [c["src"] for c in plan["columns"] if c["src"] is not None]
    if srcs == sorted(srcs):
        return []
    sheet, start, end, top = plan["sheet"], plan["start"], plan["end"], plan["header_row"]
    out: list[str] = []

    def spans(a: str, b: str | None) -> bool:
        if not b:
            return False
        _, c1, _, r1 = _split(a)
        _, c2, _, r2 = _split(b)
        if r1 and r2 and max(int(r1), int(r2)) < top:
            return False
        lo, hi = min(c1, c2), max(c1, c2)
        return lo != hi and not (hi < start or lo > end)

    def in_formula(f: str, own: str | None) -> bool:
        ps = STR_SPLIT.split(f)
        for i in range(0, len(ps), 2):
            for m in REF_RE.finditer(ps[i]):
                sh = m["sheet"]
                if (_unq(sh[:-1]) if sh else own) == sheet and spans(m["a"], m["b"]):
                    return True
        return False
    for name, part in parts.items():
        x = files[part].decode("utf-8")
        if any(in_formula(unescape(f), name) for f in re.findall(r"<f\b[^>]*>(.*?)</f>", x, re.S)):
            out.append(f"formulas on sheet '{name}'")
        if name != sheet:
            continue
        for kind, refs in (("merged cells", re.findall(r'<mergeCell\b[^>]*ref="([^"]*)"', x)),
                           ("conditional formatting / data validation",
                            re.findall(r'<(?:conditionalFormatting|dataValidation)\b[^>]*sqref="([^"]*)"', x))):
            if any(spans(*(tok.partition(":")[::2])) for r in refs for tok in r.split()):
                out.append(kind)
    wb = files["xl/workbook.xml"].decode("utf-8")
    for dn in re.finditer(r"<definedName\b([^>]*)>(.*?)</definedName>", wb, re.S):
        if _attrs(dn[1]).get("name") != "_xlnm._FilterDatabase" and in_formula(unescape(dn[2]), None):
            out.append(f"name '{_attrs(dn[1]).get('name')}'")
    return list(dict.fromkeys(out))


SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
SHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def _add_sheets(files: dict, plan: dict, letters: dict[str, str], info: dict, cats: list[str],
                changed: list[str]) -> list[str]:
    """New worksheets for the plan's extra tables (title row 1, headers row 3, data from row 4, optional totals row).
    Styles are reused from the template table; no style is created. Returns the new part names."""
    main, rows = plan["sheet"], plan["rows"]
    wb = files["xl/workbook.xml"].decode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    rp = re.search(r'xmlns:(\w+)="http://schemas.openxmlformats.org/officeDocument/2006/relationships"', wb)
    rpre = rp[1] if rp else "r"
    nums = [int(n) for n in re.findall(r"xl/worksheets/sheet(\d+)\.xml", " ".join(files))]
    ids = [int(n) for n in re.findall(r'\bsheetId="(\d+)"', wb)]
    rids = set(re.findall(r'\bId="([^"]+)"', rels))
    head_s = info[min(info)]["head_style"]
    tot = plan.get("totals")
    ranges = {(main, h): f"{_qname(main)}!${l}${rows[0]}:${l}${rows[-1]}" for h, l in letters.items()}
    for sh in plan["sheets"]:
        n = len(sh["data"])
        for j, c in enumerate(sh["columns"]):
            l = get_column_letter(j + 1)
            ranges[(sh["name"], c["header"])] = f"{_qname(sh['name'])}!${l}$4:${l}${3 + max(n, 1)}"
    added = []
    for sh in plan["sheets"]:
        n = len(sh["data"])
        num = max(nums or [0]) + 1
        nums.append(num)
        sid = max(ids or [0]) + 1
        ids.append(sid)
        rid = next(f"rId{k}" for k in range(1, 10000) if f"rId{k}" not in rids)
        rids.add(rid)
        cols = sh["columns"]
        own = {c["header"]: get_column_letter(j + 1) for j, c in enumerate(cols)}

        def cellf(tmpl: str, r: int) -> str:
            def sub(m: re.Match) -> str:
                key = m[1]
                if "!" in key:
                    s_, _, h = key.rpartition("!")
                    return ranges[(s_, h)]
                return f"{own[key]}{r}"
            body = KEY_RE.sub(sub, tmpl.strip().lstrip("="))
            return blank_guard(body, [f"{own[k]}{r}" for k in KEY_RE.findall(tmpl) if "!" not in k])
        out_rows = [f'<row r="1"><c r="A1" s="{plan.get("title_style", 0)}" t="inlineStr"><is><t xml:space="preserve">'
                    f'{_x(sh["purpose"] or sh["name"])}</t></is></c></row>',
                    '<row r="3">' + "".join(_value_xml(f"{own[c['header']]}3", head_s, c["header"], "string")
                                            for c in cols) + "</row>"]
        for i, d in enumerate(sh["data"]):
            r = 4 + i
            cs = []
            for c in cols:
                ref = f"{own[c['header']]}{r}"
                s = _pick_style(info, c["action"] == "formula", c["fmt"], plan["start"], cats)
                if c["action"] == "formula":
                    cs.append(_formula_xml(ref, s, cellf(c["formula"], r)))
                elif c["action"] == "sequence":
                    cs.append(_value_xml(ref, s, i + 1, "number"))
                else:
                    v = d.get(c["header"]) or {}
                    cs.append(_value_xml(ref, s, v.get("value"), v.get("value_type", "string")))
            out_rows.append(f'<row r="{r}">' + "".join(cs) + "</row>")
        last = 3 + n
        if any(c["total"] for c in cols) and n:
            r = last + 1
            ts = tot["style"] if tot else head_s
            cs = []
            for j, c in enumerate(cols):
                ref = f"{own[c['header']]}{r}"
                if c["total"]:
                    cs.append(_formula_xml(ref, ts, f"SUM({own[c['header']]}4:{own[c['header']]}{last})"))
                elif j == 0:
                    cs.append(_value_xml(ref, ts, "Tổng cộng" if plan.get("vi") else "Total", "string"))
            out_rows.append(f'<row r="{r}">' + "".join(cs) + "</row>")
            last = r
        widths = "".join(f'<col min="{j + 1}" max="{j + 1}" width="{min(max(len(c["header"]) * 1.15 + 2, 12), 40):.2f}"'
                         f' customWidth="1"/>' for j, c in enumerate(cols))
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               f'<dimension ref="A1:{get_column_letter(len(cols))}{last}"/>'
               '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
               f'<cols>{widths}</cols><sheetData>{"".join(out_rows)}</sheetData>'
               '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
        part = f"xl/worksheets/sheet{num}.xml"
        files[part] = xml.encode("utf-8")
        added.append(part)
        wb = wb.replace("</sheets>", f'<sheet name="{_xa(sh["name"])}" sheetId="{sid}" {rpre}:id="{rid}"/></sheets>', 1)
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{SHEET_REL}" '
                                                f'Target="worksheets/sheet{num}.xml"/></Relationships>', 1)
        ct = ct.replace("</Types>", f'<Override PartName="/{part}" ContentType="{SHEET_CT}"/></Types>', 1)
        changed.append(f"{_qname(sh['name'])}!A1:{get_column_letter(len(cols))}{last}")
    files["xl/workbook.xml"] = wb.encode("utf-8")
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    files["[Content_Types].xml"] = ct.encode("utf-8")
    return added


def _split_sheet_data(xml: str) -> tuple[str, str, str]:
    from xlsx_template import ooxml
    return ooxml.split_sheet_data(xml)


def _remap_sqref(s: str, cm: ColMap) -> str | None:
    out = []
    for tok in s.split():
        a, _, b = tok.partition(":")
        try:
            n = _map_ref(a, b or None, cm)
        except (TypeError, ValueError, AttributeError):
            n = tok
        if n:
            out.append(n)
    return " ".join(out) or None


def _rewrite_sheet_extras(xml: str, sheet: str, cm: ColMap, plan: dict, items, final_col, info, dim) -> str:
    xml = re.sub(r'<dimension ref="[^"]*"\s*/>', f'<dimension ref="{dim}"/>', xml)
    # <cols>: widths/styles/hidden follow the columns; new columns copy the width of their look-alike
    m = re.search(r"<cols>(.*?)</cols>", xml, re.S)
    widths: dict[int, dict] = {}
    if m:
        for tag in re.findall(r"<col\b[^>]*/>", m[1]):
            a = _attrs(tag)
            for c in range(int(a["min"]), min(int(a["max"]), MAX_COL) + 1):
                nc = cm(c)
                if nc is not None:
                    widths[nc] = {k: v for k, v in a.items() if k not in ("min", "max")}
    if m or plan["columns"]:
        for k, (kind, v) in enumerate(items):
            nc = final_col[k]
            if kind == "kept":
                d = next(d for d in plan["dropped"] if d["src"] == v)
                if d["mode"] == "hide":
                    widths.setdefault(nc, {"width": "9.140625"})["hidden"] = "1"
                continue
            c = plan["columns"][v]
            if c["src"] is None:
                base = dict(widths.get(cm(_near_old(plan, v)) or 0) or {"width": "12.7109375", "customWidth": "1"})
                want = max(len(c["header"]) * 1.15 + 2, *(len(str((r.get(c["header"]) or {}).get("value") or ""))
                                                          * 1.1 + 2 for r in (plan.get("data") or [{}])))
                base["width"] = str(round(min(max(float(base.get("width", 9)), want), 45), 4))
                base["customWidth"] = "1"
                base.pop("hidden", None)
                widths[nc] = base
        runs: list[list] = []
        for c in sorted(widths):
            if runs and runs[-1][1] == c - 1 and runs[-1][2] == widths[c]:
                runs[-1][1] = c
            else:
                runs.append([c, c, widths[c]])
        cols_xml = "<cols>" + "".join(f'<col min="{a}" max="{b}"{_attr_str(w)}/>' for a, b, w in runs) + "</cols>"
        if m:
            xml = xml[:m.start()] + cols_xml + xml[m.end():]
        elif "<sheetData" in xml:
            xml = xml.replace("<sheetData", cols_xml + "<sheetData", 1)
    xml = re.sub(r'(<pane\b[^>]*\btopLeftCell=")([^"]*)(")',
                 lambda m: m[1] + (_remap_sqref(m[2], cm) or "A1") + m[3], xml)
    xml = re.sub(r'(<selection\b[^>]*?)\s+activeCell="[^"]*"', r"\1", xml)
    xml = re.sub(r'(<selection\b[^>]*?)\s+sqref="[^"]*"', r"\1", xml)
    # autofilter: when it covered the whole table, it covers the whole new table
    def af(mm: re.Match) -> str:
        a = _attrs(mm[1])
        ref = a.get("ref", "")
        ok = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
        new_end = plan["start"] + len(items) - 1
        if ok and column_index_from_string(ok[1]) == plan["start"] and column_index_from_string(ok[3]) == plan["end"]:
            new = f"{ok[1]}{ok[2]}:{get_column_letter(new_end)}{ok[4]}"
        else:
            new = _remap_sqref(ref, cm)
        if not new:
            return ""
        a["ref"] = new
        body = mm[2] or ""
        s0 = column_index_from_string(re.match(r"[A-Z]+", ref)[0]) if ref else 1
        n0 = column_index_from_string(re.match(r"[A-Z]+", new)[0])

        def fc(f: re.Match) -> str:
            nc = cm(s0 + int(f[2]))
            return "" if nc is None else f[1] + str(nc - n0) + f[3]
        body = re.sub(r'(<filterColumn\b[^>]*colId=")(\d+)("[^>]*?(?:/>|>.*?</filterColumn>))', fc, body, flags=re.S)
        return f"<autoFilter{_attr_str(a)}>{body}</autoFilter>" if body else f"<autoFilter{_attr_str(a)}/>"
    xml = re.sub(r"<autoFilter\b([^>]*?)(?:/>|>(.*?)</autoFilter>)", af, xml, flags=re.S)

    def tag_ref(mm: re.Match) -> str:
        new = _remap_sqref(unescape(mm[3]), cm)
        return "" if new is None else f"{mm[1]}{mm[2]}=\"{new}\"{mm[4]}"
    xml = re.sub(r'(<(?:mergeCell|hyperlink|ignoredError|protectedRange)\b[^>]*?\b)(ref|sqref)="([^"]*)"([^>]*/>)',
                 tag_ref, xml)
    if "<mergeCells" in xml:
        n = len(re.findall(r"<mergeCell\b", xml))
        xml = re.sub(r"<mergeCells\b[^>]*>\s*</mergeCells>", "", xml) if n == 0 else \
            re.sub(r'(<mergeCells\b[^>]*?count=")\d+', rf"\g<1>{n}", xml)

    def block(mm: re.Match) -> str:
        new = _remap_sqref(unescape(mm[2]), cm)
        if new is None:
            return ""
        rest = re.sub(r"(<(?:formula|formula1|formula2)>)(.*?)(</(?:formula|formula1|formula2)>)",
                      lambda f: f[1] + _x(remap_formula(unescape(f[2]), sheet, sheet, cm)) + f[3], mm[3], flags=re.S)
        return f'{mm[1]}sqref="{new}"{rest}'
    xml = re.sub(r'(<conditionalFormatting\b[^>]*?)sqref="([^"]*)"(.*?</conditionalFormatting>)', block, xml,
                 flags=re.S)
    xml = re.sub(r'(<dataValidation\b[^>]*?)sqref="([^"]*)"((?:[^>]*/>)|(?:.*?</dataValidation>))', block, xml,
                 flags=re.S)
    xml = re.sub(r"(<xm:sqref>)(.*?)(</xm:sqref>)", lambda f: f[1] + (_remap_sqref(f[2], cm) or "A1") + f[3], xml)
    xml = re.sub(r"(<xm:f>)(.*?)(</xm:f>)", lambda f: f[1] + _x(remap_formula(unescape(f[2]), sheet, sheet, cm)) + f[3],
                 xml, flags=re.S)
    return xml


def _rewrite_other_sheet(xml: str, own: str, target: str, cm: ColMap) -> str:
    def cell(m: re.Match) -> str:
        inner = m[2]
        if inner is None or "<f" not in inner:
            return m[0]
        new = re.sub(r"(<f\b[^>]*>)(.*?)(</f>)", lambda f: f[1] + _x(remap_formula(unescape(f[2]), own, target, cm))
                     + f[3], inner, flags=re.S)
        if new == inner:
            return m[0]
        a = _attrs(m[1])
        if a.get("t") in ("str", "e", "b", "n"):
            a.pop("t")
        new = re.sub(r"<v>.*?</v>|<v/>", "", new, flags=re.S)       # cached result is stale; recalculated on open
        return f"<c{_attr_str(a)}>{new}</c>"
    xml = CELL_RE.sub(cell, xml)
    return re.sub(r"(<(formula|formula1|formula2|xm:f)>)(.*?)(</\2>)",
                  lambda f: f[1] + _x(remap_formula(unescape(f[3]), own, target, cm)) + f[4], xml, flags=re.S)


def _rewrite_workbook(xml: str, sheet: str, cm: ColMap, plan: dict, items, parts: dict) -> str:
    idx = list(parts).index(sheet)

    def dn(m: re.Match) -> str:
        a = _attrs(m[1])
        if a.get("name") == "_xlnm._FilterDatabase" and a.get("localSheetId") == str(idx):
            ok = re.fullmatch(r"(.*!)\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", unescape(m[2]))
            if ok and column_index_from_string(ok[2]) == plan["start"] and column_index_from_string(ok[4]) == plan["end"]:
                end = get_column_letter(plan["start"] + len(items) - 1)
                return f"<definedName{m[1]}>{_x(ok[1])}${ok[2]}${ok[3]}:${end}${ok[5]}</definedName>"
        local = a.get("localSheetId")
        own = list(parts)[int(local)] if local is not None and local.isdigit() and int(local) < len(parts) else None
        return f"<definedName{m[1]}>{_x(remap_formula(unescape(m[2]), own, sheet, cm))}</definedName>"
    xml = re.sub(r"<definedName\b([^>]*)>(.*?)</definedName>", dn, xml, flags=re.S)
    if "<calcPr" in xml:
        xml = re.sub(r"<calcPr\b([^>]*?)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m[1])
                     + ' fullCalcOnLoad="1"/>', xml)
    else:
        xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    return xml


# ---------------------------------------------------------------- validation of the produced file
def _ref_errors(z: zipfile.ZipFile) -> int:
    n = 0
    for name in z.namelist():
        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
            n += sum(f.count("#REF!") for f in re.findall(r"<f\b[^>]*>(.*?)</f>", z.read(name).decode("utf-8"), re.S))
    n += z.read("xl/workbook.xml").decode("utf-8").count("#REF!")
    return n


def validate(src: str | Path, out: str | Path, plan: dict) -> dict:
    errors: list[str] = []
    warnings: list[str] = []
    checks = {"parts": 0, "xml_parsed": 0, "new_ref_errors": 0, "headers": 0}
    allowed = re.compile(r"xl/worksheets/sheet\d+\.xml$|xl/workbook\.xml$|xl/charts/chart\d+\.xml$|"
                         r"\[Content_Types\]\.xml$|xl/_rels/workbook\.xml\.rels$")
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out) as zo:
        bad = zo.testzip()
        if bad:
            errors.append(f"L0: corrupt zip member {bad}")
        before, after = set(zi.namelist()), set(zo.namelist())
        new_parts = after - before
        if plan.get("sheets"):                           # the planned extra worksheets, nothing else
            extra = {n for n in new_parts if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)}
            if len(extra) != len(plan["sheets"]):
                errors.append(f"L1: expected {len(plan['sheets'])} new worksheet(s), found {len(extra)}")
            new_parts -= extra
        if new_parts:
            errors.append(f"L1: unexpected new parts {sorted(new_parts)}")
        if before - after - {"xl/calcChain.xml"}:
            errors.append(f"L1: parts removed {sorted(before - after)}")
        for n in sorted(after):
            data = zo.read(n)
            checks["parts"] += 1
            if n.endswith((".xml", ".rels")):
                checks["xml_parsed"] += 1
                try:
                    ET.fromstring(data)
                except ET.ParseError as e:
                    errors.append(f"L0: {n} is not well-formed XML ({e})")
            if n in before and zi.read(n) != data and not allowed.match(n):
                errors.append(f"L1: part {n} changed but restructuring must not touch it")
        if zi.read("xl/styles.xml") != zo.read("xl/styles.xml"):
            errors.append("L1: styles changed")
        new_refs = _ref_errors(zo) - _ref_errors(zi)
        checks["new_ref_errors"] = max(new_refs, 0)
        if new_refs > 0:
            errors.append(f"L2: {new_refs} formula reference(s) became #REF!")
    if not errors:
        import warnings as _w
        from openpyxl import load_workbook
        try:
            with _w.catch_warnings():
                _w.simplefilter("ignore")
                wb = load_workbook(out)
            try:
                ws = wb[plan["sheet"]]
                items, _ = _layout(plan)
                for k, (kind, v) in enumerate(items):
                    if kind != "col":
                        continue
                    want = plan["columns"][v]["header"]
                    got = ws.cell(plan["header_row"], plan["start"] + k).value
                    checks["headers"] += 1
                    if str(got or "").strip() != want.strip():
                        errors.append(f"L3: header {get_column_letter(plan['start'] + k)}{plan['header_row']} is "
                                      f"'{got}', expected '{want}'")
                for sh in plan.get("sheets") or []:
                    if sh["name"] not in wb.sheetnames:
                        errors.append(f"L3: sheet '{sh['name']}' missing")
                        continue
                    for j, c in enumerate(sh["columns"]):
                        checks["headers"] += 1
                        if str(wb[sh["name"]].cell(3, j + 1).value or "").strip() != c["header"]:
                            errors.append(f"L3: header of '{sh['name']}' column {j + 1} is not '{c['header']}'")
            finally:
                wb.close()
        except Exception as e:                           # noqa: BLE001
            errors.append(f"L4: openpyxl could not open the result ({type(e).__name__}: {str(e)[:120]})")
    return {"pass": not errors, "errors": errors, "warnings": warnings,
            "level": "L4" if not errors else "failed", "checks": checks}


# ---------------------------------------------------------------- grounding (values only from quotes)
COMPOUND_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(tỷ|triệu|tr)\s*(\d+)\s*(triệu|nghìn|ngàn|k)\b", re.I)
THOUSAND_RE = re.compile(r"(?<![\w.,])(\d+(?:[.,]\d+)?)\s*(nghìn|ngàn)\b", re.I)
NEG_RE = re.compile(r"\b(không|no|without|none|zero)\b", re.I)
DATE_TOKEN_RE = re.compile(r"\b(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\b|\b(\d{4})-(\d{2})-(\d{2})\b")
MULT = {"tỷ": 1e9, "triệu": 1e6, "tr": 1e6, "nghìn": 1e3, "ngàn": 1e3, "k": 1e3}
# "3 triệu 200" (colloquial: the trailing number is in the next unit down, 3.200.000)
COLLOQ_RE = re.compile(r"(?<![\d.,])(\d+(?:[.,]\d+)?)\s*(tỷ|triệu|tr)\s+(\d{1,3})(?!\d|[.,]\d|\s*%|\s*(?:tỷ|triệu|tr|"
                       r"nghìn|ngàn|k|đ|đồng|vnd|chiếc|cái|lần|người)(?!\w))", re.I)
WORD_NUM = {"không": 0, "một": 1, "hai": 2, "ba": 3, "bốn": 4, "tư": 4, "năm": 5, "sáu": 6, "bảy": 7, "bẩy": 7,
            "tám": 8, "chín": 9, "mười": 10, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
            "seven": 7, "eight": 8, "nine": 9, "ten": 10}
WORD_RE = re.compile(r"(?<!\w)(" + "|".join(sorted(WORD_NUM, key=len, reverse=True)) + r")(?!\w)", re.I)


def _num(s: str) -> float:
    return float(s.replace(",", "."))


def number_candidates(quote: str) -> set[float]:
    out: set[float] = set()
    for tok in sources.NUM_TOKEN_RE.findall(quote):
        v, unit = sources.parse_number(tok)
        if v is not None:
            out.add(round(v, 9))
            if unit == "%":
                out.add(round(v * 100, 9))
    for m in COMPOUND_RE.finditer(quote):
        out.add(round(_num(m[1]) * MULT[m[2].lower()] + _num(m[3]) * MULT[m[4].lower()], 6))
    for m in THOUSAND_RE.finditer(quote):
        out.add(round(_num(m[1]) * 1e3, 6))
    for m in COLLOQ_RE.finditer(quote):
        unit = MULT[m[2].lower()]
        out.add(round(_num(m[1]) * unit + int(m[3]) * unit / 1000, 6))
    for m in WORD_RE.finditer(quote):                    # "một màn hình và ba chuột"
        out.add(float(WORD_NUM[m[1].lower()]))
    return out


def _norm_ws(s: str) -> str:
    return " ".join(str(s).split())


def find_quote(b: sources.SourceBundle, prompt: str, quote: str) -> str | None:
    """Locator of the verbatim quote in a source text or in the user's prompt, else None."""
    q = _norm_ws(quote).casefold()                       # verbatim up to whitespace and letter case
    if not q or len(q) > 400:
        return None
    for fid, text in b.texts.items():
        for n, line in enumerate(text.splitlines(), start=1):
            if q in _norm_ws(line).casefold():
                return f"{b.names.get(fid, fid)}:line:{n}"
    return "prompt" if q in _norm_ws(prompt).casefold() else None


ELLIPSIS_RE = re.compile(r"\s*(?:\.{3,}|…)\s*")


def _elided_quote(b: sources.SourceBundle, prompt: str, quote: str) -> str | None:
    """Locator of a quote shortened with an ellipsis ("MN27 ... bán 3.450.000đ") when every part is verbatim, in
    order, on one source line (or in the prompt)."""
    parts = [_norm_ws(x).casefold() for x in ELLIPSIS_RE.split(quote) if x.strip()]
    if len(parts) < 2 or len(quote) > 400:
        return None

    def in_order(text: str) -> bool:
        pos = 0
        for x in parts:
            pos = text.find(x, pos)
            if pos < 0:
                return False
            pos += len(x)
        return True
    for fid, text in b.texts.items():
        for n, line in enumerate(text.splitlines(), start=1):
            if in_order(_norm_ws(line).casefold()):
                return f"{b.names.get(fid, fid)}:line:{n}"
    return "prompt" if in_order(_norm_ws(prompt).casefold()) else None


def _verbatim_span(b: sources.SourceBundle, prompt: str, quote: str, value: str) -> tuple[str, str] | None:
    """The longest run of >= 3 words of a not-quite-verbatim quote that is verbatim in the sources and still
    contains `value` ("12 chiếc LT14 đã giao cho khách" -> "12 chiếc LT14"). Only the verified span is kept."""
    words = _norm_ws(quote).split()
    target = sources.parse_number(value)[0]
    if target is None or len(words) < 3:
        return None
    hit = lambda text: any(abs(target - c) <= 1e-9 * max(1.0, abs(c)) for c in number_candidates(text))
    # the word after the number (unit, "triệu", "USD", "chiếc") must stay: "750.000 USD" is not "750.000 đồng"
    need = [k for k, w in enumerate(words) if re.search(r"\d", w) and hit(" ".join(words[k:k + 2]))]
    for size in range(min(len(words), 30) - 1, 2, -1):
        for i in range(len(words) - size + 1):
            span = " ".join(words[i:i + size])
            if any(i <= k and min(k + 1, len(words) - 1) < i + size for k in need) and hit(span):
                loc = find_quote(b, prompt, span)
                if loc:
                    return span, loc
    return None


def ground(value: Any, quote: str, fmt: str, b: sources.SourceBundle, prompt: str) -> tuple | str:
    """-> (typed value, value_type, locator[, verified quote]) when `quote` is in the sources and contains `value`;
    else a reason. Numbers carry the verified quote, which may be a verbatim span of a paraphrased quote."""
    if value is None or str(value).strip() == "":
        return "empty"
    loc = find_quote(b, prompt, quote) or _elided_quote(b, prompt, quote)
    if not loc and fmt in ("number", "percent_number", "percent"):
        found = _verbatim_span(b, prompt, quote, str(value).strip())
        if found:
            quote, loc = found
    if not loc:
        return "quote not found in the sources"
    sval = str(value).strip()
    if fmt in ("number", "percent_number", "percent"):
        v, unit = sources.parse_number(sval)
        if v is None:
            v = next(iter(sorted(c for c in number_candidates(sval))), None)
        if v is None:
            return f"'{sval}' is not a number"
        if fmt == "percent" and unit != "%" and v > 1:
            v = v / 100
        if fmt == "percent_number" and unit == "%":
            v = v * 100
        cands = number_candidates(quote)
        ok = any(abs(v - c) <= 1e-9 * max(1.0, abs(c)) for c in cands) or (
            fmt == "percent" and any(abs(v * 100 - c) <= 1e-9 * max(1.0, abs(c)) for c in cands))
        if not ok and v == 0 and NEG_RE.search(quote):
            ok = True
        if not ok:
            return f"{sval} does not appear in the quote"
        v = round(v, 9)
        return (int(v) if float(v).is_integer() else v), "number", loc, quote
    if fmt == "date":
        d = builder.to_date(sval)
        if d is None and (m := DATE_TOKEN_RE.search(sval)):
            d = builder.to_date(m[0]) if not m[4] else builder.to_date(m[0])
        if d is None:
            return f"'{sval}' is not a date (dd/mm/yyyy)"
        for m in DATE_TOKEN_RE.finditer(quote):
            q = builder.to_date(m[0]) if m[4] else _dmy(m)
            if q and q.date() == d.date():
                return d.date().isoformat(), "date", loc
        return f"date {sval} is not in the quote"
    if _norm_ws(sval).casefold() in _norm_ws(quote).casefold():
        return sval, "string", loc
    return f"'{sval}' is not in the quote"


def _dmy(m: re.Match) -> dt.datetime | None:
    try:
        return dt.datetime(int(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None


# ---------------------------------------------------------------- LLM proposal
FMT = Literal["text", "number", "percent_number", "percent", "date"]


class LlmUnderstanding(BaseModel):
    template_purpose: str = Field(default="", max_length=300, description="what the template is for, e.g. an "
                                  "invoice for one customer")
    document_summary: str = Field(default="", max_length=500, description="what the documents contain")
    report_intent: str = Field(default="", max_length=300, description="the report the documents/request ask for")
    fit: Literal["fits", "partial", "mismatch"] = Field(default="fits", description="fits = the template already "
                                                        "matches the data; partial = usable after column/label changes; "
                                                        "mismatch = a different kind of document")
    fit_reason: str = Field(default="", max_length=300)
    entities: list[str] = Field(default_factory=list, max_length=12, description="record dimensions found, e.g. "
                                "'2 chi nhánh: Hà Nội, Hải Phòng'")
    derivable: list[str] = Field(default_factory=list, max_length=12, description="metrics computable by formulas "
                                 "from stated values")
    missing: list[str] = Field(default_factory=list, max_length=12, description="data the report needs that the "
                               "documents do not give")
    assumptions: list[str] = Field(default_factory=list, max_length=8)


class LlmColumn(BaseModel):
    source_header: str | None = Field(default=None, description="exact header of the EXISTING template column this "
                                      "column keeps or renames; null for a NEW column")
    header: str = Field(max_length=80, description="final header text")
    kind: Literal["keep", "input", "formula", "sequence"] = Field(
        description="keep = existing column unchanged (its input or formula stays); input = typed values (an "
                    "existing formula is removed); formula = computed with `formula`; sequence = row number 1..n")
    formula: str | None = Field(default=None, description="kind=formula only: row formula using FINAL headers in "
                                "braces, e.g. {Số lượng}*{Đơn giá}")
    format: FMT = Field(description="percent_number = a percent typed as a plain number (5 for 5%) that formulas "
                        "divide by 100; percent = fraction shown with %")
    total: bool = Field(default=False, description="sum this column in the template's totals row")
    default: str | None = Field(default=None, description="default for every row, ONLY if a source states one")
    default_quote: str | None = Field(default=None, description="verbatim source text that states the default")
    reason: str = Field(default="", max_length=160, description="why (user's language, <= 15 words)")


class LlmValue(BaseModel):
    header: str = Field(description="final column header")
    value: str = Field(description="the value as written in the source (dates dd/mm/yyyy)")
    quote: str = Field(description="short text copied EXACTLY from one source line that contains the value (for "
                       "basis=implied: the sentence it is implied by)")
    basis: Literal["explicit", "implied"] = Field(default="explicit", description="implied = 0 because a complete "
                                                  "list in the quote does not include this record")
    note: str | None = Field(default=None, max_length=160, description="basis=implied: why")


class LlmRow(BaseModel):
    cells: list[LlmValue] = Field(default_factory=list, max_length=40)


class LlmKv(BaseModel):
    label: str = Field(description="CURRENT label of a header field of the template, e.g. 'Năm'")
    value: str
    quote: str = Field(description="text copied EXACTLY from one source line that contains the value")


class LlmLabel(BaseModel):
    current: str = Field(description="exact current text of a cell listed in template_texts")
    new: str = Field(max_length=120, description="replacement text")
    reason: str = Field(default="", max_length=160)


class LlmSheetColumn(BaseModel):
    header: str = Field(max_length=80)
    kind: Literal["input", "formula", "sequence"]
    formula: str | None = Field(default=None, description="same-row {Header}; a whole data column of another table "
                                "as {Sheet!Header}, only inside SUM/SUMIF/SUMIFS/COUNTIF/COUNTIFS/AVERAGEIF(S)")
    format: FMT = "number"
    total: bool = False


class LlmSheet(BaseModel):
    name: str = Field(max_length=31)
    purpose: str = Field(default="", max_length=200)
    reason: str = Field(default="", max_length=200)
    columns: list[LlmSheetColumn] = Field(min_length=1, max_length=20)
    rows: list[LlmRow] = Field(default_factory=list, max_length=60)


class LlmStructure(BaseModel):
    understanding: LlmUnderstanding | None = None
    columns: list[LlmColumn] = Field(min_length=1, max_length=40)
    labels: list[LlmLabel] = Field(default_factory=list, max_length=12)
    kv: list[LlmKv] = Field(default_factory=list, max_length=20)
    rows: list[LlmRow] = Field(default_factory=list, max_length=60)
    sheets: list[LlmSheet] = Field(default_factory=list, max_length=4)
    notes: list[str] = Field(default_factory=list, max_length=6)


SYSTEM = """You turn business documents (often prose) into a report built on an existing Excel TEMPLATE. Work like an
accountant: understand first, then adapt the template, then map data.
1. understanding: what the template is for, what the documents contain and ask for, whether the template fits the
   data (fits / partial / mismatch), the record dimensions (entities), metrics derivable by formulas, data that is
   missing, assumptions. Instructions in the documents about the report (what to include, what NOT to add, units)
   are binding.
2. columns: the FINAL ordered column list of the template table, so that one row = one record at the documents'
   finest grain (e.g. branch x product). Keep useful existing columns (source_header = exact current header; rename
   via `header` when the meaning changes, e.g. a quantity that becomes the net quantity sold); add the dimension and
   measure columns the records need; existing columns you leave out are removed. Prefer formulas for anything
   computable (net = shipped - returned, amount = qty * price, discount amount, net revenue); keep a template formula
   that still means the right thing (kind keep). If the template already fits the data, keep its columns unchanged.
   Kinds: keep, input, formula ({Final header} refs; + - * / ( ) SUM ROUND IF IFERROR MIN MAX ABS; no cell
   addresses; constants only 0-12, 100, 1000 or numbers written in the documents), sequence (row number, e.g. STT).
   format percent_number when formulas divide that column by 100 (users type 5 for 5%); percent when it holds a
   fraction shown with % (formulas do NOT divide by 100). total=true for quantity/money columns to sum in the
   template's totals row.
   The documents override the template: a formula they state for a column is used exactly; a column they say users
   enter is kind input. A rate column you add must be used by the formulas that need it.
3. labels: change texts listed in template_texts (title, field labels) when their wording no longer fits the report
   (e.g. an invoice title for a monthly sales report; a 'Khách hàng' field when there is no single customer may become
   'Đơn vị'/'Kỳ báo cáo'). Do not relabel what still fits.
4. kv: values for the template's header fields (use the CURRENT label), each with a verbatim quote.
5. rows: one per record. Turn prose into records: "Hà Nội xuất 18 laptop LT14, 30 màn hình MN27; khách trả một màn
   hình" gives (Hà Nội, LT14, xuất 18, trả 0) and (Hà Nội, MN27, xuất 30, trả 1). Give ONLY input values, never
   values of formula/sequence columns, never computed numbers. Each value: `quote` = a short span (<= 12 words)
   copied character-for-character from ONE source line that contains the value; a value stated once for many
   records (a unit price, a branch discount rate, a branch name in a heading) is repeated in each row with that
   quote. The quote must itself contain the value (never "...", never a sentence that only refers to it, e.g. "giá
   áp dụng cho cả hai địa điểm"): quote "giá niêm yết 15,9 triệu" in every LT14 row. Convert units as the documents ask (15,9 triệu -> 15900000). basis=implied ONLY for value 0 when a complete
   list in the quote leaves the record out (returns listed for two products -> the others returned 0), with a note.
6. sheets (optional, max 4): extra worksheets the documents explicitly ask for (summary by branch, receivables,
   costs, inventory). Each is a small table of columns + rows with quoted input values; aggregate the main table with
   formulas such as SUMIF({MainSheet!Chi nhánh},{Chi nhánh},{MainSheet!Doanh thu ròng}) (MainSheet = template_sheet),
   never typed totals.
Never invent values, rates, taxes or defaults; when the documents say not to add something (e.g. VAT), do not add it.
Leave missing data blank and list it in understanding.missing. Headers, labels and reasons in the documents'
language. Return only the JSON object."""


def _context(table: dict, kv_labels: list[str], b: sources.SourceBundle, prompt: str, texts: list[dict] | None = None,
             totals: dict | None = None) -> str:
    cols = [{"header": d["header"], "kind": "formula" if d["formula"] else "input",
             "formula": d["formula_h"], "used_by_other_formulas": d["used_by"]} for d in table["cols"].values()]
    docs = []
    budget = 24000
    for fid, text in b.texts.items():
        chunk = text[:min(12000, budget)]
        budget -= len(chunk)
        if chunk:
            docs.append({"file": b.names.get(fid, fid), "text": chunk})
    return json.dumps({"request": prompt[:3000], "template_sheet": table["sheet"], "table_columns": cols,
                       "table_rows_available": table.get("rows"), "totals_row": bool(totals),
                       "header_fields": kv_labels, "template_texts": [t["text"] for t in texts or []][:40],
                       "documents": docs}, ensure_ascii=False)


def unused_inputs(res: LlmStructure, template_formulas: dict[str, str] | None = None) -> list[str]:
    """Numeric input columns that no formula reads although the table computes things (e.g. 'Số lượng trả' while
    revenue and cost use the shipped quantity): usually a formula that should use them."""
    formulas = [c.formula or "" for c in res.columns if c.kind == "formula"]
    formulas += [c.formula or "" for sh in res.sheets for c in sh.columns if c.kind == "formula"]
    formulas += [(template_formulas or {}).get(templates.norm(c.source_header or ""), "") for c in res.columns
                 if c.kind == "keep"]                    # a kept template formula, in {header} terms
    if not any(formulas):
        return []
    used = {templates.norm(k.split("!")[-1]) for f in formulas for k in KEY_RE.findall(f)}
    return [f"Column '{c.header}' is a numeric input but no formula uses it" for c in res.columns
            if c.kind == "input" and c.format in NUMERIC and templates.norm(c.header) not in used]


def review_context(context: str, res: LlmStructure, issues: list[str]) -> str:
    data = json.loads(context)
    data["review"] = {
        "your_previous_answer": res.model_dump(exclude_none=True),
        "problems": issues,
        "task": "Return the corrected full answer. For each problem, re-read the documents: if they say a figure "
                "must be computed from it (e.g. revenue and cost of goods on the quantity AFTER returns), fix the "
                "formulas (add a net column such as {shipped}-{returned} and use it). If the column is genuinely "
                "informational, keep it unchanged. Keep everything else as it was."}
    return json.dumps(data, ensure_ascii=False)


def _llm(context: str) -> tuple[LlmStructure, str]:
    errs: list[str] = []
    for name, llm in config.chat_models(max_tokens=20000, timeout=240):
        for attempt in (0, 1):
            try:
                out = llm.with_structured_output(LlmStructure, method="json_schema").invoke(
                    [("system", SYSTEM), ("user", context)])
                if isinstance(out, LlmStructure):
                    return out, name
                errs.append(f"{name}: empty structured output")
                break
            except Exception as e:                       # noqa: BLE001 - retry once on a per-minute limit, then next
                raw = str(e)
                errs.append(f"{name}: {type(e).__name__}: {raw[:300]}")
                if attempt or not re.search(r"rate.?limit|\b429\b", raw, re.I):
                    break
                wait = re.search(r"try again in ([\d.]+)s", raw)
                time.sleep(min(float(wait[1]) if wait else 20.0, 30.0) + 1)
    raise RuntimeError(" | ".join(errs) or "no language model configured")


def check_formula(template: str, keys: set[str], allowed: set[float]) -> str | None:
    body = template.strip().lstrip("=")
    if not body or len(body) > 400:
        return "empty or too long"
    for k in KEY_RE.findall(body):
        if k not in keys:
            return f"unknown column {{{k}}}"
    bare = KEY_RE.sub("", body)
    if re.search(r"\$?[A-Za-z]{1,3}\$?\d+|!|\[|'", bare):
        return "direct cell/sheet references are not allowed"
    funcs = formula_eval.functions_used("=" + KEY_RE.sub("A1", body))
    if "?" in funcs or not funcs <= formula_eval.SUPPORTED_FUNCTIONS:
        return f"unsupported function(s) {sorted(funcs - formula_eval.SUPPORTED_FUNCTIONS)}"
    for lit in re.findall(r"(?<![A-Za-z_\d.])\d+(?:\.\d+)?", bare):
        if float(lit) not in spec_gen.SAFE_LITERALS and round(float(lit), 9) not in allowed:
            return f"ungrounded constant {lit}"
    return None


LOOKUP_RE = re.compile(r"\b(VLOOKUP|HLOOKUP|XLOOKUP|LOOKUP|INDEX|MATCH)\s*\(", re.I)


def fix_percent(formula: str, fmts: dict[str, str | None]) -> str:
    """Make percent scaling follow each column's format: a 'percent' column already holds a fraction (5% = 0.05),
    so a following /100 is dropped; a 'percent_number' column holds 5, so it is divided by 100 where it is used."""
    def one(m: re.Match) -> str:
        key, div = m.group(1), m.group(2)
        fmt = fmts.get(key)
        if fmt == "percent" and div:
            return "{" + key + "}"
        if fmt == "percent_number" and not div:
            return "({" + key + "}/100)"
        return m.group(0)
    # {Key} optionally followed by /100 (not a power, not a longer number); not when the key itself is a divisor
    return re.sub(r"(?<![/^])\{([^{}]+)\}(\s*/\s*100(?![\d.^]))?", one, formula)


def table_of(slots: templates.Slots) -> templates.TableSlot | None:
    return next((t for t in slots.tables if t.rows and t.sheet not in slots.hidden_sheets), None)


def _mark_dates(path: str | Path, sheet: str, row: int, info: dict[int, dict]) -> None:
    """Columns whose first data cell has a date number format: the AI's format guess is not trusted for them."""
    from openpyxl import load_workbook
    from openpyxl.styles.numbers import is_date_format
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb[sheet]
        for c, d in info.items():
            d["date"] = bool(is_date_format(ws.cell(row, c).number_format or "General"))
    finally:
        wb.close()


def propose(path: str | Path, slots: templates.Slots, b: sources.SourceBundle, prompt: str,
            llm: Callable[[str], tuple[LlmStructure, str]] | None = None) -> dict | None:
    """The AI's structure proposal, grounded and normalized into a plan, then dry-run on a copy of the template so
    the preview shows exactly what the file will contain. None when the template has no input table."""
    t = table_of(slots)
    if t is None:
        return None
    cols = [c.col for c in t.columns]
    ins = inspect(path, t.sheet, t.header_row, cols, t.rows[0], t.rows[-1])
    _mark_dates(path, t.sheet, t.rows[0], ins["cols"])
    table = {"sheet": t.sheet, "cols": ins["cols"], "rows": len(t.rows)}
    kv_labels = [k.label for k in slots.kv if k.sheet == t.sheet]
    ctx = _context(table, kv_labels, b, prompt, ins["texts"], ins["totals"])
    res, model = (llm or _llm)(ctx)
    kept_f = {templates.norm(d["header"]): d["formula_h"] or "" for d in ins["cols"].values()}
    issues = unused_inputs(res, kept_f)
    if issues:                                           # one review round: the model fixes its own formulas
        try:
            fixed, m2 = (llm or _llm)(review_context(ctx, res, issues))
            if len(unused_inputs(fixed, kept_f)) < len(issues):
                res, model = fixed, m2
        except Exception:                                # noqa: BLE001 - keep the first answer (rate limit etc.)
            pass
    plan = normalize(res, t, ins, slots, b, prompt, model if ":" in model else f"azure:{model}")
    plan["preview"] = dry_run(path, plan)
    if not plan["preview"]["ok"] and plan["preview"].get("code") == "reorder":
        plan["warnings"].append(plan["preview"]["errors"][0] + " The columns keep the template's order instead.")
        template_order(plan)
        plan["preview"] = dry_run(path, plan)
    return plan


def template_order(plan: dict) -> None:
    """Existing columns back in their template order (new columns stay after their predecessor)."""
    cols = plan["columns"]
    kept = iter(sorted((c for c in cols if c["src"] is not None), key=lambda c: c["src"]))
    plan["columns"] = [next(kept) if c["src"] is not None else c for c in cols]
    for c in plan["columns"]:
        c["changes"] = [x for x in c["changes"] if x != "move"] or ["keep"]


VI_RE = re.compile(r"[ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ]", re.I)
NUMERIC = ("number", "percent_number", "percent")
SEQ_HEADER_RE = re.compile(r"^(?:stt|tt|so tt|so thu tu|no|nr|sn|s n|sr no)$")     # templates.norm() form
AGG_FNS = {"SUM", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "MINIFS", "MAXIFS",
           "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"}


def _ground_cell(cell: LlmValue, fmt: str, b: sources.SourceBundle, prompt: str) -> dict | str:
    """A model value -> {value, value_type, quote, locator, prov, note} or the reason it is left blank."""
    if cell.basis == "implied":
        loc = find_quote(b, prompt, cell.quote)
        if not loc:                                      # "một chiếc LT14 và hai bàn phím" -> verbatim part only
            words = _norm_ws(cell.quote).split()
            for size in range(min(len(words), 30) - 1, 2, -1):
                hits = [(span, find_quote(b, prompt, span)) for span in
                        (" ".join(words[i:i + size]) for i in range(len(words) - size + 1))]
                span, loc = next(((s_, l_) for s_, l_ in hits if l_), (None, None))
                if loc:
                    cell = cell.model_copy(update={"quote": span})
                    break
        num = sources.parse_number(str(cell.value).strip())[0]
        if fmt not in NUMERIC or num != 0:
            return "only 0 may be implied"
        if not loc:
            return "quote not found in the sources"
        return {"value": 0, "value_type": "number", "quote": _norm_ws(cell.quote)[:200], "locator": loc,
                "prov": "assumed", "note": (cell.note or "")[:160] or None}
    g = ground(cell.value, cell.quote, fmt, b, prompt)
    if isinstance(g, str):
        return g
    return {"value": g[0], "value_type": g[1], "quote": _norm_ws(g[3] if len(g) > 3 else cell.quote)[:200],
            "locator": g[2], "prov": "explicit", "note": None}


def _in_aggregate(body: str, pos: int) -> bool:
    """Is position `pos` of formula `body` inside an argument of an aggregate function?"""
    masked = KEY_RE.sub(lambda m: "_" * len(m[0]), body[:pos])
    stack: list[str] = []
    for m in re.finditer(r"([A-Za-z][A-Za-z0-9.]*)\s*\(|\(|\)", masked):
        if m[0] == ")":
            if stack:
                stack.pop()
        else:
            stack.append((m[1] or "").upper())
    return any(s in AGG_FNS for s in stack)


def _sheet_name(name: str, taken: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", " ", " ".join(name.split())).strip(" '")[:31] or "Sheet"
    out, i = base, 2
    while out.casefold() in taken:
        suffix = f" ({i})"
        out, i = base[:31 - len(suffix)] + suffix, i + 1
    taken.add(out.casefold())
    return out


def _normalize_sheets(res: LlmStructure, main: str, columns: list[dict], sheet_names: list[str],
                      b: sources.SourceBundle, prompt: str, allowed: set[float], warnings: list[str]) -> list[dict]:
    taken = {s.casefold() for s in sheet_names}
    tables: dict[str, set[str]] = {main: {c["header"] for c in columns}}
    alias = {templates.norm(main): main, "mainsheet": main, "main": main}
    out = []
    for ls in res.sheets[:4]:
        name = _sheet_name(ls.name, taken)
        alias[templates.norm(ls.name)] = alias[templates.norm(name)] = name
        cols, seen = [], set()
        for lc in ls.columns:
            h = " ".join(lc.header.split())[:80]
            if h and h not in seen:
                seen.add(h)
                cols.append({"header": h, "action": lc.kind, "formula": (lc.formula or "").strip().lstrip("=") or None,
                             "fmt": lc.format, "total": bool(lc.total)})
        own = {c["header"] for c in cols}
        for c in cols:
            if c["action"] != "formula":
                continue
            f = c["formula"] or ""

            def canon(m: re.Match) -> str:
                key = m[1]
                if "!" not in key:
                    return m[0]
                sh, _, h = key.rpartition("!")
                return "{" + alias.get(templates.norm(sh.strip("'")), sh) + "!" + h.strip() + "}"
            f = KEY_RE.sub(canon, f)
            cross = {f"{s}!{h}" for s, hs in tables.items() for h in hs}
            err = check_formula(f, (own - {c["header"]}) | cross, allowed)
            if not err:
                for m in KEY_RE.finditer(f):
                    if "!" in m[1] and not _in_aggregate(f, m.start()):
                        err = f"{{{m[1]}}} is used outside an aggregate function"
                        break
            if err:
                warnings.append(f"Sheet '{name}': formula for '{c['header']}' rejected ({err}); the column stays blank.")
                c["action"], c["formula"] = "input", None
            else:
                c["formula"] = f
        fmts = {c["header"]: c["fmt"] for c in cols}
        data = []
        for lr in ls.rows[:60]:
            row = {}
            for cell in lr.cells:
                h = next((x for x in own if templates.norm(x) == templates.norm(cell.header)), None)
                c = next((x for x in cols if x["header"] == h), None)
                if not c or c["action"] != "input":
                    continue
                g = _ground_cell(cell, fmts[h], b, prompt)
                if isinstance(g, dict):
                    row[h] = g
            data.append(row)
        has_input = any(c["action"] == "input" for c in cols)
        data = [r for r in data if r] if has_input else (data or [{}])
        if not data:
            warnings.append(f"Sheet '{name}' was not added: none of its values is backed by a source quote.")
            continue
        tables[name] = own
        out.append({"name": name, "purpose": ls.purpose[:200], "reason": ls.reason[:200], "columns": cols,
                    "data": data})
    return out


def normalize(res: LlmStructure, t: templates.TableSlot, ins: dict, slots: templates.Slots,
              b: sources.SourceBundle, prompt: str, mode: str) -> dict:
    info: dict[int, dict] = ins["cols"]
    warnings: list[str] = []
    locked = ins["has_table_part"] or ins["shared_formulas"]
    if locked:
        warnings.append("This template's table uses Excel features (table object / shared formulas) that cannot be "
                        "restructured yet; only values will be filled.")
    by_head = {templates.norm(d["header"]): c for c, d in info.items() if d["header"]}
    columns: list[dict] = []
    used_src: set[int] = set()
    seen: set[str] = set()
    for lc in res.columns:
        header = " ".join(lc.header.split())[:80]
        if not header or header in seen:
            continue
        src = by_head.get(templates.norm(lc.source_header or "")) if lc.source_header else None
        if src is not None and src in used_src:
            src = None
        if src is None and lc.kind == "keep":
            src = by_head.get(templates.norm(header))
            if src is None or src in used_src:
                warnings.append(f"Column '{header}' was proposed as kept but is not in the template; added as input.")
                src = None
        if src is not None:
            used_src.add(src)
        seen.add(header)
        action = {"keep": "keep" if src is not None else "input", "formula": "formula",
                  "sequence": "sequence"}.get(lc.kind, "input")
        columns.append({"src": src, "from": info[src]["header"] if src is not None else None, "header": header,
                        "action": action, "formula": (lc.formula or "").strip().lstrip("=") or None,
                        "fmt": "number" if action == "sequence" else lc.format,
                        "total": bool(lc.total) and action != "sequence",
                        "default": None, "default_type": None, "default_src": None,
                        "reason": (lc.reason or "")[:160], "_default": lc.default, "_dq": lc.default_quote})
    for c in columns:                                    # "STT"/"No." is a row number, never a value to look up
        seq_head = SEQ_HEADER_RE.match(templates.norm(c["header"])) or c["header"].strip() == "#"
        if c["action"] in ("input", "keep") and seq_head and not (c["src"] is not None and info[c["src"]]["formula"]):
            c.update(action="sequence", fmt="number", total=False, _auto_seq=c["action"] == "keep")
    if locked:                                           # values only, on the unchanged structure
        columns = [{"src": c, "from": d["header"], "header": d["header"], "action": "keep", "formula": None,
                    "fmt": "text", "total": False, "default": None, "default_type": None, "default_src": None,
                    "reason": "", "_default": None, "_dq": None} for c, d in info.items()]
    headers = {c["header"] for c in columns}
    allowed = sources.SourceBundle.allowed_numbers(b) | {round(v, 9) for v in number_candidates(prompt)}
    for c in columns:                                    # the template's date format wins over the AI's guess
        if c["src"] is not None and info[c["src"]].get("date"):
            c["fmt"] = "date"
    fmts = {c["header"]: c["fmt"] for c in columns}
    for c in columns:
        if c["action"] == "formula":
            err = check_formula(c["formula"] or "", headers - {c["header"]}, allowed | _all_numbers(b))
            if err:
                warnings.append(f"Formula for '{c['header']}' rejected ({err}).")
                c["action"] = "keep" if c["src"] is not None and info[c["src"]]["formula"] else "input"
                c["formula"] = None
            else:
                fixed = fix_percent(c["formula"], fmts)
                if fixed != c["formula"]:
                    warnings.append(f"Formula for '{c['header']}' adjusted so percent columns are scaled once: "
                                    f"{c['formula']} -> {fixed}")
                    c["formula"] = fixed
        if c["action"] == "keep" and c["src"] is not None:
            c["fmt"] = c["fmt"] if c["fmt"] else "text"
        c["total"] = c["total"] and c["fmt"] in NUMERIC
        if c["_default"] not in (None, "") and c["action"] == "input":
            g = ground(c["_default"], c["_dq"] or "", c["fmt"], b, prompt)
            if isinstance(g, tuple):
                c["default"], c["default_type"], c["default_src"] = g[0], g[1], g[2]
            else:
                warnings.append(f"Default for '{c['header']}' not used ({g}).")
    used = " ".join(c["formula"] or "" for c in columns if c["action"] == "formula")
    for c in columns:
        if c["src"] is None and c["action"] == "input" and c["fmt"] in NUMERIC and "{" + c["header"] + "}" not in used                 and any(x["action"] == "formula" for x in columns):
            warnings.append(f"New column '{c['header']}' is not used by any formula; check the formulas that "
                            "should use it.")
    # a template lookup (e.g. VLOOKUP into a catalog) never overrides values the documents state for that column
    for c in columns:
        if c["action"] != "keep" or c["src"] is None or not LOOKUP_RE.search(info[c["src"]]["formula"] or ""):
            continue
        stated = [x for r in res.rows for x in r.cells if templates.norm(x.header) == templates.norm(c["header"])]
        if any(isinstance(ground(x.value, x.quote, c["fmt"], b, prompt), tuple) for x in stated):
            c["action"] = "input"
            c["reason"] = c["reason"] or "tài liệu ghi giá trị cụ thể; không tra cứu danh mục"
            warnings.append(f"'{c['header']}' becomes an input column: the documents state its values, so the "
                            "template's lookup formula is not used for it.")
    src_order = [c["src"] for c in columns if c["src"] is not None]
    rank = {s: i for i, s in enumerate(sorted(src_order))}
    for i, s in enumerate(src_order):
        if rank[s] != i:
            next(c for c in columns if c["src"] == s)["_moved"] = True
    for c in columns:
        c.pop("_default", None)
        c.pop("_dq", None)
        changes = []
        if c["src"] is None:
            changes.append("add")
        else:
            if c["header"] != info[c["src"]]["header"]:
                changes.append("rename")
            was_f = bool(info[c["src"]]["formula"])
            if c["action"] in ("input", "sequence") and was_f:
                changes.append("to_input")
            if c["action"] == "formula":
                changes.append("formula" if was_f else "to_formula")
            if c.pop("_moved", False):
                changes.append("move")
        if c["action"] == "sequence" and not c.pop("_auto_seq", False):     # a kept STT is just numbered
            changes.append("sequence")
        if c["default"] is not None:
            changes.append("default")
        if c["total"]:
            changes.append("total")
        c["changes"] = changes or ["keep"]
    # dropped columns: hide when something still uses them
    dropped = []
    keep_srcs = {c["src"] for c in columns if c["src"] is not None}
    for col, d in info.items():
        if col in keep_srcs:
            continue
        users = [u for u in d["used_by"] if not _only_dropped(u, info, keep_srcs)]
        new_formulas_use = [c["header"] for c in columns if c["action"] == "formula" and c["formula"]
                            and "{" + d["header"] + "}" in c["formula"]]
        users += [f"cột '{h}' (mới)" for h in new_formulas_use]
        dropped.append({"src": col, "header": d["header"], "used_by": users, "mode": "hide" if users else "delete"})
    # label edits (title, field labels) on texts that exist in the template
    texts = {}
    for tx in ins.get("texts") or []:
        texts.setdefault(templates.norm(tx["text"]), tx)
    numbers_ok = allowed | _all_numbers(b)
    labels: list[dict] = []
    for lb in res.labels if not locked else []:
        tx = texts.get(templates.norm(lb.current))
        new = " ".join(lb.new.split())[:120]
        if not tx:
            warnings.append(f"Label '{lb.current[:60]}' is not in the template; not changed.")
            continue
        if not new or new == tx["text"] or any(x["cell"] == tx["cell"] for x in labels):
            continue
        bad = [tok for tok in sources.NUM_TOKEN_RE.findall(new)
               if (v := sources.parse_number(tok)[0]) is not None and round(v, 9) not in numbers_ok]
        if bad:
            warnings.append(f"Label '{new}' not used: {bad[:3]} not in the sources.")
            continue
        if tx["text"].rstrip().endswith(":") and not new.endswith(":"):
            new += ":"
        labels.append({"cell": tx["cell"], "row": tx["row"], "col": tx["col"], "current": tx["text"], "new": new,
                       "style": tx["style"], "reason": (lb.reason or "")[:160]})
    relabel = {templates.norm(x["current"]): x["new"].rstrip(":").strip() for x in labels}
    # header fields + rows, grounded
    kv_slots = {}
    for k in slots.kv:
        if k.sheet == t.sheet:
            kv_slots[templates.norm(k.label)] = k
            if templates.norm(k.label) in relabel:
                kv_slots.setdefault(templates.norm(relabel[templates.norm(k.label)]), k)
    kv = []
    for lk in res.kv:
        slot = kv_slots.get(templates.norm(lk.label))
        if not slot or any(x["cell"] == slot.cell for x in kv):
            continue
        fmt = "number" if sources.parse_number(lk.value)[0] is not None and not re.fullmatch(r"(19|20)\d{2}",
                                                                                            lk.value.strip()) else "text"
        g = ground(lk.value, lk.quote, fmt, b, prompt)
        if isinstance(g, str):
            warnings.append(f"Header field '{slot.label}' left blank ({g}).")
            continue
        kv.append({"sheet": slot.sheet, "cell": slot.cell, "orig_label": slot.label,
                   "label": relabel.get(templates.norm(slot.label), slot.label),
                   "value": g[0], "value_type": g[1], "quote": _norm_ws(g[3] if len(g) > 3 else lk.quote)[:200],
                   "locator": g[2]})
    by_final = {templates.norm(c["header"]): c for c in columns}
    data: list[dict] = []
    skipped = 0
    for lr in res.rows[:len(t.rows)]:
        row: dict[str, dict] = {}
        for cell in lr.cells:
            c = by_final.get(templates.norm(cell.header))
            if not c or c["action"] in ("formula", "sequence") or (c["action"] == "keep" and info[c["src"]]["formula"]):
                continue
            g = _ground_cell(cell, c["fmt"], b, prompt)
            if isinstance(g, str):
                skipped += 1
                if skipped <= 6:
                    warnings.append(f"Row {len(data) + 1}, '{c['header']}' left blank ({g}).")
                continue
            row[c["header"]] = g
        if row:
            data.append(row)
    overflow = max(0, len(res.rows) - len(t.rows))
    if overflow:
        warnings.append(f"{overflow} record(s) did not fit into the template's {len(t.rows)} rows (inserting rows is "
                        "not supported yet); they are not in the file.")
    if skipped > 6:
        warnings.append(f"{skipped - 6} more value(s) left blank because no source quote backs them.")
    sheets = [] if locked else _normalize_sheets(res, t.sheet, columns, ins.get("sheet_names") or [t.sheet], b,
                                                 prompt, numbers_ok, warnings)
    warnings += [f"AI: {n[:200]}" for n in res.notes[:4]]
    # what is missing, deterministically: header fields without a value, input columns without any value
    filled = {x["cell"] for x in kv}
    missing = [{"kind": "field", "what": relabel.get(templates.norm(k.label), k.label)}
               for k in slots.kv if k.sheet == t.sheet and k.cell not in filled]
    for c in columns:
        if c["action"] == "input" and c["default"] is None and data and not any(c["header"] in r for r in data):
            missing.append({"kind": "column", "what": c["header"]})
    und = res.understanding.model_dump() if res.understanding else None
    if und:
        missing += [{"kind": "ai", "what": m[:200]} for m in und.get("missing") or []]
    texts_all = " ".join(b.texts.values()) + " " + prompt
    title = next((x for x in ins.get("texts") or [] if x["row"] == min(y["row"] for y in ins["texts"])), None)
    return {"sheet": t.sheet, "header_row": t.header_row, "rows": list(t.rows), "start": min(info), "end": max(info),
            "columns": columns, "dropped": dropped, "kv": kv, "data": data, "warnings": warnings, "mode": mode,
            "info": {str(k): v for k, v in info.items()}, "apply": True, "understanding": und, "labels": labels,
            "sheets": sheets, "sheets_on": True, "missing": missing, "overflow": overflow,
            "totals": ins.get("totals"), "title_style": title["style"] if title else 0,
            "vi": bool(VI_RE.search(texts_all))}


def _all_numbers(b: sources.SourceBundle) -> set[float]:
    out: set[float] = set()
    for text in b.texts.values():
        out |= number_candidates(text)
    return out


def _only_dropped(user: str, info: dict, keep: set[int]) -> bool:
    """A use by another column of this table that is itself removed does not keep the column alive."""
    m = re.fullmatch(r"cột '(.*)'", user)
    if not m:
        return False
    col = next((c for c, d in info.items() if d["header"] == m[1]), None)
    return col is not None and col not in keep


def fill_only(plan: dict) -> dict:
    """The user declined structural changes: original columns, values only where the column is an input."""
    info = {int(k): v for k, v in plan["info"].items()}
    by_head = {c["header"]: c for c in plan["columns"]}
    seq = {c["src"] for c in plan["columns"] if c["action"] == "sequence" and c["src"] is not None
           and not info[c["src"]]["formula"]}
    cols = [{"src": c, "from": d["header"], "header": d["header"], "action": "sequence" if c in seq else "keep",
             "formula": None, "fmt": "number" if c in seq else "text", "total": False, "default": None,
             "default_type": None, "default_src": None, "reason": "", "changes": ["keep"]}
            for c, d in sorted(info.items())]
    rename = {c["header"]: info[c["src"]]["header"] for c in plan["columns"] if c["src"] is not None}
    data = []
    for row in plan.get("data") or []:
        new = {}
        for h, v in row.items():
            orig = rename.get(h)
            if orig and not info[by_head[h]["src"]]["formula"]:
                new[orig] = v
        if new:
            data.append(new)
    kv = [{**d, "label": d.get("orig_label", d["label"])} for d in plan.get("kv") or []]
    return {**plan, "columns": cols, "dropped": [], "data": data, "kv": kv, "labels": [], "sheets": [],
            "apply": False}


def effective(plan: dict) -> dict:
    if not plan.get("apply", True):
        return fill_only(plan)
    return plan if plan.get("sheets_on", True) else {**plan, "sheets": []}


# ---------------------------------------------------------------- dry run: what the file will contain
def _grid(path: str | Path) -> tuple[dict, dict]:
    import warnings as _w
    from openpyxl import load_workbook
    cells: dict[str, dict] = {}
    formulas: dict[str, dict] = {}
    with _w.catch_warnings():
        _w.simplefilter("ignore")
        wb = load_workbook(path)                         # read only; never saved
    try:
        for ws in wb.worksheets:
            cs, fs = cells.setdefault(ws.title, {}), formulas.setdefault(ws.title, {})
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        fs[(c.row, c.column)] = v
                    elif v is not None:
                        cs[(c.row, c.column)] = v.date().isoformat() if isinstance(v, dt.datetime) else v
    finally:
        wb.close()
    return cells, formulas


def _show(v: Any) -> Any:
    if v is formula_eval.UNSUPPORTED:
        return None
    if isinstance(v, float):
        return round(v, 6)
    return v


def table_preview(path: str | Path, plan: dict) -> dict:
    """Values (computed with the safe evaluator), formulas and provenance of the produced table(s)."""
    cells, formulas = _grid(path)
    wbx = formula_eval.Workbook(cells, formulas)
    items, _ = _layout(plan)
    sheet, hr, start = plan["sheet"], plan["header_row"], plan["start"]
    n = len(plan.get("data") or [])
    cols = []
    for k, (kind, v) in enumerate(items):
        c = plan["columns"][v] if kind == "col" else None
        cols.append({"letter": get_column_letter(start + k), "col": start + k,
                     "header": str(cells.get(sheet, {}).get((hr, start + k)) or ""),
                     "hidden": kind == "kept" and next(d for d in plan["dropped"] if d["src"] == v)["mode"] == "hide",
                     "action": c["action"] if c else "kept"})

    def cell(sh: str, r: int, col: int, data_row: dict | None, header: str) -> dict:
        f = formulas.get(sh, {}).get((r, col))
        if f:
            val = wbx.value(sh, r, col)
            return {"v": _show(val), "f": f, "prov": "derived", "pending": val is formula_eval.UNSUPPORTED}
        val = cells.get(sh, {}).get((r, col))
        d = (data_row or {}).get(header) or {}
        if val is None:
            return {"v": None, "prov": "missing"}
        if d:
            return {"v": val, "prov": d.get("prov", "explicit"), "src": d.get("locator"), "quote": d.get("quote"),
                    "note": d.get("note")}
        return {"v": val, "prov": "derived" if header in plan.get("_seq", ()) else "explicit"}
    plan["_seq"] = {c["header"] for c in plan["columns"] if c["action"] == "sequence"}
    rows = []
    for i, r in enumerate(plan["rows"][:n]):
        d = plan["data"][i]
        rows.append({"row": r, "cells": [cell(sheet, r, c["col"], d, c["header"]) for c in cols]})
    tot = plan.get("totals")
    totals = {"row": tot["row"], "cells": [cell(sheet, tot["row"], c["col"], None, "") for c in cols]} if tot else None
    extra = []
    for sh in plan.get("sheets") or []:
        hdr = [c["header"] for c in sh["columns"]]
        srows = []
        for i, d in enumerate(sh["data"]):
            srows.append({"row": 4 + i, "cells": [cell(sh["name"], 4 + i, j + 1, d, h) for j, h in enumerate(hdr)]})
        stot = None
        if any(c["total"] for c in sh["columns"]):
            r = 4 + len(sh["data"])
            stot = {"row": r, "cells": [cell(sh["name"], r, j + 1, None, "") for j in range(len(hdr))]}
        extra.append({"name": sh["name"], "purpose": sh["purpose"], "reason": sh["reason"],
                      "columns": [{"letter": get_column_letter(j + 1), "header": h, "action": c["action"],
                                   "formula": c["formula"]} for j, (h, c) in enumerate(zip(hdr, sh["columns"]))],
                      "rows": srows, "totals": stot})
    plan.pop("_seq", None)
    others = {}                                          # every other formula cell on the main sheet (totals, VAT ...)
    for (r, col), f in formulas.get(sheet, {}).items():
        if r < plan["rows"][0] or r > plan["rows"][-1]:
            others[f"{get_column_letter(col)}{r}"] = {"f": f, "v": _show(wbx.value(sheet, r, col))}
    return {"columns": cols, "rows": rows, "totals": totals, "sheets": extra, "other_formulas": others}


def dry_run(src: str | Path, plan: dict) -> dict:
    """Apply the effective plan to a temporary copy, validate it and read back what Excel will show."""
    import tempfile
    eff = effective(plan)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "dry.xlsx"
        try:
            apply_plan(src, out, eff)
        except InvalidOperation as e:
            msg = getattr(e, "message", str(e))
            return {"ok": False, "errors": [msg], "code": "reorder" if "order" in msg else "blocked",
                    "applied": eff.get("apply", True)}
        rep = validate(src, out, eff)
        if not rep["pass"]:
            return {"ok": False, "errors": rep["errors"], "code": "invalid", "applied": eff.get("apply", True)}
        grid = table_preview(out, eff)
    return {"ok": True, "errors": [], "validation": rep["level"], "applied": eff.get("apply", True), **grid}


def view(plan: dict) -> dict:
    """What the review screen shows: understanding, template vs proposed columns with reasons, label edits, extra
    sheets, the dry-run data preview (values, formulas, provenance) and what is missing."""
    eff = effective(plan)
    info = {int(k): v for k, v in eff["info"].items()}
    items, _ = _layout(eff)
    cols = []
    for k, (kind, v) in enumerate(items):
        letter = get_column_letter(eff["start"] + k)
        if kind == "kept":
            d = next(d for d in eff["dropped"] if d["src"] == v)
            cols.append({"letter": letter, "header": d["header"], "from": d["header"],
                         "changes": ["hidden"] if d["mode"] == "hide" else ["keep"], "kind": "formula" if
                         info[v]["formula"] else "input", "formula": info[v]["formula_h"], "fmt": None,
                         "default": None, "total": False, "reason": "; ".join(d["used_by"])[:200]})
            continue
        c = eff["columns"][v]
        f = c["formula"] if c["action"] == "formula" else (info[c["src"]]["formula_h"] if c["src"] is not None and
                                                            c["action"] == "keep" else None)
        cols.append({"letter": letter, "header": c["header"], "from": c["from"], "changes": c["changes"],
                     "kind": "sequence" if c["action"] == "sequence" else ("formula" if f else "input"), "formula": f,
                     "fmt": c["fmt"], "default": c["default"], "total": c.get("total", False), "reason": c["reason"]})
    removed = [{"header": d["header"], "mode": d["mode"], "used_by": d["used_by"]} for d in eff["dropped"]]
    template = [{"letter": get_column_letter(c), "header": d["header"], "kind": "formula" if d["formula"] else "input",
                 "formula": d["formula_h"]} for c, d in sorted(info.items())]
    prev = plan.get("preview") or {}
    return {"sheet": eff["sheet"], "applied": eff.get("apply", True), "columns": cols, "removed": removed,
            "template_columns": template,
            "kv": [{"label": d["label"], "cell": d["cell"], "value": d["value"], "source": d["locator"],
                    "quote": d.get("quote")} for d in eff["kv"]],
            "rows": [{h: v["value"] for h, v in r.items()} for r in eff["data"][:20]], "row_count": len(eff["data"]),
            "mode": eff["mode"], "understanding": plan.get("understanding"), "labels": eff.get("labels") or [],
            "sheets_proposed": len(plan.get("sheets") or []), "sheets_on": plan.get("sheets_on", True),
            "missing": plan.get("missing") or [], "overflow": plan.get("overflow", 0),
            "approved": bool(plan.get("approved")), "preview": prev,
            "valid": bool(prev.get("ok")) if prev else True}


# ---------------------------------------------------------------- generate
def run(run_: Any, ws: dict, state: dict) -> dict:
    from . import service, workspaces
    from .generate import _filename, _step
    from .template_fill import _template_source
    plan = effective(state["structure"])
    src, tpl_name = _template_source(state)
    meta = service.create_from_path(src, _filename(state, Path(tpl_name).stem), "template", ws["id"],
                                    extra={"template_file_id": state.get("template_file_id"),
                                           "template_id": state.get("template_id"), "llm_mode": plan["mode"]})
    wb_id = meta["id"]
    note = "restructure + fill from sources" if plan.get("apply", True) else "fill from sources"
    try:
        res = _step(run_, "restructure_template", lambda: service.apply_transform(
            wb_id, 1, f"restructure-{state['id']}", lambda s, o: _transform(s, o, plan), note=note),
            lambda r: f"revision {r.revision}, {len(r.changed_ranges)} range(s)")
    except Exception as e:                               # keep revision 1 (the template), report honestly
        msg = getattr(e, "message", str(e))
        run_.emit("validation_failed", {"errors": [msg] + list(getattr(e, "extra", {}).get("errors", []))[:10]})
        return {"errors": [msg], "workbook_id": wb_id}
    warnings = list(dict.fromkeys(state.get("warnings", []) + plan["warnings"] + res.validation.warnings))
    hidden = [d["header"] for d in plan["dropped"] if d["mode"] == "hide"]
    if hidden:
        warnings.append(f"Hidden instead of deleted (still used by formulas): {hidden}")
    # the file must show what the approved preview showed
    approved = (state["structure"].get("preview") or {})
    try:
        got = table_preview(service.revision_path(wb_id, res.revision), plan)
        same = {k: got[k] for k in ("rows", "totals", "sheets")} == {k: approved.get(k) for k in ("rows", "totals",
                                                                                                   "sheets")}
    except Exception as e:                               # noqa: BLE001
        same, got = False, None
        warnings.append(f"Could not compare the file with the preview ({type(e).__name__}).")
    if not same:
        warnings.append("The generated file differs from the approved preview; review it before use.")
    report = [{"sheet": d["sheet"], "cell": d["cell"], "label": d["label"], "value": d["value"],
               "locator": d["locator"], "via": "quote"} for d in plan["kv"]]
    items, _ = _layout(plan)
    letter = {plan["columns"][v]["header"]: get_column_letter(plan["start"] + k)
              for k, (kind, v) in enumerate(items) if kind == "col"}
    for r, row in zip(plan["rows"], plan["data"]):
        report += [{"sheet": plan["sheet"], "cell": f"{letter[h]}{r}", "label": h, "value": v["value"],
                    "locator": v["locator"], "via": "assumed" if v.get("prov") == "assumed" else "quote"}
                   for h, v in row.items() if h in letter]
    for sh in plan.get("sheets") or []:
        for i, row in enumerate(sh["data"]):
            hs = [c["header"] for c in sh["columns"]]
            report += [{"sheet": sh["name"], "cell": f"{get_column_letter(hs.index(h) + 1)}{4 + i}", "label": h,
                        "value": v["value"], "locator": v["locator"], "via": "quote"} for h, v in row.items()]
    service.update_metadata(wb_id, warnings=warnings, fill_report=report, structure=view(state["structure"]),
                            matches_preview=same)
    cur = workspaces.get(ws["id"])
    cur["workbook_id"] = wb_id
    cur["runs"] = (cur.get("runs") or [])[-20:] + [run_.id]
    workspaces.save(cur)
    run_.emit("workbook_revision", {"workbook_id": wb_id, "revision": res.revision,
                                    "changed_ranges": res.changed_ranges[:200]})
    return {"workbook_id": wb_id, "revision": res.revision, "warnings": warnings, "llm_mode": plan["mode"],
            "matches_preview": same}


def _transform(src: str, out: str, plan: dict) -> tuple[list[str], dict]:
    res = apply_plan(src, out, plan)
    return res["changed"], validate(src, out, plan)
