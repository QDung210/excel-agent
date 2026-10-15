"""Typed value edits applied directly to the OOXML of an existing .xlsx (no openpyxl save).

Only these parts may change:
  - worksheet XML of edited sheets (cells inside <sheetData>, <dimension>)
  - worksheet XML of sheets whose formulas depend on an edited sheet (cached <v> stripped)
  - xl/sharedStrings.xml (strings appended, never removed or renumbered)
  - xl/workbook.xml (calcPr fullCalcOnLoad="1" when formula caches were stripped)
Every other ZIP member is copied byte for byte. Excel recalculates on open; we never claim
computed values.
"""
from __future__ import annotations

import datetime as dt
import re
import zipfile
from dataclasses import dataclass, field
from html import escape, unescape
from typing import Any, Iterable

from openpyxl.utils import get_column_letter
from xlsx_template import ooxml
from xlsx_template.writer import _force_recalc

from .errors import InvalidOperation
from .schemas import ClearValue, SetValue, parse_a1, parse_range

SST_PART = "xl/sharedStrings.xml"
ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
BAD_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
DATE_FMT_IDS = set(range(14, 23)) | {45, 46, 47}


@dataclass
class EditStats:
    changed: list[str] = field(default_factory=list)          # "Sheet!B3"
    intended: dict[tuple[str, str], Any] = field(default_factory=dict)   # (sheet, cell) -> value or None
    touched_parts: set[str] = field(default_factory=set)
    formula_cache_stripped: int = 0
    strings_added: int = 0
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------- helpers
def _xml_text(s: str) -> str:
    if BAD_XML_CHARS.search(s):
        raise InvalidOperation("Text contains control characters that cannot be stored in Excel.")
    return escape(s, quote=False)


def _num(v: float | int) -> str:
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer() and abs(v) < 1e15):
        return str(int(v))
    return repr(float(v))


def _date_serial(value: Any, date1904: bool) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    try:
        d = dt.datetime.fromisoformat(str(value).strip())
    except ValueError:
        raise InvalidOperation(f"{value!r} is not an ISO date (YYYY-MM-DD).") from None
    epoch = dt.datetime(1904, 1, 1) if date1904 else dt.datetime(1899, 12, 30)
    delta = d - epoch
    serial = delta.days + delta.seconds / 86400
    if not date1904 and serial < 61:
        raise InvalidOperation("Dates before 1900-03-01 are not supported.")
    return serial


def coerce(action: SetValue, date1904: bool = False) -> tuple[str, Any]:
    """-> (kind, value) with kind in number|bool|string|blank. Formulas are never written."""
    v, vt = action.value, action.value_type
    if v is None or (isinstance(v, str) and v == "" and vt in ("auto", "number", "date")):
        return "blank", None
    if vt == "string":
        return "string", str(v)
    if vt == "bool":
        if isinstance(v, bool):
            return "bool", v
        s = str(v).strip().lower()
        if s in ("true", "1", "yes"):
            return "bool", True
        if s in ("false", "0", "no"):
            return "bool", False
        raise InvalidOperation(f"{v!r} is not a boolean.")
    if vt == "date":
        return "date", _date_serial(v, date1904)
    if isinstance(v, bool):
        return "bool", v
    if isinstance(v, (int, float)):
        return "number", v
    s = str(v)
    if vt == "number":
        t = s.strip().replace(",", "")
        if not NUM_RE.match(t):
            raise InvalidOperation(f"{v!r} is not a number.")
        return "number", float(t) if any(ch in t for ch in ".eE") else int(t)
    # auto + text: plain numerals become numbers, everything else (including "=...") stays literal text
    t = s.strip()
    if NUM_RE.match(t) and not (len(t) > 1 and t.startswith("0") and t[1].isdigit()):
        return "number", float(t) if any(ch in t for ch in ".eE") else int(t)
    return "string", s


class _Styles:
    def __init__(self, xml: str | None):
        self.xf_numfmt: list[int] = []
        self.custom: dict[int, str] = {}
        if not xml:
            return
        for m in re.finditer(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', xml):
            self.custom[int(m.group(1))] = unescape(m.group(2))
        cx = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", xml, re.S)
        if cx:
            for m in re.finditer(r"<xf\b([^>]*)", cx.group(1)):
                self.xf_numfmt.append(int(ooxml.attrs(m.group(1)).get("numFmtId", "0")))

    def is_date(self, s: int) -> bool:
        if s >= len(self.xf_numfmt):
            return False
        fid = self.xf_numfmt[s]
        if fid in DATE_FMT_IDS:
            return True
        code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", self.custom.get(fid, "")).lower()
        return bool(re.search(r"[dy]", code))


class _SST:
    def __init__(self, xml: str | None):
        self.xml = xml
        self.index: dict[str, int] = {}
        self.added: list[str] = []
        self.count = 0
        if xml is None:
            return
        items = re.findall(r"<si>(.*?)</si>|<si/>", xml, re.S)
        self.count = len(items)
        for i, body in enumerate(items):
            if body and "<r>" not in body:     # plain (non rich-text) strings are reusable
                m = re.fullmatch(r"\s*<t(?:\s[^>]*)?>(.*?)</t>\s*|\s*<t\s*/>\s*", body, re.S)
                if m:
                    self.index.setdefault(unescape(m.group(1) or ""), i)

    def ref(self, text: str) -> int:
        if text in self.index:
            return self.index[text]
        i = self.count + len(self.added)
        self.added.append(text)
        self.index[text] = i
        return i

    def render(self, total_refs: int) -> str:
        assert self.xml is not None
        new = "".join(f'<si><t xml:space="preserve">{_xml_text(t)}</t></si>' for t in self.added)
        xml = self.xml
        if re.search(r"<sst\b[^>]*/>", xml):
            xml = re.sub(r"<sst\b([^>]*)/>", r"<sst\1></sst>", xml, count=1)
        xml = xml.replace("</sst>", new + "</sst>", 1)
        head_end = xml.index(">", xml.index("<sst")) + 1
        head = xml[:head_end]
        uniq = self.count + len(self.added)
        head = re.sub(r'\suniqueCount="\d+"', f' uniqueCount="{uniq}"', head) if "uniqueCount=" in head \
            else head[:-1] + f' uniqueCount="{uniq}">'
        head = re.sub(r'\scount="\d+"', f' count="{total_refs}"', head) if re.search(r'\scount="', head) \
            else head[:-1] + f' count="{total_refs}">'
        return head + xml[head_end:]


def _merged_ranges(xml: str) -> list[tuple[int, int, int, int]]:
    return [parse_range(r) for r in re.findall(r'<mergeCell\b[^>]*\bref="([^"]+)"', xml)]


def _col_styles(xml: str) -> dict[int, str]:
    out: dict[int, str] = {}
    for m in re.finditer(r"<col\b([^>]*)/?>", xml):
        a = ooxml.attrs(m.group(1))
        if "style" in a and a["style"] != "0":
            for c in range(int(a.get("min", "1")), min(int(a.get("max", "1")), 16384) + 1):
                out[c] = a["style"]
    return out


def _sheet_refs(formula_text: str) -> set[str]:
    out = set()
    for m in re.finditer(r"'((?:[^']|'')+)'!|([A-Za-z_À-￿][\w.À-￿]*)!", formula_text):
        out.add((m.group(1) or "").replace("''", "'") if m.group(1) else m.group(2))
    return out


# ---------------------------------------------------------------- sheet patch
def _cell_xml(ref: str, style: str | None, kind: str, value: Any, sst: _SST, keep: bool = False) -> str:
    """keep=True: an existing cell keeps its style attribute verbatim (even an explicit s="0")."""
    s = f' s="{style}"' if style and (keep or style != "0") else ""
    if kind == "blank":
        return f'<c r="{ref}"{s}/>'
    if kind in ("number", "date"):
        return f'<c r="{ref}"{s}><v>{_num(value)}</v></c>'
    if kind == "bool":
        return f'<c r="{ref}"{s} t="b"><v>{1 if value else 0}</v></c>'
    text = str(value)
    if len(text) > 32767:
        raise InvalidOperation("Excel cells hold at most 32767 characters.")
    if sst.xml is not None:
        _xml_text(text)
        return f'<c r="{ref}"{s} t="s"><v>{sst.ref(text)}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{_xml_text(text)}</t></is></c>'


def _patch_sheet(sheet: str, xml: str, edits: dict[tuple[int, int], tuple[str, Any]], sst: _SST,
                 styles: _Styles, stats: EditStats) -> str:
    merged = _merged_ranges(xml)
    for (r, c) in edits:
        for (r1, c1, r2, c2) in merged:
            if r1 <= r <= r2 and c1 <= c <= c2 and (r, c) != (r1, c1):
                raise InvalidOperation(f"{sheet}!{get_column_letter(c)}{r} is inside a merged range; edit its top-left cell "
                                       f"{get_column_letter(c1)}{r1} instead.")
    col_style = _col_styles(xml)
    before, inner, after = ooxml.split_sheet_data(xml)

    by_row: dict[int, dict[int, tuple[str, Any]]] = {}
    for (r, c), v in edits.items():
        by_row.setdefault(r, {})[c] = v

    pieces: list[str] = []
    pos = 0
    pending = sorted(by_row)
    pi = 0

    def new_row(r: int) -> str:
        cells = "".join(_cell_xml(f"{get_column_letter(c)}{r}", col_style.get(c), k, v, sst)
                        for c, (k, v) in sorted(by_row[r].items()) if k != "blank")
        for c, (k, v) in sorted(by_row[r].items()):
            stats.changed.append(f"{sheet}!{get_column_letter(c)}{r}")
        return f'<row r="{r}">{cells}</row>' if cells else ""

    for m in ROW_RE.finditer(inner):
        a = ooxml.attrs(m.group(1))
        if "r" not in a:
            raise InvalidOperation(f"Sheet {sheet!r} has rows without explicit numbers; editing it is not supported yet.")
        rnum = int(a["r"])
        while pi < len(pending) and pending[pi] < rnum:          # rows that do not exist yet
            pieces.append(inner[pos:m.start()])
            pos = m.start()
            pieces.append(new_row(pending[pi]))
            pi += 1
        if pi < len(pending) and pending[pi] == rnum:
            pieces.append(inner[pos:m.start()])
            pieces.append(_patch_row(sheet, rnum, m.group(1), m.group(2) or "", by_row[rnum], col_style, sst, styles, stats))
            pos = m.end()
            pi += 1
    pieces.append(inner[pos:])
    while pi < len(pending):
        pieces.append(new_row(pending[pi]))
        pi += 1
    before = _grow_dimension(before, edits)
    return before + "".join(pieces) + after


def _patch_row(sheet: str, rnum: int, row_attrs: str, body: str, cols: dict[int, tuple[str, Any]],
               col_style: dict[int, str], sst: _SST, styles: _Styles, stats: EditStats) -> str:
    ra = ooxml.attrs(row_attrs)
    row_style = ra.get("s") if ra.get("customFormat") in ("1", "true") else None
    out: list[str] = []
    pos = 0
    remaining = dict(cols)

    def emit_new(upto: int | None) -> None:
        for c in sorted(list(remaining)):
            if upto is not None and c >= upto:
                break
            k, v = remaining.pop(c)
            ref = f"{get_column_letter(c)}{rnum}"
            stats.changed.append(f"{sheet}!{ref}")
            if k != "blank":
                out.append(_cell_xml(ref, row_style or col_style.get(c), k, v, sst))

    for m in ooxml.CELL_RE.finditer(body):
        a = ooxml.attrs(m.group(1))
        if "r" not in a:
            raise InvalidOperation(f"Sheet {sheet!r} has cells without references; editing it is not supported yet.")
        _, c = parse_a1(a["r"])
        out.append(body[pos:m.start()])
        pos = m.end()
        emit_new(c)
        if c in remaining:
            cbody = m.group(2) or ""
            if "<f" in cbody:
                raise InvalidOperation(f"{sheet}!{a['r']} contains a formula; formula edits are not supported.")
            k, v = remaining.pop(c)
            if k == "date" and not styles.is_date(int(a.get("s", "0"))):
                stats.warnings.append(f"{sheet}!{a['r']}: stored as a date serial but the cell has no date format.")
            stats.changed.append(f"{sheet}!{a['r']}")
            out.append(_cell_xml(a["r"], a.get("s"), k, v, sst, keep=True))
        else:
            out.append(m.group(0))
    out.append(body[pos:])
    emit_new(None)
    new_body = "".join(out)

    spans = ra.get("spans")
    if spans and re.fullmatch(r"\d+:\d+", spans):
        lo, hi = map(int, spans.split(":"))
        if any(c < lo or c > hi for c in cols):
            row_attrs = re.sub(r'\sspans="[^"]*"', "", row_attrs)
    if not new_body:
        return f"<row{row_attrs}/>"
    return f"<row{row_attrs}>{new_body}</row>"


def _grow_dimension(before: str, edits: Iterable[tuple[int, int]]) -> str:
    m = re.search(r'<dimension\b[^>]*\bref="([^"]+)"', before)
    if not m:
        return before
    try:
        r1, c1, r2, c2 = parse_range(m.group(1))
    except ValueError:
        return before
    nr1, nc1, nr2, nc2 = r1, c1, r2, c2
    for r, c in edits:
        nr1, nc1, nr2, nc2 = min(nr1, r), min(nc1, c), max(nr2, r), max(nc2, c)
    if (nr1, nc1, nr2, nc2) == (r1, c1, r2, c2):
        return before
    ref = f"{get_column_letter(nc1)}{nr1}:{get_column_letter(nc2)}{nr2}"
    return before[: m.start(1)] + ref + before[m.end(1):]


def _strip_formula_cache(xml: str) -> tuple[str, int]:
    before, inner, after = ooxml.split_sheet_data(xml)
    n = 0

    def repl(m: re.Match) -> str:
        nonlocal n
        attr_s, body = m.group(1), m.group(2) or ""
        if "<f" in body and re.search(r"<v[ >/]", body):
            n += 1
            body = re.sub(r"<v(?:\s[^>]*)?>.*?</v>|<v(?:\s[^>]*)?/>", "", body, flags=re.S)   # incl. xml:space
            attr_s = re.sub(r'\st="(?:str|n|b|e)"', "", attr_s)
            return f"<c{attr_s}>{body}</c>"
        return m.group(0)

    return before + ooxml.CELL_RE.sub(repl, inner) + after, n


# ---------------------------------------------------------------- public
def date1904(z: zipfile.ZipFile) -> bool:
    return bool(re.search(r'<workbookPr\b[^>]*date1904="(?:1|true)"', z.read("xl/workbook.xml").decode("utf-8")))


def apply_actions(src: str, dst: str, actions: list[SetValue | ClearValue]) -> EditStats:
    stats = EditStats()
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        raw = {i.filename: z.read(i.filename) for i in infos}
        sheets = ooxml.sheet_parts(z)
        is1904 = date1904(z)

    edits: dict[str, dict[tuple[int, int], tuple[str, Any]]] = {}
    for act in actions:
        if act.sheet not in sheets:
            raise InvalidOperation(f"Sheet {act.sheet!r} does not exist.", sheet=act.sheet)
        r, c = parse_a1(act.cell)
        kind, val = ("blank", None) if isinstance(act, ClearValue) else coerce(act, is1904)
        edits.setdefault(act.sheet, {})[(r, c)] = (kind, val)       # last action on a cell wins
        stats.intended[(act.sheet, f"{get_column_letter(c)}{r}")] = val

    sst = _SST(raw[SST_PART].decode("utf-8") if SST_PART in raw else None)
    styles = _Styles(raw["xl/styles.xml"].decode("utf-8") if "xl/styles.xml" in raw else None)
    changed: dict[str, str] = {}
    for name, cells in edits.items():
        part = sheets[name]
        if part not in raw:
            raise InvalidOperation(f"Sheet {name!r} has no worksheet part (chart sheet?).")
        xml = raw[part].decode("utf-8")
        if "<sheetData" not in xml:
            raise InvalidOperation(f"Sheet {name!r} is not a worksheet.")
        changed[part] = _patch_sheet(name, xml, cells, sst, styles, stats)

    # formula caches: edited sheets + sheets whose formulas reference a dirty sheet (fixpoint)
    wb_xml = raw["xl/workbook.xml"].decode("utf-8")
    name_refs: dict[str, set[str]] = {}
    for m in re.finditer(r'<definedName\b([^>]*)>(.*?)</definedName>', wb_xml, re.S):
        nm = ooxml.attrs(m.group(1)).get("name", "")
        if nm and not nm.startswith("_xlnm."):
            name_refs[nm] = _sheet_refs(unescape(m.group(2)))
    part_xml = {n: changed.get(p) or raw[p].decode("utf-8") for n, p in sheets.items() if p in raw}
    deps: dict[str, set[str]] = {}
    for n, xml in part_xml.items():
        fs = " ".join(unescape(f) for f in re.findall(r"<f\b[^>]*>(.*?)</f>", xml, re.S))
        refs = _sheet_refs(fs)
        for nm, targets in name_refs.items():
            if re.search(rf"(?<![\w.]){re.escape(nm)}(?![\w(])", fs):
                refs |= targets
        if "<f" in xml:
            refs.add(n)         # same-sheet formulas
        deps[n] = refs
    dirty = set(edits)
    grew = True
    while grew:
        grew = False
        for n, refs in deps.items():
            if n not in dirty and refs & dirty:
                dirty.add(n)
                grew = True
    for n in dirty:
        if n in part_xml and "<f" in part_xml[n]:
            new_xml, k = _strip_formula_cache(part_xml[n])
            if k:
                changed[sheets[n]] = new_xml
                stats.formula_cache_stripped += k
    if "<externalReference" in wb_xml:
        stats.warnings.append("Workbook has external links; their cached values were not refreshed.")

    if sst.added or (sst.xml is not None and any(p in changed for p in sheets.values())):
        total = sum(len(re.findall(r'<c\b[^>]*\bt="s"', changed.get(p) or raw[p].decode("utf-8")))
                    for p in sheets.values() if p in raw)
        new_sst = sst.render(total)
        if new_sst != sst.xml:
            changed[SST_PART] = new_sst
        stats.strings_added = len(sst.added)
    if stats.formula_cache_stripped:
        changed["xl/workbook.xml"] = _force_recalc(wb_xml)

    with zipfile.ZipFile(dst, "w") as out:
        for info in infos:
            if info.filename in changed and changed[info.filename].encode("utf-8") != raw[info.filename]:
                stats.touched_parts.add(info.filename)
                out.writestr(info, changed[info.filename].encode("utf-8"), compress_type=info.compress_type)
            else:
                out.writestr(info, raw[info.filename], compress_type=info.compress_type)
    return stats
