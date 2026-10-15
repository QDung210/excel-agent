"""Stage 5: prove the template only lost the values we meant to clear.

Checks original vs output:
  - same zip parts; untouched parts byte-identical
  - every changed XML part is well-formed
  - per sheet: everything outside <sheetData>/<hyperlinks> identical
    (merges, column widths, data validations incl. x14, conditional formats...)
  - per cell: same style, same formula; cleared cells empty; all other constants unchanged
  - output opens with openpyxl
"""
from __future__ import annotations

import re
import warnings
import zipfile
from dataclasses import dataclass, field

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple

from . import ooxml

ALLOWED_CHANGES = {"xl/sharedStrings.xml", "xl/workbook.xml"}


@dataclass
class ValidationReport:
    ok: bool = True
    errors: list[str] = field(default_factory=list)
    checks: dict[str, int] = field(default_factory=dict)

    def fail(self, msg: str) -> None:
        self.ok = False
        if len(self.errors) < 50:
            self.errors.append(msg)


def _sst(z: zipfile.ZipFile) -> list[str]:
    try:
        xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    return ["".join(re.findall(r"<t\b[^>]*>(.*?)</t>|<t/>", si, re.S)) for si in re.findall(r"<si>.*?</si>|<si/>", xml, re.S)]


def _cells(xml: str, sst: list[str]) -> dict[str, tuple[str | None, str | None, str | None]]:
    """ref -> (style, formula xml, resolved constant value)."""
    _, inner, _ = ooxml.split_sheet_data(xml)
    out = {}
    for m in ooxml.CELL_RE.finditer(inner):
        a = ooxml.attrs(m.group(1))
        body = m.group(2) or ""
        f = re.search(r"<f\b.*?(?:/>|</f>)", body, re.S)
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        value = None
        if f is None:
            if a.get("t") == "s" and v:
                value = sst[int(v.group(1))]
            elif "<is>" in body:
                value = "".join(re.findall(r"<t\b[^>]*>(.*?)</t>", body, re.S))
            elif v:
                value = v.group(1)
        out[a["r"]] = (a.get("s"), f.group(0) if f else None, value)
    return out


def _outside(xml: str) -> str:
    before, _, after = ooxml.split_sheet_data(xml)
    after = re.sub(r"<hyperlinks>.*?</hyperlinks>", "", after, flags=re.S)
    return before + after


def validate(src: str, dst: str, clear: dict[str, set[tuple[int, int]]]) -> ValidationReport:
    rep = ValidationReport()
    with zipfile.ZipFile(src) as zo, zipfile.ZipFile(dst) as zn:
        names_o, names_n = zo.namelist(), zn.namelist()
        if names_o != names_n:
            rep.fail(f"zip parts differ: {set(names_o) ^ set(names_n)}")
        sheets = ooxml.sheet_parts(zo)
        sheet_of = {p: n for n, p in sheets.items()}
        sst_o, sst_n = _sst(zo), _sst(zn)
        n_same = n_changed = 0
        for part in names_o:
            if part not in names_n:
                continue
            a, b = zo.read(part), zn.read(part)
            if a == b:
                n_same += 1
                continue
            n_changed += 1
            is_ws = part.startswith("xl/worksheets/") and part.endswith(".xml") and "/_rels/" not in part
            is_rels = part.startswith("xl/worksheets/_rels/")
            if not (is_ws or is_rels or part in ALLOWED_CHANGES):
                rep.fail(f"{part}: unexpected change")
            if part.endswith((".xml", ".rels")):
                try:
                    etree.fromstring(b)
                except etree.XMLSyntaxError as e:
                    rep.fail(f"{part}: malformed XML: {e}")
        rep.checks["parts_identical"] = n_same
        rep.checks["parts_changed"] = n_changed

        cleared_ok = kept_ok = formulas = 0
        for part, name in sheet_of.items():
            xo, xn = zo.read(part).decode("utf-8"), zn.read(part).decode("utf-8")
            if _outside(xo) != _outside(xn):
                rep.fail(f"{name}: sheet structure outside cell data changed")
            co, cn = _cells(xo, sst_o), _cells(xn, sst_n)
            if co.keys() != cn.keys():
                rep.fail(f"{name}: cell set changed ({len(co)} -> {len(cn)})")
            targets = clear.get(name, set())
            for ref, (s, f, v) in co.items():
                if ref not in cn:
                    continue
                s2, f2, v2 = cn[ref]
                if s != s2:
                    rep.fail(f"{name}!{ref}: style changed")
                if f != f2:
                    rep.fail(f"{name}!{ref}: formula changed")
                if f is not None:
                    formulas += 1
                    continue
                if coordinate_to_tuple(ref) in targets:
                    if v2 is not None:
                        rep.fail(f"{name}!{ref}: should be cleared but has value")
                    else:
                        cleared_ok += 1
                elif v != v2:
                    rep.fail(f"{name}!{ref}: kept value changed {v!r} -> {v2!r}")
                elif v is not None:
                    kept_ok += 1
        rep.checks["cells_cleared"] = cleared_ok
        rep.checks["constants_kept"] = kept_ok
        rep.checks["formulas_preserved"] = formulas

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            load_workbook(dst).close()
        rep.checks["opens_in_openpyxl"] = 1
    except Exception as e:
        rep.fail(f"openpyxl cannot open output: {e}")
    return rep
