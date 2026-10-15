"""Small helpers for reading the raw OOXML parts of an .xlsx (zip)."""
from __future__ import annotations

import posixpath
import re
import zipfile
from html import unescape

CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
ATTR_RE = re.compile(r'([\w:]+)="([^"]*)"')


def attrs(s: str) -> dict[str, str]:
    return dict(ATTR_RE.findall(s))


def sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    """Sheet name -> zip part path (e.g. 'xl/worksheets/sheet1.xml'), in workbook order."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        a = attrs(rel)
        rid_target[a.get("Id")] = a.get("Target", "")
    out = {}
    for tag in re.findall(r"<sheet\b[^>]*>", wb):
        a = attrs(tag)
        target = rid_target.get(a.get("r:id"), "")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        out[unescape(a["name"])] = part
    return out


def hidden_sheets(z: zipfile.ZipFile) -> set[str]:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    out = set()
    for tag in re.findall(r"<sheet\b[^>]*>", wb):
        a = attrs(tag)
        if a.get("state") in ("hidden", "veryHidden"):
            out.add(unescape(a["name"]))
    return out


def rels_part(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def split_sheet_data(xml: str) -> tuple[str, str, str]:
    """(before, sheetData inner, after). sheetData may be self-closing."""
    m = re.search(r"<sheetData\s*/>", xml)
    if m:
        return xml[: m.start()] + "<sheetData>", "", "</sheetData>" + xml[m.end():]
    s = xml.index("<sheetData")
    s = xml.index(">", s) + 1
    e = xml.index("</sheetData>")
    return xml[:s], xml[s:e], xml[e:]
