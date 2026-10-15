"""Stage 4: edit the ORIGINAL .xlsx in place at the XML level.

Why not openpyxl: it silently drops x14 data validations (dropdowns referencing other
sheets), named sheet views, and other parts it does not model. Here we only touch:
  - <c> elements of cleared cells inside <sheetData> (value removed, style kept)
  - cached <v> of formula cells (old results would leak data; Excel recalculates on open)
  - sharedStrings.xml (unused strings removed, indexes remapped)
  - <hyperlink> of cleared cells (+ its relationship, e.g. mailto: links)
  - <calcPr fullCalcOnLoad="1"> in workbook.xml
Every other part is copied byte for byte.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass, field

from openpyxl.utils import coordinate_to_tuple

from . import ooxml

SST_PART = "xl/sharedStrings.xml"


@dataclass
class WriteStats:
    cleared: int = 0
    formula_cache_stripped: int = 0
    hyperlinks_removed: int = 0
    shared_strings_before: int = 0
    shared_strings_after: int = 0
    warnings: list[str] = field(default_factory=list)


def _patch_cells(inner: str, clear: set[tuple[int, int]], strip_cache: bool, stats: WriteStats) -> str:
    def repl(m: re.Match) -> str:
        attr_s, body = m.group(1), m.group(2) or ""
        a = ooxml.attrs(attr_s)
        ref = a.get("r")
        if ref is None:
            return m.group(0)
        if "<f" in body:
            if strip_cache and re.search(r"<v[ >/]", body):
                stats.formula_cache_stripped += 1
                body = re.sub(r"<v>.*?</v>|<v/>", "", body, flags=re.S)
                return f"<c{attr_s}>{body}</c>"
            return m.group(0)
        if coordinate_to_tuple(ref) in clear and (re.search(r"<v[ >]", body) or "<is>" in body):
            stats.cleared += 1
            keep = f' r="{ref}"' + (f' s="{a["s"]}"' if "s" in a else "")
            return f"<c{keep}/>"
        return m.group(0)

    return ooxml.CELL_RE.sub(repl, inner)


def _remove_hyperlinks(xml: str, clear: set[tuple[int, int]], stats: WriteStats) -> tuple[str, set[str]]:
    removed_rids: set[str] = set()

    def repl(m: re.Match) -> str:
        a = ooxml.attrs(m.group(0))
        ref = a.get("ref", "")
        if ":" not in ref and ref and coordinate_to_tuple(ref) in clear:
            stats.hyperlinks_removed += 1
            if a.get("r:id"):
                removed_rids.add(a["r:id"])
            return ""
        return m.group(0)

    xml = re.sub(r"<hyperlink\b[^>]*?(?:/>|>.*?</hyperlink>)", repl, xml, flags=re.S)
    xml = re.sub(r"<hyperlinks>\s*</hyperlinks>", "", xml)
    return xml, removed_rids


def _remap_shared_strings(parts: dict[str, str], sst_xml: str, stats: WriteStats) -> str:
    """Drop shared strings no cell references any more; renumber references."""
    items = re.findall(r"<si>.*?</si>|<si/>", sst_xml, re.S)
    stats.shared_strings_before = len(items)
    used: list[int] = []
    seen: set[int] = set()
    total_refs = 0
    s_cell = re.compile(r'<c\b[^>]*\bt="s"[^>]*>\s*<v>(\d+)</v>')
    for xml in parts.values():
        for m in s_cell.finditer(xml):
            i = int(m.group(1))
            total_refs += 1
            if i not in seen:
                seen.add(i)
                used.append(i)
    used.sort()
    new_idx = {old: new for new, old in enumerate(used)}

    for name in parts:
        parts[name] = s_cell.sub(lambda m: re.sub(r"<v>\d+</v>$", f"<v>{new_idx[int(m.group(1))]}</v>", m.group(0)), parts[name])

    head = sst_xml[: sst_xml.index(">", sst_xml.index("<sst")) + 1]
    head = re.sub(r'\scount="\d+"', f' count="{total_refs}"', head)
    head = re.sub(r'\suniqueCount="\d+"', f' uniqueCount="{len(used)}"', head)
    stats.shared_strings_after = len(used)
    return head + "".join(items[i] for i in used) + "</sst>"


def _force_recalc(wb_xml: str) -> str:
    if "<calcPr" in wb_xml:
        def repl(m: re.Match) -> str:
            tag = m.group(0)
            if "fullCalcOnLoad=" in tag:
                return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag)
            return tag.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        return re.sub(r"<calcPr\b[^>]*>", repl, wb_xml, count=1)
    for anchor in ("</definedNames>", "</externalReferences>", "</functionGroups>", "</sheets>"):
        if anchor in wb_xml:
            return wb_xml.replace(anchor, anchor + '<calcPr fullCalcOnLoad="1"/>', 1)
    return wb_xml


def write_template(src: str, dst: str, clear: dict[str, set[tuple[int, int]]],
                   strip_formula_cache: bool = True) -> WriteStats:
    stats = WriteStats()
    if src != dst:
        shutil.copyfile(src, dst)  # keeps a copy even if patching fails below
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        raw = {i.filename: z.read(i.filename) for i in infos}
        sheets = ooxml.sheet_parts(z)

    changed: dict[str, str] = {}
    all_ws = [p for p in raw if p.startswith("xl/worksheets/") and p.endswith(".xml") and "/_rels/" not in p]
    for part in all_ws:
        xml = raw[part].decode("utf-8")
        name = next((n for n, p in sheets.items() if p == part), None)
        cells = clear.get(name, set()) if name else set()
        before, inner, after = ooxml.split_sheet_data(xml)
        inner = _patch_cells(inner, cells, strip_formula_cache, stats)
        after, rids = _remove_hyperlinks(after, cells, stats)
        changed[part] = before + inner + after
        if rids:
            rp = ooxml.rels_part(part)
            if rp in raw:
                rels = raw[rp].decode("utf-8")
                for rid in rids:
                    rels = re.sub(rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*/>', "", rels)
                changed[rp] = rels
        if name is None:
            stats.warnings.append(f"{part}: not listed in workbook.xml, only formula caches touched")

    if SST_PART in raw:
        ws_xml = {p: changed[p] for p in all_ws}
        changed[SST_PART] = _remap_shared_strings(ws_xml, raw[SST_PART].decode("utf-8"), stats)
        changed.update(ws_xml)

    if strip_formula_cache:
        changed["xl/workbook.xml"] = _force_recalc(raw["xl/workbook.xml"].decode("utf-8"))

    for p in raw:
        if re.match(r"xl/comments\d*\.xml|xl/threadedComments/", p):
            stats.warnings.append(f"{p}: cell comments are kept as-is (may contain instance data)")

    with zipfile.ZipFile(dst, "w") as out:
        for info in infos:
            data = changed[info.filename].encode("utf-8") if info.filename in changed else raw[info.filename]
            out.writestr(info, data, compress_type=info.compress_type)
    return stats
