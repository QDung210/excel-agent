"""Decision cache: reuse keep/clear decisions for regions with the same structure.

Two levels, so data changes still hit while a different kind of workbook never does:
  namespace  = hash of the template skeleton: sheet names + texts of label / header-band regions
               (not bold titles: a project title changes per project)
  region key = sheet, column(s), first row, parser annotation, type, style, header band,
               header above / label on the left, sequence, dropdown flag  (no data values, no
               last row, so a report with more rows this week still matches)
Only LLM answers keep/clear are stored; "review" and fallbacks are always asked again.
"""
from __future__ import annotations

import hashlib
import json
import os
import tempfile
import threading
import time
from pathlib import Path

from openpyxl.utils import get_column_letter

from .regions import ParsedWorkbook, Region

# Bump when the prompt / region builder changes in a way that should invalidate old answers.
CACHE_VERSION = 2
DEFAULT_PATH = Path(__file__).resolve().parent.parent / "cache" / "decisions.json"
_lock = threading.Lock()


def _h(obj) -> str:
    return hashlib.sha1(json.dumps(obj, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:16]


def namespace(pw: ParsedWorkbook) -> str:
    skeleton = sorted(
        (r.sheet, r.samples[0] if r.samples else "")
        for r in pw.regions if r.annotation == "label" or r.in_header
    )
    return _h({"v": CACHE_VERSION, "sheets": pw.sheets, "skeleton": skeleton})


def region_key(r: Region) -> str:
    cols = sorted({c for _, c in r.cells})
    return _h({
        "sheet": r.sheet, "cols": [get_column_letter(c) for c in cols], "row0": min(x for x, _ in r.cells),
        "ann": r.annotation, "type": r.vtype, "bold": r.bold, "fill": bool(r.fill), "hdr": r.in_header,
        "above": r.header_above, "left": r.label_left if r.vtype == "text" else "",
        "seq": bool(r.sequence), "dd": r.dropdown_source,
    })


class DecisionCache:
    def __init__(self, path: str | Path = DEFAULT_PATH):
        self.path = Path(path)

    def _load(self) -> dict:
        try:
            with open(self.path, encoding="utf-8") as f:
                data = json.load(f)
            return data if data.get("version") == CACHE_VERSION else {"version": CACHE_VERSION, "ns": {}}
        except (FileNotFoundError, json.JSONDecodeError):
            return {"version": CACHE_VERSION, "ns": {}}

    def lookup(self, pw: ParsedWorkbook) -> tuple[dict[int, str], list[Region]]:
        """-> ({region id: action} for hits, regions still to decide)."""
        with _lock:
            entries = self._load()["ns"].get(namespace(pw), {})
        hits, pending = {}, []
        for r in pw.regions:
            e = entries.get(region_key(r))
            if e:
                hits[r.id] = e["a"]
            else:
                pending.append(r)
        return hits, pending

    def store(self, pw: ParsedWorkbook, actions: dict[int, str]) -> int:
        """actions: {region id: keep|clear}; others are ignored. Returns number stored."""
        by_id = {r.id: r for r in pw.regions}
        new = {region_key(by_id[i]): {"a": a, "t": int(time.time())}
               for i, a in actions.items() if a in ("keep", "clear") and i in by_id}
        if not new:
            return 0
        with _lock:
            data = self._load()
            data["ns"].setdefault(namespace(pw), {}).update(new)
            self.path.parent.mkdir(parents=True, exist_ok=True)
            fd, tmp = tempfile.mkstemp(dir=self.path.parent, suffix=".tmp")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            os.replace(tmp, self.path)          # atomic: a crash never leaves half a file
        return len(new)
