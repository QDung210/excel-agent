"""txt/csv sources -> SourceFacts with provenance + simple tables (deterministic, no LLM).

Every number that ends up in a generated workbook must trace back to one of these facts, a prompt number, or a
clarification answer (see grounding in spec_gen.py).
"""
from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from typing import Any

from . import upload
from .schemas import Provenance, SourceFact

CURRENCY_PATTERNS = [
    ("VND", re.compile(r"\bVND\b|\bVNĐ\b|₫|(?<=\d)\s?đ\b|\bđồng\b|\btriệu\b|\btỷ\b", re.I)),
    ("USD", re.compile(r"\bUSD\b|US\$|\$\s?\d", re.I)),
    ("EUR", re.compile(r"\bEUR\b|€", re.I)),
]
MONTHS_EN = "january|february|march|april|may|june|july|august|september|october|november|december"
PERIOD_RE = re.compile(
    rf"(tháng\s*\d{{1,2}}\s*[/\-]\s*\d{{4}}|tháng\s*\d{{1,2}}\s+năm\s+\d{{4}}|\b(?:q|quý\s*)[1-4]\s*[/\-]?\s*\d{{4}}"
    rf"|\b(?:{MONTHS_EN})\s+\d{{4}}\b|\bnăm\s+\d{{4}}\b|\bFY\s?\d{{2,4}}\b|\b\d{{1,2}}/\d{{4}}\b|\b\d{{4}}-\d{{2}}(?:-\d{{2}})?\b)",
    re.I)
NUM_TOKEN_RE = re.compile(
    r"(?<![\w.,])[-+]?(?:\$\s?)?\d[\d.,\s]*\d(?:\s?(?:%|tr\b|triệu|tỷ|k\b|m\b))?|(?<![\w.,])[-+]?\d(?:\s?%)?", re.I)
KV_RE = re.compile(r"^\s*[-*•]?\s*([^:=\t|]{1,80}?)\s*[:=]\s*(.+?)\s*$")
MONEY_WORDS = re.compile(r"giá|price|doanh\s*thu|revenue|amount|cost|chi\s*phí|tiền|thành\s*tiền|lương|salary|budget|"
                         r"ngân\s*sách|total|tổng|profit|lợi\s*nhuận|sales|bán", re.I)
QTY_WORDS = re.compile(r"số\s*lượng|quantity|qty|units?|sl\b|count", re.I)
UNIT_PRICE_WORDS = re.compile(r"đơn\s*giá|unit\s*price|price|giá", re.I)
AMOUNT_WORDS = re.compile(r"thành\s*tiền|amount|revenue|doanh\s*thu|total", re.I)
PERCENT_WORDS = re.compile(r"%|percent|tỷ\s*lệ|ratio|rate", re.I)
DATE_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$|^(\d{1,2})/(\d{1,2})/(\d{4})$")


def parse_number(text: Any) -> tuple[float | None, str | None]:
    """'15.000.000 đ' -> (15000000, 'VND'); '1,5' -> 1.5; '12%' -> (0.12, '%'); '2 tỷ' -> 2e9."""
    if isinstance(text, bool) or text is None:
        return None, None
    if isinstance(text, (int, float)):
        return float(text), None
    s = str(text).strip()
    if not s or len(s) > 40 or DATE_RE.match(s):      # '15/09/2026' is a date, never 15092026
        return None, None
    if re.match(r"0\d", s):                             # '0100000000' tax code, '0901 234 567' phone: keep the zero
        return None, None
    unit = None
    for code, pat in CURRENCY_PATTERNS:
        if pat.search(s):
            unit = code
    low = s.lower()
    mult = 1.0
    if low.endswith("%"):
        unit, mult = "%", 0.01
    elif re.search(r"\btỷ\b", low):
        mult = 1e9
    elif re.search(r"\b(triệu|tr)\b", low):
        mult = 1e6
    elif re.search(r"\d\s?k$", low):
        mult = 1e3
    digit_runs = re.findall(r"\d[\d.,]*", s)
    if len(digit_runs) > 1 and not re.search(r"\d{1,3}(?: \d{3})+", s):
        return None, None                                 # 'Quý 3/2026', 'A1-B2' are not one number
    core = re.sub(r"[^\d.,\-+]", "", s)
    if not re.search(r"\d", core):
        return None, None
    neg = core.startswith("-")
    core = core.lstrip("+-")
    if "." in core and "," in core:
        dec = "." if core.rfind(".") > core.rfind(",") else ","
        thou = "," if dec == "." else "."
        core = core.replace(thou, "").replace(dec, ".")
    elif "." in core or "," in core:
        sep = "." if "." in core else ","
        groups = core.split(sep)
        if len(groups) > 2 or (len(groups) == 2 and len(groups[1]) == 3 and mult == 1.0 and unit != "%"):
            if not all(len(g) == 3 for g in groups[1:]):
                return None, None
            core = "".join(groups)                        # thousands separators
        else:
            core = core.replace(sep, ".")                # decimal separator
    try:
        v = float(core) * mult
    except ValueError:
        return None, None
    return (-v if neg else v), unit


def detect_currency(*texts: str) -> str | None:
    found = []
    for t in texts:
        for code, pat in CURRENCY_PATTERNS:
            if t and pat.search(t) and code not in found:
                found.append(code)
    return found[0] if len(found) == 1 else None


def detect_period(*texts: str) -> str | None:
    for t in texts:
        m = PERIOD_RE.search(t or "")
        if m:
            return m.group(0).strip()
    return None


def _cell_value(raw: str) -> Any:
    s = (raw or "").strip()
    if s == "":
        return None
    m = DATE_RE.match(s)
    if m:
        return s
    v, _ = parse_number(s)
    if v is not None and re.fullmatch(r"[-+$€₫\d.,\s%]+(?:đ|VND|USD|EUR|tr|triệu|tỷ|k)?", s, re.I):
        if re.fullmatch(r"0\d+", s):                   # codes with leading zeros stay text
            return s
        return int(v) if float(v).is_integer() and abs(v) < 1e15 else v
    return s


@dataclass
class Table:
    file_id: str
    name: str
    headers: list[str]
    rows: list[list[Any]]
    locators: list[str]
    fact_ids: list[list[str | None]] = field(default_factory=list)


@dataclass
class SourceBundle:
    facts: list[SourceFact] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    texts: dict[str, str] = field(default_factory=dict)
    names: dict[str, str] = field(default_factory=dict)
    currency: str | None = None
    period: str | None = None
    warnings: list[str] = field(default_factory=list)

    def numeric_facts(self) -> list[SourceFact]:
        return [f for f in self.facts if isinstance(f.value, (int, float)) and not isinstance(f.value, bool)]

    def allowed_numbers(self) -> set[float]:
        return {round(float(f.value), 9) for f in self.numeric_facts()}

    def add(self, field_: str, value: Any, file_id: str | None, locator: str, quote: str | None = None,
            unit: str | None = None, status: str = "observed") -> SourceFact:
        f = SourceFact(id=f"f-{len(self.facts) + 1:03d}", field=field_[:80], value=value, unit=unit,
                       provenance=Provenance(file_id=file_id, locator=locator, quote=(quote or "")[:300] or None),
                       status=status)
        self.facts.append(f)
        return f


def _add_table(b: SourceBundle, file_id: str, name: str, grid: list[list[str]], first_line: int, kind: str) -> None:
    headers = [h.strip() or f"Column {i + 1}" for i, h in enumerate(grid[0])]
    t = Table(file_id, name, headers, [], [])
    for i, raw in enumerate(grid[1:], start=1):
        if not any((c or "").strip() for c in raw):
            continue
        raw = (raw + [""] * len(headers))[:len(headers)]
        row = [_cell_value(c) for c in raw]
        loc_row = f"{kind}:{first_line + i}"
        ids: list[str | None] = []
        for h, v, rc in zip(headers, row, raw):
            if isinstance(v, (int, float)):
                ids.append(b.add(h, v, file_id, f"{loc_row},col:{h}", rc.strip(), parse_number(rc)[1]).id)
            else:
                ids.append(None)
        t.rows.append(row)
        t.locators.append(loc_row)
        t.fact_ids.append(ids)
    if t.rows:
        b.tables.append(t)


def _sniff(text: str) -> str:
    sample = text[:4096]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def _read_csv(b: SourceBundle, file_id: str, name: str, text: str) -> None:
    rows = [r for r in csv.reader(io.StringIO(text), delimiter=_sniff(text))]
    rows = [r for r in rows if any(c.strip() for c in r)]
    if len(rows) < 2:
        b.warnings.append(f"{name}: CSV has no data rows.")
        return
    _add_table(b, file_id, name, rows, 1, "row")


def _split_delimited(line: str) -> list[str] | None:
    s = line.strip()
    if "|" in s and s.count("|") >= 1:
        cells = [c.strip() for c in s.strip("|").split("|")]
        return cells if len(cells) >= 2 else None
    if "\t" in s:
        cells = [c.strip() for c in s.split("\t")]
        return cells if len(cells) >= 2 else None
    return None


def _read_txt(b: SourceBundle, file_id: str, name: str, text: str) -> None:
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        cells = _split_delimited(lines[i])
        if cells:                                       # markdown / tab table block
            block, start = [], i
            while i < len(lines) and (c := _split_delimited(lines[i])):
                if not all(re.fullmatch(r":?-{2,}:?", x) for x in c if x):
                    block.append(c)
                i += 1
            if len(block) >= 2 and len({len(r) for r in block}) == 1:
                _add_table(b, file_id, f"{name} table {len(b.tables) + 1}", block, start + 1, "line")
                continue
            i = start
        line = lines[i]
        loc = f"line:{i + 1}"
        m = KV_RE.match(line)
        if m and not re.fullmatch(r"https?", m.group(1).strip(), re.I):
            label, raw = m.group(1).strip(), m.group(2).strip()
            v, unit = parse_number(raw)
            if v is not None and re.fullmatch(r"[-+$€₫\d.,\s%]+\s*(?:đ|đồng|VND|VNĐ|USD|EUR|tr|triệu|tỷ|k|%)?\s*"
                                              r"(?:[\wÀ-ỹ/]{0,12})?", raw, re.I):
                b.add(label, int(v) if v.is_integer() else v, file_id, loc, line.strip(), unit)
            else:
                b.add(label, raw[:300], file_id, loc, line.strip())
        else:
            for tok in NUM_TOKEN_RE.findall(line):
                v, unit = parse_number(tok)
                if v is not None and not re.fullmatch(r"(19|20)\d{2}", tok.strip()):
                    b.add(f"number on line {i + 1}", int(v) if v.is_integer() else v, file_id, loc, line.strip(), unit)
        i += 1


def prompt_numbers(b: SourceBundle, prompt: str) -> None:
    """Numbers typed by the user in the prompt count as user-provided facts."""
    for tok in NUM_TOKEN_RE.findall(prompt or ""):
        v, unit = parse_number(tok)
        if v is not None and not re.fullmatch(r"(19|20)\d{2}|\d{1,2}", tok.strip()):
            b.add("number in prompt", int(v) if v.is_integer() else v, None, "prompt", tok.strip(), unit,
                  status="user_confirmed")


def answer_figures(b: SourceBundle, text: str, keep_text: bool = False) -> None:
    for i, line in enumerate((text or "").splitlines(), start=1):
        m = KV_RE.match(line)
        if m:
            v, unit = parse_number(m.group(2))
            if v is not None:
                b.add(m.group(1).strip(), int(v) if v.is_integer() else v, None, f"answer:figures:{i}",
                      line.strip(), unit, status="user_confirmed")
            elif keep_text and m.group(2).strip():
                b.add(m.group(1).strip(), m.group(2).strip()[:300], None, f"answer:figures:{i}", line.strip(),
                      status="user_confirmed")


def load(file_ids: list[str], prompt: str = "") -> SourceBundle:
    b = SourceBundle()
    for fid in file_ids:
        meta = upload.get(fid)
        text = upload.read_text(fid)
        b.texts[fid], b.names[fid] = text, meta.filename
        if meta.kind == "csv":
            _read_csv(b, fid, meta.filename, text)
        else:
            _read_txt(b, fid, meta.filename, text)
    prompt_numbers(b, prompt)
    headers = " ".join(h for t in b.tables for h in t.headers)
    b.currency = detect_currency(prompt, headers, *b.texts.values())
    b.period = detect_period(prompt, *b.texts.values())
    return b
