"""Small, safe Excel formula evaluator (no eval/exec).

Used to (a) give XlsxWriter-generated formulas a real cached value and (b) show values in the preview for formulas
whose cache was stripped after an edit. Anything outside the supported subset returns UNSUPPORTED, and callers then
leave the value empty instead of guessing. Excel recalculates everything on open (fullCalcOnLoad), so this is a
display aid, not the source of truth.
"""
from __future__ import annotations

import math
import re
from typing import Any, Callable

from openpyxl.utils import column_index_from_string

UNSUPPORTED = object()
MAX_RANGE_CELLS = 200_000


class XLError(Exception):
    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


class _Unsupported(Exception):
    pass


TOKEN_RE = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?(?![A-Za-z0-9_(]))
  | (?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)
  | (?P<func>(?:_xlfn\.|_xlws\.)?[A-Za-z][A-Za-z0-9_.]*(?=\())
  | (?P<bool>TRUE|FALSE)(?![A-Za-z0-9_(])
  | (?P<op><>|<=|>=|[-+*/^&=<>%(),:])
""", re.X | re.I)

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


class Range(list):
    """Flattened values of a range (row-major)."""


Getter = Callable[[str, int, int], Any]       # (sheet, row1, col1) -> value or None


def _tokens(src: str) -> list[tuple[str, str]]:
    out, pos = [], 0
    while pos < len(src):
        m = TOKEN_RE.match(src, pos)
        if not m:
            raise _Unsupported(f"unexpected {src[pos:pos + 10]!r}")
        pos = m.end()
        kind = m.lastgroup
        if kind != "ws":
            out.append((kind, m.group()))
    return out


def _split_ref(ref: str, default_sheet: str) -> tuple[str, str]:
    if "!" in ref:
        sheet, addr = ref.rsplit("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        return sheet, addr
    return default_sheet, ref


def _cell(addr: str) -> tuple[int, int]:
    m = CELL_RE.match(addr)
    if not m:
        raise _Unsupported(addr)
    return int(m.group(2)), column_index_from_string(m.group(1).upper())


# ---------------------------------------------------------------- coercion helpers
def _num(v: Any) -> float:
    if isinstance(v, XLError):
        raise v
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except ValueError:
            raise XLError("#VALUE!") from None
    raise _Unsupported(type(v).__name__)


def _scalar(v: Any) -> Any:
    if isinstance(v, Range):
        if len(v) == 1:
            return v[0]
        raise _Unsupported("range used as scalar")
    return v


def _text(v: Any) -> str:
    v = _scalar(v)
    if isinstance(v, XLError):
        raise v
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _flat_numbers(args: list[Any], count_text: bool = False) -> list[float]:
    out: list[float] = []
    for a in args:
        if isinstance(a, Range):
            for v in a:
                if isinstance(v, XLError):
                    raise v
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(float(v))
        else:
            a = _scalar(a)
            if a is None:
                continue
            out.append(_num(a))
    return out


def _criteria(crit: Any) -> Callable[[Any], bool]:
    crit = _scalar(crit)
    if isinstance(crit, (int, float)) and not isinstance(crit, bool):
        return lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == float(crit)
    s = _text(crit)
    m = re.match(r"^(<>|<=|>=|=|<|>)?(.*)$", s, re.S)
    op, rhs = m.group(1) or "=", m.group(2)
    try:
        rnum: float | None = float(rhs)
    except ValueError:
        rnum = None
    if "*" in rhs or "?" in rhs:
        raise _Unsupported("wildcard criteria")

    def test(v: Any) -> bool:
        if rnum is not None and isinstance(v, (int, float)) and not isinstance(v, bool):
            a, b = float(v), rnum
        else:
            a, b = ("" if v is None else str(v)).lower(), rhs.lower()
            if op not in ("=", "<>"):
                return False
        return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]
    return test


def _round(x: float, n: int) -> float:
    q = 10.0 ** n
    return math.floor(abs(x) * q + 0.5) / q * (1 if x >= 0 else -1)


def _bool(v: Any) -> bool:
    v = _scalar(v)
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        raise XLError("#VALUE!")
    return bool(_num(v))


def _fn(name: str, args: list[Any]) -> Any:
    n = _canon(name)
    if n in ("AND", "OR"):
        vals = []
        for a in args:
            if isinstance(a, Range):
                vals.extend(_bool(v) for v in a if isinstance(v, (bool, int, float)))
            else:
                vals.append(_bool(a))
        if not vals:
            raise XLError("#VALUE!")
        return all(vals) if n == "AND" else any(vals)
    if n == "NOT":
        return not _bool(args[0])
    if n in ("SUMIFS", "COUNTIFS", "MINIFS", "MAXIFS", "AVERAGEIFS"):
        pairs = args if n == "COUNTIFS" else args[1:]
        if len(pairs) % 2 or not pairs:
            raise _Unsupported("criteria pairs")
        ranges = [p if isinstance(p, Range) else Range([_scalar(p)]) for p in pairs[0::2]]
        tests = [_criteria(c) for c in pairs[1::2]]
        size = len(ranges[0])
        if any(len(r) != size for r in ranges):
            raise _Unsupported("criteria shape")
        hits = [i for i in range(size) if all(t(r[i]) for r, t in zip(ranges, tests))]
        if n == "COUNTIFS":
            return float(len(hits))
        target = args[0]
        if not isinstance(target, Range) or len(target) != size:
            raise _Unsupported("target shape")
        nums = [float(target[i]) for i in hits if isinstance(target[i], (int, float)) and not isinstance(target[i], bool)]
        if n == "SUMIFS":
            return float(sum(nums))
        if n == "AVERAGEIFS":
            if not nums:
                raise XLError("#DIV/0!")
            return sum(nums) / len(nums)
        if not nums:
            return 0.0
        return min(nums) if n == "MINIFS" else max(nums)
    if n == "SUM":
        return sum(_flat_numbers(args))
    if n in ("AVERAGE", "MIN", "MAX", "COUNT"):
        nums = _flat_numbers(args)
        if n == "COUNT":
            return float(len(nums))
        if not nums:
            if n == "AVERAGE":
                raise XLError("#DIV/0!")
            return 0.0
        return {"AVERAGE": sum(nums) / len(nums), "MIN": min(nums), "MAX": max(nums)}[n]
    if n == "COUNTA":
        total = 0
        for a in args:
            vals = a if isinstance(a, Range) else [_scalar(a)]
            total += sum(1 for v in vals if v is not None and v != "")
        return float(total)
    if n == "ROUND":
        return _round(_num(_scalar(args[0])), int(_num(_scalar(args[1])) if len(args) > 1 else 0))
    if n == "ABS":
        return abs(_num(_scalar(args[0])))
    if n == "IF":
        cond = _scalar(args[0])
        if isinstance(cond, XLError):
            raise cond
        truthy = bool(_num(cond)) if not isinstance(cond, str) else (cond.upper() == "TRUE")
        if truthy:
            return _scalar(args[1]) if len(args) > 1 else True
        return _scalar(args[2]) if len(args) > 2 else False
    if n == "IFERROR":
        return _scalar(args[0])                      # errors are caught by the caller (see _call)
    if n in ("SUMIF", "COUNTIF", "AVERAGEIF"):
        rng, test = args[0], _criteria(args[1])
        if not isinstance(rng, Range):
            rng = Range([_scalar(rng)])
        if n == "COUNTIF":
            return float(sum(1 for v in rng if test(v)))
        target = args[2] if len(args) > 2 else rng
        if not isinstance(target, Range) or len(target) != len(rng):
            raise _Unsupported("SUMIF shape")
        picked = [t for v, t in zip(rng, target) if test(v) and isinstance(t, (int, float)) and not isinstance(t, bool)]
        if n == "SUMIF":
            return float(sum(picked))
        if not picked:
            raise XLError("#DIV/0!")
        return sum(picked) / len(picked)
    raise _Unsupported(f"function {n}")


SUPPORTED_FUNCTIONS = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "ROUND", "ABS", "IF", "IFERROR",
                       "SUMIF", "COUNTIF", "AVERAGEIF", "AND", "OR", "NOT", "SUMIFS", "COUNTIFS", "MINIFS", "MAXIFS",
                       "AVERAGEIFS"}


def _canon(name: str) -> str:
    n = name.upper()
    for prefix in ("_XLFN.", "_XLWS."):
        if n.startswith(prefix):
            n = n[len(prefix):]
    return n


class _Parser:
    def __init__(self, toks: list[tuple[str, str]], sheet: str, get: Getter):
        self.t, self.i, self.sheet, self.get = toks, 0, sheet, get

    def peek(self) -> tuple[str, str] | None:
        return self.t[self.i] if self.i < len(self.t) else None

    def take(self, val: str | None = None) -> tuple[str, str]:
        tok = self.peek()
        if tok is None or (val is not None and tok[1] != val):
            raise _Unsupported(f"expected {val}")
        self.i += 1
        return tok

    def parse(self) -> Any:
        v = self.compare()
        if self.peek() is not None:
            raise _Unsupported("trailing tokens")
        return v

    def compare(self) -> Any:
        a = self.concat()
        while (tok := self.peek()) and tok[1] in ("=", "<>", "<", ">", "<=", ">="):
            self.take()
            b = self.concat()
            a, b = _scalar(a), _scalar(b)
            for x in (a, b):
                if isinstance(x, XLError):
                    raise x
            if isinstance(a, str) or isinstance(b, str):
                x, y = _text(a).lower(), _text(b).lower()
            else:
                x, y = _num(a), _num(b)
            a = {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[tok[1]]
        return a

    def concat(self) -> Any:
        a = self.additive()
        while (tok := self.peek()) and tok[1] == "&":
            self.take()
            a = _text(a) + _text(self.additive())
        return a

    def additive(self) -> Any:
        a = self.term()
        while (tok := self.peek()) and tok[1] in ("+", "-"):
            self.take()
            b = self.term()
            a = _num(_scalar(a)) + _num(_scalar(b)) if tok[1] == "+" else _num(_scalar(a)) - _num(_scalar(b))
        return a

    def term(self) -> Any:
        a = self.power()
        while (tok := self.peek()) and tok[1] in ("*", "/"):
            self.take()
            b = self.power()
            if tok[1] == "*":
                a = _num(_scalar(a)) * _num(_scalar(b))
            else:
                d = _num(_scalar(b))
                if d == 0:
                    raise XLError("#DIV/0!")
                a = _num(_scalar(a)) / d
        return a

    def power(self) -> Any:
        a = self.unary()
        while (tok := self.peek()) and tok[1] == "^":
            self.take()
            a = _num(_scalar(a)) ** _num(_scalar(self.unary()))
        return a

    def unary(self) -> Any:
        tok = self.peek()
        if tok and tok[1] in ("-", "+"):
            self.take()
            v = _num(_scalar(self.unary()))
            return -v if tok[1] == "-" else v
        v = self.primary()
        while (tok := self.peek()) and tok[1] == "%":
            self.take()
            v = _num(_scalar(v)) / 100
        return v

    def primary(self) -> Any:
        tok = self.peek()
        if tok is None:
            raise _Unsupported("unexpected end")
        kind, text = tok
        if kind == "num":
            self.take()
            return float(text)
        if kind == "str":
            self.take()
            return text[1:-1].replace('""', '"')
        if kind == "bool":
            self.take()
            return text.upper() == "TRUE"
        if kind == "ref":
            self.take()
            return self.ref(text)
        if kind == "func":
            self.take()
            return self.call(text)
        if text == "(":
            self.take()
            v = self.compare()
            self.take(")")
            return v
        raise _Unsupported(f"token {text}")

    def ref(self, text: str) -> Any:
        sheet, addr = _split_ref(text, self.sheet)
        if ":" in addr:
            a, b = addr.split(":")
            (r1, c1), (r2, c2) = _cell(a), _cell(b)
            r1, r2, c1, c2 = min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)
            if (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS:
                raise _Unsupported("range too large")
            return Range(self.get(sheet, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
        r, c = _cell(addr)
        return self.get(sheet, r, c)

    def call(self, name: str) -> Any:
        if _canon(name) not in SUPPORTED_FUNCTIONS:
            raise _Unsupported(f"function {name}")
        self.take("(")
        args: list[Any] = []
        if self.peek() and self.peek()[1] == ")":
            self.take(")")
            return _fn(name, args)
        if _canon(name) == "IFERROR":
            start = self.i
            try:
                first = _scalar(self.compare())
                if isinstance(first, XLError):
                    raise first
                ok = True
            except XLError:
                ok, first = False, None
                self.i = start
                self._skip_arg()
            self.take(",")
            fallback = self.compare()
            self.take(")")
            return first if ok else _scalar(fallback)
        while True:
            args.append(self.compare())
            tok = self.take()
            if tok[1] == ")":
                break
            if tok[1] != ",":
                raise _Unsupported("bad argument list")
        return _fn(name, args)

    def _skip_arg(self) -> None:
        depth = 0
        while (tok := self.peek()) is not None:
            if tok[1] == "(":
                depth += 1
            elif tok[1] == ")":
                if depth == 0:
                    return
                depth -= 1
            elif tok[1] == "," and depth == 0:
                return
            self.i += 1


def functions_used(formula: str) -> set[str]:
    try:
        return {_canon(t) for k, t in _tokens(formula.lstrip("=")) if k == "func"}
    except _Unsupported:
        return {"?"}


def evaluate(formula: str, sheet: str, get: Getter) -> Any:
    """Returns float/str/bool, an Excel error string like '#DIV/0!', or UNSUPPORTED."""
    try:
        v = _Parser(_tokens(formula.lstrip("=")), sheet, get).parse()
        v = _scalar(v)
        if isinstance(v, XLError):
            return v.code
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return "#NUM!"
        return v
    except XLError as e:
        return e.code
    except (_Unsupported, IndexError, OverflowError, RecursionError, ValueError, ZeroDivisionError):
        return UNSUPPORTED


class Workbook:
    """Lazily evaluates formulas over a {sheet: {(row, col): value | "=formula"}} grid with cycle protection."""

    def __init__(self, cells: dict[str, dict[tuple[int, int], Any]], formulas: dict[str, dict[tuple[int, int], str]]):
        self.cells, self.formulas = cells, formulas
        self.done: dict[tuple[str, int, int], Any] = {}
        self.active: set[tuple[str, int, int]] = set()

    def get(self, sheet: str, r: int, c: int) -> Any:
        key = (sheet, r, c)
        f = self.formulas.get(sheet, {}).get((r, c))
        if f is None:
            return self.cells.get(sheet, {}).get((r, c))
        if key in self.done:
            v = self.done[key]
        elif key in self.active:
            raise _Unsupported("circular reference")
        else:
            self.active.add(key)
            try:
                v = evaluate(f, sheet, self.get)
            finally:
                self.active.discard(key)
            self.done[key] = v
        if v is UNSUPPORTED:
            raise _Unsupported("dependency unsupported")
        if isinstance(v, str) and v.startswith("#"):
            return XLError(v)
        return v

    def value(self, sheet: str, r: int, c: int) -> Any:
        try:
            v = self.get(sheet, r, c)
        except _Unsupported:
            return UNSUPPORTED
        except XLError as e:
            return e.code
        return v.code if isinstance(v, XLError) else v
