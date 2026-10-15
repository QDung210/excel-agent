"""Pydantic v2 contracts (spec 08). TypeScript mirrors live in frontend/src/lib/types.ts."""
from __future__ import annotations

import re
from typing import Annotated, Any, Literal, Union

from openpyxl.utils import column_index_from_string
from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from . import config

A1_RE = re.compile(r"^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$")
SHEET_BAD = re.compile(r"[\[\]:*?/\\]")
MAX_ROW, MAX_COL = 1_048_576, 16_384


def parse_a1(ref: str) -> tuple[int, int]:
    """'B12' -> (12, 2). Raises ValueError outside Excel bounds."""
    m = A1_RE.match((ref or "").strip().upper())
    if not m:
        raise ValueError(f"invalid cell reference {ref!r}")
    col, row = column_index_from_string(m.group(1)), int(m.group(2))
    if col > MAX_COL or row > MAX_ROW:
        raise ValueError(f"cell {ref!r} outside Excel bounds")
    return row, col


def parse_range(ref: str) -> tuple[int, int, int, int]:
    """'A1:C3' or 'B2' -> (r1, c1, r2, c2)."""
    parts = (ref or "").strip().upper().split(":")
    if len(parts) == 1:
        r, c = parse_a1(parts[0])
        return r, c, r, c
    if len(parts) != 2:
        raise ValueError(f"invalid range {ref!r}")
    (r1, c1), (r2, c2) = parse_a1(parts[0]), parse_a1(parts[1])
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def check_sheet_name(name: str) -> str:
    name = (name or "").strip()
    if not name or len(name) > 31 or SHEET_BAD.search(name) or name.startswith("'") or name.endswith("'"):
        raise ValueError(f"invalid sheet name {name!r}")
    return name


class Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


# ---------------------------------------------------------------- create / clarify
class CreateWorkspace(Strict):
    mode: Literal["scratch", "template"]
    name: str | None = Field(default=None, max_length=120)


class ClarifyQuestion(Strict):
    id: str = Field(pattern=r"^[a-z0-9_]{1,40}$")
    type: Literal["select", "boolean", "text", "number"]
    question: str = Field(max_length=300)
    options: list[str] | None = None
    required: bool = False
    default: Any = None
    explanation: str | None = Field(default=None, max_length=400)


class AnalyzeRequest(Strict):
    prompt: str = Field(default="", max_length=8000)
    source_file_ids: list[str] = Field(default_factory=list, max_length=10)
    template_file_id: str | None = None
    template_id: str | None = None
    output_filename: str | None = Field(default=None, max_length=120)
    auto_approve: bool = False                    # skip the review of an AI-proposed structure (preview must be valid)


class MappingRow(BaseModel):
    """Template mode: where one input cell/column will take its value from (proposal shown before generating)."""
    target: str                                   # Sheet!B3 or Sheet!B9:B20
    field: str                                    # template label / column header
    source: str | None = None                     # 'Bên mua' (yeu_cau.txt line:2) | column 'Qty' (...) | None
    value: Any = None                             # value preview read from the source (never model text)
    via: str | None = None                        # label | header | llm | fuzzy | user | sequence
    status: Literal["auto", "confirm", "missing"]
    reason: str = ""


class AnalyzeResponse(BaseModel):
    status: Literal["ready", "needs_clarification", "blocked"]
    clarification_id: str | None = None
    questions: list[ClarifyQuestion] = Field(default_factory=list)
    reason: str | None = None
    plan: list[str] = Field(default_factory=list)
    mapping: list[MappingRow] = Field(default_factory=list)
    structure: dict[str, Any] | None = None         # template restructuring proposal (restructure.view)
    facts: int = 0
    warnings: list[str] = Field(default_factory=list)


class ExtractRequest(Strict):
    file_id: str | None = None
    template_id: str | None = None


class ClarifyAnswers(Strict):
    clarification_id: str
    answers: dict[str, Any] = Field(default_factory=dict)


# ---------------------------------------------------------------- source facts
class Provenance(Strict):
    file_id: str | None = None
    locator: str                                  # line:4 | row:3,col:Revenue | prompt | answer:currency
    quote: str | None = Field(default=None, max_length=300)


class SourceFact(Strict):
    id: str
    field: str
    value: Any = None
    unit: str | None = None
    provenance: Provenance
    status: Literal["observed", "deterministically_derived", "user_confirmed", "unknown"] = "observed"


# ---------------------------------------------------------------- workbook spec (scratch)
ColumnType = Literal["text", "number", "integer", "currency", "percent", "date"]
StylePreset = Literal["clean_finance", "minimal", "bold_header"]


class ColumnSpec(Strict):
    key: str = Field(pattern=r"^[A-Za-z_][A-Za-z0-9_]{0,40}$")
    header: str = Field(max_length=120)
    type: ColumnType = "text"
    width: float | None = Field(default=None, ge=4, le=80)
    formula: str | None = Field(default=None, max_length=400)   # row template, e.g. "={quantity}*{unit_price}"


class SheetSpec(Strict):
    name: str
    title: str | None = Field(default=None, max_length=200)
    subtitle: str | None = Field(default=None, max_length=300)
    columns: list[ColumnSpec] = Field(min_length=1, max_length=config.MAX_COLUMNS)
    rows: list[dict[str, Any]] = Field(default_factory=list, max_length=config.MAX_ROWS_PER_SHEET)
    blank_rows: int = Field(default=0, ge=0, le=500)          # empty input rows for templates without data
    totals: list[str] = Field(default_factory=list)            # column keys that get a SUM row
    formulas: dict[str, str] = Field(default_factory=dict)     # extra absolute formulas {"F2": "=..."}
    notes: list[str] = Field(default_factory=list, max_length=20)
    freeze_header: bool = True

    _check_name = field_validator("name")(lambda cls, v: check_sheet_name(v))

    @model_validator(mode="after")
    def _check(self) -> "SheetSpec":
        keys = [c.key for c in self.columns]
        if len(keys) != len(set(keys)):
            raise ValueError("duplicate column keys")
        unknown = set(self.totals) - set(keys)
        if unknown:
            raise ValueError(f"totals reference unknown columns {sorted(unknown)}")
        for row in self.rows:
            extra = set(row) - set(keys)
            if extra:
                raise ValueError(f"row has unknown columns {sorted(extra)}")
            for v in row.values():
                if isinstance(v, str) and len(v) > config.MAX_STRING:
                    raise ValueError("cell text too long")
        if len(self.formulas) > config.MAX_FORMULAS:
            raise ValueError("too many formulas")
        for ref in self.formulas:
            parse_a1(ref)
        return self


class WorkbookSpec(Strict):
    title: str = Field(max_length=200)
    sheets: list[SheetSpec] = Field(min_length=1, max_length=config.MAX_SHEETS)
    style_preset: StylePreset = "clean_finance"
    currency: str | None = Field(default=None, max_length=8)

    @model_validator(mode="after")
    def _unique(self) -> "WorkbookSpec":
        names = [s.name.lower() for s in self.sheets]
        if len(names) != len(set(names)):
            raise ValueError("duplicate sheet names")
        return self


# ---------------------------------------------------------------- operations
CellValue = Union[str, int, float, bool, None]


class SetValue(Strict):
    type: Literal["set_value"]
    sheet: str
    cell: str
    value: CellValue
    value_type: Literal["auto", "string", "number", "bool", "date"] = "auto"
    source_fact_id: str | None = None

    _c = field_validator("cell")(lambda cls, v: (parse_a1(v), v.strip().upper().replace("$", ""))[1])

    @field_validator("value")
    @classmethod
    def _len(cls, v: CellValue) -> CellValue:
        if isinstance(v, str) and len(v) > 32767:
            raise ValueError("Excel cells hold at most 32767 characters")
        if isinstance(v, float) and (v != v or v in (float("inf"), float("-inf"))):
            raise ValueError("NaN/Infinity cannot be stored in Excel")
        return v


class ClearValue(Strict):
    type: Literal["clear_value"]
    sheet: str
    cell: str

    _c = field_validator("cell")(lambda cls, v: (parse_a1(v), v.strip().upper().replace("$", ""))[1])


Action = Annotated[Union[SetValue, ClearValue], Field(discriminator="type")]


class OperationBatch(Strict):
    base_revision: int = Field(ge=0)
    operation_id: str = Field(min_length=6, max_length=80, pattern=r"^[A-Za-z0-9_\-:.]+$")
    actions: list[Action] = Field(min_length=1, max_length=config.MAX_ACTIONS_PER_BATCH)
    origin: Literal["manual", "agent", "system"] = "manual"
    note: str | None = Field(default=None, max_length=300)


class UndoRequest(Strict):
    base_revision: int = Field(ge=0)
    target_revision: int | None = Field(default=None, ge=1)


class ValidationSummary(BaseModel):
    passed: bool = Field(alias="pass", serialization_alias="pass")
    level: str = "L4"
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)
    checks: dict[str, int] = Field(default_factory=dict)

    model_config = ConfigDict(populate_by_name=True)


class RevisionResult(BaseModel):
    workbook_id: str
    revision: int
    parent_revision: int | None
    status: Literal["validated", "rejected"]
    changed_ranges: list[str] = Field(default_factory=list)
    validation: ValidationSummary
    preview_url: str
    idempotent_replay: bool = False


# ---------------------------------------------------------------- agent
class PlanApply(Strict):
    base_revision: int = Field(ge=1)


class AgentTurnRequest(Strict):
    workbook_id: str
    revision: int = Field(ge=1)
    message: str = Field(min_length=1, max_length=4000)
    thread_id: str | None = None
    selection: str | None = Field(default=None, max_length=80)     # "Sheet1!B2:C4"
    active_sheet: str | None = None
