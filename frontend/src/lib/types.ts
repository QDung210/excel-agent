// TypeScript mirror of docs/api-contract.md and workspace/schemas.py.

export type Id = string;
/** Backend timestamps are epoch seconds (float). */
export type EpochSeconds = number;

export type ErrorCode =
  | "UNSUPPORTED_FILE"
  | "NOT_FOUND"
  | "REVISION_CONFLICT"
  | "FILE_TOO_LARGE"
  | "INVALID_OPERATION"
  | "VALIDATION_FAILED"
  | "INVALID_REQUEST"
  | "MODEL_UNAVAILABLE"
  | (string & {});

export interface ApiErrorBody {
  code: ErrorCode;
  message: string;
  current_revision?: number;
  [extra: string]: unknown;
}

// ---------------------------------------------------------------- health
export interface Health {
  ok: boolean;
  llm_mode: "mock" | "azure";
  model: string | null;
  max_upload_mb: number;
}

// ---------------------------------------------------------------- workspaces
export type WorkspaceMode = "scratch" | "template";
export type FileRole = "source" | "template";
export type FileKind = "txt" | "csv" | "xlsx";

export interface WorkspaceFile {
  id: Id;
  filename: string;
  kind: FileKind;
  size: number;
  role: FileRole;
}

export interface Workspace {
  id: Id;
  mode: WorkspaceMode;
  name: string | null;
  files: WorkspaceFile[];
  template_id: Id | null;
  prompt: string | null;
  clarification: unknown;
  workbook_id: Id | null;
  runs: Id[];
}

export type QuestionType = "select" | "boolean" | "text" | "number";

export interface ClarifyQuestion {
  id: string;
  type: QuestionType;
  question: string;
  options?: string[] | null;
  required: boolean;
  default?: unknown;
  explanation?: string | null;
}

export interface AnalyzeRequest {
  prompt: string;
  source_file_ids?: Id[];
  template_file_id?: Id | null;
  template_id?: Id | null;
  output_filename?: string | null;
  /** Skip the review of an AI-proposed structure; the backend still refuses a plan whose dry run is invalid. */
  auto_approve?: boolean;
}

export interface AnalyzeResponse {
  status: "ready" | "needs_clarification" | "blocked";
  clarification_id: Id | null;
  questions: ClarifyQuestion[];
  reason: string | null;
  plan: string[];
  mapping?: MappingRow[];
  structure?: StructureView | null;
  facts: number;
  warnings: string[];
}

export type StructureChange =
  | "add" | "rename" | "to_input" | "formula" | "to_formula" | "default" | "keep" | "hidden" | "move" | "sequence" | "total";

/** Where a previewed value comes from: quoted, computed by a formula, assumed (an implied 0) or missing. */
export type Provenance = "explicit" | "derived" | "assumed" | "missing";

export interface PlanCell {
  v: unknown;
  f?: string;
  prov: Provenance;
  pending?: boolean;
  src?: string | null;
  quote?: string | null;
  note?: string | null;
}

export interface PlanRow {
  row: number;
  cells: PlanCell[];
}

/** The dry run: the plan applied to a copy of the template and read back, i.e. what the file will contain. */
export interface StructurePreview {
  ok: boolean;
  errors: string[];
  code?: string;
  applied?: boolean;
  columns?: { letter: string; col: number; header: string; hidden: boolean; action: string }[];
  rows?: PlanRow[];
  totals?: PlanRow | null;
  sheets?: {
    name: string;
    purpose: string;
    reason: string;
    columns: { letter: string; header: string; action: string; formula: string | null }[];
    rows: PlanRow[];
    totals: PlanRow | null;
  }[];
  other_formulas?: Record<string, { f: string; v: unknown }>;
}

/** Template mode with prose sources: the column structure the AI proposes for the template's table. */
export interface StructureView {
  sheet: string;
  applied: boolean;
  columns: {
    letter: string;
    header: string;
    from: string | null;
    changes: StructureChange[];
    kind: "formula" | "input" | "sequence";
    formula: string | null;
    fmt: string | null;
    default: unknown;
    total?: boolean;
    reason: string;
  }[];
  removed: { header: string; mode: "hide" | "delete" | "keep"; used_by: string[] }[];
  template_columns?: { letter: string; header: string; kind: "formula" | "input"; formula: string | null }[];
  kv: { label: string; cell: string; value: unknown; source: string | null; quote?: string | null }[];
  rows: Record<string, unknown>[];
  row_count: number;
  mode: string;
  understanding?: {
    template_purpose: string;
    document_summary: string;
    report_intent: string;
    fit: "fits" | "partial" | "mismatch";
    fit_reason: string;
    entities: string[];
    derivable: string[];
    missing: string[];
    assumptions: string[];
  } | null;
  labels?: { cell: string; current: string; new: string; reason: string }[];
  sheets_proposed?: number;
  sheets_on?: boolean;
  missing?: { kind: "field" | "column" | "ai"; what: string }[];
  overflow?: number;
  approved?: boolean;
  valid?: boolean;
  preview?: StructurePreview | Record<string, never>;
}

/** Template mode: where one input cell/column takes its value from (values are read from sources, never invented). */
export interface MappingRow {
  target: string;
  field: string;
  source: string | null;
  value: unknown;
  via: string | null;
  status: "auto" | "confirm" | "missing";
  reason: string;
}

export type ClarifyAnswerValue = string | number | boolean | null;

export interface ClarifyAnswers {
  clarification_id: Id;
  answers: Record<string, ClarifyAnswerValue>;
}

export interface TemplateEntry {
  id: Id;
  name: string;
  description: string | null;
  origin: "builtin" | "user";
  sheets: string[];
  created: EpochSeconds;
}

// ---------------------------------------------------------------- runs / SSE
export interface SseEnvelope<T = Record<string, unknown>> {
  event: string;
  run_id: Id;
  seq: number;
  data: T;
}

export interface Run {
  id: Id;
  kind: string;
  status: "running" | "done" | "error" | "cancelled";
  result: unknown;
  events: SseEnvelope[];
  created: EpochSeconds;
}

export interface ToolStartedData {
  tool: string;
  args?: Record<string, unknown>;
}
export interface ToolCompletedData {
  tool: string;
  status: string;
  summary?: string;
}
export interface WorkbookRevisionData {
  workbook_id: Id;
  revision: number;
  changed_ranges: string[];
}
export interface RunDoneData {
  workbook_id?: Id;
  revision?: number;
  warnings?: string[];
  text?: string;
}
export interface RunErrorData {
  code: string;
  message: string;
}

// ---------------------------------------------------------------- workbooks
export interface RevisionEntry {
  revision: number;
  parent: number | null;
  created: EpochSeconds;
  origin: string;
  note: string | null;
  changed_count: number;
  restored: number | null;
}

export interface WorkbookSummary {
  id: Id;
  name: string;
  filename: string;
  updated: EpochSeconds;
  current_revision: number;
  origin: string;
}

export interface WorkbookMeta extends WorkbookSummary {
  created: EpochSeconds;
  workspace_id: Id | null;
  sheets: string[];
  llm_mode: string | null;
  warnings: string[] | null;
  revisions: RevisionEntry[];
  can_undo: boolean;
}

/** FortuneSheet Cell subset produced by the backend preview. */
export interface PreviewCell {
  v?: string | number | boolean | null;
  m?: string | number | null;
  f?: string;
  ct?: { fa?: string; t?: string; s?: Array<{ v?: string }> };
  bl?: number;
  it?: number;
  un?: number;
  cl?: number;
  fs?: number;
  ff?: string;
  fc?: string;
  bg?: string;
  ht?: number;
  vt?: number;
  tb?: number | string;
  mc?: { r: number; c: number; rs?: number; cs?: number };
}

export interface PreviewSheet {
  name: string;
  id?: string;
  index?: number | string;
  order: number;
  status: number;
  hide?: number;
  celldata: Array<{ r: number; c: number; v: PreviewCell | null }>;
  config: {
    merge?: Record<string, unknown>;
    borderInfo?: unknown[];
    columnlen?: Record<string, number>;
    rowlen?: Record<string, number>;
    rowhidden?: Record<string, number>;
    colhidden?: Record<string, number>;
  };
  frozen?: { type: string; range?: { row_focus: number; column_focus: number } };
  row: number;
  column: number;
  warnings?: string[];
  dimensions?: string;
}

export interface Preview {
  workbook_id: Id;
  revision: number;
  sheets: PreviewSheet[];
  warnings: string[];
  stale_formulas: number;
  computed_formulas?: number;
}

export type CellValue = string | number | boolean | null;

export type Action =
  | {
      type: "set_value";
      sheet: string;
      cell: string;
      value: CellValue;
      value_type?: "auto" | "string" | "number" | "bool" | "date";
      source_fact_id?: string;
    }
  | { type: "clear_value"; sheet: string; cell: string };

export interface OperationBatch {
  base_revision: number;
  operation_id: string;
  origin: "manual" | "agent" | "system";
  note?: string;
  actions: Action[];
}

export interface ValidationSummary {
  pass: boolean;
  level: string;
  warnings: string[];
  errors: string[];
  checks: Record<string, number>;
}

export interface RevisionResult {
  workbook_id: Id;
  revision: number;
  parent_revision: number | null;
  status: "validated" | "rejected";
  changed_ranges: string[];
  validation: ValidationSummary;
  preview_url: string;
  idempotent_replay: boolean;
}

export interface HistoryEntry {
  t?: EpochSeconds;
  event: string;
  revision: number;
  origin?: string;
  note?: string | null;
  [extra: string]: unknown;
}

// ---------------------------------------------------------------- agent
export interface AgentTurnRequest {
  workbook_id: Id;
  revision: number;
  message: string;
  thread_id?: Id | null;
  selection?: string | null;
  active_sheet?: string | null;
}

export interface AgentTurnResponse {
  turn_id: Id;
  run_id: Id;
  thread_id: Id;
}

export interface DiffRow {
  cell: string;
  before: unknown;
  after: unknown;
}

export interface ConfirmationRequiredData {
  plan_id: Id;
  base_revision?: number;
  actions: Action[];
  diff: DiffRow[];
  reason?: string;
}
