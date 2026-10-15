import type {
  AgentTurnRequest,
  AgentTurnResponse,
  AnalyzeRequest,
  AnalyzeResponse,
  ApiErrorBody,
  ClarifyAnswers,
  Health,
  HistoryEntry,
  OperationBatch,
  Preview,
  RevisionResult,
  TemplateEntry,
  WorkbookMeta,
  WorkbookSummary,
  Workspace,
  WorkspaceFile,
  WorkspaceMode,
  FileRole,
} from "./types";

/** Error thrown for any non-2xx response. Carries the backend `{code, message}` body. */
export class ApiError extends Error {
  readonly status: number;
  readonly code: string;
  readonly body: Partial<ApiErrorBody>;

  constructor(status: number, code: string, message: string, body: Partial<ApiErrorBody> = {}) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
    this.body = body;
  }

  /** True when the route itself does not exist yet (backend phase not built), as opposed to a missing resource. */
  get routeMissing(): boolean {
    if (this.status === 405) return true;
    if (this.status !== 404) return false;
    // FastAPI default 404 ({detail}) or the SPA catch-all "Unknown API route."
    return this.code === "HTTP_404" || /unknown api route/i.test(this.message);
  }
}

export function errorMessage(err: unknown): string {
  if (err instanceof ApiError) {
    if (err.routeMissing) return "This feature is not available on the backend yet.";
    return err.message;
  }
  if (err instanceof TypeError) return "Cannot reach the API server. Is the backend running?";
  if (err instanceof Error) return err.message;
  return String(err);
}

async function parseError(res: Response): Promise<ApiError> {
  let body: Record<string, unknown> = {};
  try {
    body = await res.json();
  } catch {
    /* non-JSON error body */
  }
  if (typeof body.code === "string") {
    return new ApiError(res.status, body.code, String(body.message ?? res.statusText), body as ApiErrorBody);
  }
  // FastAPI default shape {detail: string | [...]}
  const detail = body.detail;
  const msg =
    typeof detail === "string"
      ? detail
      : Array.isArray(detail)
        ? detail.map((d: { msg?: string }) => d?.msg).filter(Boolean).join("; ")
        : res.statusText || `HTTP ${res.status}`;
  return new ApiError(res.status, `HTTP_${res.status}`, msg, { message: msg });
}

interface RequestOptions {
  signal?: AbortSignal;
}

async function request<T>(method: string, path: string, body?: unknown, opts: RequestOptions = {}): Promise<T> {
  const init: RequestInit = { method, signal: opts.signal, headers: { Accept: "application/json" } };
  if (body instanceof FormData) {
    init.body = body;
  } else if (body !== undefined) {
    init.body = JSON.stringify(body);
    (init.headers as Record<string, string>)["Content-Type"] = "application/json";
  }
  const res = await fetch(path, init);
  if (!res.ok) throw await parseError(res);
  if (res.status === 204) return undefined as T;
  return (await res.json()) as T;
}

const enc = encodeURIComponent;

export const api = {
  health: (o?: RequestOptions) => request<Health>("GET", "/api/health", undefined, o),

  // workspaces
  createWorkspace: (mode: WorkspaceMode, name?: string) =>
    request<Workspace>("POST", "/api/workspaces", name ? { mode, name } : { mode }),
  getWorkspace: (id: string) => request<Workspace>("GET", `/api/workspaces/${enc(id)}`),
  uploadWorkspaceFile: (id: string, file: File, role: FileRole) => {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("role", role);
    return request<WorkspaceFile>("POST", `/api/workspaces/${enc(id)}/files`, fd);
  },
  removeWorkspaceFile: (id: string, fileId: string) =>
    request<Workspace>("DELETE", `/api/workspaces/${enc(id)}/files/${enc(fileId)}`),
  analyze: (id: string, body: AnalyzeRequest) =>
    request<AnalyzeResponse>("POST", `/api/workspaces/${enc(id)}/analyze`, body),
  clarify: (id: string, body: ClarifyAnswers) =>
    request<AnalyzeResponse>("POST", `/api/workspaces/${enc(id)}/clarify`, body),
  generate: (id: string) => request<{ run_id: string }>("POST", `/api/workspaces/${enc(id)}/generate`, {}),

  // templates
  listTemplates: () => request<TemplateEntry[]>("GET", "/api/templates"),
  templatePreview: (id: string, o?: RequestOptions) =>
    request<Preview>("GET", `/api/templates/${enc(id)}/preview`, undefined, o),
  /** Read-only preview of an .xlsx attached to a workspace (e.g. an uploaded template). */
  workspaceFilePreview: (wsId: string, fileId: string, o?: RequestOptions) =>
    request<Preview>("GET", `/api/workspaces/${enc(wsId)}/files/${enc(fileId)}/preview`, undefined, o),

  // workbooks
  listWorkbooks: () => request<WorkbookSummary[]>("GET", "/api/workbooks"),
  uploadWorkbook: (file: File) => {
    const fd = new FormData();
    fd.append("file", file);
    return request<WorkbookMeta>("POST", "/api/workbooks/upload", fd);
  },
  metadata: (id: string, o?: RequestOptions) =>
    request<WorkbookMeta>("GET", `/api/workbooks/${enc(id)}/metadata`, undefined, o),
  preview: (id: string, revision?: number, o?: RequestOptions) =>
    request<Preview>(
      "GET",
      `/api/workbooks/${enc(id)}/preview${revision ? `?revision=${revision}` : ""}`,
      undefined,
      o,
    ),
  operations: (id: string, batch: OperationBatch) =>
    request<RevisionResult>("POST", `/api/workbooks/${enc(id)}/operations`, batch),
  undo: (id: string, base_revision: number, target_revision?: number) =>
    request<RevisionResult>(
      "POST",
      `/api/workbooks/${enc(id)}/undo`,
      target_revision ? { base_revision, target_revision } : { base_revision },
    ),
  history: (id: string) => request<HistoryEntry[]>("GET", `/api/workbooks/${enc(id)}/history`),
  exportUrl: (id: string) => `/api/workbooks/${enc(id)}/export`,

  // agent
  startTurn: (body: AgentTurnRequest, o?: RequestOptions) =>
    request<AgentTurnResponse>("POST", "/api/agent/turns", body, o),
  applyPlan: (planId: string, base_revision: number) =>
    request<RevisionResult>("POST", `/api/agent/plans/${enc(planId)}/apply`, { base_revision }),
  rejectPlan: (planId: string) => request<{ ok: boolean }>("POST", `/api/agent/plans/${enc(planId)}/reject`, {}),

  runEventsUrl: (runId: string) => `/api/runs/${enc(runId)}/events`,
  turnEventsUrl: (turnId: string) => `/api/agent/turns/${enc(turnId)}/events`,
};

export function newOperationId(): string {
  const uuid =
    typeof crypto !== "undefined" && "randomUUID" in crypto
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  return `ui-${uuid}`;
}
