import { useCallback, useEffect, useRef, useState, type DragEvent, type ReactNode } from "react";
import { Link, Navigate, useNavigate, useParams } from "react-router-dom";
import { Group, Panel, Separator } from "react-resizable-panels";
import {
  ArrowLeft,
  ArrowUp,
  Check,
  CircleAlert,
  Eye,
  FileText,
  Loader2,
  MessagesSquare,
  OctagonX,
  Paperclip,
  Play,
  RotateCcw,
  Sparkles,
  X,
} from "lucide-react";
import { toast } from "sonner";
import { AppMark } from "@/components/AppMark";
import { LlmBadge } from "@/components/LlmBadge";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { ClarifyForm } from "@/components/wizard/ClarifyForm";
import { MappingTable } from "@/components/wizard/MappingTable";
import { StructurePlan } from "@/components/wizard/StructurePlan";
import { RunTimeline } from "@/components/wizard/RunTimeline";
import { PreviewGrid } from "@/components/workspace/PreviewGrid";
import { TemplatePicker } from "@/components/workspace/TemplatePicker";
import { useHealth } from "@/hooks/useHealth";
import { useMediaQuery } from "@/hooks/useMediaQuery";
import { ApiError, api, errorMessage } from "@/lib/api";
import { openEventStream, type StreamHandle } from "@/lib/sse";
import { cn } from "@/lib/utils";
import type {
  AnalyzeResponse,
  ClarifyAnswerValue,
  Preview,
  SseEnvelope,
  TemplateEntry,
  Workspace,
  WorkspaceFile,
  WorkspaceMode,
} from "@/lib/types";

/** One item of the create conversation. Every entry mirrors a real backend call or result. */
type Entry =
  | { id: number; kind: "user"; text: string; files: string[]; template: string | null }
  | { id: number; kind: "analysis"; res: AnalyzeResponse; ver: number }
  | { id: number; kind: "answers"; lines: string[] }
  | { id: number; kind: "note"; text: string; tone: "info" | "error" }
  | { id: number; kind: "run"; events: SseEnvelope[]; live: boolean; failed: string | null };

type NewEntry = Entry extends infer E ? (E extends Entry ? Omit<E, "id"> : never) : never;

type PreviewState = { key: string; name: string; data: Preview | null; loading: boolean; error: string | null };

const isMod = (e: React.KeyboardEvent) => (e.ctrlKey || e.metaKey) && e.key === "Enter";

function answerText(v: ClarifyAnswerValue): string {
  if (v === null || v === "") return "(blank)";
  if (typeof v === "boolean") return v ? "Yes" : "No";
  return String(v);
}

export default function Studio() {
  const { mode } = useParams<{ mode: string }>();
  if (mode !== "scratch" && mode !== "template") return <Navigate to="/" replace />;
  return <StudioInner key={mode} mode={mode} />;
}

function StudioInner({ mode }: { mode: WorkspaceMode }) {
  const navigate = useNavigate();
  const { health } = useHealth();
  const wide = useMediaQuery("(min-width: 900px)");

  // ------------------------------------------------------------ workspace
  const [ws, setWs] = useState<Workspace | null>(null);
  const [wsError, setWsError] = useState<string | null>(null);
  const created = useRef(false);

  const createWs = useCallback(() => {
    setWsError(null);
    api.createWorkspace(mode).then(setWs, (e) => {
      const msg = errorMessage(e);
      setWsError(msg);
      toast.error(`Could not create workspace: ${msg}`);
    });
  }, [mode]);

  useEffect(() => {
    if (created.current) return; // StrictMode double-invoke guard
    created.current = true;
    createWs();
  }, [createWs]);

  // ------------------------------------------------------------ conversation log
  const [log, setLog] = useState<Entry[]>([]);
  const nextId = useRef(1);
  const push = useCallback((e: NewEntry) => {
    const id = nextId.current++;
    setLog((l) => [...l, { ...e, id } as Entry]);
    return id;
  }, []);

  /** Bumped whenever the template or sources change, so an older plan can no longer be generated. */
  const [inputsVer, setInputsVer] = useState(0);
  const hasPlan = log.some((e) => e.kind === "analysis");
  /** Requests sent since the inputs last changed; follow-ups refine an AI structure proposal instead of replacing it. */
  const prompts = useRef<string[]>([]);
  const inputsChanged = useCallback(
    (what: string) => {
      setInputsVer((v) => v + 1);
      prompts.current = [];
      if (hasPlan) push({ kind: "note", tone: "info", text: `${what} Send your request again so I can re-plan.` });
    },
    [hasPlan, push],
  );

  // ------------------------------------------------------------ templates + left preview
  const [templates, setTemplates] = useState<TemplateEntry[] | null>(null);
  const [galleryAvailable, setGalleryAvailable] = useState<boolean | null>(null);
  const [templateId, setTemplateId] = useState<string | null>(null);
  const [templateFile, setTemplateFile] = useState<WorkspaceFile | null>(null);
  const [uploadingTemplate, setUploadingTemplate] = useState(false);
  const [pv, setPv] = useState<PreviewState | null>(null);
  const pvCtl = useRef<AbortController | null>(null);
  /** Read-only previews of the gallery templates, drawn as thumbnails in the picker. */
  const [thumbs, setThumbs] = useState<Record<string, Preview | null | undefined>>({});

  useEffect(() => {
    if (mode !== "template") return;
    const ctl = new AbortController();
    api.listTemplates().then(
      (list) => {
        setTemplates(list);
        setGalleryAvailable(true);
        for (const t of list)
          api.templatePreview(t.id, { signal: ctl.signal }).then(
            (p) => setThumbs((m) => ({ ...m, [t.id]: p })),
            () => !ctl.signal.aborted && setThumbs((m) => ({ ...m, [t.id]: null })),
          );
      },
      (e) => {
        // No gallery route -> only upload, never a fake gallery.
        setGalleryAvailable(false);
        if (!(e instanceof ApiError && (e.routeMissing || e.status === 404))) toast.error(`Templates: ${errorMessage(e)}`);
      },
    );
    return () => ctl.abort();
  }, [mode]);

  const loadPreview = useCallback((key: string, name: string, fetcher: (signal: AbortSignal) => Promise<Preview>) => {
    pvCtl.current?.abort();
    const ctl = new AbortController();
    pvCtl.current = ctl;
    setPv({ key, name, data: null, loading: true, error: null });
    fetcher(ctl.signal).then(
      (data) => !ctl.signal.aborted && setPv({ key, name, data, loading: false, error: null }),
      (e) => !ctl.signal.aborted && setPv({ key, name, data: null, loading: false, error: errorMessage(e) }),
    );
  }, []);
  useEffect(() => () => pvCtl.current?.abort(), []);

  const previewBuiltin = useCallback(
    (t: TemplateEntry) => {
      const cached = thumbs[t.id];
      loadPreview(`tpl:${t.id}`, t.name, (signal) => (cached ? Promise.resolve(cached) : api.templatePreview(t.id, { signal })));
    },
    [loadPreview, thumbs],
  );
  const previewUploaded = useCallback(
    (wsId: string, f: WorkspaceFile) =>
      loadPreview(`file:${f.id}`, f.filename, (signal) => api.workspaceFilePreview(wsId, f.id, { signal })),
    [loadPreview],
  );

  const chooseBuiltin = async (t: TemplateEntry) => {
    if (!ws || templateId === t.id) return;
    if (templateFile) {
      // An attached upload would take precedence on the backend: detach it (the stored upload is kept).
      try {
        await api.removeWorkspaceFile(ws.id, templateFile.id);
      } catch (e) {
        return void toast.error(`Could not switch template: ${errorMessage(e)}`);
      }
      setTemplateFile(null);
    }
    setTemplateId(t.id);
    previewBuiltin(t);
    inputsChanged(`Template switched to “${t.name}”.`);
  };

  const uploadTemplate = async (f: File) => {
    if (!ws) return;
    if (!/\.xlsx$/i.test(f.name)) return void toast.error("Templates must be .xlsx files.");
    if (health && f.size > health.max_upload_mb * 1024 * 1024)
      return void toast.error(`${f.name} is larger than ${health.max_upload_mb} MB.`);
    setUploadingTemplate(true);
    try {
      const wf = await api.uploadWorkspaceFile(ws.id, f, "template");
      setTemplateFile(wf);
      setTemplateId(null);
      previewUploaded(ws.id, wf);
      inputsChanged(`Template switched to “${wf.filename}”.`);
    } catch (e) {
      toast.error(`Template upload failed: ${errorMessage(e)}`);
    } finally {
      setUploadingTemplate(false);
    }
  };

  const templateName = templateFile?.filename ?? templates?.find((t) => t.id === templateId)?.name ?? null;
  const templateReady = mode === "scratch" || Boolean(templateFile || templateId);

  // ------------------------------------------------------------ sources
  const [sources, setSources] = useState<WorkspaceFile[]>([]);
  const [uploadingSources, setUploadingSources] = useState(false);
  const [removing, setRemoving] = useState<string | null>(null);

  const uploadSources = async (files: File[]) => {
    if (!ws) return;
    const limit = health ? health.max_upload_mb * 1024 * 1024 : Infinity;
    setUploadingSources(true);
    let added = 0;
    for (const f of files) {
      if (!/\.(txt|csv)$/i.test(f.name)) {
        toast.error(`${f.name}: only .txt and .csv sources are supported.`);
        continue;
      }
      if (f.size > limit) {
        toast.error(`${f.name} is larger than ${health?.max_upload_mb} MB.`);
        continue;
      }
      try {
        const wf = await api.uploadWorkspaceFile(ws.id, f, "source");
        setSources((s) => [...s, wf]);
        added++;
      } catch (e) {
        toast.error(`${f.name}: ${errorMessage(e)}`);
      }
    }
    setUploadingSources(false);
    if (added) inputsChanged("Source files changed.");
  };

  const removeSource = async (f: WorkspaceFile) => {
    if (!ws) return;
    setRemoving(f.id);
    try {
      await api.removeWorkspaceFile(ws.id, f.id); // detaches only; the stored upload is not deleted
      setSources((s) => s.filter((x) => x.id !== f.id));
      inputsChanged("Source files changed.");
    } catch (e) {
      toast.error(`Could not remove ${f.filename}: ${errorMessage(e)}`);
    } finally {
      setRemoving(null);
    }
  };

  // ------------------------------------------------------------ prompt -> analyze -> clarify
  const [draft, setDraft] = useState("");
  const draftRef = useRef<HTMLTextAreaElement>(null);
  /** Proposals the user cancelled: they can no longer be approved or generated. */
  const [cancelled, setCancelled] = useState<Set<number>>(() => new Set());
  const [outputName, setOutputName] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [clarifying, setClarifying] = useState(false);
  const [running, setRunning] = useState(false);
  const busy = analyzing || clarifying || running;

  const send = async () => {
    const text = draft.trim();
    if (!ws || !text || busy) return;
    if (!templateReady) return void toast.error("Choose or upload a template first.");
    if (uploadingSources || uploadingTemplate) return void toast.error("Wait for the uploads to finish.");
    push({ kind: "user", text, files: sources.map((s) => s.filename), template: mode === "template" ? templateName : null });
    setDraft("");
    setAnalyzing(true);
    const refine = !!lastAnalysis?.res.structure && lastAnalysis.ver === inputsVer;
    const prompt = refine ? [...prompts.current, text].join("\n") : text;
    prompts.current = refine ? [...prompts.current, text] : [text];
    try {
      const res = await api.analyze(ws.id, {
        prompt,
        source_file_ids: sources.map((s) => s.id),
        template_file_id: templateFile?.id ?? null,
        template_id: templateId,
        output_filename: outputName.trim() || null,
      });
      push({ kind: "analysis", res, ver: inputsVer });
    } catch (e) {
      const msg = errorMessage(e);
      toast.error(`Analyze failed: ${msg}`);
      push({ kind: "note", tone: "error", text: `I could not analyze that request: ${msg}` });
      setDraft((d) => d || text);
    } finally {
      setAnalyzing(false);
    }
  };

  const submitClarify = async (res: AnalyzeResponse, answers: Record<string, ClarifyAnswerValue>) => {
    if (!ws || !res.clarification_id) return;
    setClarifying(true);
    try {
      const next = await api.clarify(ws.id, { clarification_id: res.clarification_id, answers });
      push({
        kind: "answers",
        lines: res.questions.map((q) => `${q.question} → ${answerText(answers[q.id] ?? null)}`),
      });
      push({ kind: "analysis", res: next, ver: inputsVer });
    } catch (e) {
      toast.error(`Could not submit answers: ${errorMessage(e)}`);
    } finally {
      setClarifying(false);
    }
  };

  // ------------------------------------------------------------ generate
  const stream = useRef<StreamHandle | null>(null);
  useEffect(() => () => stream.current?.close(), []);

  const generate = async () => {
    if (!ws || busy) return;
    setRunning(true);
    const id = push({ kind: "run", events: [], live: true, failed: null });
    const patch = (fn: (e: Extract<Entry, { kind: "run" }>) => Partial<Extract<Entry, { kind: "run" }>>) =>
      setLog((l) => l.map((e) => (e.id === id && e.kind === "run" ? { ...e, ...fn(e) } : e)));
    const fail = (msg: string) => {
      setRunning(false);
      patch(() => ({ live: false, failed: msg }));
    };
    let runId: string;
    try {
      runId = (await api.generate(ws.id)).run_id;
    } catch (e) {
      const msg = errorMessage(e);
      toast.error(`Generate failed: ${msg}`);
      return fail(msg);
    }
    stream.current?.close();
    stream.current = openEventStream(api.runEventsUrl(runId), {
      onEvent: (env) => {
        patch((e) => ({ events: [...e.events, env] }));
        const d = env.data as Record<string, unknown>;
        if (env.event === "done") {
          if (typeof d.workbook_id === "string") {
            setRunning(false);
            patch(() => ({ live: false }));
            toast.success("Workbook generated");
            navigate(`/w/${d.workbook_id}`);
          } else {
            const errs = Array.isArray(d.errors) ? (d.errors as unknown[]).map(String).join("; ") : "";
            fail(errs || "The run finished without creating a workbook.");
          }
        } else if (env.event === "error") fail(String(d.message ?? "Generation failed."));
        else if (env.event === "cancelled") fail("The run was cancelled.");
      },
      onFailure: fail,
    });
  };

  // The newest plan is actionable only while inputs are unchanged and it has not been run yet.
  const lastAnalysis = [...log].reverse().find((e) => e.kind === "analysis") as Extract<Entry, { kind: "analysis" }> | undefined;
  const lastIdx = lastAnalysis ? log.indexOf(lastAnalysis) : -1;
  const ranAfter = lastIdx >= 0 && log.slice(lastIdx + 1).some((e) => e.kind === "run");
  const actionableId =
    lastAnalysis && lastAnalysis.ver === inputsVer && !ranAfter && !cancelled.has(lastAnalysis.id) ? lastAnalysis.id : null;
  const lastRunId = [...log].reverse().find((e) => e.kind === "run")?.id ?? null;

  // ------------------------------------------------------------ chat scroll + drop
  const viewport = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = viewport.current;
    if (el) el.scrollTo({ top: el.scrollHeight, behavior: "smooth" });
  }, [log, analyzing, clarifying]);

  const [dragOver, setDragOver] = useState(false);
  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (busy || !ws) return;
    const files = Array.from(e.dataTransfer.files);
    const xlsx = files.filter((f) => /\.xlsx$/i.test(f.name));
    const rest = files.filter((f) => !/\.xlsx$/i.test(f.name));
    if (xlsx.length) {
      if (mode === "template") void uploadTemplate(xlsx[0]);
      else toast.error("Templates (.xlsx) are only used in “New from template”. Sources must be .txt or .csv.");
    }
    if (rest.length) void uploadSources(rest);
  };

  const attachInput = useRef<HTMLInputElement>(null);
  const templateInput = useRef<HTMLInputElement>(null);

  // ------------------------------------------------------------ render
  const header = (
    <header className="flex h-(--app-chrome-height) shrink-0 items-center gap-2 border-b border-stroke-subtle px-3 sm:px-4">
      <Button variant="ghost" size="icon-sm" asChild>
        <Link to="/" aria-label="Back to start">
          <ArrowLeft aria-hidden />
        </Link>
      </Button>
      <div className="hidden sm:block">
        <AppMark />
      </div>
      <div className="mx-1 hidden h-5 w-px bg-border sm:block" aria-hidden />
      <h1 className="min-w-0 flex-1 truncate text-sm font-semibold">
        {mode === "template" ? "New from template" : "New from scratch"}
      </h1>
      <LlmBadge />
    </header>
  );

  if (wsError && !ws) {
    return (
      <div className="flex h-dvh flex-col bg-background">
        {header}
        <div className="mx-auto w-full max-w-lg p-6">
          <Alert variant="destructive">
            <OctagonX aria-hidden />
            <AlertTitle>Could not start a workspace</AlertTitle>
            <AlertDescription>
              <p>{wsError}</p>
              <div className="mt-3 flex gap-2">
                <Button size="sm" variant="outline" onClick={createWs}>
                  <RotateCcw aria-hidden /> Retry
                </Button>
                <Button size="sm" variant="ghost" asChild>
                  <Link to="/">Back home</Link>
                </Button>
              </div>
            </AlertDescription>
          </Alert>
        </div>
      </div>
    );
  }

  const left = (
    <section className="flex h-full min-h-0 flex-col" aria-label="Template preview">
      <div className="flex h-10 shrink-0 items-center gap-2 border-b border-stroke-subtle px-3 text-sm">
        <Eye className="size-4 text-muted-foreground" aria-hidden />
        {mode === "scratch" ? (
          <span className="truncate text-muted-foreground">The new workbook opens here for editing after you generate it.</span>
        ) : pv ? (
          <>
            <span className="truncate font-medium">{pv.name}</span>
            <span className="shrink-0 rounded-md bg-muted px-1.5 py-0.5 text-xs text-muted-foreground">Preview · read-only</span>
            {pv.data && pv.data.warnings.length > 0 && (
              <span className="flex shrink-0 items-center gap-1 text-xs text-warning" title={pv.data.warnings.join("\n")}>
                <CircleAlert className="size-3.5" aria-hidden /> {pv.data.warnings.length} preview note
                {pv.data.warnings.length > 1 ? "s" : ""}
              </span>
            )}
          </>
        ) : (
          <span className="truncate text-muted-foreground">Pick a template in the chat to preview it here.</span>
        )}
      </div>
      <div className="min-h-0 flex-1">
        <PreviewGrid
          preview={pv?.data ?? null}
          previewKey={pv?.key ?? "empty"}
          loading={pv?.loading}
          error={pv?.error}
          onRetry={() => {
            const t = templates?.find((x) => x.id === templateId);
            if (t) previewBuiltin(t);
            else if (ws && templateFile) previewUploaded(ws.id, templateFile);
          }}
        />
      </div>
    </section>
  );

  const right = (
    <aside
      className={cn("relative flex h-full min-h-0 flex-col bg-background", dragOver && "ring-2 ring-inset ring-brand")}
      aria-label="Assistant"
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setDragOver(false);
      }}
      onDrop={onDrop}
    >
      <div className="flex h-10 shrink-0 items-center gap-2 border-b border-stroke-subtle px-3">
        <MessagesSquare className="size-4 text-brand" aria-hidden />
        <h2 className="flex-1 text-sm font-semibold">Assistant</h2>
      </div>

      <div ref={viewport} className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto px-3 py-3" aria-live="polite">
        {!ws ? (
          <div className="space-y-2" aria-busy="true" aria-label="Creating workspace">
            <Skeleton className="h-4 w-3/4" />
            <Skeleton className="h-16 w-full" />
          </div>
        ) : (
          <>
            <Bot>
              {mode === "template" ? (
                <>
                  <p>Which template should I fill? Pick one below to preview it on the left, or upload your own .xlsx.</p>
                  <TemplatePicker
                    templates={templates ?? (galleryAvailable === false ? [] : null)}
                    thumbs={thumbs}
                    selectedId={templateId}
                    uploaded={templateFile}
                    uploadedPreview={
                      templateFile && pv?.key === `file:${templateFile.id}` ? (pv.data ?? (pv.loading ? undefined : null)) : undefined
                    }
                    disabled={busy}
                    uploading={uploadingTemplate}
                    onChoose={(t) => void chooseBuiltin(t)}
                    onUpload={() => templateInput.current?.click()}
                  />
                    <input
                      ref={templateInput}
                      type="file"
                      accept=".xlsx"
                      className="hidden"
                      aria-hidden
                      tabIndex={-1}
                      aria-label="Upload template"
                      onChange={(e) => {
                        const f = e.target.files?.[0];
                        e.target.value = "";
                        if (f) void uploadTemplate(f);
                      }}
                    />
                  {templateReady && (
                    <p className="mt-2">
                      Next: attach the .txt/.csv files I should read (optional) with <Paperclip className="inline size-3.5" aria-label="the clip button" />, then
                      tell me what to fill.
                    </p>
                  )}
                </>
              ) : (
                <p>
                  Describe the workbook you want. Attach .txt or .csv files with <Paperclip className="inline size-3.5" aria-label="the clip button" /> if
                  I should use your data; I only use values found in them or that you give me.
                </p>
              )}
            </Bot>

            {log.map((e) => {
              switch (e.kind) {
                case "user":
                  return (
                    <UserBubble key={e.id}>
                      <p className="whitespace-pre-wrap">{e.text}</p>
                      {(e.template || e.files.length > 0) && (
                        <p className="mt-1 text-xs opacity-80">
                          {[e.template && `Template: ${e.template}`, e.files.length > 0 && `Sources: ${e.files.join(", ")}`]
                            .filter(Boolean)
                            .join(" · ")}
                        </p>
                      )}
                    </UserBubble>
                  );
                case "answers":
                  return (
                    <UserBubble key={e.id}>
                      <ul className="space-y-0.5">
                        {e.lines.map((l, i) => (
                          <li key={i}>{l}</li>
                        ))}
                      </ul>
                    </UserBubble>
                  );
                case "note":
                  return (
                    <Bot key={e.id}>
                      <p className={cn(e.tone === "error" ? "text-destructive" : "text-muted-foreground")}>{e.text}</p>
                    </Bot>
                  );
                case "analysis":
                  return (
                    <Bot key={e.id}>
                      <AnalysisView
                        res={e.res}
                        actionable={e.id === actionableId && !busy}
                        stale={e.ver !== inputsVer}
                        clarifying={clarifying}
                        onClarify={(a) => void submitClarify(e.res, a)}
                        onGenerate={() => void generate()}
                        cancelled={cancelled.has(e.id)}
                        onRequestChanges={() => {
                          draftRef.current?.focus();
                          toast.info("Describe what to change (columns, rows, sheets); the AI revises this proposal.");
                        }}
                        onCancel={() => {
                          setCancelled((c) => new Set(c).add(e.id));
                          push({ kind: "note", tone: "info", text: "Proposal cancelled — nothing was generated. Send a new request to start again." });
                        }}
                      />
                    </Bot>
                  );
                case "run":
                  return (
                    <Bot key={e.id}>
                      <p className="mb-2 font-medium">{e.live ? "Building and validating the workbook…" : e.failed ? "The run did not complete." : "Done."}</p>
                      <RunTimeline events={e.events} live={e.live} />
                      {e.failed && (
                        <Alert variant="destructive" className="mt-3">
                          <OctagonX aria-hidden />
                          <AlertTitle>Generation did not complete</AlertTitle>
                          <AlertDescription>
                            <p>{e.failed}</p>
                            {e.id === lastRunId && !busy && (
                              <Button size="sm" className="mt-2" onClick={() => void generate()}>
                                <RotateCcw aria-hidden /> Retry
                              </Button>
                            )}
                          </AlertDescription>
                        </Alert>
                      )}
                    </Bot>
                  );
              }
            })}

            {(analyzing || clarifying) && (
              <p className="flex items-center gap-2 text-sm text-muted-foreground" role="status">
                <Loader2 className="size-4 animate-spin" aria-hidden />
                {analyzing ? "Reading the template and sources, planning…" : "Applying your answers…"}
              </p>
            )}
          </>
        )}
      </div>

      {/* composer */}
      <div className="shrink-0 px-3 pb-3 pt-1">
        <div className="rounded-xl border border-border bg-card shadow-xs focus-within:border-brand/60 focus-within:ring-2 focus-within:ring-brand/20">
          {(sources.length > 0 || uploadingSources) && (
            <ul className="flex flex-wrap gap-1.5 px-2 pt-2" aria-label="Attached sources">
              {sources.map((f) => (
                <li key={f.id} className="flex max-w-full items-center gap-1 rounded-md bg-muted px-1.5 py-0.5 text-xs">
                  <FileText className="size-3 shrink-0 text-muted-foreground" aria-hidden />
                  <span className="truncate">{f.filename}</span>
                  <button
                    type="button"
                    className="rounded p-0.5 text-muted-foreground hover:bg-background hover:text-foreground disabled:opacity-50"
                    onClick={() => void removeSource(f)}
                    disabled={busy || removing === f.id}
                    aria-label={`Remove ${f.filename}`}
                  >
                    {removing === f.id ? <Loader2 className="size-3 animate-spin" aria-hidden /> : <X className="size-3" aria-hidden />}
                  </button>
                </li>
              ))}
              {uploadingSources && (
                <li className="flex items-center gap-1 px-1.5 py-0.5 text-xs text-muted-foreground">
                  <Loader2 className="size-3 animate-spin" aria-hidden /> Uploading
                </li>
              )}
            </ul>
          )}
          <textarea
            ref={draftRef}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (isMod(e)) {
                e.preventDefault();
                void send();
              }
            }}
            maxLength={8000}
            rows={3}
            disabled={!ws}
            aria-label="Describe what to build"
            placeholder={
              mode === "template"
                ? templateReady
                  ? "e.g. Fill this quote for the customer in the attached notes"
                  : "Choose a template above first"
                : "e.g. A monthly budget tracker from the attached CSV, with totals"
            }
            className="block max-h-40 min-h-12 w-full resize-none bg-transparent px-3 py-2 text-sm outline-none placeholder:text-muted-foreground"
          />
          <div className="flex items-center gap-1.5 px-2 pb-2">
            <Button
              variant="ghost"
              size="icon-sm"
              onClick={() => attachInput.current?.click()}
              disabled={!ws || busy || uploadingSources}
              aria-label="Attach .txt or .csv sources"
              title={`Attach .txt or .csv sources${health ? ` (up to ${health.max_upload_mb} MB each)` : ""}`}
            >
              <Paperclip aria-hidden />
            </Button>
            <input
              ref={attachInput}
              type="file"
              multiple
              accept=".txt,.csv,text/plain,text/csv"
              className="hidden"
              aria-hidden
              tabIndex={-1}
              aria-label="Attach sources"
              onChange={(e) => {
                const files = Array.from(e.target.files ?? []);
                e.target.value = "";
                if (files.length) void uploadSources(files);
              }}
            />
            <input
              value={outputName}
              onChange={(e) => setOutputName(e.target.value)}
              maxLength={120}
              placeholder="Output name (optional)"
              aria-label="Output filename"
              className="h-7 min-w-0 flex-1 rounded-md bg-transparent px-1.5 text-xs outline-none placeholder:text-muted-foreground focus:bg-muted"
            />
            <span className="hidden shrink-0 text-[11px] text-muted-foreground lg:inline">Ctrl+Enter</span>
            <Button
              size="icon-sm"
              onClick={() => void send()}
              disabled={!ws || busy || !draft.trim() || !templateReady}
              aria-label="Send"
            >
              {analyzing ? <Loader2 className="animate-spin" aria-hidden /> : <ArrowUp aria-hidden />}
            </Button>
          </div>
        </div>
      </div>
    </aside>
  );

  return (
    <div className="flex h-dvh flex-col bg-background">
      {header}
      <main className="min-h-0 flex-1">
        {wide ? (
          <Group orientation="horizontal" className="h-full">
            <Panel id="preview" defaultSize="62%" minSize="40%">
              {left}
            </Panel>
            <Separator
              className="w-px bg-border transition-colors outline-none hover:bg-brand/60 focus-visible:bg-brand data-[separator=active]:bg-brand"
              aria-label="Resize assistant panel"
            />
            <Panel id="chat" defaultSize="38%" minSize="28%">
              {right}
            </Panel>
          </Group>
        ) : (
          <div className="flex h-full flex-col">
            <div className="h-[38dvh] shrink-0 border-b border-border">{left}</div>
            <div className="min-h-0 flex-1">{right}</div>
          </div>
        )}
      </main>
    </div>
  );
}

function Bot({ children }: { children: ReactNode }) {
  return (
    <div className="flex gap-2 text-sm">
      <span className="mt-0.5 grid size-6 shrink-0 place-items-center rounded-full bg-brand-soft text-brand" aria-hidden>
        <Sparkles className="size-3.5" />
      </span>
      <div className="min-w-0 flex-1 leading-relaxed">{children}</div>
    </div>
  );
}

function UserBubble({ children }: { children: ReactNode }) {
  return (
    <div className="flex justify-end">
      <div className="max-w-[85%] rounded-2xl rounded-br-md bg-brand px-3 py-2 text-sm text-brand-foreground">{children}</div>
    </div>
  );
}

function AnalysisView({
  res,
  actionable,
  stale,
  clarifying,
  onClarify,
  onGenerate,
  cancelled,
  onRequestChanges,
  onCancel,
}: {
  res: AnalyzeResponse;
  actionable: boolean;
  stale: boolean;
  clarifying: boolean;
  onClarify: (answers: Record<string, ClarifyAnswerValue>) => void;
  onGenerate: () => void;
  cancelled: boolean;
  onRequestChanges: () => void;
  onCancel: () => void;
}) {
  const st = res.structure;
  const reviewActions = st ? (
    <>
      <Button type="button" variant="outline" onClick={onRequestChanges}>
        <MessagesSquare aria-hidden /> Request changes
      </Button>
      <Button type="button" variant="ghost" onClick={onCancel}>
        <X aria-hidden /> Cancel
      </Button>
    </>
  ) : null;
  const canGenerate = !st || (!!st.approved && st.valid !== false);
  const mapping = res.structure ? (
    <section aria-label="Proposed structure" className="mt-3 flex flex-col gap-2">
      <h3 className="text-sm font-semibold">
        Proposed report for “{res.structure.sheet}”
        <span className="font-normal text-muted-foreground"> · {res.structure.approved ? "approved" : "review before generating"}</span>
      </h3>
      <StructurePlan plan={res.structure} />
    </section>
  ) : res.mapping && res.mapping.length > 0 ? (
    <section aria-label="Field mapping" className="mt-3 flex flex-col gap-2">
      <h3 className="text-sm font-semibold">How the template will be filled</h3>
      <MappingTable rows={res.mapping} />
    </section>
  ) : null;

  if (res.status === "blocked") {
    return (
      <Alert variant="destructive">
        <OctagonX aria-hidden />
        <AlertTitle>I cannot build this workbook</AlertTitle>
        <AlertDescription>{res.reason ?? "The request was blocked."} Send a different request to try again.</AlertDescription>
      </Alert>
    );
  }

  if (res.status === "needs_clarification") {
    return (
      <div>
        <p>A few questions first, so the workbook matches what you need.</p>
        {mapping}
        {actionable ? (
          <div className="mt-3 rounded-lg border border-border p-3">
            <ClarifyForm
              key={res.clarification_id ?? "q"}
              questions={res.questions}
              submitting={clarifying}
              onSubmit={onClarify}
              submitLabel={st ? "Approve plan" : undefined}
              extra={reviewActions}
            />
          </div>
        ) : (
          <p className="mt-2 text-xs text-muted-foreground">
            {cancelled ? "Cancelled." : stale ? "Inputs changed since these questions." : "Answered."}
          </p>
        )}
      </div>
    );
  }

  return (
    <div>
      <p className="mb-2 flex items-center gap-2 font-medium">
        Here is the plan <span className="font-normal text-muted-foreground">· {res.facts} source facts</span>
      </p>
      <ul className="space-y-1.5">
        {res.plan.map((p, i) => (
          <li key={i} className="flex gap-2">
            <Check className="mt-0.5 size-4 shrink-0 text-success" aria-hidden />
            <span>{p}</span>
          </li>
        ))}
      </ul>
      {mapping}
      {res.warnings.length > 0 && (
        <Alert className="mt-3 border-warning/30 bg-warning-soft">
          <CircleAlert className="text-warning" aria-hidden />
          <AlertTitle>Warnings</AlertTitle>
          <AlertDescription>
            <ul className="list-disc pl-4">
              {res.warnings.map((w, i) => (
                <li key={i}>{w}</li>
              ))}
            </ul>
          </AlertDescription>
        </Alert>
      )}
      {actionable ? (
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <Button onClick={onGenerate} disabled={!canGenerate}>
            <Play aria-hidden /> Generate workbook
          </Button>
          {reviewActions}
          {!canGenerate && <span className="text-xs text-destructive">The plan is not approved or its preview is invalid.</span>}
        </div>
      ) : cancelled ? (
        <p className="mt-2 text-xs text-muted-foreground">Cancelled.</p>
      ) : (
        stale && <p className="mt-2 text-xs text-muted-foreground">Inputs changed since this plan.</p>
      )}
    </div>
  );
}
