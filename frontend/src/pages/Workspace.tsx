import { useCallback, useEffect, useMemo, useState } from "react";
import { Link, useParams } from "react-router-dom";
import { AssistantRuntimeProvider, useLocalRuntime } from "@assistant-ui/react";
import { Group, Panel, Separator, usePanelRef } from "react-resizable-panels";
import { ArrowLeft, CircleAlert, Download, Loader2, MessagesSquare, PanelRightOpen, Undo2 } from "lucide-react";
import { toast } from "sonner";
import { AppMark } from "@/components/AppMark";
import { LlmBadge } from "@/components/LlmBadge";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Sheet, SheetContent, SheetDescription, SheetTitle } from "@/components/ui/sheet";
import { Skeleton } from "@/components/ui/skeleton";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { ChatPanel } from "@/components/chat/ChatPanel";
import { createAgentAdapter } from "@/components/chat/adapter";
import { HistoryPopover } from "@/components/workspace/HistoryPopover";
import { SpreadsheetPanel } from "@/components/workspace/SpreadsheetPanel";
import { useMediaQuery } from "@/hooks/useMediaQuery";
import { ApiError, api, errorMessage } from "@/lib/api";
import { useWorkbookStore } from "@/lib/store";
import type { WorkbookMeta } from "@/lib/types";

export default function Workspace() {
  const { workbookId = "" } = useParams();
  // Remount everything (store, runtime/thread, grid) when switching workbooks.
  return <WorkspaceInner key={workbookId} workbookId={workbookId} />;
}

function isEditableTarget(el: Element | null): boolean {
  if (!el) return false;
  if (el.closest("[data-grid-root]")) return true;
  if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement) return true;
  return (el as HTMLElement).isContentEditable === true;
}

function WorkspaceInner({ workbookId }: { workbookId: string }) {
  const [meta, setMeta] = useState<WorkbookMeta | null>(null);
  const [loadError, setLoadError] = useState<{ message: string; notFound: boolean } | null>(null);
  const [ready, setReady] = useState(false);
  const [undoing, setUndoing] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [chatCollapsed, setChatCollapsed] = useState(false);
  const wide = useMediaQuery("(min-width: 900px)");
  const chatRef = usePanelRef();

  const revision = useWorkbookStore((s) => s.revision);
  const canUndo = useWorkbookStore((s) => s.canUndo);
  const saving = useWorkbookStore((s) => s.saving);
  const metaNonce = useWorkbookStore((s) => s.metaNonce);

  const adapter = useMemo(() => createAgentAdapter(workbookId), [workbookId]);
  const runtime = useLocalRuntime(adapter);

  // Initial load: metadata seeds the shared store, then the grid mounts.
  const loadInitial = useCallback(() => {
    setLoadError(null);
    setReady(false);
    const ctl = new AbortController();
    api.metadata(workbookId, { signal: ctl.signal }).then(
      (m) => {
        setMeta(m);
        useWorkbookStore.getState().reset(workbookId, m.current_revision, m.can_undo);
        setReady(true);
      },
      (e) => {
        if (ctl.signal.aborted) return;
        setLoadError({ message: errorMessage(e), notFound: e instanceof ApiError && e.status === 404 && !e.routeMissing });
      },
    );
    return () => ctl.abort();
  }, [workbookId]);
  useEffect(() => loadInitial(), [loadInitial]);

  // Refresh can_undo / name after every revision change.
  useEffect(() => {
    if (!ready || metaNonce === 0) return;
    const ctl = new AbortController();
    api.metadata(workbookId, { signal: ctl.signal }).then(
      (m) => {
        setMeta(m);
        const st = useWorkbookStore.getState();
        st.setCanUndo(m.can_undo);
        if (m.current_revision > st.revision) st.setRevision(m.current_revision);
      },
      () => {
        /* keep the last known metadata; the grid surfaces its own errors */
      },
    );
    return () => ctl.abort();
  }, [metaNonce, ready, workbookId]);

  const undo = useCallback(async () => {
    const st = useWorkbookStore.getState();
    if (!st.canUndo || undoing) return;
    setUndoing(true);
    try {
      const res = await api.undo(workbookId, st.revision);
      st.setRevision(res.revision);
      toast.success(`Undone. Now at revision ${res.revision}.`);
    } catch (e) {
      if (e instanceof ApiError && e.status === 409) {
        const cur = e.body.current_revision;
        if (typeof cur === "number") st.setRevision(cur);
        toast.error("The workbook changed in the meantime. Reloaded the latest revision; try again.");
      } else toast.error(`Undo failed: ${errorMessage(e)}`);
    } finally {
      setUndoing(false);
    }
  }, [workbookId, undoing]);

  // Ctrl/Cmd+Z = server undo, but only when focus is outside the grid and text fields.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey || e.key.toLowerCase() !== "z") return;
      if (isEditableTarget(document.activeElement)) return;
      e.preventDefault();
      void undo();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [undo]);

  const toggleChat = () => {
    if (!wide) return setDrawerOpen((o) => !o);
    const p = chatRef.current;
    if (!p) return;
    if (p.isCollapsed()) p.expand();
    else p.collapse();
  };

  const title = meta?.name || meta?.filename || workbookId;

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <div className="flex h-dvh flex-col bg-background">
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
          <div className="flex min-w-0 flex-1 items-center gap-2">
            {meta ? (
              <h1 className="truncate text-sm font-semibold" title={meta.filename}>
                {title}
              </h1>
            ) : (
              <Skeleton className="h-4 w-40" />
            )}
            {ready && (
              <span
                className="shrink-0 rounded-md bg-muted px-1.5 py-0.5 font-mono text-xs tabular-nums text-muted-foreground"
                aria-label={`Revision ${revision}`}
                data-testid="revision"
              >
                r{revision}
              </span>
            )}
            <span className="shrink-0 text-xs text-muted-foreground" role="status" aria-live="polite">
              {saving ? (
                <span className="inline-flex items-center gap-1">
                  <Loader2 className="size-3 animate-spin" aria-hidden /> Saving
                </span>
              ) : null}
            </span>
          </div>
          <div className="hidden md:block">
            <LlmBadge />
          </div>
          <Tooltip>
            <TooltipTrigger asChild>
              <span>
                <Button variant="ghost" size="sm" onClick={undo} disabled={!canUndo || undoing || !ready} aria-label="Undo last revision">
                  {undoing ? <Loader2 className="animate-spin" aria-hidden /> : <Undo2 aria-hidden />}
                  <span className="hidden lg:inline">Undo</span>
                </Button>
              </span>
            </TooltipTrigger>
            <TooltipContent>{canUndo ? "Undo last revision (Ctrl+Z outside the grid)" : "Nothing to undo"}</TooltipContent>
          </Tooltip>
          {ready && <HistoryPopover workbookId={workbookId} />}
          <Button variant="outline" size="sm" asChild>
            <a href={api.exportUrl(workbookId)} download aria-label="Export .xlsx">
              <Download aria-hidden />
              <span className="hidden sm:inline">Export</span>
            </a>
          </Button>
          {(!wide || chatCollapsed) && (
            <Button variant="ghost" size="icon-sm" onClick={toggleChat} aria-label="Open assistant" aria-expanded={wide ? !chatCollapsed : drawerOpen}>
              {wide ? <PanelRightOpen aria-hidden /> : <MessagesSquare aria-hidden />}
            </Button>
          )}
        </header>

        <main className="min-h-0 flex-1">
          {loadError ? (
            <div className="mx-auto max-w-lg p-6">
              <Alert variant="destructive">
                <CircleAlert aria-hidden />
                <AlertTitle>{loadError.notFound ? "Workbook not found" : "Could not open the workbook"}</AlertTitle>
                <AlertDescription>
                  <p>{loadError.message}</p>
                  <div className="mt-3 flex gap-2">
                    {!loadError.notFound && (
                      <Button size="sm" variant="outline" onClick={loadInitial}>
                        Retry
                      </Button>
                    )}
                    <Button size="sm" variant="ghost" asChild>
                      <Link to="/">Back to start</Link>
                    </Button>
                  </div>
                </AlertDescription>
              </Alert>
            </div>
          ) : !ready ? (
            <div className="flex h-full flex-col gap-2 p-4" aria-busy="true" aria-label="Loading workbook">
              <Skeleton className="h-6 w-1/3" />
              <Skeleton className="flex-1" />
            </div>
          ) : wide ? (
            <Group orientation="horizontal" className="h-full">
              <Panel id="editor" defaultSize="68%" minSize="48%">
                <SpreadsheetPanel workbookId={workbookId} />
              </Panel>
              <Separator
                className="w-px bg-border transition-colors outline-none hover:bg-brand/60 focus-visible:bg-brand data-[separator=active]:bg-brand"
                aria-label="Resize assistant panel"
              />
              <Panel
                id="chat"
                panelRef={chatRef}
                defaultSize="32%"
                minSize="22%"
                collapsible
                collapsedSize="0%"
                onResize={(size) => setChatCollapsed(size.asPercentage < 1)}
              >
                <ChatPanel onCollapse={() => chatRef.current?.collapse()} />
              </Panel>
            </Group>
          ) : (
            <>
              <SpreadsheetPanel workbookId={workbookId} />
              <Sheet open={drawerOpen} onOpenChange={setDrawerOpen}>
                <SheetContent side="right" className="w-full p-0 sm:max-w-md">
                  <SheetTitle className="sr-only">Assistant</SheetTitle>
                  <SheetDescription className="sr-only">Chat with the assistant about this workbook.</SheetDescription>
                  <ChatPanel />
                </SheetContent>
              </Sheet>
            </>
          )}
        </main>
      </div>
    </AssistantRuntimeProvider>
  );
}
