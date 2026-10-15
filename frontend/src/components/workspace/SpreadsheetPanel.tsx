import { memo, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Workbook, type WorkbookInstance } from "@fortune-sheet/react";
import type { Op, Selection, Sheet } from "@fortune-sheet/core";
import "@fortune-sheet/react/dist/index.css";
import { CircleAlert, RotateCcw } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { CompatBanner } from "./CompatBanner";
import { ApiError, api, errorMessage, newOperationId } from "@/lib/api";
import { rangeToA1 } from "@/lib/a1";
import { applyOpToCell, cellKey, classifyOps, translateCell } from "@/lib/ops";
import { useWorkbookStore } from "@/lib/store";
import type { Action, Preview, PreviewCell } from "@/lib/types";

declare global {
  interface Window {
    /** Dev-only log of raw FortuneSheet ops, used to verify the onOp translation. */
    __opLog?: Op[][];
  }
}

const DEBOUNCE_MS = 400;
const CELL_CONTEXT_MENU = ["copy", "paste", "|", "clear"];
const HEADER_CONTEXT_MENU = ["copy", "paste"];

interface Loaded {
  preview: Preview;
  nonce: number;
}

type Pending = { sheetId: string; r: number; c: number; cell: PreviewCell | null };

/**
 * FortuneSheet grid. Isolated from chat state: it subscribes only to `revision`/`reloadNonce`, and it is
 * remounted (key) whenever a new canonical preview arrives.
 */
export const SpreadsheetPanel = memo(function SpreadsheetPanel({ workbookId }: { workbookId: string }) {
  const revision = useWorkbookStore((s) => s.revision);
  const reloadNonce = useWorkbookStore((s) => s.reloadNonce);

  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [retry, setRetry] = useState(0);
  const hasLoadedRef = useRef(false);
  useEffect(() => {
    hasLoadedRef.current = loaded !== null;
  }, [loaded]);

  // UI state preserved across remounts
  const activeSheetRef = useRef<string | null>(null);
  const selectionRef = useRef<{ sheetId: string; row: [number, number]; column: [number, number] } | null>(null);
  const scrollRef = useRef<{ left: number; top: number }>({ left: 0, top: 0 });

  const wbRef = useRef<WorkbookInstance>(null);
  const hostRef = useRef<HTMLDivElement>(null);

  // ------------------------------------------------------------------ load canonical preview
  useEffect(() => {
    const ctrl = new AbortController();
    setError(null);
    api.preview(workbookId, undefined, { signal: ctrl.signal }).then(
      (preview) => {
        setLoaded({ preview, nonce: reloadNonce });
        toast.dismiss("preview-refresh");
        const st = useWorkbookStore.getState();
        if (preview.revision > st.revision) st.setRevision(preview.revision);
      },
      (e) => {
        if (ctrl.signal.aborted) return;
        const msg = errorMessage(e);
        setError(msg);
        // A grid is already showing: keep it, but say it may be stale and offer a reload.
        if (hasLoadedRef.current) {
          toast.error(`Could not refresh the grid for revision ${useWorkbookStore.getState().revision}: ${msg}`, {
            id: "preview-refresh",
            description: "The change was saved; the grid may show stale values until it reloads.",
            action: { label: "Retry", onClick: () => setRetry((n) => n + 1) },
            duration: 10000,
          });
        }
      },
    );
    return () => ctrl.abort();
  }, [workbookId, revision, reloadNonce, retry]);

  const preview = loaded?.preview ?? null;

  // canonical cell map + sheet id -> name
  const canonical = useMemo(() => {
    const map = new Map<string, PreviewCell | null>();
    preview?.sheets.forEach((s, i) => {
      const id = s.id ?? String(i);
      for (const cd of s.celldata) map.set(cellKey(id, cd.r, cd.c), cd.v);
    });
    return map;
  }, [preview]);
  const canonicalRef = useRef(canonical);
  canonicalRef.current = canonical;

  const sheetNames = useMemo(() => {
    const m = new Map<string, string>();
    preview?.sheets.forEach((s, i) => m.set(s.id ?? String(i), s.name));
    return m;
  }, [preview]);
  const sheetNamesRef = useRef(sheetNames);
  sheetNamesRef.current = sheetNames;

  // FortuneSheet data (stable per loaded preview)
  const data: Sheet[] = useMemo(() => {
    if (!preview) return [];
    const ids = preview.sheets.map((s, i) => s.id ?? String(i));
    const visible = preview.sheets.filter((s) => !s.hide).map((s) => s.id ?? "");
    let active = activeSheetRef.current;
    if (!active || !ids.includes(active) || !visible.includes(active)) {
      active = preview.sheets.find((s) => s.status === 1 && !s.hide)?.id ?? visible[0] ?? ids[0];
    }
    activeSheetRef.current = active;
    const sel = selectionRef.current;
    return preview.sheets.map((s, i) => {
      const id = ids[i];
      const sheet: Sheet = {
        id,
        name: s.name,
        order: s.order ?? i,
        status: id === active ? 1 : 0,
        hide: s.hide ?? 0,
        celldata: s.celldata as Sheet["celldata"],
        config: s.config as Sheet["config"],
        // leave room to type beyond the used range
        row: Math.max((s.row ?? 0) + 20, 60),
        column: Math.max((s.column ?? 0) + 4, 26),
      };
      if (s.frozen && s.frozen.type) sheet.frozen = s.frozen as Sheet["frozen"];
      if (sel && sel.sheetId === id) sheet.luckysheet_select_save = [{ row: sel.row, column: sel.column }];
      return sheet;
    });
  }, [preview]);

  // publish initial active sheet + selection to the store (for the chat chip)
  useEffect(() => {
    if (!preview) return;
    const id = activeSheetRef.current;
    const sel = selectionRef.current;
    const name = id ? (sheetNames.get(id) ?? null) : null;
    const range = sel && sel.sheetId === id ? rangeToA1(sel.row[0], sel.column[0], sel.row[1], sel.column[1]) : null;
    useWorkbookStore.getState().setSelection(name, range);
  }, [preview, sheetNames]);

  // restore scroll after a remount
  useEffect(() => {
    if (!loaded) return;
    const { left, top } = scrollRef.current;
    if (!left && !top) return;
    const t = window.setTimeout(() => {
      try {
        wbRef.current?.scroll({ scrollLeft: left, scrollTop: top });
      } catch {
        /* grid not ready */
      }
    }, 30);
    return () => window.clearTimeout(t);
  }, [loaded]);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const onScroll = (e: Event) => {
      const el = e.target as HTMLElement;
      if (el.classList?.contains("fortune-scrollbar-x")) scrollRef.current.left = el.scrollLeft;
      else if (el.classList?.contains("fortune-scrollbar-y")) scrollRef.current.top = el.scrollTop;
    };
    host.addEventListener("scroll", onScroll, true);
    return () => host.removeEventListener("scroll", onScroll, true);
  }, []);

  // FortuneSheet only re-measures its canvas on window "resize". Forward host size changes
  // (banner expand/dismiss, panel drag, chat collapse) as a synthetic resize, once per frame.
  useEffect(() => {
    const host = hostRef.current;
    if (!host || typeof ResizeObserver === "undefined") return;
    let frame = 0;
    let last = "";
    const ro = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect;
      const key = `${Math.round(width)}x${Math.round(height)}`;
      if (key === last) return;
      last = key;
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
    });
    ro.observe(host);
    return () => {
      cancelAnimationFrame(frame);
      ro.disconnect();
    };
  }, []);

  // ------------------------------------------------------------------ manual edits
  const pending = useRef(new Map<string, Pending>());
  const reasons = useRef(new Set<string>());
  const timer = useRef<number | null>(null);
  const flushing = useRef(false);
  /** Revision produced by our own saves that is not yet published to the store (to avoid a remount mid-typing). */
  const savedRev = useRef(0);

  const flush = useCallback(async () => {
    timer.current = null;
    if (flushing.current) {
      timer.current = window.setTimeout(() => void flush(), DEBOUNCE_MS);
      return;
    }
    flushing.current = true;
    const store = useWorkbookStore.getState();
    let needsReload = false;
    const baseRevision = Math.max(store.revision, savedRev.current);
    let newRevision = baseRevision;
    try {
      const entries = [...pending.current.values()];
      pending.current.clear();
      const why = new Set(reasons.current);
      reasons.current.clear();

      const actions: Action[] = [];
      const savedCells: Pending[] = [];
      for (const p of entries) {
        const name = sheetNamesRef.current.get(p.sheetId);
        if (!name) {
          why.add("unknown sheet");
          continue;
        }
        const res = translateCell(name, p.r, p.c, p.cell, canonicalRef.current.get(cellKey(p.sheetId, p.r, p.c)));
        if ("action" in res) {
          actions.push(res.action);
          savedCells.push(p);
        } else if ("unsupported" in res) why.add(res.unsupported);
      }

      if (actions.length) {
        store.setSaving(true);
        try {
          const result = await api.operations(workbookId, {
            base_revision: baseRevision,
            operation_id: newOperationId(),
            origin: "manual",
            actions,
          });
          newRevision = result.revision;
          for (const p of savedCells) canonicalRef.current.set(cellKey(p.sheetId, p.r, p.c), p.cell);
          if (result.validation?.warnings?.length) toast.warning(result.validation.warnings.join(" "));
        } catch (e) {
          needsReload = true;
          if (e instanceof ApiError && e.status === 409) {
            const cur = e.body.current_revision;
            toast.error(`This edit cannot be saved to the XLSX yet: ${e.message}`);
            if (typeof cur === "number") newRevision = cur;
          } else {
            toast.error(`This edit cannot be saved to the XLSX yet: ${errorMessage(e)}`);
          }
        } finally {
          store.setSaving(false);
        }
      }

      if (why.size) {
        needsReload = true;
        toast.error(`This edit cannot be saved to the XLSX yet: ${[...why].join("; ")}.`);
      }
    } finally {
      flushing.current = false;
    }

    // Only remount when no more edits are queued; otherwise the next flush carries on from newRevision.
    savedRev.current = newRevision;
    if (pending.current.size || reasons.current.size) return;
    const st = useWorkbookStore.getState();
    if (newRevision !== st.revision) st.setRevision(newRevision);
    else if (needsReload) st.reloadCanonical();
  }, [workbookId]);

  // flush on unmount (route change) so a typed value is not lost
  useEffect(
    () => () => {
      if (timer.current) {
        window.clearTimeout(timer.current);
        void flush();
      }
    },
    [flush],
  );

  const onOp = useCallback(
    (ops: Op[]) => {
      if (import.meta.env.DEV) {
        (window.__opLog ??= []).push(ops);
        console.debug("[onOp]", JSON.stringify(ops));
      }
      const fallback = activeSheetRef.current ?? "";
      const { unsupported } = classifyOps(ops, fallback);
      for (const r of unsupported) reasons.current.add(r);

      for (const op of ops) {
        const p = op.path;
        if (p?.[0] !== "data" || p.length < 3 || typeof p[1] !== "number" || typeof p[2] !== "number") continue;
        if (p.length > 3 && !["v", "m", "ct", "f", "qp", "spl"].includes(String(p[3]))) continue;
        if (op.op !== "replace" && op.op !== "add" && op.op !== "remove") continue;
        const sheetId = op.id ?? fallback;
        const key = cellKey(sheetId, p[1], p[2]);
        const base = pending.current.get(key)?.cell ?? canonicalRef.current.get(key) ?? null;
        pending.current.set(key, { sheetId, r: p[1], c: p[2], cell: applyOpToCell(base, op) });
      }

      if (!pending.current.size && !reasons.current.size) return;
      if (timer.current) window.clearTimeout(timer.current);
      timer.current = window.setTimeout(() => void flush(), DEBOUNCE_MS);
    },
    [flush],
  );

  const hooks = useMemo(
    () => ({
      afterSelectionChange: (sheetId: string, sel: Selection) => {
        if (!sel?.row || !sel?.column) return;
        const row: [number, number] = [sel.row[0], sel.row[1] ?? sel.row[0]];
        const column: [number, number] = [sel.column[0], sel.column[1] ?? sel.column[0]];
        selectionRef.current = { sheetId, row, column };
        useWorkbookStore
          .getState()
          .setSelection(sheetNamesRef.current.get(sheetId) ?? null, rangeToA1(row[0], column[0], row[1], column[1]));
      },
      afterActivateSheet: (id: string) => {
        activeSheetRef.current = id;
        selectionRef.current = null;
        useWorkbookStore.getState().setSelection(sheetNamesRef.current.get(id) ?? null, null);
      },
    }),
    [],
  );

  // ------------------------------------------------------------------ render
  if (error && !loaded) {
    return (
      <div className="grid h-full place-items-center p-6">
        <div className="max-w-sm text-center" role="alert">
          <CircleAlert className="mx-auto size-8 text-destructive" aria-hidden />
          <p className="mt-2 font-medium">Could not load the workbook preview</p>
          <p className="mt-1 text-sm text-muted-foreground">{error}</p>
          <Button className="mt-4" variant="outline" onClick={() => setRetry((n) => n + 1)}>
            <RotateCcw aria-hidden /> Retry
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-0 flex-col">
      {preview && (
        <CompatBanner
          workbookId={workbookId}
          warnings={preview.warnings ?? []}
          staleFormulas={preview.stale_formulas ?? 0}
        />
      )}
      <div ref={hostRef} className="sheet-host min-h-0 flex-1" data-grid-root aria-label="Spreadsheet editor" role="region">
        {!loaded ? (
          <GridSkeleton />
        ) : (
          <Workbook
            key={`${loaded.preview.revision}:${loaded.nonce}`}
            ref={wbRef}
            data={data}
            onOp={onOp}
            hooks={hooks}
            showToolbar={false}
            showFormulaBar
            allowEdit
            lang="en"
            cellContextMenu={CELL_CONTEXT_MENU}
            headerContextMenu={HEADER_CONTEXT_MENU}
            sheetTabContextMenu={[]}
          />
        )}
      </div>
    </div>
  );
});

function GridSkeleton() {
  return (
    <div className="flex h-full flex-col gap-px bg-border p-px" aria-busy="true" aria-label="Loading workbook">
      <Skeleton className="h-8 rounded-none" />
      {Array.from({ length: 14 }, (_, i) => (
        <div key={i} className="flex gap-px">
          {Array.from({ length: 8 }, (_, j) => (
            <Skeleton key={j} className="h-6 flex-1 rounded-none" style={{ opacity: 1 - i * 0.05 }} />
          ))}
        </div>
      ))}
    </div>
  );
}
