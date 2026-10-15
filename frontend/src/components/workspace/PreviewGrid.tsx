import { memo, useEffect, useMemo, useRef } from "react";
import { Workbook } from "@fortune-sheet/react";
import type { Sheet } from "@fortune-sheet/core";
import "@fortune-sheet/react/dist/index.css";
import { CircleAlert, RotateCcw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import type { Preview } from "@/lib/types";

const EMPTY: Sheet[] = [{ id: "empty", name: "Sheet1", order: 0, status: 1, celldata: [], config: {}, row: 60, column: 26 }];

interface PreviewGridProps {
  /** null = nothing selected yet: an empty read-only sheet is shown. */
  preview: Preview | null;
  /** Remount key, e.g. the template id. */
  previewKey: string;
  loading?: boolean;
  error?: string | null;
  onRetry?: () => void;
}

/** Read-only FortuneSheet view of a backend preview (a template before generation). Nothing typed here is saved. */
export const PreviewGrid = memo(function PreviewGrid({ preview, previewKey, loading, error, onRetry }: PreviewGridProps) {
  const hostRef = useRef<HTMLDivElement>(null);

  const data: Sheet[] = useMemo(() => {
    if (!preview || preview.sheets.length === 0) return EMPTY;
    const ids = preview.sheets.map((s, i) => s.id ?? String(i));
    const active = preview.sheets.find((s) => s.status === 1 && !s.hide)?.id ?? preview.sheets.find((s) => !s.hide)?.id ?? ids[0];
    return preview.sheets.map((s, i) => {
      const sheet: Sheet = {
        id: ids[i],
        name: s.name,
        order: s.order ?? i,
        status: ids[i] === active ? 1 : 0,
        hide: s.hide ?? 0,
        celldata: s.celldata as Sheet["celldata"],
        config: s.config as Sheet["config"],
        row: Math.max((s.row ?? 0) + 10, 40),
        column: Math.max((s.column ?? 0) + 2, 26),
      };
      if (s.frozen && s.frozen.type) sheet.frozen = s.frozen as Sheet["frozen"];
      return sheet;
    });
  }, [preview]);

  // FortuneSheet only re-measures on window "resize"; forward panel size changes.
  useEffect(() => {
    const host = hostRef.current;
    if (!host || typeof ResizeObserver === "undefined") return;
    let frame = 0;
    const ro = new ResizeObserver(() => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
    });
    ro.observe(host);
    return () => {
      cancelAnimationFrame(frame);
      ro.disconnect();
    };
  }, []);

  if (error) {
    return (
      <div className="grid h-full place-items-center p-6">
        <div className="max-w-sm text-center" role="alert">
          <CircleAlert className="mx-auto size-8 text-destructive" aria-hidden />
          <p className="mt-2 font-medium">Could not load the template preview</p>
          <p className="mt-1 text-sm text-muted-foreground">{error}</p>
          {onRetry && (
            <Button className="mt-4" variant="outline" onClick={onRetry}>
              <RotateCcw aria-hidden /> Retry
            </Button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div ref={hostRef} className="sheet-host h-full min-h-0" data-grid-root aria-label="Template preview (read-only)" role="region">
      {loading ? (
        <div className="flex h-full flex-col gap-px bg-border p-px" aria-busy="true" aria-label="Loading template preview">
          {Array.from({ length: 14 }, (_, i) => (
            <div key={i} className="flex gap-px">
              {Array.from({ length: 8 }, (_, j) => (
                <Skeleton key={j} className="h-6 flex-1 rounded-none" style={{ opacity: 1 - i * 0.05 }} />
              ))}
            </div>
          ))}
        </div>
      ) : (
        <Workbook
          key={previewKey}
          data={data}
          showToolbar={false}
          showFormulaBar
          allowEdit={false}
          lang="en"
          cellContextMenu={[]}
          headerContextMenu={[]}
          sheetTabContextMenu={[]}
        />
      )}
    </div>
  );
});
