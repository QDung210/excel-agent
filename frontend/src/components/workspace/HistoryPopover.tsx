import { useState } from "react";
import { History, Loader2, RotateCcw } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Skeleton } from "@/components/ui/skeleton";
import { api, errorMessage } from "@/lib/api";
import { relativeTime } from "@/lib/format";
import { useWorkbookStore } from "@/lib/store";
import type { HistoryEntry } from "@/lib/types";

const EVENT_LABEL: Record<string, string> = {
  created: "Created",
  operations: "Edited",
  undo: "Undo",
  restore: "Restored",
};

export function HistoryPopover({ workbookId }: { workbookId: string }) {
  const revision = useWorkbookStore((s) => s.revision);
  const [entries, setEntries] = useState<HistoryEntry[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [restoring, setRestoring] = useState<number | null>(null);

  const load = () => {
    setError(null);
    setEntries(null);
    api.history(workbookId).then(
      (list) => setEntries([...list].reverse()),
      (e) => setError(errorMessage(e)),
    );
  };

  const restore = async (target: number) => {
    setRestoring(target);
    try {
      const res = await api.undo(workbookId, useWorkbookStore.getState().revision, target);
      useWorkbookStore.getState().setRevision(res.revision);
      toast.success(`Restored revision ${target} as revision ${res.revision}`);
      load();
    } catch (e) {
      toast.error(`Restore failed: ${errorMessage(e)}`);
    } finally {
      setRestoring(null);
    }
  };

  return (
    <Popover onOpenChange={(o) => o && load()}>
      <PopoverTrigger asChild>
        <Button variant="ghost" size="sm" aria-label="Revision history">
          <History aria-hidden />
          <span className="hidden lg:inline">History</span>
        </Button>
      </PopoverTrigger>
      <PopoverContent align="end" className="w-80 p-0">
        <div className="border-b border-border px-3 py-2 text-sm font-semibold">Revision history</div>
        <div className="max-h-80 overflow-y-auto">
          {error ? (
            <p className="p-3 text-sm text-destructive">{error}</p>
          ) : entries === null ? (
            <div className="space-y-2 p-3" aria-busy="true">
              <Skeleton className="h-4 w-3/4" />
              <Skeleton className="h-4 w-2/3" />
              <Skeleton className="h-4 w-1/2" />
            </div>
          ) : entries.length === 0 ? (
            <p className="p-3 text-sm text-muted-foreground">No history yet.</p>
          ) : (
            <ol className="divide-y divide-border">
              {entries.map((h, i) => {
                const current = h.revision === revision;
                return (
                  <li key={`${h.revision}-${i}`} className="flex items-center gap-3 px-3 py-2">
                    <span
                      className={`grid h-6 min-w-8 place-items-center rounded-md px-1 font-mono text-xs tabular-nums ${current ? "bg-brand text-brand-foreground" : "bg-muted"}`}
                    >
                      r{h.revision}
                    </span>
                    <div className="min-w-0 flex-1">
                      <p className="truncate text-sm">
                        {EVENT_LABEL[h.event] ?? h.event}
                        {h.origin ? <span className="text-muted-foreground"> · {h.origin}</span> : null}
                      </p>
                      <p className="truncate text-xs text-muted-foreground">
                        {h.note ? `${h.note} · ` : ""}
                        {relativeTime(h.t)}
                      </p>
                    </div>
                    {!current && h.revision < revision && (
                      <Button
                        variant="ghost"
                        size="icon-xs"
                        disabled={restoring !== null}
                        onClick={() => restore(h.revision)}
                        aria-label={`Restore revision ${h.revision}`}
                        title={`Restore revision ${h.revision}`}
                      >
                        {restoring === h.revision ? <Loader2 className="animate-spin" aria-hidden /> : <RotateCcw aria-hidden />}
                      </Button>
                    )}
                  </li>
                );
              })}
            </ol>
          )}
        </div>
      </PopoverContent>
    </Popover>
  );
}
