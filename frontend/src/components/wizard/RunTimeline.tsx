import { AnimatePresence, motion } from "motion/react";
import { CheckCircle2, CircleAlert, CircleDot, Loader2, OctagonX, Play, Sparkles } from "lucide-react";
import { cn } from "@/lib/utils";
import type { SseEnvelope } from "@/lib/types";

const pretty = (tool: string) => tool.replace(/_/g, " ").replace(/^\w/, (c) => c.toUpperCase());

interface Row {
  key: string;
  icon: typeof CircleDot;
  tone: "muted" | "brand" | "success" | "error" | "warning";
  title: string;
  detail?: string;
  spinning?: boolean;
}

/** Collapse raw run events into timeline rows; a tool_started row is replaced by its tool_completed. */
function toRows(events: SseEnvelope[]): Row[] {
  const rows: Row[] = [];
  const openTool = new Map<string, number>();
  for (const e of events) {
    const d = e.data as Record<string, unknown>;
    switch (e.event) {
      case "run_started":
        rows.push({ key: `s${e.seq}`, icon: Play, tone: "muted", title: "Run started" });
        break;
      case "analysis_ready":
        rows.push({ key: `a${e.seq}`, icon: Sparkles, tone: "brand", title: "Plan ready" });
        break;
      case "tool_started": {
        const tool = String(d.tool ?? "tool");
        openTool.set(tool, rows.length);
        rows.push({ key: `t${e.seq}`, icon: Loader2, tone: "brand", title: pretty(tool), detail: "Running", spinning: true });
        break;
      }
      case "tool_completed": {
        const tool = String(d.tool ?? "tool");
        const ok = d.status === "ok" || d.status === "success";
        const row: Row = {
          key: `t${e.seq}`,
          icon: ok ? CheckCircle2 : CircleAlert,
          tone: ok ? "success" : "error",
          title: pretty(tool),
          detail: d.summary ? String(d.summary) : String(d.status ?? ""),
        };
        const idx = openTool.get(tool);
        if (idx !== undefined) {
          rows[idx] = { ...row, key: rows[idx].key };
          openTool.delete(tool);
        } else rows.push(row);
        break;
      }
      case "validation_failed": {
        const errs = Array.isArray(d.errors) ? (d.errors as unknown[]).map(String) : [];
        rows.push({ key: `v${e.seq}`, icon: OctagonX, tone: "error", title: "Validation failed", detail: errs.join("; ") });
        break;
      }
      case "workbook_revision":
        rows.push({ key: `r${e.seq}`, icon: CircleDot, tone: "brand", title: `Revision ${d.revision ?? ""} created` });
        break;
      case "done":
        rows.push({
          key: `d${e.seq}`,
          icon: d.workbook_id ? CheckCircle2 : CircleAlert,
          tone: d.workbook_id ? "success" : "warning",
          title: d.workbook_id ? "Done" : "Finished without a workbook",
        });
        break;
      case "error":
        rows.push({ key: `e${e.seq}`, icon: OctagonX, tone: "error", title: "Error", detail: String(d.message ?? d.code ?? "") });
        break;
      case "cancelled":
        rows.push({ key: `c${e.seq}`, icon: CircleAlert, tone: "warning", title: "Cancelled" });
        break;
    }
  }
  return rows;
}

const toneCls: Record<Row["tone"], string> = {
  muted: "text-muted-foreground",
  brand: "text-brand",
  success: "text-success",
  error: "text-destructive",
  warning: "text-warning",
};

export function RunTimeline({ events, live }: { events: SseEnvelope[]; live: boolean }) {
  const rows = toRows(events);
  return (
    <ol className="relative flex flex-col gap-0.5" aria-label="Generation progress" aria-live="polite">
      <AnimatePresence initial={false}>
        {rows.map((r, i) => (
          <motion.li
            key={r.key}
            layout
            initial={{ opacity: 0, x: -6 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ duration: 0.18 }}
            className="relative flex gap-3 pb-3"
          >
            {i < rows.length - 1 && <span className="absolute top-6 bottom-0 left-[9px] w-px bg-border" aria-hidden />}
            <r.icon className={cn("mt-0.5 size-5 shrink-0", toneCls[r.tone], r.spinning && "animate-spin")} aria-hidden />
            <div className="min-w-0">
              <p className="text-sm font-medium">{r.title}</p>
              {r.detail && <p className="text-xs break-words text-muted-foreground">{r.detail}</p>}
            </div>
          </motion.li>
        ))}
      </AnimatePresence>
      {live && rows.length === 0 && (
        <li className="flex items-center gap-3 text-sm text-muted-foreground">
          <Loader2 className="size-5 animate-spin" aria-hidden /> Connecting to run
        </li>
      )}
    </ol>
  );
}
