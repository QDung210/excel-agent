import { useState } from "react";
import type { ToolCallMessagePartComponent } from "@assistant-ui/react";
import { CheckCircle2, ChevronRight, Loader2, Wrench, XCircle } from "lucide-react";
import type { ToolCardResult } from "./adapter";

const pretty = (name: string) => name.replace(/[_-]+/g, " ").replace(/^\w/, (c) => c.toUpperCase());

/** Compact card for a backend tool invocation (tool_started / tool_completed). */
export const ToolCard: ToolCallMessagePartComponent = ({ toolName, args: rawArgs, result: rawResult }) => {
  const args = rawArgs as Record<string, unknown> | undefined;
  const result = rawResult as ToolCardResult | undefined;
  const [open, setOpen] = useState(false);
  const status = result?.status ?? "running";
  const hasArgs = args && Object.keys(args).length > 0;
  const Icon = status === "running" ? Loader2 : status === "ok" ? CheckCircle2 : XCircle;
  const tone = status === "ok" ? "text-success" : status === "error" ? "text-destructive" : "text-muted-foreground";

  return (
    <div className="my-1.5 rounded-lg border border-border bg-muted/40 text-sm" data-tool={toolName}>
      <button
        type="button"
        className="flex w-full items-center gap-2 px-2.5 py-1.5 text-left disabled:cursor-default"
        onClick={() => setOpen((o) => !o)}
        disabled={!hasArgs}
        aria-expanded={hasArgs ? open : undefined}
        aria-label={`${pretty(toolName)}: ${status}`}
      >
        <Wrench className="size-3.5 shrink-0 text-muted-foreground" aria-hidden />
        <span className="font-medium">{pretty(toolName)}</span>
        <span className="min-w-0 flex-1 truncate text-xs text-muted-foreground">{result?.summary}</span>
        <Icon className={`size-4 shrink-0 ${tone} ${status === "running" ? "animate-spin" : ""}`} aria-hidden />
        {hasArgs && (
          <ChevronRight className={`size-3.5 shrink-0 transition-transform ${open ? "rotate-90" : ""}`} aria-hidden />
        )}
      </button>
      {open && hasArgs && (
        <pre className="max-h-48 overflow-auto border-t border-border px-2.5 py-1.5 font-mono text-[11px] leading-relaxed">
          {JSON.stringify(args, null, 2)}
        </pre>
      )}
    </div>
  );
};
