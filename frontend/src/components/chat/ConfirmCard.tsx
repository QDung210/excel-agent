import type { ToolCallMessagePartComponent } from "@assistant-ui/react";
import { Check, Loader2, ShieldQuestion, X } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { api, errorMessage } from "@/lib/api";
import { usePlanStore, useWorkbookStore } from "@/lib/store";
import type { ConfirmationRequiredData } from "@/lib/types";

function show(v: unknown): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

/** Confirmation card for `confirmation_required`: shows the diff and applies/rejects the plan via REST. */
export const ConfirmCard: ToolCallMessagePartComponent = ({ args: rawArgs }) => {
  const args = rawArgs as ConfirmationRequiredData;
  const planId = args?.plan_id;
  const plan = usePlanStore((s) => (planId ? s.plans[planId] : undefined));
  const setPlan = usePlanStore((s) => s.setPlan);
  if (!planId) return null;
  const status = plan?.status ?? "pending";
  const diff = args.diff ?? [];
  const busy = status === "applying" || status === "rejecting";
  const settled = status === "applied" || status === "rejected";

  const apply = async () => {
    setPlan(planId, "applying");
    try {
      const res = await api.applyPlan(planId, useWorkbookStore.getState().revision);
      useWorkbookStore.getState().setRevision(res.revision);
      setPlan(planId, "applied", `Applied as revision ${res.revision}`);
    } catch (e) {
      const msg = errorMessage(e);
      setPlan(planId, "error", msg);
      toast.error(`Could not apply changes: ${msg}`);
    }
  };
  const reject = async () => {
    setPlan(planId, "rejecting");
    try {
      await api.rejectPlan(planId);
      setPlan(planId, "rejected", "Changes discarded");
    } catch (e) {
      const msg = errorMessage(e);
      setPlan(planId, "error", msg);
      toast.error(`Could not reject: ${msg}`);
    }
  };

  return (
    <section
      className="my-2 overflow-hidden rounded-lg border border-brand/30 bg-card text-sm shadow-xs"
      aria-label="Proposed changes awaiting confirmation"
    >
      <header className="flex items-start gap-2 border-b border-border bg-brand-soft px-3 py-2">
        <ShieldQuestion className="mt-0.5 size-4 shrink-0 text-brand" aria-hidden />
        <div className="min-w-0">
          <p className="font-medium">
            Confirm {diff.length || args.actions?.length || 0} change{(diff.length || args.actions?.length) === 1 ? "" : "s"}
          </p>
          {args.reason && <p className="text-xs text-muted-foreground">{args.reason}</p>}
        </div>
      </header>
      {diff.length > 0 && (
        <div className="max-h-56 overflow-auto">
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-muted text-left text-muted-foreground">
              <tr>
                <th scope="col" className="px-3 py-1 font-medium">Cell</th>
                <th scope="col" className="px-3 py-1 font-medium">Before</th>
                <th scope="col" className="px-3 py-1 font-medium">After</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-border">
              {diff.map((row, i) => (
                <tr key={`${row.cell}-${i}`}>
                  <td className="whitespace-nowrap px-3 py-1 font-mono">{row.cell}</td>
                  <td className="max-w-32 truncate px-3 py-1 text-muted-foreground line-through decoration-destructive/50" title={show(row.before)}>
                    {show(row.before) || <span className="italic no-underline">empty</span>}
                  </td>
                  <td className="max-w-32 truncate px-3 py-1 font-medium" title={show(row.after)}>
                    {show(row.after) || <span className="italic text-muted-foreground">empty</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      <footer className="flex items-center gap-2 border-t border-border px-3 py-2">
        {settled ? (
          <p className={`text-xs ${status === "applied" ? "text-success" : "text-muted-foreground"}`} role="status">
            {plan?.message}
          </p>
        ) : (
          <>
            {status === "error" && (
              <p className="min-w-0 flex-1 truncate text-xs text-destructive" role="alert" title={plan?.message}>
                {plan?.message}
              </p>
            )}
            <div className="ml-auto flex gap-2">
              <Button size="sm" variant="outline" onClick={reject} disabled={busy}>
                {status === "rejecting" ? <Loader2 className="animate-spin" aria-hidden /> : <X aria-hidden />}
                Reject
              </Button>
              <Button size="sm" onClick={apply} disabled={busy}>
                {status === "applying" ? <Loader2 className="animate-spin" aria-hidden /> : <Check aria-hidden />}
                Apply
              </Button>
            </div>
          </>
        )}
      </footer>
    </section>
  );
};
