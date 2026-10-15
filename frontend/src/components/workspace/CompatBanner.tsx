import { useState } from "react";
import { AnimatePresence, motion } from "motion/react";
import { ChevronDown, TriangleAlert, X } from "lucide-react";
import { Button } from "@/components/ui/button";

const dismissKey = (id: string) => `compat-dismissed:${id}`;

function readDismissed(id: string): boolean {
  try {
    return sessionStorage.getItem(dismissKey(id)) === "1";
  } catch {
    return false;
  }
}

/** Dismissible notice listing preview fidelity warnings (what the grid cannot show faithfully). */
export function CompatBanner({
  workbookId,
  warnings,
  staleFormulas,
}: {
  workbookId: string;
  warnings: string[];
  staleFormulas: number;
}) {
  const [dismissed, setDismissed] = useState(() => readDismissed(workbookId));
  const [open, setOpen] = useState(false);
  const items = [...warnings];
  if (staleFormulas > 0)
    items.push(`${staleFormulas} formula value(s) may be out of date until the file is recalculated in Excel.`);
  if (!items.length) return null;

  const dismiss = () => {
    setDismissed(true);
    try {
      sessionStorage.setItem(dismissKey(workbookId), "1");
    } catch {
      /* storage unavailable */
    }
  };

  return (
    <AnimatePresence initial={false}>
      {!dismissed && (
        <motion.div
          initial={{ height: 0, opacity: 0 }}
          animate={{ height: "auto", opacity: 1 }}
          exit={{ height: 0, opacity: 0 }}
          transition={{ duration: 0.18 }}
          className="overflow-hidden border-b border-warning/25 bg-warning-soft"
          role="status"
        >
          <div className="flex items-start gap-2 px-3 py-2 text-sm">
            <TriangleAlert className="mt-0.5 size-4 shrink-0 text-warning" aria-hidden />
            <div className="min-w-0 flex-1">
              <button
                type="button"
                className="flex items-center gap-1 text-left font-medium"
                aria-expanded={open}
                aria-controls="compat-list"
                onClick={() => setOpen((o) => !o)}
              >
                Preview compatibility: {items.length} note{items.length === 1 ? "" : "s"}
                <ChevronDown className={`size-4 transition-transform ${open ? "rotate-180" : ""}`} aria-hidden />
              </button>
              <p className="text-xs text-muted-foreground">The exported .xlsx keeps everything; only this preview is simplified.</p>
              {open && (
                <ul id="compat-list" className="mt-1.5 list-disc space-y-0.5 pl-4 text-xs">
                  {items.map((w, i) => (
                    <li key={i}>{w}</li>
                  ))}
                </ul>
              )}
            </div>
            <Button variant="ghost" size="icon-xs" onClick={dismiss} aria-label="Dismiss compatibility notes">
              <X aria-hidden />
            </Button>
          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );
}
