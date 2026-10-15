import { useCallback, useEffect, useRef, useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import { motion } from "motion/react";
import { ArrowRight, FileSpreadsheet, FileUp, LayoutTemplate, Loader2, RefreshCw, Wand2 } from "lucide-react";
import { toast } from "sonner";
import { AppMark } from "@/components/AppMark";
import { LlmBadge } from "@/components/LlmBadge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { api, errorMessage } from "@/lib/api";
import { relativeTime } from "@/lib/format";
import type { WorkbookSummary } from "@/lib/types";
import { useHealth } from "@/hooks/useHealth";

const MODES = [
  {
    mode: "scratch",
    title: "Generate from scratch",
    body: "Describe the workbook you need, attach notes or CSV exports, and let the assistant build it.",
    icon: Wand2,
  },
  {
    mode: "template",
    title: "Generate with template",
    body: "Start from an .xlsx template and fill it from your sources while keeping its layout and styles.",
    icon: LayoutTemplate,
  },
] as const;

export default function Landing() {
  const navigate = useNavigate();
  const { health } = useHealth();
  const [workbooks, setWorkbooks] = useState<WorkbookSummary[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const load = useCallback(() => {
    setLoadError(null);
    setWorkbooks(null);
    api.listWorkbooks().then(
      (list) => setWorkbooks([...list].sort((a, b) => (b.updated ?? 0) - (a.updated ?? 0))),
      (e) => setLoadError(errorMessage(e)),
    );
  }, []);

  useEffect(load, [load]);

  const onUpload = async (file: File | undefined) => {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
      toast.error("Only .xlsx workbooks can be opened.");
      return;
    }
    if (health && file.size > health.max_upload_mb * 1024 * 1024) {
      toast.error(`File is larger than the ${health.max_upload_mb} MB limit.`);
      return;
    }
    setUploading(true);
    try {
      const wb = await api.uploadWorkbook(file);
      toast.success(`Opened ${wb.name}`);
      navigate(`/w/${wb.id}`);
    } catch (e) {
      toast.error(`Upload failed: ${errorMessage(e)}`);
    } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  return (
    <div className="flex min-h-full flex-col bg-background">
      <header className="flex h-(--app-chrome-height) items-center justify-between border-b border-stroke-subtle px-4 sm:px-6">
        <AppMark />
        <LlmBadge />
      </header>

      <main className="mx-auto grid w-full max-w-6xl flex-1 gap-10 px-4 py-10 sm:px-6 lg:grid-cols-[1.45fr_1fr] lg:py-16">
        <section aria-labelledby="start-heading" className="flex flex-col gap-8">
          <div className="max-w-xl">
            <p className="text-sm font-medium text-brand">AI Excel Workspace</p>
            <h1 id="start-heading" className="mt-2 text-3xl font-semibold tracking-tight text-balance sm:text-4xl">
              Real .xlsx files, built and edited with an assistant.
            </h1>
            <p className="mt-3 text-muted-foreground text-pretty">
              Every change is a validated revision of the actual workbook. Undo any step, export at any time.
            </p>
          </div>

          <div className="grid gap-4 sm:grid-cols-2">
            {MODES.map((m, i) => (
              <motion.div
                key={m.mode}
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.05 * i, duration: 0.25, ease: "easeOut" }}
              >
                <Link
                  to={`/new/${m.mode}`}
                  className="group flex h-full flex-col gap-4 rounded-xl border border-border bg-card p-5 shadow-xs transition-[border-color,box-shadow,transform] duration-(--transition-fast) hover:-translate-y-0.5 hover:border-brand/40 hover:shadow-md"
                >
                  <span className="grid size-10 place-items-center rounded-lg bg-brand-soft text-brand">
                    <m.icon className="size-5" aria-hidden />
                  </span>
                  <span className="flex flex-1 flex-col gap-1.5">
                    <span className="font-semibold">{m.title}</span>
                    <span className="text-sm text-muted-foreground">{m.body}</span>
                  </span>
                  <span className="flex items-center gap-1 text-sm font-medium text-brand">
                    Start <ArrowRight className="size-4 transition-transform group-hover:translate-x-0.5" aria-hidden />
                  </span>
                </Link>
              </motion.div>
            ))}
          </div>

          <div className="flex flex-wrap items-center gap-3 rounded-xl border border-dashed border-border p-4">
            <FileUp className="size-5 text-muted-foreground" aria-hidden />
            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium">Already have a workbook?</p>
              <p className="text-xs text-muted-foreground">
                Upload an .xlsx{health ? ` up to ${health.max_upload_mb} MB` : ""} and edit it here.
              </p>
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              className="sr-only"
              tabIndex={-1}
              aria-hidden="true"
              id="open-xlsx"
              onChange={(e) => onUpload(e.target.files?.[0])}
            />
            <Button variant="outline" disabled={uploading} onClick={() => fileRef.current?.click()}>
              {uploading ? <Loader2 className="animate-spin" aria-hidden /> : <FileUp aria-hidden />}
              {uploading ? "Uploading" : "Open existing .xlsx"}
            </Button>
          </div>
        </section>

        <section aria-labelledby="recent-heading" className="flex min-h-0 flex-col">
          <div className="mb-3 flex items-center justify-between">
            <h2 id="recent-heading" className="text-sm font-semibold">
              Recent workbooks
            </h2>
            <Button variant="ghost" size="icon-sm" onClick={load} aria-label="Refresh recent workbooks">
              <RefreshCw aria-hidden />
            </Button>
          </div>
          <div className="rounded-xl border border-border bg-card">
            {loadError ? (
              <Alert variant="destructive" className="border-0">
                <AlertTitle>Could not load workbooks</AlertTitle>
                <AlertDescription>{loadError}</AlertDescription>
              </Alert>
            ) : workbooks === null ? (
              <ul className="divide-y divide-border" aria-busy="true" aria-label="Loading workbooks">
                {[0, 1, 2].map((i) => (
                  <li key={i} className="flex items-center gap-3 p-3">
                    <Skeleton className="size-8 rounded-md" />
                    <div className="flex-1 space-y-1.5">
                      <Skeleton className="h-3.5 w-2/3" />
                      <Skeleton className="h-3 w-1/3" />
                    </div>
                  </li>
                ))}
              </ul>
            ) : workbooks.length === 0 ? (
              <p className="p-6 text-center text-sm text-muted-foreground">
                No workbooks yet. Generate one or open an existing file.
              </p>
            ) : (
              <ul className="max-h-[28rem] divide-y divide-border overflow-y-auto">
                {workbooks.map((wb) => (
                  <li key={wb.id}>
                    <Link
                      to={`/w/${wb.id}`}
                      className="flex items-center gap-3 p-3 transition-colors hover:bg-muted/60 focus-visible:bg-muted/60"
                    >
                      <span className="grid size-8 shrink-0 place-items-center rounded-md bg-success/10 text-success">
                        <FileSpreadsheet className="size-4" aria-hidden />
                      </span>
                      <span className="min-w-0 flex-1">
                        <span className="block truncate text-sm font-medium">{wb.name}</span>
                        <span className="block truncate text-xs text-muted-foreground">
                          r{wb.current_revision} · {wb.origin} · {relativeTime(wb.updated)}
                        </span>
                      </span>
                      <ArrowRight className="size-4 text-muted-foreground" aria-hidden />
                    </Link>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </section>
      </main>
    </div>
  );
}
