import { useState } from "react";
import { Check, ChevronDown, FileSpreadsheet, Loader2, Upload } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Skeleton } from "@/components/ui/skeleton";
import { TemplateThumb } from "@/components/workspace/TemplateThumb";
import { cn } from "@/lib/utils";
import type { Preview, TemplateEntry, WorkspaceFile } from "@/lib/types";

interface TemplatePickerProps {
  /** null = still loading the gallery. */
  templates: TemplateEntry[] | null;
  /** Thumbnail data per template id: undefined = loading, null = failed. */
  thumbs: Record<string, Preview | null | undefined>;
  selectedId: string | null;
  uploaded: WorkspaceFile | null;
  uploadedPreview: Preview | null | undefined;
  disabled?: boolean;
  uploading?: boolean;
  onChoose: (t: TemplateEntry) => void;
  onUpload: () => void;
}

/** Dropdown of the template gallery; every entry shows a picture of the template. */
export function TemplatePicker({
  templates,
  thumbs,
  selectedId,
  uploaded,
  uploadedPreview,
  disabled,
  uploading,
  onChoose,
  onUpload,
}: TemplatePickerProps) {
  const [open, setOpen] = useState(false);
  const selected = templates?.find((t) => t.id === selectedId) ?? null;

  if (templates === null) return <Skeleton className="mt-2 h-16 w-full rounded-lg" />;

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <button
          type="button"
          disabled={disabled}
          aria-label={selected || uploaded ? `Template: ${uploaded?.filename ?? selected?.name}. Change template` : "Choose a template"}
          className={cn(
            "mt-2 flex w-full items-center gap-3 rounded-lg border border-border bg-background p-2 text-left transition-colors hover:border-brand/40 disabled:opacity-60 data-[state=open]:border-brand",
            (selected || uploaded) && "border-brand/60",
          )}
        >
          {uploaded ? (
            <TemplateThumb preview={uploadedPreview} width={84} height={52} label={`Preview of ${uploaded.filename}`} />
          ) : selected ? (
            <TemplateThumb preview={thumbs[selected.id]} width={84} height={52} label={`Preview of ${selected.name}`} />
          ) : null}
          <span className="min-w-0 flex-1">
            <span className="block truncate text-sm font-medium">
              {uploaded?.filename ?? selected?.name ?? "Choose a template…"}
            </span>
            <span className="block truncate text-xs text-muted-foreground">
              {uploaded ? "Your upload" : selected ? (selected.description ?? selected.sheets.join(" · ")) : `${templates.length} templates, or upload your own .xlsx`}
            </span>
          </span>
          <ChevronDown className={cn("size-4 shrink-0 text-muted-foreground transition-transform", open && "rotate-180")} aria-hidden />
        </button>
      </PopoverTrigger>
      <PopoverContent align="start" className="w-(--radix-popover-trigger-width) min-w-72 gap-1 p-1.5">
        <div role="listbox" aria-label="Templates" className="flex max-h-[60dvh] flex-col gap-1 overflow-y-auto">
          {templates.map((t) => {
            const active = !uploaded && t.id === selectedId;
            return (
              <button
                key={t.id}
                type="button"
                role="option"
                aria-selected={active}
                onClick={() => {
                  setOpen(false);
                  onChoose(t);
                }}
                className={cn(
                  "flex items-center gap-3 rounded-md p-1.5 text-left outline-none hover:bg-muted focus-visible:bg-muted",
                  active && "bg-brand-soft hover:bg-brand-soft",
                )}
              >
                <TemplateThumb preview={thumbs[t.id]} width={184} height={116} label={`Preview of ${t.name}`} />
                <span className="min-w-0 flex-1">
                  <span className="flex items-center gap-1.5 text-sm font-medium">
                    <span className="truncate">{t.name}</span>
                    {active && <Check className="size-4 shrink-0 text-brand" aria-hidden />}
                  </span>
                  {t.description && <span className="line-clamp-2 text-xs text-muted-foreground">{t.description}</span>}
                  <span className="mt-0.5 block truncate text-xs text-muted-foreground">{t.sheets.join(" · ")}</span>
                </span>
              </button>
            );
          })}
          {templates.length === 0 && <p className="px-2 py-3 text-sm text-muted-foreground">No templates in the gallery yet.</p>}
        </div>
        <div className="mt-1 border-t border-border pt-1">
          <button
            type="button"
            disabled={uploading}
            onClick={() => {
              setOpen(false);
              onUpload();
            }}
            className="flex w-full items-center gap-2 rounded-md px-2 py-2 text-left text-sm outline-none hover:bg-muted focus-visible:bg-muted disabled:opacity-60"
          >
            {uploading ? <Loader2 className="size-4 animate-spin" aria-hidden /> : <Upload className="size-4 text-brand" aria-hidden />}
            Upload your own template (.xlsx)
            {uploaded && (
              <span className="ml-auto flex min-w-0 items-center gap-1 text-xs text-muted-foreground">
                <FileSpreadsheet className="size-3.5 shrink-0" aria-hidden />
                <span className="truncate">{uploaded.filename}</span>
              </span>
            )}
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
}
