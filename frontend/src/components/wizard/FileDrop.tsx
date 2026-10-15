import { useId, useRef, useState, type DragEvent } from "react";
import { FileText, Loader2, Upload, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { formatBytes } from "@/lib/format";
import type { WorkspaceFile } from "@/lib/types";

interface FileDropProps {
  label: string;
  hint: string;
  accept: string;
  multiple?: boolean;
  maxMb?: number;
  busy?: boolean;
  files: WorkspaceFile[];
  onFiles: (files: File[]) => void;
  onRemove?: (file: WorkspaceFile) => void;
  removing?: string | null;
}

export function FileDrop({ label, hint, accept, multiple, maxMb, busy, files, onFiles, onRemove, removing }: FileDropProps) {
  const inputId = useId();
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setOver(false);
    const list = Array.from(e.dataTransfer.files);
    if (list.length) onFiles(multiple ? list : list.slice(0, 1));
  };

  return (
    <div className="flex flex-col gap-3">
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={onDrop}
        className={cn(
          "flex flex-col items-center justify-center gap-2 rounded-xl border border-dashed border-border bg-muted/30 px-4 py-8 text-center transition-colors",
          over && "border-brand bg-brand-soft",
        )}
      >
        {busy ? (
          <Loader2 className="size-6 animate-spin text-muted-foreground" aria-hidden />
        ) : (
          <Upload className="size-6 text-muted-foreground" aria-hidden />
        )}
        <div>
          <label htmlFor={inputId} className="text-sm font-medium">
            {label}
          </label>
          <p className="text-xs text-muted-foreground">
            {hint}
            {maxMb ? ` · up to ${maxMb} MB each` : ""}
          </p>
        </div>
        <input
          ref={inputRef}
          id={inputId}
          type="file"
          className="sr-only"
          tabIndex={-1}
          aria-hidden="true"
          accept={accept}
          multiple={multiple}
          disabled={busy}
          onChange={(e) => {
            const list = Array.from(e.target.files ?? []);
            if (list.length) onFiles(list);
            e.target.value = "";
          }}
        />
        <Button type="button" variant="outline" size="sm" disabled={busy} onClick={() => inputRef.current?.click()}>
          Browse files
        </Button>
      </div>

      {files.length > 0 && (
        <ul className="divide-y divide-border rounded-lg border border-border" aria-label="Attached files">
          {files.map((f) => (
            <li key={f.id} className="flex items-center gap-3 px-3 py-2">
              <FileText className="size-4 shrink-0 text-muted-foreground" aria-hidden />
              <span className="min-w-0 flex-1 truncate text-sm">{f.filename}</span>
              <span className="text-xs text-muted-foreground tabular-nums">{formatBytes(f.size)}</span>
              {onRemove && (
                <Button
                  type="button"
                  variant="ghost"
                  size="icon-xs"
                  disabled={removing === f.id}
                  onClick={() => onRemove(f)}
                  aria-label={`Remove ${f.filename}`}
                >
                  {removing === f.id ? <Loader2 className="animate-spin" aria-hidden /> : <X aria-hidden />}
                </Button>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
