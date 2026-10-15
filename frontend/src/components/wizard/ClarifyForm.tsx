import { useMemo, useState, type FormEvent, type ReactNode } from "react";
import { Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import type { ClarifyAnswerValue, ClarifyQuestion } from "@/lib/types";

interface ClarifyFormProps {
  questions: ClarifyQuestion[];
  submitting: boolean;
  onSubmit: (answers: Record<string, ClarifyAnswerValue>) => void;
  /** Label of the submit button (e.g. "Approve plan"). */
  submitLabel?: string;
  /** More actions shown next to the submit button. */
  extra?: ReactNode;
}

function initialValue(q: ClarifyQuestion): ClarifyAnswerValue {
  const d = q.default;
  if (q.type === "boolean") return typeof d === "boolean" ? d : false;
  if (q.type === "number") return typeof d === "number" ? d : d != null && d !== "" ? Number(d) : null;
  if (q.type === "select") return typeof d === "string" && q.options?.includes(d) ? d : null;
  return d == null ? "" : String(d);
}

/** Renders every clarification question at once, typed by question.type. */
export function ClarifyForm({ questions, submitting, onSubmit, submitLabel, extra }: ClarifyFormProps) {
  const [values, setValues] = useState<Record<string, ClarifyAnswerValue>>(() =>
    Object.fromEntries(questions.map((q) => [q.id, initialValue(q)])),
  );
  const [touched, setTouched] = useState(false);

  const missing = useMemo(
    () =>
      questions
        .filter((q) => q.required && q.type !== "boolean")
        .filter((q) => {
          const v = values[q.id];
          return v === null || v === "" || (typeof v === "number" && Number.isNaN(v));
        })
        .map((q) => q.id),
    [questions, values],
  );

  const set = (id: string, v: ClarifyAnswerValue) => setValues((s) => ({ ...s, [id]: v }));

  const submit = (e: FormEvent) => {
    e.preventDefault();
    setTouched(true);
    if (missing.length) return;
    // Omit empty optional answers so the backend applies its defaults.
    const answers = Object.fromEntries(Object.entries(values).filter(([, v]) => v !== "" && v !== null));
    onSubmit(answers);
  };

  return (
    <form onSubmit={submit} className="flex flex-col gap-5" noValidate aria-label="Clarification questions">
      {questions.map((q) => {
        const id = `q-${q.id}`;
        const invalid = touched && missing.includes(q.id);
        const descId = q.explanation ? `${id}-desc` : undefined;
        const errId = invalid ? `${id}-err` : undefined;
        const describedBy = [descId, errId].filter(Boolean).join(" ") || undefined;
        return (
          <div key={q.id} className="flex flex-col gap-2">
            {q.type === "boolean" ? (
              <div className="flex items-start justify-between gap-4">
                <Label htmlFor={id} className="leading-snug">
                  {q.question}
                </Label>
                <Switch
                  id={id}
                  checked={Boolean(values[q.id])}
                  onCheckedChange={(v) => set(q.id, v)}
                  aria-describedby={describedBy}
                />
              </div>
            ) : (
              <Label htmlFor={id} className="leading-snug">
                {q.question}
                {q.required && (
                  <span className="text-destructive" aria-hidden>
                    *
                  </span>
                )}
              </Label>
            )}

            {q.type === "select" && (
              <Select value={(values[q.id] as string | null) ?? undefined} onValueChange={(v) => set(q.id, v)}>
                <SelectTrigger id={id} className="w-full" aria-invalid={invalid} aria-describedby={describedBy} aria-required={q.required}>
                  <SelectValue placeholder="Choose an option" />
                </SelectTrigger>
                <SelectContent>
                  {(q.options ?? []).map((o) => (
                    <SelectItem key={o} value={o}>
                      {o}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
            {q.type === "text" && (
              <Input
                id={id}
                value={(values[q.id] as string) ?? ""}
                onChange={(e) => set(q.id, e.target.value)}
                aria-invalid={invalid}
                aria-required={q.required}
                aria-describedby={describedBy}
              />
            )}
            {q.type === "number" && (
              <Input
                id={id}
                type="number"
                inputMode="decimal"
                value={values[q.id] == null || Number.isNaN(values[q.id] as number) ? "" : String(values[q.id])}
                onChange={(e) => set(q.id, e.target.value === "" ? null : Number(e.target.value))}
                aria-invalid={invalid}
                aria-required={q.required}
                aria-describedby={describedBy}
              />
            )}

            {q.explanation && (
              <p id={descId} className="text-xs text-muted-foreground">
                {q.explanation}
              </p>
            )}
            {invalid && (
              <p id={errId} className="text-xs text-destructive">
                This answer is required.
              </p>
            )}
          </div>
        );
      })}
      <div className="flex flex-wrap gap-2">
        <Button type="submit" disabled={submitting}>
          {submitting && <Loader2 className="animate-spin" aria-hidden />}
          {submitLabel ?? "Submit answers"}
        </Button>
        {extra}
      </div>
    </form>
  );
}
