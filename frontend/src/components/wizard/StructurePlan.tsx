import { Badge } from "@/components/ui/badge";
import type { PlanCell, PlanRow, Provenance, StructureChange, StructurePreview, StructureView } from "@/lib/types";
import { cn } from "@/lib/utils";

const CHANGE: Record<StructureChange, { label: string; variant: "secondary" | "outline" | "destructive" | "default" }> = {
  add: { label: "New", variant: "default" },
  rename: { label: "Renamed", variant: "secondary" },
  to_input: { label: "Formula → input", variant: "outline" },
  formula: { label: "Formula", variant: "secondary" },
  to_formula: { label: "Input → formula", variant: "secondary" },
  default: { label: "Default value", variant: "outline" },
  keep: { label: "Unchanged", variant: "outline" },
  hidden: { label: "Hidden", variant: "destructive" },
  move: { label: "Moved", variant: "secondary" },
  sequence: { label: "Numbered 1..n", variant: "outline" },
  total: { label: "In totals row", variant: "outline" },
};

const MODE: Record<StructureView["removed"][number]["mode"], string> = {
  hide: "hidden (still used by other formulas)",
  delete: "deleted",
  keep: "kept visible",
};

const FIT: Record<string, string> = { fits: "Template fits", partial: "Template partly fits", mismatch: "Template does not fit" };

const PROV: Record<Provenance, { label: string; cls: string }> = {
  explicit: { label: "Stated in the documents", cls: "" },
  derived: { label: "Computed by a formula", cls: "text-primary" },
  assumed: { label: "Assumed (implied by the documents)", cls: "bg-warning-soft" },
  missing: { label: "Missing — left blank", cls: "text-muted-foreground" },
};

const fmtNum = (v: unknown) =>
  typeof v === "number" ? v.toLocaleString("vi-VN", { maximumFractionDigits: 4 }) : v == null || v === "" ? "" : String(v);

function Cell({ c }: { c: PlanCell }) {
  const tip = [
    PROV[c.prov]?.label,
    c.f,
    c.pending ? "value computed when opened in Excel" : null,
    c.quote ? `“${c.quote}”` : null,
    c.src,
    c.note,
  ]
    .filter(Boolean)
    .join("\n");
  return (
    <td title={tip} className={cn("max-w-48 truncate px-2 py-1.5 text-right tabular-nums", PROV[c.prov]?.cls)}>
      {c.pending ? <span className="text-muted-foreground">= (Excel)</span> : c.v == null || c.v === "" ? <span className="text-muted-foreground">·</span> : fmtNum(c.v)}
      {c.f && !c.pending && <sup className="ml-0.5 text-[10px] text-primary">ƒ</sup>}
    </td>
  );
}

function Grid({ heads, rows, totals, caption }: { heads: string[]; rows: PlanRow[]; totals?: PlanRow | null; caption: string }) {
  return (
    <div className="overflow-x-auto rounded-md border">
      <table className="w-full whitespace-nowrap text-xs">
        <caption className="sr-only">{caption}</caption>
        <thead className="bg-muted/50 text-left text-muted-foreground">
          <tr>
            <th scope="col" className="px-2 py-1.5 font-medium">Row</th>
            {heads.map((h, i) => (
              <th key={i} scope="col" className="px-2 py-1.5 font-medium">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.row} className="border-t">
              <td className="px-2 py-1.5 font-mono text-muted-foreground">{r.row}</td>
              {r.cells.map((c, i) => (
                <Cell key={i} c={c} />
              ))}
            </tr>
          ))}
          {totals && (
            <tr className="border-t-2 font-semibold">
              <td className="px-2 py-1.5 font-mono text-muted-foreground">{totals.row}</td>
              {totals.cells.map((c, i) => (
                <Cell key={i} c={c} />
              ))}
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function Section({ letter, title, children }: { letter: string; title: string; children: React.ReactNode }) {
  return (
    <section className="flex flex-col gap-2">
      <h4 className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
        {letter}. {title}
      </h4>
      {children}
    </section>
  );
}

function DataPreview({ plan, prev }: { plan: StructureView; prev: StructurePreview }) {
  const cols = prev.columns ?? [];
  const show = cols.map((c, i) => ({ c, i })).filter(({ c }) => !c.hidden);
  const pick = (r: PlanRow): PlanRow => ({ row: r.row, cells: show.map(({ i }) => r.cells[i]) });
  const others = Object.entries(prev.other_formulas ?? {});
  return (
    <>
      <p className="text-xs text-muted-foreground">
        Read back from a dry run of the plan on a copy of the template — this is what the file will contain. Hover a cell for its
        formula or source quote.{" "}
        <span className="text-primary">ƒ = formula</span> · <span className="bg-warning-soft px-1">assumed</span> · “·” = blank
      </p>
      {(prev.rows ?? []).length > 0 ? (
        <Grid
          heads={show.map(({ c }) => `${c.letter} ${c.header}`)}
          rows={(prev.rows ?? []).map(pick)}
          totals={prev.totals ? pick(prev.totals) : null}
          caption={`Data preview of ${plan.sheet}`}
        />
      ) : (
        <p className="text-xs text-muted-foreground">No row values were found with a quote in the sources; rows stay blank.</p>
      )}
      {(plan.overflow ?? 0) > 0 && (
        <p className="text-xs text-destructive">{plan.overflow} record(s) do not fit into the template's rows and are not in the file.</p>
      )}
      {others.length > 0 && (
        <ul className="text-xs text-muted-foreground">
          {others.map(([ref, o]) => (
            <li key={ref}>
              <span className="font-mono">{ref}</span> <span className="font-mono">{o.f}</span> → {o.v == null ? "computed in Excel" : fmtNum(o.v)}
            </li>
          ))}
        </ul>
      )}
      {(prev.sheets ?? []).map((s) => (
        <div key={s.name} className="flex flex-col gap-1">
          <div className="text-sm font-medium">
            New sheet “{s.name}” <span className="text-xs font-normal text-muted-foreground">· {s.reason || s.purpose}</span>
          </div>
          <Grid heads={s.columns.map((c) => `${c.letter} ${c.header}`)} rows={s.rows} totals={s.totals} caption={`Preview of ${s.name}`} />
        </div>
      ))}
    </>
  );
}

/** The AI's reading of the documents, its proposed changes to the template (with reasons) and a dry-run preview of the data. */
export function StructurePlan({ plan }: { plan: StructureView }) {
  const u = plan.understanding;
  const prev = plan.preview as StructurePreview | undefined;
  return (
    <div className="flex flex-col gap-4">
      {u && (
        <Section letter="A" title="What the documents say">
          <div className="rounded-md border p-3 text-sm">
            <div className="mb-1 flex flex-wrap items-center gap-2">
              <Badge variant={u.fit === "fits" ? "outline" : u.fit === "partial" ? "secondary" : "destructive"}>{FIT[u.fit] ?? u.fit}</Badge>
              {u.fit_reason && <span className="text-xs text-muted-foreground">{u.fit_reason}</span>}
            </div>
            {u.document_summary && <p>{u.document_summary}</p>}
            {u.template_purpose && <p className="text-xs text-muted-foreground">Template: {u.template_purpose}</p>}
            {u.report_intent && <p className="text-xs text-muted-foreground">Report wanted: {u.report_intent}</p>}
            {u.entities.length > 0 && <p className="text-xs">Records by: {u.entities.join(", ")}</p>}
            {u.derivable.length > 0 && <p className="text-xs">Computed with formulas: {u.derivable.join(", ")}</p>}
            {u.assumptions.length > 0 && <p className="text-xs text-warning">Assumptions: {u.assumptions.join("; ")}</p>}
          </div>
        </Section>
      )}

      <Section letter="B" title={`Template vs proposed — “${plan.sheet}”`}>
        {!plan.applied && (
          <p className="text-xs text-muted-foreground">You chose to keep the template's columns; only matching values are filled.</p>
        )}
        {plan.template_columns && plan.template_columns.length > 0 && (
          <p className="text-xs text-muted-foreground">
            Template now:{" "}
            {plan.template_columns.map((c) => `${c.letter} ${c.header}${c.formula ? " (=…)" : ""}`).join(" · ")}
          </p>
        )}
        <div className="overflow-x-auto rounded-md border">
          <table className="w-full text-sm">
            <caption className="sr-only">Proposed columns of {plan.sheet}</caption>
            <thead className="bg-muted/50 text-left text-xs text-muted-foreground">
              <tr>
                <th scope="col" className="px-3 py-2 font-medium">Col</th>
                <th scope="col" className="px-3 py-2 font-medium">Header · why</th>
                <th scope="col" className="px-3 py-2 font-medium">Change</th>
                <th scope="col" className="px-3 py-2 font-medium">Formula / format</th>
              </tr>
            </thead>
            <tbody>
              {plan.columns.map((c) => (
                <tr key={c.letter} className={"border-t align-top" + (c.changes.includes("hidden") ? " opacity-60" : "")}>
                  <td className="px-3 py-2 font-mono text-xs">{c.letter}</td>
                  <td className="px-3 py-2">
                    <div className="font-medium">{c.header}</div>
                    {c.from && c.from !== c.header && <div className="text-xs text-muted-foreground">was “{c.from}”</div>}
                    {c.reason && <div className="text-xs italic text-muted-foreground">{c.reason}</div>}
                  </td>
                  <td className="px-3 py-2">
                    <div className="flex flex-wrap gap-1">
                      {c.changes.map((k) => (
                        <Badge key={k} variant={CHANGE[k]?.variant ?? "outline"}>
                          {CHANGE[k]?.label ?? k}
                        </Badge>
                      ))}
                    </div>
                  </td>
                  <td className="max-w-72 px-3 py-2 text-xs">
                    {c.formula && <div className="break-all font-mono">={c.formula.replace(/^=/, "")}</div>}
                    {c.fmt && <div className="text-muted-foreground">format {c.fmt}</div>}
                    {c.default != null && c.default !== "" && <div className="text-muted-foreground">default {String(c.default)}</div>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {plan.removed.length > 0 && (
          <ul className="list-disc pl-5 text-sm">
            {plan.removed.map((d) => (
              <li key={d.header}>
                Removed “{d.header}” — {MODE[d.mode]}
                {d.used_by.length > 0 && <span className="text-xs text-muted-foreground"> · used by {d.used_by.slice(0, 4).join(", ")}</span>}
              </li>
            ))}
          </ul>
        )}

        {(plan.labels ?? []).length > 0 && (
          <ul className="text-sm">
            {(plan.labels ?? []).map((l) => (
              <li key={l.cell}>
                <span className="font-mono text-xs text-muted-foreground">{l.cell}</span> “{l.current}” → <span className="font-medium">“{l.new}”</span>
                {l.reason && <span className="text-xs italic text-muted-foreground"> · {l.reason}</span>}
              </li>
            ))}
          </ul>
        )}

        {plan.kv.length > 0 && (
          <ul className="text-sm">
            {plan.kv.map((k) => (
              <li key={k.cell} title={k.quote ?? undefined}>
                <span className="font-mono text-xs text-muted-foreground">{k.cell}</span> {k.label}: <span className="font-medium">{String(k.value)}</span>
                <span className="text-xs text-muted-foreground"> · {k.source ?? "—"}</span>
              </li>
            ))}
          </ul>
        )}

        {(plan.sheets_proposed ?? 0) > 0 && !plan.sheets_on && (
          <p className="text-xs text-muted-foreground">{plan.sheets_proposed} proposed sheet(s) will not be added.</p>
        )}
      </Section>

      <Section letter="C" title="Data preview">
        {prev && "ok" in prev && !prev.ok ? (
          <p className="text-sm text-destructive">The plan cannot be applied safely: {prev.errors.join("; ")}</p>
        ) : prev && "ok" in prev ? (
          <DataPreview plan={plan} prev={prev} />
        ) : (
          <p className="text-xs text-muted-foreground">No preview available.</p>
        )}
        {(plan.missing ?? []).length > 0 && (
          <div className="rounded-md border border-warning/30 bg-warning-soft p-2 text-xs">
            <span className="font-semibold">Not in the documents (left blank, never invented): </span>
            {(plan.missing ?? []).map((m) => m.what).join("; ")}
          </div>
        )}
      </Section>
    </div>
  );
}
