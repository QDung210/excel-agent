import { Badge } from "@/components/ui/badge";
import type { MappingRow } from "@/lib/types";

const STATUS: Record<MappingRow["status"], { label: string; variant: "secondary" | "outline" | "destructive" }> = {
  auto: { label: "Matched", variant: "secondary" },
  confirm: { label: "Confirm", variant: "outline" },
  missing: { label: "No source", variant: "destructive" },
};

const VIA: Record<string, string> = {
  label: "same label",
  header: "same header",
  llm: "AI",
  fuzzy: "similar label",
  user: "your choice",
  sequence: "row number",
};

/** Where each template input will take its value from: the AI's decisions, shown before anything is written. */
export function MappingTable({ rows }: { rows: MappingRow[] }) {
  if (!rows.length) return null;
  return (
    <div className="overflow-x-auto rounded-md border">
      <table className="w-full text-sm">
        <caption className="sr-only">Field mapping</caption>
        <thead className="bg-muted/50 text-left text-xs text-muted-foreground">
          <tr>
            <th scope="col" className="px-3 py-2 font-medium">Template field</th>
            <th scope="col" className="px-3 py-2 font-medium">From the source</th>
            <th scope="col" className="px-3 py-2 font-medium">Value</th>
            <th scope="col" className="px-3 py-2 font-medium">Status</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.target} className="border-t align-top">
              <td className="px-3 py-2">
                <div className="font-medium">{r.field}</div>
                <div className="font-mono text-xs text-muted-foreground">{r.target}</div>
              </td>
              <td className="px-3 py-2">
                {r.source ?? <span className="text-muted-foreground">—</span>}
                {r.via && <div className="text-xs text-muted-foreground">via {VIA[r.via] ?? r.via}</div>}
                {r.reason && <div className="text-xs italic text-muted-foreground">{r.reason}</div>}
              </td>
              <td className="max-w-48 truncate px-3 py-2" title={r.value == null ? undefined : String(r.value)}>
                {r.value == null ? <span className="text-muted-foreground">blank</span> : String(r.value)}
              </td>
              <td className="px-3 py-2">
                <Badge variant={STATUS[r.status].variant}>{STATUS[r.status].label}</Badge>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
