import type { Op } from "@fortune-sheet/core";
import { toA1 } from "./a1";
import type { Action, PreviewCell } from "./types";

/**
 * Translate FortuneSheet `onOp` patches into the backend's allowlisted operations.
 *
 * Only single-cell value changes are saveable: `replace` / `add` / `remove` on
 *   ["data", r, c]                 (whole cell object, or null)
 *   ["data", r, c, "v" | "m" | "ct" | "f" | ...value-ish keys]
 * Anything else (styles, merges, row/col structure, sheets, config) is rejected and the grid is
 * reloaded from the canonical preview.
 */

/** Keys whose change is a pure value change. "f" is included so a formula entry is detected and refused, not missed. */
const VALUE_KEYS = new Set(["v", "m", "ct", "f", "qp", "spl"]);
const STYLE_KEYS = new Set(["bl", "it", "ff", "fs", "fc", "bg", "ht", "vt", "tb", "tr", "rt", "un", "cl", "ps", "lo", "hl"]);
/** Sheet-level bookkeeping FortuneSheet emits during normal navigation/editing; never persisted. */
const IGNORED_ROOTS = new Set([
  "luckysheet_select_save",
  "luckysheet_selection_range",
  "calcChain",
  "dynamicArray_compute",
  "dynamicArray",
  "zoomRatio",
  "scrollLeft",
  "scrollTop",
]);

export type CellPatch = { sheetId: string; r: number; c: number };

export interface TranslateResult {
  /** Cells whose value changed (deduplicated, in order). */
  touched: CellPatch[];
  /** Human-readable reasons for ops that cannot be saved. */
  unsupported: string[];
}

export function cellKey(sheetId: string, r: number, c: number) {
  return `${sheetId}:${r}:${c}`;
}

const isIdx = (x: unknown): x is number => typeof x === "number" && Number.isInteger(x) && x >= 0;

export function classifyOps(ops: Op[], fallbackSheetId: string): TranslateResult {
  const touched: CellPatch[] = [];
  const seen = new Set<string>();
  const unsupported = new Set<string>();

  for (const op of ops) {
    const sheetId = op.id ?? fallbackSheetId;
    const p = op.path ?? [];
    switch (op.op) {
      case "insertRowCol":
      case "deleteRowCol":
        unsupported.add("inserting or deleting rows/columns is not supported");
        continue;
      case "addSheet":
      case "deleteSheet":
        unsupported.add("adding or deleting sheets is not supported");
        continue;
    }
    const root = p[0];
    if (typeof root === "string" && IGNORED_ROOTS.has(root)) continue;

    if (root === "data") {
      if (p.length >= 3 && isIdx(p[1]) && isIdx(p[2])) {
        const key = p[3];
        if (p.length === 3 || (typeof key === "string" && VALUE_KEYS.has(key))) {
          const k = cellKey(sheetId, p[1], p[2]);
          if (!seen.has(k)) {
            seen.add(k);
            touched.push({ sheetId, r: p[1], c: p[2] });
          }
        } else if (key === "mc") unsupported.add("merging cells is not supported");
        else if (typeof key === "string" && STYLE_KEYS.has(key)) unsupported.add("formatting changes are not supported");
        else unsupported.add(`cell property "${String(key)}" is not supported`);
      } else {
        unsupported.add("bulk sheet data changes are not supported");
      }
      continue;
    }
    if (root === "config") {
      const sub = String(p[1] ?? "");
      if (sub === "merge") unsupported.add("merging cells is not supported");
      else if (sub === "borderInfo") unsupported.add("border changes are not supported");
      else if (sub === "rowlen" || sub === "columnlen" || sub === "customHeight" || sub === "customWidth")
        unsupported.add("row height / column width changes are not supported");
      else if (sub === "rowhidden" || sub === "colhidden") unsupported.add("hiding rows/columns is not supported");
      else unsupported.add(`sheet setting "${sub || "config"}" is not supported`);
      continue;
    }
    if (root === "name") unsupported.add("renaming sheets is not supported");
    else if (root === "frozen") unsupported.add("freezing panes is not supported");
    else if (root === "hyperlink") unsupported.add("hyperlinks are not supported");
    else if (root === "images") unsupported.add("images are not supported");
    else if (typeof root === "string" && /conditionformat|alternateformat/.test(root))
      unsupported.add("conditional formatting is not supported");
    else unsupported.add(`"${String(root ?? op.op)}" changes are not supported`);
  }
  return { touched, unsupported: [...unsupported] };
}

/** Apply one op to a cell snapshot so we can track the final cell without reading grid internals. */
export function applyOpToCell(cell: PreviewCell | null, op: Op): PreviewCell | null {
  const p = op.path;
  if (p.length === 3) return op.op === "remove" ? null : ((op.value ?? null) as PreviewCell | null);
  const next: Record<string, unknown> = { ...(cell ?? {}) };
  if (p.length === 4) {
    const key = String(p[3]);
    if (op.op === "remove") delete next[key];
    else next[key] = op.value;
    return next as PreviewCell;
  }
  // deeper path, e.g. ["data", r, c, "ct", "s"] -> write nested
  let target = next;
  for (let i = 3; i < p.length - 1; i++) {
    const k = String(p[i]);
    const child = target[k];
    target[k] = Array.isArray(child) ? [...child] : { ...((child as object) ?? {}) };
    target = target[k] as Record<string, unknown>;
  }
  const last = String(p[p.length - 1]);
  if (op.op === "remove") delete target[last];
  else target[last] = op.value;
  return next as PreviewCell;
}

export type CellTranslation = { action: Action } | { unsupported: string } | { noop: true };

function inlineText(cell: PreviewCell): string | null {
  if (cell.ct?.t === "inlineStr" && Array.isArray(cell.ct.s)) return cell.ct.s.map((s) => s?.v ?? "").join("");
  return null;
}

/** Comparable scalar value of a cell, for no-op detection. */
export function cellScalar(cell: PreviewCell | null | undefined): string | number | boolean | null {
  if (!cell) return null;
  const inline = inlineText(cell);
  if (inline !== null) return inline;
  const v = cell.v;
  return v === undefined || v === "" ? null : v;
}

function styleSignature(cell: PreviewCell | null | undefined): string {
  if (!cell) return "";
  const rec = cell as Record<string, unknown>;
  return JSON.stringify([...STYLE_KEYS].map((k) => rec[k] ?? null).concat([rec.mc ? "mc" : null]));
}

/** Turn the final state of one edited cell into a set_value / clear_value action. */
export function translateCell(
  sheetName: string,
  r: number,
  c: number,
  next: PreviewCell | null,
  canonical: PreviewCell | null | undefined,
): CellTranslation {
  const cell = A1Safe(r, c);
  if (next && (next.f || (typeof next.v === "string" && next.v.startsWith("=")))) {
    return { unsupported: "formulas cannot be entered here yet" };
  }
  if (canonical?.f) return { unsupported: `${cell} contains a formula and cannot be overwritten` };
  // A whole-cell replace (e.g. paste) can carry formatting; refuse silently dropping it.
  if (next && styleSignature(next) !== styleSignature(canonical) && Object.keys(next).some((k) => STYLE_KEYS.has(k))) {
    return { unsupported: "formatting changes are not supported" };
  }

  const value = cellScalar(next);
  if (value === cellScalar(canonical)) return { noop: true };
  if (value === null) return { action: { type: "clear_value", sheet: sheetName, cell } };

  const t = next?.ct?.t;
  if (inlineText(next!) !== null) return { action: { type: "set_value", sheet: sheetName, cell, value: String(value), value_type: "string" } };
  if (t === "d" && typeof value === "number") return { action: { type: "set_value", sheet: sheetName, cell, value, value_type: "date" } };
  if (t === "s" || next?.ct?.fa === "@") return { action: { type: "set_value", sheet: sheetName, cell, value: String(value), value_type: "string" } };
  if (typeof value === "boolean" || t === "b") {
    const b = typeof value === "boolean" ? value : /^true$/i.test(String(value));
    return { action: { type: "set_value", sheet: sheetName, cell, value: b, value_type: "bool" } };
  }
  if (typeof value === "number") return { action: { type: "set_value", sheet: sheetName, cell, value, value_type: "number" } };
  return { action: { type: "set_value", sheet: sheetName, cell, value: String(value), value_type: "auto" } };
}

function A1Safe(r: number, c: number) {
  return toA1(r, c);
}
