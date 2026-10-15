import { memo, useEffect, useRef } from "react";
import { Skeleton } from "@/components/ui/skeleton";
import { cn } from "@/lib/utils";
import type { Preview, PreviewCell, PreviewSheet } from "@/lib/types";

const DEFAULT_COL = 73;
const DEFAULT_ROW = 20;
const HEAD_H = 18; // column-letter strip, in sheet pixels
const HEAD_W = 28; // row-number strip, in sheet pixels
/** How many sheet pixels the thumbnail spans horizontally (the rest is scaled to fit). */
const SPAN = 460;

function colName(c: number): string {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function activeSheet(p: Preview): PreviewSheet | undefined {
  return p.sheets.find((s) => s.status === 1 && !s.hide) ?? p.sheets.find((s) => !s.hide) ?? p.sheets[0];
}

function cellText(v: PreviewCell): string {
  const t = v.m ?? v.v;
  return t === null || t === undefined ? "" : String(t);
}

/** Paints the top-left of the active sheet (fills, text, merges) like a screenshot of the grid. */
function paint(canvas: HTMLCanvasElement, sheet: PreviewSheet, w: number, h: number) {
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.round(w * dpr);
  canvas.height = Math.round(h * dpr);
  const ctx = canvas.getContext("2d");
  if (!ctx) return;
  const scale = w / SPAN;
  const spanH = h / scale;
  ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);

  const cfg = sheet.config ?? {};
  const xs: number[] = []; // left edge of each visible column
  const ws: number[] = [];
  for (let c = 0, x = HEAD_W; x < SPAN && c < 200; c++) {
    const cw = cfg.colhidden?.[c] !== undefined ? 0 : (cfg.columnlen?.[c] ?? DEFAULT_COL);
    xs.push(x);
    ws.push(cw);
    x += cw;
  }
  const ys: number[] = [];
  const hs: number[] = [];
  for (let r = 0, y = HEAD_H; y < spanH && r < 500; r++) {
    const rh = cfg.rowhidden?.[r] !== undefined ? 0 : (cfg.rowlen?.[r] ?? DEFAULT_ROW);
    ys.push(y);
    hs.push(rh);
    y += rh;
  }
  const nC = xs.length;
  const nR = ys.length;
  const spanW = (c: number, cs = 1) => ws.slice(c, c + cs).reduce((a, b) => a + b, 0);
  const spanR = (r: number, rs = 1) => hs.slice(r, r + rs).reduce((a, b) => a + b, 0);

  const cells = sheet.celldata.filter((d) => d.v && d.r < nR && d.c < nC) as Array<{ r: number; c: number; v: PreviewCell }>;
  const covered = new Set<string>(); // non-anchor cells of merges
  const filled = new Set<string>();
  for (const { r, c, v } of cells) {
    if (v.mc && v.mc.rs === undefined) covered.add(`${r}_${c}`);
    if (cellText(v)) filled.add(`${r}_${c}`);
  }

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, SPAN, spanH);

  // fills
  for (const { r, c, v } of cells) {
    if (!v.bg || covered.has(`${r}_${c}`)) continue;
    ctx.fillStyle = v.bg;
    ctx.fillRect(xs[c], ys[r], spanW(c, v.mc?.cs), spanR(r, v.mc?.rs));
  }

  // gridlines
  ctx.strokeStyle = "#e3e3e3";
  ctx.lineWidth = 1 / scale;
  ctx.beginPath();
  for (let c = 0; c <= nC; c++) {
    const x = c < nC ? xs[c] : xs[nC - 1] + ws[nC - 1];
    ctx.moveTo(x, HEAD_H);
    ctx.lineTo(x, spanH);
  }
  for (let r = 0; r <= nR; r++) {
    const y = r < nR ? ys[r] : ys[nR - 1] + hs[nR - 1];
    ctx.moveTo(HEAD_W, y);
    ctx.lineTo(SPAN, y);
  }
  ctx.stroke();

  // text
  ctx.textBaseline = "middle";
  for (const { r, c, v } of cells) {
    const text = cellText(v);
    if (!text || covered.has(`${r}_${c}`) || !ws[c] || !hs[r]) continue;
    const fs = ((v.fs ?? 11) * 4) / 3;
    ctx.font = `${v.it ? "italic " : ""}${v.bl ? "bold " : ""}${fs}px Calibri, Arial, sans-serif`;
    ctx.fillStyle = v.fc ?? "#1f1f1f";
    const numeric = typeof v.v === "number";
    const align = v.ht === 0 ? "center" : v.ht === 2 || (v.ht === undefined && numeric) ? "right" : "left";
    let width = spanW(c, v.mc?.cs);
    if (!v.mc && align === "left") {
      // Text overflows into empty neighbours, as in Excel.
      for (let k = c + 1; k < nC && !filled.has(`${r}_${k}`); k++) width += ws[k];
    }
    const top = ys[r];
    const height = spanR(r, v.mc?.rs);
    ctx.save();
    ctx.beginPath();
    ctx.rect(xs[c], top, width, height);
    ctx.clip();
    ctx.textAlign = align;
    const x = align === "left" ? xs[c] + 3 : align === "right" ? xs[c] + width - 3 : xs[c] + width / 2;
    ctx.fillText(text, x, top + height / 2);
    ctx.restore();
  }

  // headers
  ctx.fillStyle = "#f4f4f2";
  ctx.fillRect(0, 0, SPAN, HEAD_H);
  ctx.fillRect(0, 0, HEAD_W, spanH);
  ctx.fillStyle = "#8a8a86";
  ctx.font = "11px Arial, sans-serif";
  ctx.textAlign = "center";
  for (let c = 0; c < nC; c++) if (ws[c]) ctx.fillText(colName(c), xs[c] + ws[c] / 2, HEAD_H / 2);
  for (let r = 0; r < nR; r++) if (hs[r] >= 10) ctx.fillText(String(r + 1), HEAD_W / 2, ys[r] + hs[r] / 2);
  ctx.strokeStyle = "#d6d6d2";
  ctx.beginPath();
  ctx.moveTo(0, HEAD_H);
  ctx.lineTo(SPAN, HEAD_H);
  ctx.moveTo(HEAD_W, 0);
  ctx.lineTo(HEAD_W, spanH);
  ctx.stroke();
}

interface TemplateThumbProps {
  /** undefined = still loading; null = could not be loaded. */
  preview: Preview | null | undefined;
  width: number;
  height: number;
  className?: string;
  label: string;
}

/** Small picture of a template's first screen, drawn from its read-only backend preview. */
export const TemplateThumb = memo(function TemplateThumb({ preview, width, height, className, label }: TemplateThumbProps) {
  const ref = useRef<HTMLCanvasElement>(null);
  const sheet = preview ? activeSheet(preview) : undefined;

  useEffect(() => {
    if (ref.current && sheet) paint(ref.current, sheet, width, height);
  }, [sheet, width, height]);

  const box = cn("shrink-0 overflow-hidden rounded-md border border-border bg-white", className);
  if (preview === undefined) return <Skeleton className={cn("shrink-0 rounded-md", className)} style={{ width, height }} />;
  if (!sheet)
    return (
      <div className={cn(box, "grid place-items-center text-[10px] text-muted-foreground")} style={{ width, height }}>
        No preview
      </div>
    );
  return <canvas ref={ref} role="img" aria-label={label} className={box} style={{ width, height }} />;
});
