/** 0-based column index -> letters (0 -> A, 26 -> AA). */
export function colToLetters(c: number): string {
  let n = c + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 0-based (r, c) -> "B3". */
export function toA1(r: number, c: number): string {
  return `${colToLetters(c)}${r + 1}`;
}

/** 0-based inclusive range -> "B2:C4" (or "B2" for one cell). */
export function rangeToA1(r1: number, c1: number, r2: number, c2: number): string {
  const a = toA1(Math.min(r1, r2), Math.min(c1, c2));
  const b = toA1(Math.max(r1, r2), Math.max(c1, c2));
  return a === b ? a : `${a}:${b}`;
}

/** Sheet-qualified reference, quoting names with spaces: 'High Level Plan'!B2 */
export function qualify(sheet: string, ref: string): string {
  const quoted = /^[A-Za-z0-9_]+$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return `${quoted}!${ref}`;
}
