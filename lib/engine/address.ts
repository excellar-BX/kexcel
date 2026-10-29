import type { CellAddress } from "./types";

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function colToLetters(col: number): string {
  let s = "";
  let n = col + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" -> 0, "AA" -> 26 */
export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function addressToName({ row, col }: CellAddress): string {
  return `${colToLetters(col)}${row + 1}`;
}

const NAME_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

export function nameToAddress(name: string): CellAddress | null {
  const m = NAME_RE.exec(name);
  if (!m) return null;
  const row = parseInt(m[2], 10) - 1;
  if (row < 0) return null;
  return { row, col: lettersToCol(m[1]) };
}

/** Map key for a cell. Rows and columns are 0-based. */
export const cellKey = (row: number, col: number) => `${row},${col}`;

export function parseKey(key: string): CellAddress {
  const i = key.indexOf(",");
  return { row: Number(key.slice(0, i)), col: Number(key.slice(i + 1)) };
}

/** "A1:C9" for a range; a single cell is written as "A1". */
export function rangeToName(r: { r0: number; r1: number; c0: number; c1: number }): string {
  const a = addressToName({ row: r.r0, col: r.c0 });
  return r.r0 === r.r1 && r.c0 === r.c1 ? a : `${a}:${addressToName({ row: r.r1, col: r.c1 })}`;
}

/** Reads "A1:C9" or "A1" (in either corner order); null if it isn't a range. */
export function nameToRange(text: string): { r0: number; r1: number; c0: number; c1: number } | null {
  const [x, y, extra] = text.trim().toUpperCase().split(":");
  if (extra !== undefined) return null;
  const a = nameToAddress(x ?? "");
  const b = y === undefined ? a : nameToAddress(y);
  if (!a || !b) return null;
  return { r0: Math.min(a.row, b.row), r1: Math.max(a.row, b.row), c0: Math.min(a.col, b.col), c1: Math.max(a.col, b.col) };
}
