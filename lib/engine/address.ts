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
