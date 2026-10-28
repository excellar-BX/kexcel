import {
  isError,
  isFormula,
  parseLiteral,
  shiftFormula,
  type CellPatch,
  type Scalar,
  type Workbook,
} from "@/lib/engine";

/** An inclusive block of cells. */
export interface Range {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

const mod = (n: number, m: number) => ((n % m) + m) % m;

/**
 * Cell patches that extend `src` to fill `dst`, which contains `src` and
 * grows it in one direction. In "series" mode a run of two or more numbers
 * continues its trend (1, 2 → 3, 4); otherwise the source repeats, with
 * formulas shifted to their new position, like dragging the fill handle.
 */
export function fillPatches(
  wb: Workbook,
  sheet: string,
  src: Range,
  dst: Range,
  mode: "series" | "copy",
): CellPatch[] {
  const vertical = dst.r0 !== src.r0 || dst.r1 !== src.r1;
  const patches: CellPatch[] = [];

  // Work along "lines": columns when filling vertically, rows when filling sideways.
  const lineStart = vertical ? src.c0 : src.r0;
  const lineEnd = vertical ? src.c1 : src.r1;
  const srcFrom = vertical ? src.r0 : src.c0;
  const srcTo = vertical ? src.r1 : src.c1;
  const dstFrom = vertical ? dst.r0 : dst.c0;
  const dstTo = vertical ? dst.r1 : dst.c1;
  const n = srcTo - srcFrom + 1;
  const at = (line: number, pos: number) => (vertical ? { row: pos, col: line } : { row: line, col: pos });

  for (let line = lineStart; line <= lineEnd; line++) {
    const sources = Array.from({ length: n }, (_, i) => {
      const a = at(line, srcFrom + i);
      return { ...a, raw: wb.getRaw(sheet, a.row, a.col), format: wb.getFormat(sheet, a.row, a.col) };
    });
    const numbers = sources.map((s) => (isFormula(s.raw) ? null : parseLiteral(s.raw)));
    const series = mode === "series" && n >= 2 && numbers.every((v) => typeof v === "number");
    const first = numbers[0] as number;
    const step = series ? ((numbers[n - 1] as number) - first) / (n - 1) : 0;

    for (let pos = dstFrom; pos <= dstTo; pos++) {
      if (pos >= srcFrom && pos <= srcTo) continue;
      const target = at(line, pos);
      const source = sources[mod(pos - srcFrom, n)];
      let raw: string;
      if (series) {
        raw = String(Number((first + step * (pos - srcFrom)).toPrecision(15)));
      } else {
        raw = shiftFormula(source.raw, target.row - source.row, target.col - source.col);
      }
      patches.push({ sheet, ...target, raw, format: source.format ?? null });
    }
  }
  return patches;
}

const rank = (v: Scalar) =>
  v === null || v === "" ? 4 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3;

/** Orders values like a spreadsheet sort: numbers, text, booleans, errors, then blanks. */
function compareForSort(a: Scalar, b: Scalar): number {
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "number") return a - (b as number);
  if (typeof a === "string") return a.localeCompare(b as string, undefined, { sensitivity: "base", numeric: true });
  if (typeof a === "boolean") return Number(a) - Number(b);
  if (isError(a) && isError(b)) return a.code.localeCompare(b.code);
  return 0;
}

/** Cell patches that sort the rows of `range` by the values in column `byCol`. */
export function sortPatches(wb: Workbook, sheet: string, range: Range, byCol: number, ascending: boolean): CellPatch[] {
  const rows = Array.from({ length: range.r1 - range.r0 + 1 }, (_, i) => range.r0 + i);
  const order = [...rows].sort((x, y) => {
    const a = wb.getValue(sheet, x, byCol);
    const b = wb.getValue(sheet, y, byCol);
    const blankA = rank(a) === 4;
    const blankB = rank(b) === 4;
    // Blanks stay at the bottom in both directions.
    if (blankA || blankB) return Number(blankA) - Number(blankB) || x - y;
    const c = compareForSort(a, b);
    return (ascending ? c : -c) || x - y;
  });

  const patches: CellPatch[] = [];
  order.forEach((srcRow, i) => {
    const row = range.r0 + i;
    for (let col = range.c0; col <= range.c1; col++) {
      patches.push({
        sheet,
        row,
        col,
        raw: shiftFormula(wb.getRaw(sheet, srcRow, col), row - srcRow, 0),
        format: wb.getFormat(sheet, srcRow, col) ?? null,
      });
    }
  });
  return patches;
}
