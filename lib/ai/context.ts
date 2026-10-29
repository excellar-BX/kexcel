import { addressToName, colToLetters, formatValue, isFormula, type Workbook } from "@/lib/engine";
import type { Range } from "@/lib/sheet-ops";

export const MAX_CONTEXT_ROWS = 300;
export const MAX_CONTEXT_COLS = 30;
const MAX_CONTEXT_CHARS = 60_000;

export interface SheetContext {
  text: string;
  /** Set when the sheet was larger than what was sent, so the UI can say so. */
  truncated: { shownRows: number; totalRows: number; shownCols: number; totalCols: number } | null;
}

/**
 * Describes a sheet for the model: the tab names, the selection, and every
 * filled cell of the used area as "A1: value" (formulas as "=... → result").
 */
export function describeSheet(wb: Workbook, sheet: string, selection: Range, active: { row: number; col: number }): SheetContext {
  const lines: string[] = [];
  const names = wb.sheets.map((s) => s.name);
  lines.push(`Workbook tabs: ${names.map((n) => JSON.stringify(n)).join(", ")}`);
  lines.push(`Current tab: ${JSON.stringify(wb.sheetName(sheet))}`);
  lines.push(`Active cell: ${addressToName(active)}`);
  const sel =
    selection.r0 === selection.r1 && selection.c0 === selection.c1
      ? addressToName({ row: selection.r0, col: selection.c0 })
      : `${addressToName({ row: selection.r0, col: selection.c0 })}:${addressToName({ row: selection.r1, col: selection.c1 })}`;
  lines.push(`Selection: ${sel}`);

  const used = wb.usedBounds(sheet);
  if (!used) {
    lines.push("The tab is empty.");
    return { text: lines.join("\n"), truncated: null };
  }

  const rows = Math.min(used.rows, MAX_CONTEXT_ROWS);
  const cols = Math.min(used.cols, MAX_CONTEXT_COLS);
  lines.push(`Used area: A1:${colToLetters(used.cols - 1)}${used.rows}`);
  lines.push("Cells (empty cells omitted; formulas shown with their current result):");

  let chars = lines.join("\n").length;
  let shownRows = rows;
  outer: for (let r = 0; r < rows; r++) {
    const cells: string[] = [];
    for (let c = 0; c < cols; c++) {
      const raw = wb.getRaw(sheet, r, c);
      if (raw === "") continue;
      const shown = formatValue(wb.getValue(sheet, r, c), wb.getFormat(sheet, r, c));
      cells.push(`${addressToName({ row: r, col: c })}: ${isFormula(raw) ? `${raw} → ${shown}` : JSON.stringify(shown)}`);
    }
    if (cells.length === 0) continue;
    const line = cells.join(" | ");
    if (chars + line.length > MAX_CONTEXT_CHARS) {
      shownRows = r;
      break outer;
    }
    lines.push(line);
    chars += line.length + 1;
  }

  const notes = [...wb.commentKeys(sheet)].slice(0, 50).map((k) => {
    const [row, col] = k.split(",").map(Number);
    return `${addressToName({ row, col })}: ${JSON.stringify(wb.getComment(sheet, row, col))}`;
  });
  if (notes.length) lines.push(`Cell comments: ${notes.join(" | ")}`);

  const truncated =
    shownRows < used.rows || cols < used.cols
      ? { shownRows, totalRows: used.rows, shownCols: cols, totalCols: used.cols }
      : null;
  if (truncated) {
    lines.push(
      `Note: only rows 1-${truncated.shownRows} and columns A-${colToLetters(truncated.shownCols - 1)} are shown; the tab has ${truncated.totalRows} rows and ${truncated.totalCols} columns.`,
    );
  }
  return { text: lines.join("\n"), truncated };
}
