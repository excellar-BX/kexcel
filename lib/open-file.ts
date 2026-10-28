import { colToLetters, type WorkbookJSON } from "@/lib/engine";
import { newSheetId, saveSheet } from "@/lib/storage";
import { parseCSV } from "@/lib/tsv";
import { readXlsx, type ImportLimits } from "@/lib/xlsx/read";

export const OPEN_FILE_ACCEPT =
  ".xlsx,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const NOTES_PREFIX = "kexcel:open-notes:";

/** Reads an .xlsx or .csv file into a workbook. Throws with a readable message otherwise. */
async function readFile(file: File, limits: ImportLimits): Promise<{ workbook: WorkbookJSON; notes: string[] }> {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx")) return readXlsx(new Uint8Array(await file.arrayBuffer()), limits);
  if (name.endsWith(".csv") || file.type === "text/csv") {
    const rows = parseCSV(await file.text());
    const cells: Record<string, string> = {};
    let clipped = false;
    rows.forEach((r, row) =>
      r.forEach((raw, col) => {
        if (row >= limits.maxRows || col >= limits.maxCols) clipped = true;
        else if (raw !== "") cells[`${colToLetters(col)}${row + 1}`] = raw;
      }),
    );
    const notes = clipped
      ? [`Only the first ${limits.maxRows.toLocaleString()} rows and ${limits.maxCols} columns were imported.`]
      : [];
    return { workbook: { sheets: [{ id: "s1", name: "Sheet1", cells }] }, notes };
  }
  if (name.endsWith(".xls")) throw new Error("Kexcel can't open old .xls files. Save it as .xlsx in Excel first.");
  throw new Error("Kexcel can open .xlsx and .csv files.");
}

/**
 * Opens a file as a new Kexcel spreadsheet saved in this browser and returns
 * its id. Notes about anything that didn't carry over are kept for the
 * editor to show once it opens.
 */
export async function openFileAsSpreadsheet(file: File, limits: ImportLimits): Promise<string> {
  const { workbook, notes } = await readFile(file, limits);
  const id = newSheetId();
  const title = file.name.replace(/\.[^.]+$/, "") || "Imported spreadsheet";
  if (!saveSheet(id, { title, workbook })) {
    throw new Error("This file is too large to keep in this browser's storage.");
  }
  if (notes.length) {
    try {
      sessionStorage.setItem(NOTES_PREFIX + id, JSON.stringify(notes));
    } catch {
      // The notes just won't be shown.
    }
  }
  return id;
}

/** Notes left by openFileAsSpreadsheet for this spreadsheet, removed once read. */
export function takeOpenNotes(id: string): string[] {
  try {
    const text = sessionStorage.getItem(NOTES_PREFIX + id);
    if (!text) return [];
    sessionStorage.removeItem(NOTES_PREFIX + id);
    return JSON.parse(text) as string[];
  } catch {
    return [];
  }
}
