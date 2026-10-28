// Sheets are kept in this browser's localStorage until accounts and cloud sync exist.
// Every access is wrapped because storage can be unavailable (private mode, quota, blocked).

const PREFIX = "kexcel:sheet:";

export interface SavedSheet {
  title: string;
  /** Raw input keyed by cell name, e.g. { A1: "=SUM(B1:B3)" }. */
  cells: Record<string, string>;
  /** Custom widths keyed by column index. */
  colWidths: Record<string, number>;
  updatedAt: number;
}

export interface SheetSummary {
  id: string;
  title: string;
  updatedAt: number;
}

export function loadSheet(id: string): SavedSheet | null {
  try {
    const text = localStorage.getItem(PREFIX + id);
    if (!text) return null;
    const data = JSON.parse(text) as Partial<SavedSheet>;
    return {
      title: data.title ?? "Untitled sheet",
      cells: data.cells ?? {},
      colWidths: data.colWidths ?? {},
      updatedAt: data.updatedAt ?? 0,
    };
  } catch {
    return null;
  }
}

/** Returns false when the browser refused the write. */
export function saveSheet(id: string, data: Omit<SavedSheet, "updatedAt">): boolean {
  try {
    localStorage.setItem(PREFIX + id, JSON.stringify({ ...data, updatedAt: Date.now() }));
    return true;
  } catch {
    return false;
  }
}

export function deleteSheet(id: string) {
  try {
    localStorage.removeItem(PREFIX + id);
  } catch {
    // Nothing to do if storage is unavailable.
  }
}

export function listSheets(): SheetSummary[] {
  const out: SheetSummary[] = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key?.startsWith(PREFIX)) continue;
      const id = key.slice(PREFIX.length);
      const data = loadSheet(id);
      if (data) out.push({ id, title: data.title, updatedAt: data.updatedAt });
    }
  } catch {
    return [];
  }
  return out.sort((a, b) => b.updatedAt - a.updatedAt);
}

export function newSheetId(): string {
  return Math.random().toString(36).slice(2, 10);
}
