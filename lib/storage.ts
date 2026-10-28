// Sheets are kept in this browser's localStorage until accounts and cloud sync exist.
// Every access is wrapped because storage can be unavailable (private mode, quota, blocked).

import type { WorkbookJSON } from "@/lib/engine";

const PREFIX = "kexcel:sheet:";

export interface SavedSheet {
  title: string;
  workbook: WorkbookJSON;
  /** The tab that was open last. */
  activeSheet?: string;
  updatedAt: number;
}

export interface SheetSummary {
  id: string;
  title: string;
  updatedAt: number;
}

/** The first saved format held one sheet's cells and widths at the top level. */
interface LegacySheet {
  title?: string;
  cells?: Record<string, string>;
  colWidths?: Record<string, number>;
  updatedAt?: number;
}

export function loadSheet(id: string): SavedSheet | null {
  try {
    const text = localStorage.getItem(PREFIX + id);
    if (!text) return null;
    const data = JSON.parse(text) as Partial<SavedSheet> & LegacySheet;
    const workbook: WorkbookJSON = data.workbook ?? {
      sheets: [{ id: "s1", name: "Sheet1", cells: data.cells ?? {}, colWidths: data.colWidths ?? {} }],
    };
    return {
      title: data.title ?? "Untitled sheet",
      workbook,
      activeSheet: data.activeSheet,
      updatedAt: data.updatedAt ?? 0,
    };
  } catch {
    return null;
  }
}

/** Returns false when the browser refused the write, e.g. because storage is full. */
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
