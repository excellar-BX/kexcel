// Server-only: spreadsheets saved to an account.
import type { WorkbookJSON } from "@/lib/engine";
import { db } from "./db";

/** The largest spreadsheet the server keeps, as saved JSON. */
export const MAX_WORKBOOK_BYTES = 5_000_000;

export interface StoredWorkbook {
  id: string;
  title: string;
  workbook: WorkbookJSON;
  activeSheet?: string;
  version: number;
  updatedAt: number;
}

export interface WorkbookSummary {
  id: string;
  title: string;
  version: number;
  updatedAt: number;
}

interface Row {
  id: string;
  owner_id: string;
  title: string;
  data: string;
  version: number;
  updated_at: number;
}

/** What's stored in the data column: the document and the tab that was open. */
interface Data {
  workbook: WorkbookJSON;
  activeSheet?: string;
}

const toStored = (row: Row): StoredWorkbook => {
  const data = JSON.parse(row.data) as Data;
  return { id: row.id, title: row.title, workbook: data.workbook, activeSheet: data.activeSheet, version: row.version, updatedAt: row.updated_at };
};

export function listWorkbooks(ownerId: string): WorkbookSummary[] {
  const rows = db()
    .prepare("select id, title, version, updated_at from workbooks where owner_id = ? order by updated_at desc")
    .all(ownerId) as unknown as Pick<Row, "id" | "title" | "version" | "updated_at">[];
  return rows.map((r) => ({ id: r.id, title: r.title, version: r.version, updatedAt: r.updated_at }));
}

type Lookup = { found: StoredWorkbook } | { missing: true } | { forbidden: true };

/** Another account's spreadsheet reads as forbidden, never as its contents. */
export function getWorkbook(ownerId: string, id: string): Lookup {
  const row = db().prepare("select * from workbooks where id = ?").get(id) as Row | undefined;
  if (!row) return { missing: true };
  if (row.owner_id !== ownerId) return { forbidden: true };
  return { found: toStored(row) };
}

export type SaveResult =
  | { saved: StoredWorkbook }
  | { conflict: StoredWorkbook }
  | { forbidden: true }
  | { tooLarge: true };

/**
 * Creates or updates a spreadsheet. An update must say which version it
 * was based on; if someone saved in between, nothing is written and the
 * newer copy comes back as a conflict.
 */
export function saveWorkbook(
  ownerId: string,
  id: string,
  title: string,
  data: Data,
  baseVersion: number | null,
): SaveResult {
  const json = JSON.stringify(data);
  if (json.length > MAX_WORKBOOK_BYTES) return { tooLarge: true };
  const now = Date.now();
  const existing = db().prepare("select * from workbooks where id = ?").get(id) as Row | undefined;
  if (existing && existing.owner_id !== ownerId) return { forbidden: true };

  if (!existing) {
    db()
      .prepare("insert into workbooks (id, owner_id, title, data, version, created_at, updated_at) values (?, ?, ?, ?, 1, ?, ?)")
      .run(id, ownerId, title, json, now, now);
    return { saved: { id, title, ...data, version: 1, updatedAt: now } };
  }
  if (baseVersion !== existing.version) return { conflict: toStored(existing) };

  const version = existing.version + 1;
  // The version check in the WHERE clause makes this safe if two saves race.
  const result = db()
    .prepare("update workbooks set title = ?, data = ?, version = ?, updated_at = ? where id = ? and version = ?")
    .run(title, json, version, now, id, existing.version);
  if (Number(result.changes) === 0) {
    const latest = db().prepare("select * from workbooks where id = ?").get(id) as unknown as Row;
    return { conflict: toStored(latest) };
  }
  return { saved: { id, title, ...data, version, updatedAt: now } };
}

export function deleteWorkbook(ownerId: string, id: string): boolean {
  return Number(db().prepare("delete from workbooks where id = ? and owner_id = ?").run(id, ownerId).changes) > 0;
}
