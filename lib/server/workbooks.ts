// Server-only: spreadsheets saved to accounts, and who else may open them.
import type { WorkbookJSON } from "@/lib/engine";
import { db } from "./db";

/** The largest spreadsheet the server keeps, as saved JSON. */
export const MAX_WORKBOOK_BYTES = 5_000_000;

/** What someone may do with a spreadsheet: everything, edit it, or only look. */
export type Role = "owner" | "edit" | "view";
export type LinkAccess = "none" | "view";

export interface StoredWorkbook {
  id: string;
  title: string;
  workbook: WorkbookJSON;
  activeSheet?: string;
  version: number;
  updatedAt: number;
  role: Role;
  ownerName: string;
}

export interface WorkbookSummary {
  id: string;
  title: string;
  version: number;
  updatedAt: number;
  role: Role;
  ownerName: string;
}

export interface Share {
  userId: string;
  email: string;
  name: string;
  role: "view" | "edit";
}

interface Row {
  id: string;
  owner_id: string;
  title: string;
  data: string;
  version: number;
  updated_at: number;
  link_access: LinkAccess;
  owner_name: string;
}

/** What's stored in the data column: the document and the tab that was open. */
interface Data {
  workbook: WorkbookJSON;
  activeSheet?: string;
}

const SELECT = "select w.*, u.name as owner_name from workbooks w join users u on u.id = w.owner_id";

const rowById = (id: string) => db().prepare(`${SELECT} where w.id = ?`).get(id) as unknown as Row | undefined;

/** The best role a person has on a spreadsheet; userId null means signed out. */
function roleOf(row: Row, userId: string | null): Role | null {
  if (userId && row.owner_id === userId) return "owner";
  if (userId) {
    const share = db().prepare("select role from workbook_shares where workbook_id = ? and user_id = ?").get(row.id, userId) as
      | { role: "view" | "edit" }
      | undefined;
    if (share) return share.role;
  }
  return row.link_access === "view" ? "view" : null;
}

const toStored = (row: Row, role: Role): StoredWorkbook => {
  const data = JSON.parse(row.data) as Data;
  return {
    id: row.id,
    title: row.title,
    workbook: data.workbook,
    activeSheet: data.activeSheet,
    version: row.version,
    updatedAt: row.updated_at,
    role,
    ownerName: row.owner_name,
  };
};

/** A person's own spreadsheets and the ones shared with them, newest first. */
export function listWorkbooks(userId: string): WorkbookSummary[] {
  const rows = db()
    .prepare(
      `select w.id, w.title, w.version, w.updated_at, u.name as owner_name,
              case when w.owner_id = ? then 'owner' else s.role end as role
         from workbooks w
         join users u on u.id = w.owner_id
         left join workbook_shares s on s.workbook_id = w.id and s.user_id = ?
        where w.owner_id = ? or s.user_id is not null
        order by w.updated_at desc`,
    )
    .all(userId, userId, userId) as unknown as (Pick<Row, "id" | "title" | "version" | "updated_at" | "owner_name"> & { role: Role })[];
  return rows.map((r) => ({ id: r.id, title: r.title, version: r.version, updatedAt: r.updated_at, role: r.role, ownerName: r.owner_name }));
}

type Lookup = { found: StoredWorkbook } | { missing: true } | { forbidden: true };

/** Anyone without access gets "forbidden", never the contents. */
export function getWorkbook(userId: string | null, id: string): Lookup {
  const row = rowById(id);
  if (!row) return { missing: true };
  const role = roleOf(row, userId);
  if (!role) return { forbidden: true };
  return { found: toStored(row, role) };
}

/** Just the version, for cheap checks on whether someone else has saved. */
export function workbookVersion(userId: string | null, id: string): number | null {
  const row = rowById(id);
  return row && roleOf(row, userId) ? row.version : null;
}

export type SaveResult =
  | { saved: { version: number; updatedAt: number } }
  | { conflict: StoredWorkbook }
  | { forbidden: true }
  | { tooLarge: true };

/**
 * Creates or updates a spreadsheet. An update must say which version it was
 * based on; if someone saved in between, nothing is written and the newer
 * copy comes back as a conflict. Owners and editors may save.
 */
export function saveWorkbook(userId: string, id: string, title: string, data: Data, baseVersion: number | null): SaveResult {
  const json = JSON.stringify(data);
  if (json.length > MAX_WORKBOOK_BYTES) return { tooLarge: true };
  const now = Date.now();
  const existing = rowById(id);

  if (!existing) {
    db()
      .prepare("insert into workbooks (id, owner_id, title, data, version, created_at, updated_at) values (?, ?, ?, ?, 1, ?, ?)")
      .run(id, userId, title, json, now, now);
    return { saved: { version: 1, updatedAt: now } };
  }
  const role = roleOf(existing, userId);
  if (role !== "owner" && role !== "edit") return { forbidden: true };
  if (baseVersion !== existing.version) return { conflict: toStored(existing, role) };

  const version = existing.version + 1;
  // The version check in the WHERE clause makes this safe if two saves race.
  const result = db()
    .prepare("update workbooks set title = ?, data = ?, version = ?, updated_at = ? where id = ? and version = ?")
    .run(title, json, version, now, id, existing.version);
  if (Number(result.changes) === 0) return { conflict: toStored(rowById(id)!, role) };
  return { saved: { version, updatedAt: now } };
}

/** Only the owner can delete. */
export function deleteWorkbook(userId: string, id: string): boolean {
  return Number(db().prepare("delete from workbooks where id = ? and owner_id = ?").run(id, userId).changes) > 0;
}

// ---- Sharing (owner only) ------------------------------------------------------

const isOwner = (userId: string, id: string) => rowById(id)?.owner_id === userId;

export function sharing(userId: string, id: string): { shares: Share[]; linkAccess: LinkAccess } | null {
  const row = rowById(id);
  if (!row || row.owner_id !== userId) return null;
  const shares = db()
    .prepare(
      "select s.user_id, u.email, u.name, s.role from workbook_shares s join users u on u.id = s.user_id where s.workbook_id = ? order by s.created_at",
    )
    .all(id) as unknown as { user_id: string; email: string; name: string; role: "view" | "edit" }[];
  return { shares: shares.map((s) => ({ userId: s.user_id, email: s.email, name: s.name, role: s.role })), linkAccess: row.link_access };
}

export type ShareResult = { ok: true } | { error: string; status: number };

export function shareWith(ownerId: string, id: string, email: string, role: "view" | "edit"): ShareResult {
  if (!isOwner(ownerId, id)) return { error: "Only the owner can share this spreadsheet.", status: 403 };
  const user = db().prepare("select id from users where email = ?").get(email.trim().toLowerCase()) as { id: string } | undefined;
  if (!user) return { error: "There's no Kexcel account with that email. Ask them to sign up first.", status: 404 };
  if (user.id === ownerId) return { error: "That's your own account.", status: 400 };
  db()
    .prepare(
      "insert into workbook_shares (workbook_id, user_id, role, created_at) values (?, ?, ?, ?) on conflict (workbook_id, user_id) do update set role = excluded.role",
    )
    .run(id, user.id, role, Date.now());
  return { ok: true };
}

export function unshare(ownerId: string, id: string, userId: string): ShareResult {
  if (!isOwner(ownerId, id)) return { error: "Only the owner can change sharing.", status: 403 };
  db().prepare("delete from workbook_shares where workbook_id = ? and user_id = ?").run(id, userId);
  return { ok: true };
}

export function setLinkAccess(ownerId: string, id: string, access: LinkAccess): ShareResult {
  if (!isOwner(ownerId, id)) return { error: "Only the owner can change sharing.", status: 403 };
  db().prepare("update workbooks set link_access = ? where id = ?").run(access, id);
  return { ok: true };
}
