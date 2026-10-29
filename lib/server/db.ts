// Server-only: the account and spreadsheet database, a SQLite file on this server.
// Swap this module for Postgres (or similar) when deploying somewhere without a disk.
import { mkdirSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";

const SCHEMA = `
  create table if not exists users (
    id text primary key,
    email text not null unique,
    name text not null,
    password_hash text not null,
    created_at integer not null
  );
  create table if not exists sessions (
    token_hash text primary key,
    user_id text not null references users(id) on delete cascade,
    expires_at integer not null,
    created_at integer not null
  );
  create index if not exists sessions_user on sessions(user_id);
  create table if not exists workbooks (
    id text primary key,
    owner_id text not null references users(id) on delete cascade,
    title text not null,
    data text not null,
    version integer not null,
    created_at integer not null,
    updated_at integer not null
  );
  create index if not exists workbooks_owner on workbooks(owner_id, updated_at);
  create table if not exists workbook_shares (
    workbook_id text not null references workbooks(id) on delete cascade,
    user_id text not null references users(id) on delete cascade,
    role text not null check (role in ('view', 'edit')),
    created_at integer not null,
    primary key (workbook_id, user_id)
  );
  create index if not exists shares_user on workbook_shares(user_id);
`;

/** Changes to tables that already exist, applied once each. */
function migrate(conn: DatabaseSync) {
  const columns = (conn.prepare("pragma table_info(workbooks)").all() as { name: string }[]).map((c) => c.name);
  if (!columns.includes("link_access")) {
    conn.exec("alter table workbooks add column link_access text not null default 'none'");
  }
}

/** Bump when SCHEMA or migrate() change, so a connection kept across reloads upgrades too. */
const SCHEMA_VERSION = 2;

const globalForDb = globalThis as unknown as { kexcelDb?: DatabaseSync; kexcelSchema?: number };

function upgrade(conn: DatabaseSync) {
  conn.exec(SCHEMA);
  migrate(conn);
  globalForDb.kexcelSchema = SCHEMA_VERSION;
}

/** The shared connection, opened (and the tables created) on first use. */
export function db(): DatabaseSync {
  const cached = globalForDb.kexcelDb;
  if (cached) {
    // In development the connection outlives code reloads; bring its tables up to date.
    if (globalForDb.kexcelSchema !== SCHEMA_VERSION) upgrade(cached);
    return cached;
  }
  // Scoped to ./data so the build doesn't trace the whole project as server files.
  const file = process.env.KEXCEL_DB_PATH
    ? resolve(/*turbopackIgnore: true*/ process.env.KEXCEL_DB_PATH)
    : join(process.cwd(), "data", "kexcel.db");
  mkdirSync(dirname(file), { recursive: true });
  const conn = new DatabaseSync(file);
  conn.exec("pragma journal_mode = wal; pragma foreign_keys = on;");
  upgrade(conn);
  globalForDb.kexcelDb = conn;
  return conn;
}
