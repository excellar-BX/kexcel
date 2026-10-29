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
`;

const globalForDb = globalThis as unknown as { kexcelDb?: DatabaseSync };

/** The shared connection, opened (and the tables created) on first use. */
export function db(): DatabaseSync {
  if (globalForDb.kexcelDb) return globalForDb.kexcelDb;
  // Scoped to ./data so the build doesn't trace the whole project as server files.
  const file = process.env.KEXCEL_DB_PATH
    ? resolve(/*turbopackIgnore: true*/ process.env.KEXCEL_DB_PATH)
    : join(process.cwd(), "data", "kexcel.db");
  mkdirSync(dirname(file), { recursive: true });
  const conn = new DatabaseSync(file);
  conn.exec("pragma journal_mode = wal; pragma foreign_keys = on;");
  conn.exec(SCHEMA);
  globalForDb.kexcelDb = conn;
  return conn;
}
