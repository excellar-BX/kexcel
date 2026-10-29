// Types for the parts of Node's built-in SQLite module Kexcel uses. The installed
// @types/node predates node:sqlite; remove this file after upgrading it to v22.5+.
declare module "node:sqlite" {
  type SQLValue = null | number | bigint | string | Uint8Array;

  export class StatementSync {
    run(...params: SQLValue[]): { changes: number | bigint; lastInsertRowid: number | bigint };
    get(...params: SQLValue[]): Record<string, SQLValue> | undefined;
    all(...params: SQLValue[]): Record<string, SQLValue>[];
  }

  export class DatabaseSync {
    constructor(path: string, options?: { open?: boolean; readOnly?: boolean });
    exec(sql: string): void;
    prepare(sql: string): StatementSync;
    close(): void;
  }
}
