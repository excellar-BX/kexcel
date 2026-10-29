import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { WorkbookJSON } from "./engine";

// ---- openFromCloud, with the network and browser storage stubbed --------------

const fetchCloud = vi.fn();
vi.mock("./cloud", async (orig) => ({ ...(await orig<typeof import("./cloud")>()), fetchCloud: (id: string) => fetchCloud(id) }));

const memory = new Map<string, string>();
vi.stubGlobal("localStorage", {
  getItem: (k: string) => memory.get(k) ?? null,
  setItem: (k: string, v: string) => void memory.set(k, v),
  removeItem: (k: string) => void memory.delete(k),
  key: (i: number) => [...memory.keys()][i] ?? null,
  get length() {
    return memory.size;
  },
});

const book = (a1: string): WorkbookJSON => ({ sheets: [{ id: "s1", name: "Sheet1", cells: { A1: a1 } }] });
const server = (version: number, a1: string) => ({ id: "doc1", title: "Server", workbook: book(a1), version, updatedAt: 1 });

describe("opening a spreadsheet while signed in", async () => {
  const { openFromCloud } = await import("./sync");
  const { loadSheet, saveSheet } = await import("./storage");
  const { CloudError } = await import("./cloud");

  beforeEach(() => {
    memory.clear();
    fetchCloud.mockReset();
  });

  it("uses this device's copy when the account doesn't have it yet", async () => {
    saveSheet("doc1", { title: "Mine", workbook: book("1") });
    fetchCloud.mockResolvedValue(null);
    expect(await openFromCloud("doc1", "u1")).toEqual({ kind: "ready", role: null });
    expect(loadSheet("doc1")?.title).toBe("Mine");
  });

  it("takes the account's newer copy when nothing changed here", async () => {
    saveSheet("doc1", { title: "Old", workbook: book("1"), cloud: { owner: "u1", version: 2, synced: true } });
    fetchCloud.mockResolvedValue(server(3, "new"));
    expect(await openFromCloud("doc1", "u1")).toEqual({ kind: "ready" });
    expect(loadSheet("doc1")?.workbook.sheets[0].cells.A1).toBe("new");
    expect(loadSheet("doc1")?.cloud).toEqual({ owner: "u1", version: 3, synced: true });
  });

  it("keeps unsent changes when the account hasn't moved on", async () => {
    saveSheet("doc1", { title: "Edited", workbook: book("mine"), cloud: { owner: "u1", version: 3, synced: false } });
    fetchCloud.mockResolvedValue(server(3, "theirs"));
    expect(await openFromCloud("doc1", "u1")).toEqual({ kind: "ready" });
    expect(loadSheet("doc1")?.workbook.sheets[0].cells.A1).toBe("mine");
  });

  it("asks which to keep when both changed", async () => {
    saveSheet("doc1", { title: "Edited", workbook: book("mine"), cloud: { owner: "u1", version: 2, synced: false } });
    fetchCloud.mockResolvedValue(server(4, "theirs"));
    const result = await openFromCloud("doc1", "u1");
    expect(result.kind).toBe("conflict");
    expect(loadSheet("doc1")?.workbook.sheets[0].cells.A1).toBe("mine");
  });

  it("downloads to a device that has never seen it, and refuses other accounts' spreadsheets", async () => {
    fetchCloud.mockResolvedValue(server(5, "fresh"));
    expect(await openFromCloud("doc1", "u1")).toEqual({ kind: "ready" });
    expect(loadSheet("doc1")?.workbook.sheets[0].cells.A1).toBe("fresh");
    fetchCloud.mockRejectedValue(new CloudError("nope", 403));
    expect(await openFromCloud("doc2", "u1")).toEqual({ kind: "forbidden" });
  });

  it("opens this device's copy when offline", async () => {
    saveSheet("doc1", { title: "Offline", workbook: book("1") });
    fetchCloud.mockRejectedValue(new CloudError("offline", 0, true));
    expect(await openFromCloud("doc1", "u1")).toEqual({ kind: "ready", role: null });
  });
});

// ---- The server's store, on a throwaway database ------------------------------

describe("saving to an account", async () => {
  process.env.KEXCEL_DB_PATH = join(mkdtempSync(join(tmpdir(), "kexcel-")), "test.db");
  const { db } = await import("./server/db");
  const store = await import("./server/workbooks");
  for (const id of ["alice", "bob"]) {
    db().prepare("insert into users (id, email, name, password_hash, created_at) values (?, ?, ?, ?, ?)").run(id, `${id}@x.io`, id, "x", 0);
  }
  const data = (a1: string) => ({ workbook: book(a1) });

  it("creates, updates and refuses stale saves", () => {
    expect("saved" in store.saveWorkbook("alice", "wb1", "Budget", data("1"), null)).toBe(true);
    const second = store.saveWorkbook("alice", "wb1", "Budget", data("2"), 1);
    expect("saved" in second && second.saved.version).toBe(2);
    const stale = store.saveWorkbook("alice", "wb1", "Budget", data("old"), 1);
    expect("conflict" in stale && stale.conflict.workbook.sheets[0].cells.A1).toBe("2");
    expect(store.listWorkbooks("alice").map((w) => [w.id, w.version])).toEqual([["wb1", 2]]);
  });

  it("keeps each account's spreadsheets to itself", () => {
    expect(store.getWorkbook("bob", "wb1")).toEqual({ forbidden: true });
    expect(store.saveWorkbook("bob", "wb1", "Mine now", data("x"), 2)).toEqual({ forbidden: true });
    expect(store.deleteWorkbook("bob", "wb1")).toBe(false);
    expect(store.listWorkbooks("bob")).toEqual([]);
    expect(store.deleteWorkbook("alice", "wb1")).toBe(true);
    expect(store.getWorkbook("alice", "wb1")).toEqual({ missing: true });
  });

  it("lets owners share for viewing or editing, and by link", () => {
    store.saveWorkbook("alice", "team", "Team budget", data("1"), null);
    db().prepare("insert into users (id, email, name, password_hash, created_at) values (?, ?, ?, ?, ?)").run("cara", "cara@x.io", "Cara", "x", 0);

    expect(store.shareWith("bob", "team", "cara@x.io", "edit")).toMatchObject({ status: 403 }); // not the owner
    expect(store.shareWith("alice", "team", "nobody@x.io", "view")).toMatchObject({ status: 404 });
    expect(store.shareWith("alice", "team", "bob@x.io", "view")).toEqual({ ok: true });
    expect(store.shareWith("alice", "team", "CARA@x.io", "edit")).toEqual({ ok: true });

    // A viewer can read but not save; an editor can save but not delete or reshare.
    expect(store.getWorkbook("bob", "team")).toMatchObject({ found: { role: "view", ownerName: "alice" } });
    expect(store.saveWorkbook("bob", "team", "x", data("bob"), 1)).toEqual({ forbidden: true });
    expect(store.saveWorkbook("cara", "team", "Team budget", data("cara"), 1)).toMatchObject({ saved: { version: 2 } });
    expect(store.deleteWorkbook("cara", "team")).toBe(false);
    expect(store.shareWith("cara", "team", "bob@x.io", "edit")).toMatchObject({ status: 403 });
    expect(store.listWorkbooks("cara").map((w) => [w.id, w.role])).toEqual([["team", "edit"]]);

    // The link opens it for anyone, read-only, only once the owner turns that on.
    expect(store.getWorkbook(null, "team")).toEqual({ forbidden: true });
    store.setLinkAccess("alice", "team", "view");
    expect(store.getWorkbook(null, "team")).toMatchObject({ found: { role: "view" } });
    expect(store.workbookVersion(null, "team")).toBe(2);

    // Removing someone takes their access away; the owner sees the current list.
    store.unshare("alice", "team", "bob");
    store.setLinkAccess("alice", "team", "none");
    expect(store.getWorkbook("bob", "team")).toEqual({ forbidden: true });
    expect(store.sharing("alice", "team")).toMatchObject({ linkAccess: "none", shares: [{ email: "cara@x.io", role: "edit" }] });
    expect(store.sharing("cara", "team")).toBeNull();
  });

  it("refuses spreadsheets that are too large", () => {
    const huge = { workbook: book("x".repeat(store.MAX_WORKBOOK_BYTES)) };
    expect(store.saveWorkbook("alice", "wb2", "Huge", huge, null)).toEqual({ tooLarge: true });
  });
});
