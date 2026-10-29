import { CloudError, fetchCloud, type CloudWorkbook, type Role } from "@/lib/cloud";
import { loadSheet, saveSheet } from "@/lib/storage";

export type OpenResult =
  /** `role` is null when the spreadsheet isn't in any account yet (it's new or only on this device). */
  | { kind: "ready"; role: Role | null; ownerName?: string }
  /** Both this device and the account changed since they last matched. */
  | { kind: "conflict"; server: CloudWorkbook }
  | { kind: "forbidden" };

/**
 * Brings this device's copy of a spreadsheet up to date with the account
 * before it opens. The account's copy wins when only it has changed; when
 * both have, the editor asks which to keep. Offline, the device copy is used.
 */
export async function openFromCloud(id: string, owner: string): Promise<OpenResult> {
  const local = loadSheet(id);
  let server: CloudWorkbook | null;
  try {
    server = await fetchCloud(id);
  } catch (e) {
    if (e instanceof CloudError && e.status === 403) return { kind: "forbidden" };
    // Offline: work from this device, as whatever this copy was last opened as.
    return { kind: "ready", role: local?.cloud?.owner === owner ? (local.cloud.role ?? "owner") : null };
  }
  if (!server) return { kind: "ready", role: null };

  const ready = { kind: "ready" as const, role: server.role, ownerName: server.ownerName };
  const linked = local?.cloud?.owner === owner ? local.cloud : undefined;
  const takeServer = () =>
    saveSheet(id, {
      title: server.title,
      workbook: server.workbook,
      activeSheet: server.activeSheet,
      cloud: { owner, version: server.version, synced: true, role: server.role },
    });

  // A viewer's copy is never edited, so the account's copy always wins.
  if (!local || !linked || server.role === "view") {
    takeServer();
    return ready;
  }
  if (server.version === linked.version) return ready;
  if (linked.synced) {
    takeServer();
    return ready;
  }
  return { kind: "conflict", server };
}
