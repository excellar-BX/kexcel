import { CloudError, fetchCloud, type CloudWorkbook } from "@/lib/cloud";
import { loadSheet, saveSheet } from "@/lib/storage";

export type OpenResult =
  | { kind: "ready" }
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
    return { kind: "ready" };
  }
  // Not in the account yet: it's new or only on this device, and the editor uploads it.
  if (!server) return { kind: "ready" };

  const linked = local?.cloud?.owner === owner ? local.cloud : undefined;
  const takeServer = () =>
    saveSheet(id, {
      title: server.title,
      workbook: server.workbook,
      activeSheet: server.activeSheet,
      cloud: { owner, version: server.version, synced: true },
    });

  if (!local || !linked) {
    takeServer();
    return { kind: "ready" };
  }
  if (server.version === linked.version) return { kind: "ready" };
  if (linked.synced) {
    takeServer();
    return { kind: "ready" };
  }
  return { kind: "conflict", server };
}
