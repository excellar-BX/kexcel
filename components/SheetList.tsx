"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useEffect, useRef, useState } from "react";
import { useSession } from "@/components/auth/useSession";
import { deleteCloud, listCloud, pushCloud, type CloudSummary } from "@/lib/cloud";
import { COL_COUNT, ROW_COUNT } from "@/components/grid/constants";
import { OPEN_FILE_ACCEPT, openFileAsSpreadsheet } from "@/lib/open-file";
import { deleteSheet, listSheets, loadSheet, newSheetId, saveSheet } from "@/lib/storage";

const formatWhen = (ms: number) =>
  ms ? new Date(ms).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" }) : "";

export default function SheetList() {
  const router = useRouter();
  // Rendered client-only, so localStorage can be read during initialisation.
  const [sheets, setSheets] = useState(listSheets);
  const session = useSession();
  const account = session.status === "ready" ? session.account : null;
  // The account's spreadsheets; null while loading or when signed out.
  const [cloud, setCloud] = useState<{ owner: string; items: CloudSummary[] } | null>(null);
  const [cloudError, setCloudError] = useState<string | null>(null);

  useEffect(() => {
    if (!account) return;
    let live = true;
    listCloud()
      .then((items) => live && setCloud({ owner: account.id, items }))
      .catch(() => live && setCloudError("Couldn't load your account's spreadsheets. Showing the ones on this device."));
    return () => {
      live = false;
    };
  }, [account]);

  const cloudItems = account && cloud?.owner === account.id ? cloud.items : null;
  const inCloud = new Set(cloudItems?.map((c) => c.id));
  /** Everything to list: the account's spreadsheets, then any only on this device. */
  const rows = [
    ...(cloudItems ?? []).map((c) => ({ id: c.id, title: c.title, updatedAt: c.updatedAt, where: "account" as const })),
    ...sheets.filter((s) => !inCloud.has(s.id)).map((s) => ({ ...s, where: "device" as const })),
  ];

  const upload = async (id: string) => {
    const doc = loadSheet(id);
    if (!doc || !account) return;
    try {
      const res = await pushCloud(id, { title: doc.title, workbook: doc.workbook, activeSheet: doc.activeSheet }, null);
      if ("version" in res) {
        saveSheet(id, { ...doc, cloud: { owner: account.id, version: res.version, synced: true } });
        setCloud((c) => c && { ...c, items: [{ id, title: doc.title, version: res.version, updatedAt: Date.now() }, ...c.items] });
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Couldn't save it to your account.");
    }
  };
  const [error, setError] = useState<string | null>(null);
  const [opening, setOpening] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const open = async (file: File) => {
    setOpening(true);
    setError(null);
    try {
      router.push(`/sheet/${await openFileAsSpreadsheet(file, { maxRows: ROW_COUNT, maxCols: COL_COUNT })}`);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Couldn't open that file.");
      setOpening(false);
    }
  };

  return (
    <div className="w-full">
      <div className="flex flex-col gap-3 sm:flex-row">
        <button
          onClick={() => router.push(`/sheet/${newSheetId()}`)}
          className="rounded-lg bg-[var(--accent)] px-5 py-3 font-medium text-white shadow-sm transition hover:opacity-90"
        >
          New spreadsheet
        </button>
        <button
          onClick={() => fileRef.current?.click()}
          disabled={opening}
          className="rounded-lg border border-[var(--border)] px-5 py-3 font-medium transition hover:bg-[var(--hover)] disabled:opacity-50"
        >
          {opening ? "Opening…" : "Open Excel or CSV file"}
        </button>
        <input
          ref={fileRef}
          type="file"
          accept={OPEN_FILE_ACCEPT}
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            e.target.value = "";
            if (file) void open(file);
          }}
        />
      </div>
      {error && (
        <p role="alert" className="mt-3 text-sm text-[var(--error)]">
          {error}
        </p>
      )}

      {!account && session.status === "ready" && (
        <p className="mt-6 text-sm text-[var(--muted)]">
          <Link href="/login?mode=signup" className="font-medium text-[var(--accent)] hover:underline">
            Create a free account
          </Link>{" "}
          to keep your spreadsheets on every device and use the AI assistant.
        </p>
      )}
      {cloudError && <p className="mt-3 text-sm text-[var(--muted)]">{cloudError}</p>}

      {rows.length > 0 && (
        <div className="mt-10">
          <h2 className="mb-3 text-sm font-medium text-[var(--muted)]">Your spreadsheets</h2>
          <ul className="divide-y divide-[var(--border)] rounded-lg border border-[var(--border)]">
            {rows.map((s) => (
              <li key={s.id} className="flex items-center gap-3 px-4 py-3">
                <Link href={`/sheet/${s.id}`} className="min-w-0 flex-1 truncate font-medium hover:underline">
                  {s.title || "Untitled sheet"}
                </Link>
                {account && s.where === "device" && (
                  <>
                    <span className="hidden shrink-0 text-xs text-[var(--muted)] sm:inline">Only on this device</span>
                    <button
                      onClick={() => void upload(s.id)}
                      className="shrink-0 rounded px-2 py-1 text-xs font-medium text-[var(--accent)] hover:bg-[var(--hover)]"
                    >
                      Save to account
                    </button>
                  </>
                )}
                <span className="hidden shrink-0 text-xs text-[var(--muted)] sm:inline">{formatWhen(s.updatedAt)}</span>
                <button
                  onClick={async () => {
                    const where = s.where === "account" ? " from your account" : "";
                    if (!confirm(`Delete "${s.title}"${where}? This can't be undone.`)) return;
                    if (s.where === "account") {
                      try {
                        await deleteCloud(s.id);
                      } catch (e) {
                        setError(e instanceof Error ? e.message : "Couldn't delete it.");
                        return;
                      }
                      setCloud((c) => c && { ...c, items: c.items.filter((x) => x.id !== s.id) });
                    }
                    deleteSheet(s.id);
                    setSheets(listSheets());
                  }}
                  className="shrink-0 rounded px-2 py-1 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
                  aria-label={`Delete ${s.title}`}
                >
                  Delete
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}
