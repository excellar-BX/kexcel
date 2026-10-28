"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useRef, useState } from "react";
import { COL_COUNT, ROW_COUNT } from "@/components/grid/constants";
import { OPEN_FILE_ACCEPT, openFileAsSpreadsheet } from "@/lib/open-file";
import { deleteSheet, listSheets, newSheetId } from "@/lib/storage";

const formatWhen = (ms: number) =>
  ms ? new Date(ms).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" }) : "";

export default function SheetList() {
  const router = useRouter();
  // Rendered client-only, so localStorage can be read during initialisation.
  const [sheets, setSheets] = useState(listSheets);
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

      {sheets.length > 0 && (
        <div className="mt-10">
          <h2 className="mb-3 text-sm font-medium text-[var(--muted)]">Your sheets</h2>
          <ul className="divide-y divide-[var(--border)] rounded-lg border border-[var(--border)]">
            {sheets.map((s) => (
              <li key={s.id} className="flex items-center gap-3 px-4 py-3">
                <Link href={`/sheet/${s.id}`} className="min-w-0 flex-1 truncate font-medium hover:underline">
                  {s.title || "Untitled sheet"}
                </Link>
                <span className="hidden shrink-0 text-xs text-[var(--muted)] sm:inline">{formatWhen(s.updatedAt)}</span>
                <button
                  onClick={() => {
                    if (!confirm(`Delete "${s.title}"? This can't be undone.`)) return;
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
