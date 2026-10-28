"use client";

import dynamic from "next/dynamic";

// These read localStorage while rendering, so they skip server rendering.

export const Spreadsheet = dynamic(() => import("./grid/Spreadsheet"), {
  ssr: false,
  loading: () => <div className="grid h-dvh place-items-center text-sm text-[var(--muted)]">Loading sheet…</div>,
});

export const SheetList = dynamic(() => import("./SheetList"), { ssr: false });
