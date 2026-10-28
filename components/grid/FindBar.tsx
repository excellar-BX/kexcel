"use client";

import { useMemo, useState } from "react";
import { ChevronDown, ChevronUp, SearchIcon, XIcon } from "@/components/ui/icons";
import { formatValue, parseKey, type CellAddress, type Workbook } from "@/lib/engine";

/** Cells whose shown text or formula contains the query, in reading order. */
function findMatches(wb: Workbook, sheet: string, query: string): CellAddress[] {
  const q = query.toLowerCase();
  if (!q) return [];
  const out: CellAddress[] = [];
  for (const key of wb.cellKeys(sheet)) {
    const a = parseKey(key);
    const shown = formatValue(wb.getValue(sheet, a.row, a.col), wb.getFormat(sheet, a.row, a.col));
    if (shown.toLowerCase().includes(q) || wb.getRaw(sheet, a.row, a.col).toLowerCase().includes(q)) out.push(a);
  }
  return out.sort((a, b) => a.row - b.row || a.col - b.col);
}

export default function FindBar({
  wb,
  sheet,
  version,
  onGoTo,
  onClose,
}: {
  wb: Workbook;
  sheet: string;
  /** Changes whenever the workbook does, so matches stay current. */
  version: number;
  onGoTo: (cell: CellAddress) => void;
  onClose: () => void;
}) {
  const [query, setQuery] = useState("");
  const [index, setIndex] = useState(0);
  const matches = useMemo(() => {
    void version;
    return findMatches(wb, sheet, query);
  }, [wb, sheet, query, version]);

  const go = (dir: 1 | -1) => {
    if (matches.length === 0) return;
    const next = (index + dir + matches.length) % matches.length;
    setIndex(next);
    onGoTo(matches[next]);
  };

  return (
    <div className="absolute right-4 top-2 z-40 flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--panel)] px-3 py-2 shadow-lg">
      <SearchIcon size={15} className="text-[var(--muted)]" />
      <input
        autoFocus
        value={query}
        onChange={(e) => {
          setQuery(e.target.value);
          setIndex(-1);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            go(e.shiftKey ? -1 : 1);
          } else if (e.key === "Escape") {
            e.preventDefault();
            onClose();
          }
        }}
        placeholder="Find in sheet"
        aria-label="Find in sheet"
        className="w-44 bg-transparent text-sm outline-none"
      />
      <span className="min-w-14 text-right text-xs tabular-nums text-[var(--muted)]">
        {query ? `${matches.length ? Math.max(index, 0) + 1 : 0} of ${matches.length}` : ""}
      </span>
      <button type="button" onClick={() => go(-1)} aria-label="Previous match" className="grid h-6 w-6 place-items-center rounded hover:bg-[var(--hover)]">
        <ChevronUp size={15} />
      </button>
      <button type="button" onClick={() => go(1)} aria-label="Next match" className="grid h-6 w-6 place-items-center rounded hover:bg-[var(--hover)]">
        <ChevronDown size={15} />
      </button>
      <button type="button" onClick={onClose} aria-label="Close find" className="grid h-6 w-6 place-items-center rounded hover:bg-[var(--hover)]">
        <XIcon size={15} />
      </button>
    </div>
  );
}
