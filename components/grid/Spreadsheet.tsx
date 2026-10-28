"use client";

import Link from "next/link";
import { useParams } from "next/navigation";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  addressToName,
  colToLetters,
  displayValue,
  isError,
  parseKey,
  Sheet,
  shiftFormula,
  type CellAddress,
  type CellUpdate,
} from "@/lib/engine";
import { loadSheet, saveSheet } from "@/lib/storage";
import { parseTSV, toTSV } from "@/lib/tsv";
import {
  COL_COUNT,
  COL_HEADER_HEIGHT,
  DEFAULT_COL_WIDTH,
  MIN_COL_WIDTH,
  OVERSCAN,
  ROW_COUNT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants";

interface Selection {
  /** The active cell, where the selection started. */
  anchor: CellAddress;
  /** The opposite corner, moved by shift+arrows or dragging. */
  focus: CellAddress;
}

interface Editing {
  row: number;
  col: number;
  value: string;
  /** "enter": started by typing, arrows commit. "edit": arrows move the caret. */
  mode: "enter" | "edit";
  source: "cell" | "bar";
}

interface Batch {
  before: CellUpdate[];
  after: CellUpdate[];
}

interface Clip {
  text: string;
  origin: CellAddress;
  raws: string[][];
  cut: boolean;
}

interface Range {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

const MAX_HISTORY = 200;

const clamp = (n: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, n));

const clampAddr = ({ row, col }: CellAddress): CellAddress => ({
  row: clamp(row, 0, ROW_COUNT - 1),
  col: clamp(col, 0, COL_COUNT - 1),
});

function rangeOf({ anchor, focus }: Selection): Range {
  return {
    r0: Math.min(anchor.row, focus.row),
    r1: Math.max(anchor.row, focus.row),
    c0: Math.min(anchor.col, focus.col),
    c1: Math.max(anchor.col, focus.col),
  };
}

const inRange = (r: Range, row: number, col: number) =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

const normalizeClipText = (s: string) => s.replace(/\r\n?/g, "\n").replace(/\n$/, "");

const ARROWS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

function initialWidths(saved: Record<string, number> | undefined): number[] {
  const widths = Array<number>(COL_COUNT).fill(DEFAULT_COL_WIDTH);
  for (const [col, w] of Object.entries(saved ?? {})) {
    const c = Number(col);
    if (c >= 0 && c < COL_COUNT && w >= MIN_COL_WIDTH) widths[c] = w;
  }
  return widths;
}

export default function Spreadsheet() {
  const { id: sheetId } = useParams<{ id: string }>();
  // Rendered client-only, so reading localStorage while initialising state is safe.
  const [saved] = useState(() => loadSheet(sheetId));
  // The sheet is mutated in place; `version` tells React when to re-render.
  const [sheet] = useState(() => (saved ? Sheet.fromJSON(saved.cells) : new Sheet()));
  const [version, setVersion] = useState(0);
  const [title, setTitle] = useState(saved?.title ?? "Untitled sheet");
  const [colWidths, setColWidths] = useState(() => initialWidths(saved?.colWidths));
  const [sel, setSel] = useState<Selection>({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
  const [editing, setEditing] = useState<Editing | null>(null);
  const [viewport, setViewport] = useState({ top: 0, left: 0, width: 1200, height: 800 });
  const [saveStatus, setSaveStatus] = useState<"saved" | "failed" | null>(saved ? "saved" : null);

  const dirtyRef = useRef(false);
  const historyRef = useRef<{ undo: Batch[]; redo: Batch[] }>({ undo: [], redo: [] });
  const clipRef = useRef<Clip | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const cellsRef = useRef<HTMLDivElement>(null);
  const scrollFrameRef = useRef(0);

  const colLefts = useMemo(() => {
    const lefts = [0];
    for (const w of colWidths) lefts.push(lefts[lefts.length - 1] + w);
    return lefts;
  }, [colWidths]);
  const totalWidth = colLefts[COL_COUNT];
  const totalHeight = ROW_COUNT * ROW_HEIGHT;

  /** The column under an x offset within the cell area. */
  const colAt = (x: number) => {
    let lo = 0;
    let hi = COL_COUNT - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (colLefts[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };

  const range = rangeOf(sel);
  const active = sel.anchor;

  // ---- Persistence -------------------------------------------------------

  useEffect(() => {
    if (!dirtyRef.current) return;
    const timer = setTimeout(() => {
      const widths: Record<string, number> = {};
      colWidths.forEach((w, c) => {
        if (w !== DEFAULT_COL_WIDTH) widths[c] = w;
      });
      const ok = saveSheet(sheetId, { title, cells: sheet.toJSON(), colWidths: widths });
      setSaveStatus(ok ? "saved" : "failed");
    }, 400);
    return () => clearTimeout(timer);
  }, [version, title, colWidths, sheet, sheetId]);

  // ---- Viewport tracking -------------------------------------------------

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const observer = new ResizeObserver(() =>
      setViewport({ top: el.scrollTop, left: el.scrollLeft, width: el.clientWidth, height: el.clientHeight }),
    );
    observer.observe(el);
    el.focus({ preventScroll: true });
    return () => observer.disconnect();
  }, []);

  const onScroll = () => {
    cancelAnimationFrame(scrollFrameRef.current);
    scrollFrameRef.current = requestAnimationFrame(() => {
      const el = scrollRef.current;
      if (el) setViewport({ top: el.scrollTop, left: el.scrollLeft, width: el.clientWidth, height: el.clientHeight });
    });
  };

  /** Scrolls just enough to bring a cell fully into view. */
  const ensureVisible = ({ row, col }: CellAddress) => {
    const el = scrollRef.current;
    if (!el) return;
    const viewH = el.clientHeight - COL_HEADER_HEIGHT;
    const viewW = el.clientWidth - ROW_HEADER_WIDTH;
    const top = row * ROW_HEIGHT;
    const left = colLefts[col];
    const right = colLefts[col + 1];
    if (top < el.scrollTop) el.scrollTop = top;
    else if (top + ROW_HEIGHT > el.scrollTop + viewH) el.scrollTop = top + ROW_HEIGHT - viewH;
    if (left < el.scrollLeft) el.scrollLeft = left;
    else if (right > el.scrollLeft + viewW) el.scrollLeft = Math.min(left, right - viewW);
  };

  const focusGrid = () => scrollRef.current?.focus({ preventScroll: true });

  // ---- Edits and history -------------------------------------------------

  const applyEdits = (updates: CellUpdate[], record = true) => {
    const before: CellUpdate[] = [];
    const after: CellUpdate[] = [];
    for (const u of updates) {
      const old = sheet.getRaw(u.row, u.col);
      if (old === u.raw) continue;
      before.push({ row: u.row, col: u.col, raw: old });
      after.push(u);
    }
    if (after.length === 0) return;
    sheet.setCells(after);
    if (record) {
      const h = historyRef.current;
      h.undo.push({ before, after });
      if (h.undo.length > MAX_HISTORY) h.undo.shift();
      h.redo = [];
    }
    dirtyRef.current = true;
    setVersion((v) => v + 1);
  };

  const undo = () => {
    const batch = historyRef.current.undo.pop();
    if (!batch) return;
    historyRef.current.redo.push(batch);
    applyEdits(batch.before, false);
    selectCells(batch.before);
  };

  const redo = () => {
    const batch = historyRef.current.redo.pop();
    if (!batch) return;
    historyRef.current.undo.push(batch);
    applyEdits(batch.after, false);
    selectCells(batch.after);
  };

  /** Selects the bounding box of a set of cells, e.g. after undo. */
  const selectCells = (cells: CellUpdate[]) => {
    if (cells.length === 0) return;
    const rows = cells.map((c) => c.row);
    const cols = cells.map((c) => c.col);
    const anchor = { row: Math.min(...rows), col: Math.min(...cols) };
    setSel({ anchor, focus: { row: Math.max(...rows), col: Math.max(...cols) } });
    ensureVisible(anchor);
  };

  /** Non-empty cells inside a range. */
  const filledIn = (r: Range): CellAddress[] => {
    const out: CellAddress[] = [];
    for (const key of sheet.keys()) {
      const a = parseKey(key);
      if (inRange(r, a.row, a.col)) out.push(a);
    }
    return out;
  };

  const clearRange = (r: Range) => applyEdits(filledIn(r).map((a) => ({ ...a, raw: "" })));

  // ---- Cell editing ------------------------------------------------------

  const startEditing = (mode: Editing["mode"], value?: string, source: Editing["source"] = "cell") => {
    setEditing({
      row: active.row,
      col: active.col,
      value: value ?? sheet.getRaw(active.row, active.col),
      mode,
      source,
    });
  };

  const commitEdit = (move?: [number, number]) => {
    if (!editing) return;
    applyEdits([{ row: editing.row, col: editing.col, raw: editing.value }]);
    setEditing(null);
    if (move) {
      const next = clampAddr({ row: editing.row + move[0], col: editing.col + move[1] });
      setSel({ anchor: next, focus: next });
      ensureVisible(next);
    }
    focusGrid();
  };

  const cancelEdit = () => {
    setEditing(null);
    focusGrid();
  };

  const onEditorKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!editing) return;
    if (e.key === "Enter") {
      e.preventDefault();
      commitEdit([e.shiftKey ? -1 : 1, 0]);
    } else if (e.key === "Tab") {
      e.preventDefault();
      commitEdit([0, e.shiftKey ? -1 : 1]);
    } else if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
    } else if (editing.mode === "enter" && e.key in ARROWS) {
      e.preventDefault();
      commitEdit(ARROWS[e.key]);
    }
  };

  const placeCaretAtEnd = (e: React.FocusEvent<HTMLInputElement>) => {
    const len = e.target.value.length;
    e.target.setSelectionRange(len, len);
  };

  // ---- Keyboard navigation -----------------------------------------------

  /** Ctrl+arrow: jump to the edge of the current block of data, like Excel. */
  const jump = (from: CellAddress, dr: number, dc: number): CellAddress => {
    const inBounds = (r: number, c: number) => r >= 0 && r < ROW_COUNT && c >= 0 && c < COL_COUNT;
    const filled = (r: number, c: number) => sheet.getRaw(r, c) !== "";
    let r = from.row;
    let c = from.col;
    if (!inBounds(r + dr, c + dc)) return from;
    if (filled(r, c) && filled(r + dr, c + dc)) {
      while (inBounds(r + dr, c + dc) && filled(r + dr, c + dc)) {
        r += dr;
        c += dc;
      }
      return { row: r, col: c };
    }
    r += dr;
    c += dc;
    while (!filled(r, c) && inBounds(r + dr, c + dc)) {
      r += dr;
      c += dc;
    }
    return { row: r, col: c };
  };

  const onGridKeyDown = (e: React.KeyboardEvent) => {
    if (editing) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key;

    if (ctrl && key.toLowerCase() === "z") {
      e.preventDefault();
      if (e.shiftKey) redo();
      else undo();
      return;
    }
    if (ctrl && key.toLowerCase() === "y") {
      e.preventDefault();
      redo();
      return;
    }
    if (ctrl && key.toLowerCase() === "a") {
      e.preventDefault();
      setSel({ anchor: { row: 0, col: 0 }, focus: { row: ROW_COUNT - 1, col: COL_COUNT - 1 } });
      return;
    }

    const move = (target: CellAddress, extend: boolean) => {
      const next = clampAddr(target);
      setSel(extend ? { anchor: sel.anchor, focus: next } : { anchor: next, focus: next });
      ensureVisible(next);
    };

    if (key in ARROWS) {
      e.preventDefault();
      const [dr, dc] = ARROWS[key];
      const from = e.shiftKey ? sel.focus : sel.anchor;
      move(ctrl ? jump(from, dr, dc) : { row: from.row + dr, col: from.col + dc }, e.shiftKey);
      return;
    }

    const pageRows = Math.max(1, Math.floor((viewport.height - COL_HEADER_HEIGHT) / ROW_HEIGHT) - 1);
    switch (key) {
      case "Tab":
        e.preventDefault();
        move({ row: active.row, col: active.col + (e.shiftKey ? -1 : 1) }, false);
        return;
      case "Home":
        e.preventDefault();
        move(ctrl ? { row: 0, col: 0 } : { row: active.row, col: 0 }, e.shiftKey);
        return;
      case "PageDown":
      case "PageUp": {
        e.preventDefault();
        const dir = key === "PageDown" ? 1 : -1;
        move({ row: active.row + dir * pageRows, col: active.col }, e.shiftKey);
        return;
      }
      case "Enter":
      case "F2":
        e.preventDefault();
        startEditing("edit");
        return;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        clearRange(range);
        return;
    }

    if (key.length === 1 && !ctrl && !e.altKey) {
      // Typing replaces the cell's content, like Excel.
      e.preventDefault();
      startEditing("enter", key);
    }
  };

  // ---- Clipboard -----------------------------------------------------------

  // Re-registered every render so the handlers always see the current selection.
  useEffect(() => {
    const ours = () => document.activeElement === scrollRef.current;

    const copy = (e: ClipboardEvent, cut: boolean) => {
      if (!ours() || !e.clipboardData) return;
      const used = sheet.usedBounds();
      // Ignore the empty part of huge selections such as select-all.
      const r: Range = {
        ...range,
        r1: Math.max(range.r0, Math.min(range.r1, (used?.rows ?? 1) - 1)),
        c1: Math.max(range.c0, Math.min(range.c1, (used?.cols ?? 1) - 1)),
      };
      const raws: string[][] = [];
      const shown: string[][] = [];
      for (let row = r.r0; row <= r.r1; row++) {
        const rawRow: string[] = [];
        const shownRow: string[] = [];
        for (let col = r.c0; col <= r.c1; col++) {
          rawRow.push(sheet.getRaw(row, col));
          shownRow.push(displayValue(sheet.getValue(row, col)));
        }
        raws.push(rawRow);
        shown.push(shownRow);
      }
      const text = toTSV(shown);
      e.clipboardData.setData("text/plain", text);
      e.preventDefault();
      clipRef.current = { text, origin: { row: r.r0, col: r.c0 }, raws, cut };
      if (cut) clearRange(r);
    };

    const paste = (e: ClipboardEvent) => {
      if (!ours() || !e.clipboardData) return;
      e.preventDefault();
      const text = e.clipboardData.getData("text/plain");
      if (!text) return;
      const clip = clipRef.current;
      // Our own copy carries formulas; anything else is pasted as plain text.
      const internal = clip && normalizeClipText(clip.text) === normalizeClipText(text) ? clip : null;
      const raws = internal ? internal.raws : parseTSV(text);
      if (raws.length === 0) return;

      const single = raws.length === 1 && raws[0].length === 1;
      const fill = single && (range.r1 > range.r0 || range.c1 > range.c0);
      const rows = fill ? range.r1 - range.r0 + 1 : raws.length;
      const cols = fill ? range.c1 - range.c0 + 1 : Math.max(...raws.map((r) => r.length));

      const updates: CellUpdate[] = [];
      for (let i = 0; i < rows && range.r0 + i < ROW_COUNT; i++) {
        for (let j = 0; j < cols && range.c0 + j < COL_COUNT; j++) {
          let raw = fill ? raws[0][0] : (raws[i][j] ?? "");
          if (internal && !internal.cut) {
            const srcRow = internal.origin.row + (fill ? 0 : i);
            const srcCol = internal.origin.col + (fill ? 0 : j);
            raw = shiftFormula(raw, range.r0 + i - srcRow, range.c0 + j - srcCol);
          }
          updates.push({ row: range.r0 + i, col: range.c0 + j, raw });
        }
      }
      applyEdits(updates);
      // A cut moves its cells once; pasting again would duplicate them.
      if (internal?.cut) clipRef.current = null;
      setSel({
        anchor: { row: range.r0, col: range.c0 },
        focus: clampAddr({ row: range.r0 + rows - 1, col: range.c0 + cols - 1 }),
      });
    };

    const onCopy = (e: ClipboardEvent) => copy(e, false);
    const onCut = (e: ClipboardEvent) => copy(e, true);
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCut);
    document.addEventListener("paste", paste);
    return () => {
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("cut", onCut);
      document.removeEventListener("paste", paste);
    };
  });

  // ---- Mouse ---------------------------------------------------------------

  const pointToCell = (clientX: number, clientY: number): CellAddress => {
    const rect = cellsRef.current!.getBoundingClientRect();
    return clampAddr({ row: Math.floor((clientY - rect.top) / ROW_HEIGHT), col: colAt(clientX - rect.left) });
  };

  /** Tracks the mouse until the button is released. */
  const trackDrag = (onMove: (e: MouseEvent) => void) => {
    const up = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", up);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", up);
  };

  const onCellsMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    if (editing) commitEdit();
    const cell = pointToCell(e.clientX, e.clientY);
    const anchor = e.shiftKey ? sel.anchor : cell;
    setSel({ anchor, focus: cell });
    trackDrag((m) => setSel({ anchor, focus: pointToCell(m.clientX, m.clientY) }));
  };

  const onColumnHeaderMouseDown = (e: React.MouseEvent, col: number) => {
    if (e.button !== 0) return;
    if (editing) commitEdit();
    const startCol = e.shiftKey ? sel.anchor.col : col;
    const select = (c: number) =>
      setSel({ anchor: { row: 0, col: startCol }, focus: { row: ROW_COUNT - 1, col: c } });
    select(col);
    trackDrag((m) => select(pointToCell(m.clientX, m.clientY).col));
  };

  const onRowHeaderMouseDown = (e: React.MouseEvent, row: number) => {
    if (e.button !== 0) return;
    if (editing) commitEdit();
    const startRow = e.shiftKey ? sel.anchor.row : row;
    const select = (r: number) =>
      setSel({ anchor: { row: startRow, col: 0 }, focus: { row: r, col: COL_COUNT - 1 } });
    select(row);
    trackDrag((m) => select(pointToCell(m.clientX, m.clientY).row));
  };

  const onResizeMouseDown = (e: React.MouseEvent, col: number) => {
    e.stopPropagation();
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = colWidths[col];
    trackDrag((m) => {
      const width = Math.max(MIN_COL_WIDTH, Math.round(startWidth + m.clientX - startX));
      dirtyRef.current = true;
      setColWidths((ws) => ws.map((w, i) => (i === col ? width : w)));
    });
  };

  // ---- Derived view data ---------------------------------------------------

  const rowStart = Math.max(0, Math.floor(viewport.top / ROW_HEIGHT) - OVERSCAN);
  const rowEnd = Math.min(ROW_COUNT - 1, Math.ceil((viewport.top + viewport.height) / ROW_HEIGHT) + OVERSCAN);
  const colStart = Math.max(0, colAt(viewport.left) - 1);
  const colEnd = Math.min(COL_COUNT - 1, colAt(viewport.left + viewport.width) + 1);

  const stats = useMemo(() => {
    void version; // recompute after every edit
    let count = 0;
    let numbers = 0;
    let sum = 0;
    for (const a of filledIn(range)) {
      count++;
      const v = sheet.getValue(a.row, a.col);
      if (typeof v === "number") {
        numbers++;
        sum += v;
      }
    }
    return { count, numbers, sum };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [version, range.r0, range.r1, range.c0, range.c1]);

  const multiCell = range.r1 > range.r0 || range.c1 > range.c0;
  const selectionName = multiCell
    ? `${addressToName({ row: range.r0, col: range.c0 })}:${addressToName({ row: range.r1, col: range.c1 })}`
    : addressToName(active);

  const cellNodes: React.ReactNode[] = [];
  for (let row = rowStart; row <= rowEnd; row++) {
    for (let col = colStart; col <= colEnd; col++) {
      const value = sheet.getValue(row, col);
      if (value === null) continue;
      const text = displayValue(value);
      const error = isError(value);
      const align = typeof value === "number" ? "right" : typeof value === "string" ? "left" : "center";
      let width = colWidths[col];
      // Text spills into empty neighbours on the right, as in Excel.
      if (align === "left") {
        for (let c = col + 1; c <= colEnd && sheet.getRaw(row, c) === ""; c++) width += colWidths[c];
      }
      cellNodes.push(
        <div
          key={`${row},${col}`}
          className={`absolute overflow-hidden whitespace-nowrap px-1.5 leading-6 ${error ? "text-[var(--error)]" : ""}`}
          style={{ left: colLefts[col], top: row * ROW_HEIGHT, width, height: ROW_HEIGHT, textAlign: align }}
          title={error ? value.message : undefined}
        >
          {text}
        </div>,
      );
    }
  }

  const colLines: React.ReactNode[] = [];
  for (let col = colStart; col <= colEnd; col++) {
    colLines.push(
      <div
        key={col}
        className="absolute top-0 h-full w-px bg-[var(--grid-line)]"
        style={{ left: colLefts[col + 1] - 1 }}
      />,
    );
  }

  const selLeft = colLefts[range.c0];
  const selTop = range.r0 * ROW_HEIGHT;
  const selWidth = colLefts[range.c1 + 1] - selLeft;
  const selHeight = (range.r1 - range.r0 + 1) * ROW_HEIGHT;

  const barValue = editing ? editing.value : sheet.getRaw(active.row, active.col);

  return (
    <div className="flex h-dvh flex-col bg-[var(--bg)] text-[var(--fg)]">
      {/* Top bar */}
      <header className="flex items-center gap-3 border-b border-[var(--border)] px-3 py-2">
        <Link href="/" className="flex items-center gap-2 font-semibold text-[var(--accent)]">
          <span className="grid h-7 w-7 place-items-center rounded-md bg-[var(--accent)] text-sm text-white">K</span>
        </Link>
        <input
          value={title}
          onChange={(e) => {
            dirtyRef.current = true;
            setTitle(e.target.value);
          }}
          onKeyDown={(e) => e.key === "Enter" && focusGrid()}
          aria-label="Sheet title"
          className="min-w-0 flex-1 rounded px-2 py-1 text-base font-medium outline-none hover:bg-[var(--hover)] focus:bg-[var(--hover)] sm:max-w-sm"
        />
        <span className="ml-auto text-xs text-[var(--muted)]">
          {saveStatus === "failed" ? "Couldn't save in this browser" : saveStatus === "saved" ? "Saved on this device" : ""}
        </span>
      </header>

      {/* Formula bar */}
      <div className="flex items-center gap-2 border-b border-[var(--border)] px-3 py-1 text-sm">
        <div className="w-24 shrink-0 truncate rounded border border-[var(--border)] px-2 py-0.5 text-center font-mono text-xs">
          {selectionName}
        </div>
        <span className="font-serif text-sm italic text-[var(--muted)]">fx</span>
        <input
          value={barValue}
          aria-label="Formula bar"
          onFocus={(e) => {
            if (editing) setEditing({ ...editing, mode: "edit", source: "bar" });
            else startEditing("edit", undefined, "bar");
            placeCaretAtEnd(e);
          }}
          onChange={(e) =>
            setEditing((ed) =>
              ed
                ? { ...ed, value: e.target.value }
                : { row: active.row, col: active.col, value: e.target.value, mode: "edit", source: "bar" },
            )
          }
          onKeyDown={onEditorKeyDown}
          className="min-w-0 flex-1 rounded px-2 py-0.5 font-mono text-sm outline-none focus:bg-[var(--hover)]"
          spellCheck={false}
        />
      </div>

      {/* Grid */}
      <div
        ref={scrollRef}
        tabIndex={0}
        onScroll={onScroll}
        onKeyDown={onGridKeyDown}
        className="relative min-h-0 flex-1 select-none overflow-auto text-[13px] outline-none"
        role="grid"
        aria-label="Spreadsheet"
        aria-rowcount={ROW_COUNT}
        aria-colcount={COL_COUNT}
      >
        <div style={{ width: ROW_HEADER_WIDTH + totalWidth, height: COL_HEADER_HEIGHT + totalHeight }}>
          {/* Column headers */}
          <div
            className="sticky top-0 z-20 flex border-b border-[var(--border)] bg-[var(--header-bg)]"
            style={{ height: COL_HEADER_HEIGHT, width: ROW_HEADER_WIDTH + totalWidth }}
          >
            <div
              className="sticky left-0 z-30 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)] hover:bg-[var(--hover)]"
              style={{ width: ROW_HEADER_WIDTH }}
              onMouseDown={() =>
                setSel({ anchor: { row: 0, col: 0 }, focus: { row: ROW_COUNT - 1, col: COL_COUNT - 1 } })
              }
              title="Select all"
            />
            <div className="relative" style={{ width: totalWidth }}>
              {Array.from({ length: colEnd - colStart + 1 }, (_, i) => {
                const col = colStart + i;
                const selected = col >= range.c0 && col <= range.c1;
                return (
                  <div
                    key={col}
                    onMouseDown={(e) => onColumnHeaderMouseDown(e, col)}
                    className={`absolute top-0 grid h-full place-items-center border-r border-[var(--grid-line)] text-xs ${
                      selected ? "bg-[var(--header-selected)] font-semibold text-[var(--accent)]" : "text-[var(--muted)]"
                    }`}
                    style={{ left: colLefts[col], width: colWidths[col] }}
                  >
                    {colToLetters(col)}
                    <div
                      onMouseDown={(e) => onResizeMouseDown(e, col)}
                      className="absolute -right-1 top-0 z-10 h-full w-2 cursor-col-resize hover:bg-[var(--accent)]/40"
                    />
                  </div>
                );
              })}
            </div>
          </div>

          <div className="flex">
            {/* Row headers */}
            <div
              className="sticky left-0 z-10 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)]"
              style={{ width: ROW_HEADER_WIDTH, height: totalHeight }}
            >
              {Array.from({ length: rowEnd - rowStart + 1 }, (_, i) => {
                const row = rowStart + i;
                const selected = row >= range.r0 && row <= range.r1;
                return (
                  <div
                    key={row}
                    onMouseDown={(e) => onRowHeaderMouseDown(e, row)}
                    className={`absolute left-0 grid w-full place-items-center border-b border-[var(--grid-line)] text-xs ${
                      selected ? "bg-[var(--header-selected)] font-semibold text-[var(--accent)]" : "text-[var(--muted)]"
                    }`}
                    style={{ top: row * ROW_HEIGHT, height: ROW_HEIGHT }}
                  >
                    {row + 1}
                  </div>
                );
              })}
            </div>

            {/* Cells */}
            <div
              ref={cellsRef}
              className="relative"
              style={{
                width: totalWidth,
                height: totalHeight,
                backgroundImage: `linear-gradient(to bottom, transparent ${ROW_HEIGHT - 1}px, var(--grid-line) ${ROW_HEIGHT - 1}px)`,
                backgroundSize: `100% ${ROW_HEIGHT}px`,
              }}
              onMouseDown={onCellsMouseDown}
              onDoubleClick={() => startEditing("edit")}
            >
              {colLines}
              {cellNodes}

              {/* Selection */}
              <div
                className="pointer-events-none absolute border border-[var(--accent)] bg-[var(--selection)]"
                style={{ left: selLeft - 1, top: selTop - 1, width: selWidth + 1, height: selHeight + 1 }}
              />
              <div
                className="pointer-events-none absolute border-2 border-[var(--accent)]"
                style={{
                  left: colLefts[active.col] - 1,
                  top: active.row * ROW_HEIGHT - 1,
                  width: colWidths[active.col] + 1,
                  height: ROW_HEIGHT + 1,
                }}
              />

              {editing && (
                <input
                  value={editing.value}
                  autoFocus={editing.source === "cell"}
                  onFocus={(e) => {
                    if (editing.source !== "cell") setEditing({ ...editing, source: "cell", mode: "edit" });
                    placeCaretAtEnd(e);
                  }}
                  onChange={(e) => setEditing({ ...editing, value: e.target.value })}
                  onKeyDown={onEditorKeyDown}
                  onMouseDown={(e) => e.stopPropagation()}
                  onDoubleClick={(e) => e.stopPropagation()}
                  spellCheck={false}
                  className="absolute z-10 border-2 border-[var(--accent)] bg-[var(--bg)] px-1 font-[inherit] text-[13px] shadow-lg outline-none"
                  style={{
                    left: colLefts[editing.col] - 1,
                    top: editing.row * ROW_HEIGHT - 1,
                    height: ROW_HEIGHT + 1,
                    width: Math.min(600, Math.max(colWidths[editing.col] + 1, editing.value.length * 8 + 24)),
                  }}
                />
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Status bar */}
      <footer className="flex h-7 items-center justify-end gap-4 border-t border-[var(--border)] px-3 text-xs text-[var(--muted)]">
        {multiCell && stats.count > 0 && (
          <>
            {stats.numbers > 0 && <span>Sum: {displayValue(stats.sum)}</span>}
            {stats.numbers > 0 && <span>Average: {displayValue(stats.sum / stats.numbers)}</span>}
            <span>Count: {stats.count}</span>
          </>
        )}
      </footer>
    </div>
  );
}
