"use client";

import Link from "next/link";
import { useParams, useRouter } from "next/navigation";
import { useEffect, useMemo, useRef, useState } from "react";
import { flushSync } from "react-dom";
import AIPanel, { type AITab } from "@/components/ai/AIPanel";
import PrintDialog, { loadOptions, type PrintScope } from "@/components/print/PrintDialog";
import PrintView, { type PrintJob, type PrintOptions } from "@/components/print/PrintView";
import MenuBar from "@/components/toolbar/MenuBar";
import SheetTabs from "@/components/toolbar/SheetTabs";
import ShortcutsDialog from "@/components/toolbar/ShortcutsDialog";
import Toolbar from "@/components/toolbar/Toolbar";
import type { BorderMode } from "@/components/ui/icons";
import {
  addressToName,
  DEFAULT_FONT_SIZE,
  displayValue,
  formatValue,
  mergeFormat,
  nameToAddress,
  parseKey,
  shiftFormula,
  shownDecimals,
  Workbook,
  type CellAddress,
  type CellFormat,
  type CellPatch,
  type WorkbookJSON,
} from "@/lib/engine";
import { fillPatches, sortPatches, type Range } from "@/lib/sheet-ops";
import { loadSheet, newSheetId, saveSheet } from "@/lib/storage";
import { parseCSV, parseTSV, toCSV, toTSV } from "@/lib/tsv";
import { COL_COUNT, DEFAULT_COL_WIDTH, MIN_COL_WIDTH, ROW_COUNT } from "./constants";
import CellContextMenu from "./CellContextMenu";
import FindBar from "./FindBar";
import Grid, { inRange, rangeOf } from "./Grid";
import { GridLayout, ptToPx } from "./layout";
import type { Commands, Editing, EditorState, Selection, ToggleKey } from "./types";

type HistoryEntry =
  | { kind: "patch"; sheet: string; undo: CellPatch[]; redo: CellPatch[] }
  | { kind: "snapshot"; before: WorkbookJSON; after: WorkbookJSON; sheetBefore: string; sheetAfter: string };

interface Clip {
  text: string;
  sheet: string;
  origin: CellAddress;
  raws: string[][];
  formats: (CellFormat | undefined)[][];
  cut: boolean;
}

const MAX_HISTORY = 200;
const ORIGIN: Selection = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };

const clamp = (n: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, n));
const clampAddr = ({ row, col }: CellAddress): CellAddress => ({
  row: clamp(row, 0, ROW_COUNT - 1),
  col: clamp(col, 0, COL_COUNT - 1),
});
const normalizeClipText = (s: string) => s.replace(/\r\n?/g, "\n").replace(/\n$/, "");

const ARROWS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/** Remounts the editor when the URL switches to another document. */
export default function SpreadsheetRoute() {
  const { id } = useParams<{ id: string }>();
  return <Spreadsheet key={id} docId={id} />;
}

function Spreadsheet({ docId }: { docId: string }) {
  const router = useRouter();
  // Rendered client-only, so reading localStorage while initialising state is safe.
  const [saved] = useState(() => loadSheet(docId));
  // The workbook is mutated in place; `version` tells React when to re-render.
  const [wb] = useState(() => (saved ? Workbook.fromJSON(saved.workbook) : Workbook.withSheet()));
  const [version, setVersion] = useState(0);
  const [title, setTitle] = useState(saved?.title ?? "Untitled spreadsheet");
  const [activeSheet, setActiveSheet] = useState(() =>
    saved?.activeSheet && wb.hasSheet(saved.activeSheet) ? saved.activeSheet : wb.sheets[0].id,
  );
  const [sel, setSel] = useState<Selection>(ORIGIN);
  const [editing, setEditing] = useState<Editing | null>(null);
  const [zoom, setZoom] = useState(1);
  const [gridlines, setGridlines] = useState(true);
  const [currency, setCurrency] = useState("USD");
  const [historySize, setHistorySize] = useState({ undo: 0, redo: 0 });
  const [saveStatus, setSaveStatus] = useState<"saved" | "failed" | null>(saved ? "saved" : null);
  const [menu, setMenu] = useState<{ x: number; y: number; target: "cell" | "row" | "col" } | null>(null);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [nameBox, setNameBox] = useState<string | null>(null);
  const [printOpen, setPrintOpen] = useState(false);
  const [printJob, setPrintJob] = useState<PrintJob | null>(null);
  const [aiTab, setAiTab] = useState<AITab | null>(null);

  const dirtyRef = useRef(false);
  const historyRef = useRef<{ undo: HistoryEntry[]; redo: HistoryEntry[] }>({ undo: [], redo: [] });
  const clipRef = useRef<Clip | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const titleRef = useRef<HTMLInputElement>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const layout = useMemo(() => {
    void version; // sizes and fonts change with edits
    return GridLayout.forSheet(wb, activeSheet, zoom);
  }, [wb, activeSheet, zoom, version]);

  const range = rangeOf(sel);
  const active = sel.anchor;
  const activeFormat = wb.getFormat(activeSheet, active.row, active.col) ?? {};

  // ---- Persistence -------------------------------------------------------

  useEffect(() => {
    if (!dirtyRef.current) return;
    const timer = setTimeout(() => {
      const ok = saveSheet(docId, { title, workbook: wb.toJSON(), activeSheet });
      setSaveStatus(ok ? "saved" : "failed");
    }, 500);
    return () => clearTimeout(timer);
  }, [version, title, activeSheet, wb, docId]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    scrollRef.current?.focus({ preventScroll: true });
  }, []);

  // ---- Helpers -------------------------------------------------------------

  const focusGrid = () => scrollRef.current?.focus({ preventScroll: true });

  const changed = () => {
    dirtyRef.current = true;
    setVersion((v) => v + 1);
  };

  /** Scrolls just enough to bring a cell fully into view. */
  const ensureVisible = ({ row, col }: CellAddress) => {
    const el = scrollRef.current;
    if (!el) return;
    const viewH = el.clientHeight - layout.headerHeight;
    const viewW = el.clientWidth - layout.headerWidth;
    const top = layout.rowTop(row);
    const bottom = layout.rowTop(row + 1);
    const left = layout.colLeft(col);
    const right = layout.colLeft(col + 1);
    // Frozen rows and columns are always visible and cover the top and left of the view.
    const freeze = wb.getFreeze(activeSheet);
    const frozenH = layout.rowTop(Math.min(freeze.rows, ROW_COUNT - 1));
    const frozenW = layout.colLeft(Math.min(freeze.cols, COL_COUNT - 1));
    if (row >= freeze.rows) {
      if (top < el.scrollTop + frozenH) el.scrollTop = top - frozenH;
      else if (bottom > el.scrollTop + viewH) el.scrollTop = bottom - viewH;
    }
    if (col >= freeze.cols) {
      if (left < el.scrollLeft + frozenW) el.scrollLeft = left - frozenW;
      else if (right > el.scrollLeft + viewW) el.scrollLeft = Math.min(left - frozenW, right - viewW);
    }
  };

  const select = (next: Selection, scroll = true) => {
    setSel(next);
    if (scroll) ensureVisible(next.focus);
  };

  const selectRange = (r: Range) => select({ anchor: { row: r.r0, col: r.c0 }, focus: { row: r.r1, col: r.c1 } });

  const pushHistory = (entry: HistoryEntry) => {
    const h = historyRef.current;
    h.undo.push(entry);
    if (h.undo.length > MAX_HISTORY) h.undo.shift();
    h.redo = [];
    setHistorySize({ undo: h.undo.length, redo: 0 });
  };

  /** Applies cell changes as one undoable step. Returns false when nothing changed. */
  const commitPatches = (patches: CellPatch[]) => {
    const real = wb.effective(patches);
    if (real.length === 0) return false;
    const undo = wb.apply(real);
    pushHistory({ kind: "patch", sheet: activeSheet, undo, redo: real });
    changed();
    return true;
  };

  /** Runs a change to sheets or structure as one undoable step. `fn` may return the sheet to show next. */
  const structural = (fn: () => string | void) => {
    const before = wb.toJSON();
    let next: string | void;
    try {
      next = fn();
    } catch (e) {
      setToast(e instanceof Error ? e.message : String(e));
      return;
    }
    const sheetAfter = next && wb.hasSheet(next) ? next : wb.hasSheet(activeSheet) ? activeSheet : wb.sheets[0].id;
    pushHistory({ kind: "snapshot", before, after: wb.toJSON(), sheetBefore: activeSheet, sheetAfter });
    if (sheetAfter !== activeSheet) {
      setActiveSheet(sheetAfter);
      setSel(ORIGIN);
    }
    changed();
  };

  const boundsOf = (patches: CellPatch[]): Range | null => {
    if (patches.length === 0) return null;
    const rows = patches.map((p) => p.row);
    const cols = patches.map((p) => p.col);
    return { r0: Math.min(...rows), r1: Math.max(...rows), c0: Math.min(...cols), c1: Math.max(...cols) };
  };

  const replay = (direction: "undo" | "redo") => {
    const h = historyRef.current;
    const entry = (direction === "undo" ? h.undo : h.redo).pop();
    if (!entry) return;
    (direction === "undo" ? h.redo : h.undo).push(entry);
    setHistorySize({ undo: h.undo.length, redo: h.redo.length });
    setEditing(null);
    if (entry.kind === "patch") {
      const patches = direction === "undo" ? entry.undo : entry.redo;
      wb.apply(patches);
      if (wb.hasSheet(entry.sheet)) setActiveSheet(entry.sheet);
      const b = boundsOf(patches);
      if (b) selectRange(b);
    } else {
      wb.load(direction === "undo" ? entry.before : entry.after);
      const sheet = direction === "undo" ? entry.sheetBefore : entry.sheetAfter;
      setActiveSheet(wb.hasSheet(sheet) ? sheet : wb.sheets[0].id);
    }
    changed();
  };

  /** Filled cells inside a range. */
  const filledIn = (r: Range): CellAddress[] => {
    const out: CellAddress[] = [];
    for (const key of wb.cellKeys(activeSheet)) {
      const a = parseKey(key);
      if (inRange(r, a.row, a.col)) out.push(a);
    }
    return out;
  };

  /** Whole-row or whole-column selections are limited to just past the used area. */
  const formatTarget = (r: Range): Range => {
    const used = wb.usedBounds(activeSheet);
    const fullCols = r.r0 === 0 && r.r1 === ROW_COUNT - 1;
    const fullRows = r.c0 === 0 && r.c1 === COL_COUNT - 1;
    return {
      ...r,
      r1: fullCols ? Math.max(r.r0, Math.min(r.r1, (used?.rows ?? 0) + 100)) : r.r1,
      c1: fullRows ? Math.max(r.c0, Math.min(r.c1, (used?.cols ?? 0) + 10)) : r.c1,
    };
  };

  const updateFormats = (fn: (f: CellFormat | undefined, row: number, col: number, r: Range) => CellFormat | null) => {
    const r = formatTarget(range);
    const patches: CellPatch[] = [];
    for (let row = r.r0; row <= r.r1; row++) {
      for (let col = r.c0; col <= r.c1; col++) {
        patches.push({ sheet: activeSheet, row, col, format: fn(wb.getFormat(activeSheet, row, col), row, col, r) });
      }
    }
    commitPatches(patches);
  };

  const clearValues = (r: Range) =>
    commitPatches(filledIn(r).map((a) => ({ sheet: activeSheet, ...a, raw: "" })));

  // ---- Cell editing ------------------------------------------------------

  const startEditing = (mode: Editing["mode"], value?: string, source: Editing["source"] = "cell") => {
    setEditing({
      row: active.row,
      col: active.col,
      value: value ?? wb.getRaw(activeSheet, active.row, active.col),
      mode,
      source,
    });
  };

  const commitEdit = (move?: [number, number]) => {
    if (!editing) return;
    commitPatches([{ sheet: activeSheet, row: editing.row, col: editing.col, raw: editing.value }]);
    setEditing(null);
    if (move) {
      const next = clampAddr({ row: editing.row + move[0], col: editing.col + move[1] });
      select({ anchor: next, focus: next });
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

  // ---- Clipboard -----------------------------------------------------------

  const copyRange = (cut: boolean): string | null => {
    const used = wb.usedBounds(activeSheet);
    if (!used && !cut) return null;
    // Ignore the empty part of huge selections such as select-all.
    const r: Range = {
      ...range,
      r1: Math.max(range.r0, Math.min(range.r1, (used?.rows ?? 1) - 1)),
      c1: Math.max(range.c0, Math.min(range.c1, (used?.cols ?? 1) - 1)),
    };
    const raws: string[][] = [];
    const formats: (CellFormat | undefined)[][] = [];
    const shown: string[][] = [];
    for (let row = r.r0; row <= r.r1; row++) {
      raws.push([]);
      formats.push([]);
      shown.push([]);
      for (let col = r.c0; col <= r.c1; col++) {
        const format = wb.getFormat(activeSheet, row, col);
        raws[raws.length - 1].push(wb.getRaw(activeSheet, row, col));
        formats[formats.length - 1].push(format);
        shown[shown.length - 1].push(formatValue(wb.getValue(activeSheet, row, col), format));
      }
    }
    const text = toTSV(shown);
    clipRef.current = { text, sheet: activeSheet, origin: { row: r.r0, col: r.c0 }, raws, formats, cut };
    if (cut) {
      const patches: CellPatch[] = [];
      for (let row = r.r0; row <= r.r1; row++) {
        for (let col = r.c0; col <= r.c1; col++) patches.push({ sheet: activeSheet, row, col, raw: "", format: null });
      }
      commitPatches(patches);
    }
    return text;
  };

  const pasteText = (text: string) => {
    if (!text) return;
    const clip = clipRef.current;
    // Our own copy carries formulas and formats; anything else is pasted as plain text.
    const internal = clip && normalizeClipText(clip.text) === normalizeClipText(text) ? clip : null;
    const raws = internal ? internal.raws : parseTSV(text);
    if (raws.length === 0) return;

    const single = raws.length === 1 && raws[0].length === 1;
    const fill = single && (range.r1 > range.r0 || range.c1 > range.c0);
    const target = fill ? formatTarget(range) : range;
    const rows = fill ? target.r1 - target.r0 + 1 : raws.length;
    const cols = fill ? target.c1 - target.c0 + 1 : Math.max(...raws.map((r) => r.length));

    const patches: CellPatch[] = [];
    for (let i = 0; i < rows && range.r0 + i < ROW_COUNT; i++) {
      for (let j = 0; j < cols && range.c0 + j < COL_COUNT; j++) {
        const si = fill ? 0 : i;
        const sj = fill ? 0 : j;
        let raw = raws[si][sj] ?? "";
        const patch: CellPatch = { sheet: activeSheet, row: range.r0 + i, col: range.c0 + j, raw };
        if (internal) {
          if (!internal.cut) {
            raw = shiftFormula(raw, range.r0 + i - (internal.origin.row + si), range.c0 + j - (internal.origin.col + sj));
          }
          patch.raw = raw;
          patch.format = internal.formats[si]?.[sj] ?? null;
        }
        patches.push(patch);
      }
    }
    commitPatches(patches);
    // A cut moves its cells once; pasting again would duplicate them.
    if (internal?.cut) clipRef.current = null;
    select({
      anchor: { row: range.r0, col: range.c0 },
      focus: clampAddr({ row: range.r0 + rows - 1, col: range.c0 + cols - 1 }),
    });
  };

  // Re-registered every render so the handlers always see the current selection.
  useEffect(() => {
    const ours = () => document.activeElement === scrollRef.current;
    const onCopy = (e: ClipboardEvent, cut: boolean) => {
      if (!ours() || !e.clipboardData) return;
      const text = copyRange(cut);
      if (text === null) return;
      e.clipboardData.setData("text/plain", text);
      e.preventDefault();
    };
    const copy = (e: ClipboardEvent) => onCopy(e, false);
    const cut = (e: ClipboardEvent) => onCopy(e, true);
    const paste = (e: ClipboardEvent) => {
      if (!ours() || !e.clipboardData) return;
      e.preventDefault();
      pasteText(e.clipboardData.getData("text/plain"));
    };
    document.addEventListener("copy", copy);
    document.addEventListener("cut", cut);
    document.addEventListener("paste", paste);
    return () => {
      document.removeEventListener("copy", copy);
      document.removeEventListener("cut", cut);
      document.removeEventListener("paste", paste);
    };
  });

  // ---- Printing ----------------------------------------------------------------

  /** The cells to print for a scope: the used area, or the selection trimmed to it. */
  const printSheets = (scope: PrintScope): PrintJob["sheets"] => {
    const used = (id: string): Range | null => {
      const b = wb.usedBounds(id);
      return b ? { r0: 0, r1: b.rows - 1, c0: 0, c1: b.cols - 1 } : null;
    };
    if (scope === "all") {
      return wb.sheets.flatMap((s) => {
        const r = used(s.id);
        return r ? [{ id: s.id, range: r }] : [];
      });
    }
    const area = used(activeSheet);
    if (!area) return [];
    if (scope === "sheet") return [{ id: activeSheet, range: area }];
    const r = { ...range };
    if (r.r0 === 0 && r.r1 === ROW_COUNT - 1) r.r1 = Math.max(r.r0, area.r1);
    if (r.c0 === 0 && r.c1 === COL_COUNT - 1) r.c1 = Math.max(r.c0, area.c1);
    return [{ id: activeSheet, range: r }];
  };

  const startPrint = (scope: PrintScope, options: PrintOptions) => {
    setPrintOpen(false);
    setPrintJob({ sheets: printSheets(scope), options });
  };

  // Open the browser's print dialog once the print view is on the page.
  useEffect(() => {
    if (!printJob) return;
    const done = () => setPrintJob(null);
    window.addEventListener("afterprint", done, { once: true });
    window.print();
    return () => window.removeEventListener("afterprint", done);
  }, [printJob]);

  // Ctrl+P anywhere opens our dialog; printing from the browser menu still gets a clean sheet.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === "p") {
        e.preventDefault();
        if (editing) commitEdit();
        setPrintOpen(true);
      }
    };
    const onBeforePrint = () => {
      if (!printJob) flushSync(() => setPrintJob({ sheets: printSheets("sheet"), options: loadOptions(gridlines) }));
    };
    const onAfterPrint = () => setPrintJob(null);
    document.addEventListener("keydown", onKey, true);
    window.addEventListener("beforeprint", onBeforePrint);
    window.addEventListener("afterprint", onAfterPrint);
    return () => {
      document.removeEventListener("keydown", onKey, true);
      window.removeEventListener("beforeprint", onBeforePrint);
      window.removeEventListener("afterprint", onAfterPrint);
    };
  });

  // ---- Files -----------------------------------------------------------------

  const importFile = async (file: File) => {
    const rows = parseCSV(await file.text());
    const truncated = rows.length > ROW_COUNT || rows.some((r) => r.length > COL_COUNT);
    const updates: { row: number; col: number; raw: string }[] = [];
    rows.slice(0, ROW_COUNT).forEach((r, row) =>
      r.slice(0, COL_COUNT).forEach((raw, col) => raw !== "" && updates.push({ row, col, raw })),
    );
    structural(() => {
      const id = wb.addSheet(file.name.replace(/\.[^.]+$/, "").slice(0, 100), activeSheet);
      wb.setCells(id, updates);
      return id;
    });
    setToast(
      truncated
        ? `Imported the first ${ROW_COUNT.toLocaleString()} rows and ${COL_COUNT} columns of ${file.name}.`
        : `Imported ${file.name} into a new tab.`,
    );
  };

  const exportCSV = () => {
    const used = wb.usedBounds(activeSheet);
    const rows: string[][] = [];
    for (let row = 0; row < (used?.rows ?? 0); row++) {
      const line: string[] = [];
      for (let col = 0; col < (used?.cols ?? 0); col++) {
        line.push(formatValue(wb.getValue(activeSheet, row, col), wb.getFormat(activeSheet, row, col)));
      }
      rows.push(line);
    }
    const blob = new Blob([toCSV(rows)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title || "Spreadsheet"} - ${wb.sheetName(activeSheet)}.csv`.replace(/[\\/:*?"<>|]/g, "_");
    a.click();
    URL.revokeObjectURL(url);
  };

  /** Widest text in a column, measured with the cell fonts, in unscaled pixels. */
  const fitColumnWidth = (col: number) => {
    const ctx = document.createElement("canvas").getContext("2d");
    if (!ctx) return DEFAULT_COL_WIDTH;
    const family = getComputedStyle(document.body).fontFamily;
    let widest = 0;
    for (const key of wb.cellKeys(activeSheet)) {
      const a = parseKey(key);
      if (a.col !== col) continue;
      const f = wb.getFormat(activeSheet, a.row, a.col);
      ctx.font = `${f?.italic ? "italic " : ""}${f?.bold ? "700 " : ""}${ptToPx(f?.fontSize ?? DEFAULT_FONT_SIZE)}px ${family}`;
      widest = Math.max(widest, ctx.measureText(formatValue(wb.getValue(activeSheet, a.row, a.col), f)).width);
    }
    return widest ? Math.max(MIN_COL_WIDTH, Math.ceil(widest + 14)) : DEFAULT_COL_WIDTH;
  };

  // ---- Commands ----------------------------------------------------------------

  const fillFromEdge = (axis: "down" | "right") => {
    const r = range;
    const down = axis === "down";
    const single = down ? r.r0 === r.r1 : r.c0 === r.c1;
    if (single && (down ? r.r0 === 0 : r.c0 === 0)) return;
    const src: Range = single
      ? down
        ? { ...r, r0: r.r0 - 1, r1: r.r0 - 1 }
        : { ...r, c0: r.c0 - 1, c1: r.c0 - 1 }
      : down
        ? { ...r, r1: r.r0 }
        : { ...r, c1: r.c0 };
    const dst: Range = down ? { ...r, r0: src.r0 } : { ...r, c0: src.c0 };
    commitPatches(fillPatches(wb, activeSheet, src, dst, "copy"));
  };

  const setBorders = (mode: BorderMode) =>
    updateFormats((f, row, col, r) => {
      if (mode === "none") return mergeFormat(f, { borders: undefined });
      const b = { ...f?.borders };
      const top = row === r.r0;
      const bottom = row === r.r1;
      const left = col === r.c0;
      const right = col === r.c1;
      if (mode === "all") Object.assign(b, { top: true, bottom: true, left: true, right: true });
      if (mode === "outer") Object.assign(b, { top: b.top || top, bottom: b.bottom || bottom, left: b.left || left, right: b.right || right });
      if (mode === "inner") Object.assign(b, { top: b.top || !top, bottom: b.bottom || !bottom, left: b.left || !left, right: b.right || !right });
      if (mode === "top" && top) b.top = true;
      if (mode === "bottom" && bottom) b.bottom = true;
      if (mode === "left" && left) b.left = true;
      if (mode === "right" && right) b.right = true;
      return mergeFormat(f, { borders: b });
    });

  const insertFunction = (name: string) => {
    const multi = range.r1 > range.r0 || range.c1 > range.c0;
    if (!multi) {
      startEditing("enter", `=${name}(`);
      return;
    }
    const ref = (r0: number, c0: number, r1: number, c1: number) =>
      `${addressToName({ row: r0, col: c0 })}:${addressToName({ row: r1, col: c1 })}`;
    const patches: CellPatch[] = [];
    if (range.r1 > range.r0 && range.r1 + 1 < ROW_COUNT) {
      // A total under each selected column.
      for (let col = range.c0; col <= range.c1; col++) {
        patches.push({ sheet: activeSheet, row: range.r1 + 1, col, raw: `=${name}(${ref(range.r0, col, range.r1, col)})` });
      }
    } else if (range.c1 + 1 < COL_COUNT) {
      patches.push({ sheet: activeSheet, row: range.r0, col: range.c1 + 1, raw: `=${name}(${ref(range.r0, range.c0, range.r0, range.c1)})` });
    }
    if (commitPatches(patches)) selectRange(boundsOf(patches)!);
  };

  const sort = (ascending: boolean) => {
    const used = wb.usedBounds(activeSheet);
    if (!used) return;
    const multi = range.r1 > range.r0 || range.c1 > range.c0;
    // A single cell sorts every row of data by its column.
    const r = multi ? formatTarget(range) : { r0: 0, r1: used.rows - 1, c0: 0, c1: used.cols - 1 };
    const target = { ...r, r1: Math.min(r.r1, used.rows - 1) };
    if (target.r1 <= target.r0) return;
    commitPatches(sortPatches(wb, activeSheet, target, active.col, ascending));
  };

  // Each command refocuses the grid afterwards so the keyboard keeps working.
  const commands: Commands = {
    undo: () => {
      replay("undo");
      focusGrid();
    },
    redo: () => {
      replay("redo");
      focusGrid();
    },
    cut: () => {
      focusGrid();
      document.execCommand("cut");
    },
    copy: () => {
      focusGrid();
      document.execCommand("copy");
    },
    paste: () => {
      focusGrid();
      navigator.clipboard
        ?.readText()
        .then(pasteText)
        .catch(() => setToast("Your browser blocked clipboard access. Use Ctrl+V to paste."));
    },
    clearValues: () => {
      clearValues(range);
      focusGrid();
    },
    selectAll: () => {
      select({ anchor: { row: 0, col: 0 }, focus: { row: ROW_COUNT - 1, col: COL_COUNT - 1 } }, false);
      focusGrid();
    },
    find: () => setFindOpen(true),

    toggle: (key: ToggleKey) => {
      const on = !activeFormat[key];
      updateFormats((f) => mergeFormat(f, { [key]: on }));
      focusGrid();
    },
    applyFormat: (patch) => {
      updateFormats((f) => mergeFormat(f, patch));
      focusGrid();
    },
    clearFormatting: () => {
      updateFormats(() => null);
      focusGrid();
    },
    setNumberFormat: (numFmt, cur) => {
      if (cur) setCurrency(cur);
      updateFormats((f) =>
        mergeFormat(f, {
          numFmt: numFmt === "general" ? undefined : numFmt,
          currency: numFmt === "currency" ? (cur ?? currency) : undefined,
          decimals: undefined,
        }),
      );
      focusGrid();
    },
    changeDecimals: (delta) => {
      updateFormats((f, row, col) =>
        mergeFormat(f, { decimals: clamp(shownDecimals(wb.getValue(activeSheet, row, col), f) + delta, 0, 10) }),
      );
      focusGrid();
    },
    setFontSize: (size) => {
      updateFormats((f) => mergeFormat(f, { fontSize: size === DEFAULT_FONT_SIZE ? undefined : size }));
      focusGrid();
    },
    setBorders: (mode) => {
      setBorders(mode);
      focusGrid();
    },

    insertRows: (where) => {
      // A whole-column selection inserts one row, not one per row of the sheet.
      const count = range.r0 === 0 && range.r1 === ROW_COUNT - 1 ? 1 : range.r1 - range.r0 + 1;
      structural(() => wb.insert(activeSheet, "row", where === "above" ? range.r0 : range.r1 + 1, count));
      focusGrid();
    },
    insertCols: (where) => {
      const count = range.c0 === 0 && range.c1 === COL_COUNT - 1 ? 1 : range.c1 - range.c0 + 1;
      structural(() => wb.insert(activeSheet, "col", where === "left" ? range.c0 : range.c1 + 1, count));
      focusGrid();
    },
    deleteRows: () => {
      structural(() => wb.remove(activeSheet, "row", range.r0, range.r1 - range.r0 + 1));
      select({ anchor: { row: range.r0, col: active.col }, focus: { row: range.r0, col: active.col } });
      focusGrid();
    },
    deleteCols: () => {
      structural(() => wb.remove(activeSheet, "col", range.c0, range.c1 - range.c0 + 1));
      select({ anchor: { row: active.row, col: range.c0 }, focus: { row: active.row, col: range.c0 } });
      focusGrid();
    },
    fillDown: () => {
      fillFromEdge("down");
      focusGrid();
    },
    fillRight: () => {
      fillFromEdge("right");
      focusGrid();
    },
    sort: (asc) => {
      sort(asc);
      focusGrid();
    },
    insertFunction: (name) => {
      insertFunction(name);
      if (range.r1 > range.r0 || range.c1 > range.c0) focusGrid();
    },

    addSheet: () => {
      structural(() => wb.addSheet(undefined, activeSheet));
      focusGrid();
    },
    switchSheet: (id) => {
      if (editing) commitEdit();
      if (id !== activeSheet) {
        setActiveSheet(id);
        setSel(ORIGIN);
        scrollRef.current?.scrollTo(0, 0);
        dirtyRef.current = true;
      }
      focusGrid();
    },
    renameSheet: (id, name) => {
      const error = wb.validateSheetName(name, id);
      if (error) return error;
      if (name.trim() !== wb.sheetName(id)) structural(() => wb.renameSheet(id, name));
      focusGrid();
      return null;
    },
    duplicateSheet: (id) => {
      structural(() => wb.duplicateSheet(id));
      focusGrid();
    },
    deleteSheet: (id) => {
      if (!confirm(`Delete the sheet "${wb.sheetName(id)}"? You can undo this with Ctrl+Z.`)) return;
      structural(() => wb.removeSheet(id));
      focusGrid();
    },
    moveSheet: (id, delta) => {
      structural(() => {
        wb.moveSheet(id, wb.sheets.findIndex((s) => s.id === id) + delta);
        return id;
      });
      focusGrid();
    },

    newWorkbook: () => router.push(`/sheet/${newSheetId()}`),
    makeCopy: () => {
      const id = newSheetId();
      if (saveSheet(id, { title: `Copy of ${title}`, workbook: wb.toJSON(), activeSheet })) router.push(`/sheet/${id}`);
      else setToast("Couldn't save a copy in this browser.");
    },
    importCSV: () => fileRef.current?.click(),
    exportCSV,
    print: () => {
      if (editing) commitEdit();
      setPrintOpen(true);
    },
    renameWorkbook: () => {
      titleRef.current?.focus();
      titleRef.current?.select();
    },

    setZoom: (z) => {
      setZoom(z);
      focusGrid();
    },
    setFreeze: (axis, count) => {
      const current = wb.getFreeze(activeSheet);
      const next = axis === "rows" ? { ...current, rows: count } : { ...current, cols: count };
      structural(() => wb.setFreeze(activeSheet, next));
      focusGrid();
    },
    toggleGridlines: () => {
      setGridlines((g) => !g);
      focusGrid();
    },
    showShortcuts: () => setShortcutsOpen(true),
    openAI: (tab) => {
      if (editing) commitEdit();
      setAiTab(tab);
    },
  };

  /** Selects a reference written like "B4", "C2:C9" or "'My data'!A1", switching tabs if needed. */
  const goToReference = (ref: string) => {
    const bang = ref.lastIndexOf("!");
    const sheetPart = bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'") : null;
    const [a, b] = (bang >= 0 ? ref.slice(bang + 1) : ref).trim().toUpperCase().split(":");
    const start = nameToAddress(a ?? "");
    const end = b ? nameToAddress(b) : start;
    if (!start || !end) return;
    const target = sheetPart ? wb.sheets.find((s) => s.name.toLowerCase() === sheetPart.toLowerCase())?.id : activeSheet;
    if (!target) return;
    if (target !== activeSheet) setActiveSheet(target);
    select({ anchor: clampAddr(start), focus: clampAddr(end) });
  };

  // ---- Keyboard ---------------------------------------------------------------

  /** Ctrl+arrow: jump to the edge of the current block of data. */
  const jump = (from: CellAddress, dr: number, dc: number): CellAddress => {
    const inBounds = (r: number, c: number) => r >= 0 && r < ROW_COUNT && c >= 0 && c < COL_COUNT;
    const filled = (r: number, c: number) => wb.getRaw(activeSheet, r, c) !== "";
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
    const key = e.key.toLowerCase();

    const shortcut = (fn: () => void) => {
      e.preventDefault();
      fn();
    };

    if (ctrl && !e.altKey) {
      if (key === "z") return shortcut(e.shiftKey ? commands.redo : commands.undo);
      if (key === "y") return shortcut(commands.redo);
      if (key === "a") return shortcut(commands.selectAll);
      if (key === "b") return shortcut(() => commands.toggle("bold"));
      if (key === "i") return shortcut(() => commands.toggle("italic"));
      if (key === "u") return shortcut(() => commands.toggle("underline"));
      if (key === "d") return shortcut(commands.fillDown);
      if (key === "r") return shortcut(commands.fillRight);
      if (key === "f") return shortcut(commands.find);
      if (key === "\\") return shortcut(commands.clearFormatting);
      if (key === "/") return shortcut(commands.showShortcuts);
      if (key === " ") {
        return shortcut(() =>
          select({ anchor: { row: 0, col: range.c0 }, focus: { row: ROW_COUNT - 1, col: range.c1 } }, false),
        );
      }
    }
    if (e.altKey && e.shiftKey && (e.key === "5" || e.code === "Digit5")) return shortcut(() => commands.toggle("strike"));
    if (e.shiftKey && !ctrl && e.key === " ") {
      return shortcut(() =>
        select({ anchor: { row: range.r0, col: 0 }, focus: { row: range.r1, col: COL_COUNT - 1 } }, false),
      );
    }

    const move = (target: CellAddress, extend: boolean) => {
      const next = clampAddr(target);
      select(extend ? { anchor: sel.anchor, focus: next } : { anchor: next, focus: next });
    };

    if (e.key in ARROWS) {
      e.preventDefault();
      const [dr, dc] = ARROWS[e.key];
      const from = e.shiftKey ? sel.focus : sel.anchor;
      move(ctrl ? jump(from, dr, dc) : { row: from.row + dr, col: from.col + dc }, e.shiftKey);
      return;
    }

    const el = scrollRef.current;
    const pageRows = Math.max(1, Math.floor(((el?.clientHeight ?? 600) - layout.headerHeight) / (24 * zoom)) - 1);
    switch (e.key) {
      case "Tab":
        e.preventDefault();
        move({ row: active.row, col: active.col + (e.shiftKey ? -1 : 1) }, false);
        return;
      case "Home":
        e.preventDefault();
        move(ctrl ? { row: 0, col: 0 } : { row: active.row, col: 0 }, e.shiftKey);
        return;
      case "PageDown":
      case "PageUp":
        e.preventDefault();
        move({ row: active.row + (e.key === "PageDown" ? 1 : -1) * pageRows, col: active.col }, e.shiftKey);
        return;
      case "Enter":
      case "F2":
        e.preventDefault();
        startEditing("edit");
        return;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        clearValues(range);
        return;
    }

    if (e.key.length === 1 && !ctrl && !e.altKey) {
      // Typing replaces the cell's content.
      e.preventDefault();
      startEditing("enter", e.key);
    }
  };

  // ---- Derived view data ---------------------------------------------------------

  const stats = useMemo(() => {
    void version; // recompute after every edit
    let count = 0;
    let numbers = 0;
    let sum = 0;
    for (const key of wb.cellKeys(activeSheet)) {
      const a = parseKey(key);
      if (!inRange(range, a.row, a.col)) continue;
      count++;
      const v = wb.getValue(activeSheet, a.row, a.col);
      if (typeof v === "number") {
        numbers++;
        sum += v;
      }
    }
    return { count, numbers, sum };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wb, activeSheet, version, range.r0, range.r1, range.c0, range.c1]);

  const multiCell = range.r1 > range.r0 || range.c1 > range.c0;
  const selectionName = multiCell
    ? `${addressToName({ row: range.r0, col: range.c0 })}:${addressToName({ row: range.r1, col: range.c1 })}`
    : addressToName(active);
  const barValue = editing ? editing.value : wb.getRaw(activeSheet, active.row, active.col);

  const state: EditorState = {
    format: activeFormat,
    zoom,
    gridlines,
    canUndo: historySize.undo > 0,
    canRedo: historySize.redo > 0,
    sheets: wb.sheets,
    activeSheet,
    currency,
    freeze: wb.getFreeze(activeSheet),
    activeCell: active,
  };

  return (
    <>
      <div className="flex h-dvh flex-col bg-[var(--bg)] text-[var(--fg)] print:hidden">
      {/* Title and menus */}
      <header className="flex items-start gap-3 px-3 pt-2">
        <Link
          href="/"
          title="All spreadsheets"
          className="mt-1 grid h-9 w-9 shrink-0 place-items-center rounded-lg bg-[var(--accent)] font-bold text-white"
        >
          K
        </Link>
        <div className="min-w-0 flex-1">
          <input
            ref={titleRef}
            value={title}
            onChange={(e) => {
              dirtyRef.current = true;
              setTitle(e.target.value);
            }}
            onKeyDown={(e) => (e.key === "Enter" || e.key === "Escape") && focusGrid()}
            aria-label="Spreadsheet title"
            className="-ml-1 w-full max-w-md rounded px-1 text-lg outline-none hover:ring-1 hover:ring-[var(--border)] focus:ring-2 focus:ring-[var(--accent)]"
          />
          <div className="-ml-2 overflow-x-auto [scrollbar-width:none]">
            <MenuBar commands={commands} state={state} />
          </div>
        </div>
        <span className="mt-2 hidden shrink-0 text-xs text-[var(--muted)] sm:block">
          {saveStatus === "failed" ? "Couldn't save in this browser" : saveStatus === "saved" ? "Saved on this device" : ""}
        </span>
      </header>

      <div className="px-2 py-1.5">
        <Toolbar commands={commands} state={state} />
      </div>

      {/* Formula bar */}
      <div className="flex items-center gap-2 border-y border-[var(--border)] px-3 py-1 text-sm">
        <input
          value={nameBox ?? selectionName}
          onFocus={(e) => {
            setNameBox(selectionName);
            e.target.select();
          }}
          onChange={(e) => setNameBox(e.target.value)}
          onBlur={() => setNameBox(null)}
          onKeyDown={(e) => {
            if (e.key === "Escape") {
              setNameBox(null);
              focusGrid();
            }
            if (e.key !== "Enter") return;
            const [a, b] = (nameBox ?? "").trim().toUpperCase().split(":");
            const start = nameToAddress(a ?? "");
            const end = b ? nameToAddress(b) : start;
            if (start && end) select({ anchor: clampAddr(start), focus: clampAddr(end) });
            else setToast(`"${nameBox}" isn't a cell or range, like B3 or A1:C10.`);
            setNameBox(null);
            focusGrid();
          }}
          aria-label="Name box: type a cell to go to it"
          title="Type a cell or range to go to it"
          className="w-24 shrink-0 rounded border border-[var(--border)] bg-transparent px-2 py-0.5 text-center font-mono text-xs outline-none focus:border-[var(--accent)]"
        />
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
          className="min-w-0 flex-1 rounded bg-transparent px-2 py-0.5 font-mono text-sm outline-none focus:bg-[var(--hover)]"
          spellCheck={false}
        />
      </div>

      {/* Grid, with the AI panel beside it */}
      <div className="flex min-h-0 flex-1">
      <div className="relative flex min-h-0 min-w-0 flex-1 flex-col">
        <Grid
          wb={wb}
          sheet={activeSheet}
          layout={layout}
          gridlines={gridlines}
          freeze={wb.getFreeze(activeSheet)}
          sel={sel}
          editing={editing}
          scrollRef={scrollRef}
          onKeyDown={onGridKeyDown}
          onSelect={(s) => select(s, false)}
          onPointerStart={() => {
            if (editing) commitEdit();
            setMenu(null);
          }}
          onEdit={() => startEditing("edit")}
          onEditorChange={(value) => editing && setEditing({ ...editing, value })}
          onEditorKeyDown={onEditorKeyDown}
          onEditorFocus={(e) => {
            if (editing && editing.source !== "cell") setEditing({ ...editing, source: "cell", mode: "edit" });
            placeCaretAtEnd(e);
          }}
          onColResize={(col, width) => {
            wb.setColWidth(activeSheet, col, width === DEFAULT_COL_WIDTH ? undefined : width);
            changed();
          }}
          onRowResize={(row, height) => {
            wb.setRowHeight(activeSheet, row, height === 24 ? undefined : height);
            changed();
          }}
          onAutoFitCol={(col) => {
            wb.setColWidth(activeSheet, col, fitColumnWidth(col));
            changed();
          }}
          onFill={(src, dst) => {
            if (commitPatches(fillPatches(wb, activeSheet, src, dst, "series"))) selectRange(dst);
          }}
          onContextMenu={(x, y, target) => setMenu({ x, y, target })}
        />
        {findOpen && (
          <FindBar
            wb={wb}
            sheet={activeSheet}
            version={version}
            onGoTo={(cell) => select({ anchor: cell, focus: cell })}
            onClose={() => {
              setFindOpen(false);
              focusGrid();
            }}
          />
        )}
      </div>
      {aiTab && (
        <AIPanel
          wb={wb}
          sheet={activeSheet}
          version={version}
          selection={range}
          active={active}
          tab={aiTab}
          onTab={setAiTab}
          onClose={() => {
            setAiTab(null);
            focusGrid();
          }}
          onInsert={(cell, raw) => {
            commitPatches([{ sheet: activeSheet, row: cell.row, col: cell.col, raw }]);
            select({ anchor: cell, focus: cell });
          }}
          onGoTo={goToReference}
        />
      )}
      </div>

      {/* Tabs and status */}
      <footer className="flex h-10 items-center gap-4 border-t border-[var(--border)] bg-[var(--header-bg)] px-2 text-xs text-[var(--muted)]">
        <div className="min-w-0 flex-1">
          <SheetTabs commands={commands} state={state} />
        </div>
        {multiCell && stats.count > 0 && (
          <div className="hidden shrink-0 gap-4 sm:flex">
            {stats.numbers > 0 && <span>Sum: {displayValue(stats.sum)}</span>}
            {stats.numbers > 0 && <span>Average: {displayValue(stats.sum / stats.numbers)}</span>}
            <span>Count: {stats.count}</span>
          </div>
        )}
      </footer>

      {menu && <CellContextMenu {...menu} commands={commands} onClose={() => setMenu(null)} />}
      {printOpen && (
        <PrintDialog
          gridlines={gridlines}
          selection={multiCell ? range : null}
          sheetCount={wb.sheets.length}
          hasContent={(scope) => printSheets(scope).length > 0}
          onPrint={startPrint}
          onClose={() => {
            setPrintOpen(false);
            focusGrid();
          }}
        />
      )}
      {shortcutsOpen && (
        <ShortcutsDialog
          onClose={() => {
            setShortcutsOpen(false);
            focusGrid();
          }}
        />
      )}
      {toast && (
        <div
          role="status"
          className="fixed bottom-14 left-1/2 z-50 max-w-[90vw] -translate-x-1/2 rounded-lg bg-[var(--fg)] px-4 py-2 text-sm text-[var(--bg)] shadow-lg"
        >
          {toast}
        </div>
      )}
      <input
        ref={fileRef}
        type="file"
        accept=".csv,text/csv"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (file) void importFile(file);
        }}
      />
      </div>
      {printJob && <PrintView wb={wb} job={printJob} title={title} />}
    </>
  );
}
