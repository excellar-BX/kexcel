"use client";

import { useEffect, useRef, useState, type CSSProperties, type RefObject } from "react";
import {
  colToLetters,
  formatValue,
  isError,
  type CellAddress,
  type CellFormat,
  type Freeze,
  type Scalar,
  type Workbook,
} from "@/lib/engine";
import type { Look } from "@/lib/conditional";
import type { Range } from "@/lib/sheet-ops";
import { COL_COUNT, MIN_COL_WIDTH, OVERSCAN, ROW_COUNT, ROW_HEIGHT } from "./constants";
import { ChevronDown } from "@/components/ui/icons";
import { ptToPx, WRAP_LINE_HEIGHT, type GridLayout } from "./layout";
import type { Editing, Selection } from "./types";

const MIN_ROW_HEIGHT = 16;

export function rangeOf({ anchor, focus }: Selection): Range {
  return {
    r0: Math.min(anchor.row, focus.row),
    r1: Math.max(anchor.row, focus.row),
    c0: Math.min(anchor.col, focus.col),
    c1: Math.max(anchor.col, focus.col),
  };
}

export const inRange = (r: Range, row: number, col: number) =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

const FONT_STACKS = {
  sans: "var(--font-sans), system-ui, sans-serif",
  serif: "Georgia, 'Times New Roman', serif",
  mono: "var(--font-mono), ui-monospace, monospace",
};

/** CSS for a cell's text. */
export function textStyle(f: CellFormat | undefined, zoom: number): CSSProperties {
  const decoration = [f?.underline && "underline", f?.strike && "line-through"].filter(Boolean).join(" ");
  return {
    fontSize: ptToPx(f?.fontSize ?? 10) * zoom,
    fontFamily: f?.fontFamily ? FONT_STACKS[f.fontFamily] : undefined,
    fontWeight: f?.bold ? 700 : undefined,
    fontStyle: f?.italic ? "italic" : undefined,
    textDecoration: decoration || undefined,
    color: f?.color,
  };
}

interface Viewport {
  top: number;
  left: number;
  width: number;
  height: number;
}

export interface GridProps {
  wb: Workbook;
  sheet: string;
  layout: GridLayout;
  gridlines: boolean;
  freeze: Freeze;
  sel: Selection;
  editing: Editing | null;
  scrollRef: RefObject<HTMLDivElement | null>;
  onKeyDown: (e: React.KeyboardEvent) => void;
  onSelect: (sel: Selection) => void;
  /** Called before any mouse interaction so an open edit can be committed. */
  onPointerStart: () => void;
  onEdit: () => void;
  onEditorChange: (value: string) => void;
  onEditorKeyDown: (e: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>) => void;
  onEditorFocus: (e: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => void;
  onColResize: (col: number, width: number) => void;
  onRowResize: (row: number, height: number) => void;
  onAutoFitCol: (col: number) => void;
  onFill: (src: Range, dst: Range) => void;
  onContextMenu: (x: number, y: number, target: "cell" | "row" | "col") => void;
  /** Conditional formatting for a cell, if any applies. */
  look: (row: number, col: number, value: Scalar) => Look | null;
  /** Whether a cell holds a value its dropdown list doesn't offer. */
  invalid: (row: number, col: number) => boolean;
  /** Whether a cell has a dropdown list. */
  hasList: (row: number, col: number) => boolean;
  /** Opens the dropdown list under the given button. */
  onOpenList: (anchor: DOMRect) => void;
  /** Charts and other things floating over the cells. */
  objects?: React.ReactNode;
}

export default function Grid(props: GridProps) {
  const { wb, sheet, layout, gridlines, freeze, sel, editing, scrollRef } = props;
  const z = layout.zoom;
  const fr = Math.min(freeze.rows, ROW_COUNT - 1);
  const fc = Math.min(freeze.cols, COL_COUNT - 1);
  const frozenH = layout.rowTop(fr);
  const frozenW = layout.colLeft(fc);
  const frameRef = useRef(0);
  const [viewport, setViewport] = useState<Viewport>({ top: 0, left: 0, width: 1200, height: 800 });
  const [fillPreview, setFillPreview] = useState<Range | null>(null);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const observer = new ResizeObserver(() =>
      setViewport({ top: el.scrollTop, left: el.scrollLeft, width: el.clientWidth, height: el.clientHeight }),
    );
    observer.observe(el);
    return () => observer.disconnect();
  }, [scrollRef]);

  const onScroll = () => {
    cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(() => {
      const el = scrollRef.current;
      if (el) setViewport({ top: el.scrollTop, left: el.scrollLeft, width: el.clientWidth, height: el.clientHeight });
    });
  };

  // ---- Mouse -----------------------------------------------------------------

  /** The cell under a point. Inside the frozen area positions don't scroll; elsewhere they do. */
  const pointToCell = (clientX: number, clientY: number): CellAddress => {
    const el = scrollRef.current!;
    const rect = el.getBoundingClientRect();
    const vx = clientX - rect.left - layout.headerWidth;
    const vy = clientY - rect.top - layout.headerHeight;
    const x = vx < frozenW ? vx : vx + el.scrollLeft;
    const y = vy < frozenH ? vy : vy + el.scrollTop;
    return { row: layout.rowAt(y), col: layout.colAt(x) };
  };

  /** Scrolls when a drag reaches past the edge of the grid. */
  const autoScroll = (e: MouseEvent) => {
    const el = scrollRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const step = 24;
    if (e.clientY > rect.bottom - 8) el.scrollTop += step;
    else if (e.clientY < rect.top + layout.headerHeight) el.scrollTop -= step;
    if (e.clientX > rect.right - 8) el.scrollLeft += step;
    else if (e.clientX < rect.left + layout.headerWidth) el.scrollLeft -= step;
  };

  const trackDrag = (onMove: (e: MouseEvent) => void, onUp?: () => void) => {
    const move = (e: MouseEvent) => {
      autoScroll(e);
      onMove(e);
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      onUp?.();
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  };

  const range = rangeOf(sel);

  const onCellsMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    props.onPointerStart();
    const cell = pointToCell(e.clientX, e.clientY);
    const anchor = e.shiftKey ? sel.anchor : cell;
    props.onSelect({ anchor, focus: cell });
    trackDrag((m) => props.onSelect({ anchor, focus: pointToCell(m.clientX, m.clientY) }));
  };

  const onFillHandleMouseDown = (e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    props.onPointerStart();
    const src = range;
    let dst = src;
    trackDrag(
      (m) => {
        const cell = pointToCell(m.clientX, m.clientY);
        const dr = cell.row > src.r1 ? cell.row - src.r1 : cell.row < src.r0 ? cell.row - src.r0 : 0;
        const dc = cell.col > src.c1 ? cell.col - src.c1 : cell.col < src.c0 ? cell.col - src.c0 : 0;
        if (dr === 0 && dc === 0) dst = src;
        else if (Math.abs(dr) >= Math.abs(dc)) {
          dst = { ...src, r0: Math.min(src.r0, cell.row), r1: Math.max(src.r1, cell.row) };
        } else {
          dst = { ...src, c0: Math.min(src.c0, cell.col), c1: Math.max(src.c1, cell.col) };
        }
        setFillPreview(dst);
      },
      () => {
        setFillPreview(null);
        if (dst !== src) props.onFill(src, dst);
      },
    );
  };

  const onColumnHeaderMouseDown = (e: React.MouseEvent, col: number) => {
    if (e.button !== 0) return;
    props.onPointerStart();
    const startCol = e.shiftKey ? sel.anchor.col : col;
    const select = (c: number) =>
      props.onSelect({ anchor: { row: 0, col: startCol }, focus: { row: ROW_COUNT - 1, col: c } });
    select(col);
    trackDrag((m) => select(pointToCell(m.clientX, m.clientY).col));
  };

  const onRowHeaderMouseDown = (e: React.MouseEvent, row: number) => {
    if (e.button !== 0) return;
    props.onPointerStart();
    const startRow = e.shiftKey ? sel.anchor.row : row;
    const select = (r: number) =>
      props.onSelect({ anchor: { row: startRow, col: 0 }, focus: { row: r, col: COL_COUNT - 1 } });
    select(row);
    trackDrag((m) => select(pointToCell(m.clientX, m.clientY).row));
  };

  const onColResizeMouseDown = (e: React.MouseEvent, col: number) => {
    e.stopPropagation();
    e.preventDefault();
    const startX = e.clientX;
    const start = layout.baseColWidths[col];
    trackDrag((m) => props.onColResize(col, Math.max(MIN_COL_WIDTH, Math.round(start + (m.clientX - startX) / z))));
  };

  const onRowResizeMouseDown = (e: React.MouseEvent, row: number) => {
    e.stopPropagation();
    e.preventDefault();
    const startY = e.clientY;
    const start = layout.baseRowHeights[row];
    trackDrag((m) => props.onRowResize(row, Math.max(MIN_ROW_HEIGHT, Math.round(start + (m.clientY - startY) / z))));
  };

  const onContextMenu = (e: React.MouseEvent, target: "cell" | "row" | "col", at?: number) => {
    e.preventDefault();
    props.onPointerStart();
    if (target === "cell") {
      const cell = pointToCell(e.clientX, e.clientY);
      if (!inRange(range, cell.row, cell.col)) props.onSelect({ anchor: cell, focus: cell });
    } else if (target === "col" && at !== undefined && (at < range.c0 || at > range.c1 || range.r1 - range.r0 < ROW_COUNT - 1)) {
      props.onSelect({ anchor: { row: 0, col: at }, focus: { row: ROW_COUNT - 1, col: at } });
    } else if (target === "row" && at !== undefined && (at < range.r0 || at > range.r1 || range.c1 - range.c0 < COL_COUNT - 1)) {
      props.onSelect({ anchor: { row: at, col: 0 }, focus: { row: at, col: COL_COUNT - 1 } });
    }
    props.onContextMenu(e.clientX, e.clientY, target);
  };

  // ---- Visible window --------------------------------------------------------

  // Frozen rows and columns are drawn in their own pinned layers; the main layer
  // scrolls underneath them. All layers share the grid's coordinates.
  const rowStart = Math.max(fr, layout.rowAt(viewport.top + frozenH) - OVERSCAN);
  const rowEnd = Math.min(ROW_COUNT - 1, layout.rowAt(viewport.top + viewport.height) + OVERSCAN);
  const colStart = Math.max(fc, layout.colAt(viewport.left + frozenW) - 1);
  const colEnd = Math.min(COL_COUNT - 1, layout.colAt(viewport.left + viewport.width) + 1);

  type Layer = "main" | "top" | "left" | "corner";
  const layerOf = (row: number, col: number): Layer =>
    row < fr ? (col < fc ? "corner" : "top") : col < fc ? "left" : "main";

  /** One cell's text, spanning `width` x `height`. Wrapped text breaks lines instead of spilling. */
  const cellText = (
    key: string,
    value: Scalar,
    format: CellFormat | undefined,
    left: number,
    top: number,
    width: number,
    height: number,
    spillTo: number,
  ) => {
    if (value === null) return null;
    const error = isError(value);
    const defaultAlign = typeof value === "number" ? "right" : typeof value === "string" ? "left" : "center";
    const align = format?.align ?? defaultAlign;
    const wrap = !!format?.wrap;
    let textWidth = width;
    // Text spills into empty neighbours on the right, as in other spreadsheets.
    if (!wrap && align === "left" && typeof value === "string") {
      const row = layout.rowAt(top);
      for (let c = layout.colAt(left + width); c <= spillTo && wb.getRaw(sheet, row, c) === "" && !wb.mergeAt(sheet, row, c); c++) {
        textWidth += layout.colWidth(c);
      }
    }
    return (
      <div
        key={key}
        className={`pointer-events-none absolute overflow-hidden ${wrap ? "whitespace-pre-wrap break-words" : "whitespace-pre"} ${
          error ? "text-[var(--error)]" : ""
        }`}
        style={{
          left,
          top,
          width: textWidth,
          height,
          lineHeight: wrap ? WRAP_LINE_HEIGHT : `${height}px`,
          padding: wrap ? `${3 * z}px ${4 * z}px` : `0 ${4 * z}px`,
          textAlign: align,
          ...textStyle(format, z),
        }}
        title={error ? value.message : undefined}
      >
        {formatValue(value, format)}
      </div>
    );
  };

  /** Background colour and borders for a cell or merged range. */
  const cellFill = (key: string, format: CellFormat | undefined, left: number, top: number, width: number, height: number) => {
    if (!format?.fill && !format?.borders) return null;
    const b = format.borders;
    const line = "1px solid var(--fg)";
    return (
      <div
        key={key}
        className="absolute"
        style={{
          left: left - (b?.left ? 1 : 0),
          top: top - (b?.top ? 1 : 0),
          width: width + (b?.left ? 1 : 0),
          height: height + (b?.top ? 1 : 0),
          background: format.fill,
          borderTop: b?.top ? line : undefined,
          borderLeft: b?.left ? line : undefined,
          borderRight: b?.right ? line : undefined,
          borderBottom: b?.bottom ? line : undefined,
        }}
      />
    );
  };

  /** A cell's format with conditional formatting laid over it. */
  const withLook = (format: CellFormat | undefined, look: Look | null): CellFormat | undefined =>
    look
      ? { ...format, fill: look.fill ?? format?.fill, color: look.color ?? format?.color, bold: look.bold || format?.bold }
      : format;

  /** A data bar: a tint of the first series colour, as long as the value's share of the range. */
  const dataBar = (key: string, fraction: number, left: number, top: number, width: number, height: number) => (
    <div
      key={key}
      className="absolute rounded-r-[3px]"
      style={{
        left: left + 1,
        top: top + 3 * z,
        width: Math.max(0, (width - 3) * fraction),
        height: Math.max(0, height - 7 * z),
        background: "color-mix(in oklab, var(--series-1) 35%, transparent)",
      }}
    />
  );

  /** A small corner flag on cells whose value isn't in their dropdown list. */
  const invalidMark = (key: string, right: number, top: number) => (
    <div
      key={key}
      className="pointer-events-none absolute"
      title="This value isn't in the cell's list"
      style={{
        left: right - 7,
        top,
        width: 0,
        height: 0,
        borderTop: "6px solid var(--error)",
        borderLeft: "6px solid transparent",
      }}
    />
  );

  /** Fills, borders, gridlines and text for a block of cells, with merged ranges drawn as one cell. */
  const renderCells = (r0: number, r1: number, c0: number, c1: number): React.ReactNode[] => {
    const fills: React.ReactNode[] = [];
    const lines: React.ReactNode[] = [];
    const covers: React.ReactNode[] = [];
    const texts: React.ReactNode[] = [];

    const merges = wb.merges(sheet).filter((m) => m.r1 >= r0 && m.r0 <= r1 && m.c1 >= c0 && m.c0 <= c1);
    const covered = (row: number, col: number) =>
      merges.some((m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1);

    for (let row = r0; row <= r1; row++) {
      const top = layout.rowTop(row);
      const height = layout.rowHeight(row);
      for (let col = c0; col <= c1; col++) {
        if (merges.length && covered(row, col)) continue;
        const value = wb.getValue(sheet, row, col);
        const look = props.look(row, col, value);
        const format = withLook(wb.getFormat(sheet, row, col), look);
        const left = layout.colLeft(col);
        const width = layout.colWidth(col);
        fills.push(cellFill(`f${row},${col}`, format, left, top, width, height));
        if (look?.bar) fills.push(dataBar(`b${row},${col}`, look.bar, left, top, width, height));
        if (value !== null && props.invalid(row, col)) texts.push(invalidMark(`i${row},${col}`, left + width, top));
        texts.push(cellText(`t${row},${col}`, value, format, left, top, width, height, c1));
      }
    }

    for (const m of merges) {
      // The top-left cell's content and look fill the whole range; a cover hides the gridlines inside it.
      const value = wb.getValue(sheet, m.r0, m.c0);
      const look = props.look(m.r0, m.c0, value);
      const format = withLook(wb.getFormat(sheet, m.r0, m.c0), look);
      const left = layout.colLeft(m.c0);
      const top = layout.rowTop(m.r0);
      const width = layout.colLeft(m.c1 + 1) - left;
      const height = layout.rowTop(m.r1 + 1) - top;
      const key = `${m.r0},${m.c0}`;
      covers.push(
        <div
          key={`m${key}`}
          className="absolute"
          style={{ left, top, width: width - 1, height: height - 1, background: format?.fill ?? "var(--bg)" }}
        />,
      );
      covers.push(cellFill(`mf${key}`, format?.borders ? { borders: format.borders } : undefined, left, top, width, height));
      if (look?.bar) covers.push(dataBar(`mb${key}`, look.bar, left, top, width, height));
      texts.push(cellText(`mt${key}`, value, format, left, top, width, height, m.c1));
    }

    if (gridlines) {
      for (let col = c0; col <= c1; col++) {
        lines.push(
          <div
            key={`v${col}`}
            className="absolute top-0 h-full w-px bg-[var(--grid-line)]"
            style={{ left: layout.colLeft(col + 1) - 1 }}
          />,
        );
      }
      for (let row = r0; row <= r1; row++) {
        lines.push(
          <div
            key={`h${row}`}
            className="absolute left-0 h-px w-full bg-[var(--grid-line)]"
            style={{ top: layout.rowTop(row + 1) - 1 }}
          />,
        );
      }
    }
    return [...fills, ...lines, ...covers, ...texts];
  };

  const box = (r: Range) => ({
    left: layout.colLeft(r.c0) - 1,
    top: layout.rowTop(r.r0) - 1,
    width: layout.colLeft(r.c1 + 1) - layout.colLeft(r.c0) + 1,
    height: layout.rowTop(r.r1 + 1) - layout.rowTop(r.r0) + 1,
  });
  const selBox = box(range);
  const active = sel.anchor;
  const cellBox = (row: number, col: number) => box(wb.mergeAt(sheet, row, col) ?? { r0: row, r1: row, c0: col, c1: col });
  const headerFont = 11 * z;

  /** The editor grows to show long or multi-line content (Alt+Enter adds a line). */
  const editorBox = () => {
    if (!editing) return {};
    const b = cellBox(editing.row, editing.col);
    const f = wb.getFormat(sheet, editing.row, editing.col);
    const fontPx = ptToPx(f?.fontSize ?? 10) * z;
    const lines = editing.value.split("\n");
    const longest = Math.max(...lines.map((l) => l.length));
    const width = f?.wrap ? b.width : Math.min(600 * z, Math.max(b.width, (longest * 8 + 24) * z));
    const height = Math.max(b.height, Math.ceil(lines.length * fontPx * WRAP_LINE_HEIGHT + 8 * z));
    return { ...b, width, height };
  };

  /** Selection, fill handle and editor, drawn in every layer and clipped by it. */
  const overlays = (layer: Layer) => (
    <>
      <div className="pointer-events-none absolute border border-[var(--accent)] bg-[var(--selection)]" style={selBox} />
      <div
        className="pointer-events-none absolute border-2 border-[var(--accent)]"
        style={cellBox(active.row, active.col)}
      />
      {!editing && props.hasList(active.row, active.col) && layerOf(active.row, active.col) === layer && (
        <button
          type="button"
          data-list-button=""
          aria-label="Show the list for this cell"
          onMouseDown={(e) => e.stopPropagation()}
          onClick={(e) => props.onOpenList(e.currentTarget.getBoundingClientRect())}
          className="absolute z-[6] grid place-items-center rounded-sm bg-[var(--bg)] text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--fg)]"
          style={(() => {
            const b = cellBox(active.row, active.col);
            const size = Math.min(18 * z, b.height - 4);
            return { left: b.left + b.width - size - 2, top: b.top + (b.height - size) / 2, width: size, height: size };
          })()}
        >
          <ChevronDown size={Math.round(13 * z)} />
        </button>
      )}
      {fillPreview && (
        <div className="pointer-events-none absolute border-2 border-dashed border-[var(--accent)]" style={box(fillPreview)} />
      )}
      {!editing && layerOf(range.r1, range.c1) === layer && (
        <div
          onMouseDown={onFillHandleMouseDown}
          title="Drag to fill"
          className="absolute z-[5] h-2 w-2 cursor-crosshair border border-[var(--bg)] bg-[var(--accent)]"
          style={{ left: selBox.left + selBox.width - 4, top: selBox.top + selBox.height - 4 }}
        />
      )}
      {editing && layerOf(editing.row, editing.col) === layer && (
        <textarea
          value={editing.value}
          rows={1}
          autoFocus={editing.source === "cell"}
          onFocus={props.onEditorFocus}
          onChange={(e) => props.onEditorChange(e.target.value)}
          onKeyDown={props.onEditorKeyDown}
          onMouseDown={(e) => e.stopPropagation()}
          onDoubleClick={(e) => e.stopPropagation()}
          spellCheck={false}
          aria-label="Cell editor"
          className="absolute z-10 resize-none overflow-hidden border-2 border-[var(--accent)] bg-[var(--bg)] text-[var(--fg)] shadow-lg outline-none"
          style={{
            ...editorBox(),
            padding: `${2 * z}px ${3 * z}px`,
            lineHeight: WRAP_LINE_HEIGHT,
            whiteSpace: wb.getFormat(sheet, editing.row, editing.col)?.wrap ? "pre-wrap" : "pre",
            ...textStyle(wb.getFormat(sheet, editing.row, editing.col), z),
            color: undefined,
          }}
        />
      )}
    </>
  );

  const cellHandlers = {
    onMouseDown: onCellsMouseDown,
    onDoubleClick: props.onEdit,
    onContextMenu: (e: React.MouseEvent) => onContextMenu(e, "cell"),
  };

  /** The heavier line along the edge of the frozen area. */
  const freezeEdge = (side: "bottom" | "right") => (
    <div
      className={`pointer-events-none absolute bg-[var(--muted)]/45 ${
        side === "bottom" ? "bottom-0 left-0 h-[2px] w-full" : "right-0 top-0 h-full w-[2px]"
      }`}
    />
  );

  const colHeader = (col: number) => {
    const selected = col >= range.c0 && col <= range.c1;
    return (
      <div
        key={col}
        onMouseDown={(e) => onColumnHeaderMouseDown(e, col)}
        onContextMenu={(e) => onContextMenu(e, "col", col)}
        className={`absolute top-0 grid h-full place-items-center border-r border-[var(--grid-line)] ${
          selected ? "bg-[var(--header-selected)] font-semibold text-[var(--accent)]" : "text-[var(--muted)]"
        }`}
        style={{ left: layout.colLeft(col), width: layout.colWidth(col), fontSize: headerFont }}
      >
        {colToLetters(col)}
        <div
          onMouseDown={(e) => onColResizeMouseDown(e, col)}
          onDoubleClick={() => props.onAutoFitCol(col)}
          title="Drag to resize, double-click to fit"
          className="absolute -right-1 top-0 z-10 h-full w-2 cursor-col-resize hover:bg-[var(--accent)]/40"
        />
      </div>
    );
  };

  const rowHeader = (row: number) => {
    const selected = row >= range.r0 && row <= range.r1;
    return (
      <div
        key={row}
        onMouseDown={(e) => onRowHeaderMouseDown(e, row)}
        onContextMenu={(e) => onContextMenu(e, "row", row)}
        className={`absolute left-0 grid w-full place-items-center border-b border-[var(--grid-line)] ${
          selected ? "bg-[var(--header-selected)] font-semibold text-[var(--accent)]" : "text-[var(--muted)]"
        }`}
        style={{ top: layout.rowTop(row), height: layout.rowHeight(row), fontSize: headerFont }}
      >
        {row + 1}
        <div
          onMouseDown={(e) => onRowResizeMouseDown(e, row)}
          onDoubleClick={() => props.onRowResize(row, ROW_HEIGHT)}
          title="Drag to resize, double-click to reset"
          className="absolute -bottom-1 left-0 z-10 h-2 w-full cursor-row-resize hover:bg-[var(--accent)]/40"
        />
      </div>
    );
  };

  const span = (from: number, to: number) => Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => from + i);
  const headerW = layout.headerWidth;
  const headerH = layout.headerHeight;

  return (
    <div
      ref={scrollRef}
      tabIndex={0}
      onScroll={onScroll}
      onKeyDown={props.onKeyDown}
      className="relative min-h-0 flex-1 select-none overflow-auto outline-none"
      role="grid"
      aria-label="Spreadsheet"
      aria-rowcount={ROW_COUNT}
      aria-colcount={COL_COUNT}
    >
      <div style={{ width: headerW + layout.totalWidth, height: headerH + layout.totalHeight }}>
        {/* Column headers */}
        <div
          className="sticky top-0 z-30 flex border-b border-[var(--border)] bg-[var(--header-bg)]"
          style={{ height: headerH, width: headerW + layout.totalWidth }}
        >
          <div
            className="sticky left-0 z-40 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)] hover:bg-[var(--hover)]"
            style={{ width: headerW }}
            onMouseDown={() => {
              props.onPointerStart();
              props.onSelect({ anchor: { row: 0, col: 0 }, focus: { row: ROW_COUNT - 1, col: COL_COUNT - 1 } });
            }}
            title="Select all"
          />
          {fc > 0 && (
            <div
              className="sticky z-[35] h-full shrink-0 bg-[var(--header-bg)]"
              style={{ left: headerW, width: frozenW, marginRight: -frozenW }}
            >
              {span(0, fc - 1).map(colHeader)}
              {freezeEdge("right")}
            </div>
          )}
          <div className="relative" style={{ width: layout.totalWidth }}>
            {span(colStart, colEnd).map(colHeader)}
          </div>
        </div>

        {/* Frozen rows, pinned under the column headers */}
        {fr > 0 && (
          <div
            className="sticky z-20 flex"
            style={{ top: headerH, height: frozenH, marginBottom: -frozenH, width: headerW + layout.totalWidth }}
          >
            <div
              className="sticky left-0 z-[3] shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)]"
              style={{ width: headerW, height: frozenH }}
            >
              {span(0, fr - 1).map(rowHeader)}
              {freezeEdge("bottom")}
            </div>
            {fc > 0 && (
              <div
                className="sticky z-[2] shrink-0 overflow-hidden bg-[var(--bg)]"
                style={{ left: headerW, width: frozenW, height: frozenH, marginRight: -frozenW }}
                {...cellHandlers}
              >
                {renderCells(0, fr - 1, 0, fc - 1)}
                {overlays("corner")}
                {freezeEdge("bottom")}
                {freezeEdge("right")}
              </div>
            )}
            <div
              className="relative shrink-0 overflow-hidden bg-[var(--bg)]"
              style={{ width: layout.totalWidth, height: frozenH }}
              {...cellHandlers}
            >
              {renderCells(0, fr - 1, colStart, colEnd)}
              {overlays("top")}
              {freezeEdge("bottom")}
            </div>
          </div>
        )}

        <div className="flex">
          {/* Row headers */}
          <div
            className="sticky left-0 z-10 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)]"
            style={{ width: headerW, height: layout.totalHeight }}
          >
            {span(rowStart, rowEnd).map(rowHeader)}
          </div>

          {/* Frozen columns, pinned beside the row headers */}
          {fc > 0 && (
            <div
              className="sticky z-[5] shrink-0 overflow-hidden bg-[var(--bg)]"
              style={{ left: headerW, width: frozenW, height: layout.totalHeight, marginRight: -frozenW }}
              {...cellHandlers}
            >
              {renderCells(rowStart, rowEnd, 0, fc - 1)}
              {overlays("left")}
              {freezeEdge("right")}
            </div>
          )}

          {/* Cells */}
          <div
            className="relative shrink-0 overflow-hidden"
            style={{ width: layout.totalWidth, height: layout.totalHeight }}
            {...cellHandlers}
          >
            {renderCells(rowStart, rowEnd, colStart, colEnd)}
            {overlays("main")}
            {props.objects}
          </div>
        </div>
      </div>
    </div>
  );
}
