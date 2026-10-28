"use client";

import { useEffect, useRef, useState, type CSSProperties, type RefObject } from "react";
import { colToLetters, formatValue, isError, type CellAddress, type CellFormat, type Workbook } from "@/lib/engine";
import type { Range } from "@/lib/sheet-ops";
import { COL_COUNT, MIN_COL_WIDTH, OVERSCAN, ROW_COUNT, ROW_HEIGHT } from "./constants";
import { ptToPx, type GridLayout } from "./layout";
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
  sel: Selection;
  editing: Editing | null;
  scrollRef: RefObject<HTMLDivElement | null>;
  onKeyDown: (e: React.KeyboardEvent) => void;
  onSelect: (sel: Selection) => void;
  /** Called before any mouse interaction so an open edit can be committed. */
  onPointerStart: () => void;
  onEdit: () => void;
  onEditorChange: (value: string) => void;
  onEditorKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => void;
  onEditorFocus: (e: React.FocusEvent<HTMLInputElement>) => void;
  onColResize: (col: number, width: number) => void;
  onRowResize: (row: number, height: number) => void;
  onAutoFitCol: (col: number) => void;
  onFill: (src: Range, dst: Range) => void;
  onContextMenu: (x: number, y: number, target: "cell" | "row" | "col") => void;
}

export default function Grid(props: GridProps) {
  const { wb, sheet, layout, gridlines, sel, editing, scrollRef } = props;
  const z = layout.zoom;
  const cellsRef = useRef<HTMLDivElement>(null);
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

  const pointToCell = (clientX: number, clientY: number): CellAddress => {
    const rect = cellsRef.current!.getBoundingClientRect();
    return { row: layout.rowAt(clientY - rect.top), col: layout.colAt(clientX - rect.left) };
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

  const rowStart = Math.max(0, layout.rowAt(viewport.top) - OVERSCAN);
  const rowEnd = Math.min(ROW_COUNT - 1, layout.rowAt(viewport.top + viewport.height) + OVERSCAN);
  const colStart = Math.max(0, layout.colAt(viewport.left) - 1);
  const colEnd = Math.min(COL_COUNT - 1, layout.colAt(viewport.left + viewport.width) + 1);

  const fills: React.ReactNode[] = [];
  const texts: React.ReactNode[] = [];
  for (let row = rowStart; row <= rowEnd; row++) {
    const top = layout.rowTop(row);
    const height = layout.rowHeight(row);
    for (let col = colStart; col <= colEnd; col++) {
      const format = wb.getFormat(sheet, row, col);
      const left = layout.colLeft(col);
      const width = layout.colWidth(col);

      if (format?.fill || format?.borders) {
        const b = format.borders;
        const line = "1px solid var(--fg)";
        fills.push(
          <div
            key={`f${row},${col}`}
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
          />,
        );
      }

      const value = wb.getValue(sheet, row, col);
      if (value === null) continue;
      const error = isError(value);
      const defaultAlign = typeof value === "number" ? "right" : typeof value === "string" ? "left" : "center";
      const align = format?.align ?? defaultAlign;
      let textWidth = width;
      // Text spills into empty neighbours on the right, as in other spreadsheets.
      if (align === "left" && typeof value === "string") {
        for (let c = col + 1; c <= colEnd && wb.getRaw(sheet, row, c) === ""; c++) textWidth += layout.colWidth(c);
      }
      texts.push(
        <div
          key={`t${row},${col}`}
          className={`pointer-events-none absolute overflow-hidden whitespace-pre ${error ? "text-[var(--error)]" : ""}`}
          style={{
            left,
            top,
            width: textWidth,
            height,
            lineHeight: `${height}px`,
            padding: `0 ${4 * z}px`,
            textAlign: align,
            ...textStyle(format, z),
          }}
          title={error ? value.message : undefined}
        >
          {formatValue(value, format)}
        </div>,
      );
    }
  }

  const lines: React.ReactNode[] = [];
  if (gridlines) {
    for (let col = colStart; col <= colEnd; col++) {
      lines.push(
        <div
          key={`v${col}`}
          className="absolute top-0 h-full w-px bg-[var(--grid-line)]"
          style={{ left: layout.colLeft(col + 1) - 1 }}
        />,
      );
    }
    for (let row = rowStart; row <= rowEnd; row++) {
      lines.push(
        <div
          key={`h${row}`}
          className="absolute left-0 h-px w-full bg-[var(--grid-line)]"
          style={{ top: layout.rowTop(row + 1) - 1 }}
        />,
      );
    }
  }

  const box = (r: Range) => ({
    left: layout.colLeft(r.c0) - 1,
    top: layout.rowTop(r.r0) - 1,
    width: layout.colLeft(r.c1 + 1) - layout.colLeft(r.c0) + 1,
    height: layout.rowTop(r.r1 + 1) - layout.rowTop(r.r0) + 1,
  });
  const selBox = box(range);
  const active = sel.anchor;
  const headerFont = 11 * z;

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
      <div style={{ width: layout.headerWidth + layout.totalWidth, height: layout.headerHeight + layout.totalHeight }}>
        {/* Column headers */}
        <div
          className="sticky top-0 z-20 flex border-b border-[var(--border)] bg-[var(--header-bg)]"
          style={{ height: layout.headerHeight, width: layout.headerWidth + layout.totalWidth }}
        >
          <div
            className="sticky left-0 z-30 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)] hover:bg-[var(--hover)]"
            style={{ width: layout.headerWidth }}
            onMouseDown={() => {
              props.onPointerStart();
              props.onSelect({ anchor: { row: 0, col: 0 }, focus: { row: ROW_COUNT - 1, col: COL_COUNT - 1 } });
            }}
            title="Select all"
          />
          <div className="relative" style={{ width: layout.totalWidth }}>
            {Array.from({ length: colEnd - colStart + 1 }, (_, i) => {
              const col = colStart + i;
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
            })}
          </div>
        </div>

        <div className="flex">
          {/* Row headers */}
          <div
            className="sticky left-0 z-10 shrink-0 border-r border-[var(--border)] bg-[var(--header-bg)]"
            style={{ width: layout.headerWidth, height: layout.totalHeight }}
          >
            {Array.from({ length: rowEnd - rowStart + 1 }, (_, i) => {
              const row = rowStart + i;
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
            })}
          </div>

          {/* Cells */}
          <div
            ref={cellsRef}
            className="relative overflow-hidden"
            style={{ width: layout.totalWidth, height: layout.totalHeight }}
            onMouseDown={onCellsMouseDown}
            onDoubleClick={props.onEdit}
            onContextMenu={(e) => onContextMenu(e, "cell")}
          >
            {fills}
            {lines}
            {texts}

            <div
              className="pointer-events-none absolute border border-[var(--accent)] bg-[var(--selection)]"
              style={selBox}
            />
            <div
              className="pointer-events-none absolute border-2 border-[var(--accent)]"
              style={box({ r0: active.row, r1: active.row, c0: active.col, c1: active.col })}
            />
            {fillPreview && (
              <div
                className="pointer-events-none absolute border-2 border-dashed border-[var(--accent)]"
                style={box(fillPreview)}
              />
            )}
            {!editing && (
              <div
                onMouseDown={onFillHandleMouseDown}
                title="Drag to fill"
                className="absolute z-[5] h-2 w-2 cursor-crosshair border border-[var(--bg)] bg-[var(--accent)]"
                style={{ left: selBox.left + selBox.width - 4, top: selBox.top + selBox.height - 4 }}
              />
            )}

            {editing && (
              <input
                value={editing.value}
                autoFocus={editing.source === "cell"}
                onFocus={props.onEditorFocus}
                onChange={(e) => props.onEditorChange(e.target.value)}
                onKeyDown={props.onEditorKeyDown}
                onMouseDown={(e) => e.stopPropagation()}
                onDoubleClick={(e) => e.stopPropagation()}
                spellCheck={false}
                aria-label="Cell editor"
                className="absolute z-10 border-2 border-[var(--accent)] bg-[var(--bg)] text-[var(--fg)] shadow-lg outline-none"
                style={{
                  ...box({ r0: editing.row, r1: editing.row, c0: editing.col, c1: editing.col }),
                  width: Math.min(
                    600 * z,
                    Math.max(layout.colWidth(editing.col) + 1, (editing.value.length * 8 + 24) * z),
                  ),
                  padding: `0 ${3 * z}px`,
                  ...textStyle(wb.getFormat(sheet, editing.row, editing.col), z),
                  color: undefined,
                }}
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
