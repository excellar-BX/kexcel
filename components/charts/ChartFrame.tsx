"use client";

import { useState } from "react";
import type { ChartData } from "@/lib/charts";
import type { ChartSpec } from "@/lib/engine";
import ChartView from "./ChartView";

const MIN_W = 220;
const MIN_H = 160;
const TITLE_H = 26;

interface Props {
  chart: ChartSpec;
  data: ChartData;
  zoom: number;
  selected: boolean;
  onSelect: () => void;
  /** Called once a move or resize ends, with unscaled pixels. */
  onChange: (patch: Pick<ChartSpec, "x" | "y" | "width" | "height">) => void;
  onEdit: () => void;
  onDelete: () => void;
}

/** A chart placed on the sheet: drag to move, drag the corner to resize. */
export default function ChartFrame({ chart, data, zoom, selected, onSelect, onChange, onEdit, onDelete }: Props) {
  // While dragging, the frame moves on its own; the sheet only hears about it at the end.
  const [live, setLive] = useState<Pick<ChartSpec, "x" | "y" | "width" | "height"> | null>(null);
  const box = live ?? chart;

  const drag = (e: React.MouseEvent, mode: "move" | "resize") => {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    onSelect();
    const startX = e.clientX;
    const startY = e.clientY;
    const start = { x: chart.x, y: chart.y, width: chart.width, height: chart.height };
    let last = start;
    const move = (m: MouseEvent) => {
      const dx = (m.clientX - startX) / zoom;
      const dy = (m.clientY - startY) / zoom;
      last =
        mode === "move"
          ? { ...start, x: Math.max(0, Math.round(start.x + dx)), y: Math.max(0, Math.round(start.y + dy)) }
          : { ...start, width: Math.max(MIN_W, Math.round(start.width + dx)), height: Math.max(MIN_H, Math.round(start.height + dy)) };
      setLive(last);
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      setLive(null);
      if (last !== start) onChange(last);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  };

  const width = box.width * zoom;
  const height = box.height * zoom;

  return (
    <div
      role="figure"
      aria-label={chart.title || "Chart"}
      onMouseDown={(e) => drag(e, "move")}
      onDoubleClick={(e) => {
        e.stopPropagation();
        onEdit();
      }}
      onContextMenu={(e) => e.stopPropagation()}
      className={`absolute z-[6] flex cursor-move select-none flex-col rounded-lg border bg-[var(--bg)] shadow-md ${
        selected ? "border-[var(--accent)] ring-2 ring-[var(--accent)]/30" : "border-[var(--border)]"
      }`}
      style={{ left: box.x * zoom, top: box.y * zoom, width, height }}
    >
      <div className="flex shrink-0 items-center gap-1 px-3" style={{ height: TITLE_H }}>
        <span className="min-w-0 flex-1 truncate text-sm font-medium">{chart.title}</span>
        {selected && (
          <>
            <button
              type="button"
              onMouseDown={(e) => e.stopPropagation()}
              onClick={onEdit}
              className="rounded px-1.5 py-0.5 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--fg)]"
            >
              Edit
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.stopPropagation()}
              onClick={onDelete}
              className="rounded px-1.5 py-0.5 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
            >
              Delete
            </button>
          </>
        )}
      </div>
      <div className="flex min-h-0 flex-1 cursor-default px-2 pb-2">
        <ChartView data={data} type={chart.type} width={Math.max(40, width - 16)} height={Math.max(40, height - TITLE_H - 8)} />
      </div>
      {data.truncated && (
        <div className="absolute bottom-1 left-3 text-[10px] text-[var(--muted)]">Showing the first 8 series and 200 rows</div>
      )}
      <div
        onMouseDown={(e) => drag(e, "resize")}
        title="Drag to resize"
        className="absolute -bottom-1 -right-1 h-3.5 w-3.5 cursor-nwse-resize rounded-sm border-2 border-[var(--bg)] bg-[var(--accent)]"
        style={{ display: selected ? undefined : "none" }}
      />
    </div>
  );
}
