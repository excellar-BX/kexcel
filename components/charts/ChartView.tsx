"use client";

import { useState } from "react";
import { compactNumber, MAX_PIE_SLICES, niceTicks, type ChartData } from "@/lib/charts";
import type { ChartType } from "@/lib/engine";

const seriesColor = (i: number) => `var(--series-${(i % 8) + 1})`;
const BAR_MAX = 24;
const GAP = 2;

interface Hover {
  x: number;
  y: number;
  title: string;
  rows: { name: string; value: string; color: string }[];
}

const fmt = (v: number | null) => (v === null ? "No value" : Number(v.toPrecision(10)).toLocaleString());

/** Rounded data end, square at the baseline, as one path. */
function barPath(x: number, y: number, w: number, h: number, horizontal: boolean, negative: boolean) {
  const r = Math.min(4, horizontal ? h / 2 : w / 2, horizontal ? Math.abs(w) : Math.abs(h));
  if (!horizontal) {
    const top = negative ? y + h : y;
    const bottom = negative ? y : y + h;
    // Round the end away from the baseline.
    return negative
      ? `M${x},${bottom} h${w} v${top - bottom - r} q0,${r} -${r},${r} h${-(w - 2 * r)} q-${r},0 -${r},-${r} Z`
      : `M${x},${bottom} v${-(bottom - top - r)} q0,-${r} ${r},-${r} h${w - 2 * r} q${r},0 ${r},${r} v${bottom - top - r} Z`;
  }
  const left = negative ? x + w : x;
  const right = negative ? x : x + w;
  return negative
    ? `M${right},${y} v${h} h${left - right + r} q-${r},0 -${r},-${r} v${-(h - 2 * r)} q0,-${r} ${r},-${r} Z`
    : `M${left},${y} h${right - left - r} q${r},0 ${r},${r} v${h - 2 * r} q0,${r} -${r},${r} h${-(right - left - r)} Z`;
}

function Legend({ names }: { names: string[] }) {
  if (names.length < 2) return null;
  return (
    <div className="flex flex-wrap justify-center gap-x-4 gap-y-1 px-2 pb-1 text-[11px] text-[var(--muted)]">
      {names.map((n, i) => (
        <span key={i} className="flex items-center gap-1.5">
          <span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ background: seriesColor(i) }} />
          {n}
        </span>
      ))}
    </div>
  );
}

function Tooltip({ hover }: { hover: Hover | null }) {
  if (!hover) return null;
  return (
    <div
      className="pointer-events-none absolute z-10 min-w-28 rounded-md border border-[var(--border)] bg-[var(--panel)] px-2.5 py-1.5 text-xs shadow-lg"
      style={{ left: hover.x + 12, top: Math.max(0, hover.y - 10) }}
    >
      <div className="mb-1 font-medium">{hover.title}</div>
      {hover.rows.map((r, i) => (
        <div key={i} className="flex items-center gap-2">
          <span className="inline-block h-2 w-2 rounded-sm" style={{ background: r.color }} />
          <span className="flex-1 text-[var(--muted)]">{r.name}</span>
          <span className="tabular-nums">{r.value}</span>
        </div>
      ))}
    </div>
  );
}

function PieChart({ data, width, height, onHover }: { data: ChartData; width: number; height: number; onHover: (h: Hover | null) => void }) {
  const first = data.series[0];
  let slices = data.categories
    .map((name, i) => ({ name, value: first?.values[i] ?? 0 }))
    .filter((s): s is { name: string; value: number } => typeof s.value === "number" && s.value > 0);
  // Past eight slices the rest fold into "Other" instead of inventing more colours.
  if (slices.length > MAX_PIE_SLICES) {
    const rest = slices.slice(MAX_PIE_SLICES - 1).reduce((sum, s) => sum + s.value, 0);
    slices = [...slices.slice(0, MAX_PIE_SLICES - 1), { name: "Other", value: rest }];
  }
  const total = slices.reduce((sum, s) => sum + s.value, 0);
  if (total <= 0) return <Empty text="A pie chart needs positive numbers." />;
  const cx = width / 2;
  const cy = height / 2;
  const r = Math.max(10, Math.min(width, height) / 2 - 8);
  // Each slice's start and end angle, clockwise from twelve o'clock.
  const arcs: { a0: number; a1: number }[] = [];
  for (let i = 0, angle = -Math.PI / 2; i < slices.length; i++) {
    const a1 = angle + (slices[i].value / total) * Math.PI * 2;
    arcs.push({ a0: angle, a1 });
    angle = a1;
  }
  return (
    <>
      <svg width={width} height={height} role="img" aria-label="Pie chart">
        {slices.map((s, i) => {
          const { a0, a1 } = arcs[i];
          const sweep = a1 - a0;
          const large = sweep > Math.PI ? 1 : 0;
          const p0 = [cx + r * Math.cos(a0), cy + r * Math.sin(a0)];
          const p1 = [cx + r * Math.cos(a1), cy + r * Math.sin(a1)];
          const d =
            slices.length === 1
              ? `M${cx - r},${cy} a${r},${r} 0 1,0 ${2 * r},0 a${r},${r} 0 1,0 ${-2 * r},0`
              : `M${cx},${cy} L${p0[0]},${p0[1]} A${r},${r} 0 ${large} 1 ${p1[0]},${p1[1]} Z`;
          return (
            <path
              key={i}
              d={d}
              fill={seriesColor(i)}
              stroke="var(--bg)"
              strokeWidth={GAP}
              onMouseMove={(e) => {
                const box = (e.currentTarget.ownerSVGElement as SVGSVGElement).getBoundingClientRect();
                onHover({
                  x: e.clientX - box.left,
                  y: e.clientY - box.top,
                  title: s.name,
                  rows: [{ name: `${Math.round((s.value / total) * 1000) / 10}%`, value: fmt(s.value), color: seriesColor(i) }],
                });
              }}
              onMouseLeave={() => onHover(null)}
            />
          );
        })}
      </svg>
      <Legend names={slices.map((s) => s.name)} />
    </>
  );
}

function Empty({ text }: { text: string }) {
  return <div className="grid flex-1 place-items-center px-4 text-center text-xs text-[var(--muted)]">{text}</div>;
}

/** Column, bar, line and area charts on one value axis. */
function AxisChart({
  data,
  type,
  width,
  height,
  onHover,
}: {
  data: ChartData;
  type: Exclude<ChartType, "pie">;
  width: number;
  height: number;
  onHover: (h: Hover | null) => void;
}) {
  const horizontal = type === "bar";
  const values = data.series.flatMap((s) => s.values).filter((v): v is number => v !== null);
  if (values.length === 0) return <Empty text="Select cells with numbers to chart." />;
  const ticks = niceTicks(Math.min(...values), Math.max(...values));
  const lo = ticks[0];
  const hi = ticks[ticks.length - 1];

  const pad = { top: 8, right: 12, bottom: horizontal ? 22 : 28, left: horizontal ? 72 : 44 };
  const plotW = Math.max(10, width - pad.left - pad.right);
  const plotH = Math.max(10, height - pad.top - pad.bottom);
  const n = data.categories.length;
  const along = horizontal ? plotH : plotW;
  const band = along / Math.max(1, n);
  const valueTo = (v: number) =>
    horizontal ? pad.left + ((v - lo) / (hi - lo)) * plotW : pad.top + plotH - ((v - lo) / (hi - lo)) * plotH;
  const zero = valueTo(Math.max(lo, Math.min(hi, 0)));
  const center = (i: number) => (horizontal ? pad.top : pad.left) + band * i + band / 2;

  const groups = data.series.length;
  const barW = Math.min(BAR_MAX, Math.max(2, (band * 0.7 - GAP * (groups - 1)) / groups));
  const groupW = barW * groups + GAP * (groups - 1);

  const labelEvery = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(along / (horizontal ? 16 : 48)))));

  const hoverCategory = (i: number, x: number, y: number) =>
    onHover({
      x,
      y,
      title: data.categories[i],
      rows: data.series.map((s, k) => ({ name: s.name, value: fmt(s.values[i]), color: seriesColor(k) })),
    });

  return (
    <>
      <svg
        width={width}
        height={height}
        role="img"
        aria-label={`${type} chart`}
        onMouseLeave={() => onHover(null)}
        onMouseMove={(e) => {
          if (type !== "line" && type !== "area") return;
          // Lines and areas: the nearest category under the pointer, with a crosshair.
          const box = e.currentTarget.getBoundingClientRect();
          const x = e.clientX - box.left;
          const i = Math.max(0, Math.min(n - 1, Math.floor((x - pad.left) / band)));
          hoverCategory(i, x, e.clientY - box.top);
          const line = e.currentTarget.querySelector<SVGLineElement>("[data-crosshair]");
          if (line) {
            line.setAttribute("x1", String(center(i)));
            line.setAttribute("x2", String(center(i)));
            line.style.display = "";
          }
        }}
      >
        {/* Gridlines and value labels */}
        {ticks.map((t) => {
          const p = valueTo(t);
          return horizontal ? (
            <g key={t}>
              <line x1={p} x2={p} y1={pad.top} y2={pad.top + plotH} stroke="var(--grid-line)" />
              <text x={p} y={height - 6} textAnchor="middle" fontSize={10} fill="var(--muted)">
                {compactNumber(t)}
              </text>
            </g>
          ) : (
            <g key={t}>
              <line x1={pad.left} x2={pad.left + plotW} y1={p} y2={p} stroke="var(--grid-line)" />
              <text x={pad.left - 6} y={p + 3} textAnchor="end" fontSize={10} fill="var(--muted)">
                {compactNumber(t)}
              </text>
            </g>
          );
        })}
        {/* Baseline */}
        {horizontal ? (
          <line x1={zero} x2={zero} y1={pad.top} y2={pad.top + plotH} stroke="var(--muted)" strokeOpacity={0.5} />
        ) : (
          <line x1={pad.left} x2={pad.left + plotW} y1={zero} y2={zero} stroke="var(--muted)" strokeOpacity={0.5} />
        )}
        {/* Category labels, thinned out when they'd collide */}
        {data.categories.map((c, i) =>
          i % labelEvery !== 0 ? null : horizontal ? (
            <text key={i} x={pad.left - 6} y={center(i) + 3} textAnchor="end" fontSize={10} fill="var(--muted)">
              {c.length > 12 ? c.slice(0, 11) + "…" : c}
            </text>
          ) : (
            <text key={i} x={center(i)} y={pad.top + plotH + 14} textAnchor="middle" fontSize={10} fill="var(--muted)">
              {c.length > 10 ? c.slice(0, 9) + "…" : c}
            </text>
          ),
        )}

        {type === "column" || type === "bar"
          ? data.categories.map((_, i) => (
              <g key={i}>
                {data.series.map((s, k) => {
                  const v = s.values[i];
                  if (v === null) return null;
                  const start = center(i) - groupW / 2 + k * (barW + GAP);
                  const end = valueTo(v);
                  const negative = v < 0;
                  const d = horizontal
                    ? barPath(Math.min(zero, end), start, Math.abs(end - zero), barW, true, negative)
                    : barPath(start, Math.min(zero, end), barW, Math.abs(end - zero), false, negative);
                  return <path key={k} d={d} fill={seriesColor(k)} />;
                })}
                {/* A hit area covering the whole category band, bigger than the bars */}
                <rect
                  x={horizontal ? pad.left : pad.left + band * i}
                  y={horizontal ? pad.top + band * i : pad.top}
                  width={horizontal ? plotW : band}
                  height={horizontal ? band : plotH}
                  fill="transparent"
                  onMouseMove={(e) => {
                    const box = (e.currentTarget.ownerSVGElement as SVGSVGElement).getBoundingClientRect();
                    hoverCategory(i, e.clientX - box.left, e.clientY - box.top);
                  }}
                />
              </g>
            ))
          : data.series.map((s, k) => {
              const pts = s.values.map((v, i) => (v === null ? null : ([center(i), valueTo(v)] as const)));
              const segments: (readonly [number, number])[][] = [[]];
              for (const p of pts) {
                if (p) segments[segments.length - 1].push(p);
                else if (segments[segments.length - 1].length) segments.push([]);
              }
              return (
                <g key={k}>
                  {segments.map((seg, j) =>
                    seg.length === 0 ? null : (
                      <g key={j}>
                        {type === "area" && seg.length > 1 && (
                          <path
                            d={`M${seg[0][0]},${zero} ${seg.map((p) => `L${p[0]},${p[1]}`).join(" ")} L${seg[seg.length - 1][0]},${zero} Z`}
                            fill={seriesColor(k)}
                            fillOpacity={0.1}
                          />
                        )}
                        <path
                          d={seg.map((p, idx) => `${idx ? "L" : "M"}${p[0]},${p[1]}`).join(" ")}
                          fill="none"
                          stroke={seriesColor(k)}
                          strokeWidth={2}
                          strokeLinejoin="round"
                          strokeLinecap="round"
                        />
                      </g>
                    ),
                  )}
                  {n <= 40 &&
                    pts.map((p, i) =>
                      p ? <circle key={i} cx={p[0]} cy={p[1]} r={4} fill={seriesColor(k)} stroke="var(--bg)" strokeWidth={2} /> : null,
                    )}
                </g>
              );
            })}
        {(type === "line" || type === "area") && (
          <line data-crosshair="" y1={pad.top} y2={pad.top + plotH} stroke="var(--muted)" strokeOpacity={0.5} style={{ display: "none" }} />
        )}
      </svg>
      <Legend names={data.series.map((s) => s.name)} />
    </>
  );
}

/** A chart's drawing area: picks the form, and owns the hover tooltip. */
export default function ChartView({ data, type, width, height }: { data: ChartData; type: ChartType; width: number; height: number }) {
  const [hover, setHover] = useState<Hover | null>(null);
  const legendRows = type === "pie" ? 1 : data.series.length > 1 ? 1 : 0;
  const plotHeight = Math.max(40, height - legendRows * 22);
  return (
    <div className="relative flex min-h-0 flex-1 flex-col">
      {data.series.length === 0 ? (
        <Empty text="Select cells with numbers to chart." />
      ) : type === "pie" ? (
        <PieChart data={data} width={width} height={plotHeight} onHover={setHover} />
      ) : (
        <AxisChart data={data} type={type} width={width} height={plotHeight} onHover={setHover} />
      )}
      <Tooltip hover={hover} />
    </div>
  );
}
