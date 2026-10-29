import { formatValue, isError, type CellRange, type Workbook } from "@/lib/engine";

/** The eight series colours, in their fixed order (see the --series-N tokens). */
export const MAX_SERIES = 8;
export const MAX_CATEGORIES = 200;
export const MAX_PIE_SLICES = 8;

export interface ChartSeries {
  name: string;
  values: (number | null)[];
}

export interface ChartData {
  categories: string[];
  series: ChartSeries[];
  /** Set when the range had more series or categories than a chart shows. */
  truncated: boolean;
}

const isText = (v: unknown) => typeof v === "string" && v.trim() !== "";

/**
 * Reads a range the way spreadsheets usually lay out chart data: each column
 * is a series, each row a category. A first row with text is taken as series
 * names, and a first column with text as category labels.
 */
export function chartData(wb: Workbook, sheet: string, range: CellRange): ChartData {
  const value = (r: number, c: number) => wb.getValue(sheet, r, c);
  const width = range.c1 - range.c0 + 1;

  const headerRow = Array.from({ length: width }, (_, i) => value(range.r0, range.c0 + i)).some(isText);
  const firstData = headerRow ? range.r0 + 1 : range.r0;
  let labelCol = false;
  if (width > 1) {
    for (let r = firstData; r <= range.r1; r++) {
      const v = value(r, range.c0);
      if (isText(v)) labelCol = true;
    }
  }

  const seriesCols: number[] = [];
  for (let c = labelCol ? range.c0 + 1 : range.c0; c <= range.c1; c++) seriesCols.push(c);
  const rows: number[] = [];
  for (let r = firstData; r <= range.r1; r++) rows.push(r);

  const truncated = seriesCols.length > MAX_SERIES || rows.length > MAX_CATEGORIES;
  const shownRows = rows.slice(0, MAX_CATEGORIES);
  const categories = shownRows.map((r, i) =>
    labelCol ? formatValue(value(r, range.c0), wb.getFormat(sheet, r, range.c0)) : String(i + 1),
  );
  const series = seriesCols.slice(0, MAX_SERIES).map((c, i) => {
    const head = headerRow ? value(range.r0, c) : null;
    return {
      name: head !== null && !isError(head) && head !== "" ? formatValue(head) : `Series ${i + 1}`,
      values: shownRows.map((r) => {
        const v = value(r, c);
        return typeof v === "number" && Number.isFinite(v) ? v : null;
      }),
    };
  });
  return { categories, series, truncated };
}

/** Clean axis ticks from 0 (or the minimum, when negative) to just past the maximum. */
export function niceTicks(min: number, max: number, target = 5): number[] {
  const lo = Math.min(0, min);
  const hi = Math.max(0, max);
  if (lo === hi) return [0, 1];
  const raw = (hi - lo) / target;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  const ticks: number[] = [];
  for (let t = Math.floor(lo / step) * step; t <= hi + step * 0.001; t += step) ticks.push(Number(t.toPrecision(12)));
  if (ticks[ticks.length - 1] < hi) ticks.push(Number((ticks[ticks.length - 1] + step).toPrecision(12)));
  return ticks;
}

/** 1,284 / 12.9K / 4.2M style labels for axes. */
export function compactNumber(n: number): string {
  const abs = Math.abs(n);
  if (abs >= 1e9) return `${Number((n / 1e9).toPrecision(3))}B`;
  if (abs >= 1e6) return `${Number((n / 1e6).toPrecision(3))}M`;
  if (abs >= 1e4) return `${Number((n / 1e3).toPrecision(3))}K`;
  return Number(n.toPrecision(6)).toLocaleString();
}
