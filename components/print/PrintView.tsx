"use client";

import { colToLetters, formatValue, isError, type Workbook } from "@/lib/engine";
import type { Range } from "@/lib/sheet-ops";
import { textStyle } from "@/components/grid/Grid";
import { GridLayout } from "@/components/grid/layout";

export type Paper = "A4" | "Letter";
export type Orientation = "portrait" | "landscape";

export interface PrintOptions {
  paper: Paper;
  orientation: Orientation;
  fitWidth: boolean;
  gridlines: boolean;
  headers: boolean;
}

export interface PrintJob {
  sheets: { id: string; range: Range }[];
  options: PrintOptions;
}

const MARGIN_MM = 12;
const PAPER_MM: Record<Paper, [number, number]> = {
  A4: [210, 297],
  Letter: [215.9, 279.4],
};
const mmToPx = (mm: number) => (mm * 96) / 25.4;

/** Width of the printable area of a page in CSS pixels. */
export function printableWidth({ paper, orientation }: PrintOptions) {
  const [w, h] = PAPER_MM[paper];
  return mmToPx((orientation === "portrait" ? w : h) - 2 * MARGIN_MM);
}

function SheetTable({
  wb,
  sheet,
  range,
  options,
}: {
  wb: Workbook;
  sheet: string;
  range: Range;
  options: PrintOptions;
}) {
  const layout = GridLayout.forSheet(wb, sheet, 1);
  const cols = Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i);
  const rows = Array.from({ length: range.r1 - range.r0 + 1 }, (_, i) => range.r0 + i);
  const headerWidth = options.headers ? 40 : 0;
  const tableWidth = headerWidth + cols.reduce((sum, c) => sum + layout.colWidth(c), 0);
  const scale = options.fitWidth ? Math.min(1, printableWidth(options) / tableWidth) : 1;
  const grid = options.gridlines ? "1px solid #d0d0d0" : undefined;
  const headerCell = "border border-[#c0c0c0] bg-[#f3f3f3] text-center text-[10px] text-[#555]";

  return (
    <table
      className="border-collapse"
      style={{
        tableLayout: "fixed",
        width: tableWidth,
        zoom: scale,
        color: "#000",
      }}
    >
      <colgroup>
        {options.headers && <col style={{ width: headerWidth }} />}
        {cols.map((c) => (
          <col key={c} style={{ width: layout.colWidth(c) }} />
        ))}
      </colgroup>
      {options.headers && (
        <thead>
          <tr>
            <th className={headerCell} />
            {cols.map((c) => (
              <th key={c} className={headerCell}>
                {colToLetters(c)}
              </th>
            ))}
          </tr>
        </thead>
      )}
      <tbody>
        {rows.map((r) => (
          <tr key={r} style={{ height: layout.rowHeight(r) }}>
            {options.headers && <th className={headerCell}>{r + 1}</th>}
            {cols.map((c) => {
              // A merged range prints as one cell spanning its rows and columns (within the printed area).
              const merge = wb.mergeAt(sheet, r, c);
              const topLeft = merge && { row: Math.max(merge.r0, range.r0), col: Math.max(merge.c0, range.c0) };
              if (topLeft && (topLeft.row !== r || topLeft.col !== c)) return null;
              const value = merge ? wb.getValue(sheet, merge.r0, merge.c0) : wb.getValue(sheet, r, c);
              const f = merge ? wb.getFormat(sheet, merge.r0, merge.c0) : wb.getFormat(sheet, r, c);
              const b = f?.borders;
              const line = "1px solid #000";
              const align =
                f?.align ?? (typeof value === "number" ? "right" : typeof value === "string" ? "left" : "center");
              return (
                <td
                  key={c}
                  rowSpan={merge ? Math.min(merge.r1, range.r1) - r + 1 : undefined}
                  colSpan={merge ? Math.min(merge.c1, range.c1) - c + 1 : undefined}
                  className={`overflow-hidden px-1 ${f?.wrap ? "whitespace-pre-wrap break-words" : "whitespace-nowrap"}`}
                  style={{
                    ...textStyle(f, 1),
                    color: isError(value) ? "#c62828" : (f?.color ?? "#000"),
                    textAlign: align,
                    background: f?.fill,
                    borderTop: b?.top ? line : grid,
                    borderBottom: b?.bottom ? line : grid,
                    borderLeft: b?.left ? line : grid,
                    borderRight: b?.right ? line : grid,
                  }}
                >
                  {formatValue(value, f)}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/** What goes on paper: only shown when printing, while the app itself is hidden. */
export default function PrintView({ wb, job, title }: { wb: Workbook; job: PrintJob; title: string }) {
  const { paper, orientation } = job.options;
  const many = job.sheets.length > 1;
  return (
    <div
      className="hidden bg-white text-black print:block"
      style={{ printColorAdjust: "exact", WebkitPrintColorAdjust: "exact" }}
    >
      <style>{`@page { size: ${paper} ${orientation}; margin: ${MARGIN_MM}mm; }`}</style>
      {job.sheets.map(({ id, range }, i) => (
        <section key={id} style={{ breakBefore: i > 0 ? "page" : undefined }}>
          {many && (
            <h2 className="mb-2 text-sm font-semibold">
              {title} — {wb.sheetName(id)}
            </h2>
          )}
          <SheetTable wb={wb} sheet={id} range={range} options={job.options} />
        </section>
      ))}
    </div>
  );
}
