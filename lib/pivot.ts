import {
  addressToName,
  colToLetters,
  formatValue,
  isError,
  isFormula,
  sheetPrefix,
  type CellFormat,
  type CellRange,
  type Scalar,
  type SheetJSON,
  type Workbook,
} from "@/lib/engine";

export type Aggregate = "sum" | "count" | "average" | "min" | "max";

export const AGGREGATES: { value: Aggregate; label: string }[] = [
  { value: "sum", label: "Sum" },
  { value: "count", label: "Count" },
  { value: "average", label: "Average" },
  { value: "min", label: "Smallest" },
  { value: "max", label: "Largest" },
];

export interface PivotSpec {
  /** The data, with a header row. */
  source: CellRange;
  /** Absolute column indexes within the source sheet. */
  rowField: number;
  colField: number | null;
  valueField: number;
  aggregate: Aggregate;
}

const MAX_ROWS = 1000;
const MAX_COLS = 50;

/** The header names of a range's first row, by absolute column. */
export function pivotFields(wb: Workbook, sheet: string, source: CellRange): { col: number; name: string }[] {
  const out: { col: number; name: string }[] = [];
  for (let c = source.c0; c <= source.c1; c++) {
    const v = wb.getValue(sheet, source.r0, c);
    out.push({ col: c, name: v === null || v === "" ? `Column ${colToLetters(c)}` : formatValue(v) });
  }
  return out;
}

/** Distinct non-empty values of a column, sorted: numbers first, then text, ignoring case. */
function distinct(wb: Workbook, sheet: string, col: number, r0: number, r1: number) {
  const seen = new Map<string, { value: Scalar; raw: string }>();
  for (let r = r0; r <= r1; r++) {
    const v = wb.getValue(sheet, r, col);
    if (v === null || v === "" || isError(v)) continue;
    // A total row inside the range isn't a group of its own.
    if (typeof v === "string" && /^(grand )?totals?$/i.test(v.trim())) continue;
    const key = typeof v === "string" ? v.trim().toLowerCase() : formatValue(v);
    if (seen.has(key)) continue;
    const raw = wb.getRaw(sheet, r, col);
    // Keys are copied as typed; formula results are copied as their value.
    seen.set(key, { value: v, raw: isFormula(raw) ? (typeof v === "string" ? `'${v}` : formatValue(v)) : raw });
  }
  return [...seen.values()].sort((a, b) => {
    const na = typeof a.value === "number";
    const nb = typeof b.value === "number";
    if (na !== nb) return na ? -1 : 1;
    if (na && nb) return (a.value as number) - (b.value as number);
    return formatValue(a.value).localeCompare(formatValue(b.value), undefined, { sensitivity: "base", numeric: true });
  });
}

const absolute = (prefix: string, col: number, r0: number, r1: number) =>
  `${prefix}$${colToLetters(col)}$${r0 + 1}:$${colToLetters(col)}$${r1 + 1}`;

/**
 * Builds a pivot table as a new sheet. Every value cell is a live *IFS
 * formula over the source, so totals follow the data; the list of row and
 * column labels is taken when the table is made.
 */
export function buildPivot(wb: Workbook, sheet: string, spec: PivotSpec, name: string): SheetJSON {
  const { source, rowField, colField, valueField, aggregate } = spec;
  if (source.r1 <= source.r0) throw new Error("A pivot table needs a header row and at least one row of data.");
  const first = source.r0 + 1;
  const last = source.r1;
  const prefix = sheetPrefix(wb.sheetName(sheet));
  const fields = pivotFields(wb, sheet, source);
  const fieldName = (col: number) => fields.find((f) => f.col === col)?.name ?? colToLetters(col);

  const rows = distinct(wb, sheet, rowField, first, last);
  const cols = colField === null ? [] : distinct(wb, sheet, colField, first, last);
  if (rows.length === 0) throw new Error(`"${fieldName(rowField)}" has no values to group by.`);
  if (rows.length > MAX_ROWS) throw new Error(`"${fieldName(rowField)}" has more than ${MAX_ROWS} different values.`);
  if (cols.length > MAX_COLS) throw new Error(`"${fieldName(colField!)}" has more than ${MAX_COLS} different values.`);

  const valueRange = absolute(prefix, valueField, first, last);
  const rowRange = absolute(prefix, rowField, first, last);
  const colRange = colField === null ? "" : absolute(prefix, colField, first, last);

  /** The aggregate over the value column, for rows matching the given (range, criterion) pairs. */
  const formula = (criteria: [string, string][]) => {
    const pairs = criteria.map(([r, c]) => `${r},${c}`).join(",");
    if (criteria.length === 0) {
      return aggregate === "count" ? `=COUNTA(${rowRange})` : `=${aggregate === "average" ? "AVERAGE" : aggregate.toUpperCase()}(${valueRange})`;
    }
    switch (aggregate) {
      case "sum":
        return `=SUMIFS(${valueRange},${pairs})`;
      case "count":
        return `=COUNTIFS(${pairs})`;
      case "average":
        return `=IFERROR(AVERAGEIFS(${valueRange},${pairs}),"")`;
      case "min":
        return `=MINIFS(${valueRange},${pairs})`;
      case "max":
        return `=MAXIFS(${valueRange},${pairs})`;
    }
  };

  const cells: Record<string, string> = {};
  const formats: Record<string, CellFormat> = {};
  const header: CellFormat = { bold: true, borders: { bottom: true } };
  const total: CellFormat = { bold: true, borders: { top: true } };
  // Values keep the number format of the source column (currency and so on), except counts.
  const valueFormat = aggregate === "count" ? undefined : wb.getFormat(sheet, first, valueField);
  const numberLook: CellFormat = valueFormat?.numFmt ? { numFmt: valueFormat.numFmt, currency: valueFormat.currency, decimals: valueFormat.decimals } : {};
  const set = (row: number, col: number, raw: string, format?: CellFormat) => {
    const key = addressToName({ row, col });
    cells[key] = raw;
    if (format && Object.keys(format).length) formats[key] = format;
  };

  const label = `${AGGREGATES.find((a) => a.value === aggregate)!.label} of ${fieldName(valueField)}`;
  set(0, 0, `'${fieldName(rowField)}`, header);
  const valueCols = colField === null ? [{ raw: `'${label}` }] : cols;
  valueCols.forEach((c, j) => set(0, j + 1, c.raw, header));
  if (colField !== null) set(0, cols.length + 1, "'Total", header);

  rows.forEach((r, i) => {
    const row = i + 1;
    set(row, 0, r.raw);
    const rowKey = `$A${row + 1}`;
    if (colField === null) {
      set(row, 1, formula([[rowRange, rowKey]]), numberLook);
    } else {
      cols.forEach((_, j) => {
        const colKey = `${colToLetters(j + 1)}$1`;
        set(row, j + 1, formula([[rowRange, rowKey], [colRange, colKey]]), numberLook);
      });
      set(row, cols.length + 1, formula([[rowRange, rowKey]]), { ...numberLook, bold: true });
    }
  });

  // Totals add up the table's own rows, so they match it even when the source has total
  // rows or blank labels of its own. An average of averages would be wrong, so averages
  // go back to the source, leaving out blank and total rows.
  const totalRow = rows.length + 1;
  const totalOf = (col: number, extra: [string, string][]) => {
    const letter = colToLetters(col);
    const own = `${letter}2:${letter}${totalRow}`;
    if (aggregate === "sum" || aggregate === "count") return `=SUM(${own})`;
    if (aggregate === "min" || aggregate === "max") return `=${aggregate.toUpperCase()}(${own})`;
    const skip: [string, string][] = [
      [rowRange, '"<>"'],
      [rowRange, '"<>Total"'],
      [rowRange, '"<>Grand total"'],
    ];
    return formula([...skip, ...extra]);
  };
  set(totalRow, 0, "'Total", total);
  if (colField === null) {
    set(totalRow, 1, totalOf(1, []), { ...numberLook, ...total });
  } else {
    cols.forEach((_, j) =>
      set(totalRow, j + 1, totalOf(j + 1, [[colRange, `${colToLetters(j + 1)}$1`]]), { ...numberLook, ...total }),
    );
    set(totalRow, cols.length + 1, totalOf(cols.length + 1, []), { ...numberLook, ...total });
  }

  return {
    id: "",
    name,
    cells,
    formats,
    colWidths: { 0: 180 },
    freeze: { rows: 1, cols: 1 },
  };
}
