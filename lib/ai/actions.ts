import { HIGHLIGHT_STYLES } from "@/lib/conditional";
import {
  impliedDateKind,
  lettersToCol,
  mergeFormat,
  nameToRange,
  parseKey,
  type CellFormat,
  type CellRange,
  type ConditionalRule,
  type Workbook,
} from "@/lib/engine";
import { buildPivot, type Aggregate } from "@/lib/pivot";
import { sortPatches } from "@/lib/sheet-ops";
import { newSheetId } from "@/lib/storage";
import type { SheetAction } from "./schemas";

export interface ActionLimits {
  maxRows: number;
  maxCols: number;
}

export interface ApplyResult {
  /** The tab that should be showing afterwards. */
  sheet: string;
  applied: number;
  /** Steps that were skipped, and why, in plain words. */
  problems: string[];
}

/** Sizes of a single format or clear step; bigger ranges are refused rather than freezing the page. */
const MAX_CELLS = 50_000;

const STYLE_INDEX: Record<NonNullable<SheetAction["style"]>, number> = {
  red: 0,
  yellow: 1,
  green: 2,
  blue: 3,
  red_text: 4,
  bold: 5,
};

const HEX = /^#[0-9a-f]{6}$/i;

class StepError extends Error {}

/** Reads B2:D9 or B5; whole columns (C or A:C) and rows (5 or 2:4) are limited to the sheet's size. */
function parseRange(text: string | undefined, limits: ActionLimits): CellRange {
  const t = (text ?? "").trim().toUpperCase().replace(/\$/g, "");
  if (!t) throw new StepError("it didn't say which cells");
  const cols = /^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/.exec(t);
  if (cols) {
    const a = lettersToCol(cols[1]);
    const b = lettersToCol(cols[2] ?? cols[1]);
    return { r0: 0, r1: limits.maxRows - 1, c0: Math.min(a, b), c1: Math.max(a, b) };
  }
  const rows = /^(\d+)(?::(\d+))?$/.exec(t);
  if (rows) {
    const a = Number(rows[1]) - 1;
    const b = Number(rows[2] ?? rows[1]) - 1;
    return { r0: Math.min(a, b), r1: Math.max(a, b), c0: 0, c1: limits.maxCols - 1 };
  }
  const r = nameToRange(t.replace(/^.*!/, ""));
  if (!r) throw new StepError(`"${text}" isn't a cell range`);
  if (r.r1 >= limits.maxRows || r.c1 >= limits.maxCols) throw new StepError(`${text} is outside the sheet`);
  return r;
}

/** Limits whole-row or whole-column ranges to the used area, for formatting and clearing. */
function trimToUsed(wb: Workbook, sheet: string, r: CellRange, limits: ActionLimits): CellRange {
  const used = wb.usedBounds(sheet);
  const out = { ...r };
  if (r.r0 === 0 && r.r1 === limits.maxRows - 1) out.r1 = Math.max(r.r0, (used?.rows ?? 1) - 1);
  if (r.c0 === 0 && r.c1 === limits.maxCols - 1) out.c1 = Math.max(r.c0, (used?.cols ?? 1) - 1);
  if ((out.r1 - out.r0 + 1) * (out.c1 - out.c0 + 1) > MAX_CELLS) throw new StepError("that's too many cells at once");
  return out;
}

function formatPatch(a: SheetAction): CellFormat {
  const patch: CellFormat = {};
  if (a.bold !== undefined) patch.bold = a.bold;
  if (a.italic !== undefined) patch.italic = a.italic;
  if (a.underline !== undefined) patch.underline = a.underline;
  if (a.wrap !== undefined) patch.wrap = a.wrap;
  if (a.align) patch.align = a.align;
  if (a.fill !== undefined) {
    if (a.fill && !HEX.test(a.fill)) throw new StepError(`"${a.fill}" isn't a colour like #ffeeaa`);
    patch.fill = a.fill || undefined;
  }
  if (a.text_color !== undefined) {
    if (a.text_color && !HEX.test(a.text_color)) throw new StepError(`"${a.text_color}" isn't a colour like #224488`);
    patch.color = a.text_color || undefined;
  }
  if (a.number_format) {
    patch.numFmt = a.number_format === "general" ? undefined : a.number_format;
    patch.currency = a.number_format === "currency" ? (a.currency ?? "USD").toUpperCase() : undefined;
    patch.decimals = undefined;
  }
  if (a.decimals !== undefined) patch.decimals = a.decimals;
  // mergeFormat drops keys set to undefined, so "general" clears the number format.
  if (Object.keys(patch).length === 0) throw new StepError("it didn't say how to format the cells");
  return patch;
}

function findSheet(wb: Workbook, name: string | undefined): string {
  const want = (name ?? "").trim().toLowerCase();
  const found = wb.sheets.find((s) => s.name.toLowerCase() === want);
  if (!found) throw new StepError(`there's no tab called "${name}"`);
  return found.id;
}

function fieldColumn(wb: Workbook, sheet: string, source: CellRange, header: string | undefined): number {
  const want = (header ?? "").trim().toLowerCase();
  if (!want) throw new StepError("it didn't say which column to use");
  for (let c = source.c0; c <= source.c1; c++) {
    if (String(wb.getValue(sheet, source.r0, c) ?? "").trim().toLowerCase() === want) return c;
  }
  if (/^[a-z]{1,3}$/i.test(want)) return lettersToCol(want);
  throw new StepError(`there's no "${header}" column in the header row`);
}

const overlaps = (a: CellRange, b: CellRange) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0;

/** Carries out one step on `sheet`; returns the sheet later steps should work on. */
function applyOne(wb: Workbook, sheet: string, a: SheetAction, limits: ActionLimits): string {
  switch (a.type) {
    case "set_values": {
      const at = parseRange(a.range, limits);
      const rows = a.values ?? [];
      if (rows.length === 0) throw new StepError("it had no values to write");
      if (at.r0 + rows.length > limits.maxRows) throw new StepError("the values run past the last row");
      const updates = rows.flatMap((cells, i) =>
        cells.slice(0, limits.maxCols - at.c0).map((raw, j) => ({ row: at.r0 + i, col: at.c0 + j, raw: String(raw ?? "") })),
      );
      wb.setCells(sheet, updates);
      // Typed dates and date formulas get a date format, as when a person types them.
      for (const u of updates) {
        const kind = impliedDateKind(u.raw);
        const f = wb.getFormat(sheet, u.row, u.col);
        if (kind && !f?.numFmt) wb.setFormat(sheet, u.row, u.col, mergeFormat(f, { numFmt: kind }));
      }
      return sheet;
    }
    case "format":
    case "clear_format": {
      const r = trimToUsed(wb, sheet, parseRange(a.range, limits), limits);
      const patch = a.type === "format" ? formatPatch(a) : null;
      for (let row = r.r0; row <= r.r1; row++) {
        for (let col = r.c0; col <= r.c1; col++) {
          wb.setFormat(sheet, row, col, patch ? mergeFormat(wb.getFormat(sheet, row, col), patch) : null);
        }
      }
      return sheet;
    }
    case "clear": {
      const r = trimToUsed(wb, sheet, parseRange(a.range, limits), limits);
      const filled = [...wb.cellKeys(sheet)].map(parseKey).filter((p) => p.row >= r.r0 && p.row <= r.r1 && p.col >= r.c0 && p.col <= r.c1);
      wb.setCells(sheet, filled.map((p) => ({ ...p, raw: "" })));
      return sheet;
    }
    case "insert_rows":
    case "delete_rows": {
      const r = parseRange(a.range, limits);
      const count = a.count ?? r.r1 - r.r0 + 1;
      if (a.type === "insert_rows") wb.insert(sheet, "row", r.r0, Math.min(count, 1000));
      else wb.remove(sheet, "row", r.r0, Math.min(count, limits.maxRows - r.r0));
      return sheet;
    }
    case "insert_columns":
    case "delete_columns": {
      const r = parseRange(a.range, limits);
      const count = a.count ?? r.c1 - r.c0 + 1;
      if (a.type === "insert_columns") wb.insert(sheet, "col", r.c0, Math.min(count, 50));
      else wb.remove(sheet, "col", r.c0, Math.min(count, limits.maxCols - r.c0));
      return sheet;
    }
    case "sort": {
      const r = trimToUsed(wb, sheet, parseRange(a.range, limits), limits);
      if (wb.merges(sheet).some((m) => overlaps(m, r))) throw new StepError("the range has merged cells");
      const by = a.column ? lettersToCol(a.column.replace(/[^a-z]/gi, "")) : r.c0;
      if (by < r.c0 || by > r.c1) throw new StepError(`column ${a.column} isn't inside ${a.range}`);
      wb.apply(sortPatches(wb, sheet, r, by, a.ascending ?? true));
      return sheet;
    }
    case "add_sheet":
      return wb.addSheet(a.name, sheet);
    case "switch_sheet":
      return findSheet(wb, a.name);
    case "rename_sheet": {
      const name = (a.name ?? "").trim();
      const problem = wb.validateSheetName(name, sheet);
      if (problem) throw new StepError(problem);
      wb.renameSheet(sheet, name);
      return sheet;
    }
    case "chart": {
      const r = parseRange(a.range, limits);
      const charts = wb.charts(sheet);
      wb.setCharts(sheet, [
        ...charts,
        {
          id: newSheetId(),
          type: a.chart_type ?? "column",
          range: r,
          title: a.title ?? "Chart",
          // To the right of the data, stacked below any charts already there.
          x: (r.c1 + 1) * 100 + 16,
          y: r.r0 * 24 + charts.length * 24,
          width: 480,
          height: 300,
        },
      ]);
      return sheet;
    }
    case "highlight": {
      const r = parseRange(a.range, limits);
      if (!a.condition) throw new StepError("it didn't say which cells to highlight");
      const rule: ConditionalRule = {
        id: newSheetId(),
        range: r,
        kind: a.condition,
        values: a.compare_values,
        style: HIGHLIGHT_STYLES[STYLE_INDEX[a.style ?? "red"]].style,
      };
      wb.setConditionalRules(sheet, [...wb.conditionalRules(sheet), rule]);
      return sheet;
    }
    case "color_scale":
    case "data_bars": {
      const r = parseRange(a.range, limits);
      const rule: ConditionalRule =
        a.type === "color_scale"
          ? { id: newSheetId(), range: r, kind: "scale", scale: a.scale ?? "blue" }
          : { id: newSheetId(), range: r, kind: "bar" };
      wb.setConditionalRules(sheet, [...wb.conditionalRules(sheet), rule]);
      return sheet;
    }
    case "dropdown": {
      const r = parseRange(a.range, limits);
      const items = (a.items ?? []).map((x) => x.trim()).filter(Boolean);
      if (items.length === 0) throw new StepError("it had no choices for the list");
      wb.setValidations(sheet, [...wb.validations(sheet), { id: newSheetId(), range: r, items: [...new Set(items)], strict: true }]);
      return sheet;
    }
    case "freeze":
      wb.setFreeze(sheet, { rows: a.rows ?? 0, cols: a.columns ?? 0 });
      return sheet;
    case "merge": {
      const r = parseRange(a.range, limits);
      if (r.r0 === r.r1 && r.c0 === r.c1) throw new StepError("merging needs two or more cells");
      const keep = wb.getRaw(sheet, r.r0, r.c0);
      const others = [...wb.cellKeys(sheet)]
        .map(parseKey)
        .filter((p) => p.row >= r.r0 && p.row <= r.r1 && p.col >= r.c0 && p.col <= r.c1 && (p.row !== r.r0 || p.col !== r.c0));
      if (others.length && !keep) throw new StepError("the top-left cell is empty but other cells have values");
      wb.setCells(sheet, others.map((p) => ({ ...p, raw: "" })));
      wb.setMerges(sheet, [...wb.merges(sheet).filter((m) => !overlaps(m, r)), r]);
      return sheet;
    }
    case "unmerge": {
      const r = parseRange(a.range, limits);
      wb.setMerges(sheet, wb.merges(sheet).filter((m) => !overlaps(m, r)));
      return sheet;
    }
    case "column_width": {
      const r = parseRange(a.range, limits);
      if (!a.width) throw new StepError("it didn't give a width");
      for (let c = r.c0; c <= r.c1; c++) wb.setColWidth(sheet, c, a.width);
      return sheet;
    }
    case "comment": {
      const r = parseRange(a.range, limits);
      wb.setComment(sheet, r.r0, r.c0, a.text ?? null);
      return sheet;
    }
    case "pivot": {
      const source = parseRange(a.range, limits);
      const spec = {
        source,
        rowField: fieldColumn(wb, sheet, source, a.row_field),
        colField: a.column_field ? fieldColumn(wb, sheet, source, a.column_field) : null,
        valueField: fieldColumn(wb, sheet, source, a.value_field),
        aggregate: (a.aggregate ?? "sum") as Aggregate,
      };
      const name = a.name?.trim() || `Pivot of ${wb.sheetName(sheet)}`;
      return wb.addSheetWith(buildPivot(wb, sheet, spec, name), sheet);
    }
  }
}

/**
 * Carries out a plan's steps in order on the workbook. Each step that can't
 * be done is skipped with a reason, so one bad step doesn't spoil the rest.
 * Run it inside an undo snapshot so the whole plan undoes in one go.
 */
export function applyActions(wb: Workbook, startSheet: string, actions: SheetAction[], limits: ActionLimits): ApplyResult {
  let sheet = startSheet;
  let applied = 0;
  const problems: string[] = [];
  actions.forEach((a, i) => {
    try {
      sheet = applyOne(wb, sheet, a, limits);
      applied++;
    } catch (e) {
      const why = e instanceof Error ? e.message : String(e);
      problems.push(`Step ${i + 1} (${a.description}) was skipped: ${why}.`);
    }
  });
  return { sheet, applied, problems };
}
