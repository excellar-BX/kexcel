import { describe, expect, it } from "vitest";
import { applyActions } from "./ai/actions";
import type { SheetAction } from "./ai/schemas";
import { formatValue, Workbook } from "./engine";
import { buildPivot } from "./pivot";

const LIMITS = { maxRows: 10_000, maxCols: 52 };

const sales = () =>
  Workbook.fromJSON({
    sheets: [
      {
        id: "s1",
        name: "Sales",
        cells: {
          A1: "Region", B1: "Month", C1: "Amount",
          A2: "Lagos", B2: "Jan", C2: "1200",
          A3: "Abuja", B3: "Jan", C3: "800",
          A4: "Lagos", B4: "Feb", C4: "1500",
          A5: "Kano", B5: "Feb", C5: "650",
          A6: "Total", C6: "=SUM(C2:C5)",
        },
      },
    ],
  });

const shown = (wb: Workbook, sheet: string, name: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(name)!;
  const col = m[1].split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const row = Number(m[2]) - 1;
  return formatValue(wb.getValue(sheet, row, col), wb.getFormat(sheet, row, col));
};

describe("*IFS functions", () => {
  it("sum, count and average with several conditions", () => {
    const wb = sales();
    wb.setCells("s1", [
      { row: 9, col: 0, raw: '=SUMIFS(C2:C5,A2:A5,"Lagos",B2:B5,"Feb")' },
      { row: 9, col: 1, raw: '=COUNTIFS(A2:A5,"Lagos")' },
      { row: 9, col: 2, raw: '=AVERAGEIFS(C2:C5,B2:B5,"Jan")' },
      { row: 9, col: 3, raw: '=MAXIFS(C2:C5,A2:A5,"<>Lagos")' },
      { row: 9, col: 4, raw: '=MINIFS(C2:C5,A2:A5,"Lagos")' },
      { row: 9, col: 5, raw: "=SUMIFS(C2:C5,A2:A3,1)" },
    ]);
    expect(["A10", "B10", "C10", "D10", "E10", "F10"].map((n) => shown(wb, "s1", n))).toEqual([
      "1500", "2", "1000", "800", "1200", "#VALUE!",
    ]);
  });
});

describe("pivot tables", () => {
  it("groups rows with live formulas, skipping total rows", () => {
    const wb = sales();
    const id = wb.addSheetWith(
      buildPivot(wb, "s1", { source: { r0: 0, r1: 5, c0: 0, c1: 2 }, rowField: 0, colField: null, valueField: 2, aggregate: "sum" }, "Pivot"),
    );
    expect(["A1", "B1", "A2", "B2", "A3", "B3", "A4", "B4", "A5", "B5"].map((n) => shown(wb, id, n))).toEqual([
      "Region", "Sum of Amount", "Abuja", "800", "Kano", "650", "Lagos", "2700", "Total", "4150",
    ]);
    // The totals follow the data.
    wb.setCell("s1", 1, 2, "2000");
    expect(shown(wb, id, "B4")).toBe("3500");
    expect(wb.getFreeze(id)).toEqual({ rows: 1, cols: 1 });
  });

  it("spreads a second field across columns", () => {
    const wb = sales();
    const id = wb.addSheetWith(
      buildPivot(wb, "s1", { source: { r0: 0, r1: 4, c0: 0, c1: 2 }, rowField: 0, colField: 1, valueField: 2, aggregate: "sum" }, "Pivot"),
    );
    expect(["B1", "C1", "D1"].map((n) => shown(wb, id, n))).toEqual(["Feb", "Jan", "Total"]);
    expect(["A4", "B4", "C4", "D4"].map((n) => shown(wb, id, n))).toEqual(["Lagos", "1500", "1200", "2700"]);
    expect(["B5", "C5", "D5"].map((n) => shown(wb, id, n))).toEqual(["2150", "2000", "4150"]);
  });
});

describe("AI actions", () => {
  const step = (a: Partial<SheetAction> & Pick<SheetAction, "type">): SheetAction => ({ description: a.type, ...a });

  it("carry out a plan in order, across a new tab", () => {
    const wb = sales();
    const r = applyActions(
      wb,
      "s1",
      [
        step({ type: "format", range: "A1:C1", bold: true, fill: "#e3eefc" }),
        step({ type: "format", range: "C2:C6", number_format: "currency", currency: "NGN", decimals: 0 }),
        step({ type: "highlight", range: "C2:C5", condition: "greater", compare_values: ["1000"], style: "green" }),
        step({ type: "sort", range: "A2:C5", column: "C", ascending: false }),
        step({ type: "freeze", rows: 1 }),
        step({ type: "comment", range: "C6", text: "Sum of all regions" }),
        step({ type: "add_sheet", name: "Notes" }),
        step({ type: "set_values", range: "A1", values: [["Checked", "=TODAY()"]] }),
      ],
      LIMITS,
    );
    expect(r.problems).toEqual([]);
    expect(r.applied).toBe(8);
    expect(wb.sheetName(r.sheet)).toBe("Notes");
    expect(wb.getFormat("s1", 0, 0)).toEqual({ bold: true, fill: "#e3eefc" });
    expect(shown(wb, "s1", "C2")).toBe("₦1,500");
    expect(shown(wb, "s1", "A5")).toBe("Kano");
    expect(wb.conditionalRules("s1")).toHaveLength(1);
    expect(wb.getFreeze("s1")).toEqual({ rows: 1, cols: 0 });
    expect(wb.getComment("s1", 5, 2)).toBe("Sum of all regions");
    expect(wb.getFormat(r.sheet, 0, 1)).toEqual({ numFmt: "date" });
  });

  it("skip steps that can't be done, with a reason, and keep going", () => {
    const wb = sales();
    const r = applyActions(
      wb,
      "s1",
      [
        step({ type: "format", range: "nonsense", bold: true, description: "Bold the header" }),
        step({ type: "switch_sheet", name: "Nowhere", description: "Go to Nowhere" }),
        step({ type: "format", range: "A1", fill: "blue", description: "Colour A1" }),
        step({ type: "set_values", range: "D1", values: [["Kept"]] }),
      ],
      LIMITS,
    );
    expect(r.applied).toBe(1);
    expect(r.problems).toEqual([
      'Step 1 (Bold the header) was skipped: "nonsense" isn\'t a cell range.',
      'Step 2 (Go to Nowhere) was skipped: there\'s no tab called "Nowhere".',
      'Step 3 (Colour A1) was skipped: "blue" isn\'t a colour like #ffeeaa.',
    ]);
    expect(wb.getRaw("s1", 0, 3)).toBe("Kept");
  });

  it("make pivot tables by header name", () => {
    const wb = sales();
    const r = applyActions(wb, "s1", [step({ type: "pivot", range: "A1:C5", row_field: "region", value_field: "Amount", aggregate: "count" })], LIMITS);
    expect(r.problems).toEqual([]);
    expect(wb.sheetName(r.sheet)).toBe("Pivot of Sales");
    expect(shown(wb, r.sheet, "B4")).toBe("2");
  });
});

describe("comments", () => {
  it("are saved and move with their cells", () => {
    const wb = sales();
    wb.setComment("s1", 2, 0, "Check with Ada");
    wb.insert("s1", "row", 0, 1);
    expect(wb.getComment("s1", 3, 0)).toBe("Check with Ada");
    const copy = Workbook.fromJSON(wb.toJSON());
    expect(copy.getComment("s1", 3, 0)).toBe("Check with Ada");
    copy.setComment("s1", 3, 0, "  ");
    expect(copy.getComment("s1", 3, 0)).toBeUndefined();
  });
});

describe("pivot averages", () => {
  it("average the source rows, not the group averages", () => {
    const wb = sales();
    const id = wb.addSheetWith(
      buildPivot(wb, "s1", { source: { r0: 0, r1: 5, c0: 0, c1: 2 }, rowField: 0, colField: null, valueField: 2, aggregate: "average" }, "P"),
    );
    // Lagos averages 1350 and the others 800 and 650; all four rows average 1037.5.
    expect(shown(wb, id, "B4")).toBe("1350");
    expect(shown(wb, id, "B5")).toBe("1037.5");
  });
});
