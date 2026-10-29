import { describe, expect, it } from "vitest";
import { chartData, compactNumber, niceTicks } from "./charts";
import { describeRule, lookAt, prepareRules } from "./conditional";
import { Workbook, type ConditionalRule } from "./engine";
import { isAllowed, isInvalid, listItems, sourceProblem } from "./validation";

const range = (r0: number, r1: number, c0: number, c1: number) => ({ r0, r1, c0, c1 });

function book(cells: Record<string, string>, rules: ConditionalRule[] = []) {
  const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "Data", cells }] });
  wb.setConditionalRules("s1", rules);
  const prepared = prepareRules(wb, "s1");
  const look = (row: number, col: number) => lookAt(prepared, wb.getValue("s1", row, col), row, col);
  return { wb, look };
}

describe("conditional formatting", () => {
  const red = { fill: "#fde7e7", color: "#a52a24" };

  it("highlights values by comparison, text and emptiness", () => {
    const { look } = book({ A1: "5", A2: "15", A3: "apple pie", A5: "10" }, [
      { id: "1", range: range(0, 4, 0, 0), kind: "greater", values: ["10"], style: red },
      { id: "2", range: range(0, 4, 0, 0), kind: "contains", values: ["PIE"], style: { bold: true } },
      { id: "3", range: range(0, 4, 0, 0), kind: "blank", style: { fill: "#eee" } },
    ]);
    expect(look(0, 0)).toBeNull();
    expect(look(1, 0)).toEqual(red);
    expect(look(2, 0)).toEqual({ bold: true });
    expect(look(3, 0)).toEqual({ fill: "#eee" });
    expect(look(4, 0)).toBeNull(); // 10 is not greater than 10
  });

  it("lets earlier rules win and handles between and duplicates", () => {
    const { look } = book({ A1: "3", A2: "7", A3: "x", A4: "X" }, [
      { id: "1", range: range(0, 3, 0, 0), kind: "between", values: ["1", "5"], style: { fill: "first" } },
      { id: "2", range: range(0, 3, 0, 0), kind: "notBlank", style: { fill: "second", bold: true } },
      { id: "3", range: range(0, 3, 0, 0), kind: "duplicate", style: { color: "dup" } },
    ]);
    expect(look(0, 0)).toEqual({ fill: "first", bold: true });
    expect(look(1, 0)).toEqual({ fill: "second", bold: true });
    expect(look(2, 0)).toEqual({ fill: "second", bold: true, color: "dup" });
  });

  it("works out colour scales and data bars from the range's own values", () => {
    const { look } = book({ A1: "0", A2: "50", A3: "100", B1: "-10", B2: "0", B3: "30" }, [
      { id: "1", range: range(0, 2, 0, 0), kind: "bar" },
      { id: "2", range: range(0, 2, 1, 1), kind: "scale", scale: "diverging" },
    ]);
    expect(look(1, 0)?.bar).toBe(0.5);
    expect(look(2, 0)?.bar).toBe(1);
    expect(look(0, 1)?.fill).toContain("--series-8"); // below zero
    expect(look(1, 1)).toBeNull(); // the neutral middle
    expect(look(2, 1)?.fill).toContain("--series-1");
  });

  it("describes rules in words", () => {
    expect(describeRule({ id: "x", range: range(0, 0, 0, 0), kind: "between", values: ["1", "9"] })).toBe("Between 1 and 9");
  });
});

describe("dropdown lists", () => {
  it("reads typed items or a source range, and checks values", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        { id: "s1", name: "Orders", cells: { A1: "Paid", A2: "Maybe" } },
        { id: "s2", name: "My lists", cells: { A1: "Paid", A2: "Unpaid", A3: "paid", A4: "" } },
      ],
    });
    wb.setValidations("s1", [{ id: "v1", range: range(0, 9, 0, 0), source: "'My lists'!A1:A5", strict: false }]);
    const rule = wb.validationAt("s1", 1, 0)!;
    expect(listItems(wb, "s1", rule)).toEqual(["Paid", "Unpaid"]);
    expect(isAllowed(["Paid", "Unpaid"], " paid ")).toBe(true);
    expect(isAllowed(["Paid"], "")).toBe(true);
    expect(isInvalid(wb, "s1", 0, 0)).toBe(false);
    expect(isInvalid(wb, "s1", 1, 0)).toBe(true);
    expect(sourceProblem(wb, "Nope!A1:A3")).toMatch(/no tab/);
    expect(sourceProblem(wb, "A1:A3")).toBeNull();
  });

  it("moves with inserted rows and is saved", () => {
    const wb = Workbook.withSheet();
    wb.setValidations("s1", [{ id: "v1", range: range(2, 5, 1, 1), items: ["Yes", "No"], strict: true }]);
    wb.insert("s1", "row", 0, 2);
    expect(wb.validations("s1")[0].range).toEqual(range(4, 7, 1, 1));
    expect(Workbook.fromJSON(wb.toJSON()).validations("s1")).toEqual(wb.validations("s1"));
  });
});

describe("chart data", () => {
  it("takes series names from the first row and labels from the first column", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        { id: "s1", name: "S", cells: { A1: "Month", B1: "Lagos", C1: "Abuja", A2: "Jan", B2: "10", C2: "4", A3: "Feb", B3: "12", C3: "x" } },
      ],
    });
    expect(chartData(wb, "s1", range(0, 2, 0, 2))).toEqual({
      categories: ["Jan", "Feb"],
      series: [
        { name: "Lagos", values: [10, 12] },
        { name: "Abuja", values: [4, null] },
      ],
      truncated: false,
    });
  });

  it("numbers categories when there are no labels", () => {
    const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "S", cells: { A1: "3", A2: "5" } }] });
    expect(chartData(wb, "s1", range(0, 1, 0, 0))).toEqual({
      categories: ["1", "2"],
      series: [{ name: "Series 1", values: [3, 5] }],
      truncated: false,
    });
  });

  it("makes clean axis ticks and compact labels", () => {
    expect(niceTicks(0, 87)).toEqual([0, 20, 40, 60, 80, 100]);
    expect(niceTicks(-30, 45)).toEqual([-40, -20, 0, 20, 40, 60]);
    expect(compactNumber(12900)).toBe("12.9K");
    expect(compactNumber(4_200_000)).toBe("4.2M");
  });
});
