import { describe, expect, it } from "vitest";
import { Workbook } from "@/lib/engine";
import { describeSheet, MAX_CONTEXT_ROWS } from "./context";

const sel = { r0: 0, r1: 0, c0: 0, c1: 0 };

describe("describeSheet", () => {
  it("lists tabs, selection and filled cells with formula results", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        { id: "s1", name: "Sales", cells: { A1: "Region", B1: "Total", A2: "Lagos", B2: "1200", B3: "=SUM(B2:B2)" } },
        { id: "s2", name: "Notes", cells: {} },
      ],
    });
    const { text, truncated } = describeSheet(wb, "s1", { r0: 1, r1: 2, c0: 1, c1: 1 }, { row: 1, col: 1 });
    expect(text).toContain('Workbook tabs: "Sales", "Notes"');
    expect(text).toContain("Selection: B2:B3");
    expect(text).toContain('A1: "Region" | B1: "Total"');
    expect(text).toContain("B3: =SUM(B2:B2) → 1200");
    expect(truncated).toBeNull();
  });

  it("says when a large tab was cut short", () => {
    const cells: Record<string, string> = {};
    for (let r = 1; r <= MAX_CONTEXT_ROWS + 50; r++) cells[`A${r}`] = String(r);
    const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "Big", cells }] });
    const { text, truncated } = describeSheet(wb, "s1", sel, { row: 0, col: 0 });
    expect(truncated).toEqual({ shownRows: MAX_CONTEXT_ROWS, totalRows: MAX_CONTEXT_ROWS + 50, shownCols: 1, totalCols: 1 });
    expect(text).toContain(`only rows 1-${MAX_CONTEXT_ROWS}`);
  });

  it("handles an empty tab", () => {
    const { text } = describeSheet(Workbook.withSheet(), "s1", sel, { row: 0, col: 0 });
    expect(text).toContain("The tab is empty.");
  });
});
