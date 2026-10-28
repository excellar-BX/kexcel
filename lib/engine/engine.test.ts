import { describe, expect, it } from "vitest";
import { colToLetters, lettersToCol, nameToAddress } from "./address";
import { displayValue, formatValue } from "./format";
import { adjustForStructureChange, renameSheetInFormula, shiftFormula } from "./references";
import { Workbook } from "./workbook";
import { parseCSV, parseTSV, toCSV, toTSV } from "../tsv";
import { fillPatches, sortPatches } from "../sheet-ops";

/** Builds a one-sheet workbook from { A1: "...", ... } with helpers addressed by cell name. */
function sheetWith(cells: Record<string, string>) {
  const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "Sheet1", cells }] });
  const sheet = {
    getValue: (row: number, col: number) => wb.getValue("s1", row, col),
  };
  const show = (name: string) => {
    const a = nameToAddress(name)!;
    return displayValue(wb.getValue("s1", a.row, a.col));
  };
  const raw = (name: string) => {
    const a = nameToAddress(name)!;
    return wb.getRaw("s1", a.row, a.col);
  };
  const set = (name: string, value: string) => {
    const a = nameToAddress(name)!;
    wb.setCell("s1", a.row, a.col, value);
  };
  return { wb, sheet, show, raw, set };
}

const calc = (formula: string, cells: Record<string, string> = {}) =>
  sheetWith({ ...cells, ZZ1: formula }).show("ZZ1");

describe("addresses", () => {
  it("converts column letters both ways", () => {
    expect(colToLetters(0)).toBe("A");
    expect(colToLetters(25)).toBe("Z");
    expect(colToLetters(26)).toBe("AA");
    expect(colToLetters(51)).toBe("AZ");
    expect(lettersToCol("AZ")).toBe(51);
    expect(nameToAddress("$B$3")).toEqual({ row: 2, col: 1 });
  });
});

describe("literals", () => {
  it("parses numbers, percents, booleans and text", () => {
    const { sheet } = sheetWith({ A1: "42", A2: "1,000", A3: "50%", A4: "true", A5: "hello", A6: "'007" });
    expect(sheet.getValue(0, 0)).toBe(42);
    expect(sheet.getValue(1, 0)).toBe(1000);
    expect(sheet.getValue(2, 0)).toBe(0.5);
    expect(sheet.getValue(3, 0)).toBe(true);
    expect(sheet.getValue(4, 0)).toBe("hello");
    expect(sheet.getValue(5, 0)).toBe("007");
  });
});

describe("operators", () => {
  it("follows Excel precedence", () => {
    expect(calc("=1+2*3")).toBe("7");
    expect(calc("=(1+2)*3")).toBe("9");
    expect(calc("=2^3^2")).toBe("64");
    expect(calc("=-2^2")).toBe("4");
    expect(calc("=50%*10")).toBe("5");
    expect(calc("=0.1+0.2")).toBe("0.3");
  });

  it("concatenates and compares", () => {
    expect(calc('="a"&1&TRUE')).toBe("a1TRUE");
    expect(calc('="abc"="ABC"')).toBe("TRUE");
    expect(calc("=1<2")).toBe("TRUE");
    expect(calc('=1<"a"')).toBe("TRUE");
    expect(calc("=A1=0")).toBe("TRUE");
  });

  it("produces errors", () => {
    expect(calc("=1/0")).toBe("#DIV/0!");
    expect(calc('="x"+1')).toBe("#VALUE!");
    expect(calc("=FOO(1)")).toBe("#NAME?");
    expect(calc("=1+")).toBe("#ERROR!");
    expect(calc("=A1+1", { A1: "=1/0" })).toBe("#DIV/0!");
  });
});

describe("functions", () => {
  const data = { A1: "10", A2: "20", A3: "hello", A4: "30", B1: "x", B2: "y", B3: "x", B4: "z" };

  it("aggregates ranges, skipping text", () => {
    expect(calc("=SUM(A1:A4)", data)).toBe("60");
    expect(calc("=AVERAGE(A1:A4)", data)).toBe("20");
    expect(calc("=MIN(A1:A4)", data)).toBe("10");
    expect(calc("=MAX(A1:A4, 99)", data)).toBe("99");
    expect(calc("=COUNT(A1:A4)", data)).toBe("3");
    expect(calc("=COUNTA(A1:A5)", data)).toBe("4");
    expect(calc("=COUNTBLANK(A1:A5)", data)).toBe("1");
    expect(calc("=AVERAGE(C1:C3)")).toBe("#DIV/0!");
  });

  it("handles conditionals", () => {
    expect(calc('=COUNTIF(B1:B4,"x")', data)).toBe("2");
    expect(calc('=SUMIF(B1:B4,"x",A1:A4)', data)).toBe("10");
    expect(calc('=SUMIF(A1:A4,">15")', data)).toBe("50");
    expect(calc('=COUNTIF(B1:B4,"<>x")', data)).toBe("2");
    expect(calc('=COUNTIF(A1:A4,"h*")', data)).toBe("1");
    expect(calc('=AVERAGEIF(A1:A4,">=20")', data)).toBe("25");
  });

  it("evaluates IF lazily and IFERROR", () => {
    expect(calc('=IF(A1>5,"big","small")', data)).toBe("big");
    expect(calc("=IF(FALSE,1/0,2)")).toBe("2");
    expect(calc("=IF(FALSE,1)")).toBe("FALSE");
    expect(calc('=IFERROR(1/0,"oops")')).toBe("oops");
    expect(calc("=AND(TRUE,1,A1>5)", data)).toBe("TRUE");
    expect(calc("=OR(FALSE,0)")).toBe("FALSE");
    expect(calc("=NOT(TRUE)")).toBe("FALSE");
  });

  it("rounds like Excel", () => {
    expect(calc("=ROUND(2.5)")).toBe("3");
    expect(calc("=ROUND(-2.5)")).toBe("-3");
    expect(calc("=ROUND(1.005,2)")).toBe("1.01");
    expect(calc("=ROUNDUP(1.21,1)")).toBe("1.3");
    expect(calc("=ROUNDDOWN(-1.29,1)")).toBe("-1.2");
    expect(calc("=MOD(-3,2)")).toBe("1");
    expect(calc("=INT(-1.5)")).toBe("-2");
    expect(calc("=SQRT(-1)")).toBe("#NUM!");
  });

  it("handles text functions", () => {
    expect(calc('=LEN("hello")')).toBe("5");
    expect(calc('=UPPER("hi")&LOWER("YO")')).toBe("HIyo");
    expect(calc('=TRIM("  a   b ")')).toBe("a b");
    expect(calc('=LEFT("hello",2)&RIGHT("hello",2)&MID("hello",2,3)')).toBe("heloell");
    expect(calc("=CONCAT(B1:B3)", data)).toBe("xyx");
  });

  it("looks values up", () => {
    const table = { A1: "apple", B1: "1.5", A2: "banana", B2: "0.5", A3: "cherry", B3: "3" };
    expect(calc('=VLOOKUP("banana",A1:B3,2,FALSE)', table)).toBe("0.5");
    expect(calc('=VLOOKUP("kiwi",A1:B3,2,FALSE)', table)).toBe("#N/A");
    expect(calc('=VLOOKUP("banana",A1:B3,3,FALSE)', table)).toBe("#REF!");
    expect(calc('=MATCH("cherry",A1:A3,0)', table)).toBe("3");
    expect(calc("=INDEX(A1:B3,2,2)", table)).toBe("0.5");
    const grades = { A1: "0", B1: "F", A2: "50", B2: "C", A3: "80", B3: "A" };
    expect(calc("=VLOOKUP(65,A1:B3,2)", grades)).toBe("C");
  });
});

describe("recalculation", () => {
  it("updates dependents through chains", () => {
    const { show, set } = sheetWith({ A1: "1", A2: "=A1*2", A3: "=A2+A1", B1: "=SUM(A1:A3)" });
    expect(show("A3")).toBe("3");
    expect(show("B1")).toBe("6");
    set("A1", "10");
    expect(show("A2")).toBe("20");
    expect(show("A3")).toBe("30");
    expect(show("B1")).toBe("60");
  });

  it("picks up cells filled in after the formula", () => {
    const { show, set } = sheetWith({ A1: "=SUM(B1:B3)" });
    expect(show("A1")).toBe("0");
    set("B2", "5");
    expect(show("A1")).toBe("5");
    set("B2", "");
    expect(show("A1")).toBe("0");
  });

  it("drops stale dependencies when a formula changes", () => {
    const { show, set } = sheetWith({ A1: "1", B1: "2", C1: "=A1" });
    set("C1", "=B1");
    set("A1", "100");
    expect(show("C1")).toBe("2");
  });

  it("detects cycles and recovers when they are broken", () => {
    const { show, set } = sheetWith({ A1: "=B1+1", B1: "=A1+1", C1: "=A1" });
    expect(show("A1")).toBe("#CYCLE!");
    expect(show("B1")).toBe("#CYCLE!");
    expect(show("C1")).toBe("#CYCLE!");
    set("B1", "5");
    expect(show("A1")).toBe("6");
    expect(show("C1")).toBe("6");
    expect(calc("=ZZ1")).toBe("#CYCLE!");
  });

  it("round-trips through JSON", () => {
    const { wb } = sheetWith({ A1: "3", B2: "=A1*2" });
    wb.setFormat("s1", 0, 0, { bold: true });
    const copy = Workbook.fromJSON(wb.toJSON());
    expect(copy.getValue("s1", 1, 1)).toBe(6);
    expect(copy.getFormat("s1", 0, 0)).toEqual({ bold: true });
  });
});

describe("multiple sheets", () => {
  function twoSheets() {
    return Workbook.fromJSON({
      sheets: [
        { id: "s1", name: "Sheet1", cells: { A1: "=Sales!B2*2", A2: "='My data'!A1+1" } },
        { id: "s2", name: "Sales", cells: { B2: "21" } },
        { id: "s3", name: "My data", cells: { A1: "9" } },
      ],
    });
  }

  it("reads cells on other sheets", () => {
    const wb = twoSheets();
    expect(wb.getValue("s1", 0, 0)).toBe(42);
    expect(wb.getValue("s1", 1, 0)).toBe(10);
    wb.setCell("s2", 1, 1, "50");
    expect(wb.getValue("s1", 0, 0)).toBe(100);
  });

  it("updates formulas when a sheet is renamed", () => {
    const wb = twoSheets();
    wb.renameSheet("s2", "Q1 sales");
    expect(wb.getRaw("s1", 0, 0)).toBe("='Q1 sales'!B2*2");
    expect(wb.getValue("s1", 0, 0)).toBe(42);
    expect(() => wb.renameSheet("s2", "my DATA")).toThrow();
  });

  it("shows #REF! after a referenced sheet is removed", () => {
    const wb = twoSheets();
    wb.removeSheet("s2");
    expect(displayValue(wb.getValue("s1", 0, 0))).toBe("#REF!");
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "My data"]);
  });

  it("adds, duplicates and names sheets", () => {
    const wb = Workbook.withSheet();
    const id = wb.addSheet();
    expect(wb.sheetName(id)).toBe("Sheet2");
    wb.setCell(id, 0, 0, "5");
    wb.setCell(id, 0, 1, "=A1*2");
    const copy = wb.duplicateSheet(id);
    expect(wb.sheetName(copy)).toBe("Sheet2 (copy)");
    expect(wb.getValue(copy, 0, 1)).toBe(10);
    expect(wb.sheets.map((s) => s.id)).toEqual(["s1", id, copy]);
  });
});

describe("inserting and deleting rows and columns", () => {
  it("moves cells and keeps formulas pointing at the same data", () => {
    const { wb, show, raw } = sheetWith({ A1: "1", A2: "2", A3: "3", B1: "=SUM(A1:A3)", B2: "=A3" });
    wb.insert("s1", "row", 1, 2); // two blank rows before row 2
    expect(raw("A4")).toBe("2");
    expect(raw("B1")).toBe("=SUM(A1:A5)");
    expect(raw("B4")).toBe("=A5");
    expect(show("B1")).toBe("6");
  });

  it("shrinks ranges and turns deleted references into #REF!", () => {
    const { wb, show, raw } = sheetWith({ A1: "1", A2: "2", A3: "3", B1: "=SUM(A1:A3)", C1: "=A2" });
    wb.remove("s1", "row", 1, 1);
    expect(raw("B1")).toBe("=SUM(A1:A2)");
    expect(show("B1")).toBe("4");
    expect(raw("C1")).toBe("=#REF!");
    expect(show("C1")).toBe("#REF!");
  });

  it("adjusts columns, anchored references and other sheets", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        { id: "s1", name: "Data", cells: { A1: "1", B1: "2", C1: "=$B$1+A1" } },
        { id: "s2", name: "Report", cells: { A1: "=Data!B1", B1: "=B1" } },
      ],
    });
    wb.insert("s1", "col", 1, 1);
    expect(wb.getRaw("s1", 0, 3)).toBe("=$C$1+A1");
    expect(wb.getRaw("s2", 0, 0)).toBe("=Data!C1");
    expect(wb.getRaw("s2", 0, 1)).toBe("=B1"); // its own sheet didn't change
    expect(wb.getValue("s2", 0, 0)).toBe(2);
  });

  it("moves formats and sizes with their rows", () => {
    const { wb } = sheetWith({ A2: "x" });
    wb.setFormat("s1", 1, 0, { bold: true });
    wb.setRowHeight("s1", 1, 40);
    wb.insert("s1", "row", 0, 1);
    expect(wb.getFormat("s1", 2, 0)).toEqual({ bold: true });
    expect(wb.rowHeights("s1").get(2)).toBe(40);
  });
});

describe("patches and undo", () => {
  it("returns inverse patches that restore the previous state", () => {
    const { wb, raw } = sheetWith({ A1: "1" });
    const undo = wb.apply([
      { sheet: "s1", row: 0, col: 0, raw: "2", format: { bold: true } },
      { sheet: "s1", row: 1, col: 0, raw: "=A1" },
    ]);
    expect(wb.getValue("s1", 1, 0)).toBe(2);
    wb.apply(undo);
    expect(raw("A1")).toBe("1");
    expect(raw("A2")).toBe("");
    expect(wb.getFormat("s1", 0, 0)).toBeUndefined();
  });
});

describe("number formats", () => {
  it("formats numbers, currency, percents and decimals", () => {
    expect(formatValue(1234.5, { numFmt: "number" }, "en-US")).toBe("1,234.50");
    expect(formatValue(1234.5, { numFmt: "currency", currency: "USD" }, "en-US")).toBe("$1,234.50");
    expect(formatValue(0.256, { numFmt: "percent", decimals: 1 }, "en-US")).toBe("25.6%");
    expect(formatValue(3.14159, { decimals: 2 }, "en-US")).toBe("3.14");
    expect(formatValue(12345, { numFmt: "scientific" }, "en-US")).toBe("1.23E+4");
    expect(formatValue("text", { numFmt: "currency" }, "en-US")).toBe("text");
  });
});

describe("fill and sort", () => {
  it("continues numeric series and repeats formulas", () => {
    const { wb, raw } = sheetWith({ A1: "1", A2: "3", B1: "=A1*10" });
    wb.apply(fillPatches(wb, "s1", { r0: 0, r1: 1, c0: 0, c1: 0 }, { r0: 0, r1: 4, c0: 0, c1: 0 }, "series"));
    expect([raw("A3"), raw("A4"), raw("A5")]).toEqual(["5", "7", "9"]);
    wb.apply(fillPatches(wb, "s1", { r0: 0, r1: 0, c0: 1, c1: 1 }, { r0: 0, r1: 2, c0: 1, c1: 1 }, "copy"));
    expect(raw("B3")).toBe("=A3*10");
    expect(wb.getValue("s1", 2, 1)).toBe(50);
  });

  it("sorts rows by a column, keeping blanks last", () => {
    const { wb, raw } = sheetWith({ A1: "pear", B1: "3", A2: "", B2: "9", A3: "apple", B3: "1", A4: "10", B4: "4" });
    wb.apply(sortPatches(wb, "s1", { r0: 0, r1: 3, c0: 0, c1: 1 }, 0, true));
    expect([raw("A1"), raw("A2"), raw("A3"), raw("A4")]).toEqual(["10", "apple", "pear", ""]);
    expect([raw("B1"), raw("B2"), raw("B3"), raw("B4")]).toEqual(["4", "1", "3", "9"]);
    wb.apply(sortPatches(wb, "s1", { r0: 0, r1: 3, c0: 0, c1: 1 }, 1, false));
    expect([raw("B1"), raw("B4")]).toEqual(["9", "1"]);
  });
});

describe("shiftFormula", () => {
  it("moves relative references and keeps anchored ones", () => {
    expect(shiftFormula("=A1+B2", 1, 1)).toBe("=B2+C3");
    expect(shiftFormula("=$A$1+A$1+$A1", 2, 2)).toBe("=$A$1+C$1+$A3");
    expect(shiftFormula("=SUM(A1:A3)", 0, 1)).toBe("=SUM(B1:B3)");
    expect(shiftFormula('="A1"&A1', 1, 0)).toBe('="A1"&A2');
    expect(shiftFormula("hello", 1, 1)).toBe("hello");
  });

  it("turns references pushed off the sheet into #REF!", () => {
    const shifted = shiftFormula("=A1+1", -1, 0);
    expect(shifted).toBe("=#REF!+1");
    expect(calc(shifted)).toBe("#REF!");
  });

  it("keeps sheet prefixes", () => {
    expect(shiftFormula("=Sales!A1+'My data'!B2:C3", 1, 0)).toBe("=Sales!A2+'My data'!B3:C4");
  });
});

describe("structure and rename rewriting", () => {
  const own = (ctx: { sheet: string | undefined }) => ctx.sheet === undefined;

  it("expands a range when rows are inserted inside it", () => {
    expect(adjustForStructureChange("=SUM(A2:A5)", "row", 2, 3, own)).toBe("=SUM(A2:A8)");
    expect(adjustForStructureChange("=SUM(A2:A5)", "row", 0, 1, own)).toBe("=SUM(A3:A6)");
    expect(adjustForStructureChange("=SUM(A2:A5)", "row", 9, 1, own)).toBe("=SUM(A2:A5)");
  });

  it("drops a fully deleted range", () => {
    const adjusted = adjustForStructureChange("=SUM(B1:C1)+1", "col", 1, -2, own);
    expect(adjusted).toBe("=SUM(#REF!)+1");
    expect(calc(adjusted)).toBe("#REF!");
  });

  it("renames quoted and bare sheet prefixes", () => {
    expect(renameSheetInFormula("=sales!A1+Other!A1", "Sales", "Q1 Sales")).toBe("='Q1 Sales'!A1+Other!A1");
    expect(renameSheetInFormula("='Q1 Sales'!A1", "q1 sales", "Sales")).toBe("=Sales!A1");
  });
});

describe("TSV", () => {
  it("round-trips values with tabs, quotes and newlines", () => {
    const rows = [["a", "b\tc"], ['say "hi"', "line1\nline2"]];
    expect(parseTSV(toTSV(rows))).toEqual(rows);
    expect(parseTSV("1\t2\r\n3\t4\r\n")).toEqual([["1", "2"], ["3", "4"]]);
  });

  it("reads and writes CSV", () => {
    const rows = [["name", "note"], ["Ada", "likes, commas"]];
    expect(toCSV(rows)).toBe('name,note\nAda,"likes, commas"');
    expect(parseCSV("﻿" + toCSV(rows))).toEqual(rows);
  });
});
