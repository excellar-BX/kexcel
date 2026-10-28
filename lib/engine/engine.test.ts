import { describe, expect, it } from "vitest";
import { colToLetters, lettersToCol, nameToAddress } from "./address";
import { displayValue, Sheet, shiftFormula } from "./sheet";
import { parseTSV, toTSV } from "../tsv";

/** Builds a sheet from { A1: "...", ... } and returns a display-value reader. */
function sheetWith(cells: Record<string, string>) {
  const sheet = Sheet.fromJSON(cells);
  const show = (name: string) => {
    const a = nameToAddress(name)!;
    return displayValue(sheet.getValue(a.row, a.col));
  };
  const set = (name: string, raw: string) => {
    const a = nameToAddress(name)!;
    sheet.setCell(a.row, a.col, raw);
  };
  return { sheet, show, set };
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
    const { sheet } = sheetWith({ A1: "3", B2: "=A1*2" });
    const copy = Sheet.fromJSON(sheet.toJSON());
    expect(copy.getValue(1, 1)).toBe(6);
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
});

describe("TSV", () => {
  it("round-trips values with tabs, quotes and newlines", () => {
    const rows = [["a", "b\tc"], ['say "hi"', "line1\nline2"]];
    expect(parseTSV(toTSV(rows))).toEqual(rows);
    expect(parseTSV("1\t2\r\n3\t4\r\n")).toEqual([["1", "2"], ["3", "4"]]);
  });
});
