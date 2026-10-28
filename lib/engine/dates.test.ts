import { describe, expect, it } from "vitest";
import { nameToAddress } from "./address";
import { dateToSerial, formatSerial, parseDateText } from "./dates";
import { displayValue, formatValue } from "./format";
import { impliedDateKind } from "./literal";
import { Workbook } from "./workbook";

function sheetWith(cells: Record<string, string>) {
  const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "Sheet1", cells }] });
  const value = (name: string) => {
    const a = nameToAddress(name)!;
    return wb.getValue("s1", a.row, a.col);
  };
  return { wb, value, show: (name: string) => displayValue(value(name)) };
}

const calc = (formula: string, cells: Record<string, string> = {}) => sheetWith({ ...cells, ZZ1: formula }).show("ZZ1");

describe("date serials", () => {
  it("match Excel's serial numbers", () => {
    expect(dateToSerial(1900, 3, 1)).toBe(61);
    expect(dateToSerial(2025, 1, 1)).toBe(45658);
    expect(dateToSerial(2025, 13, 1)).toBe(dateToSerial(2026, 1, 1));
  });

  it("read typed dates and times", () => {
    expect(parseDateText("2025-01-15")).toEqual({ serial: 45672, kind: "date" });
    expect(parseDateText("2025/1/5")).toEqual({ serial: 45662, kind: "date" });
    expect(parseDateText("15 Jan 2025")?.serial).toBe(45672);
    expect(parseDateText("January 15, 2025")?.serial).toBe(45672);
    expect(parseDateText("14:30")).toEqual({ serial: 14.5 / 24, kind: "time" });
    expect(parseDateText("2:30 pm")?.serial).toBe(14.5 / 24);
    expect(parseDateText("2025-01-15 09:00")).toEqual({ serial: 45672.375, kind: "datetime" });
    expect(parseDateText("2025-02-30")).toBeNull();
    expect(parseDateText("01/02/2025")).toBeNull(); // day and month order is ambiguous
    expect(parseDateText("hello")).toBeNull();
  });

  it("format dates, times and both", () => {
    expect(formatSerial(45672, "date")).toBe("2025-01-15");
    expect(formatSerial(45672.375, "datetime")).toBe("2025-01-15 09:00");
    expect(formatSerial(0.5, "time")).toBe("12:00");
    expect(formatSerial(45672, "date", "medium", "en-GB")).toBe("15 Jan 2025");
    expect(formatValue(45672, { numFmt: "date" })).toBe("2025-01-15");
  });
});

describe("typed dates in cells", () => {
  it("become numbers that date functions and arithmetic understand", () => {
    const { value } = sheetWith({ A1: "2025-01-15", A2: "2025-03-01", B1: "=A2-A1", B2: "=YEAR(A1)", B3: "=A1+30" });
    expect(value("A1")).toBe(45672);
    expect(value("B1")).toBe(45);
    expect(value("B2")).toBe(2025);
    expect(value("B3")).toBe(45702);
  });

  it("tell the editor which entries should get a date format", () => {
    expect(impliedDateKind("2025-01-15")).toBe("date");
    expect(impliedDateKind("9:30")).toBe("time");
    expect(impliedDateKind("=TODAY()")).toBe("date");
    expect(impliedDateKind("=NOW()")).toBe("datetime");
    expect(impliedDateKind("=YEAR(A1)")).toBeNull();
    expect(impliedDateKind("'2025-01-15")).toBeNull();
    expect(impliedDateKind("42")).toBeNull();
  });
});

describe("date functions", () => {
  it("build and take apart dates", () => {
    expect(calc("=DATE(2025,1,15)")).toBe("45672");
    expect(calc("=DATE(2025,14,1)")).toBe(String(dateToSerial(2026, 2, 1)));
    expect(calc("=MONTH(DATE(2025,7,4))&\"/\"&DAY(DATE(2025,7,4))")).toBe("7/4");
    expect(calc('=YEAR("2024-02-29")')).toBe("2024");
    expect(calc("=HOUR(TIME(14,45,0))*100+MINUTE(TIME(14,45,0))")).toBe("1445");
    expect(calc('=DATEVALUE("15 Jan 2025")')).toBe("45672");
    expect(calc('=DATEVALUE("nope")')).toBe("#VALUE!");
  });

  it("work out weekdays, month ends and differences", () => {
    expect(calc("=WEEKDAY(DATE(2025,1,15))")).toBe("4"); // a Wednesday
    expect(calc("=WEEKDAY(DATE(2025,1,15),2)")).toBe("3");
    expect(calc("=EOMONTH(DATE(2024,1,31),1)")).toBe(String(dateToSerial(2024, 2, 29)));
    expect(calc("=EDATE(DATE(2025,1,31),1)")).toBe(String(dateToSerial(2025, 2, 28)));
    expect(calc("=DAYS(DATE(2025,3,1),DATE(2025,1,1))")).toBe("59");
    expect(calc('=DATEDIF(DATE(2000,5,20),DATE(2025,5,19),"Y")')).toBe("24");
    expect(calc('=DATEDIF(DATE(2025,1,10),DATE(2025,3,5),"M")')).toBe("1");
    expect(calc('=DATEDIF(DATE(2025,1,10),DATE(2025,3,5),"MD")')).toBe("23");
    expect(calc("=NETWORKDAYS(DATE(2025,1,13),DATE(2025,1,24))")).toBe("10");
  });

  it("recalculate TODAY on every change", () => {
    const { wb, value } = sheetWith({ A1: "=TODAY()", B1: "1" });
    const today = new Date();
    expect(value("A1")).toBe(dateToSerial(today.getFullYear(), today.getMonth() + 1, today.getDate()));
    wb.setCell("s1", 0, 0, "=TODAY()+1");
    expect(value("A1")).toBe(dateToSerial(today.getFullYear(), today.getMonth() + 1, today.getDate()) + 1);
  });
});

describe("loading typed dates", () => {
  it("gives date cells without a number format a date format", () => {
    const wb = Workbook.fromJSON({
      sheets: [{ id: "s1", name: "S", cells: { A1: "2025-01-15", A2: "9:30", A3: "'2025-01-15", A4: "5" }, formats: { A4: { bold: true } } }],
    });
    expect(wb.getFormat("s1", 0, 0)).toEqual({ numFmt: "date" });
    expect(wb.getFormat("s1", 1, 0)).toEqual({ numFmt: "time" });
    expect(wb.getFormat("s1", 2, 0)).toBeUndefined();
    expect(wb.getFormat("s1", 3, 0)).toEqual({ bold: true });
    expect(formatValue(wb.getValue("s1", 0, 0), wb.getFormat("s1", 0, 0))).toBe("2025-01-15");
  });
});

describe("merged ranges", () => {
  it("are saved, drop overlaps, and follow inserted and deleted rows", () => {
    const wb = Workbook.withSheet();
    wb.setMerges("s1", [
      { r0: 0, c0: 0, r1: 1, c1: 2 },
      { r0: 1, c0: 1, r1: 3, c1: 3 }, // overlaps the first, so it's dropped
      { r0: 5, c0: 0, r1: 5, c1: 0 }, // a single cell isn't a merge
      { r0: 5, c0: 0, r1: 6, c1: 1 },
    ]);
    expect(wb.merges("s1")).toEqual([
      { r0: 0, c0: 0, r1: 1, c1: 2 },
      { r0: 5, c0: 0, r1: 6, c1: 1 },
    ]);
    expect(wb.mergeAt("s1", 1, 2)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(Workbook.fromJSON(wb.toJSON()).merges("s1")).toEqual(wb.merges("s1"));

    wb.insert("s1", "row", 1, 2); // inside the first merge: it grows
    expect(wb.merges("s1")).toEqual([
      { r0: 0, c0: 0, r1: 3, c1: 2 },
      { r0: 7, c0: 0, r1: 8, c1: 1 },
    ]);
    wb.remove("s1", "row", 8, 1); // the second merge shrinks to one row
    expect(wb.merges("s1")[1]).toEqual({ r0: 7, c0: 0, r1: 7, c1: 1 });
    wb.remove("s1", "col", 1, 1); // ...and then to one cell, so it's gone
    expect(wb.merges("s1")).toEqual([{ r0: 0, c0: 0, r1: 3, c1: 1 }]);
  });
});
