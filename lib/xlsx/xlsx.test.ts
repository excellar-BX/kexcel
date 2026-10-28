import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Workbook } from "@/lib/engine";
import { formatFromCode, isDateFormat, serialToDateText } from "./numfmt";
import { readXlsx } from "./read";
import { writeXlsx } from "./write";
import { parseXml } from "./xml";

const LIMITS = { maxRows: 10_000, maxCols: 52 };

describe("xml", () => {
  it("parses elements, attributes, entities and CDATA", () => {
    const doc = parseXml(`<?xml version="1.0"?><x:a k='1' j="&lt;b&gt;"><b>A &amp; B &#65;</b><!-- skip --><c/><d><![CDATA[<raw>]]></d></x:a>`);
    const a = doc.children[0];
    expect(a.name).toBe("a");
    expect(a.attrs).toEqual({ k: "1", j: "<b>" });
    expect(a.children.map((c) => c.name)).toEqual(["b", "c", "d"]);
    expect(a.children[0].text).toBe("A & B A");
    expect(a.children[2].text).toBe("<raw>");
  });
});

describe("number formats", () => {
  it("maps Excel codes to Kexcel formats", () => {
    expect(formatFromCode("#,##0.00")).toEqual({ numFmt: "number", decimals: 2 });
    expect(formatFromCode("0.0%")).toEqual({ numFmt: "percent", decimals: 1 });
    expect(formatFromCode('"₦"#,##0.00')).toEqual({ numFmt: "currency", currency: "NGN", decimals: 2 });
    expect(formatFromCode("[$€-x-euro2] #,##0")).toEqual({ numFmt: "currency", currency: "EUR", decimals: 0 });
    expect(formatFromCode("0.000")).toEqual({ decimals: 3 });
    expect(formatFromCode("General")).toEqual({});
  });

  it("recognises dates and converts serial numbers", () => {
    expect(isDateFormat(14, undefined)).toBe(true);
    expect(isDateFormat(164, "dd/mm/yyyy")).toBe(true);
    expect(isDateFormat(164, '#,##0.00 "days"')).toBe(false);
    expect(serialToDateText(45658)).toBe("2025-01-01");
    expect(serialToDateText(45658.5)).toBe("2025-01-01 12:00");
  });
});

describe("xlsx round trip", () => {
  it("keeps values, formulas, formats, sizes and frozen panes", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        {
          id: "s1",
          name: "Sales",
          cells: { A1: "Region", B1: "Amount", A2: "Lagos", B2: "1200.5", A3: "Kano", B3: "800", B4: "=SUM(B2:B3)", C1: "'007", C2: "TRUE" },
          formats: {
            A1: { bold: true, fill: "#e3f2fd", borders: { bottom: true } },
            B2: { numFmt: "currency", currency: "NGN", decimals: 2 },
            B4: { numFmt: "number", decimals: 0, italic: true, color: "#c62828", align: "center", fontSize: 14 },
          },
          colWidths: { 0: 160 },
          rowHeights: { 3: 40 },
          freeze: { rows: 1, cols: 1 },
          merges: ["D1:E2"],
        },
        { id: "s2", name: "My data", cells: { A1: "=Sales!B4*2" } },
      ],
    });

    const { workbook, notes } = readXlsx(writeXlsx(wb), LIMITS);
    expect(notes).toEqual([]);
    const back = Workbook.fromJSON(workbook);
    const [sales, mine] = back.sheets;
    expect(back.sheets.map((s) => s.name)).toEqual(["Sales", "My data"]);
    expect(back.getRaw(sales.id, 3, 1)).toBe("=SUM(B2:B3)");
    expect(back.getValue(sales.id, 3, 1)).toBe(2000.5);
    expect(back.getValue(mine.id, 0, 0)).toBe(4001);
    expect(back.getRaw(sales.id, 0, 2)).toBe("'007");
    expect(back.getValue(sales.id, 1, 2)).toBe(true);
    expect(back.getFormat(sales.id, 0, 0)).toEqual({ bold: true, fill: "#e3f2fd", borders: { top: false, right: false, bottom: true, left: false } });
    expect(back.getFormat(sales.id, 1, 1)).toEqual({ numFmt: "currency", currency: "NGN", decimals: 2 });
    expect(back.getFormat(sales.id, 3, 1)).toEqual({
      italic: true,
      fontSize: 14,
      color: "#c62828",
      numFmt: "number",
      decimals: 0,
      align: "center",
    });
    expect(back.colWidths(sales.id).get(0)).toBe(160);
    expect(back.rowHeights(sales.id).get(3)).toBe(40);
    expect(back.getFreeze(sales.id)).toEqual({ rows: 1, cols: 1 });
    expect(back.merges(sales.id)).toEqual([{ r0: 0, c0: 3, r1: 1, c1: 4 }]);
  });

  it("keeps dates, times and wrapped text", () => {
    const wb = Workbook.fromJSON({
      sheets: [
        {
          id: "s1",
          name: "S",
          cells: { A1: "2025-01-15", A2: "14:30", A3: "2025-01-15 09:05", A4: "Long text" },
          formats: {
            A1: { numFmt: "date", dateStyle: "medium" },
            A2: { numFmt: "time" },
            A3: { numFmt: "datetime" },
            A4: { wrap: true },
          },
        },
      ],
    });
    const back = Workbook.fromJSON(readXlsx(writeXlsx(wb), LIMITS).workbook);
    expect(back.getValue("s1", 0, 0)).toBe(45672);
    expect(back.getFormat("s1", 0, 0)).toEqual({ numFmt: "date" });
    expect(back.getRaw("s1", 1, 0)).toBe("14:30");
    expect(back.getFormat("s1", 1, 0)).toEqual({ numFmt: "time" });
    expect(back.getRaw("s1", 2, 0)).toBe("2025-01-15 09:05");
    expect(back.getFormat("s1", 2, 0)).toEqual({ numFmt: "datetime" });
    expect(back.getFormat("s1", 3, 0)).toEqual({ wrap: true });
  });

  it("writes cached results and a recalculation flag for Excel", () => {
    const wb = Workbook.fromJSON({ sheets: [{ id: "s1", name: "S", cells: { A1: "2", A2: "=A1*3", A3: '="x"&A1' } }] });
    const files = unzipSync(writeXlsx(wb));
    const sheet = strFromU8(files["xl/worksheets/sheet1.xml"]);
    expect(sheet).toContain('<c r="A2"><f>A1*3</f><v>6</v></c>');
    expect(sheet).toContain('<c r="A3" t="str"><f>&quot;x&quot;&amp;A1</f><v>x2</v></c>');
    expect(strFromU8(files["xl/workbook.xml"])).toContain('fullCalcOnLoad="1"');
  });
});

describe("reading files written by Excel", () => {
  /** A small workbook in the shape Excel writes: shared strings, shared formulas, styles, a date and a merge. */
  function excelLikeFile() {
    const x = (s: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`);
    const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    return zipSync({
      "_rels/.rels": x(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      ),
      "xl/workbook.xml": x(
        `<workbook ${main} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Q1: Sales" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      ),
      "xl/_rels/workbook.xml.rels": x(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
          '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
          "</Relationships>",
      ),
      "xl/sharedStrings.xml": x(
        `<sst ${main} count="3" uniqueCount="3"><si><t>Item</t></si><si><r><rPr><b/></rPr><t>Pri</t></r><r><t>ce</t></r></si><si><t>0042</t></si></sst>`,
      ),
      "xl/styles.xml": x(
        `<styleSheet ${main}><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts>` +
          '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF1F4E79"/><name val="Calibri"/></font></fonts>' +
          '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0"/>' +
          '<xf numFmtId="164" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>',
      ),
      "xl/worksheets/sheet1.xml": x(
        `<worksheet ${main}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
          '<cols><col min="1" max="1" width="20.7109375" customWidth="1"/></cols><sheetData>' +
          '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>' +
          '<row r="2"><c r="A2" t="inlineStr"><is><t>Pen</t></is></c><c r="B2" s="2"><v>1.5</v></c><c r="C2"><f t="shared" ref="C2:C3" si="0">B2*2</f><v>3</v></c><c r="D2" s="3"><v>45658</v></c></row>' +
          '<row r="3" ht="30" customHeight="1"><c r="A3" t="str"><v>Ink</v></c><c r="B3" s="2"><v>2</v></c><c r="C3"><f t="shared" si="0"/><v>4</v></c><c r="D3"><f>_xlfn.CONCAT(A2,A3)</f><v>PenInk</v></c></row>' +
          '</sheetData><mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells></worksheet>',
      ),
    });
  }

  it("reads strings, styles, shared formulas, dates and panes", () => {
    const { workbook, notes } = readXlsx(excelLikeFile(), LIMITS);
    const wb = Workbook.fromJSON(workbook);
    const id = wb.sheets[0].id;
    expect(wb.sheets[0].name).toBe("Q1_ Sales");
    expect(wb.getRaw(id, 0, 1)).toBe("Price");
    expect(wb.getRaw(id, 0, 2)).toBe("'0042");
    expect(wb.getRaw(id, 1, 0)).toBe("Pen");
    expect(wb.getRaw(id, 2, 2)).toBe("=B3*2");
    expect(wb.getValue(id, 2, 2)).toBe(4);
    expect(wb.getRaw(id, 1, 3)).toBe("2025-01-01");
    expect(wb.getValue(id, 1, 3)).toBe(45658);
    expect(wb.getFormat(id, 1, 3)).toEqual({ numFmt: "date" });
    expect(wb.merges(id)).toEqual([{ r0: 0, c0: 4, r1: 0, c1: 5 }]);
    expect(wb.getRaw(id, 2, 3)).toBe("=CONCAT(A2,A3)");
    expect(wb.getValue(id, 2, 3)).toBe("PenInk");
    expect(wb.getFormat(id, 0, 0)).toEqual({ bold: true, color: "#1f4e79", fill: "#ffff00" });
    expect(wb.getFormat(id, 1, 1)).toEqual({ numFmt: "currency", currency: "USD", decimals: 2 });
    expect(wb.colWidths(id).get(0)).toBe(150);
    expect(wb.rowHeights(id).get(2)).toBe(40);
    expect(wb.getFreeze(id)).toEqual({ rows: 1, cols: 0 });
    expect(notes).toEqual([]);
  });

  it("explains files that aren't workbooks", () => {
    expect(() => readXlsx(strToU8("not a zip"), LIMITS)).toThrow(/isn't an Excel workbook/);
  });
});
