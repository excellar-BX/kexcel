import { strToU8, zipSync } from "fflate";
import {
  addressToName,
  colToLetters,
  DEFAULT_FONT_SIZE,
  isError,
  isFormula,
  parseKey,
  type CellFormat,
  type Workbook,
} from "@/lib/engine";
import { codeFromFormat } from "./numfmt";
import { escapeXml } from "./xml";

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Error values Excel knows; Kexcel's own (#CYCLE!, #ERROR!) are left for Excel to recompute. */
const EXCEL_ERRORS = new Set(["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!"]);

const FONT_NAMES = { sans: "Arial", serif: "Times New Roman", mono: "Courier New" } as const;

const argb = (hex: string) => "FF" + hex.replace("#", "").toUpperCase().padStart(6, "0").slice(-6);

/** Collects the distinct cell formats into Excel's style tables. Index 0 is the default style. */
class StyleTable {
  private fonts = new Map<string, number>();
  private fills = new Map<string, number>();
  private borders = new Map<string, number>();
  private numFmts = new Map<string, number>();
  private xfs = new Map<string, number>();
  private fontXml: string[] = [];
  private fillXml = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  private borderXml = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  private xfXml: string[] = [];

  constructor() {
    this.font({});
    this.xfXml.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    this.xfs.set("{}", 0);
  }

  private font(f: CellFormat): number {
    const parts = [
      f.bold ? "<b/>" : "",
      f.italic ? "<i/>" : "",
      f.strike ? "<strike/>" : "",
      f.underline ? "<u/>" : "",
      `<sz val="${f.fontSize ?? DEFAULT_FONT_SIZE}"/>`,
      f.color ? `<color rgb="${argb(f.color)}"/>` : "",
      `<name val="${FONT_NAMES[f.fontFamily ?? "sans"]}"/>`,
    ].join("");
    let id = this.fonts.get(parts);
    if (id === undefined) {
      id = this.fontXml.length;
      this.fonts.set(parts, id);
      this.fontXml.push(`<font>${parts}</font>`);
    }
    return id;
  }

  private fill(color: string | undefined): number {
    if (!color) return 0;
    let id = this.fills.get(color);
    if (id === undefined) {
      id = this.fillXml.length;
      this.fills.set(color, id);
      this.fillXml.push(`<fill><patternFill patternType="solid"><fgColor rgb="${argb(color)}"/><bgColor indexed="64"/></patternFill></fill>`);
    }
    return id;
  }

  private border(b: CellFormat["borders"]): number {
    if (!b || !Object.values(b).some(Boolean)) return 0;
    const side = (name: string, on?: boolean) => (on ? `<${name} style="thin"><color auto="1"/></${name}>` : `<${name}/>`);
    const xml = `<border>${side("left", b.left)}${side("right", b.right)}${side("top", b.top)}${side("bottom", b.bottom)}<diagonal/></border>`;
    let id = this.borders.get(xml);
    if (id === undefined) {
      id = this.borderXml.length;
      this.borders.set(xml, id);
      this.borderXml.push(xml);
    }
    return id;
  }

  private numFmt(f: CellFormat): number {
    const code = codeFromFormat(f);
    if (code === null) return 0;
    let id = this.numFmts.get(code);
    if (id === undefined) {
      id = 164 + this.numFmts.size;
      this.numFmts.set(code, id);
    }
    return id;
  }

  /** The style index for a cell format. */
  styleOf(f: CellFormat | undefined): number {
    if (!f) return 0;
    const key = JSON.stringify(f);
    let id = this.xfs.get(key);
    if (id !== undefined) return id;
    const numFmtId = this.numFmt(f);
    const fontId = this.font(f);
    const fillId = this.fill(f.fill);
    const borderId = this.border(f.borders);
    const align =
      f.align || f.wrap
        ? `<alignment${f.align ? ` horizontal="${f.align}"` : ""}${f.wrap ? ' wrapText="1"' : ""}/>`
        : "";
    id = this.xfXml.length;
    this.xfs.set(key, id);
    this.xfXml.push(
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
        `${numFmtId ? ' applyNumberFormat="1"' : ""}${fontId ? ' applyFont="1"' : ""}${fillId ? ' applyFill="1"' : ""}` +
        `${borderId ? ' applyBorder="1"' : ""}${align ? ' applyAlignment="1">' + align + "</xf>" : "/>"}`,
    );
    return id;
  }

  toXml(): string {
    const numFmts = [...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`);
    return (
      XML_HEAD +
      `<styleSheet xmlns="${NS_MAIN}">` +
      (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join("")}</numFmts>` : "") +
      `<fonts count="${this.fontXml.length}">${this.fontXml.join("")}</fonts>` +
      `<fills count="${this.fillXml.length}">${this.fillXml.join("")}</fills>` +
      `<borders count="${this.borderXml.length}">${this.borderXml.join("")}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfXml.length}">${this.xfXml.join("")}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      "</styleSheet>"
    );
  }
}

function cellXml(wb: Workbook, sheet: string, row: number, col: number, style: number): string {
  const ref = addressToName({ row, col });
  const s = style ? ` s="${style}"` : "";
  const raw = wb.getRaw(sheet, row, col);
  const value = wb.getValue(sheet, row, col);
  if (raw === "") return `<c r="${ref}"${s}/>`;

  if (isFormula(raw)) {
    const f = `<f>${escapeXml(raw.slice(1))}</f>`;
    // The last computed result goes along so viewers that don't recalculate show it.
    if (typeof value === "number") return `<c r="${ref}"${s}>${f}<v>${value}</v></c>`;
    if (typeof value === "boolean") return `<c r="${ref}"${s} t="b">${f}<v>${value ? 1 : 0}</v></c>`;
    if (typeof value === "string") return `<c r="${ref}"${s} t="str">${f}<v>${escapeXml(value)}</v></c>`;
    if (isError(value) && EXCEL_ERRORS.has(value.code)) return `<c r="${ref}"${s} t="e">${f}<v>${value.code}</v></c>`;
    return `<c r="${ref}"${s}>${f}</c>`;
  }
  if (typeof value === "number") return `<c r="${ref}"${s}><v>${value}</v></c>`;
  if (typeof value === "boolean") return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  const text = typeof value === "string" ? value : raw;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
}

function sheetXml(wb: Workbook, sheet: string, styles: StyleTable): string {
  const rows = new Map<number, Set<number>>();
  const add = (key: string) => {
    const { row, col } = parseKey(key);
    let cols = rows.get(row);
    if (!cols) rows.set(row, (cols = new Set()));
    cols.add(col);
  };
  for (const key of wb.cellKeys(sheet)) add(key);
  for (const key of wb.formatKeys(sheet)) add(key);
  const heights = wb.rowHeights(sheet);
  for (const row of heights.keys()) if (!rows.has(row)) rows.set(row, new Set());

  const freeze = wb.getFreeze(sheet);
  let views = '<sheetViews><sheetView workbookViewId="0"';
  if (freeze.rows || freeze.cols) {
    const pane = freeze.rows && freeze.cols ? "bottomRight" : freeze.rows ? "bottomLeft" : "topRight";
    views +=
      `><pane${freeze.cols ? ` xSplit="${freeze.cols}"` : ""}${freeze.rows ? ` ySplit="${freeze.rows}"` : ""}` +
      ` topLeftCell="${addressToName({ row: freeze.rows, col: freeze.cols })}" activePane="${pane}" state="frozen"/>` +
      `<selection pane="${pane}"/></sheetView></sheetViews>`;
  } else {
    views += "/></sheetViews>";
  }

  const widths = [...wb.colWidths(sheet)].sort((a, b) => a[0] - b[0]);
  const cols = widths.length
    ? `<cols>${widths
        .map(([c, px]) => `<col min="${c + 1}" max="${c + 1}" width="${Math.max(1, (px - 5) / 7).toFixed(2)}" customWidth="1"/>`)
        .join("")}</cols>`
    : "";

  const rowXml = [...rows.keys()]
    .sort((a, b) => a - b)
    .map((row) => {
      const h = heights.get(row);
      const attrs = h ? ` ht="${((h * 3) / 4).toFixed(2)}" customHeight="1"` : "";
      const cells = [...rows.get(row)!]
        .sort((a, b) => a - b)
        .map((col) => cellXml(wb, sheet, row, col, styles.styleOf(wb.getFormat(sheet, row, col))))
        .join("");
      return `<row r="${row + 1}"${attrs}>${cells}</row>`;
    })
    .join("");

  const merges = wb.merges(sheet);
  const mergeXml = merges.length
    ? `<mergeCells count="${merges.length}">${merges
        .map((m) => `<mergeCell ref="${addressToName({ row: m.r0, col: m.c0 })}:${addressToName({ row: m.r1, col: m.c1 })}"/>`)
        .join("")}</mergeCells>`
    : "";

  const used = wb.usedBounds(sheet);
  const dimension = used ? `A1:${colToLetters(used.cols - 1)}${used.rows}` : "A1";
  return (
    XML_HEAD +
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    `<dimension ref="${dimension}"/>${views}<sheetFormatPr defaultRowHeight="18"/>${cols}` +
    `<sheetData>${rowXml}</sheetData>${mergeXml}</worksheet>`
  );
}

/** Excel limits sheet names to 31 characters without []:*?/\ and requires them to be unique. */
function excelSheetNames(names: string[]): string[] {
  const used = new Set<string>();
  return names.map((name, i) => {
    let base = name.replace(/[[\]:*?/\\]/g, "_").slice(0, 31).trim() || `Sheet${i + 1}`;
    let n = 2;
    while (used.has(base.toLowerCase())) base = `${name.slice(0, 28)} ${n++}`;
    used.add(base.toLowerCase());
    return base;
  });
}

/** Builds an .xlsx file from a workbook. */
export function writeXlsx(wb: Workbook): Uint8Array {
  const styles = new StyleTable();
  const sheets = wb.sheets;
  const names = excelSheetNames(sheets.map((s) => s.name));
  const files: Record<string, Uint8Array> = {};

  sheets.forEach((s, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(wb, s.id, styles));
  });

  files["xl/workbook.xml"] = strToU8(
    XML_HEAD +
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets>` +
      sheets.map((_, i) => `<sheet name="${escapeXml(names[i])}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
      // Ask Excel to recalculate on open, since Kexcel computes some things differently.
      '</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>',
  );
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    XML_HEAD +
      `<Relationships xmlns="${NS_PKG_REL}">` +
      sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL_TYPE}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
      `<Relationship Id="rId${sheets.length + 1}" Type="${REL_TYPE}/styles" Target="styles.xml"/>` +
      "</Relationships>",
  );
  files["xl/styles.xml"] = strToU8(styles.toXml());
  files["_rels/.rels"] = strToU8(
    XML_HEAD +
      `<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL_TYPE}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  files["[Content_Types].xml"] = strToU8(
    XML_HEAD +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets
        .map(
          (_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join("") +
      "</Types>",
  );

  return zipSync(files, { level: 6 });
}

export const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
