import { strFromU8, unzipSync } from "fflate";
import {
  addressToName,
  isError,
  mergeFormat,
  nameToAddress,
  parseLiteral,
  shiftFormula,
  Workbook,
  type CellFormat,
  type SheetJSON,
  type WorkbookJSON,
} from "@/lib/engine";
import { builtinFormat, formatFromCode, isDateFormat, serialToDateText } from "./numfmt";
import { child, children, parseXml, type XmlNode } from "./xml";

export interface ImportLimits {
  maxRows: number;
  maxCols: number;
}

export interface ImportResult {
  workbook: WorkbookJSON;
  /** Things that didn't carry over, in plain words for the user. */
  notes: string[];
}

interface Style {
  format: CellFormat | null;
  isDate: boolean;
}

const INVALID_SHEET_CHARS = /[!'[\]*?/\\:]/g;

/** "FFRRGGBB" (or "RRGGBB") to "#rrggbb"; theme and indexed colours aren't resolved. */
function rgb(node: XmlNode | undefined): string | undefined {
  const v = node?.attrs.rgb;
  if (!v || !/^[0-9a-f]{6,8}$/i.test(v)) return undefined;
  return "#" + v.slice(-6).toLowerCase();
}

const on = (node: XmlNode | undefined) => !!node && node.attrs.val !== "0" && node.attrs.val !== "false";

function readStyles(xml: XmlNode | undefined): Style[] {
  const sheet = child(xml, "styleSheet");
  if (!sheet) return [];
  const customFormats = new Map<number, string>();
  for (const f of children(child(sheet, "numFmts"), "numFmt")) {
    customFormats.set(Number(f.attrs.numFmtId), f.attrs.formatCode ?? "");
  }

  const fontNodes = children(child(sheet, "fonts"), "font");
  const defaultSize = Number(child(fontNodes[0], "sz")?.attrs.val ?? 11);
  const fonts: CellFormat[] = fontNodes.map((f) => {
    const out: CellFormat = {};
    if (on(child(f, "b"))) out.bold = true;
    if (on(child(f, "i"))) out.italic = true;
    if (child(f, "u") && child(f, "u")!.attrs.val !== "none") out.underline = true;
    if (on(child(f, "strike"))) out.strike = true;
    const size = Number(child(f, "sz")?.attrs.val);
    if (size && size !== defaultSize) out.fontSize = size;
    const color = rgb(child(f, "color"));
    if (color && color !== "#000000") out.color = color;
    const name = child(f, "name")?.attrs.val ?? "";
    if (/mono|courier|consolas/i.test(name)) out.fontFamily = "mono";
    else if (/times|georgia|serif|cambria|garamond/i.test(name) && !/sans/i.test(name)) out.fontFamily = "serif";
    return out;
  });

  const fills = children(child(sheet, "fills"), "fill").map((f) => {
    const pattern = child(f, "patternFill");
    return pattern?.attrs.patternType === "solid" ? rgb(child(pattern, "fgColor")) : undefined;
  });

  const borders = children(child(sheet, "borders"), "border").map((b) => {
    const side = (name: string) => !!child(b, name)?.attrs.style && child(b, name)!.attrs.style !== "none";
    return { top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") };
  });

  return children(child(sheet, "cellXfs"), "xf").map((xf) => {
    const numFmtId = Number(xf.attrs.numFmtId ?? 0);
    const code = customFormats.get(numFmtId) ?? builtinFormat(numFmtId);
    const isDate = isDateFormat(numFmtId, code);
    let f: CellFormat | null = { ...fonts[Number(xf.attrs.fontId ?? 0)] };
    if (!isDate) f = mergeFormat(f, formatFromCode(code));
    const fill = fills[Number(xf.attrs.fillId ?? 0)];
    if (fill) f = mergeFormat(f ?? {}, { fill });
    const border = borders[Number(xf.attrs.borderId ?? 0)];
    if (border && Object.values(border).some(Boolean)) f = mergeFormat(f ?? {}, { borders: border });
    const h = child(xf, "alignment")?.attrs.horizontal;
    const align = h === "center" || h === "centerContinuous" ? "center" : h === "right" ? "right" : h === "left" ? "left" : undefined;
    if (align) f = mergeFormat(f ?? {}, { align });
    return { format: f && Object.keys(f).length ? f : null, isDate };
  });
}

function sharedString(si: XmlNode): string {
  // Plain <t>, or rich text runs <r><t>…</t></r>; phonetic hints (<rPh>) are skipped.
  const direct = children(si, "t").map((t) => t.text).join("");
  const runs = children(si, "r")
    .map((r) => children(r, "t").map((t) => t.text).join(""))
    .join("");
  return direct + runs;
}

/** Excel writes newer functions with prefixes such as _xlfn. that Kexcel doesn't use. */
const cleanFormula = (f: string) => f.replace(/_xlfn\.|_xlws\.|_xludf\./g, "");

/** Text that would be read back as something else (a number, TRUE, a formula) is kept as text. */
function textRaw(text: string): string {
  if (text === "") return "";
  return typeof parseLiteral(text) === "string" && !text.startsWith("=") && !text.startsWith("'") ? text : `'${text}`;
}

function resolvePath(base: string, target: string) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

/** Reads an .xlsx file into Kexcel's workbook format. Throws if the file isn't a readable workbook. */
export function readXlsx(data: Uint8Array, limits: ImportLimits): ImportResult {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(data);
  } catch {
    throw new Error("This file isn't an Excel workbook (.xlsx). Older .xls files need to be saved as .xlsx first.");
  }
  const read = (path: string) => (files[path] ? parseXml(strFromU8(files[path])) : undefined);
  const relsOf = (path: string) => {
    const dir = path.split("/").slice(0, -1).join("/");
    const file = path.split("/").pop();
    const rels = new Map<string, string>();
    for (const r of children(child(read(`${dir ? dir + "/" : ""}_rels/${file}.rels`), "Relationships"), "Relationship")) {
      rels.set(r.attrs.Id, resolvePath(path, r.attrs.Target ?? ""));
    }
    return rels;
  };

  const rootRels = children(child(read("_rels/.rels"), "Relationships"), "Relationship");
  const officeDoc = rootRels.find((r) => /officeDocument$/.test(r.attrs.Type ?? ""))?.attrs.Target;
  const workbookPath = officeDoc ? resolvePath("", officeDoc) : "xl/workbook.xml";
  const workbookXml = child(read(workbookPath), "workbook");
  if (!workbookXml) throw new Error("This file isn't an Excel workbook (.xlsx).");
  const rels = relsOf(workbookPath);

  const relTargets = [...rels.values()];
  const stringsPath = relTargets.find((t) => /sharedStrings\.xml$/i.test(t));
  const stylesPath = relTargets.find((t) => /styles\.xml$/i.test(t));
  const strings = children(child(stringsPath ? read(stringsPath) : undefined, "sst"), "si").map(sharedString);
  const styles = readStyles(stylesPath ? read(stylesPath) : undefined);

  const notes: string[] = [];
  let merged = 0;
  let clipped = false;
  const usedNames = new Set<string>();
  const sheets: SheetJSON[] = [];

  for (const [index, s] of children(child(workbookXml, "sheets"), "sheet").entries()) {
    const target = rels.get(s.attrs["r:id"] ?? "");
    const ws = child(target ? read(target) : undefined, "worksheet");
    if (!ws) continue; // chart sheets and the like

    let name = (s.attrs.name ?? `Sheet${index + 1}`).replace(INVALID_SHEET_CHARS, "_").slice(0, 100).trim() || `Sheet${index + 1}`;
    for (let n = 2; usedNames.has(name.toLowerCase()); n++) name = `${name.slice(0, 95)} ${n}`;
    usedNames.add(name.toLowerCase());

    const cells: Record<string, string> = {};
    const formats: Record<string, CellFormat> = {};
    const colWidths: Record<string, number> = {};
    const rowHeights: Record<string, number> = {};
    const shared = new Map<string, { formula: string; row: number; col: number }>();

    for (const col of children(child(ws, "cols"), "col")) {
      if (col.attrs.width === undefined || col.attrs.customWidth === "0") continue;
      const px = Math.max(16, Math.round(Number(col.attrs.width) * 7 + 5));
      const max = Math.min(Number(col.attrs.max), limits.maxCols);
      for (let c = Number(col.attrs.min) - 1; c < max; c++) colWidths[c] = px;
    }

    for (const row of children(child(ws, "sheetData"), "row")) {
      const r = Number(row.attrs.r) - 1;
      if (row.attrs.customHeight === "1" && row.attrs.ht && r < limits.maxRows) {
        rowHeights[r] = Math.max(12, Math.round((Number(row.attrs.ht) * 4) / 3));
      }
      for (const c of children(row, "c")) {
        const addr = nameToAddress(c.attrs.r ?? "");
        if (!addr) continue;
        if (addr.row >= limits.maxRows || addr.col >= limits.maxCols) {
          clipped = true;
          continue;
        }
        const name = addressToName(addr);
        const style = styles[Number(c.attrs.s ?? 0)];
        if (style?.format) formats[name] = style.format;

        const f = child(c, "f");
        const v = child(c, "v")?.text ?? "";
        let raw = "";
        if (f && (f.text.trim() || f.attrs.t === "shared")) {
          if (f.text.trim()) {
            raw = "=" + cleanFormula(f.text.trim());
            if (f.attrs.t === "shared" && f.attrs.si !== undefined) shared.set(f.attrs.si, { formula: raw, ...addr });
          } else {
            const master = shared.get(f.attrs.si ?? "");
            if (master) raw = shiftFormula(master.formula, addr.row - master.row, addr.col - master.col);
          }
        } else {
          switch (c.attrs.t) {
            case "s":
              raw = textRaw(strings[Number(v)] ?? "");
              break;
            case "inlineStr":
              raw = textRaw(sharedString(child(c, "is") ?? { name: "", attrs: {}, children: [], text: "" }));
              break;
            case "str":
              raw = textRaw(v);
              break;
            case "b":
              raw = v === "1" ? "TRUE" : "FALSE";
              break;
            case "e":
              raw = `'${v}`;
              break;
            default:
              if (v !== "") raw = style?.isDate ? textRaw(serialToDateText(Number(v))) : v;
          }
        }
        if (raw !== "") cells[name] = raw;
      }
    }

    merged += children(child(ws, "mergeCells"), "mergeCell").length;

    const pane = child(child(child(ws, "sheetViews"), "sheetView"), "pane");
    const frozen = pane && (pane.attrs.state === "frozen" || pane.attrs.state === "frozenSplit");
    const freeze = frozen
      ? { rows: Math.round(Number(pane.attrs.ySplit ?? 0)), cols: Math.round(Number(pane.attrs.xSplit ?? 0)) }
      : undefined;

    sheets.push({ id: `s${sheets.length + 1}`, name, cells, formats, colWidths, rowHeights, ...(freeze ? { freeze } : {}) });
  }

  if (sheets.length === 0) throw new Error("This workbook has no sheets Kexcel can open.");

  if (clipped) notes.push(`Only the first ${limits.maxRows.toLocaleString()} rows and ${limits.maxCols} columns of each sheet were imported.`);
  if (merged) notes.push(`${merged} merged cell range${merged === 1 ? " was" : "s were"} unmerged; Kexcel doesn't merge cells yet.`);

  // Formulas Kexcel can't run yet, such as whole-column ranges or unknown functions.
  const workbook: WorkbookJSON = { sheets };
  const check = Workbook.fromJSON(workbook);
  let unsupported = 0;
  for (const s of check.sheets) {
    for (const key of check.cellKeys(s.id)) {
      const [row, col] = key.split(",").map(Number);
      const v = check.getValue(s.id, row, col);
      if (isError(v) && (v.code === "#ERROR!" || v.code === "#NAME?")) unsupported++;
    }
  }
  if (unsupported) {
    notes.push(
      `${unsupported} formula${unsupported === 1 ? " uses" : "s use"} something Kexcel doesn't support yet and show${unsupported === 1 ? "s" : ""} an error.`,
    );
  }
  return { workbook, notes };
}
