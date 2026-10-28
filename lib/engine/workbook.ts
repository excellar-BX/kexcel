import { addressToName, cellKey, nameToAddress, parseKey } from "./address";
import { evaluateFormula, type CellReader } from "./evaluate";
import { sameFormat, type CellFormat } from "./format";
import { VOLATILE_FUNCTIONS } from "./functions";
import { impliedDateKind, isFormula, parseLiteral } from "./literal";
import { collectReferences, parseFormula, type Node } from "./parser";
import { adjustForStructureChange, renameSheetInFormula, type Axis } from "./references";
import { CellError, type Scalar } from "./types";

interface Cell {
  raw: string;
  ast: Node | null;
  value: Scalar;
  /** Global keys ("sheetId!row,col") of the cells this cell's formula reads. */
  deps: Set<string>;
  /** Uses TODAY() or NOW(), so it recalculates on every change. */
  volatile?: boolean;
}

/** An inclusive block of cells, used for merged ranges. */
export interface CellRange {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

interface SheetState {
  id: string;
  name: string;
  /** Keyed by "row,col". */
  cells: Map<string, Cell>;
  formats: Map<string, CellFormat>;
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  freeze: Freeze;
  merges: CellRange[];
}

/** Rows at the top and columns at the left that stay in place while scrolling. */
export interface Freeze {
  rows: number;
  cols: number;
}

export interface SheetInfo {
  id: string;
  name: string;
}

export interface CellUpdate {
  row: number;
  col: number;
  raw: string;
}

/** A change to one cell. Fields left undefined are untouched; `format: null` clears the format. */
export interface CellPatch {
  sheet: string;
  row: number;
  col: number;
  raw?: string;
  format?: CellFormat | null;
}

export interface SheetJSON {
  id: string;
  name: string;
  /** Raw input keyed by cell name, e.g. { A1: "=SUM(B1:B3)" }. */
  cells: Record<string, string>;
  formats?: Record<string, CellFormat>;
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
  freeze?: Freeze;
  /** Merged ranges such as "A1:C1". */
  merges?: string[];
}

export interface WorkbookJSON {
  sheets: SheetJSON[];
}

const globalKey = (sheet: string, row: number, col: number) => `${sheet}!${row},${col}`;

const INVALID_NAME_CHARS = /[!'[\]*?/\\:]/;

/**
 * A spreadsheet document: ordered sheets holding what the user typed and
 * how it looks, plus one dependency graph across all sheets so that only
 * the cells affected by a change are recalculated.
 */
export class Workbook {
  private list: SheetState[] = [];
  private byId = new Map<string, SheetState>();
  private byName = new Map<string, SheetState>();
  /** Reverse edges: global key -> global keys of formulas that read it. */
  private dependents = new Map<string, Set<string>>();
  private nextId = 1;

  static withSheet(name = "Sheet1"): Workbook {
    const wb = new Workbook();
    wb.addSheet(name);
    return wb;
  }

  static fromJSON(json: WorkbookJSON): Workbook {
    const wb = new Workbook();
    wb.load(json);
    return wb;
  }

  /** Replaces the whole document, e.g. to restore an undo snapshot. */
  load(json: WorkbookJSON) {
    this.list = [];
    this.byId.clear();
    this.byName.clear();
    this.nextId = 1;
    for (const s of json.sheets) {
      const sheet = this.createSheet(s.name, s.id);
      for (const [name, raw] of Object.entries(s.cells ?? {})) {
        const a = nameToAddress(name);
        if (a && typeof raw === "string" && raw !== "") {
          sheet.cells.set(cellKey(a.row, a.col), { raw, ast: null, value: null, deps: new Set() });
        }
      }
      for (const [name, f] of Object.entries(s.formats ?? {})) {
        const a = nameToAddress(name);
        if (a && f && typeof f === "object") sheet.formats.set(cellKey(a.row, a.col), f);
      }
      for (const [c, w] of Object.entries(s.colWidths ?? {})) sheet.colWidths.set(Number(c), w);
      for (const [r, h] of Object.entries(s.rowHeights ?? {})) sheet.rowHeights.set(Number(r), h);
      if (s.freeze) sheet.freeze = { rows: Math.max(0, s.freeze.rows | 0), cols: Math.max(0, s.freeze.cols | 0) };
      for (const ref of s.merges ?? []) {
        const [a, b] = ref.split(":").map(nameToAddress);
        if (a && b) sheet.merges.push({ r0: a.row, c0: a.col, r1: b.row, c1: b.col });
      }
    }
    if (this.list.length === 0) this.createSheet("Sheet1");
    // Typed dates without a number format (from CSV files, or saved before dates
    // were values) would show as day numbers, so they get a date format.
    for (const sheet of this.list) {
      for (const [key, cell] of sheet.cells) {
        if (isFormula(cell.raw) || sheet.formats.get(key)?.numFmt) continue;
        const kind = impliedDateKind(cell.raw);
        if (kind) sheet.formats.set(key, { ...sheet.formats.get(key), numFmt: kind });
      }
    }
    this.rebuild();
  }

  toJSON(): WorkbookJSON {
    return {
      sheets: this.list.map((s) => {
        const cells: Record<string, string> = {};
        for (const [key, cell] of s.cells) cells[addressToName(parseKey(key))] = cell.raw;
        const formats: Record<string, CellFormat> = {};
        for (const [key, f] of s.formats) formats[addressToName(parseKey(key))] = f;
        return {
          id: s.id,
          name: s.name,
          cells,
          formats,
          colWidths: Object.fromEntries(s.colWidths),
          rowHeights: Object.fromEntries(s.rowHeights),
          ...(s.freeze.rows || s.freeze.cols ? { freeze: { ...s.freeze } } : {}),
          ...(s.merges.length
            ? {
                merges: s.merges.map(
                  (m) => `${addressToName({ row: m.r0, col: m.c0 })}:${addressToName({ row: m.r1, col: m.c1 })}`,
                ),
              }
            : {}),
        };
      }),
    };
  }

  // ---- Sheets ----------------------------------------------------------------

  get sheets(): SheetInfo[] {
    return this.list.map(({ id, name }) => ({ id, name }));
  }

  hasSheet(id: string) {
    return this.byId.has(id);
  }

  sheetName(id: string): string {
    return this.byId.get(id)?.name ?? "";
  }

  /** Returns an error message, or null when the name can be used. */
  validateSheetName(name: string, exceptId?: string): string | null {
    const trimmed = name.trim();
    if (!trimmed) return "A sheet name can't be empty.";
    if (trimmed.length > 100) return "A sheet name can be at most 100 characters.";
    if (INVALID_NAME_CHARS.test(trimmed)) return "A sheet name can't contain ! ' [ ] * ? / \\ or :";
    const existing = this.byName.get(trimmed.toLowerCase());
    if (existing && existing.id !== exceptId) return `There's already a sheet called "${trimmed}".`;
    return null;
  }

  /** Adds a sheet after `afterId` (or at the end) and returns its id. */
  addSheet(name?: string, afterId?: string): string {
    let finalName = name?.trim() ?? "";
    if (!finalName || this.validateSheetName(finalName)) {
      let n = this.list.length + 1;
      while (this.byName.has(`sheet${n}`)) n++;
      finalName = `Sheet${n}`;
    }
    const sheet = this.createSheet(finalName);
    if (afterId) {
      this.list.pop();
      this.list.splice(this.indexOf(afterId) + 1, 0, sheet);
    }
    this.rebuild();
    return sheet.id;
  }

  duplicateSheet(id: string): string {
    const src = this.sheet(id);
    let n = 1;
    let name = `${src.name} (copy)`;
    while (this.byName.has(name.toLowerCase())) name = `${src.name} (copy ${++n})`;
    const copy = this.createSheet(name);
    this.list.pop();
    this.list.splice(this.indexOf(id) + 1, 0, copy);
    for (const [k, c] of src.cells) copy.cells.set(k, { raw: c.raw, ast: null, value: null, deps: new Set() });
    for (const [k, f] of src.formats) copy.formats.set(k, f);
    src.colWidths.forEach((w, c) => copy.colWidths.set(c, w));
    src.rowHeights.forEach((h, r) => copy.rowHeights.set(r, h));
    copy.freeze = { ...src.freeze };
    copy.merges = src.merges.map((m) => ({ ...m }));
    this.rebuild();
    return copy.id;
  }

  /** Renames a sheet and every formula reference to it. Throws on an invalid name. */
  renameSheet(id: string, name: string) {
    const sheet = this.sheet(id);
    const trimmed = name.trim();
    const error = this.validateSheetName(trimmed, id);
    if (error) throw new Error(error);
    if (trimmed === sheet.name) return;
    const old = sheet.name;
    this.forEachFormula((cell) => {
      cell.raw = renameSheetInFormula(cell.raw, old, trimmed);
    });
    this.byName.delete(old.toLowerCase());
    sheet.name = trimmed;
    this.byName.set(trimmed.toLowerCase(), sheet);
    this.rebuild();
  }

  /** Removes a sheet; formulas that pointed at it show #REF!. The last sheet can't be removed. */
  removeSheet(id: string) {
    if (this.list.length <= 1) throw new Error("A workbook needs at least one sheet.");
    const sheet = this.sheet(id);
    this.list.splice(this.indexOf(id), 1);
    this.byId.delete(id);
    this.byName.delete(sheet.name.toLowerCase());
    this.rebuild();
  }

  moveSheet(id: string, toIndex: number) {
    const from = this.indexOf(id);
    const [sheet] = this.list.splice(from, 1);
    this.list.splice(Math.max(0, Math.min(toIndex, this.list.length)), 0, sheet);
  }

  // ---- Cells -----------------------------------------------------------------

  getRaw(sheet: string, row: number, col: number): string {
    return this.byId.get(sheet)?.cells.get(cellKey(row, col))?.raw ?? "";
  }

  getValue(sheet: string, row: number, col: number): Scalar {
    return this.byId.get(sheet)?.cells.get(cellKey(row, col))?.value ?? null;
  }

  getFormat(sheet: string, row: number, col: number): CellFormat | undefined {
    return this.byId.get(sheet)?.formats.get(cellKey(row, col));
  }

  /** "row,col" keys of the non-empty cells in a sheet. */
  cellKeys(sheet: string): IterableIterator<string> {
    return (this.byId.get(sheet)?.cells ?? new Map<string, Cell>()).keys();
  }

  /** "row,col" keys of the formatted cells in a sheet. */
  formatKeys(sheet: string): IterableIterator<string> {
    return (this.byId.get(sheet)?.formats ?? new Map<string, CellFormat>()).keys();
  }

  /** Size of the area holding content or formatting, or null when the sheet is blank. */
  usedBounds(sheet: string): { rows: number; cols: number } | null {
    const s = this.byId.get(sheet);
    if (!s || (s.cells.size === 0 && s.formats.size === 0)) return null;
    let rows = 0;
    let cols = 0;
    for (const key of [...s.cells.keys(), ...s.formats.keys()]) {
      const { row, col } = parseKey(key);
      rows = Math.max(rows, row + 1);
      cols = Math.max(cols, col + 1);
    }
    return { rows, cols };
  }

  setCell(sheet: string, row: number, col: number, raw: string) {
    this.setCells(sheet, [{ row, col, raw }]);
  }

  /** Applies many edits at once (paste, clear, load) with a single recalculation. */
  setCells(sheetId: string, updates: CellUpdate[]) {
    const sheet = this.sheet(sheetId);
    const changed = new Set<string>();
    for (const { row, col, raw } of updates) {
      const key = cellKey(row, col);
      const gkey = globalKey(sheetId, row, col);
      const old = sheet.cells.get(key);
      if (old) this.unlink(gkey, old);
      if (raw === "") {
        sheet.cells.delete(key);
      } else {
        const cell = this.compile(sheetId, raw);
        sheet.cells.set(key, cell);
        this.link(gkey, cell);
      }
      changed.add(gkey);
    }
    this.recalculate(changed);
  }

  setFormat(sheet: string, row: number, col: number, format: CellFormat | null) {
    const formats = this.sheet(sheet).formats;
    const key = cellKey(row, col);
    if (format && Object.keys(format).length) formats.set(key, format);
    else formats.delete(key);
  }

  /** Applies cell patches and returns the patches that undo them. */
  apply(patches: CellPatch[]): CellPatch[] {
    const inverse: CellPatch[] = [];
    const raws = new Map<string, CellUpdate[]>();
    for (const p of patches) {
      if (!this.byId.has(p.sheet)) continue;
      const inv: CellPatch = { sheet: p.sheet, row: p.row, col: p.col };
      if (p.raw !== undefined) {
        inv.raw = this.getRaw(p.sheet, p.row, p.col);
        let list = raws.get(p.sheet);
        if (!list) raws.set(p.sheet, (list = []));
        list.push({ row: p.row, col: p.col, raw: p.raw });
      }
      if (p.format !== undefined) {
        inv.format = this.getFormat(p.sheet, p.row, p.col) ?? null;
        this.setFormat(p.sheet, p.row, p.col, p.format);
      }
      inverse.push(inv);
    }
    for (const [sheet, updates] of raws) this.setCells(sheet, updates);
    return inverse.reverse();
  }

  /** Drops patches that would not change anything. */
  effective(patches: CellPatch[]): CellPatch[] {
    return patches.filter(
      (p) =>
        (p.raw !== undefined && p.raw !== this.getRaw(p.sheet, p.row, p.col)) ||
        (p.format !== undefined && !sameFormat(p.format, this.getFormat(p.sheet, p.row, p.col))),
    );
  }

  // ---- Sizes -----------------------------------------------------------------

  /** Custom column widths in pixels, keyed by column. */
  colWidths(sheet: string): ReadonlyMap<number, number> {
    return this.byId.get(sheet)?.colWidths ?? new Map();
  }

  /** Custom row heights in pixels, keyed by row. */
  rowHeights(sheet: string): ReadonlyMap<number, number> {
    return this.byId.get(sheet)?.rowHeights ?? new Map();
  }

  setColWidth(sheet: string, col: number, width: number | undefined) {
    const m = this.sheet(sheet).colWidths;
    if (width === undefined) m.delete(col);
    else m.set(col, width);
  }

  setRowHeight(sheet: string, row: number, height: number | undefined) {
    const m = this.sheet(sheet).rowHeights;
    if (height === undefined) m.delete(row);
    else m.set(row, height);
  }

  getFreeze(sheet: string): Freeze {
    return this.byId.get(sheet)?.freeze ?? { rows: 0, cols: 0 };
  }

  setFreeze(sheet: string, freeze: Freeze) {
    this.sheet(sheet).freeze = { rows: Math.max(0, freeze.rows | 0), cols: Math.max(0, freeze.cols | 0) };
  }

  /** Merged ranges on a sheet. */
  merges(sheet: string): readonly CellRange[] {
    return this.byId.get(sheet)?.merges ?? [];
  }

  /** The merged range covering a cell, if any. */
  mergeAt(sheet: string, row: number, col: number): CellRange | undefined {
    return this.merges(sheet).find((m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1);
  }

  /** Replaces a sheet's merged ranges; single cells and overlaps are dropped. */
  setMerges(sheet: string, merges: CellRange[]) {
    const kept: CellRange[] = [];
    for (const m of merges) {
      if (m.r0 === m.r1 && m.c0 === m.c1) continue;
      const overlaps = kept.some((k) => m.r0 <= k.r1 && m.r1 >= k.r0 && m.c0 <= k.c1 && m.c1 >= k.c0);
      if (!overlaps) kept.push({ ...m });
    }
    this.sheet(sheet).merges = kept;
  }

  // ---- Structure -------------------------------------------------------------

  /** Inserts `count` blank rows or columns before index `at`. */
  insert(sheet: string, axis: Axis, at: number, count: number) {
    this.restructure(sheet, axis, at, count);
  }

  /** Deletes `count` rows or columns starting at index `at`. */
  remove(sheet: string, axis: Axis, at: number, count: number) {
    this.restructure(sheet, axis, at, -count);
  }

  private restructure(sheetId: string, axis: Axis, at: number, count: number) {
    if (count === 0) return;
    const target = this.sheet(sheetId);
    const end = at - count; // first index after a deleted block
    /** New index for a row or column, or null if it was deleted. */
    const move = (v: number): number | null => {
      if (v < at) return v;
      if (count > 0) return v + count;
      return v < end ? null : v + count;
    };
    const moveKey = (key: string) => {
      const { row, col } = parseKey(key);
      const v = move(axis === "row" ? row : col);
      if (v === null) return null;
      return axis === "row" ? cellKey(v, col) : cellKey(row, v);
    };
    const remap = <T>(m: Map<string, T>) => {
      const out = new Map<string, T>();
      for (const [k, v] of m) {
        const nk = moveKey(k);
        if (nk !== null) out.set(nk, v);
      }
      return out;
    };
    const remapSizes = (m: Map<number, number>) => {
      const out = new Map<number, number>();
      for (const [k, v] of m) {
        const nk = move(k);
        if (nk !== null) out.set(nk, v);
      }
      return out;
    };

    target.cells = remap(target.cells);
    target.formats = remap(target.formats);
    if (axis === "row") target.rowHeights = remapSizes(target.rowHeights);
    else target.colWidths = remapSizes(target.colWidths);

    // Merged ranges move, grow or shrink with their rows and columns.
    const lo = axis === "row" ? "r0" : "c0";
    const hi = axis === "row" ? "r1" : "c1";
    target.merges = target.merges.flatMap((m) => {
      let a = m[lo];
      let b = m[hi];
      if (count > 0) {
        if (a >= at) a += count;
        if (b >= at) b += count;
      } else {
        a = a < at ? a : a >= end ? a + count : at;
        b = b < at ? b : b >= end ? b + count : at - 1;
        if (b < a) return [];
      }
      const moved = { ...m, [lo]: a, [hi]: b };
      return moved.r0 === moved.r1 && moved.c0 === moved.c1 ? [] : [moved];
    });

    // Rows or columns added or removed inside the frozen area grow or shrink it.
    const key = axis === "row" ? "rows" : "cols";
    const frozen = target.freeze[key];
    if (at < frozen) {
      target.freeze = {
        ...target.freeze,
        [key]: count > 0 ? frozen + count : Math.max(at, frozen - Math.min(-count, frozen - at)),
      };
    }

    const targetName = target.name.toLowerCase();
    this.forEachFormula((cell, owner) => {
      cell.raw = adjustForStructureChange(cell.raw, axis, at, count, (ctx) =>
        ctx.sheet === undefined ? owner.id === sheetId : ctx.sheet.toLowerCase() === targetName,
      );
    });
    this.rebuild();
  }

  // ---- Internals -------------------------------------------------------------

  private sheet(id: string): SheetState {
    const s = this.byId.get(id);
    if (!s) throw new Error(`Unknown sheet ${id}`);
    return s;
  }

  private indexOf(id: string) {
    return this.list.findIndex((s) => s.id === id);
  }

  private createSheet(name: string, id?: string): SheetState {
    let sheetId = id;
    if (!sheetId || this.byId.has(sheetId) || sheetId.includes("!")) {
      while (this.byId.has(`s${this.nextId}`)) this.nextId++;
      sheetId = `s${this.nextId++}`;
    }
    const numeric = /^s(\d+)$/.exec(sheetId);
    if (numeric) this.nextId = Math.max(this.nextId, Number(numeric[1]) + 1);

    let finalName = name;
    for (let n = 2; this.byName.has(finalName.toLowerCase()); n++) finalName = `${name} ${n}`;
    const sheet: SheetState = {
      id: sheetId,
      name: finalName,
      cells: new Map(),
      formats: new Map(),
      colWidths: new Map(),
      rowHeights: new Map(),
      freeze: { rows: 0, cols: 0 },
      merges: [],
    };
    this.list.push(sheet);
    this.byId.set(sheetId, sheet);
    this.byName.set(finalName.toLowerCase(), sheet);
    return sheet;
  }

  private forEachFormula(fn: (cell: Cell, owner: SheetState) => void) {
    for (const sheet of this.list) {
      for (const cell of sheet.cells.values()) if (isFormula(cell.raw)) fn(cell, sheet);
    }
  }

  private compile(sheetId: string, raw: string): Cell {
    if (!isFormula(raw)) {
      return { raw, ast: null, value: parseLiteral(raw), deps: new Set() };
    }
    try {
      const ast = parseFormula(raw.slice(1));
      const deps = new Set<string>();
      for (const ref of collectReferences(ast)) {
        const target = ref.sheet === undefined ? sheetId : this.byName.get(ref.sheet.toLowerCase())?.id;
        if (target) deps.add(globalKey(target, ref.row, ref.col));
      }
      return { raw, ast, value: null, deps, volatile: callsAny(ast, VOLATILE_FUNCTIONS) };
    } catch (e) {
      const message = e instanceof Error ? e.message : "Invalid formula";
      return { raw, ast: null, value: new CellError("#ERROR!", message), deps: new Set() };
    }
  }

  private cellAt(gkey: string): Cell | undefined {
    const bang = gkey.indexOf("!");
    return this.byId.get(gkey.slice(0, bang))?.cells.get(gkey.slice(bang + 1));
  }

  private link(gkey: string, cell: Cell) {
    for (const dep of cell.deps) {
      let set = this.dependents.get(dep);
      if (!set) this.dependents.set(dep, (set = new Set()));
      set.add(gkey);
    }
  }

  private unlink(gkey: string, cell: Cell) {
    for (const dep of cell.deps) {
      const set = this.dependents.get(dep);
      set?.delete(gkey);
      if (set?.size === 0) this.dependents.delete(dep);
    }
  }

  /** Recompiles every cell and recalculates everything, after sheets or structure change. */
  private rebuild() {
    this.dependents.clear();
    const all = new Set<string>();
    for (const sheet of this.list) {
      for (const [key, old] of sheet.cells) {
        const cell = this.compile(sheet.id, old.raw);
        sheet.cells.set(key, cell);
        const gkey = `${sheet.id}!${key}`;
        this.link(gkey, cell);
        all.add(gkey);
      }
    }
    this.recalculate(all);
  }

  private readerFor(sheetId: string): CellReader {
    return (row, col, sheetName) => {
      const target = sheetName === undefined ? this.byId.get(sheetId) : this.byName.get(sheetName.toLowerCase());
      if (!target) throw new CellError("#REF!", `There's no sheet called "${sheetName}"`);
      return target.cells.get(cellKey(row, col))?.value ?? null;
    };
  }

  private recalculate(seeds: Set<string>) {
    // Everything downstream of the edited cells needs recomputing, plus TODAY() and NOW().
    const affected = new Set(seeds);
    for (const sheet of this.list) {
      for (const [key, cell] of sheet.cells) if (cell.volatile) affected.add(`${sheet.id}!${key}`);
    }
    const stack = [...affected];
    while (stack.length) {
      const key = stack.pop()!;
      for (const d of this.dependents.get(key) ?? []) {
        if (!affected.has(d)) {
          affected.add(d);
          stack.push(d);
        }
      }
    }

    // Kahn's algorithm: evaluate a cell once all of its affected inputs are done.
    const pending = new Map<string, number>();
    const ready: string[] = [];
    for (const key of affected) {
      let count = 0;
      for (const dep of this.cellAt(key)?.deps ?? []) if (affected.has(dep)) count++;
      pending.set(key, count);
      if (count === 0) ready.push(key);
    }

    while (ready.length) {
      const key = ready.pop()!;
      pending.delete(key);
      const cell = this.cellAt(key);
      if (cell?.ast) cell.value = evaluateFormula(cell.ast, this.readerFor(key.slice(0, key.indexOf("!"))));
      for (const d of this.dependents.get(key) ?? []) {
        const left = pending.get(d);
        if (left === undefined) continue;
        pending.set(d, left - 1);
        if (left === 1) ready.push(d);
      }
    }

    // Whatever never became ready sits on, or downstream of, a circular reference.
    for (const key of pending.keys()) {
      const cell = this.cellAt(key);
      if (cell?.ast) cell.value = new CellError("#CYCLE!", "Circular reference");
    }
  }
}

function callsAny(node: Node, names: Set<string>): boolean {
  switch (node.type) {
    case "call":
      return names.has(node.name) || node.args.some((a) => callsAny(a, names));
    case "unary":
    case "percent":
      return callsAny(node.arg, names);
    case "binary":
      return callsAny(node.left, names) || callsAny(node.right, names);
    default:
      return false;
  }
}
