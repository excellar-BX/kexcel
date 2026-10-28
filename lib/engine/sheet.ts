import { addressToName, cellKey, colToLetters, nameToAddress, parseKey } from "./address";
import { formatNumber, parseNumberLiteral } from "./coerce";
import { evaluateFormula } from "./evaluate";
import { collectReferences, parseFormula, type Node } from "./parser";
import { tokenize } from "./tokenizer";
import { CellError, isError, type Scalar } from "./types";

interface Cell {
  raw: string;
  ast: Node | null;
  value: Scalar;
  /** Keys of the cells this cell's formula reads. */
  deps: Set<string>;
}

export interface CellUpdate {
  row: number;
  col: number;
  raw: string;
}

/** Interprets typed input that is not a formula: numbers, booleans, or text. */
function parseLiteral(raw: string): Scalar {
  if (raw === "") return null;
  if (raw.startsWith("'")) return raw.slice(1);
  const n = parseNumberLiteral(raw);
  if (n !== null) return n;
  const upper = raw.trim().toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return raw;
}

export const isFormula = (raw: string) => raw.startsWith("=") && raw.length > 1;

/**
 * A single worksheet: stores what the user typed, keeps a dependency graph
 * between cells, and recalculates only the cells affected by each change.
 */
export class Sheet {
  private cells = new Map<string, Cell>();
  /** Reverse edges: key -> cells whose formulas read it. */
  private dependents = new Map<string, Set<string>>();

  getRaw(row: number, col: number): string {
    return this.cells.get(cellKey(row, col))?.raw ?? "";
  }

  getValue(row: number, col: number): Scalar {
    return this.cells.get(cellKey(row, col))?.value ?? null;
  }

  setCell(row: number, col: number, raw: string) {
    this.setCells([{ row, col, raw }]);
  }

  /** Applies many edits at once (paste, clear, load) with a single recalculation. */
  setCells(updates: CellUpdate[]) {
    const changed = new Set<string>();
    for (const { row, col, raw } of updates) {
      const key = cellKey(row, col);
      this.unlink(key);
      if (raw === "") {
        this.cells.delete(key);
      } else {
        this.cells.set(key, this.compile(raw));
        this.link(key);
      }
      changed.add(key);
    }
    this.recalculate(changed);
  }

  /** Keys of all non-empty cells. */
  keys(): IterableIterator<string> {
    return this.cells.keys();
  }

  /** The bottom-right corner of the used area, or null when the sheet is empty. */
  usedBounds(): { rows: number; cols: number } | null {
    if (this.cells.size === 0) return null;
    let rows = 0;
    let cols = 0;
    for (const key of this.cells.keys()) {
      const { row, col } = parseKey(key);
      rows = Math.max(rows, row + 1);
      cols = Math.max(cols, col + 1);
    }
    return { rows, cols };
  }

  toJSON(): Record<string, string> {
    const out: Record<string, string> = {};
    for (const [key, cell] of this.cells) out[addressToName(parseKey(key))] = cell.raw;
    return out;
  }

  static fromJSON(data: Record<string, string>): Sheet {
    const sheet = new Sheet();
    const updates: CellUpdate[] = [];
    for (const [name, raw] of Object.entries(data)) {
      const addr = nameToAddress(name);
      if (addr && typeof raw === "string") updates.push({ ...addr, raw });
    }
    sheet.setCells(updates);
    return sheet;
  }

  private compile(raw: string): Cell {
    if (!isFormula(raw)) {
      return { raw, ast: null, value: parseLiteral(raw), deps: new Set() };
    }
    try {
      const ast = parseFormula(raw.slice(1));
      const deps = new Set(collectReferences(ast).map((a) => cellKey(a.row, a.col)));
      return { raw, ast, value: null, deps };
    } catch (e) {
      const message = e instanceof Error ? e.message : "Invalid formula";
      return { raw, ast: null, value: new CellError("#ERROR!", message), deps: new Set() };
    }
  }

  private link(key: string) {
    for (const dep of this.cells.get(key)!.deps) {
      let set = this.dependents.get(dep);
      if (!set) this.dependents.set(dep, (set = new Set()));
      set.add(key);
    }
  }

  private unlink(key: string) {
    const cell = this.cells.get(key);
    if (!cell) return;
    for (const dep of cell.deps) {
      const set = this.dependents.get(dep);
      set?.delete(key);
      if (set?.size === 0) this.dependents.delete(dep);
    }
  }

  private recalculate(seeds: Set<string>) {
    // Everything downstream of the edited cells needs recomputing.
    const affected = new Set(seeds);
    const stack = [...seeds];
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
      for (const dep of this.cells.get(key)?.deps ?? []) if (affected.has(dep)) count++;
      pending.set(key, count);
      if (count === 0) ready.push(key);
    }

    const read = (row: number, col: number) => this.getValue(row, col);
    while (ready.length) {
      const key = ready.pop()!;
      pending.delete(key);
      const cell = this.cells.get(key);
      if (cell?.ast) cell.value = evaluateFormula(cell.ast, read);
      for (const d of this.dependents.get(key) ?? []) {
        const left = pending.get(d);
        if (left === undefined) continue;
        pending.set(d, left - 1);
        if (left === 1) ready.push(d);
      }
    }

    // Whatever never became ready sits on, or downstream of, a circular reference.
    for (const key of pending.keys()) {
      const cell = this.cells.get(key);
      if (cell?.ast) cell.value = new CellError("#CYCLE!", "Circular reference");
    }
  }
}

/** Text shown in a cell for its computed value. */
export function displayValue(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "number") return formatNumber(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isError(v)) return v.code;
  return v;
}

/**
 * Moves the relative references in a formula, as when it is pasted
 * dRow rows and dCol columns away from where it was copied. References
 * pushed off the sheet become #REF!. Anchored ($) parts stay put.
 */
export function shiftFormula(raw: string, dRow: number, dCol: number): string {
  if (!isFormula(raw) || (dRow === 0 && dCol === 0)) return raw;
  const body = raw.slice(1);
  let tokens;
  try {
    tokens = tokenize(body);
  } catch {
    return raw;
  }
  let out = "";
  let last = 0;
  for (const t of tokens) {
    if (t.type !== "ref") continue;
    const row = t.rowAbs ? t.row : t.row + dRow;
    const col = t.colAbs ? t.col : t.col + dCol;
    out += body.slice(last, t.start);
    out +=
      row < 0 || col < 0
        ? "#REF!"
        : `${t.colAbs ? "$" : ""}${colToLetters(col)}${t.rowAbs ? "$" : ""}${row + 1}`;
    last = t.end;
  }
  return "=" + out + body.slice(last);
}
