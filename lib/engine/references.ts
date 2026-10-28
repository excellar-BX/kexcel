import { colToLetters } from "./address";
import { isFormula } from "./literal";
import { tokenize, type Token } from "./tokenizer";

/** One end of a reference as written in a formula. */
export interface RefPart {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export interface RefContext {
  /** Sheet name written before the reference, or undefined for the formula's own sheet. */
  sheet: string | undefined;
}

export interface ReferenceMapper {
  /** Returns the new reference, or null to replace it with #REF!. */
  ref?: (part: RefPart, ctx: RefContext) => RefPart | null;
  /** Ranges default to mapping both ends with `ref`. */
  range?: (start: RefPart, end: RefPart, ctx: RefContext) => [RefPart, RefPart] | null;
  /** Returns the new sheet name written before a reference. */
  sheet?: (name: string) => string;
}

const BARE_SHEET_NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/;
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}\d+$/;

/** How a sheet name is written before "!" in a formula, quoted when needed. */
export function sheetPrefix(name: string): string {
  const bare = BARE_SHEET_NAME.test(name) && !LOOKS_LIKE_CELL.test(name) && !/^(TRUE|FALSE)$/i.test(name);
  return bare ? `${name}!` : `'${name.replace(/'/g, "''")}'!`;
}

const writePart = (p: RefPart) =>
  `${p.colAbs ? "$" : ""}${colToLetters(p.col)}${p.rowAbs ? "$" : ""}${p.row + 1}`;

const partOf = (t: Extract<Token, { type: "ref" }>): RefPart => ({
  row: t.row,
  col: t.col,
  rowAbs: t.rowAbs,
  colAbs: t.colAbs,
});

const offSheet = (p: RefPart) => p.row < 0 || p.col < 0;

/**
 * Rewrites the references in a formula while leaving everything else,
 * including spacing and string literals, exactly as typed.
 * Non-formulas and formulas that fail to tokenize are returned unchanged.
 */
export function mapReferences(raw: string, mapper: ReferenceMapper): string {
  if (!isFormula(raw)) return raw;
  const body = raw.slice(1);
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return raw;
  }

  const mapRef = (p: RefPart, ctx: RefContext) => {
    const out = mapper.ref ? mapper.ref(p, ctx) : p;
    return out && !offSheet(out) ? out : null;
  };

  let out = "";
  let last = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type !== "sheet" && t.type !== "ref") continue;

    const start = t.start;
    let sheet: string | undefined;
    let refTok = t;
    if (t.type === "sheet") {
      const next = tokens[i + 1];
      if (next?.type !== "ref") continue;
      sheet = t.name;
      refTok = next;
      i++;
    }
    const ctx = { sheet };
    const first = partOf(refTok as Extract<Token, { type: "ref" }>);
    let end = refTok.end;
    let text: string | null;

    const colon = tokens[i + 1];
    const second = tokens[i + 2];
    if (colon?.type === "colon" && second?.type === "ref") {
      i += 2;
      end = second.end;
      const b = partOf(second);
      let pair: [RefPart, RefPart] | null;
      if (mapper.range) {
        pair = mapper.range(first, b, ctx);
        if (pair && (offSheet(pair[0]) || offSheet(pair[1]))) pair = null;
      } else {
        const x = mapRef(first, ctx);
        const y = mapRef(b, ctx);
        pair = x && y ? [x, y] : null;
      }
      text = pair ? `${writePart(pair[0])}:${writePart(pair[1])}` : null;
    } else {
      const p = mapRef(first, ctx);
      text = p ? writePart(p) : null;
    }

    out += body.slice(last, start);
    if (text === null) {
      out += "#REF!";
    } else {
      if (sheet !== undefined) out += sheetPrefix(mapper.sheet ? mapper.sheet(sheet) : sheet);
      out += text;
    }
    last = end;
  }
  return "=" + out + body.slice(last);
}

/**
 * Moves the relative references in a formula, as when it is pasted
 * dRow rows and dCol columns away from where it was copied. References
 * pushed off the sheet become #REF!. Anchored ($) parts stay put.
 */
export function shiftFormula(raw: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return raw;
  return mapReferences(raw, {
    ref: (p) => ({
      ...p,
      row: p.rowAbs ? p.row : p.row + dRow,
      col: p.colAbs ? p.col : p.col + dCol,
    }),
  });
}

export type Axis = "row" | "col";

/**
 * Adjusts references after rows or columns are inserted (count > 0) or
 * deleted (count < 0) at `at`. `targets` decides whether a reference
 * points at the sheet that changed. Unlike pasting, this moves anchored
 * references too, as Excel does.
 */
export function adjustForStructureChange(
  raw: string,
  axis: Axis,
  at: number,
  count: number,
  targets: (ctx: RefContext) => boolean,
): string {
  const get = (p: RefPart) => (axis === "row" ? p.row : p.col);
  const set = (p: RefPart, v: number): RefPart => (axis === "row" ? { ...p, row: v } : { ...p, col: v });

  if (count > 0) {
    const move = (p: RefPart) => (get(p) >= at ? set(p, get(p) + count) : p);
    return mapReferences(raw, {
      ref: (p, ctx) => (targets(ctx) ? move(p) : p),
      range: (a, b, ctx) => (targets(ctx) ? [move(a), move(b)] : [a, b]),
    });
  }

  const removed = -count;
  const endRemoved = at + removed; // first index after the deleted block
  return mapReferences(raw, {
    ref: (p, ctx) => {
      if (!targets(ctx)) return p;
      const v = get(p);
      if (v < at) return p;
      if (v >= endRemoved) return set(p, v - removed);
      return null;
    },
    range: (a, b, ctx) => {
      if (!targets(ctx)) return [a, b];
      const lo = Math.min(get(a), get(b));
      const hi = Math.max(get(a), get(b));
      const newLo = lo < at ? lo : lo >= endRemoved ? lo - removed : at;
      const newHi = hi < at ? hi : hi >= endRemoved ? hi - removed : at - 1;
      if (newHi < newLo) return null;
      const [first, second] = get(a) <= get(b) ? [a, b] : [b, a];
      return [set(first, newLo), set(second, newHi)];
    },
  });
}

/** Rewrites "Old!A1" references to a renamed sheet. */
export function renameSheetInFormula(raw: string, oldName: string, newName: string): string {
  const old = oldName.toLowerCase();
  return mapReferences(raw, { sheet: (name) => (name.toLowerCase() === old ? newName : name) });
}
