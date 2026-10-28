import { lettersToCol } from "./address";

export type Token =
  | { type: "number"; value: number; start: number; end: number }
  | { type: "string"; value: string; start: number; end: number }
  | {
      type: "ref";
      row: number;
      col: number;
      rowAbs: boolean;
      colAbs: boolean;
      start: number;
      end: number;
    }
  | { type: "ident"; value: string; start: number; end: number }
  | { type: "op"; value: string; start: number; end: number }
  | { type: "refError"; start: number; end: number }
  /** A sheet prefix such as `Sheet2!` or `'My data'!`; the reference follows it. */
  | { type: "sheet"; name: string; start: number; end: number }
  | { type: "lparen" | "rparen" | "comma" | "colon"; start: number; end: number };

export class FormulaSyntaxError extends Error {}

const NUMBER_RE = /(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/y;
const REF_RE = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y;
const IDENT_RE = /[A-Za-z_][A-Za-z0-9_.]*/y;
const SHEET_RE = /([A-Za-z_][A-Za-z0-9_.]*)!/y;
const TWO_CHAR_OPS = ["<=", ">=", "<>"];
const ONE_CHAR_OPS = "+-*/^&=<>%";

/** Tokenizes a formula body (the text after the leading "="). */
export function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  const matchAt = (re: RegExp) => {
    re.lastIndex = i;
    return re.exec(src);
  };

  while (i < src.length) {
    const ch = src[i];

    if (ch === " " || ch === "\t" || ch === "\n") {
      i++;
      continue;
    }

    if (ch === '"') {
      let value = "";
      let j = i + 1;
      for (;;) {
        if (j >= src.length) throw new FormulaSyntaxError("Unclosed string");
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            value += '"';
            j += 2;
            continue;
          }
          break;
        }
        value += src[j++];
      }
      tokens.push({ type: "string", value, start: i, end: j + 1 });
      i = j + 1;
      continue;
    }

    if (ch === "'") {
      let name = "";
      let j = i + 1;
      for (;;) {
        if (j >= src.length) throw new FormulaSyntaxError("Unclosed sheet name");
        if (src[j] === "'") {
          if (src[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          break;
        }
        name += src[j++];
      }
      if (src[j + 1] !== "!") throw new FormulaSyntaxError("Expected ! after sheet name");
      tokens.push({ type: "sheet", name, start: i, end: j + 2 });
      i = j + 2;
      continue;
    }

    if (src.startsWith("#REF!", i)) {
      tokens.push({ type: "refError", start: i, end: i + 5 });
      i += 5;
      continue;
    }

    if (/[0-9.]/.test(ch)) {
      const m = matchAt(NUMBER_RE);
      if (m) {
        tokens.push({ type: "number", value: parseFloat(m[0]), start: i, end: i + m[0].length });
        i += m[0].length;
        continue;
      }
    }

    if (ch === "$" || /[A-Za-z_]/.test(ch)) {
      const sheet = matchAt(SHEET_RE);
      if (sheet) {
        tokens.push({ type: "sheet", name: sheet[1], start: i, end: i + sheet[0].length });
        i += sheet[0].length;
        continue;
      }
      const ref = matchAt(REF_RE);
      if (ref) {
        tokens.push({
          type: "ref",
          colAbs: ref[1] === "$",
          col: lettersToCol(ref[2]),
          rowAbs: ref[3] === "$",
          row: parseInt(ref[4], 10) - 1,
          start: i,
          end: i + ref[0].length,
        });
        i += ref[0].length;
        continue;
      }
      const id = matchAt(IDENT_RE);
      if (id) {
        tokens.push({ type: "ident", value: id[0].toUpperCase(), start: i, end: i + id[0].length });
        i += id[0].length;
        continue;
      }
    }

    const two = src.slice(i, i + 2);
    if (TWO_CHAR_OPS.includes(two)) {
      tokens.push({ type: "op", value: two, start: i, end: i + 2 });
      i += 2;
      continue;
    }
    if (ONE_CHAR_OPS.includes(ch)) {
      tokens.push({ type: "op", value: ch, start: i, end: i + 1 });
      i++;
      continue;
    }

    const punct = { "(": "lparen", ")": "rparen", ",": "comma", ":": "colon" } as const;
    if (ch in punct) {
      tokens.push({ type: punct[ch as keyof typeof punct], start: i, end: i + 1 });
      i++;
      continue;
    }

    throw new FormulaSyntaxError(`Unexpected character "${ch}"`);
  }

  return tokens;
}
