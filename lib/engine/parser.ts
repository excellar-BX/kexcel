import { FormulaSyntaxError, tokenize, type Token } from "./tokenizer";
import type { CellAddress } from "./types";

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";

export type Node =
  | { type: "number"; value: number }
  | { type: "string"; value: string }
  | { type: "boolean"; value: boolean }
  /** `sheet` is the sheet name written in the formula; absent means the formula's own sheet. */
  | { type: "ref"; addr: CellAddress; sheet?: string }
  | { type: "range"; start: CellAddress; end: CellAddress; sheet?: string }
  | { type: "name"; name: string }
  | { type: "refError" }
  | { type: "unary"; op: "-" | "+"; arg: Node }
  | { type: "percent"; arg: Node }
  | { type: "binary"; op: BinaryOp; left: Node; right: Node }
  | { type: "call"; name: string; args: Node[] };

const BINARY_PRECEDENCE: Record<string, number> = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "&": 2,
  "+": 3, "-": 3,
  "*": 4, "/": 4,
  "^": 5,
};
// Excel binds unary minus tighter than ^, so -2^2 is 4.
const UNARY_PRECEDENCE = 6;
const PERCENT_PRECEDENCE = 7;

/** Parses a formula body (without the leading "=") into an AST. */
export function parseFormula(src: string): Node {
  const tokens = tokenize(src);
  let pos = 0;

  const peek = (): Token | undefined => tokens[pos];
  const next = (): Token => {
    const t = tokens[pos++];
    if (!t) throw new FormulaSyntaxError("Unexpected end of formula");
    return t;
  };
  const expect = (type: Token["type"]) => {
    const t = next();
    if (t.type !== type) throw new FormulaSyntaxError(`Expected ${type}`);
    return t;
  };

  function parseExpression(minPrec: number): Node {
    let left = parsePrefix();
    for (;;) {
      const t = peek();
      if (!t || t.type !== "op") break;
      if (t.value === "%") {
        if (PERCENT_PRECEDENCE < minPrec) break;
        pos++;
        left = { type: "percent", arg: left };
        continue;
      }
      const prec = BINARY_PRECEDENCE[t.value];
      if (prec === undefined || prec < minPrec) break;
      pos++;
      const right = parseExpression(prec + 1);
      left = { type: "binary", op: t.value as BinaryOp, left, right };
    }
    return left;
  }

  function parsePrefix(): Node {
    const t = next();
    switch (t.type) {
      case "number":
        return { type: "number", value: t.value };
      case "string":
        return { type: "string", value: t.value };
      case "refError":
        return { type: "refError" };
      case "op":
        if (t.value === "-" || t.value === "+") {
          return { type: "unary", op: t.value, arg: parseExpression(UNARY_PRECEDENCE) };
        }
        throw new FormulaSyntaxError(`Unexpected "${t.value}"`);
      case "lparen": {
        const inner = parseExpression(0);
        expect("rparen");
        return inner;
      }
      case "sheet": {
        const ref = next();
        if (ref.type !== "ref") throw new FormulaSyntaxError("Expected a cell after the sheet name");
        return parseReference(ref, t.name);
      }
      case "ref":
        return parseReference(t);
      case "ident": {
        if (peek()?.type === "lparen") {
          pos++;
          const args: Node[] = [];
          if (peek()?.type !== "rparen") {
            for (;;) {
              args.push(parseExpression(0));
              if (peek()?.type === "comma") {
                pos++;
                continue;
              }
              break;
            }
          }
          expect("rparen");
          return { type: "call", name: t.value, args };
        }
        if (t.value === "TRUE" || t.value === "FALSE") {
          return { type: "boolean", value: t.value === "TRUE" };
        }
        return { type: "name", name: t.value };
      }
      default:
        throw new FormulaSyntaxError("Unexpected token");
    }
  }

  function parseReference(t: Extract<Token, { type: "ref" }>, sheet?: string): Node {
    const start = { row: t.row, col: t.col };
    if (peek()?.type !== "colon") return { type: "ref", addr: start, sheet };
    pos++;
    const endTok = expect("ref") as Extract<Token, { type: "ref" }>;
    const end = { row: endTok.row, col: endTok.col };
    return {
      type: "range",
      start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
      end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
      sheet,
    };
  }

  const ast = parseExpression(0);
  if (pos < tokens.length) throw new FormulaSyntaxError("Unexpected token after expression");
  return ast;
}

export interface Reference extends CellAddress {
  sheet?: string;
}

/** Collects every cell the formula reads, expanding ranges cell by cell. */
export function collectReferences(node: Node, out: Reference[] = []): Reference[] {
  switch (node.type) {
    case "ref":
      out.push({ ...node.addr, sheet: node.sheet });
      break;
    case "range":
      for (let r = node.start.row; r <= node.end.row; r++) {
        for (let c = node.start.col; c <= node.end.col; c++) out.push({ row: r, col: c, sheet: node.sheet });
      }
      break;
    case "unary":
    case "percent":
      collectReferences(node.arg, out);
      break;
    case "binary":
      collectReferences(node.left, out);
      collectReferences(node.right, out);
      break;
    case "call":
      for (const a of node.args) collectReferences(a, out);
      break;
  }
  return out;
}
