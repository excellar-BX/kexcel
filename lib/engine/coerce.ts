import { CellError, isError, isMatrix, type Scalar, type Value } from "./types";

export function formatNumber(n: number): string {
  // Trim floating point noise such as 0.1 + 0.2 = 0.30000000000000004.
  return String(Number(n.toPrecision(15)));
}

/** Collapses a value to one scalar; multi-cell ranges are not valid here. */
export function toScalar(v: Value): Scalar {
  if (!isMatrix(v)) return v;
  if (v.length === 1 && v[0].length === 1) return v[0][0];
  throw new CellError("#VALUE!", "Expected a single cell, got a range");
}

export function toNumber(v: Value): number {
  const s = toScalar(v);
  if (isError(s)) throw s;
  if (s === null) return 0;
  if (typeof s === "number") return s;
  if (typeof s === "boolean") return s ? 1 : 0;
  const n = parseNumberLiteral(s);
  if (n === null) throw new CellError("#VALUE!", `"${s}" is not a number`);
  return n;
}

export function toText(v: Value): string {
  const s = toScalar(v);
  if (isError(s)) throw s;
  if (s === null) return "";
  if (typeof s === "number") return formatNumber(s);
  if (typeof s === "boolean") return s ? "TRUE" : "FALSE";
  return s;
}

export function toBoolean(v: Value): boolean {
  const s = toScalar(v);
  if (isError(s)) throw s;
  if (s === null) return false;
  if (typeof s === "boolean") return s;
  if (typeof s === "number") return s !== 0;
  const upper = s.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  throw new CellError("#VALUE!", `"${s}" is not TRUE or FALSE`);
}

/** Parses "12", "-3.5", "1e3", "1,000" and "50%" into numbers; returns null otherwise. */
export function parseNumberLiteral(text: string): number | null {
  const t = text.trim().replace(/,/g, "");
  if (t === "") return null;
  const pct = t.endsWith("%");
  const body = pct ? t.slice(0, -1) : t;
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(body)) return null;
  const n = parseFloat(body);
  return pct ? n / 100 : n;
}

const typeRank = (v: Scalar) =>
  typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 0;

/** Excel ordering: numbers < text < booleans; text compares case-insensitively. */
export function compareScalars(a: Scalar, b: Scalar): number {
  if (isError(a)) throw a;
  if (isError(b)) throw b;
  // An empty cell takes the type of whatever it is compared against.
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "string") {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = Number(a);
  const y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

export function checkNumber(n: number): number {
  if (!Number.isFinite(n)) throw new CellError("#NUM!", "Result is not a finite number");
  return n;
}
