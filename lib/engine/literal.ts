import { parseNumberLiteral } from "./coerce";
import { parseDateText, type DateKind } from "./dates";
import { parseFormula } from "./parser";
import type { Scalar } from "./types";

export const isFormula = (raw: string) => raw.startsWith("=") && raw.length > 1;

/** Interprets typed input that is not a formula: numbers, booleans, or text. */
export function parseLiteral(raw: string): Scalar {
  if (raw === "") return null;
  if (raw.startsWith("'")) return raw.slice(1);
  const n = parseNumberLiteral(raw);
  if (n !== null) return n;
  const date = parseDateText(raw);
  if (date) return date.serial;
  const upper = raw.trim().toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return raw;
}

const DATE_RESULTS: Record<string, DateKind> = {
  TODAY: "date",
  DATE: "date",
  DATEVALUE: "date",
  EDATE: "date",
  EOMONTH: "date",
  NOW: "datetime",
  TIME: "time",
};

/**
 * Whether typed input is a date or time that should get a date format when the
 * cell has none: a typed date such as 2025-01-15, or a formula like =TODAY().
 */
export function impliedDateKind(raw: string): DateKind | null {
  if (!isFormula(raw)) return raw.startsWith("'") ? null : (parseDateText(raw)?.kind ?? null);
  try {
    const ast = parseFormula(raw.slice(1));
    return ast.type === "call" ? (DATE_RESULTS[ast.name] ?? null) : null;
  } catch {
    return null;
  }
}
