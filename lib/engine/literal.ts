import { parseNumberLiteral } from "./coerce";
import type { Scalar } from "./types";

export const isFormula = (raw: string) => raw.startsWith("=") && raw.length > 1;

/** Interprets typed input that is not a formula: numbers, booleans, or text. */
export function parseLiteral(raw: string): Scalar {
  if (raw === "") return null;
  if (raw.startsWith("'")) return raw.slice(1);
  const n = parseNumberLiteral(raw);
  if (n !== null) return n;
  const upper = raw.trim().toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return raw;
}
