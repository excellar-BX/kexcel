import {
  formatValue,
  isError,
  parseKey,
  parseLiteral,
  type ConditionalRule,
  type Scalar,
  type Workbook,
} from "@/lib/engine";

/** What conditional formatting adds to a cell's own format. `bar` is a data bar's length, 0 to 1. */
export interface Look {
  fill?: string;
  color?: string;
  bold?: boolean;
  bar?: number;
}

interface Prepared {
  rule: ConditionalRule;
  operands: Scalar[];
  min: number;
  max: number;
  duplicates: Set<string>;
}

/** Highlight styles offered in the rule editor. Text colours stay dark enough to read on their fills. */
export const HIGHLIGHT_STYLES: { label: string; style: NonNullable<ConditionalRule["style"]> }[] = [
  { label: "Red fill, dark red text", style: { fill: "#fde7e7", color: "#a52a24" } },
  { label: "Yellow fill, dark yellow text", style: { fill: "#fff3c4", color: "#7a5900" } },
  { label: "Green fill, dark green text", style: { fill: "#e3f4e6", color: "#1e6b2f" } },
  { label: "Blue fill, dark blue text", style: { fill: "#e3eefc", color: "#1c4f95" } },
  { label: "Red text", style: { color: "#d03b3b" } },
  { label: "Bold text", style: { bold: true } },
];

const inRange = (r: ConditionalRule["range"], row: number, col: number) =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

const textKey = (v: Scalar) => (typeof v === "string" ? v.trim().toLowerCase() : formatValue(v));

/** Works out per-rule figures (lowest, highest, repeated values) once per change, not per cell. */
export function prepareRules(wb: Workbook, sheet: string): Prepared[] {
  const rules = wb.conditionalRules(sheet);
  if (rules.length === 0) return [];
  const filled = [...wb.cellKeys(sheet)].map(parseKey);
  return rules.map((rule) => {
    let min = Infinity;
    let max = -Infinity;
    const seen = new Map<string, number>();
    for (const { row, col } of filled) {
      if (!inRange(rule.range, row, col)) continue;
      const v = wb.getValue(sheet, row, col);
      if (typeof v === "number") {
        min = Math.min(min, v);
        max = Math.max(max, v);
      }
      if (rule.kind === "duplicate" && v !== null && v !== "" && !isError(v)) {
        const key = textKey(v);
        seen.set(key, (seen.get(key) ?? 0) + 1);
      }
    }
    const duplicates = new Set([...seen].filter(([, n]) => n > 1).map(([k]) => k));
    const operands = (rule.values ?? []).map((v) => parseLiteral(v.trim()));
    return { rule, operands, min, max, duplicates };
  });
}

/** Compares a cell value with a typed operand: numbers as numbers, anything else as text, ignoring case. */
function compare(v: Scalar, operand: Scalar): number | null {
  if (v === null || v === "" || isError(v) || operand === null || isError(operand)) return null;
  if (typeof v === "number" && typeof operand === "number") return v - operand;
  if (typeof v === "number" || typeof operand === "number") return null;
  return String(v).localeCompare(String(operand), undefined, { sensitivity: "base", numeric: true });
}

function matches(p: Prepared, v: Scalar): boolean {
  const [a, b] = p.operands;
  switch (p.rule.kind) {
    case "greater": {
      const c = compare(v, a);
      return c !== null && c > 0;
    }
    case "less": {
      const c = compare(v, a);
      return c !== null && c < 0;
    }
    case "between": {
      const lo = compare(v, a);
      const hi = compare(v, b);
      return lo !== null && hi !== null && lo >= 0 && hi <= 0;
    }
    case "equal":
      return compare(v, a) === 0;
    case "notEqual": {
      const c = compare(v, a);
      return c !== null && c !== 0;
    }
    case "contains": {
      const needle = String(p.rule.values?.[0] ?? "").trim().toLowerCase();
      return v !== null && needle !== "" && formatValue(v).toLowerCase().includes(needle);
    }
    case "blank":
      return v === null || v === "";
    case "notBlank":
      return v !== null && v !== "";
    case "duplicate":
      return v !== null && v !== "" && !isError(v) && p.duplicates.has(textKey(v));
    default:
      return false;
  }
}

/** A tint of a theme colour; strength 0 to 1. Works in light and dark themes alike. */
const tint = (token: string, strength: number) =>
  `color-mix(in oklab, var(${token}) ${Math.round(8 + strength * 62)}%, transparent)`;

function scaleFill(p: Prepared, v: number): string | undefined {
  if (!(p.max > p.min)) return undefined;
  if (p.rule.scale === "diverging") {
    // Two hues around a neutral middle: zero when the data crosses it, otherwise the halfway point.
    const mid = p.min < 0 && p.max > 0 ? 0 : (p.min + p.max) / 2;
    if (v === mid) return undefined;
    return v < mid ? tint("--series-8", (mid - v) / (mid - p.min)) : tint("--series-1", (v - mid) / (p.max - mid));
  }
  return tint(p.rule.scale === "orange" ? "--series-2" : "--series-1", (v - p.min) / (p.max - p.min));
}

/** The combined look of every rule that applies to a cell. Earlier rules win where they overlap. */
export function lookAt(prepared: Prepared[], value: Scalar, row: number, col: number): Look | null {
  let look: Look | null = null;
  for (const p of prepared) {
    if (!inRange(p.rule.range, row, col)) continue;
    const add: Look = {};
    if (p.rule.kind === "scale") {
      const fill = typeof value === "number" ? scaleFill(p, value) : undefined;
      if (fill) add.fill = fill;
    } else if (p.rule.kind === "bar") {
      if (typeof value === "number" && p.max > Math.min(0, p.min)) {
        const lo = Math.min(0, p.min);
        add.bar = Math.max(0, Math.min(1, (value - lo) / (p.max - lo)));
      }
    } else if (matches(p, value)) {
      Object.assign(add, p.rule.style);
    }
    if (Object.keys(add).length === 0) continue;
    look = { ...add, ...(look ?? {}) };
  }
  return look;
}

/** Words for a rule, as shown in the rule list. */
export function describeRule(rule: ConditionalRule): string {
  const [a, b] = rule.values ?? [];
  switch (rule.kind) {
    case "greater":
      return `Greater than ${a}`;
    case "less":
      return `Less than ${a}`;
    case "between":
      return `Between ${a} and ${b}`;
    case "equal":
      return `Equal to ${a}`;
    case "notEqual":
      return `Not equal to ${a}`;
    case "contains":
      return `Text contains "${a}"`;
    case "blank":
      return "Is empty";
    case "notBlank":
      return "Is not empty";
    case "duplicate":
      return "Duplicate values";
    case "scale":
      return rule.scale === "diverging" ? "Colour scale: below and above the middle" : "Colour scale: low to high";
    case "bar":
      return "Data bars";
  }
}
