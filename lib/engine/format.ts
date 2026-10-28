import { formatNumber } from "./coerce";
import { formatSerial, type DateStyle } from "./dates";
import { isError, type Scalar } from "./types";

export type NumberFormat = "general" | "number" | "currency" | "percent" | "scientific" | "date" | "time" | "datetime";
export type HorizontalAlign = "left" | "center" | "right";
export type FontFamily = "sans" | "serif" | "mono";

export interface Borders {
  top?: boolean;
  right?: boolean;
  bottom?: boolean;
  left?: boolean;
}

/** How a cell looks. Every field is optional; absent means the default. */
export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** CSS colours. */
  color?: string;
  fill?: string;
  /** In points; the default is 10. */
  fontSize?: number;
  fontFamily?: FontFamily;
  align?: HorizontalAlign;
  numFmt?: NumberFormat;
  /** Digits after the decimal point; each number format has its own default. */
  decimals?: number;
  /** ISO 4217 code used by the currency format. */
  currency?: string;
  /** How date and datetime formats write the date; ISO (2025-01-15) by default. */
  dateStyle?: DateStyle;
  /** Wrap long text onto more lines instead of spilling into the next cells. */
  wrap?: boolean;
  borders?: Borders;
}

export const DEFAULT_FONT_SIZE = 10;

/** Combines formats; `undefined` or `false` in the patch removes that field. */
export function mergeFormat(base: CellFormat | undefined, patch: CellFormat): CellFormat | null {
  const out: Record<string, unknown> = { ...base };
  for (const [k, v] of Object.entries(patch)) {
    if (v === undefined || v === false) delete out[k];
    else out[k] = v;
  }
  if (out.borders && !Object.values(out.borders as Borders).some(Boolean)) delete out.borders;
  return Object.keys(out).length ? (out as CellFormat) : null;
}

export function sameFormat(a: CellFormat | null | undefined, b: CellFormat | null | undefined): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/** Decimal places a value shows by default, used when the user first nudges decimals. */
export function shownDecimals(value: Scalar, format?: CellFormat): number {
  if (format?.decimals !== undefined) return format.decimals;
  switch (format?.numFmt) {
    case "number":
    case "currency":
    case "percent":
    case "scientific":
      return 2;
    case "date":
    case "time":
    case "datetime":
      return 0;
  }
  if (typeof value !== "number") return 0;
  const text = formatNumber(value);
  const dot = text.indexOf(".");
  return dot < 0 || text.includes("e") ? 0 : text.length - dot - 1;
}

/** Text shown in a cell for its computed value. */
export function formatValue(v: Scalar, format?: CellFormat | null, locale?: string): string {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isError(v)) return v.code;
  if (typeof v === "string") return v;

  const fmt = format?.numFmt ?? "general";
  if (fmt === "general") {
    if (format?.decimals === undefined) return formatNumber(v);
    return v.toLocaleString(locale, {
      minimumFractionDigits: format.decimals,
      maximumFractionDigits: format.decimals,
      useGrouping: false,
    });
  }
  if (fmt === "date" || fmt === "time" || fmt === "datetime") return formatSerial(v, fmt, format?.dateStyle, locale);
  const decimals = format?.decimals ?? 2;
  const digits = { minimumFractionDigits: decimals, maximumFractionDigits: decimals };
  switch (fmt) {
    case "number":
      return v.toLocaleString(locale, digits);
    case "percent":
      return (v * 100).toLocaleString(locale, digits) + "%";
    case "scientific":
      return v.toExponential(decimals).toUpperCase();
    case "currency":
      try {
        return v.toLocaleString(locale, { style: "currency", currency: format?.currency ?? "USD", ...digits });
      } catch {
        return v.toLocaleString(locale, digits);
      }
  }
}

export const CURRENCIES = [
  { code: "USD", label: "US dollar" },
  { code: "EUR", label: "Euro" },
  { code: "GBP", label: "British pound" },
  { code: "NGN", label: "Nigerian naira" },
  { code: "GHS", label: "Ghanaian cedi" },
  { code: "KES", label: "Kenyan shilling" },
  { code: "ZAR", label: "South African rand" },
  { code: "INR", label: "Indian rupee" },
  { code: "CAD", label: "Canadian dollar" },
  { code: "JPY", label: "Japanese yen" },
] as const;

/** Text for a value with no formatting applied. */
export const displayValue = (v: Scalar) => formatValue(v);
