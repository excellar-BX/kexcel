import type { CellFormat } from "@/lib/engine";

/** Excel's built-in number formats by id (the ones that aren't locale-dependent dates). */
const BUILTIN: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  48: "##0.0E+0",
  49: "@",
};

/** Built-in ids that are dates or times. */
const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

const SYMBOLS: [string, string][] = [
  ["GH₵", "GHS"],
  ["KSh", "KES"],
  ["₦", "NGN"],
  ["€", "EUR"],
  ["£", "GBP"],
  ["¥", "JPY"],
  ["₹", "INR"],
  ["C$", "CAD"],
  ["CA$", "CAD"],
  ["$", "USD"],
];

export function builtinFormat(id: number): string | undefined {
  return BUILTIN[id];
}

/** The format code without quoted text, escapes, [colour]/[locale] blocks and padding. */
function bare(code: string) {
  return code
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/[_*]./g, "");
}

export function isDateFormat(id: number, code: string | undefined): boolean {
  if (DATE_IDS.has(id)) return true;
  if (!code) return false;
  const section = bare(code.split(";")[0]);
  return /[dmyhs]/i.test(section) && !/^general$/i.test(section.trim());
}

/** Turns an Excel number format code into Kexcel's number format fields. */
export function formatFromCode(code: string | undefined): Partial<CellFormat> {
  if (!code || /^general$/i.test(code.trim()) || code.trim() === "@") return {};
  const first = code.split(";")[0];
  const plain = bare(first);
  const dot = plain.indexOf(".");
  const decimals = dot < 0 ? 0 : (plain.slice(dot + 1).match(/^[0#?]*/)?.[0].length ?? 0);

  if (plain.includes("%")) return { numFmt: "percent", decimals };
  if (/E[+-]/i.test(plain)) return { numFmt: "scientific", decimals };

  const quoted = [...first.matchAll(/"([^"]*)"/g)].map((m) => m[1]).join("");
  const locale = first.match(/\[\$([^\]-]*)/)?.[1] ?? "";
  const symbolText = `${quoted}${locale}${first.replace(/"[^"]*"|\[[^\]]*\]/g, "")}`;
  const currency = SYMBOLS.find(([sym]) => symbolText.includes(sym))?.[1];
  if (currency) return { numFmt: "currency", currency, decimals };

  if (plain.includes(",")) return { numFmt: "number", decimals };
  if (/[0#]/.test(plain)) return { decimals };
  return {};
}

/** The currency symbol Excel should show for a currency code. */
function currencySymbol(currency: string) {
  try {
    const parts = new Intl.NumberFormat("en-US", { style: "currency", currency, currencyDisplay: "narrowSymbol" }).formatToParts(0);
    return parts.find((p) => p.type === "currency")?.value ?? currency;
  } catch {
    return currency;
  }
}

/** The Excel format code for a Kexcel format, or null for General. */
export function codeFromFormat(f: CellFormat): string | null {
  const zeros = (n: number) => (n > 0 ? "." + "0".repeat(n) : "");
  switch (f.numFmt) {
    case "number":
      return "#,##0" + zeros(f.decimals ?? 2);
    case "percent":
      return "0" + zeros(f.decimals ?? 2) + "%";
    case "scientific":
      return "0" + zeros(f.decimals ?? 2) + "E+00";
    case "currency": {
      const sym = currencySymbol(f.currency ?? "USD").replace(/"/g, "");
      const body = "#,##0" + zeros(f.decimals ?? 2);
      return `"${sym}"${body};-"${sym}"${body}`;
    }
    default:
      return f.decimals === undefined ? null : "0" + zeros(f.decimals);
  }
}

/** Excel stores dates as days since 1899-12-30; shows them as YYYY-MM-DD (and time if any). */
export function serialToDateText(serial: number): string {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const pad = (n: number) => String(n).padStart(2, "0");
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const hasTime = Math.abs(serial - Math.floor(serial)) > 1e-9;
  if (!hasTime) return date;
  const time = `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
  return serial < 1 ? time : `${date} ${time}`;
}
