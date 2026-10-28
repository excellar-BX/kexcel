import {
  checkNumber,
  compareScalars,
  parseNumberLiteral,
  toBoolean,
  toNumber,
  toScalar,
  toText,
} from "./coerce";
import { dateToSerial, daysInMonth, nowSerial, parseDateText, serialToParts, timeToSerial, type DateParts } from "./dates";
import { CellError, isError, isMatrix, type Matrix, type Scalar, type Value } from "./types";

export type FormulaFunction = (args: Value[]) => Value;

function arity(name: string, args: Value[], min: number, max = min) {
  if (args.length < min || args.length > max) {
    const expected = min === max ? `${min}` : max === Infinity ? `at least ${min}` : `${min}-${max}`;
    throw new CellError("#ERROR!", `${name} expects ${expected} argument(s), got ${args.length}`);
  }
}

function asMatrix(v: Value): Matrix {
  return isMatrix(v) ? v : [[v]];
}

/**
 * Numbers to aggregate. Values typed directly as arguments are coerced;
 * text, booleans and blanks inside ranges are skipped, as in Excel.
 */
function collectNumbers(args: Value[]): number[] {
  const out: number[] = [];
  for (const arg of args) {
    if (isMatrix(arg)) {
      for (const row of arg) {
        for (const v of row) {
          if (isError(v)) throw v;
          if (typeof v === "number") out.push(v);
        }
      }
    } else {
      out.push(toNumber(arg));
    }
  }
  return out;
}

function flatten(args: Value[]): Scalar[] {
  return args.flatMap((a) => (isMatrix(a) ? a.flat() : [a]));
}

function wildcardToRegex(pattern: string): RegExp {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) {
      re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    } else if (ch === "*") re += ".*";
    else if (ch === "?") re += ".";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "i");
}

/** Builds a matcher for SUMIF/COUNTIF criteria such as 5, ">10", "<>done" or "app*". */
export function makeCriteria(criteria: Scalar): (v: Scalar) => boolean {
  if (isError(criteria)) throw criteria;
  if (typeof criteria === "number" || typeof criteria === "boolean") {
    return (v) => v === criteria;
  }
  const text = criteria ?? "";
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(text)!;
  const op = m[1] ?? "=";
  const operand = m[2];
  const num = parseNumberLiteral(operand);

  if (num !== null) {
    return (v) => {
      if (typeof v !== "number") return op === "<>";
      switch (op) {
        case "=": return v === num;
        case "<>": return v !== num;
        case "<": return v < num;
        case ">": return v > num;
        case "<=": return v <= num;
        default: return v >= num;
      }
    };
  }

  if (operand === "") {
    if (op === "=") return (v) => v === null || v === "";
    if (op === "<>") return (v) => v !== null && v !== "";
  }

  if (op === "=" || op === "<>") {
    const upper = operand.toUpperCase();
    const bool = upper === "TRUE" ? true : upper === "FALSE" ? false : null;
    const re = wildcardToRegex(operand);
    const match = (v: Scalar) =>
      bool !== null ? v === bool : typeof v === "string" && re.test(v);
    return op === "=" ? match : (v) => !match(v);
  }

  return (v) => {
    if (typeof v !== "string") return false;
    const c = compareScalars(v, operand);
    return op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
  };
}

function conditionalPairs(name: string, args: Value[]) {
  arity(name, args, 2, 3);
  const range = asMatrix(args[0]);
  const test = makeCriteria(toScalar(args[1]));
  const target = args.length === 3 ? asMatrix(args[2]) : range;
  const hits: Scalar[] = [];
  for (let r = 0; r < range.length; r++) {
    for (let c = 0; c < range[r].length; c++) {
      if (test(range[r][c])) hits.push(target[r]?.[c] ?? null);
    }
  }
  return hits;
}

function roundTo(n: number, digits: number, mode: "nearest" | "up" | "down") {
  const f = 10 ** digits;
  // toPrecision strips binary noise so 1.005 rounds to 1.01 like Excel does.
  const scaled = Number((Math.abs(n) * f).toPrecision(15));
  const r = mode === "nearest" ? Math.round(scaled) : mode === "up" ? Math.ceil(scaled) : Math.floor(scaled);
  return (Math.sign(n) * r) / f;
}

function lookupPosition(needle: Scalar, haystack: Scalar[], exact: boolean): number {
  if (exact) {
    if (typeof needle === "string" && /[*?~]/.test(needle)) {
      const re = wildcardToRegex(needle);
      return haystack.findIndex((v) => typeof v === "string" && re.test(v));
    }
    return haystack.findIndex((v) => v !== null && compareScalars(v, needle) === 0);
  }
  // Approximate match: the last value <= needle, assuming ascending order.
  let found = -1;
  for (let i = 0; i < haystack.length; i++) {
    const v = haystack[i];
    if (v === null || typeof v !== typeof needle) continue;
    if (compareScalars(v, needle) <= 0) found = i;
    else break;
  }
  return found;
}

const math1 = (name: string, fn: (n: number) => number): FormulaFunction => (args) => {
  arity(name, args, 1);
  return checkNumber(fn(toNumber(args[0])));
};

export const FUNCTIONS: Record<string, FormulaFunction> = {
  SUM: (args) => checkNumber(collectNumbers(args).reduce((a, b) => a + b, 0)),
  PRODUCT: (args) => checkNumber(collectNumbers(args).reduce((a, b) => a * b, 1)),
  AVERAGE: (args) => {
    const nums = collectNumbers(args);
    if (nums.length === 0) throw new CellError("#DIV/0!", "AVERAGE of no numbers");
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  MIN: (args) => {
    const nums = collectNumbers(args);
    return nums.length ? Math.min(...nums) : 0;
  },
  MAX: (args) => {
    const nums = collectNumbers(args);
    return nums.length ? Math.max(...nums) : 0;
  },
  COUNT: (args) => flatten(args).filter((v) => typeof v === "number").length,
  COUNTA: (args) => flatten(args).filter((v) => v !== null && v !== "").length,
  COUNTBLANK: (args) => flatten(args).filter((v) => v === null || v === "").length,

  ROUND: (args) => {
    arity("ROUND", args, 1, 2);
    return roundTo(toNumber(args[0]), args[1] === undefined ? 0 : toNumber(args[1]), "nearest");
  },
  ROUNDUP: (args) => {
    arity("ROUNDUP", args, 1, 2);
    return roundTo(toNumber(args[0]), args[1] === undefined ? 0 : toNumber(args[1]), "up");
  },
  ROUNDDOWN: (args) => {
    arity("ROUNDDOWN", args, 1, 2);
    return roundTo(toNumber(args[0]), args[1] === undefined ? 0 : toNumber(args[1]), "down");
  },
  INT: math1("INT", Math.floor),
  ABS: math1("ABS", Math.abs),
  SQRT: (args) => {
    arity("SQRT", args, 1);
    const n = toNumber(args[0]);
    if (n < 0) throw new CellError("#NUM!", "SQRT of a negative number");
    return Math.sqrt(n);
  },
  POWER: (args) => {
    arity("POWER", args, 2);
    return checkNumber(toNumber(args[0]) ** toNumber(args[1]));
  },
  MOD: (args) => {
    arity("MOD", args, 2);
    const n = toNumber(args[0]);
    const d = toNumber(args[1]);
    if (d === 0) throw new CellError("#DIV/0!", "MOD by zero");
    return n - d * Math.floor(n / d);
  },

  AND: (args) => {
    if (args.length === 0) arity("AND", args, 1, Infinity);
    return flatten(args).filter((v) => v !== null).every((v) => toBoolean(v));
  },
  OR: (args) => {
    if (args.length === 0) arity("OR", args, 1, Infinity);
    return flatten(args).filter((v) => v !== null).some((v) => toBoolean(v));
  },
  NOT: (args) => {
    arity("NOT", args, 1);
    return !toBoolean(args[0]);
  },

  LEN: (args) => {
    arity("LEN", args, 1);
    return toText(args[0]).length;
  },
  UPPER: (args) => {
    arity("UPPER", args, 1);
    return toText(args[0]).toUpperCase();
  },
  LOWER: (args) => {
    arity("LOWER", args, 1);
    return toText(args[0]).toLowerCase();
  },
  TRIM: (args) => {
    arity("TRIM", args, 1);
    return toText(args[0]).trim().replace(/ +/g, " ");
  },
  LEFT: (args) => {
    arity("LEFT", args, 1, 2);
    const n = args[1] === undefined ? 1 : toNumber(args[1]);
    if (n < 0) throw new CellError("#VALUE!", "LEFT length is negative");
    return toText(args[0]).slice(0, n);
  },
  RIGHT: (args) => {
    arity("RIGHT", args, 1, 2);
    const n = args[1] === undefined ? 1 : toNumber(args[1]);
    if (n < 0) throw new CellError("#VALUE!", "RIGHT length is negative");
    const s = toText(args[0]);
    return n === 0 ? "" : s.slice(-n);
  },
  MID: (args) => {
    arity("MID", args, 3);
    const start = toNumber(args[1]);
    const len = toNumber(args[2]);
    if (start < 1 || len < 0) throw new CellError("#VALUE!", "MID start or length out of range");
    return toText(args[0]).substr(start - 1, len);
  },
  CONCAT: (args) => flatten(args).map((v) => toText(v)).join(""),
  CONCATENATE: (args) => args.map((v) => toText(v)).join(""),

  SUMIF: (args) =>
    conditionalPairs("SUMIF", args).reduce<number>((sum, v) => (typeof v === "number" ? sum + v : sum), 0),
  COUNTIF: (args) => {
    arity("COUNTIF", args, 2);
    return conditionalPairs("COUNTIF", args).length;
  },
  AVERAGEIF: (args) => {
    const nums = conditionalPairs("AVERAGEIF", args).filter((v): v is number => typeof v === "number");
    if (nums.length === 0) throw new CellError("#DIV/0!", "No matching numbers to average");
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },

  VLOOKUP: (args) => {
    arity("VLOOKUP", args, 3, 4);
    const needle = toScalar(args[0]);
    if (isError(needle)) throw needle;
    const table = asMatrix(args[1]);
    const colIndex = Math.trunc(toNumber(args[2]));
    const approx = args[3] === undefined ? true : toBoolean(args[3]);
    if (colIndex < 1) throw new CellError("#VALUE!", "Column index must be at least 1");
    if (colIndex > (table[0]?.length ?? 0)) throw new CellError("#REF!", "Column index is outside the table");
    const row = lookupPosition(needle, table.map((r) => r[0]), !approx);
    if (row < 0) throw new CellError("#N/A", "Value not found");
    return table[row][colIndex - 1];
  },
  MATCH: (args) => {
    arity("MATCH", args, 2, 3);
    const needle = toScalar(args[0]);
    if (isError(needle)) throw needle;
    const type = args[2] === undefined ? 1 : toNumber(args[2]);
    const pos = lookupPosition(needle, asMatrix(args[1]).flat(), type === 0);
    if (pos < 0) throw new CellError("#N/A", "Value not found");
    return pos + 1;
  },
  INDEX: (args) => {
    arity("INDEX", args, 2, 3);
    const table = asMatrix(args[0]);
    let row = Math.trunc(toNumber(args[1]));
    let col = args[2] === undefined ? 1 : Math.trunc(toNumber(args[2]));
    // A single row given one index is read along its columns.
    if (args[2] === undefined && table.length === 1) [row, col] = [1, row];
    const v = table[row - 1]?.[col - 1];
    if (v === undefined || row < 1 || col < 1) throw new CellError("#REF!", "Index is outside the range");
    return v;
  },

  ISBLANK: (args) => {
    arity("ISBLANK", args, 1);
    return toScalar(args[0]) === null;
  },
  ISNUMBER: (args) => {
    arity("ISNUMBER", args, 1);
    return typeof toScalar(args[0]) === "number";
  },
  ISTEXT: (args) => {
    arity("ISTEXT", args, 1);
    return typeof toScalar(args[0]) === "string";
  },
  ISERROR: (args) => {
    arity("ISERROR", args, 1);
    return isError(toScalar(args[0]));
  },

  // ---- Dates and times (serial day numbers, see dates.ts) ----
  TODAY: (args) => {
    arity("TODAY", args, 0);
    return nowSerial(false);
  },
  NOW: (args) => {
    arity("NOW", args, 0);
    return nowSerial(true);
  },
  DATE: (args) => {
    arity("DATE", args, 3);
    let y = Math.trunc(toNumber(args[0]));
    // Like Excel, years 0-1899 count from 1900.
    if (y >= 0 && y < 1900) y += 1900;
    const serial = dateToSerial(y, Math.trunc(toNumber(args[1])), Math.trunc(toNumber(args[2])));
    if (serial < 1) throw new CellError("#NUM!", "Dates before 1900 aren't supported");
    return serial;
  },
  TIME: (args) => {
    arity("TIME", args, 3);
    const t = timeToSerial(toNumber(args[0]), toNumber(args[1]), toNumber(args[2]));
    if (t < 0) throw new CellError("#NUM!", "Time can't be negative");
    return t - Math.floor(t);
  },
  DATEVALUE: (args) => {
    arity("DATEVALUE", args, 1);
    const parsed = parseDateText(toText(args[0]));
    if (!parsed || parsed.kind === "time") throw new CellError("#VALUE!", "That text isn't a date");
    return Math.floor(parsed.serial);
  },
  YEAR: datePart("YEAR", (p) => p.year),
  MONTH: datePart("MONTH", (p) => p.month),
  DAY: datePart("DAY", (p) => p.day),
  HOUR: datePart("HOUR", (p) => p.hour),
  MINUTE: datePart("MINUTE", (p) => p.minute),
  SECOND: datePart("SECOND", (p) => p.second),
  WEEKDAY: (args) => {
    arity("WEEKDAY", args, 1, 2);
    const day = serialToParts(toDate(args[0])).weekday; // 0 = Sunday
    const type = args[1] === undefined ? 1 : toNumber(args[1]);
    if (type === 1) return day + 1;
    if (type === 2) return ((day + 6) % 7) + 1;
    if (type === 3) return (day + 6) % 7;
    throw new CellError("#NUM!", "WEEKDAY type must be 1, 2 or 3");
  },
  DAYS: (args) => {
    arity("DAYS", args, 2);
    return Math.floor(toDate(args[0])) - Math.floor(toDate(args[1]));
  },
  EDATE: (args) => {
    arity("EDATE", args, 2);
    return addMonths(toDate(args[0]), Math.trunc(toNumber(args[1])), false);
  },
  EOMONTH: (args) => {
    arity("EOMONTH", args, 2);
    return addMonths(toDate(args[0]), Math.trunc(toNumber(args[1])), true);
  },
  DATEDIF: (args) => {
    arity("DATEDIF", args, 3);
    const start = Math.floor(toDate(args[0]));
    const end = Math.floor(toDate(args[1]));
    if (end < start) throw new CellError("#NUM!", "The end date is before the start date");
    const a = serialToParts(start);
    const b = serialToParts(end);
    let months = (b.year - a.year) * 12 + (b.month - a.month);
    if (b.day < a.day) months--;
    switch (toText(args[2]).toUpperCase()) {
      case "D":
        return end - start;
      case "M":
        return months;
      case "Y":
        return Math.floor(months / 12);
      case "YM":
        return months % 12;
      case "MD": {
        const anchor = addMonths(start, months, false);
        return end - anchor;
      }
      case "YD": {
        const years = Math.floor(months / 12);
        return end - addMonths(start, years * 12, false);
      }
      default:
        throw new CellError("#NUM!", 'DATEDIF unit must be "Y", "M", "D", "YM", "MD" or "YD"');
    }
  },
  NETWORKDAYS: (args) => {
    arity("NETWORKDAYS", args, 2);
    let start = Math.floor(toDate(args[0]));
    let end = Math.floor(toDate(args[1]));
    const sign = end < start ? -1 : 1;
    if (sign < 0) [start, end] = [end, start];
    if (end - start > 100_000) throw new CellError("#NUM!", "Date range is too long");
    let count = 0;
    for (let d = start; d <= end; d++) {
      const w = serialToParts(d).weekday;
      if (w !== 0 && w !== 6) count++;
    }
    return sign * count;
  },
};

/** A date argument: a serial number, or text such as "2025-01-15". */
function toDate(v: Value): number {
  const s = toScalar(v);
  if (typeof s === "string") {
    const parsed = parseDateText(s);
    if (parsed) return parsed.serial;
  }
  return toNumber(s);
}

function datePart(name: string, pick: (p: DateParts) => number): FormulaFunction {
  return (args) => {
    arity(name, args, 1);
    const serial = toDate(args[0]);
    if (serial < 0) throw new CellError("#NUM!", "Dates can't be negative");
    return pick(serialToParts(serial));
  };
}

/** Moves a date by whole months, keeping the day where it exists; `endOfMonth` snaps to the last day. */
function addMonths(serial: number, months: number, endOfMonth: boolean): number {
  const p = serialToParts(Math.floor(serial));
  const total = p.year * 12 + (p.month - 1) + months;
  const year = Math.floor(total / 12);
  const month = (total % 12) + 1;
  const last = daysInMonth(year, month);
  return dateToSerial(year, month, endOfMonth ? last : Math.min(p.day, last));
}

/** Functions whose result changes over time, so their cells recalculate on every change. */
export const VOLATILE_FUNCTIONS = new Set(["TODAY", "NOW"]);
