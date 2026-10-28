// Dates and times are numbers, as in Excel: whole days since 1899-12-30, with
// the time of day as the fraction. That epoch matches Excel's serial numbers
// for every date from 1900-03-01 on.

const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

export type DateKind = "date" | "time" | "datetime";

/** Serial for a calendar date. Out-of-range months and days roll over, as in Excel's DATE. */
export function dateToSerial(year: number, month: number, day: number): number {
  return (Date.UTC(year, month - 1, day) - EPOCH) / DAY_MS;
}

export const timeToSerial = (h: number, m: number, s: number) => (h * 3600 + m * 60 + s) / 86_400;

export interface DateParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  /** 0 = Sunday. */
  weekday: number;
}

export function serialToParts(serial: number): DateParts {
  // Round to the second so 0.5 doesn't come out as 11:59:59.999.
  const d = new Date(EPOCH + Math.round(serial * 86_400) * 1000);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds(),
    weekday: d.getUTCDay(),
  };
}

export const daysInMonth = (year: number, month: number) => new Date(Date.UTC(year, month, 0)).getUTCDate();

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const monthIndex = (name: string) => MONTHS.indexOf(name.slice(0, 3).toLowerCase()) + 1;

function validDate(y: number, m: number, d: number) {
  return y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
}

function parseTime(text: string): number | null {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?$/i.exec(text.trim());
  if (!m) return null;
  let h = Number(m[1]);
  const min = Number(m[2]);
  const s = Number(m[3] ?? 0);
  const ampm = m[4]?.[0].toLowerCase();
  if (ampm) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (ampm === "p" ? 12 : 0);
  }
  if (h > 23 || min > 59 || s > 59) return null;
  return timeToSerial(h, min, s);
}

/**
 * Reads typed dates and times: 2025-01-15, 2025/01/15, 15 Jan 2025,
 * Jan 15, 2025, 14:30, 2:30 pm, and a date followed by a time. Day-month
 * order like 01/02/2025 is ambiguous between countries, so it isn't read.
 */
export function parseDateText(text: string): { serial: number; kind: DateKind } | null {
  const t = text.trim();
  if (t.length < 4 || t.length > 40) return null;

  const time = parseTime(t);
  if (time !== null) return { serial: time, kind: "time" };

  let y: number;
  let mo: number;
  let d: number;
  let rest: string;
  let m: RegExpExecArray | null;
  if ((m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(.*)$/.exec(t))) {
    [y, mo, d, rest] = [Number(m[1]), Number(m[2]), Number(m[3]), m[4]];
  } else if ((m = /^(\d{1,2})[ -]([A-Za-z]{3,9})\.?[ -,]+(\d{4})(.*)$/.exec(t))) {
    [y, mo, d, rest] = [Number(m[3]), monthIndex(m[2]), Number(m[1]), m[4]];
  } else if ((m = /^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{4})(.*)$/.exec(t))) {
    [y, mo, d, rest] = [Number(m[3]), monthIndex(m[1]), Number(m[2]), m[4]];
  } else {
    return null;
  }
  if (!validDate(y, mo, d)) return null;
  const date = dateToSerial(y, mo, d);

  rest = rest.replace(/^[T ,]+/, "");
  if (!rest) return { serial: date, kind: "date" };
  const tod = parseTime(rest);
  return tod === null ? null : { serial: date + tod, kind: "datetime" };
}

const pad = (n: number) => String(n).padStart(2, "0");

export type DateStyle = "iso" | "medium" | "full";

/** Shows a serial as a date, time, or both. */
export function formatSerial(serial: number, kind: DateKind, style: DateStyle = "iso", locale?: string): string {
  if (!Number.isFinite(serial) || serial < -657434 || serial > 2958465) return "#NUM!";
  const p = serialToParts(serial);
  const time = `${pad(p.hour)}:${pad(p.minute)}${p.second ? `:${pad(p.second)}` : ""}`;
  if (kind === "time") return time;
  let date: string;
  if (style === "iso") {
    date = `${p.year}-${pad(p.month)}-${pad(p.day)}`;
  } else {
    const at = new Date(Date.UTC(p.year, p.month - 1, p.day));
    date = at.toLocaleDateString(locale, {
      timeZone: "UTC",
      ...(style === "full"
        ? { weekday: "long", day: "numeric", month: "long", year: "numeric" }
        : { day: "numeric", month: "short", year: "numeric" }),
    });
  }
  return kind === "datetime" ? `${date} ${time}` : date;
}

/** Today's date (and the current time, for NOW) in the viewer's time zone. */
export function nowSerial(withTime: boolean): number {
  const now = new Date();
  const date = dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  return withTime ? date + timeToSerial(now.getHours(), now.getMinutes(), now.getSeconds()) : date;
}
