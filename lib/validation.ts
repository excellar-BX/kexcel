import { formatValue, isFormula, nameToRange, type ValidationRule, type Workbook } from "@/lib/engine";

export const MAX_LIST_ITEMS = 500;

/** Splits "Lists!A1:A9" or "'My lists'!A1:A9" into a sheet name and a range. */
function parseSource(source: string): { sheetName: string | null; range: ReturnType<typeof nameToRange> } {
  const text = source.trim().replace(/^=/, "");
  const bang = text.lastIndexOf("!");
  if (bang < 0) return { sheetName: null, range: nameToRange(text) };
  const sheetName = text.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
  return { sheetName, range: nameToRange(text.slice(bang + 1)) };
}

/** Whether a list source like "Lists!A1:A9" points at a real range; returns a message if not. */
export function sourceProblem(wb: Workbook, source: string): string | null {
  const { sheetName, range } = parseSource(source);
  if (!range) return "Write the range like A1:A10, or Lists!A1:A10 for another tab.";
  if (sheetName && !wb.sheets.some((s) => s.name.toLowerCase() === sheetName.toLowerCase())) {
    return `There's no tab called "${sheetName}".`;
  }
  return null;
}

/** The choices a dropdown offers: typed items, or the shown values of its source range, without repeats. */
export function listItems(wb: Workbook, sheet: string, rule: ValidationRule): string[] {
  if (!rule.source) return (rule.items ?? []).filter((x) => x.trim() !== "");
  const { sheetName, range } = parseSource(rule.source);
  const target = sheetName ? wb.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase())?.id : sheet;
  if (!range || !target) return [];
  const out: string[] = [];
  const seen = new Set<string>();
  for (let r = range.r0; r <= range.r1 && out.length < MAX_LIST_ITEMS; r++) {
    for (let c = range.c0; c <= range.c1 && out.length < MAX_LIST_ITEMS; c++) {
      const text = formatValue(wb.getValue(target, r, c), wb.getFormat(target, r, c)).trim();
      if (text && !seen.has(text.toLowerCase())) {
        seen.add(text.toLowerCase());
        out.push(text);
      }
    }
  }
  return out;
}

/** Whether typed content fits a dropdown list. Empty cells always do. */
export function isAllowed(items: string[], text: string): boolean {
  const t = text.trim().toLowerCase();
  return t === "" || items.some((item) => item.trim().toLowerCase() === t);
}

/** Whether a cell holds something its dropdown list doesn't offer. */
export function isInvalid(wb: Workbook, sheet: string, row: number, col: number): boolean {
  const rule = wb.validationAt(sheet, row, col);
  if (!rule) return false;
  const raw = wb.getRaw(sheet, row, col);
  if (raw === "") return false;
  const shown = isFormula(raw) ? formatValue(wb.getValue(sheet, row, col), wb.getFormat(sheet, row, col)) : raw.replace(/^'/, "");
  return !isAllowed(listItems(wb, sheet, rule), shown);
}
