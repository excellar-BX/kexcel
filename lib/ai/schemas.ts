import { z } from "zod";

// Shapes the model must answer in, one per AI action. The server enforces
// them with structured outputs; the client imports only the inferred types.

export const FormulaResult = z.object({
  formula: z.string().describe("The formula, starting with ="),
  explanation: z.string().describe("One or two plain sentences on what it calculates"),
  cell: z.string().nullable().describe("Cell the formula belongs in, like D2, or null for the active cell"),
});

export const ExplainResult = z.object({
  summary: z.string().describe("What the formula calculates, in one or two plain sentences"),
  steps: z.array(z.string()).describe("Short steps walking through how it works, in order"),
});

export const AskResult = z.object({
  answer: z.string().describe("The answer in plain language, a few sentences at most"),
  cells: z.array(z.string()).describe("Cells or ranges the answer is based on, like B4 or C2:C9"),
});

export const CheckResult = z.object({
  issues: z.array(
    z.object({
      cell: z.string().describe("The cell with the problem, like C7"),
      severity: z.enum(["error", "warning", "info"]),
      problem: z.string().describe("What is wrong, in one sentence"),
      suggestion: z.string().describe("How to fix it, in one sentence"),
      fix: z.string().nullable().describe("Exact replacement content for the cell, or null if it needs a person"),
    }),
  ),
});

/** One step the AI can take on the workbook. Which fields matter depends on `type`. */
export const SheetAction = z.object({
  type: z.enum([
    "set_values",
    "format",
    "clear",
    "clear_format",
    "insert_rows",
    "insert_columns",
    "delete_rows",
    "delete_columns",
    "sort",
    "add_sheet",
    "switch_sheet",
    "rename_sheet",
    "chart",
    "highlight",
    "color_scale",
    "data_bars",
    "dropdown",
    "freeze",
    "merge",
    "unmerge",
    "column_width",
    "comment",
    "pivot",
  ]),
  description: z.string().describe("One short sentence telling the user what this step does"),
  range: z.string().optional().describe("Cells, rows or columns this step applies to, like B2:D9, B5, 5 or C"),
  values: z
    .array(z.array(z.string()))
    .optional()
    .describe("set_values: rows of cell contents from the top-left of range; formulas start with ="),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fill: z.string().optional().describe("Background colour as #rrggbb"),
  text_color: z.string().optional().describe("Text colour as #rrggbb"),
  align: z.enum(["left", "center", "right"]).optional(),
  number_format: z.enum(["general", "number", "currency", "percent", "date", "time", "datetime"]).optional(),
  decimals: z.number().int().min(0).max(10).optional(),
  currency: z.string().optional().describe("ISO currency code such as USD or NGN"),
  wrap: z.boolean().optional(),
  count: z.number().int().min(1).max(1000).optional().describe("How many rows or columns to insert or delete"),
  column: z.string().optional().describe("sort: the column letter to sort by"),
  ascending: z.boolean().optional(),
  name: z.string().optional().describe("Tab name for add_sheet, switch_sheet and rename_sheet"),
  chart_type: z.enum(["column", "bar", "line", "area", "pie"]).optional(),
  title: z.string().optional(),
  condition: z
    .enum(["greater", "less", "between", "equal", "notEqual", "contains", "blank", "notBlank", "duplicate"])
    .optional(),
  compare_values: z.array(z.string()).optional().describe("highlight: the value (or two for between) to compare with"),
  style: z.enum(["red", "yellow", "green", "blue", "red_text", "bold"]).optional(),
  scale: z.enum(["blue", "orange", "diverging"]).optional(),
  items: z.array(z.string()).optional().describe("dropdown: the choices"),
  rows: z.number().int().min(0).max(100).optional().describe("freeze: rows to freeze"),
  columns: z.number().int().min(0).max(20).optional().describe("freeze: columns to freeze"),
  width: z.number().int().min(20).max(800).optional().describe("column_width: pixels"),
  text: z.string().optional().describe("comment: the note's text"),
  row_field: z.string().optional().describe("pivot: header of the column to group rows by"),
  column_field: z.string().optional().describe("pivot: optional header to spread across columns"),
  value_field: z.string().optional().describe("pivot: header of the column to total"),
  aggregate: z.enum(["sum", "count", "average", "min", "max"]).optional(),
});

export const PlanResult = z.object({
  summary: z.string().describe("A sentence or two on what the plan does, or why it can't be done"),
  actions: z.array(SheetAction).max(40),
});

export type SheetAction = z.infer<typeof SheetAction>;
export type PlanResult = z.infer<typeof PlanResult>;
export type FormulaResult = z.infer<typeof FormulaResult>;
export type ExplainResult = z.infer<typeof ExplainResult>;
export type AskResult = z.infer<typeof AskResult>;
export type CheckResult = z.infer<typeof CheckResult>;

export type AIRequest =
  | { action: "status" }
  | { action: "formula"; context: string; prompt: string }
  | { action: "explain"; context: string; cell: string; formula: string }
  | { action: "ask"; context: string; question: string }
  | { action: "check"; context: string }
  | { action: "plan"; context: string; request: string };

export type AIResponse =
  | { ok: true; action: "status"; enabled: boolean; provider: "anthropic" | "groq" | null; model: string }
  | { ok: true; action: "formula"; result: FormulaResult }
  | { ok: true; action: "explain"; result: ExplainResult }
  | { ok: true; action: "ask"; result: AskResult }
  | { ok: true; action: "check"; result: CheckResult }
  | { ok: true; action: "plan"; result: PlanResult }
  | { ok: false; error: string };
