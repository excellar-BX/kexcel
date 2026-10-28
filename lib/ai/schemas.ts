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

export type FormulaResult = z.infer<typeof FormulaResult>;
export type ExplainResult = z.infer<typeof ExplainResult>;
export type AskResult = z.infer<typeof AskResult>;
export type CheckResult = z.infer<typeof CheckResult>;

export type AIRequest =
  | { action: "status" }
  | { action: "formula"; context: string; prompt: string }
  | { action: "explain"; context: string; cell: string; formula: string }
  | { action: "ask"; context: string; question: string }
  | { action: "check"; context: string };

export type AIResponse =
  | { ok: true; action: "status"; enabled: boolean; model: string }
  | { ok: true; action: "formula"; result: FormulaResult }
  | { ok: true; action: "explain"; result: ExplainResult }
  | { ok: true; action: "ask"; result: AskResult }
  | { ok: true; action: "check"; result: CheckResult }
  | { ok: false; error: string };
