"use client";

import { useEffect, useState, type ReactNode } from "react";
import { describeSheet, type SheetContext } from "@/lib/ai/context";
import type { AIRequest, AIResponse, AskResult, CheckResult, ExplainResult, FormulaResult } from "@/lib/ai/schemas";
import {
  addressToName,
  formatValue,
  isFormula,
  nameToAddress,
  parseFormula,
  Workbook,
  type CellAddress,
} from "@/lib/engine";
import type { Range } from "@/lib/sheet-ops";

export type AITab = "formula" | "explain" | "ask" | "check";

const TABS: { id: AITab; label: string }[] = [
  { id: "formula", label: "Write" },
  { id: "explain", label: "Explain" },
  { id: "ask", label: "Ask" },
  { id: "check", label: "Check" },
];

async function callAI(req: AIRequest): Promise<AIResponse> {
  try {
    const res = await fetch("/api/ai", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(req),
    });
    return (await res.json()) as AIResponse;
  } catch {
    return { ok: false, error: "Couldn't reach Kexcel's server. Is it still running?" };
  }
}

/** Why a formula can't be used as-is, or null when the engine can read it. */
function formulaProblem(formula: string): string | null {
  if (!isFormula(formula)) return "The AI's answer isn't a formula.";
  try {
    parseFormula(formula.slice(1));
    return null;
  } catch (e) {
    return `Kexcel can't read this formula: ${e instanceof Error ? e.message : "syntax error"}.`;
  }
}

/** What the formula would show if inserted, computed on a copy of the workbook. */
function previewValue(wb: Workbook, sheet: string, cell: CellAddress, formula: string): string {
  const copy = Workbook.fromJSON(wb.toJSON());
  copy.setCell(sheet, cell.row, cell.col, formula);
  return formatValue(copy.getValue(sheet, cell.row, cell.col), copy.getFormat(sheet, cell.row, cell.col)) || "(empty)";
}

function CellLink({ refText, onGoTo }: { refText: string; onGoTo: (ref: string) => void }) {
  return (
    <button
      type="button"
      onClick={() => onGoTo(refText)}
      className="rounded bg-[var(--header-selected)] px-1.5 py-0.5 font-mono text-xs text-[var(--accent)] hover:underline"
    >
      {refText}
    </button>
  );
}

function ErrorNote({ children }: { children: ReactNode }) {
  return (
    <p role="alert" className="rounded-md border border-[var(--error)]/40 bg-[var(--error)]/10 px-3 py-2 text-sm text-[var(--error)]">
      {children}
    </p>
  );
}

const primaryButton =
  "rounded-md bg-[var(--accent)] px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40";
const secondaryButton = "rounded-md border border-[var(--border)] px-3 py-1.5 text-sm hover:bg-[var(--hover)]";
const textArea =
  "w-full resize-none rounded-md border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm outline-none focus:border-[var(--accent)]";

export interface AIPanelProps {
  wb: Workbook;
  sheet: string;
  /** Changes whenever the workbook does. */
  version: number;
  selection: Range;
  active: CellAddress;
  tab: AITab;
  onTab: (tab: AITab) => void;
  onClose: () => void;
  /** Writes content into a cell as an undoable edit and selects it. */
  onInsert: (cell: CellAddress, raw: string) => void;
  /** Selects a cell or range written like "B4", "C2:C9" or "Sales!A1". */
  onGoTo: (ref: string) => void;
}

export default function AIPanel(props: AIPanelProps) {
  const { wb, sheet, selection, active, tab } = props;
  const [status, setStatus] = useState<{ enabled: boolean; model: string } | null>(null);

  useEffect(() => {
    let live = true;
    void callAI({ action: "status" }).then((r) => {
      if (live) setStatus(r.ok && r.action === "status" ? { enabled: r.enabled, model: r.model } : { enabled: false, model: "" });
    });
    return () => {
      live = false;
    };
  }, []);

  const context = (): SheetContext => describeSheet(wb, sheet, selection, active);

  return (
    <aside
      aria-label="AI assistant"
      className="fixed inset-0 z-40 flex flex-col bg-[var(--panel)] sm:static sm:z-auto sm:w-[360px] sm:shrink-0 sm:border-l sm:border-[var(--border)]"
    >
      <div className="flex items-center gap-2 border-b border-[var(--border)] px-3 py-2">
        <span className="text-base">✨</span>
        <h2 className="flex-1 text-sm font-semibold">AI assistant</h2>
        <button type="button" onClick={props.onClose} aria-label="Close AI assistant" className="rounded px-2 py-1 text-sm hover:bg-[var(--hover)]">
          ✕
        </button>
      </div>
      <div role="tablist" className="flex gap-1 border-b border-[var(--border)] px-2 py-1.5">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            aria-selected={tab === t.id}
            onClick={() => props.onTab(t.id)}
            className={`flex-1 rounded-md px-2 py-1 text-sm ${
              tab === t.id ? "bg-[var(--header-selected)] font-medium text-[var(--accent)]" : "hover:bg-[var(--hover)]"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div className="min-h-0 flex-1 overflow-y-auto p-3">
        {status === null ? (
          <p className="text-sm text-[var(--muted)]">Connecting…</p>
        ) : !status.enabled ? (
          <SetupNote />
        ) : (
          <>
            {/* Tabs stay mounted so switching keeps what you typed and the last result. */}
            <div hidden={tab !== "formula"}>
              <FormulaTab {...props} context={context} />
            </div>
            <div hidden={tab !== "explain"}>
              <ExplainTab {...props} context={context} />
            </div>
            <div hidden={tab !== "ask"}>
              <AskTab {...props} context={context} />
            </div>
            <div hidden={tab !== "check"}>
              <CheckTab {...props} context={context} />
            </div>
          </>
        )}
      </div>
      {status?.enabled && (
        <p className="border-t border-[var(--border)] px-3 py-2 text-[11px] text-[var(--muted)]">
          Uses {status.model}. Your sheet&apos;s contents are sent to the AI service with each request. AI can make
          mistakes, so check results before relying on them.
        </p>
      )}
    </aside>
  );
}

function SetupNote() {
  return (
    <div className="space-y-3 text-sm">
      <p className="font-medium">AI isn&apos;t switched on yet.</p>
      <p className="text-[var(--muted)]">To turn it on, add an Anthropic API key to the project:</p>
      <ol className="list-decimal space-y-1.5 pl-5 text-[var(--muted)]">
        <li>
          Create a key at <span className="font-mono text-[var(--fg)]">console.anthropic.com</span>.
        </li>
        <li>
          Add this line to <span className="font-mono text-[var(--fg)]">.env.local</span>:
          <pre className="mt-1 rounded bg-[var(--header-bg)] px-2 py-1 font-mono text-xs text-[var(--fg)]">ANTHROPIC_API_KEY=sk-ant-…</pre>
        </li>
        <li>Restart the dev server.</li>
      </ol>
    </div>
  );
}

type TabProps = AIPanelProps & { context: () => SheetContext };

function TruncatedNote({ ctx }: { ctx: SheetContext | null }) {
  if (!ctx?.truncated) return null;
  const t = ctx.truncated;
  return (
    <p className="text-xs text-[var(--muted)]">
      This tab is large, so the AI only saw the first {t.shownRows} of {t.totalRows} rows and {t.shownCols} of{" "}
      {t.totalCols} columns.
    </p>
  );
}

function FormulaTab({ wb, sheet, active, context, onInsert, onGoTo }: TabProps) {
  const [prompt, setPrompt] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{ value: FormulaResult; target: CellAddress; ctx: SheetContext } | null>(null);

  const submit = async () => {
    if (!prompt.trim() || busy) return;
    setBusy(true);
    setError(null);
    const ctx = context();
    const res = await callAI({ action: "formula", context: ctx.text, prompt });
    setBusy(false);
    if (!res.ok) return setError(res.error);
    if (res.action !== "formula") return;
    const named = res.result.cell ? nameToAddress(res.result.cell.replace(/^.*!/, "")) : null;
    setResult({ value: res.result, target: named ?? active, ctx });
  };

  const problem = result ? formulaProblem(result.value.formula) : null;
  const targetName = result ? addressToName(result.target) : addressToName(active);

  return (
    <div className="space-y-3">
      <label className="block text-sm">
        <span className="mb-1.5 block text-[var(--muted)]">Describe what you want to calculate</span>
        <textarea
          rows={3}
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              void submit();
            }
          }}
          placeholder='e.g. "total sales for Lagos" or "average of column C, ignoring zeros"'
          className={textArea}
        />
      </label>
      <button type="button" onClick={submit} disabled={busy || !prompt.trim()} className={primaryButton}>
        {busy ? "Thinking…" : "Write formula"}
      </button>
      {error && <ErrorNote>{error}</ErrorNote>}
      {result && (
        <div className="space-y-2 rounded-lg border border-[var(--border)] p-3">
          <code className="block break-all rounded bg-[var(--header-bg)] px-2 py-1.5 font-mono text-sm">{result.value.formula}</code>
          <p className="text-sm">{result.value.explanation}</p>
          {problem ? (
            <ErrorNote>{problem} Try rephrasing your request.</ErrorNote>
          ) : (
            <>
              <p className="text-sm text-[var(--muted)]">
                In <CellLink refText={targetName} onGoTo={onGoTo} /> it would show{" "}
                <span className="font-medium text-[var(--fg)]">{previewValue(wb, sheet, result.target, result.value.formula)}</span>
              </p>
              <div className="flex gap-2">
                <button type="button" className={primaryButton} onClick={() => onInsert(result.target, result.value.formula)}>
                  Insert into {targetName}
                </button>
                <button type="button" className={secondaryButton} onClick={() => void navigator.clipboard?.writeText(result.value.formula)}>
                  Copy
                </button>
              </div>
            </>
          )}
          <TruncatedNote ctx={result.ctx} />
        </div>
      )}
    </div>
  );
}

function ExplainTab({ wb, sheet, active, context, onGoTo }: TabProps) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{ cell: string; formula: string; value: ExplainResult } | null>(null);
  const raw = wb.getRaw(sheet, active.row, active.col);
  const cell = addressToName(active);

  if (!isFormula(raw) && !result) {
    return <p className="text-sm text-[var(--muted)]">Select a cell with a formula, then come back here to have it explained.</p>;
  }

  const submit = async () => {
    setBusy(true);
    setError(null);
    const res = await callAI({ action: "explain", context: context().text, cell, formula: raw });
    setBusy(false);
    if (!res.ok) return setError(res.error);
    if (res.action === "explain") setResult({ cell, formula: raw, value: res.result });
  };

  return (
    <div className="space-y-3">
      {isFormula(raw) && (
        <div className="space-y-2">
          <p className="text-sm text-[var(--muted)]">
            <CellLink refText={cell} onGoTo={onGoTo} /> contains
          </p>
          <code className="block break-all rounded bg-[var(--header-bg)] px-2 py-1.5 font-mono text-sm">{raw}</code>
          <button type="button" onClick={submit} disabled={busy} className={primaryButton}>
            {busy ? "Thinking…" : `Explain ${cell}`}
          </button>
        </div>
      )}
      {error && <ErrorNote>{error}</ErrorNote>}
      {result && (
        <div className="space-y-2 rounded-lg border border-[var(--border)] p-3">
          {result.cell !== cell && (
            <p className="text-xs text-[var(--muted)]">
              About <CellLink refText={result.cell} onGoTo={onGoTo} />: <code className="font-mono">{result.formula}</code>
            </p>
          )}
          <p className="text-sm font-medium">{result.value.summary}</p>
          <ol className="list-decimal space-y-1 pl-5 text-sm">
            {result.value.steps.map((s, i) => (
              <li key={i}>{s}</li>
            ))}
          </ol>
        </div>
      )}
    </div>
  );
}

interface Exchange {
  question: string;
  answer?: AskResult;
  error?: string;
}

function AskTab({ context, onGoTo }: TabProps) {
  const [question, setQuestion] = useState("");
  const [history, setHistory] = useState<Exchange[]>([]);
  const [busy, setBusy] = useState(false);
  const [lastCtx, setLastCtx] = useState<SheetContext | null>(null);

  const submit = async () => {
    const q = question.trim();
    if (!q || busy) return;
    setBusy(true);
    setQuestion("");
    const ctx = context();
    setLastCtx(ctx);
    setHistory((h) => [...h, { question: q }]);
    const res = await callAI({ action: "ask", context: ctx.text, question: q });
    setBusy(false);
    setHistory((h) =>
      h.map((x, i) =>
        i === h.length - 1 ? (res.ok && res.action === "ask" ? { ...x, answer: res.result } : { ...x, error: res.ok ? "Unexpected answer." : res.error }) : x,
      ),
    );
  };

  return (
    <div className="space-y-3">
      {history.length === 0 && (
        <p className="text-sm text-[var(--muted)]">
          Ask anything about this tab, like &quot;Which month had the highest sales?&quot; or &quot;How many orders are unpaid?&quot;
        </p>
      )}
      {history.map((x, i) => (
        <div key={i} className="space-y-1.5">
          <p className="ml-6 rounded-lg bg-[var(--header-selected)] px-3 py-2 text-sm">{x.question}</p>
          {x.answer ? (
            <div className="mr-6 space-y-2 rounded-lg border border-[var(--border)] px-3 py-2">
              <p className="whitespace-pre-wrap text-sm">{x.answer.answer}</p>
              {x.answer.cells.length > 0 && (
                <div className="flex flex-wrap items-center gap-1 text-xs text-[var(--muted)]">
                  Based on
                  {x.answer.cells.map((c) => (
                    <CellLink key={c} refText={c} onGoTo={onGoTo} />
                  ))}
                </div>
              )}
            </div>
          ) : x.error ? (
            <ErrorNote>{x.error}</ErrorNote>
          ) : (
            <p className="text-sm text-[var(--muted)]">Thinking…</p>
          )}
        </div>
      ))}
      <TruncatedNote ctx={lastCtx} />
      <div className="sticky bottom-0 space-y-2 bg-[var(--panel)] pt-1">
        <textarea
          rows={2}
          value={question}
          onChange={(e) => setQuestion(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              void submit();
            }
          }}
          placeholder="Ask about this tab"
          aria-label="Question about this tab"
          className={textArea}
        />
        <button type="button" onClick={submit} disabled={busy || !question.trim()} className={primaryButton}>
          {busy ? "Thinking…" : "Ask"}
        </button>
      </div>
    </div>
  );
}

const SEVERITY_STYLE: Record<CheckResult["issues"][number]["severity"], string> = {
  error: "bg-[var(--error)]/15 text-[var(--error)]",
  warning: "bg-amber-500/15 text-amber-700 dark:text-amber-400",
  info: "bg-[var(--header-selected)] text-[var(--accent)]",
};

function CheckTab({ context, onInsert, onGoTo }: TabProps) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{ value: CheckResult; ctx: SheetContext } | null>(null);
  const [applied, setApplied] = useState<Set<number>>(new Set());

  const submit = async () => {
    setBusy(true);
    setError(null);
    const ctx = context();
    const res = await callAI({ action: "check", context: ctx.text });
    setBusy(false);
    if (!res.ok) return setError(res.error);
    if (res.action === "check") {
      setResult({ value: res.result, ctx });
      setApplied(new Set());
    }
  };

  return (
    <div className="space-y-3">
      <p className="text-sm text-[var(--muted)]">
        Looks for broken formulas, ranges that miss data, totals that don&apos;t add up and other likely mistakes.
      </p>
      <button type="button" onClick={submit} disabled={busy} className={primaryButton}>
        {busy ? "Checking…" : result ? "Check again" : "Check this tab"}
      </button>
      {error && <ErrorNote>{error}</ErrorNote>}
      {result && result.value.issues.length === 0 && <p className="text-sm">No problems found. 🎉</p>}
      {result?.value.issues.map((issue, i) => {
        const target = nameToAddress(issue.cell.replace(/^.*!/, ""));
        const fixProblem = issue.fix && isFormula(issue.fix) ? formulaProblem(issue.fix) : null;
        return (
          <div key={i} className="space-y-1.5 rounded-lg border border-[var(--border)] p-3">
            <div className="flex items-center gap-2">
              <span className={`rounded px-1.5 py-0.5 text-[11px] font-semibold uppercase ${SEVERITY_STYLE[issue.severity]}`}>
                {issue.severity}
              </span>
              <CellLink refText={issue.cell} onGoTo={onGoTo} />
            </div>
            <p className="text-sm">{issue.problem}</p>
            <p className="text-sm text-[var(--muted)]">{issue.suggestion}</p>
            {issue.fix !== null && target && !fixProblem && (
              <button
                type="button"
                disabled={applied.has(i)}
                onClick={() => {
                  onInsert(target, issue.fix!);
                  setApplied((s) => new Set(s).add(i));
                }}
                className={secondaryButton}
              >
                {applied.has(i) ? "Fixed ✓" : <>Set to <code className="font-mono">{issue.fix}</code></>}
              </button>
            )}
          </div>
        );
      })}
      {result && <TruncatedNote ctx={result.ctx} />}
    </div>
  );
}
