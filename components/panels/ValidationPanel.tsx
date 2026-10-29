"use client";

import { useState } from "react";
import { nameToRange, rangeToName, type ValidationRule, type Workbook } from "@/lib/engine";
import { listItems, sourceProblem } from "@/lib/validation";
import { newSheetId } from "@/lib/storage";
import PanelShell, { FieldError, panelInput, panelLabel, primaryButton, secondaryButton } from "./PanelShell";

export default function ValidationPanel({
  wb,
  sheet,
  rules,
  selection,
  onChange,
  onClose,
}: {
  wb: Workbook;
  sheet: string;
  rules: readonly ValidationRule[];
  selection: string;
  onChange: (rules: ValidationRule[]) => void;
  onClose: () => void;
}) {
  const [adding, setAdding] = useState(rules.length === 0);
  const [range, setRange] = useState(selection);
  const [mode, setMode] = useState<"items" | "range">("items");
  const [items, setItems] = useState("");
  const [source, setSource] = useState("");
  const [strict, setStrict] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const save = () => {
    const parsed = nameToRange(range);
    if (!parsed) return setError("Write the cells like B2:B20.");
    const rule: ValidationRule = { id: newSheetId(), range: parsed, strict };
    if (mode === "items") {
      const list = items.split(/[\n,]/).map((x) => x.trim()).filter(Boolean);
      if (list.length === 0) return setError("Add at least one choice.");
      rule.items = [...new Set(list)];
    } else {
      const problem = sourceProblem(wb, source);
      if (problem) return setError(problem);
      rule.source = source.trim().replace(/^=/, "");
    }
    onChange([...rules, rule]);
    setAdding(false);
    setError(null);
    setItems("");
    setSource("");
  };

  return (
    <PanelShell title="Dropdown lists" onClose={onClose}>
      <div className="space-y-3">
        {rules.length > 0 && (
          <ul className="space-y-2">
            {rules.map((r) => {
              const choices = listItems(wb, sheet, r);
              return (
                <li key={r.id} className="space-y-1 rounded-lg border border-[var(--border)] p-2">
                  <div className="flex items-center gap-2">
                    <span className="flex-1 font-mono text-xs">{rangeToName(r.range)}</span>
                    <span className="text-xs text-[var(--muted)]">{r.strict ? "Rejects other values" : "Marks other values"}</span>
                    <button
                      type="button"
                      onClick={() => onChange(rules.filter((x) => x.id !== r.id))}
                      className="rounded px-1.5 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
                    >
                      Delete
                    </button>
                  </div>
                  <p className="truncate text-sm">
                    {r.source && <span className="text-[var(--muted)]">From {r.source}: </span>}
                    {choices.length ? choices.join(", ") : <span className="text-[var(--muted)]">No choices yet</span>}
                  </p>
                </li>
              );
            })}
          </ul>
        )}

        {!adding ? (
          <button type="button" className={secondaryButton} onClick={() => { setRange(selection); setAdding(true); }}>
            Add a dropdown list
          </button>
        ) : (
          <div className="space-y-3 rounded-lg border border-[var(--border)] p-3">
            <label className="block">
              <span className={panelLabel}>Cells</span>
              <input value={range} onChange={(e) => setRange(e.target.value)} className={`${panelInput} font-mono`} />
            </label>
            <div className="flex gap-1.5" role="radiogroup" aria-label="Where the choices come from">
              {(["items", "range"] as const).map((m) => (
                <button
                  key={m}
                  type="button"
                  role="radio"
                  aria-checked={mode === m}
                  onClick={() => setMode(m)}
                  className={`flex-1 rounded-md border px-2 py-1.5 text-sm ${
                    mode === m ? "border-[var(--accent)] bg-[var(--header-selected)] text-[var(--accent)]" : "border-[var(--border)]"
                  }`}
                >
                  {m === "items" ? "Type the choices" : "Use a range"}
                </button>
              ))}
            </div>
            {mode === "items" ? (
              <label className="block">
                <span className={panelLabel}>Choices, separated by commas or new lines</span>
                <textarea
                  rows={3}
                  value={items}
                  onChange={(e) => setItems(e.target.value)}
                  placeholder="Paid, Unpaid, Refunded"
                  className={`${panelInput} resize-none`}
                />
              </label>
            ) : (
              <label className="block">
                <span className={panelLabel}>Range holding the choices</span>
                <input
                  value={source}
                  onChange={(e) => setSource(e.target.value)}
                  placeholder="Lists!A1:A10"
                  className={`${panelInput} font-mono`}
                />
              </label>
            )}
            <label className="flex items-start gap-2 text-sm">
              <input type="checkbox" checked={strict} onChange={(e) => setStrict(e.target.checked)} className="mt-0.5" />
              <span>
                Reject other values
                <span className="block text-xs text-[var(--muted)]">Otherwise they&apos;re allowed but flagged with a red corner.</span>
              </span>
            </label>
            {error && <FieldError>{error}</FieldError>}
            <div className="flex gap-2">
              <button type="button" className={primaryButton} onClick={save}>
                Add list
              </button>
              {rules.length > 0 && (
                <button type="button" className={secondaryButton} onClick={() => setAdding(false)}>
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}
        <p className="text-xs text-[var(--muted)]">
          In a cell with a list, click the arrow or press Alt+Down to pick a value.
        </p>
      </div>
    </PanelShell>
  );
}
