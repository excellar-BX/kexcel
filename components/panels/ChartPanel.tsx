"use client";

import { useState } from "react";
import { nameToRange, rangeToName, type ChartSpec, type ChartType } from "@/lib/engine";
import PanelShell, { FieldError, panelInput, panelLabel, secondaryButton } from "./PanelShell";

const TYPES: { value: ChartType; label: string; hint: string }[] = [
  { value: "column", label: "Column", hint: "Compare amounts" },
  { value: "bar", label: "Bar", hint: "Compare with long labels" },
  { value: "line", label: "Line", hint: "Change over time" },
  { value: "area", label: "Area", hint: "Totals over time" },
  { value: "pie", label: "Pie", hint: "Parts of a whole" },
];

export default function ChartPanel({
  chart,
  onChange,
  onDelete,
  onClose,
}: {
  chart: ChartSpec;
  onChange: (patch: Partial<ChartSpec>) => void;
  onDelete: () => void;
  onClose: () => void;
}) {
  const [range, setRange] = useState(rangeToName(chart.range));
  const [title, setTitle] = useState(chart.title);
  const [error, setError] = useState<string | null>(null);

  const applyRange = () => {
    const parsed = nameToRange(range);
    if (!parsed) return setError("Write the cells like A1:C12.");
    setError(null);
    onChange({ range: parsed });
  };

  return (
    <PanelShell title="Chart" onClose={onClose}>
      <div className="space-y-4">
        <fieldset>
          <legend className={panelLabel}>Type</legend>
          <div className="grid grid-cols-2 gap-1.5">
            {TYPES.map((t) => (
              <button
                key={t.value}
                type="button"
                aria-pressed={chart.type === t.value}
                onClick={() => onChange({ type: t.value })}
                className={`rounded-md border px-2.5 py-1.5 text-left ${
                  chart.type === t.value
                    ? "border-[var(--accent)] bg-[var(--header-selected)] text-[var(--accent)]"
                    : "border-[var(--border)] hover:bg-[var(--hover)]"
                }`}
              >
                <span className="block text-sm font-medium">{t.label}</span>
                <span className="block text-[11px] text-[var(--muted)]">{t.hint}</span>
              </button>
            ))}
          </div>
        </fieldset>
        <label className="block">
          <span className={panelLabel}>Title</span>
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            onBlur={() => title !== chart.title && onChange({ title })}
            onKeyDown={(e) => e.key === "Enter" && onChange({ title })}
            className={panelInput}
          />
        </label>
        <label className="block">
          <span className={panelLabel}>Data</span>
          <input
            value={range}
            onChange={(e) => setRange(e.target.value)}
            onBlur={applyRange}
            onKeyDown={(e) => e.key === "Enter" && applyRange()}
            className={`${panelInput} font-mono`}
          />
          {error && <FieldError>{error}</FieldError>}
        </label>
        <p className="text-xs text-[var(--muted)]">
          Each column is a series and each row a category. A first row of text names the series; a first column of
          text labels the categories. Pie charts use the first series.
        </p>
        <button type="button" onClick={onDelete} className={`${secondaryButton} text-[var(--error)]`}>
          Delete chart
        </button>
      </div>
    </PanelShell>
  );
}
