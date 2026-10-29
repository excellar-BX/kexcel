"use client";

import { useState } from "react";
import { nameToRange, type Workbook } from "@/lib/engine";
import { AGGREGATES, pivotFields, type Aggregate, type PivotSpec } from "@/lib/pivot";
import PanelShell, { FieldError, panelInput, panelLabel, primaryButton } from "./PanelShell";

export default function PivotPanel({
  wb,
  sheet,
  defaultRange,
  onCreate,
  onClose,
}: {
  wb: Workbook;
  sheet: string;
  /** The selection or the used area, written like A1:D50. */
  defaultRange: string;
  onCreate: (spec: PivotSpec) => void;
  onClose: () => void;
}) {
  const [range, setRange] = useState(defaultRange);
  const source = nameToRange(range);
  const fields = source && source.r1 > source.r0 ? pivotFields(wb, sheet, source) : [];
  const [rowField, setRowField] = useState<number | null>(null);
  const [colField, setColField] = useState<number | null>(null);
  const [valueField, setValueField] = useState<number | null>(null);
  const [aggregate, setAggregate] = useState<Aggregate>("sum");

  // Until chosen: rows by the first column, values from the last.
  const rows = rowField ?? fields[0]?.col ?? null;
  const values = valueField ?? fields[fields.length - 1]?.col ?? null;
  const problem = !source
    ? "Write the data's range like A1:D50."
    : fields.length === 0
      ? "The range needs a header row and at least one row of data."
      : null;

  const select = (value: number | null, onChange: (v: number | null) => void, allowNone = false) => (
    <select
      value={value ?? ""}
      onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))}
      className={panelInput}
    >
      {allowNone && <option value="">None</option>}
      {fields.map((f) => (
        <option key={f.col} value={f.col}>
          {f.name}
        </option>
      ))}
    </select>
  );

  return (
    <PanelShell title="Pivot table" onClose={onClose}>
      <div className="space-y-3">
        <p className="text-sm text-[var(--muted)]">
          Summarise a table on a new tab, such as total sales for each region. The totals are formulas, so they follow
          changes to the data.
        </p>
        <label className="block">
          <span className={panelLabel}>Data, with its header row</span>
          <input value={range} onChange={(e) => setRange(e.target.value)} className={`${panelInput} font-mono`} />
          {problem && <FieldError>{problem}</FieldError>}
        </label>
        {fields.length > 0 && (
          <>
            <label className="block">
              <span className={panelLabel}>Rows: one for each</span>
              {select(rows, setRowField)}
            </label>
            <label className="block">
              <span className={panelLabel}>Columns: one for each (optional)</span>
              {select(colField, setColField, true)}
            </label>
            <div className="flex gap-2">
              <label className="block w-32 shrink-0">
                <span className={panelLabel}>Show the</span>
                <select value={aggregate} onChange={(e) => setAggregate(e.target.value as Aggregate)} className={panelInput}>
                  {AGGREGATES.map((a) => (
                    <option key={a.value} value={a.value}>
                      {a.label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="block min-w-0 flex-1">
                <span className={panelLabel}>of</span>
                {select(values, setValueField)}
              </label>
            </div>
          </>
        )}
        <button
          type="button"
          disabled={!!problem || rows === null || values === null}
          onClick={() =>
            source && rows !== null && values !== null && onCreate({ source, rowField: rows, colField, valueField: values, aggregate })
          }
          className={primaryButton}
        >
          Create pivot table
        </button>
      </div>
    </PanelShell>
  );
}
