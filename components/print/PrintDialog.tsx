"use client";

import { useRef, useState } from "react";
import { useDismiss } from "@/components/ui/Menu";
import type { Range } from "@/lib/sheet-ops";
import type { Orientation, Paper, PrintOptions } from "./PrintView";

export type PrintScope = "sheet" | "selection" | "all";

const OPTIONS_KEY = "kexcel:print-options";

/** Remembers the last print settings in this browser; falls back to defaults. */
export function loadOptions(gridlines: boolean): PrintOptions {
  const defaults: PrintOptions = {
    paper: "A4",
    orientation: "portrait",
    fitWidth: true,
    gridlines,
    headers: false,
  };
  try {
    return {
      ...defaults,
      ...JSON.parse(localStorage.getItem(OPTIONS_KEY) ?? "{}"),
    };
  } catch {
    return defaults;
  }
}

function saveOptions(o: PrintOptions) {
  try {
    localStorage.setItem(OPTIONS_KEY, JSON.stringify(o));
  } catch {
    // Settings just won't be remembered.
  }
}

function Choice<T extends string>({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: T;
  options: { value: T; label: string; disabled?: boolean }[];
  onChange: (v: T) => void;
}) {
  return (
    <fieldset>
      <legend className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-[var(--muted)]">{label}</legend>
      <div className="flex flex-wrap gap-1.5">
        {options.map((o) => (
          <button
            key={o.value}
            type="button"
            disabled={o.disabled}
            aria-pressed={value === o.value}
            onClick={() => onChange(o.value)}
            className={`rounded-md border px-3 py-1.5 text-sm transition disabled:opacity-40 ${
              value === o.value
                ? "border-[var(--accent)] bg-[var(--header-selected)] text-[var(--accent)]"
                : "border-[var(--border)] hover:bg-[var(--hover)]"
            }`}
          >
            {o.label}
          </button>
        ))}
      </div>
    </fieldset>
  );
}

export default function PrintDialog({
  gridlines,
  selection,
  sheetCount,
  hasContent,
  onPrint,
  onClose,
}: {
  /** Whether gridlines are showing on screen; the first default for printing them. */
  gridlines: boolean;
  /** The selected range, or null when only one cell is selected. */
  selection: Range | null;
  sheetCount: number;
  /** Whether the chosen scope has anything to print, by scope. */
  hasContent: (scope: PrintScope) => boolean;
  onPrint: (scope: PrintScope, options: PrintOptions) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  const [scope, setScope] = useState<PrintScope>(selection ? "selection" : "sheet");
  const [options, setOptions] = useState(() => loadOptions(gridlines));
  const set = (patch: Partial<PrintOptions>) => setOptions((o) => ({ ...o, ...patch }));
  const empty = !hasContent(scope);

  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/30 p-4 print:hidden">
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-labelledby="print-title"
        className="w-full max-w-md rounded-xl border border-[var(--border)] bg-[var(--panel)] p-5 shadow-2xl"
      >
        <h2 id="print-title" className="mb-4 text-lg font-semibold">
          Print
        </h2>
        <div className="space-y-4">
          <Choice
            label="Print"
            value={scope}
            onChange={setScope}
            options={[
              { value: "sheet", label: "This sheet" },
              {
                value: "selection",
                label: "Selected cells",
                disabled: !selection,
              },
              { value: "all", label: "All sheets", disabled: sheetCount < 2 },
            ]}
          />
          <Choice<Paper>
            label="Paper"
            value={options.paper}
            onChange={(paper) => set({ paper })}
            options={[
              { value: "A4", label: "A4" },
              { value: "Letter", label: "Letter" },
            ]}
          />
          <Choice<Orientation>
            label="Orientation"
            value={options.orientation}
            onChange={(orientation) => set({ orientation })}
            options={[
              { value: "portrait", label: "Portrait" },
              { value: "landscape", label: "Landscape" },
            ]}
          />
          <Choice
            label="Scale"
            value={options.fitWidth ? "fit" : "actual"}
            onChange={(v) => set({ fitWidth: v === "fit" })}
            options={[
              { value: "fit", label: "Fit to page width" },
              { value: "actual", label: "Actual size" },
            ]}
          />
          <div className="space-y-2 text-sm">
            <label className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={options.gridlines}
                onChange={(e) => set({ gridlines: e.target.checked })}
              />
              Show gridlines
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={options.headers} onChange={(e) => set({ headers: e.target.checked })} />
              Show row numbers and column letters
            </label>
          </div>
        </div>

        <div className="mt-6 flex items-center justify-end gap-2">
          {empty && <span className="mr-auto text-xs text-[var(--muted)]">Nothing to print here yet.</span>}
          <button type="button" onClick={onClose} className="rounded-md px-4 py-2 text-sm hover:bg-[var(--hover)]">
            Cancel
          </button>
          <button
            type="button"
            disabled={empty}
            onClick={() => {
              saveOptions(options);
              onPrint(scope, options);
            }}
            className="rounded-md bg-[var(--accent)] px-4 py-2 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40"
          >
            Print
          </button>
        </div>
      </div>
    </div>
  );
}
