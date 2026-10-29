"use client";

import { useState } from "react";
import { describeRule, HIGHLIGHT_STYLES } from "@/lib/conditional";
import { nameToRange, rangeToName, type ConditionalRule } from "@/lib/engine";
import { newSheetId } from "@/lib/storage";
import PanelShell, { FieldError, panelInput, panelLabel, primaryButton, secondaryButton } from "./PanelShell";

type Kind = ConditionalRule["kind"];

const KINDS: { value: Kind; label: string; operands: 0 | 1 | 2 }[] = [
  { value: "greater", label: "Greater than", operands: 1 },
  { value: "less", label: "Less than", operands: 1 },
  { value: "between", label: "Between", operands: 2 },
  { value: "equal", label: "Equal to", operands: 1 },
  { value: "notEqual", label: "Not equal to", operands: 1 },
  { value: "contains", label: "Text contains", operands: 1 },
  { value: "blank", label: "Is empty", operands: 0 },
  { value: "notBlank", label: "Is not empty", operands: 0 },
  { value: "duplicate", label: "Duplicate values", operands: 0 },
  { value: "scale", label: "Colour scale", operands: 0 },
  { value: "bar", label: "Data bars", operands: 0 },
];

const SCALES: { value: NonNullable<ConditionalRule["scale"]>; label: string; from: string; to: string }[] = [
  { value: "blue", label: "Blue, low to high", from: "--series-1", to: "--series-1" },
  { value: "orange", label: "Orange, low to high", from: "--series-2", to: "--series-2" },
  { value: "diverging", label: "Red below the middle, blue above", from: "--series-8", to: "--series-1" },
];

function Swatch({ rule }: { rule: ConditionalRule }) {
  if (rule.kind === "scale") {
    const s = SCALES.find((x) => x.value === rule.scale) ?? SCALES[0];
    const mid = s.value === "diverging" ? "transparent" : `color-mix(in oklab, var(${s.from}) 35%, transparent)`;
    return (
      <span
        className="h-5 w-10 shrink-0 rounded border border-[var(--border)]"
        style={{
          background: `linear-gradient(90deg, color-mix(in oklab, var(${s.from}) ${s.value === "diverging" ? 70 : 8}%, transparent), ${mid}, color-mix(in oklab, var(${s.to}) 70%, transparent))`,
        }}
      />
    );
  }
  if (rule.kind === "bar") {
    return (
      <span className="flex h-5 w-10 shrink-0 items-center rounded border border-[var(--border)] px-0.5">
        <span className="h-3 w-3/5 rounded-r-sm" style={{ background: "color-mix(in oklab, var(--series-1) 35%, transparent)" }} />
      </span>
    );
  }
  return (
    <span
      className="grid h-5 w-10 shrink-0 place-items-center rounded border border-[var(--border)] text-xs"
      style={{ background: rule.style?.fill, color: rule.style?.color, fontWeight: rule.style?.bold ? 700 : undefined }}
    >
      Aa
    </span>
  );
}

export default function ConditionalPanel({
  rules,
  selection,
  onChange,
  onClose,
}: {
  rules: readonly ConditionalRule[];
  /** The selected range, written like B2:B20, used as the default for new rules. */
  selection: string;
  onChange: (rules: ConditionalRule[]) => void;
  onClose: () => void;
}) {
  const [adding, setAdding] = useState(rules.length === 0);
  const [range, setRange] = useState(selection);
  const [kind, setKind] = useState<Kind>("greater");
  const [a, setA] = useState("");
  const [b, setB] = useState("");
  const [styleIndex, setStyleIndex] = useState(0);
  const [scale, setScale] = useState<NonNullable<ConditionalRule["scale"]>>("blue");
  const [error, setError] = useState<string | null>(null);

  const operands = KINDS.find((k) => k.value === kind)!.operands;

  const save = () => {
    const parsed = nameToRange(range);
    if (!parsed) return setError("Write the cells like B2:B20.");
    if ((operands >= 1 && !a.trim()) || (operands === 2 && !b.trim())) return setError("Fill in the value to compare with.");
    const rule: ConditionalRule = {
      id: newSheetId(),
      range: parsed,
      kind,
      ...(operands ? { values: operands === 2 ? [a.trim(), b.trim()] : [a.trim()] } : {}),
      ...(kind === "scale" ? { scale } : kind === "bar" ? {} : { style: HIGHLIGHT_STYLES[styleIndex].style }),
    };
    onChange([...rules, rule]);
    setAdding(false);
    setError(null);
    setA("");
    setB("");
  };

  const move = (i: number, dir: -1 | 1) => {
    const next = [...rules];
    [next[i], next[i + dir]] = [next[i + dir], next[i]];
    onChange(next);
  };

  return (
    <PanelShell title="Conditional formatting" onClose={onClose}>
      <div className="space-y-3">
        {rules.length > 0 && (
          <>
            <ul className="space-y-2">
              {rules.map((r, i) => (
                <li key={r.id} className="flex items-center gap-2 rounded-lg border border-[var(--border)] p-2">
                  <Swatch rule={r} />
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-sm">{describeRule(r)}</p>
                    <p className="font-mono text-xs text-[var(--muted)]">{rangeToName(r.range)}</p>
                  </div>
                  <button
                    type="button"
                    disabled={i === 0}
                    onClick={() => move(i, -1)}
                    aria-label="Move rule up"
                    className="rounded px-1.5 text-xs text-[var(--muted)] hover:bg-[var(--hover)] disabled:opacity-30"
                  >
                    Up
                  </button>
                  <button
                    type="button"
                    onClick={() => onChange(rules.filter((x) => x.id !== r.id))}
                    className="rounded px-1.5 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
                  >
                    Delete
                  </button>
                </li>
              ))}
            </ul>
            {rules.length > 1 && <p className="text-xs text-[var(--muted)]">Where rules overlap, the higher one wins.</p>}
          </>
        )}

        {!adding ? (
          <button type="button" className={secondaryButton} onClick={() => { setRange(selection); setAdding(true); }}>
            Add a rule
          </button>
        ) : (
          <div className="space-y-3 rounded-lg border border-[var(--border)] p-3">
            <label className="block">
              <span className={panelLabel}>Apply to cells</span>
              <input value={range} onChange={(e) => setRange(e.target.value)} className={`${panelInput} font-mono`} />
            </label>
            <label className="block">
              <span className={panelLabel}>Format cells if</span>
              <select value={kind} onChange={(e) => setKind(e.target.value as Kind)} className={panelInput}>
                {KINDS.map((k) => (
                  <option key={k.value} value={k.value}>
                    {k.label}
                  </option>
                ))}
              </select>
            </label>
            {operands >= 1 && (
              <div className="flex items-center gap-2">
                <input value={a} onChange={(e) => setA(e.target.value)} placeholder="Value" aria-label="Value" className={panelInput} />
                {operands === 2 && (
                  <>
                    <span className="text-sm text-[var(--muted)]">and</span>
                    <input value={b} onChange={(e) => setB(e.target.value)} placeholder="Value" aria-label="Second value" className={panelInput} />
                  </>
                )}
              </div>
            )}
            {kind === "scale" ? (
              <label className="block">
                <span className={panelLabel}>Colours</span>
                <select value={scale} onChange={(e) => setScale(e.target.value as typeof scale)} className={panelInput}>
                  {SCALES.map((s) => (
                    <option key={s.value} value={s.value}>
                      {s.label}
                    </option>
                  ))}
                </select>
              </label>
            ) : kind === "bar" ? (
              <p className="text-xs text-[var(--muted)]">Each cell gets a bar as long as its share of the largest value.</p>
            ) : (
              <fieldset>
                <legend className={panelLabel}>Style</legend>
                <div className="grid grid-cols-2 gap-1.5">
                  {HIGHLIGHT_STYLES.map((h, i) => (
                    <button
                      key={h.label}
                      type="button"
                      aria-pressed={styleIndex === i}
                      title={h.label}
                      onClick={() => setStyleIndex(i)}
                      className={`rounded-md border px-2 py-1.5 text-left text-xs ${
                        styleIndex === i ? "border-[var(--accent)] ring-2 ring-[var(--accent)]/30" : "border-[var(--border)]"
                      }`}
                      style={{ background: h.style.fill, color: h.style.color, fontWeight: h.style.bold ? 700 : undefined }}
                    >
                      {h.label}
                    </button>
                  ))}
                </div>
              </fieldset>
            )}
            {error && <FieldError>{error}</FieldError>}
            <div className="flex gap-2">
              <button type="button" className={primaryButton} onClick={save}>
                Add rule
              </button>
              {rules.length > 0 && (
                <button type="button" className={secondaryButton} onClick={() => setAdding(false)}>
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    </PanelShell>
  );
}
