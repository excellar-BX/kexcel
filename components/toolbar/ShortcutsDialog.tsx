"use client";

import { useRef } from "react";
import { useDismiss } from "@/components/ui/Menu";

const mod = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl";

const GROUPS: { title: string; keys: [string, string][] }[] = [
  {
    title: "Moving around",
    keys: [
      ["Arrow keys", "Move one cell"],
      [`${mod} + arrow`, "Jump to the edge of the data"],
      ["Tab / Shift + Tab", "Move right / left"],
      ["Home", "Start of row"],
      [`${mod} + Home`, "Go to A1"],
      ["Page Up / Page Down", "Move one screen"],
    ],
  },
  {
    title: "Selecting",
    keys: [
      ["Shift + arrow", "Extend the selection"],
      [`${mod} + A`, "Select all"],
      [`${mod} + Space`, "Select column"],
      ["Shift + Space", "Select row"],
    ],
  },
  {
    title: "Editing",
    keys: [
      ["Enter / F2", "Edit the cell"],
      ["Escape", "Cancel editing"],
      ["Delete", "Clear values"],
      [`${mod} + Z / ${mod} + Y`, "Undo / redo"],
      [`${mod} + C / X / V`, "Copy / cut / paste"],
      [`${mod} + D / ${mod} + R`, "Fill down / right"],
      [`${mod} + F`, "Find"],
      [`${mod} + P`, "Print"],
      [`${mod} + O`, "Open a file"],
      [`${mod} + Alt + M`, "Add or edit a comment"],
    ],
  },
  {
    title: "Formatting",
    keys: [
      [`${mod} + B / I / U`, "Bold / italic / underline"],
      ["Alt + Shift + 5", "Strikethrough"],
      [`${mod} + \\`, "Clear formatting"],
    ],
  },
];

export default function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/30 p-4">
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-labelledby="shortcuts-title"
        className="max-h-[85dvh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[var(--border)] bg-[var(--panel)] p-5 shadow-2xl"
      >
        <div className="mb-4 flex items-center justify-between">
          <h2 id="shortcuts-title" className="text-lg font-semibold">
            Keyboard shortcuts
          </h2>
          <button type="button" onClick={onClose} className="rounded px-2 py-1 text-sm hover:bg-[var(--hover)]">
            Close
          </button>
        </div>
        <div className="grid gap-6 sm:grid-cols-2">
          {GROUPS.map((g) => (
            <section key={g.title}>
              <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-[var(--muted)]">{g.title}</h3>
              <dl className="space-y-1.5 text-sm">
                {g.keys.map(([k, v]) => (
                  <div key={k} className="flex justify-between gap-4">
                    <dt className="text-[var(--muted)]">{v}</dt>
                    <dd>
                      <kbd className="rounded border border-[var(--border)] bg-[var(--header-bg)] px-1.5 py-0.5 font-mono text-xs">
                        {k}
                      </kbd>
                    </dd>
                  </div>
                ))}
              </dl>
            </section>
          ))}
        </div>
      </div>
    </div>
  );
}
