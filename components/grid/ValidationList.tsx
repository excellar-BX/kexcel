"use client";

import { useRef, useState } from "react";
import { menuPanelClass, useDismiss } from "@/components/ui/Menu";
import { CheckIcon } from "@/components/ui/icons";

/** The choices of a cell's dropdown list, opened under the cell. Type to filter. */
export default function ValidationList({
  anchor,
  items,
  current,
  onPick,
  onClose,
}: {
  anchor: DOMRect;
  items: string[];
  current: string;
  onPick: (item: string) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  const [filter, setFilter] = useState("");
  const shown = items.filter((x) => x.toLowerCase().includes(filter.trim().toLowerCase()));
  const [index, setIndex] = useState(() => Math.max(0, shown.findIndex((x) => x.toLowerCase() === current.toLowerCase())));
  const below = window.innerHeight - anchor.bottom > 240;

  return (
    <div
      ref={ref}
      className="fixed z-50 w-56"
      style={{
        left: Math.min(anchor.right - 224, window.innerWidth - 232),
        ...(below ? { top: anchor.bottom + 4 } : { bottom: window.innerHeight - anchor.top + 4 }),
      }}
    >
      <div className={`${menuPanelClass} min-w-0`}>
        <input
          autoFocus
          value={filter}
          onChange={(e) => {
            setFilter(e.target.value);
            setIndex(0);
          }}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown") {
              e.preventDefault();
              setIndex((i) => Math.min(shown.length - 1, i + 1));
            } else if (e.key === "ArrowUp") {
              e.preventDefault();
              setIndex((i) => Math.max(0, i - 1));
            } else if (e.key === "Enter") {
              e.preventDefault();
              if (shown[index] !== undefined) onPick(shown[index]);
            } else if (e.key === "Escape") {
              e.preventDefault();
              onClose();
            }
          }}
          placeholder="Search the list"
          aria-label="Search the list"
          className="mx-2 mb-1 w-[calc(100%-1rem)] rounded border border-[var(--border)] bg-[var(--bg)] px-2 py-1 text-sm outline-none focus:border-[var(--accent)]"
        />
        <ul role="listbox" aria-label="Choices" className="max-h-56 overflow-y-auto">
          {shown.length === 0 && <li className="px-3 py-1.5 text-sm text-[var(--muted)]">No matching choices</li>}
          {shown.map((item, i) => (
            <li key={item} role="option" aria-selected={i === index}>
              <button
                type="button"
                onMouseEnter={() => setIndex(i)}
                onClick={() => onPick(item)}
                className={`flex w-full items-center gap-2 px-3 py-1.5 text-left text-sm ${i === index ? "bg-[var(--hover)]" : ""}`}
              >
                <span className="flex w-4 justify-center text-[var(--accent)]">
                  {item.toLowerCase() === current.toLowerCase() && <CheckIcon size={14} />}
                </span>
                <span className="truncate">{item}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
