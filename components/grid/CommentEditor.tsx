"use client";

import { useRef, useState } from "react";
import { menuPanelClass, useDismiss } from "@/components/ui/Menu";

/** A small editor for a cell's note, opened beside the cell. Ctrl+Enter saves. */
export default function CommentEditor({
  anchor,
  cell,
  initial,
  onSave,
  onClose,
}: {
  anchor: DOMRect;
  cell: string;
  initial: string;
  onSave: (text: string) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  const [text, setText] = useState(initial);
  const left = Math.min(anchor.right + 8, window.innerWidth - 280);
  const top = Math.min(anchor.top, window.innerHeight - 200);

  return (
    <div ref={ref} className="fixed z-50 w-64" style={{ left, top }}>
      <div className={`${menuPanelClass} min-w-0 space-y-2 p-3`}>
        <p className="text-xs font-medium text-[var(--muted)]">Comment on {cell}</p>
        <textarea
          autoFocus
          rows={4}
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
              e.preventDefault();
              onSave(text);
            } else if (e.key === "Escape") {
              e.preventDefault();
              onClose();
            }
          }}
          placeholder="Write a note for this cell"
          aria-label={`Comment on ${cell}`}
          className="w-full resize-none rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
        />
        <div className="flex gap-2">
          <button
            type="button"
            onClick={() => onSave(text)}
            className="rounded-md bg-[var(--accent)] px-3 py-1 text-sm font-medium text-white hover:opacity-90"
          >
            Save
          </button>
          {initial && (
            <button
              type="button"
              onClick={() => onSave("")}
              className="rounded-md px-3 py-1 text-sm text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
            >
              Delete
            </button>
          )}
          <button type="button" onClick={onClose} className="ml-auto rounded-md px-3 py-1 text-sm hover:bg-[var(--hover)]">
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
