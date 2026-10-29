"use client";

import type { ReactNode } from "react";
import { XIcon } from "@/components/ui/icons";

/** A panel beside the grid on wide screens, or covering it on phones. */
export default function PanelShell({
  title,
  icon,
  onClose,
  children,
  footer,
}: {
  title: string;
  icon?: ReactNode;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
}) {
  return (
    <aside
      aria-label={title}
      className="fixed inset-0 z-40 flex flex-col bg-[var(--panel)] sm:static sm:z-auto sm:w-[360px] sm:shrink-0 sm:border-l sm:border-[var(--border)]"
    >
      <div className="flex items-center gap-2 border-b border-[var(--border)] px-3 py-2">
        {icon}
        <h2 className="flex-1 text-sm font-semibold">{title}</h2>
        <button
          type="button"
          onClick={onClose}
          aria-label={`Close ${title.toLowerCase()}`}
          className="grid h-7 w-7 place-items-center rounded hover:bg-[var(--hover)]"
        >
          <XIcon size={16} />
        </button>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto p-3">{children}</div>
      {footer}
    </aside>
  );
}

export const panelInput =
  "w-full rounded-md border border-[var(--border)] bg-[var(--bg)] px-2.5 py-1.5 text-sm outline-none focus:border-[var(--accent)]";
export const panelLabel = "mb-1 block text-xs font-medium text-[var(--muted)]";
export const primaryButton =
  "rounded-md bg-[var(--accent)] px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40";
export const secondaryButton = "rounded-md border border-[var(--border)] px-3 py-1.5 text-sm hover:bg-[var(--hover)]";

export function FieldError({ children }: { children: ReactNode }) {
  return (
    <p role="alert" className="mt-1 text-xs text-[var(--error)]">
      {children}
    </p>
  );
}
