"use client";

import { useCallback, useEffect, useLayoutEffect, useRef, useState, type ReactNode, type RefObject } from "react";
import { CheckIcon, ChevronRight } from "./icons";

export type MenuEntry =
  | {
      type?: "item";
      label: string;
      shortcut?: string;
      onSelect?: () => void;
      disabled?: boolean;
      checked?: boolean;
      /** Opens a submenu instead of acting. */
      items?: MenuEntry[];
      icon?: ReactNode;
    }
  | { type: "separator" };

/** Calls `onClose` on a click outside `ref` or on Escape, while `active`. */
export function useDismiss(ref: RefObject<HTMLElement | null>, onClose: () => void, active = true) {
  useEffect(() => {
    if (!active) return;
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.stopPropagation();
        onClose();
      }
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey, true);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey, true);
    };
  }, [ref, onClose, active]);
}

/** Keeps a floating element inside the window by nudging it left or up. */
function useKeepInView(ref: RefObject<HTMLElement | null>) {
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const margin = 8;
    const dx = Math.min(0, window.innerWidth - margin - rect.right);
    let dy = Math.min(0, window.innerHeight - margin - rect.bottom);
    if (rect.top + dy < margin) {
      dy = margin - rect.top;
      el.style.maxHeight = `${window.innerHeight - 2 * margin}px`;
    }
    if (dx || dy) el.style.transform = `translate(${dx}px, ${dy}px)`;
  }, [ref]);
}

export const menuPanelClass =
  "min-w-[220px] overflow-y-auto rounded-lg border border-[var(--border)] bg-[var(--panel)] py-1 text-sm shadow-xl";

export function MenuList({ items, onClose }: { items: MenuEntry[]; onClose: () => void }) {
  const [openSub, setOpenSub] = useState<number | null>(null);
  const listRef = useRef<HTMLDivElement>(null);
  useKeepInView(listRef);

  const focusItem = (dir: 1 | -1) => {
    const buttons = [...(listRef.current?.querySelectorAll<HTMLButtonElement>(":scope > div > button:not(:disabled)") ?? [])];
    if (buttons.length === 0) return;
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    buttons[(i + dir + buttons.length) % buttons.length].focus();
  };

  return (
    <div
      ref={listRef}
      role="menu"
      className={menuPanelClass}
      onKeyDown={(e) => {
        if (e.key === "ArrowDown") {
          e.preventDefault();
          focusItem(1);
        } else if (e.key === "ArrowUp") {
          e.preventDefault();
          focusItem(-1);
        }
      }}
    >
      {items.map((item, i) => {
        if (item.type === "separator") {
          return <div key={i} role="separator" className="my-1 border-t border-[var(--border)]" />;
        }
        const hasSub = !!item.items?.length;
        return (
          <div
            key={i}
            className="relative"
            onMouseEnter={() => setOpenSub(hasSub ? i : null)}
            onMouseLeave={() => hasSub && setOpenSub(null)}
          >
            <button
              type="button"
              role="menuitem"
              disabled={item.disabled}
              aria-haspopup={hasSub || undefined}
              aria-expanded={hasSub ? openSub === i : undefined}
              onClick={() => {
                if (hasSub) {
                  setOpenSub(openSub === i ? null : i);
                  return;
                }
                onClose();
                item.onSelect?.();
              }}
              onKeyDown={(e) => {
                if (hasSub && e.key === "ArrowRight") {
                  e.preventDefault();
                  setOpenSub(i);
                }
              }}
              className="flex w-full items-center gap-2 px-3 py-1.5 text-left outline-none hover:bg-[var(--hover)] focus-visible:bg-[var(--hover)] disabled:cursor-default disabled:opacity-40 disabled:hover:bg-transparent"
            >
              <span className="flex w-4 shrink-0 justify-center text-[var(--muted)]">
                {item.checked ? <CheckIcon size={15} /> : item.icon}
              </span>
              <span className="flex-1 whitespace-nowrap">{item.label}</span>
              {item.shortcut && <span className="pl-6 text-xs text-[var(--muted)]">{item.shortcut}</span>}
              {hasSub && <ChevronRight size={14} className="text-[var(--muted)]" />}
            </button>
            {hasSub && openSub === i && (
              <div
                className="absolute left-full top-0 z-10 -mt-1"
                onKeyDown={(e) => {
                  if (e.key === "ArrowLeft") {
                    e.stopPropagation();
                    setOpenSub(null);
                  }
                }}
              >
                <MenuList items={item.items!} onClose={onClose} />
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

/**
 * A button that toggles a floating panel below it. The panel is positioned
 * against the window so a scrolling toolbar can't clip it.
 */
export function Popover({
  button,
  children,
}: {
  button: (props: { open: boolean; toggle: (e: React.MouseEvent<HTMLElement>) => void }) => ReactNode;
  children: (close: () => void) => ReactNode;
}) {
  const [at, setAt] = useState<{ left: number; top: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setAt(null), []);
  useDismiss(ref, close, at !== null);

  const toggle = (e: React.MouseEvent<HTMLElement>) => {
    if (at) return setAt(null);
    const rect = e.currentTarget.getBoundingClientRect();
    setAt({ left: rect.left, top: rect.bottom + 4 });
  };

  return (
    <div ref={ref}>
      {button({ open: at !== null, toggle })}
      {at && (
        <div className="fixed z-50" style={at}>
          <FloatingPanel>{children(close)}</FloatingPanel>
        </div>
      )}
    </div>
  );
}

function FloatingPanel({ children }: { children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null);
  useKeepInView(ref);
  return <div ref={ref}>{children}</div>;
}

/** A menu that opens at a point, such as a right-click menu. */
export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuEntry[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  return (
    <div ref={ref} className="fixed z-50" style={{ left: x, top: y }} onContextMenu={(e) => e.preventDefault()}>
      <MenuList items={items} onClose={onClose} />
    </div>
  );
}
