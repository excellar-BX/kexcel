"use client";

import { useState } from "react";
import { ContextMenu } from "@/components/ui/Menu";
import { ChevronDown, PlusIcon } from "@/components/ui/icons";
import type { Commands, EditorState } from "@/components/grid/types";

export default function SheetTabs({ commands: c, state: s }: { commands: Commands; state: EditorState }) {
  const [renaming, setRenaming] = useState<{ id: string; value: string; error: string | null } | null>(null);
  const [menu, setMenu] = useState<{ id: string; x: number; y: number } | null>(null);

  const finishRename = () => {
    if (!renaming) return;
    const error = c.renameSheet(renaming.id, renaming.value);
    if (error) setRenaming({ ...renaming, error });
    else setRenaming(null);
  };

  const index = (id: string) => s.sheets.findIndex((x) => x.id === id);

  return (
    <div className="flex min-w-0 items-center gap-1" role="tablist" aria-label="Sheets">
      <button
        type="button"
        onClick={c.addSheet}
        title="Add sheet"
        aria-label="Add sheet"
        className="grid h-7 w-7 shrink-0 place-items-center rounded-md hover:bg-[var(--hover)]"
      >
        <PlusIcon size={16} />
      </button>
      <div className="flex min-w-0 items-center gap-0.5 overflow-x-auto [scrollbar-width:thin]">
        {s.sheets.map((sheet) => {
          const active = sheet.id === s.activeSheet;
          if (renaming?.id === sheet.id) {
            return (
              <div key={sheet.id} className="relative shrink-0">
                <input
                  autoFocus
                  value={renaming.value}
                  onChange={(e) => setRenaming({ ...renaming, value: e.target.value, error: null })}
                  onFocus={(e) => e.target.select()}
                  onBlur={finishRename}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") finishRename();
                    if (e.key === "Escape") setRenaming(null);
                  }}
                  aria-label="Sheet name"
                  aria-invalid={!!renaming.error}
                  className={`h-7 w-32 rounded-md border bg-[var(--bg)] px-2 text-sm outline-none ${
                    renaming.error ? "border-[var(--error)]" : "border-[var(--accent)]"
                  }`}
                />
                {renaming.error && (
                  <div className="absolute bottom-full left-0 mb-1 w-64 rounded-md bg-[var(--fg)] px-2 py-1 text-xs text-[var(--bg)] shadow-lg">
                    {renaming.error}
                  </div>
                )}
              </div>
            );
          }
          return (
            <div
              key={sheet.id}
              className={`flex h-7 shrink-0 items-center rounded-md text-sm ${
                active ? "bg-[var(--header-selected)] font-medium text-[var(--accent)]" : "hover:bg-[var(--hover)]"
              }`}
            >
              <button
                type="button"
                role="tab"
                aria-selected={active}
                onClick={() => c.switchSheet(sheet.id)}
                onDoubleClick={() => setRenaming({ id: sheet.id, value: sheet.name, error: null })}
                onContextMenu={(e) => {
                  e.preventDefault();
                  c.switchSheet(sheet.id);
                  setMenu({ id: sheet.id, x: e.clientX, y: e.clientY });
                }}
                className="max-w-40 truncate py-1 pl-3 pr-1"
              >
                {sheet.name}
              </button>
              <button
                type="button"
                aria-label={`${sheet.name} options`}
                onClick={(e) => {
                  const r = e.currentTarget.getBoundingClientRect();
                  c.switchSheet(sheet.id);
                  setMenu({ id: sheet.id, x: r.left, y: r.top - 8 });
                }}
                className="mr-1 grid h-5 w-5 place-items-center rounded opacity-60 hover:bg-[var(--hover)] hover:opacity-100"
              >
                <ChevronDown size={13} />
              </button>
            </div>
          );
        })}
      </div>

      {menu && (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={[
            {
              label: "Rename",
              onSelect: () => setRenaming({ id: menu.id, value: s.sheets[index(menu.id)]?.name ?? "", error: null }),
            },
            { label: "Duplicate", onSelect: () => c.duplicateSheet(menu.id) },
            { type: "separator" },
            { label: "Move left", disabled: index(menu.id) <= 0, onSelect: () => c.moveSheet(menu.id, -1) },
            {
              label: "Move right",
              disabled: index(menu.id) >= s.sheets.length - 1,
              onSelect: () => c.moveSheet(menu.id, 1),
            },
            { type: "separator" },
            { label: "Delete", disabled: s.sheets.length <= 1, onSelect: () => c.deleteSheet(menu.id) },
          ]}
        />
      )}
    </div>
  );
}
