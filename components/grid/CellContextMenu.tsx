"use client";

import { insertDeleteItems } from "@/components/toolbar/MenuBar";
import { ContextMenu, type MenuEntry } from "@/components/ui/Menu";
import type { Commands } from "./types";

/** Right-click menu for cells, row headers and column headers. */
export default function CellContextMenu({
  x,
  y,
  target,
  commands: c,
  hasComment,
  onClose,
}: {
  x: number;
  y: number;
  target: "cell" | "row" | "col";
  commands: Commands;
  hasComment: boolean;
  onClose: () => void;
}) {
  const clip: MenuEntry[] = [
    { label: "Cut", shortcut: "Ctrl+X", onSelect: c.cut },
    { label: "Copy", shortcut: "Ctrl+C", onSelect: c.copy },
    { label: "Paste", shortcut: "Ctrl+V", onSelect: c.paste },
    { type: "separator" },
  ];
  const [above, below, left, right, , deleteRow, deleteCol] = insertDeleteItems(c);
  const clear: MenuEntry[] = [
    { type: "separator" },
    { label: "Clear values", shortcut: "Delete", onSelect: c.clearValues },
    { label: "Clear formatting", onSelect: c.clearFormatting },
  ];
  const sort: MenuEntry[] = [
    { type: "separator" },
    { label: "Sort A to Z", onSelect: () => c.sort(true) },
    { label: "Sort Z to A", onSelect: () => c.sort(false) },
  ];

  const comment: MenuEntry[] = hasComment
    ? [
        { type: "separator" },
        { label: "Edit comment", shortcut: "Ctrl+Alt+M", onSelect: c.editComment },
        { label: "Delete comment", onSelect: c.deleteComment },
      ]
    : [{ type: "separator" }, { label: "Add comment", shortcut: "Ctrl+Alt+M", onSelect: c.editComment }];

  const items =
    target === "row"
      ? [...clip, above, below, deleteRow, ...clear]
      : target === "col"
        ? [...clip, left, right, deleteCol, ...clear, ...sort]
        : [...clip, above, below, deleteRow, { type: "separator" } as const, left, right, deleteCol, ...clear, ...sort, ...comment];

  return <ContextMenu x={x} y={y} items={items} onClose={onClose} />;
}
