"use client";

import { useCallback, useRef, useState } from "react";
import { MenuList, useDismiss, type MenuEntry } from "@/components/ui/Menu";
import { CURRENCIES } from "@/lib/engine";
import type { Commands, EditorState } from "@/components/grid/types";

const mod = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";

export const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];

export function numberFormatItems(c: Commands, s: EditorState): MenuEntry[] {
  const fmt = s.format.numFmt ?? "general";
  return [
    { label: "Automatic", checked: fmt === "general", onSelect: () => c.setNumberFormat("general") },
    { label: "Number", shortcut: "1,000.12", checked: fmt === "number", onSelect: () => c.setNumberFormat("number") },
    { label: "Percent", shortcut: "10.12%", checked: fmt === "percent", onSelect: () => c.setNumberFormat("percent") },
    {
      label: "Scientific",
      shortcut: "1.01E+3",
      checked: fmt === "scientific",
      onSelect: () => c.setNumberFormat("scientific"),
    },
    { type: "separator" },
    {
      label: "Currency",
      checked: fmt === "currency",
      items: CURRENCIES.map((cur) => ({
        label: `${cur.label} (${cur.code})`,
        checked: fmt === "currency" && (s.format.currency ?? "USD") === cur.code,
        onSelect: () => c.setNumberFormat("currency", cur.code),
      })),
    },
  ];
}

export function insertDeleteItems(c: Commands): MenuEntry[] {
  return [
    { label: "Insert row above", onSelect: () => c.insertRows("above") },
    { label: "Insert row below", onSelect: () => c.insertRows("below") },
    { label: "Insert column left", onSelect: () => c.insertCols("left") },
    { label: "Insert column right", onSelect: () => c.insertCols("right") },
    { type: "separator" },
    { label: "Delete row", onSelect: c.deleteRows },
    { label: "Delete column", onSelect: c.deleteCols },
  ];
}

const FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "MIN", "MAX", "IF", "SUMIF", "COUNTIF", "VLOOKUP", "ROUND"];

function buildMenus(c: Commands, s: EditorState): { label: string; items: MenuEntry[] }[] {
  return [
    {
      label: "File",
      items: [
        { label: "New spreadsheet", onSelect: c.newWorkbook },
        { label: "Make a copy", onSelect: c.makeCopy },
        { label: "Rename", onSelect: c.renameWorkbook },
        { type: "separator" },
        { label: "Import CSV…", onSelect: c.importCSV },
        { label: "Download as CSV (current tab)", onSelect: c.exportCSV },
        { type: "separator" },
        { label: "Print", shortcut: `${mod}P`, onSelect: c.print },
      ],
    },
    {
      label: "Edit",
      items: [
        { label: "Undo", shortcut: `${mod}Z`, onSelect: c.undo, disabled: !s.canUndo },
        { label: "Redo", shortcut: `${mod}Y`, onSelect: c.redo, disabled: !s.canRedo },
        { type: "separator" },
        { label: "Cut", shortcut: `${mod}X`, onSelect: c.cut },
        { label: "Copy", shortcut: `${mod}C`, onSelect: c.copy },
        { label: "Paste", shortcut: `${mod}V`, onSelect: c.paste },
        { type: "separator" },
        { label: "Fill down", shortcut: `${mod}D`, onSelect: c.fillDown },
        { label: "Fill right", shortcut: `${mod}R`, onSelect: c.fillRight },
        { label: "Delete values", shortcut: "Delete", onSelect: c.clearValues },
        { label: "Select all", shortcut: `${mod}A`, onSelect: c.selectAll },
        { type: "separator" },
        { label: "Find", shortcut: `${mod}F`, onSelect: c.find },
      ],
    },
    {
      label: "View",
      items: [
        { label: "Gridlines", checked: s.gridlines, onSelect: c.toggleGridlines },
        {
          label: "Zoom",
          items: ZOOM_LEVELS.map((z) => ({
            label: `${Math.round(z * 100)}%`,
            checked: s.zoom === z,
            onSelect: () => c.setZoom(z),
          })),
        },
      ],
    },
    {
      label: "Insert",
      items: [
        ...insertDeleteItems(c).slice(0, 4),
        { type: "separator" },
        { label: "Sheet", onSelect: c.addSheet },
        { label: "Function", items: FUNCTIONS.map((f) => ({ label: f, onSelect: () => c.insertFunction(f) })) },
      ],
    },
    {
      label: "Format",
      items: [
        { label: "Number", items: numberFormatItems(c, s) },
        { type: "separator" },
        { label: "Bold", shortcut: `${mod}B`, checked: !!s.format.bold, onSelect: () => c.toggle("bold") },
        { label: "Italic", shortcut: `${mod}I`, checked: !!s.format.italic, onSelect: () => c.toggle("italic") },
        {
          label: "Underline",
          shortcut: `${mod}U`,
          checked: !!s.format.underline,
          onSelect: () => c.toggle("underline"),
        },
        {
          label: "Strikethrough",
          shortcut: "Alt+Shift+5",
          checked: !!s.format.strike,
          onSelect: () => c.toggle("strike"),
        },
        {
          label: "Alignment",
          items: (["left", "center", "right"] as const).map((a) => ({
            label: a[0].toUpperCase() + a.slice(1),
            checked: s.format.align === a,
            onSelect: () => c.applyFormat({ align: a }),
          })),
        },
        { type: "separator" },
        { label: "Clear formatting", shortcut: `${mod}\\`, onSelect: c.clearFormatting },
      ],
    },
    {
      label: "Data",
      items: [
        { label: "Sort range A → Z", onSelect: () => c.sort(true) },
        { label: "Sort range Z → A", onSelect: () => c.sort(false) },
      ],
    },
    {
      label: "Help",
      items: [{ label: "Keyboard shortcuts", shortcut: `${mod}/`, onSelect: c.showShortcuts }],
    },
  ];
}

export default function MenuBar({ commands, state }: { commands: Commands; state: EditorState }) {
  // Dropdowns are positioned against the window so a scrolling parent can't clip them.
  const [open, setOpen] = useState<{ index: number; left: number; top: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(null), []);
  useDismiss(ref, close, open !== null);
  const menus = buildMenus(commands, state);

  const openAt = (index: number, button: HTMLElement) => {
    const rect = button.getBoundingClientRect();
    setOpen({ index, left: rect.left, top: rect.bottom + 2 });
  };

  return (
    <nav ref={ref} className="flex items-center text-sm" aria-label="Menu">
      {menus.map((m, i) => (
        <div key={m.label}>
          <button
            type="button"
            aria-haspopup="menu"
            aria-expanded={open?.index === i}
            onClick={(e) => (open?.index === i ? close() : openAt(i, e.currentTarget))}
            onMouseEnter={(e) => open && open.index !== i && openAt(i, e.currentTarget)}
            className={`rounded px-2 py-0.5 hover:bg-[var(--hover)] ${open?.index === i ? "bg-[var(--hover)]" : ""}`}
          >
            {m.label}
          </button>
          {open?.index === i && (
            <div className="fixed z-50" style={{ left: open.left, top: open.top }}>
              <MenuList items={m.items} onClose={close} keepInView />
            </div>
          )}
        </div>
      ))}
    </nav>
  );
}
