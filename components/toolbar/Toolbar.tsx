"use client";

import { useState, type ReactNode } from "react";
import { MenuList, menuPanelClass, Popover, type MenuEntry } from "@/components/ui/Menu";
import {
  AlignCenterIcon,
  AlignLeftIcon,
  AlignRightIcon,
  BorderIcon,
  ChartIcon,
  ChevronDown,
  ColorBarIcon,
  DecimalDecreaseIcon,
  DecimalIncreaseIcon,
  EraserIcon,
  FillIcon,
  MergeIcon,
  MinusIcon,
  PlusIcon,
  PrintIcon,
  RedoIcon,
  SortIcon,
  SparklesIcon,
  UndoIcon,
  WrapIcon,
  type BorderMode,
} from "@/components/ui/icons";
import { DEFAULT_FONT_SIZE, type FontFamily } from "@/lib/engine";
import type { Commands, EditorState } from "@/components/grid/types";
import { numberFormatItems, ZOOM_LEVELS } from "./MenuBar";

const mod = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";

const PALETTE = [
  ["#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#d9d9d9", "#efefef", "#ffffff"],
  ["#c62828", "#ef6c00", "#f9a825", "#2e7d32", "#00838f", "#1565c0", "#6a1b9a", "#ad1457"],
  ["#ef9a9a", "#ffcc80", "#fff59d", "#a5d6a7", "#80deea", "#90caf9", "#ce93d8", "#f48fb1"],
  ["#fde7e7", "#fff3e0", "#fffde7", "#e8f5e9", "#e0f7fa", "#e3f2fd", "#f3e5f5", "#fce4ec"],
];

const FONTS: { value: FontFamily; label: string }[] = [
  { value: "sans", label: "Default" },
  { value: "serif", label: "Serif" },
  { value: "mono", label: "Monospace" },
];

const FONT_SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 18, 24, 36];
const BORDER_MODES: { mode: BorderMode; label: string }[] = [
  { mode: "all", label: "All borders" },
  { mode: "inner", label: "Inner borders" },
  { mode: "outer", label: "Outer borders" },
  { mode: "top", label: "Top border" },
  { mode: "bottom", label: "Bottom border" },
  { mode: "left", label: "Left border" },
  { mode: "right", label: "Right border" },
  { mode: "none", label: "Clear borders" },
];
const QUICK_FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "MIN", "MAX"];

function Button({
  label,
  onClick,
  active,
  disabled,
  children,
  wide,
}: {
  label: string;
  onClick: (e: React.MouseEvent<HTMLButtonElement>) => void;
  active?: boolean;
  disabled?: boolean;
  children: ReactNode;
  wide?: boolean;
}) {
  return (
    <button
      type="button"
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} // keep focus on the grid
      onClick={onClick}
      className={`flex h-8 shrink-0 items-center justify-center gap-0.5 rounded-md text-[13px] transition-colors hover:bg-[var(--hover)] disabled:opacity-35 disabled:hover:bg-transparent ${
        wide ? "px-2" : "w-8"
      } ${active ? "bg-[var(--header-selected)] text-[var(--accent)]" : ""}`}
    >
      {children}
    </button>
  );
}

const Divider = () => <div className="mx-1 h-5 w-px shrink-0 bg-[var(--border)]" />;

function DropdownButton({ label, title, children }: { label: ReactNode; title: string; children: (close: () => void) => ReactNode }) {
  return (
    <Popover
      button={({ open, toggle }) => (
        <Button label={title} onClick={toggle} active={open} wide>
          {label}
          <ChevronDown size={13} className="text-[var(--muted)]" />
        </Button>
      )}
    >
      {children}
    </Popover>
  );
}

function MenuDropdown({ label, title, items }: { label: ReactNode; title: string; items: MenuEntry[] }) {
  return <DropdownButton label={label} title={title}>{(close) => <MenuList items={items} onClose={close} />}</DropdownButton>;
}

function ColorPicker({ onPick, close, resetLabel }: { onPick: (c: string | undefined) => void; close: () => void; resetLabel: string }) {
  return (
    <div className={`${menuPanelClass} min-w-0 p-2`}>
      <button
        type="button"
        onClick={() => {
          onPick(undefined);
          close();
        }}
        className="mb-2 w-full rounded px-2 py-1 text-left text-sm hover:bg-[var(--hover)]"
      >
        {resetLabel}
      </button>
      <div className="grid grid-cols-8 gap-1">
        {PALETTE.flat().map((c) => (
          <button
            key={c}
            type="button"
            title={c}
            aria-label={`Colour ${c}`}
            onClick={() => {
              onPick(c);
              close();
            }}
            className="h-5 w-5 rounded-full border border-black/10 transition hover:scale-110 dark:border-white/20"
            style={{ background: c }}
          />
        ))}
      </div>
    </div>
  );
}

function FontSizeControl({ size, onChange }: { size: number; onChange: (n: number) => void }) {
  const [text, setText] = useState(String(size));
  const commit = (n: number) => {
    const clamped = Math.min(72, Math.max(6, Math.round(n)));
    setText(String(clamped));
    onChange(clamped);
  };
  const step = (dir: 1 | -1) => {
    const next = dir > 0 ? FONT_SIZES.find((s) => s > size) : [...FONT_SIZES].reverse().find((s) => s < size);
    commit(next ?? size + dir);
  };
  return (
    <div className="flex shrink-0 items-center">
      <Button label="Decrease font size" onClick={() => step(-1)}>
        <MinusIcon size={15} />
      </Button>
      <input
        value={text}
        onChange={(e) => setText(e.target.value.replace(/[^\d]/g, ""))}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit(Number(text) || size);
        }}
        onBlur={() => setText(String(size))}
        aria-label="Font size"
        className="h-7 w-9 rounded-md border border-[var(--border)] bg-transparent text-center text-[13px] outline-none focus:border-[var(--accent)]"
      />
      <Button label="Increase font size" onClick={() => step(1)}>
        <PlusIcon size={15} />
      </Button>
    </div>
  );
}

export default function Toolbar({ commands: c, state: s }: { commands: Commands; state: EditorState }) {
  const f = s.format;
  const size = f.fontSize ?? DEFAULT_FONT_SIZE;
  const align = f.align ?? "left";
  const AlignIcon = align === "center" ? AlignCenterIcon : align === "right" ? AlignRightIcon : AlignLeftIcon;

  return (
    <div
      role="toolbar"
      aria-label="Formatting"
      className="flex items-center gap-0.5 overflow-x-auto rounded-full bg-[var(--toolbar)] px-2 py-1 [scrollbar-width:none]"
    >
      <Button label={`Undo (${mod}Z)`} onClick={c.undo} disabled={!s.canUndo}>
        <UndoIcon size={17} />
      </Button>
      <Button label={`Redo (${mod}Y)`} onClick={c.redo} disabled={!s.canRedo}>
        <RedoIcon size={17} />
      </Button>
      <Button label={`Print (${mod}P)`} onClick={c.print}>
        <PrintIcon size={17} />
      </Button>
      <Divider />

      <MenuDropdown
        title="Zoom"
        label={<span className="w-9 text-left">{Math.round(s.zoom * 100)}%</span>}
        items={ZOOM_LEVELS.map((z) => ({
          label: `${Math.round(z * 100)}%`,
          checked: s.zoom === z,
          onSelect: () => c.setZoom(z),
        }))}
      />
      <Divider />

      <Button
        label={`Format as currency (${s.currency})`}
        onClick={() => c.setNumberFormat("currency", s.currency)}
        active={f.numFmt === "currency"}
      >
        <span className="font-medium">$</span>
      </Button>
      <Button label="Format as percent" onClick={() => c.setNumberFormat("percent")} active={f.numFmt === "percent"}>
        <span className="font-medium">%</span>
      </Button>
      <Button label="Decrease decimal places" onClick={() => c.changeDecimals(-1)}>
        <DecimalDecreaseIcon size={17} />
      </Button>
      <Button label="Increase decimal places" onClick={() => c.changeDecimals(1)}>
        <DecimalIncreaseIcon size={17} />
      </Button>
      <MenuDropdown title="More number formats" label={<span className="text-xs font-medium">123</span>} items={numberFormatItems(c, s)} />
      <Divider />

      <MenuDropdown
        title="Font"
        label={<span className="w-[72px] truncate text-left">{FONTS.find((x) => x.value === (f.fontFamily ?? "sans"))!.label}</span>}
        items={FONTS.map((font) => ({
          label: font.label,
          checked: (f.fontFamily ?? "sans") === font.value,
          onSelect: () => c.applyFormat({ fontFamily: font.value === "sans" ? undefined : font.value }),
        }))}
      />
      <Divider />
      <FontSizeControl key={size} size={size} onChange={c.setFontSize} />
      <Divider />

      <Button label={`Bold (${mod}B)`} onClick={() => c.toggle("bold")} active={!!f.bold}>
        <span className="text-[15px] font-bold">B</span>
      </Button>
      <Button label={`Italic (${mod}I)`} onClick={() => c.toggle("italic")} active={!!f.italic}>
        <span className="font-serif text-[15px] italic">I</span>
      </Button>
      <Button label={`Underline (${mod}U)`} onClick={() => c.toggle("underline")} active={!!f.underline}>
        <span className="text-[15px] underline underline-offset-2">U</span>
      </Button>
      <Button label="Strikethrough (Alt+Shift+5)" onClick={() => c.toggle("strike")} active={!!f.strike}>
        <span className="text-[15px] line-through">S</span>
      </Button>
      <DropdownButton
        title="Text colour"
        label={
          <ColorBarIcon color={f.color ?? "var(--fg)"}>
            <span className="text-[14px] font-semibold">A</span>
          </ColorBarIcon>
        }
      >
        {(close) => <ColorPicker resetLabel="Default colour" close={close} onPick={(color) => c.applyFormat({ color })} />}
      </DropdownButton>
      <DropdownButton
        title="Fill colour"
        label={
          <ColorBarIcon color={f.fill ?? "transparent"}>
            <FillIcon size={15} />
          </ColorBarIcon>
        }
      >
        {(close) => <ColorPicker resetLabel="No fill" close={close} onPick={(fill) => c.applyFormat({ fill })} />}
      </DropdownButton>
      <Divider />

      <DropdownButton title="Borders" label={<BorderIcon mode="all" />}>
        {(close) => (
          <div className={`${menuPanelClass} min-w-0 p-2`}>
            <div className="grid grid-cols-4 gap-1">
              {BORDER_MODES.map((b) => (
                <button
                  key={b.mode}
                  type="button"
                  title={b.label}
                  aria-label={b.label}
                  onClick={() => {
                    c.setBorders(b.mode);
                    close();
                  }}
                  className="grid h-8 w-8 place-items-center rounded hover:bg-[var(--hover)]"
                >
                  <BorderIcon mode={b.mode} />
                </button>
              ))}
            </div>
          </div>
        )}
      </DropdownButton>
      <MenuDropdown
        title="Horizontal align"
        label={<AlignIcon size={17} />}
        items={(["left", "center", "right"] as const).map((a) => ({
          label: a[0].toUpperCase() + a.slice(1),
          checked: f.align === a,
          icon: a === "left" ? <AlignLeftIcon size={15} /> : a === "center" ? <AlignCenterIcon size={15} /> : <AlignRightIcon size={15} />,
          onSelect: () => c.applyFormat({ align: a }),
        }))}
      />
      <Button label="Wrap text" onClick={() => c.toggle("wrap")} active={!!f.wrap}>
        <WrapIcon size={17} />
      </Button>
      <Button
        label={s.merged ? "Unmerge cells" : "Merge cells"}
        onClick={s.merged ? c.unmergeCells : c.mergeCells}
        active={s.merged}
      >
        <MergeIcon size={17} />
      </Button>
      <Divider />

      <Button label={`Clear formatting (${mod}\\)`} onClick={c.clearFormatting}>
        <EraserIcon size={17} />
      </Button>
      <Button label="Sort range A to Z" onClick={() => c.sort(true)}>
        <SortIcon size={17} />
      </Button>
      <MenuDropdown
        title="Functions"
        label={<span className="text-[16px]">Σ</span>}
        items={QUICK_FUNCTIONS.map((fn) => ({ label: fn, onSelect: () => c.insertFunction(fn) }))}
      />
      <Button label="Insert chart" onClick={c.insertChart}>
        <ChartIcon size={17} />
      </Button>
      <Divider />
      <Button label="AI assistant: write, explain and check formulas" onClick={() => c.openAI("formula")} wide>
        <SparklesIcon size={16} className="text-[var(--accent)]" />
        <span className="font-medium">AI</span>
      </Button>
    </div>
  );
}
