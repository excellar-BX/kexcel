import type { AITab } from "@/components/ai/AIPanel";
import type { BorderMode } from "@/components/ui/icons";
import type { CellAddress, CellFormat, Freeze, NumberFormat, SheetInfo } from "@/lib/engine";

export interface Selection {
  /** The active cell, where the selection started. */
  anchor: CellAddress;
  /** The opposite corner, moved by shift+arrows or dragging. */
  focus: CellAddress;
}

export interface Editing {
  row: number;
  col: number;
  value: string;
  /** "enter": started by typing, arrows commit. "edit": arrows move the caret. */
  mode: "enter" | "edit";
  source: "cell" | "bar";
}

export type ToggleKey = "bold" | "italic" | "underline" | "strike";

/** What the toolbar and menus reflect: the active cell's format and view settings. */
export interface EditorState {
  format: CellFormat;
  zoom: number;
  gridlines: boolean;
  canUndo: boolean;
  canRedo: boolean;
  sheets: SheetInfo[];
  activeSheet: string;
  /** Currency the $ button applies. */
  currency: string;
  freeze: Freeze;
  activeCell: CellAddress;
}

/** Every action the menus, toolbar, shortcuts and right-click menu can trigger. */
export interface Commands {
  undo(): void;
  redo(): void;
  cut(): void;
  copy(): void;
  paste(): void;
  clearValues(): void;
  selectAll(): void;
  find(): void;

  toggle(key: ToggleKey): void;
  applyFormat(patch: CellFormat): void;
  clearFormatting(): void;
  setNumberFormat(format: NumberFormat, currency?: string): void;
  changeDecimals(delta: 1 | -1): void;
  setFontSize(size: number): void;
  setBorders(mode: BorderMode): void;

  insertRows(where: "above" | "below"): void;
  insertCols(where: "left" | "right"): void;
  deleteRows(): void;
  deleteCols(): void;
  fillDown(): void;
  fillRight(): void;
  sort(ascending: boolean): void;
  insertFunction(name: string): void;

  addSheet(): void;
  switchSheet(id: string): void;
  /** Returns an error message when the name can't be used. */
  renameSheet(id: string, name: string): string | null;
  duplicateSheet(id: string): void;
  deleteSheet(id: string): void;
  moveSheet(id: string, delta: -1 | 1): void;

  newWorkbook(): void;
  makeCopy(): void;
  importCSV(): void;
  exportCSV(): void;
  /** Opens an .xlsx or .csv file as a new spreadsheet. */
  openFile(): void;
  exportXlsx(): void;
  print(): void;
  renameWorkbook(): void;

  setZoom(zoom: number): void;
  /** Freezes this many rows or columns; 0 unfreezes. */
  setFreeze(axis: "rows" | "cols", count: number): void;
  toggleGridlines(): void;
  showShortcuts(): void;
  openAI(tab: AITab): void;
}
