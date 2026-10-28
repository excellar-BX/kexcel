import { DEFAULT_FONT_SIZE, parseKey, type Workbook } from "@/lib/engine";
import {
  COL_COUNT,
  COL_HEADER_HEIGHT,
  DEFAULT_COL_WIDTH,
  ROW_COUNT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants";

/** Points to CSS pixels. */
export const ptToPx = (pt: number) => (pt * 4) / 3;

/** Unscaled row height needed to fit text of a given size. */
const heightForFont = (pt: number) => Math.ceil(ptToPx(pt) * 1.25 + 7);

function upperBound(starts: number[], count: number, x: number) {
  let lo = 0;
  let hi = count - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid] <= x) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Positions of every row and column in on-screen pixels, at the current zoom. */
export class GridLayout {
  readonly zoom: number;
  readonly colLefts: number[];
  readonly rowTops: number[];
  readonly headerWidth: number;
  readonly headerHeight: number;
  /** Unscaled widths and heights, as stored in the workbook. */
  readonly baseColWidths: number[];
  readonly baseRowHeights: number[];

  constructor(baseColWidths: number[], baseRowHeights: number[], zoom: number) {
    this.zoom = zoom;
    this.baseColWidths = baseColWidths;
    this.baseRowHeights = baseRowHeights;
    this.colLefts = [0];
    for (const w of baseColWidths) this.colLefts.push(this.colLefts[this.colLefts.length - 1] + w * zoom);
    this.rowTops = [0];
    for (const h of baseRowHeights) this.rowTops.push(this.rowTops[this.rowTops.length - 1] + h * zoom);
    this.headerWidth = Math.round(ROW_HEADER_WIDTH * zoom);
    this.headerHeight = Math.round(COL_HEADER_HEIGHT * zoom);
  }

  static forSheet(wb: Workbook, sheet: string, zoom: number): GridLayout {
    const widths = Array<number>(COL_COUNT).fill(DEFAULT_COL_WIDTH);
    for (const [c, w] of wb.colWidths(sheet)) if (c < COL_COUNT) widths[c] = w;

    const heights = Array<number>(ROW_COUNT).fill(ROW_HEIGHT);
    const manual = wb.rowHeights(sheet);
    // Rows without a manual height grow to fit their largest font.
    for (const key of wb.formatKeys(sheet)) {
      const { row, col } = parseKey(key);
      const size = wb.getFormat(sheet, row, col)?.fontSize;
      if (row < ROW_COUNT && size && size > DEFAULT_FONT_SIZE && !manual.has(row)) {
        heights[row] = Math.max(heights[row], heightForFont(size));
      }
    }
    for (const [r, h] of manual) if (r < ROW_COUNT) heights[r] = h;
    return new GridLayout(widths, heights, zoom);
  }

  get totalWidth() {
    return this.colLefts[COL_COUNT];
  }

  get totalHeight() {
    return this.rowTops[ROW_COUNT];
  }

  colLeft(c: number) {
    return this.colLefts[c];
  }

  colWidth(c: number) {
    return this.colLefts[c + 1] - this.colLefts[c];
  }

  rowTop(r: number) {
    return this.rowTops[r];
  }

  rowHeight(r: number) {
    return this.rowTops[r + 1] - this.rowTops[r];
  }

  /** The column under an x offset within the cell area, clamped to the sheet. */
  colAt(x: number) {
    return upperBound(this.colLefts, COL_COUNT, x);
  }

  /** The row under a y offset within the cell area, clamped to the sheet. */
  rowAt(y: number) {
    return upperBound(this.rowTops, ROW_COUNT, y);
  }
}
