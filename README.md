# Kexcel

A fast, simple spreadsheet for small businesses, students and teams.

## Getting started

```bash
npm install
npm run dev     # http://localhost:3000
npm test        # formula engine tests
```

## What works today

- Grid of 10,000 rows × 52 columns that renders only the visible cells
- Formula bar, multi-cell selection, column resizing
- Keyboard navigation: arrows, Ctrl+arrows, Tab, Enter, F2, Home, Page Up/Down
- Copy, cut and paste, including to and from Excel and Google Sheets; pasted formulas shift their references
- Undo and redo
- About 40 functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTIF, SUMIF, AVERAGEIF, IF, IFERROR, AND, OR, NOT, VLOOKUP, INDEX, MATCH, ROUND, ROUNDUP, ROUNDDOWN, MOD, LEN, LEFT, RIGHT, MID, CONCAT, TRIM and more
- Recalculation of only the affected cells, with circular reference detection (`#CYCLE!`)
- Sheets saved automatically in the browser's localStorage

## Project layout

```
app/                    Next.js routes: home page and /sheet/[id] editor
components/grid/        Spreadsheet UI (grid, formula bar, status bar)
lib/engine/             Formula engine in plain TypeScript, no React
  tokenizer.ts          Formula text -> tokens
  parser.ts             Tokens -> syntax tree
  evaluate.ts           Syntax tree -> value
  functions.ts          Built-in functions
  sheet.ts              Cell storage, dependency graph, recalculation
lib/storage.ts          localStorage persistence
lib/tsv.ts              Clipboard format
```

## Roadmap

1. Formatting (bold, colours, number and currency formats), multiple tabs, `.xlsx` and `.csv` import and export
2. AI: write formulas from plain English, explain formulas, find errors, ask questions of a sheet
3. Accounts, cloud save, sharing, charts, mobile polish
4. Templates for small businesses
