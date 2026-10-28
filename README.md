# Kexcel

A fast, simple spreadsheet for small businesses, students and teams.

## Getting started

```bash
npm install
npm run dev     # http://localhost:3000
npm test        # formula engine tests
```

## AI assistant

The ✨ AI button (and the AI menu) opens a panel that can write a formula from a plain-English request, explain a formula, answer questions about a tab, and check a tab for likely mistakes. Formulas are checked by Kexcel's own engine and previewed before you insert them.

To turn it on, copy `.env.example` to `.env.local`, set `GROQ_API_KEY` (free tier at console.groq.com) or `ANTHROPIC_API_KEY`, and restart `npm run dev`. With both set, Anthropic is used unless `AI_PROVIDER=groq`. Keys stay on the server (`lib/ai/`, `app/api/ai/route.ts`); the browser never sees them. The route has a simple in-memory rate limit, so add sign-in and proper limits before deploying it publicly.

## What works today

- Menu bar (File, Edit, View, Insert, Format, Data, Help) and a formatting toolbar
- Multiple sheet tabs: add, rename, duplicate, reorder, delete; formulas across tabs (`=Sales!B2`, `='My data'!A1:A9`)
- Formatting: bold, italic, underline, strikethrough, font, size, text and fill colour, borders, alignment
- Number formats: number, percent, scientific, currency (10 currencies), more/fewer decimals
- Insert and delete rows and columns; formulas on every tab adjust like in Excel
- Fill handle (drag to continue 1, 2, 3…), Ctrl+D / Ctrl+R fill, sort A→Z / Z→A
- Find (Ctrl+F), name box to jump to a cell, zoom, gridlines toggle, right-click menus
- Copy, cut and paste with formulas and formats, also to and from Excel and Google Sheets
- Undo and redo for edits, formatting, structure and tab changes
- CSV import (into a new tab) and export
- Grid of 10,000 rows × 52 columns that renders only the visible cells; resizable rows and columns, double-click to fit
- About 40 functions with recalculation of only the affected cells and circular reference detection
- Spreadsheets saved automatically in the browser

## Project layout

```
app/                    Next.js routes: home page and /sheet/[id] editor
components/grid/        Editor: state and commands (Spreadsheet), rendering (Grid), layout, find
components/toolbar/     Menu bar, formatting toolbar, sheet tabs, shortcuts dialog
components/ui/          Menus, popovers and icons
lib/engine/             Formula engine in plain TypeScript, no React
  tokenizer.ts          Formula text -> tokens
  parser.ts             Tokens -> syntax tree
  evaluate.ts           Syntax tree -> value
  functions.ts          Built-in functions
  workbook.ts           Sheets, cells, formats, dependency graph, recalculation
  references.ts         Rewrites references for paste, insert/delete, rename
  format.ts             Cell formats and number display
lib/sheet-ops.ts        Fill and sort
lib/storage.ts          localStorage persistence
lib/tsv.ts              Clipboard (TSV) and CSV
```

## Roadmap

1. `.xlsx` import and export, freeze rows and columns, merge cells, wrap text, dates
2. AI: write formulas from plain English, explain formulas, find errors, ask questions of a sheet
3. Accounts, cloud save, sharing, charts, mobile polish
4. Templates for small businesses
