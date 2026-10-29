# Kexcel

A fast, simple spreadsheet for small businesses, students and teams.

## Getting started

```bash
npm install
npm run dev     # http://localhost:3000
npm test        # formula engine tests
```

## Accounts and cloud save

Sign in (top right) and spreadsheets are saved to your account as you type, so they open on any device. Kexcel still keeps a copy in the browser, so it works offline and catches up when the connection returns. If a spreadsheet changed on two devices before syncing, Kexcel asks which version to keep instead of overwriting one.

- Share a spreadsheet (Share button, owners only) with other accounts as Can view or Can edit, or let anyone with the link view it. Open shared spreadsheets pick up other people's saves every few seconds; view-only copies can't be changed.
- Accounts and spreadsheets live in a SQLite file at `data/kexcel.db` (Node's built-in `node:sqlite`; set `KEXCEL_DB_PATH` to move it). It is git-ignored; back it up.
- Passwords are hashed with scrypt; sessions are random tokens stored only as hashes, in an httpOnly cookie. Requests that change data must come from Kexcel's own pages.
- To deploy, use a host with a persistent disk, or swap `lib/server/db.ts` and `lib/server/workbooks.ts` for Postgres. Serverless hosts without a disk will lose the database.

## AI assistant

The AI button (and the AI menu) opens a panel that can make changes for you from a plain-English request (it shows the planned steps first, and applies them as one undoable step), write a formula, explain a formula, answer questions about a tab, and check a tab for likely mistakes. Formulas are checked by Kexcel's own engine and previewed before you insert them.

To turn it on, copy `.env.example` to `.env.local`, set `GROQ_API_KEY` (free tier at console.groq.com) or `ANTHROPIC_API_KEY`, and restart `npm run dev`. With both set, Anthropic is used unless `AI_PROVIDER=groq`. Keys stay on the server (`lib/ai/`, `app/api/ai/route.ts`); the browser never sees them. The AI needs a signed-in account (set `AI_REQUIRE_SIGNIN=false` to allow anyone), with a per-account rate limit.

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
- Open and download Excel files (.xlsx) with values, formulas, formatting, column widths, row heights and frozen panes; CSV import and export
- Pivot tables on a new tab, built from live SUMIFS-style formulas (Data > Pivot table)
- Cell comments (Ctrl+Alt+M or right-click), shown on hover
- Charts (column, bar, line, area, pie) drawn from a range, movable and resizable, with hover details
- Conditional formatting: highlight rules, colour scales and data bars (Format > Conditional formatting)
- Dropdown lists in cells from typed choices or a range, rejecting or flagging other values (Data > Dropdown lists)
- Freeze rows and columns (View > Freeze); merge cells; wrap text (Alt+Enter adds a line in a cell)
- Dates and times as real values: type 2025-01-15, 15 Jan 2025 or 14:30; date formats; TODAY, NOW, DATE, TIME, YEAR, MONTH, DAY, WEEKDAY, EDATE, EOMONTH, DAYS, DATEDIF, NETWORKDAYS, DATEVALUE
- Print with page, scale, gridline and header options
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
  dates.ts              Date serial numbers, parsing and display
lib/sheet-ops.ts        Fill and sort
lib/charts.ts           Chart data from ranges, axis ticks
lib/conditional.ts      Conditional formatting rules
lib/validation.ts       Dropdown list rules
lib/storage.ts          Browser copy of each spreadsheet (offline)
lib/cloud.ts            Account and cloud-save requests; lib/sync.ts reconciles copies
lib/server/             Database, sessions, and the account's spreadsheets (server only)
lib/tsv.ts              Clipboard (TSV) and CSV
lib/xlsx/               Excel .xlsx reader and writer (no network, runs in the browser)
lib/ai/                 AI assistant prompts and provider adapters (server only)
```

## Roadmap

1. Charts, conditional formatting and comments in .xlsx files
2. Password reset by email
3. Live editing together (see each other's cursors and changes instantly)
4. Templates for small businesses
