# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **+ Row / + Column** — grow the grid; new cells start empty
- **Sheet paste** — paste TSV / HTML tables from Google Sheets or Excel into the focused cell; values spread across the grid and the grid grows if needed
- **Cell copy/cut** — with no text selected in a cell, Ctrl/Cmd+C or X copies or cuts the whole cell
- **Row filters** — All / Non-empty / In Combined buttons hide rows that do not match (data stays intact; empty and Combined inclusion filters)
- **Row / cell Combined toggles** — checkbox-like controls on each row and each cell add that content to the Combined prompt and stay checked while included; uncheck removes the linked confirmed segment(s) and clears green. Row toggle covers the whole row; cell toggle is per-cell. Both stay in sync with confirmed links
- **Append all non-empty** — includes every not-yet-included non-empty cell of the active part (row-major)
- **Undo / Redo** — Ctrl/Cmd+Z undoes document edits (cells, Combined prompt, confirmed links, paste, append, Master insert, clear, etc.); Ctrl/Cmd+Shift+Z or Ctrl+Y redoes. Typing in a cell or Combined is coalesced into one undo step
- **Confirmed append highlighting** — appended cell text is marked green in both the source cell and the Combined prompt (editable “premade” segments). Editing a Combined segment so it no longer matches its linked cell clears green on both sides
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Persistence** — tabs, cell grids, and combined prompt are saved to a local JSON file under Electron `userData` and restored on restart
- **Export / Import** — backup and restore all tabs + combined prompt as JSON (survives app updates; userData is outside the repo). Legacy backups with a single `content` string per tab are migrated into cell `[0]` of a 3×8 grid

## Requirements

- Node.js 18+ recommended
- npm

## Install

```bash
cd click2copy
npm run setup
```

## Run

```bash
npm start
```

This launches Electron with the Click2Copy window.

## Project layout

| File | Role |
|------|------|
| `main.js` | Electron main process (window + JSON persistence) |
| `preload.js` | Secure bridge for load/save IPC |
| `index.html` | UI structure |
| `styles.css` | Dark, practical styling |
| `renderer.js` | Tabs, cell grid, append, copy, and autosave logic |

## Data shape

Each tab is stored as:

```json
{
  "id": "tab-1",
  "title": "Part 1",
  "cols": 3,
  "rows": 8,
  "cells": ["…", "…"]
}
```

`cells` is a flat row-major array of length `cols * rows`.

Confirmed append links (ranges into the combined prompt tied to source tab/cell) are stored as `confirmedLinks` on the document.

## Usage tips

1. Edit cells in the active tab’s grid.
2. Use **All** / **Non-empty** / **In Combined** to filter which rows are visible.
3. Check the box left of a row (or the small box on a cell) to include it in the Combined prompt; uncheck to remove it.
4. Or click **Append all non-empty** to include every missing non-empty cell.
5. Use **+ Row** / **+ Column** to expand the grid.
6. Paste from Google Sheets or Excel into a focused cell — tab-separated or HTML table data fills the grid (expanding rows/columns as needed) instead of dumping into one cell. Copy/cut with no text selection copies the whole focused cell.
7. Click **Copy** when ready to paste elsewhere.
8. Double-click a tab title to rename it; use **Delete** (with confirm if non-empty) to remove a tab.

Data is autosaved shortly after edits.

## Version

App version is tracked in `package.json` as thousandths (`0.001`, `0.002`, …) and bumped on each release push.
