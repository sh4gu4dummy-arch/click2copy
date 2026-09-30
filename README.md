# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **+ Row / + Column** — grow the grid; new cells start empty
- **Sheet paste** — paste TSV / HTML tables from Google Sheets or Excel into the focused cell; values spread across the grid and the grid grows if needed
- **Click2Copy** — clicking or focusing a cell with content auto-copies that cell’s text to the clipboard (editing, focus, and text selection still work; Combined / sleep toggles unchanged). Click+drag across cells keeps the full rectangular selection highlighted (every cell in range) and copies the range as TSV (tabs between columns, newlines between rows) on mouseup. Typing does not re-copy on every keystroke
- **Cell copy/cut** — with no text selected in a cell, Ctrl/Cmd+C or X copies or cuts the whole cell
- **Sticky Master insert** — Master parts picker sits above the grid; its collapse header stays sticky while scrolling the Master insert list
- **Master insert greens** — Master library cells turn green when that Master text is currently represented in Combined (direct Master link or matching confirmed part cell); green clears when unlinked
- **Row filters** — All / Non-empty / In Combined buttons hide rows that do not match (data stays intact; empty and Combined inclusion filters)
- **Column 1 value filter** — Values multi-select on the Column 1 header picks which unique first-column values to show; composes with All / Non-empty / In Combined (data stays intact)
- **Sort by column 1** — A–Z / Z–A buttons on the Column 1 header reorder all rows by first-column text (case-insensitive); cells, sleep flags, Combined checkboxes, and confirmed-link cell indices stay row-aligned (blank column-1 values sink to the bottom)
- **Checkbox drag-paint** — click+drag across Combined cell toggles, sleep (Zzz) checkboxes, or row Combined toggles to set many at once to the first control’s new value; Combined, sleep, and row drags never mix
- **Row / cell Combined toggles** — larger checkbox-like controls on each row and each cell add that content to the Combined prompt and stay checked while included; uncheck removes the linked confirmed segment(s) and clears green. Row toggle covers the whole row; cell toggle is per-cell. Both stay in sync with confirmed links
- **Sleep cell** — distinct Zzz checkbox above each cell’s Combined toggle marks that cell slept; slept cells stay visible but are skipped by Append all active (state persists with the document)
- **Append all active** — includes every not-yet-included non-empty cell of the active part that is not slept (row-major)
- **Undo / Redo** — Ctrl/Cmd+Z undoes document edits (cells, Combined prompt, confirmed links, paste, append, Master insert, clear, etc.); Ctrl/Cmd+Shift+Z or Ctrl+Y redoes. Typing in a cell or Combined is coalesced into one undo step
- **Combined auto-copy** — clicking or focusing the Combined prompt copies its current text to the clipboard (edit still works); any edit also auto-copies (the Copy button remains as a manual fallback)
- **Confirmed append highlighting** — appended cell text is marked green in both the source cell and the Combined prompt (editable “premade” segments). Editing a Combined segment so it no longer matches its linked cell clears green on both sides
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Status / under-the-hood log** — click the status line for a full-width-friendly event panel; long log lines wrap so the window does not need stretching
- **Persistence / autosave** — edits debounce (~250ms) and autosave open documents. Session state lives in Electron `userData` (`click2copy-session.json`). Untitled docs also mirror readable `latest-autosave.c2copy` + `latest-autosave.json` into `userData/backups/` **and** `Documents/Click2Copy/` so you can open or recover them if the app breaks. After **Save As**, autosave writes to that named `.c2copy` path (not a surprise alternate file).
- **Export / Import** — backup and restore all tabs + combined prompt as JSON (survives app updates; userData is outside the repo). Legacy backups with a single `content` string per tab are migrated into cell `[0]` of a 3×8 grid. You can also File → Open the recoverable `latest-autosave.c2copy` from the backup folders.

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
  "cells": ["…", "…"],
  "sleptCells": [false, false]
}
```

`cells` is a flat row-major array of length `cols * rows`. `sleptCells` is a parallel boolean array (same length); slept cells are excluded from Append all active.

Confirmed append links (ranges into the combined prompt tied to source tab/cell) are stored as `confirmedLinks` on the document. They round-trip through session JSON, untitled autosave, named `.c2copy` files, and recoverable backups so cell greens and Master-insert greens restore on launch (offsets are repaired if Combined text still matches).

## Usage tips

1. Edit cells in the active tab’s grid.
2. Use **All** / **Non-empty** / **In Combined** to filter which rows are visible. Use **Values** on the Column 1 header to multi-select which Column 1 values to show (composes with the row filters). Use **A–Z** / **Z–A** on the Column 1 header to sort rows by the first column.
3. Check the box left of a row (or the Combined box on a cell) to include it in the Combined prompt; uncheck to remove it. Use the Zzz sleep checkbox above a cell’s Combined toggle to exclude that cell from Append all active.
4. Or click **Append all active** to include every missing non-empty, non-slept cell.
5. Use **+ Row** / **+ Column** to expand the grid.
6. Click a cell to copy its text to the clipboard (or Tab/Enter onto it). Click+drag across cells to select a rectangle (full range stays highlighted while dragging) — on mouseup the range is copied as TSV for pasting into Sheets/Excel. Paste from Google Sheets or Excel into a focused cell — tab-separated or HTML table data fills the grid (expanding rows/columns as needed) instead of dumping into one cell. Copy/cut with no text selection copies the whole focused cell.
7. Click **Copy** when ready to paste elsewhere.
8. Double-click a tab title to rename it; use **Delete** (with confirm if non-empty) to remove a tab.

Data is autosaved shortly after edits. Click the status line → under-the-hood panel to see the current autosave target and recoverable backup paths (`Documents/Click2Copy/latest-autosave.c2copy` and `userData/backups/latest-autosave.c2copy`).

## Recoverable autosave locations

| Location | Files |
|----------|--------|
| `Documents/Click2Copy/` | `latest-autosave.c2copy`, `latest-autosave.json`, plus per-untitled `autosave-<id>.c2copy` |
| Electron `userData/backups/` | Same readable mirrors (app-local) |
| Electron `userData/` | `click2copy-session.json` (all open docs), `click2copy-data.json` (legacy untitled store) |
| Your Save As path | Autosaved `.c2copy` when the document has a named file |

Open a `.c2copy` with File → Open Prompt Files, or read the plain JSON mirror in any text editor.

## Version

App version is tracked in `package.json` as thousandths (`0.001`, `0.002`, …) and bumped on each release push.
