# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **+ Row / + Column** — grow the grid; new cells start empty
- **Row ↑ / ↓ move** — arrows beside each row shift that cell row up or down; ↓ pushes every row below further down (grid grows) so nothing is overwritten; ↑ reorders without clobbering; sleep flags, nested cells, Combined link indices, and row heights stay aligned (column widths unchanged)
- **Sheet paste** — paste TSV / HTML tables from Google Sheets or Excel into the focused cell; values spread across the grid and the grid grows if needed
- **Click2Copy** — a short click or focus leaves the text caret where you clicked and auto-copies the cell text (Combined / sleep toggles unchanged). Hold Ctrl/Cmd while clicking to skip auto-copy so the clipboard stays intact for paste overwrite. Click+hold+drag a cell onto another to move into an empty target or swap with a filled one (sleep flags and Combined-link indices follow the content). Immediate click+drag across cells keeps the full rectangular selection highlighted, copies the range as TSV (tabs between columns, newlines between rows) on mouseup, and **keeps the multi-cell highlight sticky** afterward (clear with Esc, click outside, edit, or Tab/arrows). Drag inside that sticky selection to move/swap the whole block (grab offset preserved; destination clamped in-grid). Typing does not re-copy on every keystroke
- **Cell copy/cut** — with no text selected in a cell, Ctrl/Cmd+C or X copies or cuts the whole cell
- **Sticky Master insert** — Master parts picker sits above the grid; its collapse header stays sticky while scrolling the Master insert list
- **Sticky Column 1 sort/filter** — grid header row (A–Z / Z–A / Values) stays sticky at the top of the cell grid while scrolling, so those controls remain reachable
- **Master insert greens** — Master library cells turn green when that Master text is represented in the **current part tab’s Combined** (direct Master link or matching confirmed part cell in that Combined scope)—not across other tabs; greens refresh on tab switch and Combined edits
- **Master insert Values / A–Z** — same idea as Column 1 controls: **Values** multi-select filters Master insert rows by first-column text; **A–Z** / **Z–A** reorder the insert list for picking (display only; Master grid data unchanged)
- **Row filters** — All / Non-empty / In Combined buttons hide rows that do not match (data stays intact; empty and Combined inclusion filters)
- **Fit row heights** — one-click toolbar button sizes every row on the active tab to its tallest wrapped content (parent cell + nested cells/pages); uses live wrap width and persists as `rowHeights` like column widths
- **Column 1 value filter** — Values multi-select on the Column 1 header picks which unique first-column values to show; composes with All / Non-empty / In Combined (data stays intact)
- **Sort by column 1** — A–Z / Z–A buttons on the Column 1 header reorder all rows by first-column text (case-insensitive); cells, sleep flags, Combined checkboxes, and confirmed-link cell indices stay row-aligned (blank column-1 values sink to the bottom)
- **Checkbox drag-paint** — click+drag across Combined cell toggles, sleep (Zzz) checkboxes, or row Combined toggles to set many at once to the first control’s new value; Combined, sleep, and row drags never mix
- **Row / cell Combined toggles** — larger checkbox-like controls on each row and each cell add that content to the Combined prompt **at the Combined caret/selection** (or last caret if focus left Combined; otherwise append at end) and stay checked while included; uncheck removes the linked confirmed segment(s) and clears green. Row toggle covers the whole row; cell toggle is per-cell. Both stay in sync with confirmed links. **Append all active** uses the same insert point
- **Sleep cell** — distinct Zzz checkbox above each cell’s Combined toggle marks that cell slept; slept cells stay visible but are skipped by Append all active (state persists with the document)
- **Nested cells** — **+** under each cell’s Combined checkbox adds an indented nested text field under that cell (multiple nests per parent). Each nest has its **own Combined checkbox** (independent of the parent; Append all / sleep follow the parent cell’s sleep flag) and **pages** (◀ ▶, `1/3` label, **+** page); arrow keys at the caret edge also flip pages (▶ past the end adds a page). **Enter** confirms and moves to the next cell (like parents); **Shift+Enter** inserts a newline. Nested content persists with the document
- **Append all active** — includes every not-yet-included non-empty cell of the active part that is not slept (row-major)
- **Undo / Redo** — Ctrl/Cmd+Z undoes document edits (cells, Combined prompt, confirmed links, paste, append, Master insert, clear, etc.); Ctrl/Cmd+Shift+Z or Ctrl+Y redoes. Typing in a cell or Combined is coalesced into one undo step
- **Combined auto-copy** — clicking or focusing the Combined prompt copies its current text to the clipboard (edit still works); any edit also auto-copies (the Copy button remains as a manual fallback). Hold Ctrl/Cmd while activating Combined to skip the activate copy (clipboard stays for paste overwrite)
- **Confirmed append highlighting** — appended cell text is marked green in both the source cell and the Combined prompt (editable “premade” segments). Editing a Combined segment so it no longer matches its linked cell clears green on both sides. **Editing a green cell or nest in the table live-replaces that confirmed segment in Combined** (keeps the link); clearing the cell/nest removes it from Combined
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Status / under-the-hood log** — click the status line for a full-width-friendly event panel; long log lines wrap so the window does not need stretching
- **Window orientation** — remembers the BrowserWindow position/size (including half-screen snap/dock) plus maximized/fullscreen across restarts via `userData/window-state.json`; restores on launch and skips off-screen coordinates if the display layout changed
- **Persistence / autosave** — edits debounce (~250ms) and autosave open documents. Session state lives in Electron `userData` (`click2copy-session.json`). Untitled docs also mirror readable `latest-autosave.c2copy` + `latest-autosave.json` into `userData/backups/` **and** `Documents/Click2Copy/` so you can open or recover them if the app breaks. After **Save As**, autosave writes to that named `.c2copy` path (not a surprise alternate file).
- **Export / Import** — backup and restore all tabs + combined prompt as JSON (survives app updates; userData is outside the repo). Legacy backups with a single `content` string per tab are migrated into cell `[0]` of a 3×8 grid. You can also File → Open the recoverable `latest-autosave.c2copy` from the backup folders.
- **Self-refresh after pull** — watches app source files under the install folder (`package.json`, `main.js`, `preload.js`, `renderer.js`, `index.html`, `styles.css`). After an external `git pull` (or any disk change to those files), a light **Update ready — Restart** banner appears; Restart calls Electron `app.relaunch()` + quit. User-data / Documents autosaves are not watched.

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
| `main.js` | Electron main process (window state + JSON persistence) |
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
  "sleptCells": [false, false],
  "nestedCells": [[{ "pages": ["…"], "page": 0 }], []]
}
```

`cells` is a flat row-major array of length `cols * rows`. `sleptCells` is a parallel boolean array (same length); slept cells are excluded from Append all active. `nestedCells` is a parallel array of nest lists; each nest is `{ pages: string[], page: number }` (`page` is the current page index). When a cell is checked / Append all, Combined gets **parent text, then each nest’s pages in order** (all non-empty pages), joined by the row separator. Optional `columnWidths` / `rowHeights` arrays (lengths `cols` / `rows`) store resized column widths and auto-fitted row heights in pixels (autofit includes nested height).

Confirmed append links (ranges into the combined prompt tied to source tab/cell, with optional `nestIndex` for nested cells) are stored as `confirmedLinks` on the document. They round-trip through session JSON, untitled autosave, named `.c2copy` files, and recoverable backups so cell greens and Master-insert greens restore on launch (offsets are repaired if Combined text still matches).

## Usage tips

1. Edit cells in the active tab’s grid.
2. Use **All** / **Non-empty** / **In Combined** to filter which rows are visible. Use **Values** on the Column 1 header to multi-select which Column 1 values to show (composes with the row filters). Use **A–Z** / **Z–A** on the Column 1 header to sort rows by the first column. The Master parts picker has its own **Values** / **A–Z** / **Z–A** (display-only) for finding insertable Master text; its greens follow the active tab’s Combined.
3. Check the box left of a row (or the Combined box on a cell) to include it in the Combined prompt at the Combined caret (place the caret first, or it appends at the end); uncheck to remove it. Use the Zzz sleep checkbox above a cell’s Combined toggle to exclude that cell from Append all active. Use **+** under the Combined checkbox to add indented nested cells (pages via ◀ ▶).
4. Or click **Append all active** to include every missing non-empty, non-slept cell.
5. Use **+ Row** / **+ Column** to expand the grid.
6. Click a cell to place the caret where you clicked and copy its text (or Tab/Enter/arrow keys onto it). Click+hold+drag a cell onto another to move/swap (sleep + Combined links stay aligned). Click+drag across cells immediately to select a rectangle (full range stays highlighted sticky after mouseup) — on mouseup the range is copied as TSV for pasting into Sheets/Excel; drag the sticky selection to move/swap the block. Paste from Google Sheets or Excel into a focused cell — tab-separated or HTML table data fills the grid (expanding rows/columns as needed) instead of dumping into one cell. Copy/cut with no text selection copies the whole focused cell.
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
