# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Tools → Rename), and delete tabs
- **Per-tab icons** — Master and part tabs each show a customizable icon from **text (Aa)**, **video**, **img**, **sparkle**, **tag**, **folder**, **layers**, **hash**. Right-click the tab to open a small icon picker (part tabs also get Rename / Delete there). Left-click selects/activates the tab only. Chosen icons persist on the tab in the project/session across reload. The **Tools** tab is **wrench-only** (fixed; not in the picker). Master / part / Tools tabs use Soft dark Draft A chrome (**Master** filled teal / library feel, **part** tabs outlined / neutral, **Tools** amber wrench utility chrome)
- **Tools tab** — shared UI tab pinned to the **far right** of the tab bar (after Master/parts and **+**); parks little-used controls (**+ Row**, **+ Column**, **Rename**, and **Part / Column / Row separators**) so the Find/filter row stays uncluttered; actions still apply to the active part tab (highlighted while Tools is open)
- **Find / Replace + part controls** — one compact row on the current part tab: **Find**, **Find All**, **Replace** (Shift+Replace = replace all) plus **All / Non-empty / In Combined**, **Fit row heights**, and **Delete** (wraps on narrow widths); Ctrl/Cmd+F focuses Find. The old “Part name (cols×rows)” caption above the grid is gone (tab title is enough).
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **Excel-style cell names** — columns labeled **A, B, C…**, rows **1, 2, 3…**; empty-cell placeholders and status/aria chrome use addresses like **A1** / **B2**
- **+ Row / + Column** — in the **Tools** tab; grow the active part’s grid; new cells start empty
- **Row ↑ / ↓ move** — arrows beside each row shift that cell row up or down; ↓ pushes every row below further down (grid grows) so nothing is overwritten; ↑ reorders without clobbering; nested cells, Combined link indices, and row heights stay aligned (column widths unchanged)
- **Sheet paste** — paste TSV / HTML tables from Google Sheets or Excel into the focused cell; values spread across the grid and the grid grows if needed
- **Click2Copy** — a short click or focus leaves the text caret where you clicked (**no** auto-copy on activate). **Enter** confirms the cell edit, copies that cell’s text, then moves to the next cell (same as before); **Shift+Enter** inserts a newline and does **not** copy. Combined click-copy / Combined auto-clipboard on edit stay unchanged. Click+drag **inside** one cell selects characters with the normal I-beam (no cell-range hijack). Click+drag **across** cells keeps the full rectangular selection highlighted, copies the range as TSV (tabs between columns, newlines between rows) on mouseup, and **keeps the multi-cell highlight sticky** afterward (clear with Esc, a normal click (including inside the selection), edit, or Tab/arrows). A fresh drag never moves/swaps cells; **Shift+Arrow** also extends the sticky highlight; drag inside that sticky selection moves/swaps the whole block (grab offset preserved; destination clamped in-grid). Typing does not re-copy on every keystroke
- **Cell copy/cut** — with no text selected in a cell, Ctrl/Cmd+C or X copies or cuts the whole cell
- **Sticky Master insert** — Master parts picker sits above the grid; its collapse header stays sticky while scrolling the Master insert list
- **Sticky Column A sort/filter** — grid header row (A–Z / Z–A / Values) stays sticky at the top of the cell grid while scrolling, so those controls remain reachable
- **Master insert greens** — Master library cells get the soft sage/teal Combined outline only when that Master text is an **exact confirmed Combined segment** in the **current part tab’s Combined** (direct Master link or a live confirmed part/nest segment with the same text)—not other tabs’ Combined, not other parts’ segments in Global Combined, and not nest-parent / stale / substring false matches; greens refresh on tab switch and Combined edits
- **Locked Master Combined + part cells** — Master-origin Combined segments are **locked read-only by default** with a **subtle green cue** (soft fill / thin border — not a permanent focus/selection ring); caret typing/deletion/paste into those spans is blocked until unlock. **Insert-from-Master** into a part-tab cell also locks that **grid cell** the same way (`cellLocks`: locked + masterOrigin + masterCellIndex) with a **thin left green accent** (still distinct, without looking always-selected). **Double-click** unlocks Combined or the cell for editing (amber edit chrome). On finish (blur / Enter): **Overwrite Master** (updates the Master library cell and other tabs’ cells still using that Master text, then re-locks), **Keep as local only** (drops Master lock/origin on the Combined segment and/or part cell), or **Cancel** (discards the edit and re-locks). Esc cancels. Locks persist on `confirmedLinks` and per-tab `cellLocks` across reload
- **Master insert Values / A–Z** — same idea as Column A controls: **Values** multi-select filters Master insert rows by first-column text; **A–Z** / **Z–A** reorder the insert list for picking (display only; Master grid data unchanged). The Values menu **stays open** while picking values or using **Clear** / **Select all** (no layout jump; dismiss by clicking elsewhere, Esc, or scrolling the Master insert panel). Selection and sort are **per part tab** — each tab remembers its own prefs, and switching tabs restores that tab’s filter/sort
- **Row filters** — All / Non-empty / In Combined (same row as Find/Replace) hide rows that do not match (data stays intact; empty and Combined inclusion filters)
- **Fit row heights** — one-click control on the Find/Replace row sizes **each row** to the height of its tallest cell (wrapped parent text + nested cells/pages, not equalize-all-rows); **shrinks** tall empty/short rows as well as growing clipped ones (replaces prior `rowHeights` / manual resize pins — not grow-only); measures at live wrap width (border-box, nest chrome beside textarea, tallest nest page), pins wrap height so `overflow:hidden` does not clip, persists as `rowHeights` like column widths; after Fit, typing in a cell/nest live-refits that row only
- **Column A value filter** — Values multi-select on the Column A header picks which unique first-column values to show; composes with All / Non-empty / In Combined (data stays intact)
- **Sort by column A** — A–Z / Z–A buttons on the Column A header reorder all rows by first-column text (case-insensitive); cells, Combined checkboxes, and confirmed-link cell indices stay row-aligned (blank column-A values sink to the bottom)
- **Checkbox drag-paint** — click+drag across Combined cell toggles or row Combined toggles to set many at once to the first control’s new value; Combined and row drags never mix
- **Row / cell Combined toggles** — larger checkbox-like controls on each row and each cell add that content to the Combined prompt **at the Combined caret/selection** (or last caret if focus left Combined; otherwise append at end) and stay checked while included; uncheck removes the linked confirmed segment(s) and clears green. Row toggle covers the whole row; cell toggle is per-cell. Both stay in sync with confirmed links. **Append all active** uses the same insert point. **Match source order** (next to **Global**) when ON reorders Combined green segments to match cell/grid source order (tab order, then row-major cell index, then parent before nests) after each add/remove/re-add; when OFF keeps caret/append order. The toggle persists with the document
- **Nested cells** — **+** under each cell’s Combined checkbox adds an indented nested text field under that cell (multiple nests per parent). Each nest has its **own Combined checkbox** (independent of the parent) and **pages** (◀ ▶, `1/3` label, **+** page); arrow keys at the caret edge also flip pages (▶ past the end adds a page). **Enter** confirms and moves to the next cell (like parents); **Shift+Enter** inserts a newline. Nested content persists with the document. Nest **+** is a quiet Soft dark control (small muted +, not a Combined checkbox tile), stays pinned under the Combined checkbox (not at the bottom of tall autofitted rows), and always attaches to the clicked parent cell; adding/removing a nest refits that row’s height.
- **Append all active** — includes every not-yet-included non-empty cell of the active part (row-major)
- **Undo / Redo** — Ctrl/Cmd+Z undoes document edits (cells, Combined prompt, confirmed links, paste, append, Master insert, clear, etc.); Ctrl/Cmd+Shift+Z or Ctrl+Y redoes. Typing in a cell or Combined is coalesced into one undo step
- **Combined auto-copy** — clicking or focusing the Combined prompt copies its current text to the clipboard (edit still works); any edit also auto-copies; **changing tabs auto-copies the Combined prompt for the tab you land on** (same clipboard behavior). The Copy button remains as a manual fallback. Hold Ctrl/Cmd while activating Combined to skip the activate copy (clipboard stays for paste overwrite)
- **Confirmed append highlighting** — checked / in-Combined cells use **Green alt 2 — soft outline**: soft sage/teal outline ring around the cell (dark interior, no neon green wash) plus Combined checkboxes that match **Global Combined** (**filled teal/green box** + **black checkmark**, same native accent look — not outline + teal check); Nest **+** under each cell is a separate quiet Soft dark control (small muted +, not checkbox-like); Combined segments get the same soft ring. Editing a Combined segment so it no longer matches its linked cell clears the outline on both sides. **Editing a confirmed cell or nest in the table live-replaces that confirmed segment in Combined** (keeps the link); clearing the cell/nest removes it from Combined
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Status / under-the-hood log** — click the status line for a full-width-friendly event panel; long log lines wrap so the window does not need stretching
- **Window orientation** — remembers the BrowserWindow position/size (including half-screen snap/dock) plus maximized/fullscreen across restarts via `userData/window-state.json`; flushes synchronously before quit and **Update ready — Restart** (`app.relaunch` + `app.exit`); restores bounds + maximized/fullscreen on every start (including relaunch) and skips off-screen coordinates if the display layout changed
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
| `styles.css` | Soft dark charcoal theme (teal accents, amber Tools; Combined confirmed = soft sage/teal outline) |
| `renderer.js` | Tabs, cell grid, append, copy, and autosave logic |

## Data shape

Each tab is stored as:

```json
{
  "id": "tab-1",
  "title": "Part 1",
  "icon": "text",
  "cols": 3,
  "rows": 8,
  "cells": ["…", "…"],
  "nestedCells": [[{ "pages": ["…"], "page": 0 }], []],
  "cellLocks": [null, { "masterOrigin": true, "locked": true, "masterCellIndex": 0 }]
}
```

`cells` is a flat row-major array of length `cols * rows`. Older documents may still carry a `sleptCells` array; it is ignored on load. Optional `icon` is one of `text` | `video` | `img` | `sparkle` | `tag` | `folder` | `layers` | `hash` (defaults: Master → `layers`, parts → `text`; Tools is UI-only wrench and is not stored). `nestedCells` is a parallel array of nest lists; each nest is `{ pages: string[], page: number }` (`page` is the current page index). Optional `cellLocks` is a parallel array (same length as `cells`) of `null` or `{ masterOrigin: true, locked: boolean, masterCellIndex?: number }` for Master-inserted part cells. When a cell is checked / Append all, Combined gets **parent text, then each nest’s pages in order** (all non-empty pages), joined by the row separator. Optional `columnWidths` / `rowHeights` arrays (lengths `cols` / `rows`) store resized column widths and auto-fitted row heights in pixels (autofit includes nested height).

Confirmed append links (ranges into the combined prompt tied to source tab/cell, with optional `nestIndex` for nested cells) are stored as `confirmedLinks` on the document. Master-origin segments also store `masterOrigin`, `locked`, and optional `masterCellIndex` so Combined Master locks restore on launch. Part-tab Master locks round-trip on `cellLocks`. They round-trip through session JSON, untitled autosave, named `.c2copy` files, and recoverable backups so cell greens and Master-insert greens restore on launch (offsets are repaired if Combined text still matches).

## Usage tips

1. Edit cells in the active tab’s grid.
2. Use **All** / **Non-empty** / **In Combined** to filter which rows are visible. Use **Values** on the Column A header to multi-select which Column A values to show (composes with the row filters). Use **A–Z** / **Z–A** on the Column A header to sort rows by the first column. The Master parts picker has its own **Values** / **A–Z** / **Z–A** (display-only) for finding insertable Master text; its greens follow exact confirmed segments in the active tab’s Combined (not other tabs / Global cross-part).
3. Check the box left of a row (or the Combined box on a cell) to include it in the Combined prompt at the Combined caret (place the caret first, or it appends at the end); uncheck to remove it. Turn on **Match source order** beside **Global** if you want Combined segments to stay in grid order as you check/uncheck. Use **+** under the Combined checkbox to add indented nested cells (pages via ◀ ▶).
4. Or click **Append all active** to include every missing non-empty cell.
5. Open the shared **Tools** tab for **+ Row** / **+ Column** / **Rename** and separator settings (they still act on the active part tab). Use the part **Find** bar to search or replace text in cells (and nests).
6. Click a cell to place the caret where you clicked (no copy yet). Press **Enter** to confirm, copy that cell, and move to the next (**Shift+Enter** = newline, no copy). Tab/arrows move without copying; **Shift+Arrow** extends a sticky multi-cell highlight. Click+drag inside one cell is native text selection; click+drag across cells selects/highlights a rectangle and copies it as TSV on mouseup; a later drag that leaves the starting cell while inside that sticky selection moves/swaps the block. Paste from Google Sheets or Excel into a focused cell — tab-separated or HTML table data fills the grid (expanding rows/columns as needed) instead of dumping into one cell. Copy/cut with no text selection copies the whole focused cell; Ctrl/Cmd+C while a multi-cell highlight is sticky copies its TSV range. Combined prompt still copies on click/focus/edit. Master-inserted part cells and Master-origin Combined segments stay locked (subtle green cue) until **double-click** unlock; finish with Overwrite Master / Keep local / Cancel.
7. Click **Copy** when ready to paste elsewhere.
8. Double-click a tab title to rename it (or use Tools → Rename); use **Delete** (with confirm if non-empty) to remove a tab. Right-click the tab to pick an icon (Master + parts); left-click selects the tab only. Tools stays wrench-only.

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
