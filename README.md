# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Tools → Rename), and delete tabs
- **Per-tab icons** — Master and part tabs each show a customizable icon from **text (Aa)**, **video**, **img**, **sparkle**, **tag**, **folder**, **layers**, **hash**. Right-click the tab to open a small icon picker (part tabs also get Rename / Delete there). Left-click selects/activates the tab only. Chosen icons persist on the tab in the project/session across reload. The **Tools** tab is **wrench-only** (fixed; not in the picker). Master / part / Tools tabs use Soft dark Draft A chrome (**Master** and **part** tabs outlined / neutral with the same inactive/active colors, **Tools** amber wrench utility chrome)
- **Tools tab** — shared UI tab pinned to the **far right** of the tab bar (after Master/parts and **+**); parks little-used controls (**+ Row**, **+ Column**, **Rename**, and **Part / Column / Row separators**) so the Find/filter row stays uncluttered; actions still apply to the active part tab (highlighted while Tools is open). Committing a separator (leave the field) rewrites existing confirmed Combined text with the new part/column/row separators — plain text before the first and after the last confirmed segment is kept; **Match source order** also re-sorts, otherwise caret/append order stays
- **Find / Replace + part controls** — **one Soft dark line** on the current part tab (compact controls sized to fit ~1000px+ without wrapping): part-page chrome (`◀ n/n ▶` + page chips — right-click a chip to **Rename**, **Add page after this** (blank), **Duplicate this page** (copy including Master locks, inserted right after), or **Delete page**; there is always at least one chip), **Find** / **Replace** boxes (no buttons: Find box Enter = next match, Shift+Enter = highlight all; Replace box Enter = replace current, Shift+Enter = replace all; hover either box for the keys), **All / Non-empty / In Combined**, **Fit row heights**, and **Delete** (deletes the whole active part **tab**, all pages — not only the current page); Ctrl/Cmd+F focuses Find. The old “Part name (cols×rows)” caption above the grid is gone (tab title is enough).
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **Excel-style cell names** — columns labeled **A, B, C…**, rows **1, 2, 3…**; empty-cell placeholders and status/aria chrome use addresses like **A1** / **B2**
- **+ Row / + Column** — in the **Tools** tab; grow the active part’s grid; new cells start empty
- **Row ↑ / ↓ move** — arrows beside each row shift that cell row up or down (quick click = one step; **hold + drag** moves the row anywhere with a drop line + faint ghost, edge auto-scroll, Esc cancels, one undo step — same remap as ↑ for locks/nests/shades/heights/Combined, including Master→part link remap); **right-click the row number** → **Clear this row** empties that row on the current page (cells + nests, unlinks Master on part tabs; keeps shade + row height; Master tab empties with normal sync), or **Delete this row** removes it and moves every row below up one (cells, nests, locks, shades, heights, Combined checks follow; an empty row is added at the bottom so the row count stays; deleted row’s checks drop; on Master, part cells linked to the deleted row keep their text and unlock, links below remap up; one undo step); ↓ pushes every row below further down (grid grows) so nothing is overwritten; ↑ reorders without clobbering; nested cells, Combined link indices, and row heights stay aligned (column widths unchanged)
- **Sheet paste** — paste TSV / HTML tables from Google Sheets or Excel into the focused cell; values spread across the grid and the grid grows if needed
- **Paragraph paste (v0.158)** — plain text with line breaks (no tabs) pastes into the one cell as-is, no overwrite prompt. Tab-separated text, spreadsheet tables, and blocks copied inside Click2Copy (even a single column, or cells with line breaks) still paste across cells.
- **Overwrite confirms (v0.157)** — one small popover (the paste-confirm look: message + action button + Cancel; Cancel is focused, Esc / click outside cancel) asks before anything replaces or removes text you typed, and only when at least one non-empty cell would actually lose or change text: Delete/Backspace on a lit block, right-click **Clear**, **Clear this row**, **Delete this row**, Col A / Col B ▾ picks and Master parts inserts over your own text (or over a cell’s own nests/pages), **Replace all** (Shift+Enter), deleting a part / cell / nest page or a whole nest that has text, and **Overwrite Master** when other unlinked cells hold the same old text (“Also update N other cells with the same text?”). On the Master tab, clear / replace confirms also say how many linked part cells will update. Swapping one Master value for another never asks (▾ pick, insert, paste). No confirm for Cut, single Replace, typing, single-cell Ctrl+V, reorders, adds, or undo/redo. **Replace / Replace all skip locked Master cells** (status says how many). A Master parts insert with no selected cell and no empty cell adds a row at the bottom (like ⤓) instead of replacing A1. Pasting **Master values** into locked Master cells (single cell or block) replaces and relinks them; any non-Master text into a locked cell is still refused
- **Click2Copy** — a short click or focus leaves the text caret where you clicked (**no** auto-copy on activate). **Triple-click** selects and **copies the entire cell** (or nest) value — including newlines / paragraph breaks — with the same status toast as other copies (overrides the browser’s paragraph-bounded native selection). **Enter** moves to the next cell without copying; **Shift+Enter** inserts a newline. Combined click/focus still copies Combined (see Combined clipboard). Click+drag **inside** one cell selects characters with the normal I-beam. A drag only becomes a multi-cell block once the pointer is clearly inside another cell (~20px past the edge); drifting back restores the text selection. The block stays sticky afterward (clear with Esc, a press outside the grid, a normal click, edit, or Tab/arrows; copy sticky TSV via **Ctrl/Cmd+C**). Plain drag never moves/swaps cells. **Shift+Arrow** / **Shift+click** extend the sticky highlight. A block and a text selection never show together (in every drag direction, including bottom-up; no system blue under the teal block during or after the drag). **Ctrl/Cmd+V** into a lit block pastes from the block’s **top-left** cell, whichever way it was dragged. Typing does not re-copy on every keystroke
- **Cell copy/cut** — with no text selected in a cell, Ctrl/Cmd+C or X copies or cuts the whole cell
- **Col A / Col B ▾ search** — shared Master-value picker search filters **live on every keystroke** (case-insensitive substring; IME/paste too); first match stays highlighted for Enter; Up/Down navigate without stealing the search caret; Esc closes; menu height stays pinned (scroll inside; “No matches” in place — no bounce)
- **Sticky Master insert** — Master parts picker sits above the grid; its collapse header stays put. On part tabs, while the picker is open, **Values** and **A–Z / Z–A** sit in a fixed toolbar above a separate scrollable results list (no sticky jank / jump). Scrolling the results list still closes the Values dropdown; scrolling inside the Values dropdown itself does not.
- **Sticky Column A sort/filter** — grid header row (A–Z / Z–A / Values) stays sticky at the top of the cell grid while scrolling, so those controls remain reachable
- **Master insert greens** — Master library cells get the soft sage/teal Combined outline only when that Master text is an **exact confirmed Combined segment** in the **current part tab’s Combined** (direct Master link or a live confirmed part/nest segment with the same text)—not other tabs’ Combined, not other parts’ segments in Global Combined, and not nest-parent / stale / substring false matches; greens refresh on tab switch and Combined edits
- **Locked Master Combined + part cells** — Master-origin Combined segments are **locked read-only by default** with a **subtle green cue** (soft fill / thin border — not a permanent focus/selection ring); caret typing/deletion/paste into those spans is blocked until unlock. **Insert-from-Master** into a part-tab cell also locks that **grid cell** the same way (`cellLocks`: locked + masterOrigin + masterCellIndex) with a **thin left green accent** (still distinct, without looking always-selected), **copies the Master cell’s nested pages** into the target (replacing any prior nests), and **auto-checks that cell's Combined checkbox** so it joins the Combined prompt (no-op if already checked). Locked inserted cells keep **nests in sync** when Master nest pages change (and on Overwrite). **Double-click** unlocks Combined or the cell/**nests** for editing (amber edit chrome). On finish (blur / Enter): **Overwrite Master** (updates the Master library cell’s parent text **and nests**, syncs other tabs’ locked cells still using that Master index, then re-locks; hidden when the edited text is empty so a clear cannot blank Master), **Keep as local only** (drops Master lock/origin on the Combined segment and/or part cell), or **Cancel** (discards the edit — including nest edits — and re-locks). Esc cancels. Right-click a cell (or a sticky multi-cell block) → **Copy** / **Cut** / shade / **Clear**. **Copy** matches Ctrl/Cmd+C (block TSV if the click is inside a sticky multi-cell selection, else selected text, else the whole cell). **Cut** copies then clears with the same Clear rules (empty + unlink Master on part tabs; Master tab empties with normal sync; one undo step). **Clear** empties the current page and unlinks Master on part tabs without touching the Master library (on the Master tab, Clear empties like a normal edit and existing sync may update linked part cells). Nest textareas keep the browser menu (no Clear there). Locks persist on `confirmedLinks` and per-tab `cellLocks` across reload
- **Master insert Values / A–Z** — same idea as Column A controls: **Values** multi-select filters Master insert rows by first-column text; **A–Z** / **Z–A** reorder the insert list for picking (display only; Master grid data unchanged). The Values menu **stays open** while picking values or using **Clear** / **Select all** (no layout jump; dismiss by clicking elsewhere, Esc, or scrolling the Master insert **results** list — not when scrolling the Values dropdown itself). Selection and sort are **per part tab** — each tab remembers its own prefs, and switching tabs restores that tab’s filter/sort. Undo/redo **prunes** orphaned prefs only (does not wipe other tabs’ remembered Values/sort). Sorting the **Master grid** A–Z remaps other tabs’ `masterCellIndex` locks so Master associations stay intact
- **Row filters** — All / Non-empty / In Combined (same row as Find/Replace) hide rows that do not match (data stays intact; empty and Combined inclusion filters)
- **Fit row heights** — one-click control on the Find/Replace row sizes **each row** to the height of its tallest cell (wrapped parent text + nested cells/pages, not equalize-all-rows); **shrinks** tall empty/short rows as well as growing clipped ones (replaces prior `rowHeights` / manual resize pins — not grow-only); measures at live wrap width (border-box, nest chrome beside textarea, tallest nest page), pins wrap height so `overflow:hidden` does not clip, persists as `rowHeights` like column widths; after Fit, typing in a cell/nest live-refits that row only
- **Column A value filter** — Values multi-select on the Column A header picks which unique first-column values to show; composes with All / Non-empty / In Combined (data stays intact). Row filter + Values selections are **per tab** (including Master), so editing one tab does not overwrite another’s remembered filter. In the Values list, the **checkbox** still toggles the filter; clicking the **value name** jumps to that Column A’s first row (scrolls it near the top, brief highlight, closes the menu) — if it was filtered out, it is checked on first so it is visible. The Master-insert **Values** menu does the same inside the Master parts results pane
- **Sort by column A** — A–Z / Z–A buttons on the Column A header reorder all rows by first-column text (case-insensitive); cells, Combined checkboxes, and confirmed-link cell indices stay row-aligned (blank column-A values sink to the bottom)
- **Checkbox drag-paint** — click+drag across Combined cell toggles or row Combined toggles to set many at once to the first control’s new value; Combined and row drags never mix
- **Col1 Combined cascade + exclusive filter (Ash)** — on part tabs, checking or unchecking a **Column A** Combined checkbox also (un)checks that row's **Column B** and every other Column A cell with the same trimmed value (exact match, same as Values filter) plus each of those rows' Column B. **Checking** a Col1 option is **exclusive**: other Column A values (≠ selected) and those rows' Column B uncheck first (radio-style category filter). **Column A is filter-only**: its text never appears in the Combined prompt (checkbox state only); Column B (and other non-A cells) still append. Nest Combined checkboxes and Master insert stay single-cell; Master grid unchanged
- **Row / cell Combined toggles** — larger checkbox-like controls on each row and each cell add that content to the Combined prompt **at the Combined caret/selection** (or last caret if focus left Combined; otherwise append at end) and stay checked while included; uncheck removes the linked confirmed segment(s) and clears green. Row toggle covers the whole row; cell toggle is per-cell. Both stay in sync with confirmed links. **Match source order** (next to **Global**) when ON reorders Combined green segments to match cell/grid source order (tab order, then row-major cell index, then parent before nests) after each add/remove/re-add; when OFF keeps caret/append order. The toggle persists with the document
- **Nested cells** — **+** under each cell’s Combined checkbox opens **Add blank** / **Copy current** (same picker as cell-page +): Copy current duplicates the current nest (last focused, else last; all pages + its Combined check) right after it and focuses it — with no nests yet it seeds the new nest with the cell’s text; one undo step; locked part cells need unlock first, and copies made on Master sync to locked part cells (multiple nests per parent). Each nest has its **own Combined checkbox** (independent of the parent) and **pages** — nest chrome is **one horizontal row** (Combined checkbox, ◀, `1/n`, ▶, **+** page, ×; wider indent preferred over tall stacked rows); arrow keys at the caret edge also flip pages (▶ past the end adds a page). **Enter** confirms and moves to the next cell (like parents); **Shift+Enter** inserts a newline. Nested content persists with the document. Nest **+** is a quiet Soft dark control (small muted +, not a Combined checkbox tile), stays pinned under the Combined checkbox (not at the bottom of tall autofitted rows), and always attaches to the clicked parent cell; adding/removing a nest refits that row’s height.
- **Undo / Redo** — Ctrl/Cmd+Z undoes document edits (cells, Combined prompt, confirmed links, paste, append, Master insert, clear, etc.); Ctrl/Cmd+Shift+Z or Ctrl+Y redoes. Typing in a cell or Combined is coalesced into one undo step
- **Combined auto-copy** — clicking or focusing the Combined prompt copies its current text to the clipboard (edit still works). Edits, tab switches, and cell Enter no longer auto-copy (see **Combined clipboard**). The Copy button remains as a manual fallback. Hold Ctrl/Cmd while activating Combined to skip the activate copy (clipboard stays for paste overwrite)
- **Confirmed append highlighting** — checked / in-Combined cells use **Green alt 2 — soft outline**: soft sage/teal outline ring around the cell (dark interior, no neon green wash) plus Combined checkboxes that match **Global Combined** (**filled teal/green box** + **black checkmark**, same native accent look — not outline + teal check); Nest **+** under each cell is a separate quiet Soft dark control (small muted +, not checkbox-like); Combined segments get the same soft ring. Editing a Combined segment so it no longer matches its linked cell clears the outline on both sides. **Editing a confirmed cell or nest in the table live-replaces that confirmed segment in Combined** (keeps the link); clearing the cell/nest removes it from Combined
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**; Soft dark native scrollbar appears **only when** Combined text overflows the box (`overflow-y: auto`, never `scroll`)
- **Status / under-the-hood log** — click the status line for a full-width-friendly event panel; long log lines wrap so the window does not need stretching
- **Window orientation** — remembers the BrowserWindow position/size (including half-screen snap/dock) plus maximized/fullscreen across restarts via `userData/window-state.json`; flushes synchronously before quit and **Update ready — Restart** (`app.relaunch` + `app.exit`); restores bounds + maximized/fullscreen on every start (including relaunch) and skips off-screen coordinates if the display layout changed
- **Persistence / autosave** — edits debounce (~250ms) and autosave open documents. Session state lives in Electron `userData` (`click2copy-session.json`). Untitled docs also mirror readable `latest-autosave.c2copy` + `latest-autosave.json` into `userData/backups/` **and** `Documents/Click2Copy/` so you can open or recover them if the app breaks. After **Save As**, autosave writes to that named `.c2copy` path (not a surprise alternate file).
- **Combined clipboard** — auto-copy only on **click/focus Combined**. No auto-copy on Combined text changes, tab switch, Tools part change, Enter in parent/nest cells, or multi-cell drag release. **Triple-click** in a cell/nest textarea also copies the **whole cell** value including paragraph breaks (does not re-add removed cell activate/Enter auto-copies). Explicit **Copy** button and **Ctrl/Cmd+C** remain (sticky multi-cell ranges still copy TSV via Ctrl/Cmd+C).
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

On Windows, double-click **`Click2Copy.vbs`** (repo root) to start the app with **no console window** — you can pin a shortcut to it. It runs `node_modules\electron\dist\electron.exe` for this folder (passing a `.c2copy` path through if given); if Electron isn't installed yet it falls back to **Start Click2Copy.bat** (visible, so setup runs). Update-ready Restart works the same either way.

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
  "cellLocks": [null, { "masterOrigin": true, "locked": true, "masterCellIndex": 0 }],
  "cellPages": [{ "pages": ["…", "…"], "page": 0 }],
  "cellShades": [null, "green"],
  "pages": [{
    "cols": 3,
    "rows": 8,
    "cells": ["…"],
    "nestedCells": [],
    "cellLocks": [],
    "cellPages": [],
    "cellShades": [],
    "partPrompt": "",
    "confirmed": [{ "cellIndex": 0 }]
  }],
  "page": 0
}
```

`cells` is a flat row-major array of length `cols * rows` (working copy of the active part page). Older documents may still carry a `sleptCells` array; it is ignored on load. Optional `icon` is one of `text` | `video` | `img` | `sparkle` | `tag` | `folder` | `layers` | `hash` (defaults: Master → `layers`, parts → `text`; Tools is UI-only wrench and is not stored). `nestedCells` is a parallel array of nest lists; each nest is `{ pages: string[], page: number }` (`page` is the current page index). Optional `cellPages` is a parallel array of parent-cell page entries `{ pages: string[], page }` (same shape as nests). Optional `cellShades` is a parallel array of `null` | `green` | `yellow` | `red`. Optional `cellLocks` is a parallel array (same length as `cells`) of `null` or `{ masterOrigin: true, locked: boolean, masterCellIndex?: number }` for Master-inserted part cells. Part tabs also store `pages` (array of full grid snapshots) and `page` (active index); adding a part page copies the current snapshot with Master locks cleared. Each part page may include `partPrompt` and `confirmed` (`[{ cellIndex, nestIndex? }]`) so switching pages swaps that part’s Combined membership. When a cell is checked, Combined gets **parent text, then each nest’s pages in order** (all non-empty pages), joined by the row separator. Optional `columnWidths` / `rowHeights` arrays (lengths `cols` / `rows`) store resized column widths and auto-fitted row heights in pixels (autofit includes nested height).

Confirmed append links (ranges into the combined prompt tied to source tab/cell, with optional `nestIndex` for nested cells) are stored as `confirmedLinks` on the document. Master-origin segments also store `masterOrigin`, `locked`, and optional `masterCellIndex` so Combined Master locks restore on launch. Part-tab Master locks round-trip on `cellLocks`. They round-trip through session JSON, untitled autosave, named `.c2copy` files, and recoverable backups so cell greens and Master-insert greens restore on launch (offsets are repaired if Combined text still matches).

## Usage tips

1. Edit cells in the active tab’s grid.
2. Use **All** / **Non-empty** / **In Combined** to filter which rows are visible. Use **Values** on the Column A header to multi-select which Column A values to show (composes with the row filters). Use **A–Z** / **Z–A** on the Column A header to sort rows by the first column. The Master parts picker has its own **Values** / **A–Z** / **Z–A** (display-only) for finding insertable Master text; its greens follow exact confirmed segments in the active tab’s Combined (not other tabs / Global cross-part).
3. Check the box left of a row (or the Combined box on a cell) to include it in the Combined prompt at the Combined caret (place the caret first, or it appends at the end); uncheck to remove it. Turn on **Match source order** beside **Global** if you want Combined segments to stay in grid order as you check/uncheck. Use **+** under the Combined checkbox to add indented nested cells (pages via ◀ ▶).
4. Open the shared **Tools** tab for **+ Row** / **+ Column** / **Rename** and separator settings (they still act on the active part tab). Use the part **Find** bar to search or replace text in cells (and nests).
5. Click a cell to place the caret where you clicked (no copy yet). Press **Enter** to confirm, copy that cell, and move to the next (**Shift+Enter** = newline, no copy). Tab/arrows move without copying; **Shift+Arrow** extends a sticky multi-cell highlight. Click+drag inside one cell is native text selection; click+drag clearly into another cell selects a sticky rectangle (Ctrl/Cmd+C copies its TSV). Plain drag never moves/swaps cells. Paste from Google Sheets or Excel into a focused cell — tab-separated or HTML table data fills the grid (expanding rows/columns as needed) instead of dumping into one cell. Copy/cut with no text selection copies the whole focused cell; Ctrl/Cmd+C while a multi-cell highlight is sticky copies its TSV range. Combined prompt still copies on click/focus/edit. Master-inserted part cells and Master-origin Combined segments stay locked (subtle green cue) until **double-click** unlock; finish with Overwrite Master / Keep local / Cancel.
6. Click **Copy** when ready to paste elsewhere.
7. Double-click a tab title to rename it (or use Tools → Rename); use **Delete** (with confirm if non-empty) to remove a tab. Right-click the tab to pick an icon (Master + parts); left-click selects the tab only. Tools stays wrench-only. Part tabs have pages (**◀ n/n ▶** + page chips in the part toolbar): right-click a chip → **Add page after this** inserts a blank page right after it, **Duplicate this page** inserts a copy (cells, nests, cell pages, shades, Combined checks, Master locks) so Master-linked cells stay locked and sync; both switch to the new page and Ctrl+Z undoes them. Renaming a tab (double-click) keeps the tab at its width — long names scroll inside the box, the tab row never re-wraps.
8. Parent cells also have pages (corner **◀ n/n ▶ +**, same idea as nest pages).

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
