# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Reorder tabs** — drag part tabs or prompt-file tabs left or right to rearrange them
- **Master part** — the fixed first part is a reusable text library; edit it there, then click its read-only items in another part to copy them into that part’s next empty cell
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **+ Row / + Column** — grow the grid; new cells start empty
- **Move rows** — use the small up/down arrows beside a row’s append control to reorder rows without losing content; moving down shifts occupied rows until an empty row, adding one if needed
- **Row append** — click the checkbox-like control to the left of any row to append that row’s non-empty cells (joined with ` | `) as a new line in the combined prompt; every click appends again
- **Append all non-empty** — appends every non-empty row of the active part (row-major)
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Global prompt toggle** — keep one combined prompt across all parts, or uncheck **Global** to keep a separate appended prompt for each part; checking it merges those part prompts into the global prompt
- **Resizable prompt pane** — drag the divider above the combined prompt to change its height; the size is remembered
- **Persistence** — tabs, cell grids, and combined prompt are saved to a local JSON file under Electron `userData` and restored on restart
- **Prompt documents** — create or open multiple `.c2copy` files in one window; each file keeps its own tabs, grid, separators, and combined prompt and autosaves independently
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

On Windows, double-click `Start Click2Copy.bat`. It installs the dependencies first if Electron is not installed yet and registers `.c2copy` files for opening with Click2Copy. After that, double-click a `.c2copy` file to open it in the running app.

Alternatively, run:

```bash
npm start
```

This launches Electron with the Click2Copy window. Use **File → New Prompt**, **Open Prompt Files…**, **Save**, and **Save As…** to manage prompt documents. The current app version appears in the window title. Saved `.c2copy` files contain versioned JSON, so they remain independent of the app install and can be opened by later app versions.

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

## Usage tips

1. Edit cells in the active tab’s grid.
2. Click the box left of a row to append that row to the combined prompt (repeatable).
3. Or click **Append all non-empty** for the whole part.
4. Use **+ Row** / **+ Column** to expand the grid.
5. Click **Copy** when ready to paste elsewhere.
6. Double-click a tab title to rename it; use **Delete** (with confirm if non-empty) to remove a tab.

Data is autosaved shortly after edits.

## Version

App version is tracked in `package.json` as thousandths (`0.001`, `0.002`, …) and bumped on each release push.
