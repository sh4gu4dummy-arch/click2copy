# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Cell grid** — each tab is a grid of independently editable cells (default 3 columns × 8 rows)
- **+ Row / + Column** — grow the grid; new cells start empty
- **Row append** — click the checkbox-like control to the left of any row to append that row’s non-empty cells (joined with ` | `) as a new line in the combined prompt; every click appends again
- **Append all non-empty** — appends every non-empty row of the active part (row-major)
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Persistence** — tabs, cell grids, and combined prompt are saved to a local JSON file under Electron `userData` and restored on restart
- **Export / Import** — backup and restore all tabs + combined prompt as JSON (survives app updates; userData is outside the repo). Legacy backups with a single `content` string per tab are migrated into cell `[0]` of a 3×8 grid

## Requirements

- Node.js 18+ recommended
- npm

## Install

```bash
cd click2copy
npm install
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
