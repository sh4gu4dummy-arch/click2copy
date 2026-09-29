# Click2Copy

A simple Electron desktop app for composing prompts from reusable parts.

## Features

- **Editable tabs** — add, rename (double-click or Rename button), and delete tabs
- **Part editor** — each tab has a large textarea for that part’s content
- **Append to prompt** — clicks append the active part into the combined prompt (does not replace)
- **Combined prompt** — editable pane with **Copy** to clipboard and **Clear**
- **Persistence** — tabs, contents, and combined prompt are saved to a local JSON file under Electron `userData` and restored on restart
- **Export / Import** — backup and restore all tabs + combined prompt as JSON (survives app updates; userData is outside the repo)

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
| `renderer.js` | Tabs, append, copy, and autosave logic |

## Usage tips

1. Edit text in the active tab’s textarea.
2. Click **Append to prompt** to add that text to the combined prompt.
3. Switch tabs to compose other parts the same way.
4. Click **Copy** when ready to paste elsewhere.
5. Double-click a tab title to rename it; use **Delete** (with confirm if non-empty) to remove a tab.

Data is autosaved shortly after edits.

## Version

App version is tracked in `package.json` as thousandths (`0.001`, `0.002`, …) and bumped on each release push.
