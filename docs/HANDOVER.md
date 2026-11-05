# Click2Copy — Handover

Written 2026-11-05 by the Click2Copy project bot (Grok Bot, "Click2Copy" agent) for a replacement agent.
It follows Meta-AI `handover-guide-standard.md`. If you are new, read this whole file first. It holds everything
that is not in the code or in `README.md`, especially Ash's rules and the memory that built up over ~200 releases.

---

## 1. Who / what

- **Project:** Click2Copy, a desktop prompt builder for Ash. Electron app with a spreadsheet-like grid of reusable
  prompt parts. Checked cells build a **Combined** prompt that Ash copies into other AI tools.
- **Version:** **v0.206** (`package.json` `"version": "0.206"`), commit `425c7c5`, tag `v0.206`, shipped 2026-11-05.
- **Repo:** https://github.com/sh4gu4dummy-arch/click2copy, branch `main` (the only branch; push straight to main).
- **Runtime:** Electron **33** (box has 33.4.11). Plain JS, no build step, no framework, no bundler.
- **Files:**

  | File | Role |
  |------|------|
  | `renderer.js` (~17.7k lines) | Almost all logic: tabs, grid, pages, nests, Master links/locks, Combined, undo, filters, menus, autosave |
  | `main.js` (~1.7k lines) | Main process: window-state, session/autosave/backups, `.c2copy` open/save, file watching, update-ready relaunch, IPC |
  | `preload.js` | `window.click2copy` IPC bridge (contextIsolation) |
  | `styles.css` | Soft dark charcoal theme (teal accents, amber Tools tab / amber edit outline) |
  | `index.html` | UI skeleton |
  | `scripts/ensure-electron.js`, `scripts/register-file-association.js` | Setup helpers (Electron binary repair; `.c2copy` association) |
  | `scripts/electron-tests/` | Real-Electron Playwright harness (`auh.mjs` + example tests; see that folder's README) |
  | `Click2Copy.vbs` | Windows launcher, no console window (Ash's normal way to start it) |
  | `Start Click2Copy.bat`, `Initial Setup.bat` | Visible-console launcher / first-time setup fallback |

- **Documents:** `.c2copy` prompt files (JSON). Each document has a **Master** tab (the library of approved values) and
  **part tabs** (where Ash builds prompts, linking to Master values). There is also a UI-only **Tools** tab pinned far right.
- **How Ash runs it:** locally on **Windows** from **`E:\AI\click2copy`**. He updates with `git pull` (his own PullSync /
  GitPullManager tooling) and the app shows **Update ready — Restart**. The app never goes online. The only
  network use is first-time setup downloading Electron.
- **Role of this bot:** the only developer. Ash reports bugs or asks for features in chat. The bot reproduces them, fixes,
  tests in real Electron, bumps the version, and pushes.

---

## 2. How to talk to Ash

- **Short, plain replies. No chat spam.** Message only when something is done or Ash must act. No "starting now…" chains.
- **"lmk" / "i'm just asking" = explain only, do not build.** Answer the question and change nothing.
- **When Ash asks for an explanation, go slowly and use concrete examples** (e.g. "Master row 3 says X; on Part 1 …").
- **"f" = fold:** make the change locally only, no commit/push, until Ash confirms. After each fold give **one** next idea.
- After a ship, tell Ash the version, what changed in 1–3 short steps, and that it was tested in the real app.
- If screenshots are sent, use placeholder data only, and delete the image file from the box afterwards.
- Ash's version labels look like `v0.206` (thousandths).
- QAsupervisor (another of Ash's agents) coordinates handovers/audits. When QAsup says "stay quiet to Ash", reply only to QAsup.

---

## 3. Standing Ash rules (still active)

**Shipping / git**
- **2026-10-18, Ash:** "stop asking for me everything just push" … "I accept the risks". Pick the sensible option,
  fix, and ship without asking. **Still nothing destructive without asking:** no force-push, no branch delete,
  no history rewrite / filter-repo.
- Ship recipe: test in **real Electron 33 under xvfb** → bump `package.json` version in thousandths (`"0.207"` style;
  **leave `package-lock.json` alone**) → commit `vX.XXX: summary` → tag `vX.XXX` → push main + tag → verify with
  `git ls-remote`. (See §7 for the exact commands.)
- Ash does not want local setup/npm to leave tracked changes or a dirty/ahead repo on his PC. Never commit changes
  that `npm install` would make on his side (package-lock churn, etc.).
- **Only one worker on Click2Copy at a time.** Never run two executors/agents on this repo in parallel.

**Scope**
- **Do exactly what Ash asks. No unrequested extras.** He was angry when a launcher request grew desktop/Start-menu
  shortcut creation. Offer extras as ideas and ask first ("just ask me first next time maybe I'll like your idea").
- Bugs found while testing something else: report and ask before fixing.
- Find the **root cause first**. Make a minimal, targeted fix.
- Shared rule across Ash's bots: quality/creativity beats rigid rules, but flag the conflict instead of silently
  breaking a rule.

**UI**
- **No bouncing UI:** nothing that shows/hides and shifts widths or rows when state changes. Fixed-width counters, etc.
- **Prefer right-click menus over buttons.** No redundant controls. Use free space on existing rows before adding rows.
- Jumps/scrolls land the target **near the top with 1–2 rows above it**.
- Persist UI session state across restarts and updates (scroll positions, open tab/page, zoom, window position).

**Data safety**
- **Never nuke user data.** User data lives outside the app folder so pulls/restarts keep everything.
- **Confirm before overwriting Ash's own typed text.** Exceptions: swapping one Master-linked value for another Master
  value (▾ pick, insert, paste), and right-click **Clear** (Ctrl+Z undoes).
- **Clear/delete/cut of a Master cell that has linked part cells always confirms and lists the linked locations.**
- Cut of a locked linked part cell = Clear (copies, unlinks, no confirm, Ctrl+Z undoes).

**Privacy / hygiene (Ash, verbatim)**
- "also don't store info about my prompts like the text etc": never store Ash's prompt/cell text in memory,
  docs, tests, fixtures or commit messages. Tests use generic placeholders (`alpha`, `prompt one`, …).
- "clean up any temp screenshots etc on your machine if no longer needed".
- Ash once pasted a GitHub PAT in chat. It was not used or stored, and the bot suggested revoking it. **Never store PATs in memory or files.**

---

## 4. Memory dump: product behaviour and history a stranger needs

### Core model
- **Master tab** = the library. **Part tabs** = prompts built from Master values. Column A (Col A) is the
  **category/filter column**, Column B (Col B) the prompt text.
- **Col A is filter-only.** Its text never goes into Combined. Clicking a Col A cell filters rows to that value
  (click again / Esc clears). Values are picked with the ▾ menu (search filters live) or typed (type-ahead suggestions).
- **Validation (part tabs):** Col A must be a Master Col A value. Col B (when Col A holds a Master value) must be one of
  that Col A's Master Col B values (strict). Bad values are **kept and marked ⚠/amber, never deleted**. Header ⚠ N badge
  click jumps to the next not-allowed cell (v0.200).
- **Master links:** inserting/picking a Master value into a part cell creates a real link in `cellLocks[i]`
  = `{ masterOrigin: true, locked: true, masterCellIndex }`. Linked cells are read-only (thin green left accent),
  sync text **and nests/pages** when Master changes. Typed/pasted text that exactly matches a Master cell auto-links on
  Enter/blur. Ash rejected a fake look-only lock: links must be real and sync.
- **Unlock/edit session:** double-click a locked cell → `masterSegmentEdit` (renderer.js ~2260) shows the **amber
  outline**. Finishing offers **Overwrite Master** (hidden if text empty so Master can't be blanked) / **Keep as local
  only** / **Cancel**; Esc cancels. Any press/focus outside ends the session (v0.203). While editing, the row stays
  visible even if the Col A filter would hide it (v0.204).
- **Linked cells and pages (v0.195/v0.198):** a linked cell keeps its own visible cell/nest page number. When Master
  reorders pages, linked cells get the same 1,2,3 order and stay on their page number, so the text shown can change.
  Ash confirmed this is intended.
- **Pages:** part tabs have pages (◀ n/n ▶ + chips; right-click chip → Rename / Add page after this / Duplicate this
  page / Delete page). Cells have their own pages and nests (indented sub-cells with pages). Right-click page counters →
  Move page back/forward, Delete page / Delete nest. There is no × button; right-click menus replaced them.

### Combined / clipboard
- Combined is generated **only from checked items on the current page** (single source of truth since v0.140).
  The **Notes** box is separate for Ash's own text. **Clear** wipes notes. **Uncheck all** is separate (undoable).
- **Auto-clipboard only on Combined click/focus**, and on **tab switch** Combined is auto-copied. No other auto-copies.
- Master insert goes into the **selected cell** and **auto-checks its Combined box**.

### Keys / mouse
- Enter = next cell, Shift+Enter = newline, Tab = next cell, arrows move cells, Shift+arrow / Shift+click extend a
  sticky multi-cell selection. Drag across cells selects a block (Ctrl+C copies TSV), **never swaps/moves cells**.
  Drag inside one cell = normal text selection (block only after ~20px into another cell).
- Triple-click copies the **full cell** (across paragraph breaks) with a soft copy flash (v0.191).
- Right-click row number / column letter selects that row/column (and opens the block cell menu, v0.199). Drag across
  headers selects ranges (v0.189). Row ↑/↓ arrows and hold-drag move rows.
- Right-click cell menu: Copy / Cut / shade (green/yellow/red/clear) / Clear / notes (Add/Edit/Delete, corner mark +
  hover tooltip, v0.190) / Extract dupes / Find similar / Add to Master / Jump to Master / text size A+ A-.
- Text size A+/A-: 30% per click, 70–220%. App zoom Ctrl+wheel / Ctrl+= / - / 0 (persisted, v0.192).
- **Fit** only changes Col A width (sizes to widest value, one line, capped 40%). Right-click Fit → Auto adjust.

### Master tools
- **Extract dupes** (right-click, Master Col B prompts): strips Master dupes from a cell or whole selected block,
  with a strike-through preview before Apply.
- Master Col B duplicates get an **orange outline**.
- **Find similar** (Master Col B, 75%+ similarity, v0.178).
- **Add to Master** (right-click on part-tab Col A or Col B): fills the first empty Master row (never overwrites),
  links the part cell, Ctrl+Z undoes. Col A support is new in **v0.206**: a new Col A value typed on a part tab now stays
  (amber ⚠) instead of being reverted, so it can be promoted.
- **Jump to Master** (right-click linked part cell, v0.201): opens the origin near top, same page, flash. If the target
  is filtered out, the status says to click All.
- Values list (Col A header and Master-insert panel): checkbox filters, clicking the name jumps to the first row (v0.148),
  and jumps can be pinned as quick buttons on the Find row (v0.161).

### Persistence
- `scheduleSave()` (renderer.js ~1433) refreshes Combined/marks, snapshots the document, and debounces `persist()` by
  **250 ms**. Session lives in Electron `userData/click2copy-session.json`. Untitled docs also mirror
  `latest-autosave.c2copy` + `.json` to `userData/backups/` **and** `Documents/Click2Copy/`. Named docs autosave to their
  own path. `window-state.json` keeps window bounds/maximize. Grid scroll positions are persisted.
- **File-changed bar (v0.205):** if an open named `.c2copy` changes on disk outside the app (or goes missing), autosave
  stops writing it. A bar offers **Reload / Keep mine / Save current as…**. There is also a pre-write outside-change check
  (`documents:check` IPC). The session backup continues meanwhile.
- `main.js normalizeTab()` must keep every tab field (icon, nests, locks, rowHeights, cellNotes, pages …). One bug
  was dropping fields on save. **When you add a new per-tab field, add it to normalizeTab** too.

### Row delete / clear
- `deleteRowFromMenu` (~5084): deletes a row and shifts everything up (cells, nests, locks, shades, heights, checks).
  On a part tab, Master-linked cells and their matching Col A are **not** counted as "filled" for the confirm (v0.202).
  Deleting a part row with a linked cell does **not** delete the Master cell.
- `clearCellsFromMenu` (~4878): empties + unlinks. On Master, cells with linked part cells always confirm with the
  locations, then clear + unlink those part cells.

### Feature requests checked against code (2026-11-05)
| Request | Status |
|---|---|
| Soft copy flash on copy | **Done** v0.191 (Option B soft accent tint, one pulse). Ash was offered grey vs blue tint and hasn't picked |
| Right-click cell notes with icon (Add/Edit/Delete, hover shows) | **Done** v0.190 |
| Drag across row numbers / column letters to multi-select | **Done** v0.189 (row-number drag starts a range select. Row moving is the ↑/↓ arrows / hold-drag on the arrows) |
| Add page while unlocked-editing a linked cell | **Done** v0.187/v0.188 |
| Values list click scrolls to first row | **Done** v0.148 (+ quick buttons v0.161) |
| Extract dupes on many cells at once | **Done** (v0.177 block + v0.199 column) |

### Pending issues (NOT started; don't start without Ash)
1. **Part cells with their own extra pages/nests never auto-lock to a matching Master cell** ("matches Master cell, not
   linked"). Options discussed: (a) copy the extra pages/nests into the Master cell, then lock; (b) per-page/per-nest locks
   instead of one whole-cell lock (bigger change: locks, sync and undo all assume one lock per cell). Ash only asked
   about (b) as a question. **Undecided.**
2. **"Scan Master for near-dupes"**: a whole-Master list of close pairs (builds on Find similar). Not started.

### Old scars / lessons
- Unrequested extras (shortcut creation) angered Ash. Ask first.
- A fix was once reported before it was pushed and Ash couldn't test it. Only say "pull it" after `git ls-remote` shows it.
- Chromium 130 (Electron 33) ignores `::details-content`. Use explicit px heights for scrolling inside `<details>`.
- CSS `display:block` overrode `[hidden]` in the ▾ menu (v0.196). Check computed styles when hiding things.
- Testing in a stub browser missed real bugs. **Always test in real Electron 33.**
- `v0.205` shipped `package.json` as `"0.205.0"` (wrong format). It was fixed to `"0.206"` without a tag rewrite.

### Open offers awaiting Ash (only if he brings them up)
- Cap Col A suggestions at ~100 with lazy load.
- Copy-flash tint: grey or blue.
- Should a page flip start a new undo step?

---

## 5. Current status (2026-11-05)

- **Idle.** Last ship **v0.206** (`425c7c5`, tag pushed). Working tree clean apart from this docs commit. No unpushed work.
- No folds pending. No branches besides `main`.
- In-repo harness: `scripts/electron-tests/` (launcher + examples). Older one-off scratch still on the box at `/workspace/click2copy-tests/` (not in git; see §7).

## 6. Waiting on Ash

- Nothing actively waiting. Ash may test **v0.205** (file-changed bar) and **v0.206** (Add to Master for Col A) and report back.
- Undecided items: pending issues 1–2 and the open offers in §4.

---

## 7. Desk / preview: how to run and test

- No ports, tunnels or Pages. It is a desktop app. Ash runs it himself. Screenshots are the only "preview" and must use placeholder data.
- **Run locally (box):** `cd /workspace/click2copy && npm start` needs a display. On the box use
  `xvfb-run -a …`. The Electron binary is `node_modules/electron/dist/electron` (33.4.11). Pass `--no-sandbox` as root/box.
- **Test harness (in repo):** `scripts/electron-tests/`
  - `auh.mjs` is the helper: `launch(name, data)` writes a throwaway session into a temp `XDG_CONFIG_HOME`
    (`/tmp/c2c-au-<name>`), launches **real Electron 33** via Playwright `_electron`, and gives helpers
    (`cells()`, `status()`, `cell(i)`, `locked()`, `tab(name)`, clipboard get/set, `close()`).
  - Resolves Playwright from `PLAYWRIGHT_NODE_MODULES`, repo `node_modules/playwright`, or `/tmp/c2c-pw`.
    Resolves Electron from `ELECTRON_PATH`, repo `node_modules/electron/dist/electron`, or `/tmp/e33`.
    `/tmp` may be wiped; reinstall with `npm i playwright@1.48.2` in `/tmp/c2c-pw` and (fallback)
    `npm i electron@33` in `/tmp/e33`, or run `npm run setup` in the repo for the Electron binary.
  - Run one test: `cd scripts/electron-tests && xvfb-run -a node atm206.mjs`. Prints PASS/FAIL lines.
  - Shipped examples: `atm206.mjs` (Add to Master Col A), `atm206b.mjs` (companion Col B + undo). See
    `scripts/electron-tests/README.md`.
  - Box-local scratch (NOT in repo, may delete anytime): `/workspace/click2copy-tests/` still holds ~150
    one-off `*.mjs` scripts named per feature/version plus old `stub*.js` / `main*.cjs` browser stubs
    (less trustworthy). Prefer the in-repo harness for new tests.
  - Fixtures use placeholder values only. Keep it that way.
- **Ship commands:**
  ```bash
  cd /workspace/click2copy
  # edit package.json "version": "0.207"  (thousandths; do not touch package-lock.json)
  git add -A && git commit -m "v0.207: …" && git tag v0.207
  git push -q "https://x-access-token:${GH_WORKFLOW_PAT}@github.com/sh4gu4dummy-arch/click2copy.git" HEAD:main v0.207
  git ls-remote https://github.com/sh4gu4dummy-arch/click2copy.git | grep -E "HEAD|v0.207"
  ```

## 8. Doc map

- `README.md`: full feature list, data shape (tab JSON fields), autosave locations, install/run. Long but accurate up to v0.206.
- `docs/HANDOVER.md`: this file.
- `scripts/electron-tests/README.md`: how to run the real-Electron Playwright harness.
- No START-HERE / AGENTS / qa-howto files exist in this repo.
- Meta-AI repo (`sh4gu4dummy-arch/Meta-AI`): `handover-guide-standard.md`, `project-bot-status.md`, `START-HERE.md`
  for cross-project agent rules.

## 9. Secrets & hygiene

- **GitHub push token:** env var **`GH_WORKFLOW_PAT`** on the shared box. Use it only inside the push URL. Never echo it, write it to files, or put it in memory. If it's missing, ask Ash / QAsupervisor.
- The app itself has no secrets, accounts, API keys or network calls.
- Never commit `.grok/`, screenshots, test fixtures with real prompt text, or Ash's `.c2copy` files.
- Ash's prompt content is private: no copies on the box, in memory, docs or tests.

## 10. Git history cleanup: soft dry-run suggestions only

**Suggestions only. Nothing here has been done. Do not rewrite, force-push or filter unless Ash confirms that specific change.**
- `v0.205` commit/tag has `package.json` `"version": "0.205.0"` (breaks the thousandths format). It's harmless and
  already fixed in v0.206. Recommend **leaving it**. A tag rewrite would break Ash's pulls.
- A few early non-version commits (`Update README.md`, `Fix sticky cell clicks…`, `Drop row-header Combined check`,
  `Merge pull request #1`) predate the `vX.XXX:` convention. They're harmless, so leave them.
- Overall recommendation: **no history cleanup needed.**

---

## Recent changelog (v0.195 – v0.206)

| Version | Change |
|---|---|
| v0.206 | Add to Master from part-tab Col A: new Col A types typed on a part tab stay (amber ⚠) and can be promoted |
| v0.205 | File changed outside the app bar (Reload / Keep mine / Save current as…), missing-file bar, pre-write outside-change check |
| v0.204 | Master-linked cell's row stays visible while being edited (Col A filter no longer ends the edit); Esc works on Overwrite prompt after outside click |
| v0.203 | Master cell edit session ends on any press/focus outside the cell (stale amber outline, lost edits) |
| v0.202 | Part-tab Delete row doesn't count Master-linked cells (and their matching Col A) as filled |
| v0.201 | Right-click Jump to Master on linked part cells |
| v0.200 | Header ⚠ N badge jumps to the next not-allowed cell |
| v0.199 | Right-click column letter selects column + opens block cell menu (Extract dupes etc.) |
| v0.198 | Move nest page back/forward in nest page counter menu (Master syncs linked nests; Ctrl+Z) |
| v0.197 | Find / All tabs / Replace moved right of the page line; page tabs get the space |
| v0.196 | ▾ value menu filter actually hides non-matching values |
| v0.195 | Linked cells keep their own visible cell/nest page; unlock + Overwrite work from any page |

Full history: `git log --oneline` (every release commit is `vX.XXX: summary` and tagged).
