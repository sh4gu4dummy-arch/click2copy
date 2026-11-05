# Electron test harness (agent / box)

Real **Electron 33** UI tests via Playwright `_electron`. Stub browser tests miss real bugs — prefer this.

## Files

| File | Role |
|------|------|
| `auh.mjs` | Shared launcher: throwaway session + helpers (`cells`, `status`, `locked`, `tab`, clipboard, `close`) |
| `atm206.mjs` | Example: v0.206 Add to Master from part-tab Column A |
| `atm206b.mjs` | Example: Col A add brings companion Col B; undo clears both |

## One-time deps on the box

Playwright and a second Electron install are **not** committed. Either:

```bash
mkdir -p /tmp/c2c-pw && cd /tmp/c2c-pw && npm i playwright@1.48.2
# Electron: prefer the repo binary after `npm run setup` in the repo root.
# Fallback: mkdir -p /tmp/e33 && cd /tmp/e33 && npm i electron@33
```

Or set `PLAYWRIGHT_NODE_MODULES` / `ELECTRON_PATH`. `/tmp` may be wiped; reinstall as above.

## Run one test

```bash
cd /workspace/click2copy/scripts/electron-tests
xvfb-run -a node atm206.mjs
```

Prints `PASS` / `FAIL` lines; exit `1` if any fail.

## Rules

- Placeholder fixture text only (`alpha`, `prompt one`, …). Never Ash's real prompts.
- Do not commit screenshots, `.grok/`, or `/tmp` deps.
- Older one-off scripts (~150) may still live under `/workspace/click2copy-tests/` on the box; they are scratch, not the canonical harness.
