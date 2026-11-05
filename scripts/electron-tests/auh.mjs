/**
 * Real-Electron Playwright launcher for Click2Copy agent tests.
 * Writes a throwaway session under /tmp/c2c-au-<name> and launches the app
 * with XDG_CONFIG_HOME/HOME pointed there so Ash's real session is never touched.
 *
 * Deps (not bundled — install once on the box):
 *   Playwright: PLAYWRIGHT_NODE_MODULES, or <repo>/node_modules/playwright, or /tmp/c2c-pw
 *   Electron:   ELECTRON_PATH, or <repo>/node_modules/electron/dist/electron, or /tmp/e33
 *
 * Run under xvfb on the box: xvfb-run -a node example.mjs
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const REPO_ROOT = path.resolve(__dirname, '..', '..');

function firstExisting(candidates) {
  for (const p of candidates) {
    if (p && fs.existsSync(p)) return p;
  }
  return null;
}

async function loadPlaywrightElectron() {
  const candidates = [
    process.env.PLAYWRIGHT_NODE_MODULES,
    path.join(REPO_ROOT, 'node_modules', 'playwright'),
    '/tmp/c2c-pw/node_modules/playwright',
  ].filter(Boolean);
  const pwRoot = firstExisting(candidates);
  if (!pwRoot) {
    throw new Error(
      'Playwright not found. Install with: mkdir -p /tmp/c2c-pw && cd /tmp/c2c-pw && npm i playwright@1.48.2\n' +
        'Or set PLAYWRIGHT_NODE_MODULES to a playwright package dir.'
    );
  }
  const mod = await import(pathToFileURL(path.join(pwRoot, 'index.mjs')).href);
  if (!mod._electron) {
    throw new Error(`playwright at ${pwRoot} has no _electron export`);
  }
  return mod._electron;
}

function resolveElectronBinary() {
  const fromEnv = process.env.ELECTRON_PATH;
  const fromRepo = (() => {
    const electronDir = path.join(REPO_ROOT, 'node_modules', 'electron');
    const pathTxt = path.join(electronDir, 'path.txt');
    if (fs.existsSync(pathTxt)) {
      const rel = fs.readFileSync(pathTxt, 'utf8').trim();
      const candidate = path.join(electronDir, 'dist', rel);
      if (fs.existsSync(candidate)) return candidate;
    }
    const linux = path.join(electronDir, 'dist', 'electron');
    if (fs.existsSync(linux)) return linux;
    return null;
  })();
  const found = firstExisting([
    fromEnv,
    fromRepo,
    '/tmp/e33/node_modules/electron/dist/electron',
  ]);
  if (!found) {
    throw new Error(
      'Electron binary not found. Run npm run setup in the repo, or set ELECTRON_PATH, ' +
        'or install: mkdir -p /tmp/e33 && cd /tmp/e33 && npm i electron@33'
    );
  }
  return found;
}

/** Absolute path of the throwaway session root for a given launch name. */
export function sessionRoot(name) {
  return path.join('/tmp', 'c2c-au-' + name);
}

/**
 * Launch Click2Copy with a throwaway session.
 * @param {string} name short id used in /tmp/c2c-au-<name>
 * @param {object} data document.data (tabs, activeTabId, …) — placeholder text only
 * @param {{ reuse?: boolean }} [opt]
 */
export async function launch(name, data, opt = {}) {
  const _electron = await loadPlaywrightElectron();
  const electronPath = resolveElectronBinary();
  const root = sessionRoot(name);
  const cfg = path.join(root, 'cfg');
  const home = path.join(root, 'home');

  if (!opt.reuse) {
    fs.rmSync(root, { recursive: true, force: true });
    fs.mkdirSync(path.join(cfg, 'click2copy'), { recursive: true });
    fs.mkdirSync(path.join(home, 'Documents'), { recursive: true });
    fs.writeFileSync(
      path.join(cfg, 'click2copy', 'click2copy-session.json'),
      JSON.stringify({
        documents: [{ id: 'd1', title: 'U', filePath: null, data }],
        activeDocumentId: 'd1',
      })
    );
  }

  const app = await _electron.launch({
    executablePath: electronPath,
    args: ['--no-sandbox', REPO_ROOT],
    env: { ...process.env, XDG_CONFIG_HOME: cfg, HOME: home },
  });
  const p = await app.firstWindow();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message));
  await p.waitForSelector('textarea.cell');
  await p.waitForTimeout(600);

  const H = {
    app,
    p,
    errs,
    root,
    cells: () => p.$$eval('textarea.cell', (a) => a.map((t) => t.value)),
    status: () => p.locator('#status').innerText(),
    cell: (i) => 'textarea.cell[data-idx="' + i + '"]',
    focus: () =>
      p.evaluate(() => {
        const a = document.activeElement;
        return a
          ? a.tagName +
              (a.dataset && a.dataset.idx != null ? '#' + a.dataset.idx : '') +
              (a.id ? '#' + a.id : '') +
              (a.className ? '.' + String(a.className).split(' ')[0] : '')
          : 'none';
      }),
    locked: () =>
      p.$$eval('.cell-wrap', (a) =>
        a
          .filter((w) => w.classList.contains('master-cell-locked'))
          .map((w) => w.querySelector('textarea.cell')?.dataset.idx)
      ),
    setClip: (t) => app.evaluate(({ clipboard }, t) => clipboard.writeText(t), t),
    getClip: () => app.evaluate(({ clipboard }) => clipboard.readText()),
    tab: async (t) => {
      await p.click('.tab:has-text("' + t + '")');
      await p.waitForTimeout(350);
    },
    async close(keep) {
      await app.close();
      if (!keep) fs.rmSync(root, { recursive: true, force: true });
    },
  };
  return H;
}

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
