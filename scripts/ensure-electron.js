/**
 * Ensure Electron binary exists. Never hard-fail npm install — report and exit 0
 * so setup can continue with a repair step.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const root = path.join(__dirname, '..');
const electronDir = path.join(root, 'node_modules', 'electron');
const installJs = path.join(electronDir, 'install.js');
const pathTxt = path.join(electronDir, 'path.txt');

function hasBinary() {
  if (!fs.existsSync(pathTxt)) return false;
  try {
    const rel = fs.readFileSync(pathTxt, 'utf8').trim();
    const distDir = process.env.ELECTRON_OVERRIDE_DIST_PATH || path.join(electronDir, 'dist');
    return !!rel && fs.existsSync(path.join(distDir, rel));
  } catch {
    return false;
  }
}

function cachedArchive() {
  if (process.platform !== 'win32') return null;

  const arch = process.env.npm_config_arch || process.arch;
  const electronVersion = require(path.join(electronDir, 'package.json')).version;
  const filename = `electron-v${electronVersion}-win32-${arch}.zip`;
  const cacheRoot = process.env.electron_config_cache ||
    path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'electron', 'Cache');
  if (!fs.existsSync(cacheRoot)) return null;

  const directories = [cacheRoot];
  for (const entry of fs.readdirSync(cacheRoot, { withFileTypes: true })) {
    if (entry.isDirectory()) directories.push(path.join(cacheRoot, entry.name));
  }

  for (const directory of directories) {
    const archive = path.join(directory, filename);
    if (fs.existsSync(archive)) return archive;
  }
  return null;
}

function extractCachedArchive(archive) {
  if (!archive) return false;

  const distDir = process.env.ELECTRON_OVERRIDE_DIST_PATH || path.join(electronDir, 'dist');
  const quote = value => `'${value.replace(/'/g, "''")}'`;
  const command = [
    '$ErrorActionPreference = "Stop"',
    `Expand-Archive -LiteralPath ${quote(archive)} -DestinationPath ${quote(distDir)} -Force`,
    `if (-not (Test-Path -LiteralPath ${quote(path.join(distDir, 'electron.exe'))})) { exit 1 }`
  ].join('; ');
  const result = spawnSync('powershell.exe', [
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', command
  ], { stdio: 'inherit', windowsHide: true, timeout: 120000 });

  if (result.error) {
    console.warn(`Could not extract cached Electron archive: ${result.error.message}`);
    return false;
  }
  if (result.status !== 0) {
    console.warn(`Cached Electron archive extraction failed (exit code ${result.status}).`);
    return false;
  }

  fs.writeFileSync(pathTxt, 'electron.exe');
  return hasBinary();
}

function runInstaller(mirror) {
  const env = { ...process.env };
  if (mirror) env.ELECTRON_MIRROR = mirror;

  const result = spawnSync(process.execPath, [installJs], {
    cwd: electronDir,
    stdio: 'inherit',
    env,
    timeout: 120000
  });

  if (result.error) {
    console.warn(`Electron installer did not complete: ${result.error.message}`);
  } else if (result.status !== 0) {
    console.warn(`Electron installer failed (exit code ${result.status}).`);
  }
}

function fail() {
  console.warn('');
  console.warn('Electron binary is still missing.');
  console.warn('Check your network access to GitHub releases or the Electron mirror, then run:');
  console.warn('  npm run setup');
  process.exit(process.argv.includes('--strict') ? 1 : 0);
}

if (!fs.existsSync(electronDir)) {
  console.warn('electron package not present yet');
  fail();
}

if (hasBinary()) process.exit(0);

if (extractCachedArchive(cachedArchive())) {
  console.log('Electron binary recovered from the cached release archive.');
  process.exit(0);
}

if (!fs.existsSync(installJs)) {
  console.warn('Electron package is incomplete: install.js is missing.');
  fail();
}

console.log('Downloading Electron binary...');
runInstaller();
if (hasBinary()) {
  console.log('Electron binary OK.');
  process.exit(0);
}

if (extractCachedArchive(cachedArchive())) {
  console.log('Electron binary recovered from the cached release archive.');
  process.exit(0);
}

if (process.platform === 'win32' && !process.env.ELECTRON_MIRROR) {
  console.log('Retrying the Electron download from the npmmirror CDN...');
  runInstaller('https://cdn.npmmirror.com/binaries/electron/');
  if (hasBinary()) {
    console.log('Electron binary OK.');
    process.exit(0);
  }

  if (extractCachedArchive(cachedArchive())) {
    console.log('Electron binary recovered from the cached release archive.');
    process.exit(0);
  }
}

fail();
