/**
 * Ensure Electron binary exists. Never hard-fail npm install — report and exit 0
 * so setup can continue with a repair step.
 */
const fs = require('fs');
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
    return !!rel && fs.existsSync(path.join(electronDir, rel));
  } catch {
    return false;
  }
}

if (!fs.existsSync(electronDir)) {
  console.warn('electron package not present yet');
  process.exit(0);
}

if (hasBinary()) process.exit(0);

console.log('Electron binary missing — running electron install.js …');
if (fs.existsSync(installJs)) {
  const r = spawnSync(process.execPath, [installJs], {
    cwd: electronDir,
    stdio: 'inherit',
    env: { ...process.env, ELECTRON_GET_USE_PROXY: process.env.ELECTRON_GET_USE_PROXY || '0' }
  });
  if (r.status === 0 && hasBinary()) {
    console.log('Electron binary OK.');
    process.exit(0);
  }
}

console.warn('');
console.warn('Electron binary still missing (install scripts may be blocked or download failed).');
console.warn('Run these in this folder, then npm start:');
console.warn('  npm config set ignore-scripts false');
console.warn('  npm approve-scripts electron');
console.warn('  set ELECTRON_GET_USE_PROXY=0');
console.warn('  node node_modules\\electron\\install.js');
process.exit(0);
