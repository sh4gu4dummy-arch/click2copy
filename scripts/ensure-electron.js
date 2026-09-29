/**
 * Ensure Electron's binary was installed (npm may block postinstall via allow-scripts).
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
  console.error('electron package missing — run: npm install');
  process.exit(1);
}

if (hasBinary()) {
  process.exit(0);
}

console.log('Electron binary missing — running electron install.js …');
if (!fs.existsSync(installJs)) {
  console.error('electron/install.js not found');
  process.exit(1);
}
const r = spawnSync(process.execPath, [installJs], {
  cwd: electronDir,
  stdio: 'inherit',
  env: process.env
});
if (r.status !== 0 || !hasBinary()) {
  console.error('');
  console.error('Electron still failed to install.');
  console.error('On newer npm, allow scripts then reinstall:');
  console.error('  npm approve-scripts electron');
  console.error('  rmdir /s /q node_modules\\electron');
  console.error('  npm install electron --save-dev');
  process.exit(1);
}
console.log('Electron binary OK.');
