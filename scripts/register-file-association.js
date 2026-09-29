const { spawnSync } = require('child_process');
const path = require('path');

const extensionKey = 'HKCU\\Software\\Classes\\.c2copy';
const programKey = 'HKCU\\Software\\Classes\\Click2Copy.Prompt';
const openCommandKey = `${programKey}\\shell\\open\\command`;
const launcher = path.resolve(__dirname, '..', 'Start Click2Copy.bat');
const openCommand = `"${launcher}" "%1"`;

function runReg(args) {
  return spawnSync('reg.exe', args, { encoding: 'utf8', windowsHide: true });
}

const existing = runReg(['query', extensionKey, '/ve']);
if (existing.status === 0) {
  const match = existing.stdout.match(/REG_SZ\s+(.+)\s*$/m);
  if (match && match[1].trim() !== 'Click2Copy.Prompt') {
    console.warn('.c2copy is already associated with another application; leaving it unchanged.');
    process.exit(0);
  }
  const currentCommand = runReg(['query', openCommandKey, '/ve']);
  const commandMatch = currentCommand.stdout && currentCommand.stdout.match(/REG_SZ\s+(.+)\s*$/m);
  if (match && commandMatch && commandMatch[1].trim() === openCommand) process.exit(0);
}

for (const [key, value] of [
  [programKey, 'Click2Copy Prompt'],
  [openCommandKey, openCommand],
  [extensionKey, 'Click2Copy.Prompt']
]) {
  const result = runReg(['add', key, '/ve', '/t', 'REG_SZ', '/d', value, '/f']);
  if (result.status !== 0) {
    console.error((result.stderr || result.stdout || 'Could not register Click2Copy prompt files.').trim());
    process.exit(result.status || 1);
  }
}
