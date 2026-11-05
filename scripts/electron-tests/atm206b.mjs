/**
 * Example: Add to Master from Col A also brings companion Col B, and undo clears both.
 * Run: cd scripts/electron-tests && xvfb-run -a node atm206b.mjs
 */
import { launch, sleep } from './auh.mjs';

const H = await launch('atm206b', {
  tabs: [
    { id: 'master', title: 'Master', cols: 2, rows: 2, cells: ['alpha', 'p1', '', ''] },
    { id: 'p1', title: 'Part 1', cols: 2, rows: 2, cells: ['newtype', 'side prompt', '', ''] },
  ],
  activeTabId: 'p1',
});
const { p } = H;
let fails = 0;
const ok = (c, l, x) => {
  console.log(c ? 'PASS' : 'FAIL', l, x === undefined ? '' : JSON.stringify(x));
  if (!c) fails++;
};

await p.click(H.cell(0), { button: 'right' });
await sleep(200);
await p.click('.part-page-menu button:has-text("Add to Master")');
await sleep(500);
ok(/Added to Master A2/i.test(await H.status()), 'added', await H.status());
await H.tab('Master');
const m = await H.cells();
ok(m[2] === 'newtype' && m[3] === 'side prompt', 'Master row got Col A + companion Col B', m);
await H.tab('Part 1');
await p.keyboard.press('Control+z');
await sleep(350);
await H.tab('Master');
ok((await H.cells())[2] === '' && (await H.cells())[3] === '', 'undo cleared both', await H.cells());
console.log('fails', fails);
await H.close();
process.exit(fails ? 1 : 0);
