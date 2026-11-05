/**
 * Example: v0.206 Add to Master from part-tab Column A.
 * Run: cd scripts/electron-tests && xvfb-run -a node atm206.mjs
 */
import { launch, sleep, sessionRoot } from './auh.mjs';
import fs from 'fs';
import path from 'path';

const mc = ['alpha', 'prompt one', '', '', 'beta', 'prompt two', '', ''];
const pc = ['', '', 'alpha', 'prompt one', '', 'new type text'];
const H = await launch('atm206', {
  tabs: [
    { id: 'master', title: 'Master', cols: 2, rows: 4, cells: mc },
    { id: 'p1', title: 'Part 1', cols: 2, rows: 3, cells: pc },
  ],
  activeTabId: 'p1',
});
const { p } = H;
let fails = 0;
const ok = (c, l, x) => {
  console.log(c ? 'PASS' : 'FAIL', l, x === undefined ? '' : JSON.stringify(x));
  if (!c) fails++;
};

// 1) Type a brand-new Col A value on empty A1 — must KEEP (not revert)
await p.click(H.cell(0));
await sleep(150);
await p.keyboard.type('gamma');
await sleep(100);
await p.keyboard.press('Tab');
await sleep(400);
const a1 = await p.$eval(H.cell(0), (e) => e.value);
ok(a1 === 'gamma', 'typed new Col A kept (not reverted)', a1);
const st1 = await H.status();
ok(
  /Add to Master/i.test(st1) || /not in Master Column A/i.test(st1),
  'status hints Add to Master or not-in-Master',
  st1
);
const marked = await p.$eval('.cell-wrap[data-row="0"][data-col="0"]', (w) =>
  w.classList.contains('col1-invalid')
);
ok(marked, 'new Col A marked not-allowed', marked);

// 2) Right-click Col A → Add to Master enabled
await p.click(H.cell(0), { button: 'right' });
await sleep(250);
const item = await p
  .$eval('.part-page-menu button:has-text("Add to Master")', (b) => ({
    dis: b.disabled,
    title: b.title,
  }))
  .catch(() => ({ dis: true, title: 'absent' }));
ok(!item.dis, 'Add to Master enabled on Col A', item);
await p.click('.part-page-menu button:has-text("Add to Master")');
await sleep(500);
const st2 = await H.status();
ok(/Added to Master/i.test(st2), 'status after add', st2);
const locked = await H.locked();
ok(locked.includes('0') || locked.includes(0), 'Col A cell locked after add', locked);

// 3) Master has gamma
await H.tab('Master');
const mcells = await H.cells();
ok(mcells.includes('gamma'), 'Master contains new Col A type gamma', mcells);

// 4) Undo restores
await H.tab('Part 1');
await p.keyboard.press('Control+z');
await sleep(400);
await H.tab('Master');
const afterUndo = await H.cells();
ok(!afterUndo.includes('gamma'), 'Ctrl+Z removed gamma from Master', afterUndo);
await H.tab('Part 1');
const locked2 = await H.locked();
ok(!locked2.includes('0') && !locked2.includes(0), 'Col A unlocked after undo', locked2);
const a1b = await p.$eval(H.cell(0), (e) => e.value);
ok(a1b === 'gamma', 'part Col A text still gamma after undo', a1b);

// 5) Col B Add to Master still works (regression)
await p.click(H.cell(5), { button: 'right' });
await sleep(200);
await p.click('.part-page-menu button:has-text("Add to Master")');
await sleep(500);
const st3 = await H.status();
ok(/Added to Master|Already in Master/i.test(st3), 'Col B Add to Master still works', st3);
await H.tab('Master');
const m2 = await H.cells();
ok(m2.includes('new type text'), 'Master got Col B text', m2);

console.log('fails', fails, 'errs', H.errs);
await H.close(true);
const sessionPath = path.join(sessionRoot('atm206'), 'cfg', 'click2copy', 'click2copy-session.json');
const s = JSON.parse(fs.readFileSync(sessionPath, 'utf8'));
const m = s.documents[0].data.tabs[0];
console.log('model master', m.rows, JSON.stringify(m.cells));
fs.rmSync(sessionRoot('atm206'), { recursive: true, force: true });
process.exit(fails ? 1 : 0);
