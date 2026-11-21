// 枠を固定した表のスクロールと見える位置：キー操作・名前ボックス・Ctrl+Home で移った先のセルが、固定した行や列の下に隠れない
// こと、固定したセルへ移ってもスクロールしている側の表示が動かないこと（行だけ・列だけの固定、すべての行や列の固定も）。
// 固定していない表は ui-scroll にある（共通の部品は lib/view.mjs）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {hidden, press, scrolled} from './lib/view.mjs';
import {blankBook} from './lib/workbook.mjs';

const env = useBrowser();

test('枠を固定した表でも、上や左へ移ったセルは固定した行・列の下に隠れない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 2, cols: 1}}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で、固定した行の下の行へ戻ったとき');
  await x.select('B5');
  await press(x, 'ArrowRight', 35);
  await press(x, 'ArrowLeft', 35);
  assert.equal(await x.nameBox(), 'B5');
  assert.equal(await hidden(x), null, '←で、固定した列のすぐ右の列へ戻ったとき');
});

// A frozen cell is always in view, so moving to and in frozen rows or columns leaves the scrolled part where it is.
test('枠を固定した表で、固定した行や列のセルへ移っても、スクロールしている側の表示は動かない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 2, cols: 1}}}));
  await x.select('B60');
  const {top} = await scrolled(x);
  assert.ok(top > 0, '前提: 下へスクロールしている');
  await x.select('B1');
  await press(x, 'ArrowRight', 3);
  assert.equal(await x.nameBox(), 'E1');
  assert.equal((await scrolled(x)).top, top, '固定した行へ移り、その中を動いても、縦の表示は変わらない');
  assert.equal(await hidden(x), null);
  await x.select('AN5');
  const {left} = await scrolled(x);
  assert.ok(left > 0, '前提: 右へスクロールしている');
  await x.select('A5');
  await press(x, 'ArrowDown', 3);
  assert.equal(await x.nameBox(), 'A8');
  assert.equal((await scrolled(x)).left, left, '固定した列へ移り、その中を動いても、横の表示は変わらない');
  assert.equal(await hidden(x), null);
});

test('行だけ・列だけを固定した表でも、上や左へ移ったセルは隠れない', async t => {
  const rows = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 1, cols: 0}}}));
  await rows.select('A5');
  await press(rows, 'ArrowRight', 35);
  await press(rows, 'ArrowLeft', 35);
  assert.equal(await rows.nameBox(), 'A5');
  assert.equal(await hidden(rows), null, '先頭行だけを固定した表で、←で戻ったとき');
  await rows.select('B60');
  await press(rows, 'ArrowUp', 40);
  assert.equal(await hidden(rows), null, '先頭行だけを固定した表で、↑で戻ったとき');
  const cols = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 0, cols: 1}}}));
  await cols.select('B60');
  await press(cols, 'ArrowUp', 40);
  assert.equal(await hidden(cols), null, '先頭列だけを固定した表で、↑で戻ったとき');
  await cols.select('B5');
  await press(cols, 'ArrowRight', 35);
  await press(cols, 'ArrowLeft', 35);
  assert.equal(await cols.nameBox(), 'B5');
  assert.equal(await hidden(cols), null, '先頭列だけを固定した表で、←で戻ったとき');
});

// The frozen part of a range that starts in frozen rows or columns is always in view; the scrolled part then shows from
// the range's first row (column) in it, so the whole range is seen when it fits.
test('名前ボックスで固定した行や列から始まる範囲を選んでも、画面に収まる範囲はすべて見える', async t => {
  const cols = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 0, cols: 1}}}));
  await cols.select('Z5');
  await cols.select('A5:D5');
  for (const a of ['A5', 'B5', 'C5', 'D5']) assert.equal(await hidden(cols, a), null, `固定した列 A から始まる範囲 A5:D5 の ${a}`);
  const rows = await env.open(t, blankBook({sheet: {rows: 300, cols: 12, freeze: {rows: 2, cols: 0}}}));
  await rows.select('B200');
  await rows.select('B1:B10');
  for (let r = 1; r <= 10; r++) assert.equal(await hidden(rows, 'B' + r), null, `固定した1〜2行目から始まる範囲 B1:B10 の B${r}`);
});

// With frozen panes Excel's Ctrl+Home goes to the first cell below and right of them, so the scrolled part shows its
// start; without them it goes to A1 (the test of Home and Ctrl+Home in ui-select.test.mjs).
test('枠を固定した表で Ctrl+Home を押すと、固定していない部分の左上のセルへ移り、その部分の表示も先頭へ戻る', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 60, freeze: {rows: 2, cols: 1}}}));
  await x.select('AD100');
  await x.press('Control+Home');
  assert.equal(await x.nameBox(), 'B3');
  assert.deepEqual(await scrolled(x), {top: 0, left: 0});
  assert.equal(await hidden(x), null);
  const rows = await env.open(t, blankBook({sheet: {rows: 300, cols: 12, freeze: {rows: 1, cols: 0}}}));
  await rows.select('C250');
  await rows.press('Control+Home');
  assert.equal(await rows.nameBox(), 'A2', '先頭行だけを固定した表では A2');
  assert.equal((await scrolled(rows)).top, 0);
});

// With every row (column) frozen nothing scrolls that way, so Ctrl+Home goes to the first row (column) then, as without
// frozen panes. Such a sheet is valid, and deleting the rows (columns) below (right of) the frozen ones leaves one.
test('すべての行や列を固定した表で Ctrl+Home を押すと、その向きでは1行目・1列目へ移る', async t => {
  const rows = await env.open(t, blankBook({sheet: {rows: 10, cols: 26, freeze: {rows: 10, cols: 1}}}));
  await rows.select('D5');
  await rows.press('Control+Home');
  assert.equal(await rows.nameBox(), 'B1', 'すべての行と1列を固定した表（最後の行の B10 へ移らない）');
  const cols = await env.open(t, blankBook({sheet: {rows: 30, cols: 6, freeze: {rows: 2, cols: 6}}}));
  await cols.select('D20');
  await cols.press('Control+Home');
  assert.equal(await cols.nameBox(), 'A3', 'すべての列と2行を固定した表（最後の列の F3 へ移らない）');
});
