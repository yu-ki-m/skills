// 拡大・縮小したときに見える位置：画面に見えていたアクティブセルが外へ出ないこと、元の倍率に戻すと左上のセルが戻ること、
// 固定した行や列の下に隠れていたアクティブセルは出さないこと。ほかの表示のテストは ui-viewport・ui-scroll・ui-scroll-frozen・
// ui-drag にある（共通の部品は lib/view.mjs）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {box, hidden, scrollBy} from './lib/view.mjs';
import {blankBook} from './lib/workbook.mjs';

const env = useBrowser();

// Excel keeps the selection in view when the zoom changes. The scroll offsets are screen pixels and the cells grow and
// shrink with the zoom, so keeping the offsets as they are would show other rows and columns at the new zoom.
// The cell at the top left of the scrolled part: just below the column headers and right of the row numbers.
const topLeft = x => x.evaluate(() => {
  const corner = document.querySelector('#grid thead th').getBoundingClientRect();
  return document.elementFromPoint(corner.right + 2, corner.bottom + 2)?.closest('td')?.id.slice('cell-'.length) ?? null;
});

test('ズームを縮小しても拡大しても、画面に見えていたアクティブセルは画面の外へ出ない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 60}}));
  await x.select('D100');
  const first = await topLeft(x);
  assert.equal(await hidden(x), null, '前提: 100% で D100 が見えている');
  await x.setControl('zoom', '50');
  assert.equal(await hidden(x), null, '50% に縮小したとき（D100 が画面の上の外へ出ない）');
  assert.equal(await hidden(x, first), null, `縮小すると見える範囲は広がるので、100% で左上に見えていた ${first} も見えたまま`);
  await x.setControl('zoom', '200');
  assert.equal(await hidden(x), null, '200% に拡大したとき（D100 が画面の下の外へ出ない）');
  await x.setControl('zoom', '100');
  await x.select('A1');
  await x.select('AD5');
  // Sideways likewise: the first column wholly seen right of the row numbers.
  const firstColumn = await x.evaluate(() => {
    const corner = document.querySelector('#grid thead th').getBoundingClientRect();
    return [...document.querySelectorAll('#grid td[data-r="4"]')].find(td => td.getBoundingClientRect().left >= corner.right).id.slice('cell-'.length);
  });
  assert.notEqual(firstColumn, 'A5', '前提: 右へスクロールしている');
  await x.setControl('zoom', '50');
  assert.equal(await hidden(x, firstColumn), null, `横も、100% で行番号のすぐ右に見えていた ${firstColumn} が、50% でも見える`);
  await x.setControl('zoom', '100');
  await x.select('A1');
  await x.select('D100');
  await x.click('#zoom-in');
  assert.equal(await x.textOf('#zoom-label'), '110%');
  assert.equal(await hidden(x), null, '＋ボタンで 110% にしたとき（画面の下の端にあった D100 が外へ出ない）');
  await x.setControl('zoom', '100');
  await x.select('A1');
  await x.select('AH120');
  await x.setControl('zoom', '50');
  assert.equal(await hidden(x), null, '右下へスクロールした表を 50% に縮小したとき');
  await x.setControl('zoom', '200');
  assert.equal(await hidden(x), null, '右下へスクロールした表を 200% に拡大したとき');
  const frozen = await env.open(t, blankBook({sheet: {rows: 200, cols: 60, freeze: {rows: 2, cols: 1}}}));
  await frozen.select('AH120');
  await frozen.setControl('zoom', '50');
  assert.equal(await hidden(frozen), null, '枠を固定した表を 50% に縮小したとき');
  await frozen.setControl('zoom', '200');
  assert.equal(await hidden(frozen), null, '枠を固定した表を 200% に拡大したとき');
});

test('長い表でズームを縮小してから元の倍率に戻すと、左上に元のセルが戻る', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 500, cols: 20}}));
  await x.select('A420');
  // The name box puts A420 at the bottom of the view, where bringing it into view again would also put it; with A420
  // moved up, only the offsets scaled with the zoom give the same view back.
  await scrollBy(x, 200);
  assert.equal(await hidden(x), null, '前提: A420 は画面の中ほどに見えている');
  const first = await topLeft(x);
  await x.setControl('zoom', '50');
  assert.equal(await hidden(x), null, '50% に縮小したとき（A420 が画面の上の外へ出ない）');
  await x.setControl('zoom', '100');
  assert.equal(await topLeft(x), first, 'スライダーで 100% に戻したとき');
  for (let i = 0; i < 5; i++) await x.click('#zoom-out');
  assert.equal(await x.textOf('#zoom-label'), '50%');
  assert.equal(await hidden(x), null, '－ボタンで 50% まで縮小したとき');
  for (let i = 0; i < 5; i++) await x.click('#zoom-in');
  assert.equal(await x.textOf('#zoom-label'), '100%');
  assert.equal(await topLeft(x), first, '＋ボタンで 100% に戻したとき');
});

// A cell under the frozen rows (columns) is not seen, though it is inside the view, so the zoom does not bring it out:
// the view stays where the scaled offsets put it.
test('アクティブセルが固定した行や列の下に隠れているときは、ズームしてもそのセルが見えるところへ表示を動かさない', async t => {
  for (const [address, sideways] of [['D60', false], ['Z5', true]]) {
    const x = await env.open(t, blankBook({sheet: {rows: 300, cols: 60, freeze: {rows: 4, cols: 2}}}));
    await x.select(address);
    // Scrolls until the middle of the cell is in the middle of the frozen rows (columns).
    const first = await box(x, sideways ? '#grid th[data-col="0"]' : '#grid th[data-row="0"]'), last = await box(x, sideways ? '#grid th[data-col="1"]' : '#grid th[data-row="3"]'), cell = await box(x, '#cell-' + address);
    await (sideways ? scrollBy(x, 0, (cell.left + cell.right - first.left - last.right) / 2) : scrollBy(x, (cell.top + cell.bottom - first.top - last.bottom) / 2));
    const where = sideways ? '固定した列' : '固定した行';
    assert.notEqual(await hidden(x), null, `前提: ${address} は${where}の下に隠れている`);
    await x.setControl('zoom', '150');
    assert.notEqual(await hidden(x), null, `150% にしても、${address} を${where}の下から出さない`);
  }
});
