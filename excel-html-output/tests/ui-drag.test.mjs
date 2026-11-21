// ドラッグで範囲を選ぶときの表示：列見出しや行番号の上、その先まで動かすと表が流れ、選択はスクロールで画面に入った行・列まで
// 広がること（拡大・縮小した表、枠を固定した表、見出しのドラッグ、フィルハンドルのドラッグも）と、フィルハンドルの上で少し
// 動かしただけではフィルしないこと。ほかの表示のテストは ui-viewport・ui-scroll・ui-scroll-frozen・ui-zoom にある（共通の
// 部品は lib/view.mjs）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {box, middle, press, scrolled} from './lib/view.mjs';
import {blankBook} from './lib/workbook.mjs';

const env = useBrowser();

// Excel scrolls the sheet while a drag goes on past the top or left edge of the cells, over the column headers or the
// row numbers, and the selection follows the scroll: it runs to the first row (column) in view below (right of) them.
// The rows and columns in view next to the headers: the first one partly seen and the first one wholly seen.
const besideHeaders = x => x.evaluate(() => {
  const C = window.ExcelHTMLCore, corner = document.querySelector('#grid thead th').getBoundingClientRect();
  const rows = [...document.querySelectorAll('#grid th[data-row]')].map(th => th.getBoundingClientRect());
  const cols = [...document.querySelectorAll('#grid th[data-col]')].map(th => th.getBoundingClientRect());
  return {
    rows: [rows.findIndex(b => b.bottom > corner.bottom) + 1, rows.findIndex(b => b.top >= corner.bottom) + 1],
    cols: [cols.findIndex(b => b.right > corner.right), cols.findIndex(b => b.left >= corner.right)].map(c => C.colName(c)),
  };
});
// Presses at from, moves onto to, keeps moving there a pixel to and fro (the page scrolls a step on every move while
// the pointer is near the top or left edge), and releases.
async function dragAndHold(x, from, to, moves) {
  const mouse = x.page.mouse;
  await mouse.move(from.x, from.y);
  await mouse.down(from.x, from.y);
  for (let i = 1; i <= 4; i++) await mouse.move(from.x + (to.x - from.x) * i / 4, from.y + (to.y - from.y) * i / 4, {buttons: 1});
  for (let i = 0; i < moves; i++) await mouse.move(to.x + (i % 2 ? -1 : 1), to.y, {buttons: 1});
  await mouse.up(to.x, to.y);
}

test('ドラッグで列見出しや行番号の上まで動かして上や左へスクロールさせると、選択はスクロールで画面に入った行・列まで広がる', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 30}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 5);
  const before = await scrolled(x);
  await dragAndHold(x, await x.point('#cell-B55'), await x.point('#grid th[data-col="3"]'), 12);
  assert.ok((await scrolled(x)).top < before.top - 200, '前提: ドラッグの間に表が上へスクロールした');
  const {rows} = await besideHeaders(x);
  assert.ok(rows.map(r => 'B' + r + ':D55').includes(await x.nameBox()),
    `選択（${await x.nameBox()}）が、列見出しのすぐ下に見えている ${rows.join('〜')} 行目まで広がる（起点の55行目だけに縮まない）`);
  assert.equal(await x.activeCell(), 'B55', '起点のセルがアクティブのまま');

  await x.select('A5');
  await press(x, 'ArrowRight', 29);
  await press(x, 'ArrowLeft', 9);
  const start = await scrolled(x);
  await dragAndHold(x, await x.point('#cell-U5'), await x.point('#grid th[data-row="7"]'), 12);
  assert.ok((await scrolled(x)).left < start.left - 200, '前提: ドラッグの間に表が左へスクロールした');
  const {cols} = await besideHeaders(x);
  assert.ok(cols.map(c => c + '5:U8').includes(await x.nameBox()),
    `選択（${await x.nameBox()}）が、行番号のすぐ右に見えている ${cols.join('〜')} 列まで広がる（起点のU列だけに縮まない）`);
});

// Frozen rows stay right below the column headers, so they are the rows in view there: a drag that goes from a frozen
// row up onto the headers runs to row 1, and never down to the scrolled rows below the frozen ones.
test('枠を固定した表で、固定した行のセルから列見出しの上へドラッグすると、選択は1行目まで広がり、起点より下へは広がらない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 30, freeze: {rows: 2, cols: 1}}}));
  await x.select('C60');
  assert.ok((await scrolled(x)).top > 0, '前提: 下へスクロールしている');
  await dragAndHold(x, await x.point('#cell-C2'), await x.point('#grid th[data-col="3"]'), 3);
  assert.equal(await x.nameBox(), 'C1:D2');
});

// The drag scrolls while the pointer is past the edge of the cells, that is over the column headers or the row numbers,
// whose size follows the zoom: at 200% the whole of the taller (wider) headers, at 50% not the first row (column) of
// cells below (right of) them.
test('拡大・縮小した表でも、ドラッグの自動スクロールは列見出しや行番号の上全体で始まり、そのすぐ下や右のセルの上では始まらない', async t => {
  const big = await env.open(t, blankBook({view: {zoom: 200}, sheet: {rows: 200, cols: 30}}));
  await big.select('B60');
  await press(big, 'ArrowUp', 3);
  let before = await scrolled(big), head = await box(big, '#grid th[data-col="3"]');
  await dragAndHold(big, await big.point('#cell-B57'), {x: middle(head).x, y: head.bottom - 4}, 6);
  assert.ok((await scrolled(big)).top < before.top, '200%: 列見出しの下の端の近く（100%の見出しの高さより下）でも、上へスクロールする');
  await big.select('AD5');
  before = await scrolled(big);
  assert.ok(before.left > 0, '前提: 右へスクロールしている');
  head = await box(big, '#grid th[data-row="4"]');
  await dragAndHold(big, await big.point('#cell-AD5'), {x: head.right - 4, y: middle(head).y}, 6);
  assert.ok((await scrolled(big)).left < before.left, '200%: 行番号の右の端の近く（100%の行番号の幅より右）でも、左へスクロールする');
  const small = await env.open(t, blankBook({view: {zoom: 50}, sheet: {rows: 300, cols: 30}}));
  await small.select('B250');
  before = await scrolled(small);
  head = await box(small, '#grid th[data-col="3"]');
  await dragAndHold(small, await small.point('#cell-B250'), {x: middle(head).x, y: head.bottom + 6}, 6);
  assert.equal((await scrolled(small)).top, before.top, '50%: 列見出しのすぐ下のセルの上では、スクロールしない');
  const wide = await env.open(t, blankBook({view: {zoom: 50}, sheet: {rows: 30, cols: 100}}));
  await wide.select('CV5');
  before = await scrolled(wide);
  assert.ok(before.left > 0, '前提: 右へスクロールしている');
  head = await box(wide, '#grid th[data-row="4"]');
  await dragAndHold(wide, await wide.point('#cell-CV5'), {x: head.right + 6, y: middle(head).y}, 6);
  assert.equal((await scrolled(wide)).left, before.left, '50%: 行番号のすぐ右のセルの上では、スクロールしない');
});

// A drag over the column headers selects whole columns, so, as in Excel, the sheet scrolls only sideways then, though
// the pointer stays past the top edge of the cells or goes past the bottom one; over the row numbers it scrolls only up
// and down, past the left edge or the right one.
test('列見出しや行番号をドラッグして列・行を選ぶ間は、表は縦（行番号なら横）にスクロールしない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 300, cols: 60}}));
  await x.select('C150');
  let from = await x.point('#grid th[data-col="2"]'), to = await x.point('#grid th[data-col="4"]'), before = await scrolled(x);
  assert.ok(before.top > 0, '前提: 下へスクロールしている');
  await dragAndHold(x, from, to, 12);
  assert.equal(await x.nameBox(), 'C1:E300');
  assert.equal((await scrolled(x)).top, before.top, '列見出しの上でドラッグしても、表は上へ流れない');
  await dragAndHold(x, from, {x: from.x, y: middle(await box(x, '.statusbar')).y}, 12);
  assert.equal(await x.nameBox(), 'C1:C300');
  assert.equal((await scrolled(x)).top, before.top, '列見出しからステータスバーの上まで動かしても、表は下へ流れない');
  await x.select('AH16');
  from = await x.point('#grid th[data-row="15"]'); to = await x.point('#grid th[data-row="18"]'); before = await scrolled(x);
  assert.ok(before.left > 0, '前提: 右へスクロールしている');
  await dragAndHold(x, from, to, 12);
  assert.equal(await x.nameBox(), 'A16:BH19');
  assert.equal((await scrolled(x)).left, before.left, '行番号の上でドラッグしても、表は左へ流れない');
  await dragAndHold(x, from, {x: (await box(x, '#viewport')).right - 4, y: from.y}, 12);
  assert.equal(await x.nameBox(), 'A16:BH16');
  assert.equal((await scrolled(x)).left, before.left, '行番号から縦のスクロールバーの上まで動かしても、表は右へ流れない');
});

// The rows (columns) in view at the bottom (right) edge of the cells, above the horizontal scroll bar (left of the vertical
// one): the last one wholly seen and the last one partly seen.
const atEnd = x => x.evaluate(() => {
  const C = window.ExcelHTMLCore, v = document.getElementById('viewport'), b = v.getBoundingClientRect();
  const bottom = b.top + v.clientTop + v.clientHeight, right = b.left + v.clientLeft + v.clientWidth;
  const rows = [...document.querySelectorAll('#grid th[data-row]')].map(th => th.getBoundingClientRect());
  const cols = [...document.querySelectorAll('#grid th[data-col]')].map(th => th.getBoundingClientRect());
  return {
    rows: [rows.findLastIndex(r => r.bottom <= bottom) + 1, rows.findLastIndex(r => r.top < bottom) + 1],
    cols: [cols.findLastIndex(c => c.right <= right), cols.findLastIndex(c => c.left < right)].map(c => C.colName(c)),
  };
});

// Past the headers the pointer is over the formula bar or the corner, past the bottom of the cells over the sheet tabs or
// the status bar, and past their right edge over the vertical scroll bar. The sheet scrolls on there, and, as in Excel,
// the selection follows it to the row or column in view at that edge. Past the end of a table that ends inside the view,
// the selection runs to its last row or column.
test('ドラッグで見出しを越えて数式バーや左上の角、表の下や右の外まで動かしても、選択はその端に見えている行・列まで広がる', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 60}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 5);
  let before = await scrolled(x);
  await dragAndHold(x, await x.point('#cell-B55'), {x: middle(await box(x, '#grid th[data-col="1"]')).x, y: middle(await box(x, '.formula-bar')).y}, 12);
  assert.ok((await scrolled(x)).top < before.top - 200, '前提: ドラッグの間に表が上へスクロールした');
  const {rows} = await besideHeaders(x);
  assert.ok(rows.map(r => 'B' + r + ':B55').includes(await x.nameBox()),
    `数式バーの上: 選択（${await x.nameBox()}）が、列見出しのすぐ下に見えている ${rows.join('〜')} 行目まで広がる`);

  await x.select('AD97');
  before = await scrolled(x);
  await dragAndHold(x, await x.point('#cell-AD97'), middle(await box(x, '#grid thead th')), 12);
  const now = await scrolled(x);
  assert.ok(now.top < before.top - 200 && now.left < before.left - 200, '前提: ドラッグの間に表が上と左へスクロールした');
  const seen = await besideHeaders(x), corner = seen.cols.flatMap(c => seen.rows.map(r => c + r + ':AD97'));
  assert.ok(corner.includes(await x.nameBox()), `左上の角の上: 選択（${await x.nameBox()}）が、左上に見えているセル（${corner.join('、')}）まで広がる`);

  await x.select('B5');
  before = await scrolled(x);
  const start = await x.point('#cell-B5');
  await dragAndHold(x, start, {x: start.x, y: middle(await box(x, '.statusbar')).y}, 12);
  assert.ok((await scrolled(x)).top > before.top + 200, '前提: ドラッグの間に表が下へスクロールした');
  const last = (await atEnd(x)).rows;
  assert.ok(last.map(r => 'B5:B' + r).includes(await x.nameBox()),
    `ステータスバーの上: 選択（${await x.nameBox()}）が、下の端に見えている ${last.join('〜')} 行目まで広がる`);

  // The page's right edge is the sheet's, so a drag to the right runs onto the vertical scroll bar, where no column header
  // is under the pointer.
  await x.select('B5');
  before = await scrolled(x);
  const from = await x.point('#cell-B5'), bar = {x: (await box(x, '#viewport')).right - 4, y: from.y};
  assert.equal(await x.evaluate(p => document.elementFromPoint(p.x, p.y)?.id, bar), 'viewport', '前提: 縦のスクロールバーの上（セルの上ではない）');
  await dragAndHold(x, from, bar, 12);
  assert.ok((await scrolled(x)).left > before.left + 200, '前提: ドラッグの間に表が右へスクロールした');
  const right = (await atEnd(x)).cols;
  assert.ok(right.map(c => 'B5:' + c + '5').includes(await x.nameBox()),
    `縦のスクロールバーの上: 選択（${await x.nameBox()}）が、右の端に見えている ${right.join('〜')} 列まで広がる`);

  const small = await env.open(t, blankBook({sheet: {rows: 12, cols: 8}}));
  const table = await box(small, '#grid'), view = await box(small, '#viewport');
  assert.ok(table.right < view.right - 100 && table.bottom < view.bottom - 100, '前提: 表の右と下に空きがある');
  await dragAndHold(small, await small.point('#cell-B5'), {x: table.right + 50, y: table.bottom + 50}, 2);
  assert.equal(await small.nameBox(), 'B5:H12', '表の右下の外まで動かすと、最後の行と列まで広がる');
});

// The fill handle is drawn over the selection's bottom-right corner, so part of it lies over the cells right of and below
// the selection. A press on the handle that moves a pixel or two and stays on it has not reached another cell: it fills
// nothing. The same press carried on into the next cell fills it.
test('フィルハンドルを押してその上で少し動かしただけで離しても、右や下の隣のセルは埋まらない', async t => {
  const x = await env.open(t, blankBook({sheet: {cells: {B2: {t: 'n', v: 5}, C2: {t: 's', v: 'keep'}, B3: {t: 's', v: 'keep'}, C3: {t: 's', v: 'keep'}}}}));
  await x.clickCell('B2');
  const handle = await box(x, '#fill-handle'), cell = await box(x, '#cell-B2'), centre = middle(handle);
  assert.ok(handle.right > cell.right + 1 && handle.bottom > cell.bottom + 1, '前提: フィルハンドルは B2 の右と下へはみ出している');
  const right = (cell.right + handle.right) / 2, below = (cell.bottom + handle.bottom) / 2;
  for (const [where, to] of [['右（C2 の上）', {x: right, y: centre.y}], ['下（B3 の上）', {x: centre.x, y: below}], ['右下（C3 の上）', {x: right, y: below}]]) {
    assert.equal(await x.evaluate(p => document.elementFromPoint(p.x, p.y)?.id, to), 'fill-handle', `前提: ${where}の点もフィルハンドルの上`);
    await x.page.mouse.move(centre.x, centre.y);
    await x.page.mouse.down(centre.x, centre.y);
    await x.page.mouse.move(to.x, to.y, {buttons: 1});
    await x.page.mouse.up(to.x, to.y);
    assert.deepEqual([await x.text('C2'), await x.text('B3'), await x.text('C3'), await x.nameBox()], ['keep', 'keep', 'keep', 'B2'],
      `フィルハンドルの中央で押して、ハンドルの${where}へ動かして離したとき`);
  }
  await x.page.drag(centre, await x.point('#cell-C2'));
  assert.equal(await x.text('C2'), '5', '前提: 同じように押して C2 まで動かせば、C2 が埋まる');
});

// A drag of the fill handle scrolls the sheet and follows the scroll as a drag over the cells does: past the bottom of the
// cells (over the status bar) the fill runs to the row in view at the bottom edge, and over the column headers to the row
// in view right below them.
test('フィルハンドルをステータスバーや列見出しの上まで動かすと、表が流れ、フィルはその端に見えている行まで広がる', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 300, cols: 8, cells: {B5: {t: 'n', v: 7}, B55: {t: 'n', v: 8}}}}));
  const columnB = middle(await box(x, '#grid th[data-col="1"]'));
  await x.clickCell('B5');
  let before = await scrolled(x);
  await dragAndHold(x, await x.point('#fill-handle'), {x: columnB.x, y: middle(await box(x, '.statusbar')).y}, 12);
  assert.ok((await scrolled(x)).top > before.top + 200, 'ステータスバーの上: 表が下へ流れる');
  const last = (await atEnd(x)).rows, down = await x.nameBox();
  assert.ok(last.map(r => 'B5:B' + r).includes(down), `ステータスバーの上: フィル（${down}）が、下の端に見えている ${last.join('〜')} 行目まで広がる`);
  assert.equal(await x.text(down.split(':')[1]), '7', `${down.split(':')[1]} まで 7 で埋まる`);

  await x.select('B60');
  await press(x, 'ArrowUp', 5);
  before = await scrolled(x);
  await dragAndHold(x, await x.point('#fill-handle'), columnB, 12);
  assert.ok((await scrolled(x)).top < before.top - 200, '列見出しの上: 表が上へ流れる');
  const {rows} = await besideHeaders(x), up = await x.nameBox();
  assert.ok(rows.map(r => 'B' + r + ':B55').includes(up), `列見出しの上: フィル（${up}）が、列見出しのすぐ下に見えている ${rows.join('〜')} 行目まで広がる`);
  assert.equal(await x.text(up.split(':')[0]), '8', `${up.split(':')[0]} まで 8 で埋まる`);
});
