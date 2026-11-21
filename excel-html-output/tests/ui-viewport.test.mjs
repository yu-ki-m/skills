// 画面の表示：シートを追加したとき（元に戻す・やり直すときも）に見える位置、見出しと選択の枠・フィルハンドル・編集中のセル・
// 左上の角の重なり（フィルハンドルに重なった列見出しの境界をドラッグしても、フィルにならないことも）、編集を始めたセルが見出し
// の下から出ること。キー操作で移った先のセルが隠れないことは ui-scroll と ui-scroll-frozen に、拡大・縮小は ui-zoom に、範囲を
// 選ぶドラッグとフィルハンドルのドラッグは ui-drag にある（共通の部品は lib/view.mjs）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {box, hidden, middle, scrollBy, scrolled} from './lib/view.mjs';
import {blankBook} from './lib/workbook.mjs';

const env = useBrowser();

// Excel keeps the view of each sheet (the top-left cell and the selection) apart, so a new sheet opens at its start, with
// A1 at the top left, wherever the sheet before it was scrolled to. When no sheet can be added, the view stays as it was.
test('スクロールした状態で＋でシートを追加すると、新しいシートは A1 を左上にして開き、アクティブセル A1 が見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 200, cols: 60}}));
  await x.select('AD150');
  const before = await scrolled(x);
  assert.ok(before.top > 0 && before.left > 0, '前提: 下と右へスクロールしている');
  await x.click('#add-sheet');
  assert.equal(await x.textOf('#sheet-tabs [aria-selected="true"]'), 'Sheet2', '追加したシートが開く');
  assert.equal(await x.nameBox(), 'A1');
  assert.deepEqual(await scrolled(x), {top: 0, left: 0}, '前のシートのスクロール位置（新しいシートの端で切り詰めた位置）のまま開かない');
  assert.equal(await hidden(x), null, 'アクティブセル A1 が画面の外に出ない');
});

test('シートが30枚あって＋で追加できないときは、今のシートの表示と選択が変わらない', async t => {
  const book = blankBook({sheet: {rows: 200, cols: 60}});
  for (let i = 2; i <= 30; i++) book.sheets.push({id: 's' + i, name: 'Sheet' + i, rows: 10, cols: 5, cells: {}, merges: [], rowHeights: [], colWidths: []});
  const x = await env.open(t, book);
  await x.select('AD150');
  const before = await scrolled(x);
  await x.click('#add-sheet');
  assert.equal((await x.dialog())?.title, '操作を完了できません');
  await x.answer('ok');
  assert.equal(await x.evaluate(() => document.querySelectorAll('#sheet-tabs [data-sheet]').length), 30, 'シートは増えない');
  assert.equal(await x.nameBox(), 'AD150');
  assert.deepEqual(await scrolled(x), before, '表示は先頭へ戻らない');
  assert.equal(await hidden(x), null, 'アクティブセル AD150 が見えたまま');
});

// Undoing or redoing the addition of a sheet changes the sheet in view, which then shows its active cell, wherever the
// other sheet was scrolled to: the sheet before comes back with its active cell in view (at its start when the cell is
// near it), and the new one opens at its start, as when it was added.
test('＋でシートを追加した後に Ctrl+Z や Ctrl+Y を押しても、開いたシートは前に見ていたシートのスクロール位置に左右されず、アクティブセルが見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 100, cols: 26}}));
  const tab = () => x.textOf('#sheet-tabs [aria-selected="true"]');
  await x.clickCell('B3');
  await x.click('#add-sheet');
  await x.select('Z100');
  const away = await scrolled(x);
  assert.ok(away.top > 0 && away.left > 0, '前提: 追加したシートで下と右へスクロールしている');
  await x.press('Control+z');
  assert.equal(await x.nameBox(), 'B3', '前提: Ctrl+Z で元のシートの B3 に戻る');
  assert.deepEqual(await scrolled(x), {top: 0, left: 0}, 'Ctrl+Z で戻ったシートは、追加したシートのスクロール位置からではなく、先頭から開く（B3 が左上の端に来て、1〜2行目と A 列が隠れることはない）');
  await x.select('Z90');
  await x.click('#add-sheet');
  assert.equal(await tab(), 'Sheet2', '前提: 追加したシートが開く');
  await x.press('Control+z');
  assert.equal(await tab(), 'Sheet1', '前提: Ctrl+Z で元のシートに戻る');
  assert.equal(await x.nameBox(), 'Z90');
  assert.equal(await hidden(x), null, 'Ctrl+Z で戻ったシートで、アクティブセル Z90 が見える');
  await x.select('Z90');
  assert.ok((await scrolled(x)).top > 0, '前提: 元のシートで下へスクロールしている');
  await x.press('Control+y');
  assert.equal(await tab(), 'Sheet2', '前提: Ctrl+Y で追加をやり直す');
  assert.equal(await x.nameBox(), 'A1');
  assert.deepEqual(await scrolled(x), {top: 0, left: 0}, 'Ctrl+Y で開いた新しいシートは、A1 を左上にして開く');
});

// In Excel the column headers and the row numbers lie outside the area the cells scroll in, so what belongs to the cells
// (the green frame, its fill handle, the cell being edited) goes out of sight with them and never covers a header.
// What is drawn in front at a point is what the pointer reaches there; the frame itself does not take the pointer
// (clicks go through it to the cells), so it is made to take it for the moment.
const inFront = (x, px, py) => x.evaluate((px, py) => {
  const frame = document.getElementById('selection');
  frame.style.pointerEvents = 'auto';
  const hit = document.elementFromPoint(px, py);
  frame.style.pointerEvents = '';
  const head = hit?.closest('#grid th');
  if (head) return head.dataset.corner ? '左上の角' : (head.dataset.col !== undefined ? '列見出し ' : '行番号 ') + head.textContent;
  return hit ? (hit.id ? '#' + hit.id : hit.tagName) : 'なし';
}, px, py);

test('選んだセルが列見出しや行番号の下へスクロールしても、緑の枠と編集中のセルは見出しの手前に描かれず、セルと一緒に隠れる', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  await x.select('B1:B90');
  await x.press('Shift+Enter');
  let frame = await box(x, '#selection'), head = await box(x, '#grid th[data-col="1"]');
  assert.ok(frame.top < head.top, '前提: 枠の上の端は画面の上の外にあり、枠の左右の線が列見出し B の位置を通る');
  assert.equal(await inFront(x, frame.left + 1, middle(head).y), '列見出し B', '枠の左の線');
  assert.equal(await inFront(x, frame.right - 1, middle(head).y), '列見出し B', '枠の右の線');

  await x.select('B5:AJ5');
  await scrollBy(x, 0, 400);
  frame = await box(x, '#selection'); head = await box(x, '#grid th[data-row="4"]');
  assert.ok(frame.left < head.left, '前提: 枠の左の端は画面の左の外にあり、枠の上下の線が行番号 5 の位置を通る');
  assert.equal(await inFront(x, middle(head).x, frame.top + 1), '行番号 5', '枠の上の線');
  assert.equal(await inFront(x, middle(head).x, frame.bottom - 1), '行番号 5', '枠の下の線');

  await x.clickCell('B4');
  await x.type('入力中');
  assert.equal(await x.editing(), true, '前提: B4 を編集している');
  head = await box(x, '#grid th[data-col="1"]');
  await scrollBy(x, (await box(x, '#cell-editor')).top - head.top);
  assert.equal(await inFront(x, middle(head).x, middle(head).y), '列見出し B', '入力中にスクロールして、編集中のセルが列見出しの位置に来たとき');

  // The headers of frozen rows and columns are drawn over the other headers, and over the frame as well.
  const frozen = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 1, cols: 1}}}));
  await frozen.select('A4:C4');
  head = await box(frozen, '#grid th[data-col="0"]');
  await scrollBy(frozen, (await box(frozen, '#selection')).bottom - middle(head).y);
  frame = await box(frozen, '#selection');
  assert.ok(frame.bottom > head.top && frame.bottom < head.bottom, '前提: 枠の下の線が、固定した列の見出し A の位置を通る');
  assert.equal(await inFront(frozen, middle(head).x, frame.bottom - 1), '列見出し A', '固定した列の見出し');
  await frozen.select('B3:Z3');
  head = await box(frozen, '#grid th[data-row="0"]');
  await scrollBy(frozen, (await box(frozen, '#selection')).top - middle(head).y, 400);
  frame = await box(frozen, '#selection');
  assert.ok(frame.left < head.left && frame.top > head.top && frame.top < head.bottom, '前提: 枠の上の線が、固定した行の行番号 1 の位置を通る');
  assert.equal(await inFront(frozen, middle(head).x, frame.top + 1), '行番号 1', '固定した行の行番号');
});

test('選んだセルのフィルハンドルが列見出しの位置までスクロールしても、見出しの境界をドラッグすると列幅が変わり、セルに値はコピーされない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 12, cells: {B4: {t: 'n', v: 7}}}}));
  await x.clickCell('B4');
  // Scroll down a little, as one turn of the mouse wheel does, until the fill handle at the bottom right of B4 is in
  // the middle of the column headers: on the boundary between B and C, where a drag widens column B.
  const head = await box(x, '#grid th[data-col="1"]');
  await scrollBy(x, middle(await box(x, '#fill-handle')).y - middle(head).y);
  const handle = await box(x, '#fill-handle'), from = {x: middle(await box(x, '#grid th[data-col="1"] .col-resize')).x, y: middle(handle).y};
  assert.ok(from.x > handle.left && from.x < handle.right && from.y > head.top && from.y < head.bottom, '前提: 列 B と C の境界が、フィルハンドルと同じ位置にある');
  assert.equal(await inFront(x, from.x, from.y), '列見出し B', 'フィルハンドルは列見出しの下に隠れる');
  await x.page.drag(from, {x: from.x + 60, y: from.y});
  assert.equal(await x.text('C4'), '', 'フィルにならない（B4 の値が C4 にコピーされない）');
  assert.ok((await box(x, '#grid th[data-col="1"]')).width >= head.width + 50, '列 B の幅が広がる');
});

// The corner (select all) stays at the top left over both headers, which scroll under it.
test('表を右や下へスクロールしても、左上の角は列見出しと行番号の手前にあり、押すとシート全体が選ばれる', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  for (const [dy, dx, name] of [[0, 300, '右へ'], [300, -300, '下へ'], [0, 300, '右と下へ']]) {
    await scrollBy(x, dy, dx);
    const corner = middle(await box(x, '#grid thead th'));
    assert.equal(await inFront(x, corner.x, corner.y), '左上の角', name + 'スクロールしたとき');
  }
  await x.click('#grid thead th');
  assert.equal(await x.nameBox(), 'A1:AN120', '角の下に流れてきた列や行ではなく、シート全体が選ばれる');
  const frozen = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 2, cols: 1}}}));
  await scrollBy(frozen, 300, 300);
  const corner = middle(await box(frozen, '#grid thead th'));
  assert.equal(await inFront(frozen, corner.x, corner.y), '左上の角', '枠を固定した表で、右と下へスクロールしたとき');
});

// Excel brings the active cell into view when editing starts in it, so a cell scrolled under the column headers or the
// row numbers comes out from under them, and what is typed is seen in the cell, not only in the formula bar.
// Whether the cell being edited is seen is checked at its top-left and bottom-right corners and its centre.
const editorHidden = x => x.evaluate(() => {
  const editor = document.getElementById('cell-editor'), r = editor.getBoundingClientRect();
  for (const [name, px, py] of [['左上', r.left + 4, r.top + 4], ['右下', r.right - 4, r.bottom - 4], ['中央', (r.left + r.right) / 2, (r.top + r.bottom) / 2]]) {
    const hit = document.elementFromPoint(px, py);
    if (hit !== editor) return '編集中のセルの' + name + 'が ' + (hit ? hit.tagName + ' ' + hit.textContent.slice(0, 8) : '画面の外') + ' の下にある';
  }
  return null;
});

test('列見出しや行番号の下へスクロールしたセルで編集を始めると、そのセルが見出しの下から出て、編集中のセルが見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  // Selects the cell and scrolls until its top (left) edge is shift pixels below (right of) the top (left) edge of the
  // headers.
  const under = async (address, sideways, shift = 0) => {
    await x.clickCell(address);
    const head = await box(x, sideways ? '#grid th[data-row="4"]' : '#grid th[data-col="1"]'), cell = await box(x, '#cell-' + address);
    await (sideways ? scrollBy(x, 0, cell.left - head.left - shift) : scrollBy(x, cell.top - head.top - shift));
    assert.notEqual(await hidden(x), null, `前提: ${address} が${sideways ? '行番号' : '列見出し'}の下にある`);
  };
  for (const [name, start] of [['文字を打ったとき', () => x.type('abc')], ['F2 を押したとき', () => x.press('F2')]]) {
    await under('B4');
    await start();
    assert.equal(await x.editing(), true, '前提: 編集を始めた');
    assert.equal(await editorHidden(x), null, '列見出しの下にあった B4 で、' + name);
    await x.press('Escape');
  }
  await under('B4', false, 12);
  const head = await box(x, '#grid th[data-col="1"]');
  await x.page.dblclick({x: middle(head).x, y: head.bottom + 4});
  assert.equal(await x.editing(), true, '前提: 半分が列見出しの下にある B4 の、見えている下の部分をダブルクリックして編集を始めた');
  assert.equal(await editorHidden(x), null, '半分が列見出しの下にあった B4 をダブルクリックしたとき');
  await x.press('Escape');
  await under('C5', true);
  await x.type('abc');
  assert.equal(await editorHidden(x), null, '行番号の下にあった C5 で、文字を打ったとき');
  await x.press('Escape');
  await under('B4');
  await x.page.compose('にほん');
  assert.equal(await editorHidden(x), null, '列見出しの下にあった B4 で、日本語入力を始めたとき');
  await x.type('にほん');
  await x.press('Escape');
});
