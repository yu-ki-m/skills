// セルの選択と移動：クリック、Shift+クリック、ドラッグ、見出し、キーボード、結合セル、名前ボックス。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {blankBook, fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

test('クリックしたセルが選ばれ、名前ボックスと数式バーにその番地と中身が出る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C4');
  assert.equal(await x.nameBox(), 'C4');
  assert.equal(await x.formula(), '3');
  await x.clickCell('E4');
  assert.equal(await x.formula(), '=C4*D4', '数式のセルは、数式バーに値ではなく数式が出る');
  assert.equal(await x.text('E4'), '¥36,000');
  assert.equal(await x.activeElement(), 'cell-editor', 'クリックの後も、キー入力はシートに届く');
});

test('Shift+クリックで範囲を広げると、選択の大きさと、ステータスバーの合計・個数が出る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C4');
  await x.clickCell('E6', {shift: true});
  assert.equal(await x.nameBox(), 'C4:E6');
  assert.equal(await x.activeCell(), 'C4', '起点のセルがアクティブのまま');
  assert.equal(await x.selectionSize(), '3行 × 3列');
  const stats = await x.stats();
  assert.match(stats, /データの個数: 9/);
  assert.match(stats, /合計: 123006/);
  assert.match(stats, /平均: 13667\.33/);
});

test('マウスでドラッグした範囲が選ばれ、緑の選択枠がその範囲に重なる', async t => {
  const x = await env.open(t, book);
  await x.page.drag(await x.point('#cell-B4'), await x.point('#cell-D6'));
  assert.equal(await x.nameBox(), 'B4:D6');
  assert.equal(await x.activeCell(), 'B4');
  assert.ok(await x.selectionFrameGap('B4', 'D6') <= 2, '選択枠の位置');
});

test('行見出し・列見出し・左上の角をクリックすると、行全体・列全体・シート全体が選ばれる', async t => {
  const x = await env.open(t, book);
  await x.click('#grid th[data-row="3"]');
  assert.equal(await x.nameBox(), 'A4:J4');
  await x.click('#grid th[data-col="3"]');
  assert.equal(await x.nameBox(), 'D1:D30');
  await x.click('#grid th[data-col="2"]');
  assert.equal(await x.nameBox(), 'C1:C30', 'Excelと同じく、結合セル（A11:C12）がかかる列でも、押した列だけを選ぶ');
  await x.click('#grid th[data-corner]');
  assert.equal(await x.nameBox(), 'A1:J30');
  assert.equal(await x.selectionSize(), '30行 × 10列');
});

test('結合セルが一部の行にあっても、見出しで選ぶと押した列・行だけが選ばれ、ドラッグや Shift で広げても結合の分まで広がらない', async t => {
  const x = await env.open(t, book);
  // A11:C12 is merged: it crosses columns A to C and rows 11 to 12.
  await x.click('#grid th[data-col="1"]');
  assert.equal(await x.nameBox(), 'B1:B30');
  assert.equal(await x.selectionSize(), '30行 × 1列');
  assert.ok(await x.selectionFrameGap('B1', 'B30') <= 1, '緑の枠も列Bだけを囲む');
  await x.click('#grid th[data-row="10"]');
  assert.equal(await x.nameBox(), 'A11:J11');
  assert.equal(await x.selectionSize(), '1行 × 10列');
  await x.page.drag(await x.point('#grid th[data-col="1"]'), await x.point('#grid th[data-col="2"]'));
  assert.equal(await x.nameBox(), 'B1:C30', 'ドラッグで選んだ2列だけ');
  await x.click('#grid th[data-col="1"]');
  await x.click('#grid th[data-col="2"]', {shift: true});
  assert.equal(await x.nameBox(), 'B1:C30', 'Shift+クリックで選んだ2列だけ');
  await x.click('#grid th[data-col="1"]');
  await x.press('Shift+ArrowRight');
  assert.equal(await x.nameBox(), 'B1:C30', 'Shift+→ で広げた2列だけ');
  await x.clickCell('B10');
  await x.clickCell('B11', {shift: true});
  assert.equal(await x.nameBox(), 'A10:C12', 'セルの範囲を選んだときは、これまでどおり結合セル全体を含む');
});

test('列の一番上のセルが結合セルの一部でも、見出しで選んだ列だけを緑の枠で囲む', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 8, cols: 5, cells: {B1: {t: 's', v: '題名'}}, merges: [{r1: 0, c1: 1, r2: 0, c2: 2}]}}));
  await x.click('#grid th[data-col="2"]');
  assert.equal(await x.nameBox(), 'C1:C8');
  const gap = await x.evaluate(() => {
    const frame = document.getElementById('selection').getBoundingClientRect(), box = s => document.querySelector(s).getBoundingClientRect();
    const col = box('#grid th[data-col="2"]'), first = box('#grid th[data-row="0"]'), last = box('#grid th[data-row="7"]');
    return Math.max(Math.abs(frame.left - col.left), Math.abs(frame.right - col.right), Math.abs(frame.top - first.top), Math.abs(frame.bottom - last.bottom));
  });
  assert.ok(gap <= 1, `枠が列Cの範囲から ${gap}px ずれている（C1 を隠す結合 B1:C1 の左端まで広がっていないこと）`);
});

test('矢印キーで移動し、Shift+矢印で範囲を広げる', async t => {
  const x = await env.open(t, book);
  await x.press('ArrowDown');
  await x.press('ArrowRight');
  assert.equal(await x.nameBox(), 'B2');
  await x.press('Shift+ArrowRight');
  await x.press('Shift+ArrowDown');
  assert.equal(await x.nameBox(), 'B2:C3');
  assert.equal(await x.activeCell(), 'B2');
  await x.press('ArrowLeft');
  assert.equal(await x.nameBox(), 'A2', '範囲選択の後の矢印は、アクティブセルから1つ動いて範囲を解く');
  await x.press('ArrowUp');
  await x.press('ArrowUp');
  await x.press('ArrowLeft');
  assert.equal(await x.nameBox(), 'A1', 'シートの端より外へは動かない');
});

test('Ctrl+矢印は、Excelと同じく値の続く端、または次の値まで飛ぶ', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C4');
  await x.press('Control+ArrowDown');
  assert.equal(await x.nameBox(), 'C6', '値が続く範囲の端');
  await x.press('Control+ArrowDown');
  assert.equal(await x.nameBox(), 'C30', '下に値がなければシートの端');
  await x.press('Control+ArrowUp');
  assert.equal(await x.nameBox(), 'C6', '空のセルからは、次に値があるセル');
  await x.clickCell('C4');
  await x.press('Control+Shift+ArrowDown');
  assert.equal(await x.nameBox(), 'C4:C6', 'Ctrl+Shift+矢印は、そこまで範囲を広げる');
});

test('Home は行の先頭、Ctrl+Home は A1、Ctrl+End は使われている範囲の右下へ移る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('E5');
  await x.press('Home');
  assert.equal(await x.nameBox(), 'A5');
  await x.clickCell('E5');
  await x.press('Control+Home');
  assert.equal(await x.nameBox(), 'A1');
  await x.press('Control+End');
  assert.equal(await x.nameBox(), 'J18');
});

// The column headers and the row numbers stay over the sheet's top and left edges (and the frozen rows and columns
// below and beside them), so a cell scrolled in from above or the left must land clear of them, not under them.
// Whether the active cell is seen is checked at its corners and centre (the fill handle sits on the bottom-right one).
const hidden = x => x.evaluate(() => {
  const td = document.querySelector('#grid td.active-cell'), r = td.getBoundingClientRect();
  for (const [name, px, py] of [['左上', r.left + 4, r.top + 4], ['右上', r.right - 4, r.top + 4], ['左下', r.left + 4, r.bottom - 4], ['中央', (r.left + r.right) / 2, (r.top + r.bottom) / 2]]) {
    const hit = document.elementFromPoint(px, py);
    if (!(td === hit || td.contains(hit))) return td.id + ' の' + name + 'が ' + (hit ? hit.tagName + (hit.className ? '.' + hit.className : '') : '画面の外') + ' の下にある';
  }
  return null;
});
const press = async (x, key, times) => { for (let i = 0; i < times; i++) await x.press(key); };

test('スクロールした表でキーボードで上や左へ移っても、アクティブセルは列見出しや行番号に隠れずに見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  await x.select('A61');
  await x.press('Control+Home');
  assert.equal(await hidden(x), null, '下へ移ってから Ctrl+Home で A1 に戻ったとき（1行目が列見出しの下に入らない）');
  await x.select('A60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で画面の上の外へ出たとき');
  await x.select('A100');
  await press(x, 'PageUp', 3);
  assert.equal(await hidden(x), null, 'PageUp を続けて押したとき');
  await x.select('A5');
  await press(x, 'ArrowRight', 35);
  await press(x, 'ArrowLeft', 34);
  assert.equal(await x.nameBox(), 'B5');
  assert.equal(await hidden(x), null, '右へスクロールした後に←で戻ったとき（左の端が行番号の下に入らない）');
  await x.select('A1');
  await press(x, 'ArrowDown', 60);
  assert.equal(await hidden(x), null, '↓で画面の下の外へ出たとき');
  await x.select('A1');
  await press(x, 'ArrowRight', 30);
  assert.equal(await hidden(x), null, '→で画面の右の外へ出たとき');
});

test('範囲の中を Enter で巡って画面の上の外のセルに移っても、隠れずに見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 12}}));
  await x.select('B1:B90');
  await x.press('Enter');
  assert.equal(await x.nameBox(), 'B1:B90');
  assert.equal(await x.evaluate(() => document.querySelector('#grid td.active-cell').id), 'cell-B2', 'アクティブセルは範囲の2行目へ移る');
  assert.equal(await hidden(x), null);
});

test('枠を固定した表でも、上や左へ移ったセルは固定した行・列の下に隠れない', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 2, cols: 1}}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で、固定した行の下の行へ戻ったとき');
  await press(x, 'ArrowUp', 18);
  assert.equal(await x.nameBox(), 'B2');
  assert.equal(await hidden(x), null, '固定した行に入ったとき');
  await x.select('B5');
  await press(x, 'ArrowRight', 35);
  await press(x, 'ArrowLeft', 35);
  assert.equal(await x.nameBox(), 'B5');
  assert.equal(await hidden(x), null, '←で、固定した列のすぐ右の列へ戻ったとき');
});

test('拡大した表でも、上や左へ移ったセルは見出しに隠れない', async t => {
  const x = await env.open(t, blankBook({view: {zoom: 150}, sheet: {rows: 120, cols: 40, freeze: {rows: 1, cols: 1}}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で戻ったとき');
  await x.select('B5');
  await press(x, 'ArrowRight', 30);
  await press(x, 'ArrowLeft', 28);
  assert.equal(await hidden(x), null, '←で戻ったとき');
});

test('PageDown と PageUp は、画面の高さ分だけ上下に移る', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 300}}));
  await x.press('PageDown');
  const [, row] = /^A(\d+)$/.exec(await x.nameBox());
  assert.ok(+row > 10, `1画面分（10行以上）下がる: A${row}`);
  await x.press('PageUp');
  assert.equal(await x.nameBox(), 'A1');
});

test('Enter で下、Tab で右、Shift を付けると逆へ移る', async t => {
  const x = await env.open(t, book);
  await x.press('Enter');
  assert.equal(await x.nameBox(), 'A2');
  await x.press('Tab');
  assert.equal(await x.nameBox(), 'B2');
  await x.press('Shift+Enter');
  assert.equal(await x.nameBox(), 'B1');
  await x.press('Shift+Tab');
  assert.equal(await x.nameBox(), 'A1');
});

test('範囲を選んだまま Enter / Tab を押すと、範囲の中でアクティブセルが順に巡る', async t => {
  const x = await env.open(t, book);
  await x.select('B4:C5');
  const order = [];
  for (let i = 0; i < 4; i++) { await x.press('Enter'); order.push(await x.activeCell()); }
  assert.deepEqual(order, ['B5', 'C4', 'C5', 'B4'], 'Enter は列の順');
  order.length = 0;
  for (let i = 0; i < 4; i++) { await x.press('Tab'); order.push(await x.activeCell()); }
  assert.deepEqual(order, ['C4', 'B5', 'C5', 'B4'], 'Tab は行の順');
  await x.press('Shift+Enter');
  assert.equal(await x.activeCell(), 'C5', 'Shift+Enter は逆順');
  assert.equal(await x.nameBox(), 'B4:C5', '選択範囲は変わらない');
  assert.equal(await x.formula(), '2', '数式バーはアクティブセルの中身');
});

test('結合セルは全体が選ばれ、矢印で出るときは結合の幅を飛び越す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('A11');
  assert.equal(await x.selectionSize(), '2行 × 3列');
  assert.equal(await x.formula(), '結合セル');
  await x.press('ArrowRight');
  assert.equal(await x.nameBox(), 'D11');
  await x.clickCell('A11');
  await x.press('ArrowDown');
  assert.equal(await x.nameBox(), 'A13');
});

test('結合セルに一部だけかかる範囲を選ぶと、結合全体を含む範囲に広がる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B10');
  await x.clickCell('B11', {shift: true});
  assert.equal(await x.nameBox(), 'A10:C12');
  assert.equal(await x.selectionSize(), '3行 × 3列');
});

test('名前ボックスに番地や範囲を入れて Enter を押すと、そこが選ばれる', async t => {
  const x = await env.open(t, book);
  await x.select('c3:e5');
  assert.equal(await x.nameBox(), 'C3:E5', '小文字でも受け付ける');
  assert.equal(await x.activeCell(), 'C3');
  assert.equal(await x.activeElement(), 'cell-editor', '選んだ後はシートにキー入力が届く');
  await x.select('D8');
  assert.equal(await x.formula(), '消費税');
});

test('名前ボックスの不正な番地とシートの外の番地は、ダイアログで知らせて選択を変えない', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B2');
  await x.select('A0');
  let dialog = await x.dialog();
  assert.equal(dialog?.title, '操作を完了できません');
  assert.match(dialog.body, /セル番地が不正です/);
  await x.answer('ok');
  assert.equal(await x.activeCell(), 'B2');
  await x.select('K1');
  dialog = await x.dialog();
  assert.match(dialog?.body ?? '', /シート範囲外です/);
  await x.answer('ok');
  assert.equal(await x.activeCell(), 'B2');
});

test('Ctrl+A は1回目で使われている範囲、2回目でシート全体を選ぶ', async t => {
  const x = await env.open(t, book);
  await x.press('Control+a');
  assert.equal(await x.nameBox(), 'A1:J18');
  await x.press('Control+a');
  assert.equal(await x.nameBox(), 'A1:J30');
});
