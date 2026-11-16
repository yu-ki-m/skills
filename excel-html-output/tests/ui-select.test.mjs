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
  assert.equal(await x.nameBox(), 'A1:C30', 'Excelと同じく、列にかかる結合セル（A11:C12）の列まで広がる');
  await x.click('#grid th[data-corner]');
  assert.equal(await x.nameBox(), 'A1:J30');
  assert.equal(await x.selectionSize(), '30行 × 10列');
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

test('スクロールした表でキーボードで上や左へ移っても、アクティブセルは見出しに隠れずに見える', {
  todo: '既知の不具合: 移動先を scrollIntoView で表示するとき、画面に固定した列見出し・行見出しの分を考えないため、セルがその下に隠れる（app.js の select）',
}, async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  const covered = () => x.evaluate(() => {
    const td = document.querySelector('#grid td.active-cell'), r = td.getBoundingClientRect(), hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
    return td === hit || td.contains(hit) ? null : td.id + ' が ' + hit?.tagName + ' の下にある';
  });
  await x.select('A100');
  await x.press('PageUp');
  assert.equal(await covered(), null, 'PageUp で上へ戻ったとき');
  await x.select('AN21');
  for (let i = 0; i < 14; i++) await x.press('ArrowLeft');
  assert.equal(await covered(), null, '矢印で左へ戻ったとき');
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
