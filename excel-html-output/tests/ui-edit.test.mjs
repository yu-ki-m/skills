// 入力と編集：型の推定、F2・ダブルクリック・Esc、数式と再計算、セル内の改行、日本語入力、数式バー、クリア、表示形式。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

test('文字を打つとセルの中身が置き換わり、Enter で確定して下へ移る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.type('点検作業');
  assert.equal(await x.editing(), true, '打ち始めると編集になる');
  assert.equal(await x.formula(), '点検作業', '数式バーにも打った文字が出る');
  await x.press('Enter');
  assert.equal(await x.text('B4'), '点検作業');
  assert.equal(await x.nameBox(), 'B5');
  assert.equal(await x.status(), '準備完了');
  const cells = await x.cells();
  assert.deepEqual([cells.B4.t, cells.B4.v], ['s', '点検作業']);
});

test('編集中の Tab は確定して右へ、Shift+Tab は確定して左へ移る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('F4');
  await x.type('右へ');
  await x.press('Tab');
  assert.equal(await x.nameBox(), 'G4');
  await x.type('左へ');
  await x.press('Shift+Tab');
  assert.equal(await x.nameBox(), 'F4');
  const cells = await x.cells();
  assert.deepEqual([cells.F4.v, cells.G4.v], ['右へ', '左へ']);
});

test('入力した文字は、Excelと同じ規則で数値・文字列・真偽値に分かれる', async t => {
  const x = await env.open(t, book);
  const inputs = {F4: '123', F5: '-1.5', F6: '00123', F7: 'TRUE', F8: "'=A1", F9: '1234567890123456', F10: '1e3'};
  for (const [addr, text] of Object.entries(inputs)) await x.enter(addr, text);
  const cells = await x.cells();
  assert.deepEqual([cells.F4.t, cells.F4.v], ['n', 123]);
  assert.deepEqual([cells.F5.t, cells.F5.v], ['n', -1.5]);
  assert.deepEqual([cells.F6.t, cells.F6.v], ['s', '00123'], '先頭のゼロは文字列として残る');
  assert.deepEqual([cells.F7.t, cells.F7.v], ['b', true]);
  assert.deepEqual([cells.F8.t, cells.F8.v], ['s', '=A1'], "先頭の ' は文字列の印として消える");
  assert.deepEqual([cells.F9.t, cells.F9.v], ['s', '1234567890123456'], '15桁を超える数字は精度を落とさないよう文字列');
  assert.deepEqual([cells.F10.t, cells.F10.v], ['n', 1000]);
  assert.equal((await x.cell('F4')).style.textAlign, 'right', '数値は右寄せ');
  assert.equal((await x.cell('F6')).style.textAlign, 'left', '文字列は左寄せ');
  assert.equal((await x.cell('F7')).style.textAlign, 'center', '真偽値は中央');
  assert.equal(await x.text('F8'), '=A1');
});

test('表示形式「文字列」のセルに数字を打つと、数値にならずに文字列のまま入る', async t => {
  const x = await env.open(t, book);
  await x.enter('A4', '00999');
  await x.enter('A5', '123');
  const cells = await x.cells();
  assert.deepEqual([cells.A4.t, cells.A4.v, cells.A4.s.numFmt], ['s', '00999', '@']);
  assert.deepEqual([cells.A5.t, cells.A5.v], ['s', '123']);
});

test('F2 は今の中身の末尾から編集し、Esc は編集を取り消す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.press('F2');
  assert.equal(await x.editing(), true);
  await x.type('（月額）');
  await x.press('Enter');
  assert.equal(await x.text('B4'), '保守作業（月額）');
  await x.clickCell('B5');
  await x.type('書き換え');
  await x.press('Escape');
  assert.equal(await x.text('B5'), '監視', 'Esc で元の値に戻る');
  assert.equal(await x.editing(), false);
  assert.equal(await x.nameBox(), 'B5', 'Esc では移動しない');
  assert.equal(await x.status(), '編集を取り消しました');
});

test('セルをダブルクリックすると、中身を残したまま編集になる', async t => {
  const x = await env.open(t, book);
  await x.dblclickCell('B5');
  assert.equal(await x.editing(), true);
  await x.type('業務');
  await x.press('Enter');
  assert.equal(await x.text('B5'), '監視業務');
});

test('数式を入れると計算され、参照先を変えると、つながる数式がすべて再計算される', async t => {
  const x = await env.open(t, book);
  await x.enter('C4', '10');
  assert.equal(await x.text('E4'), '¥120,000', '金額 = 数量 × 単価');
  assert.equal(await x.text('E7'), '¥160,000', '小計');
  assert.equal(await x.text('E8'), '¥16,000', '消費税（1円未満切り捨て）');
  assert.equal(await x.text('E9'), '¥176,000', '合計');
  await x.enter('F4', '=E4*2');
  assert.equal(await x.text('F4'), '240000');
  await x.clickCell('F4');
  assert.equal(await x.formula(), '=E4*2');
  const cells = await x.cells();
  assert.deepEqual([cells.F4.t, cells.F4.f], ['f', '=E4*2']);
});

test('計算できない数式は、Excelと同じエラー値を表示する', async t => {
  const x = await env.open(t, book);
  await x.enter('G4', '=1/0');
  await x.enter('G5', '=G5+1');
  await x.enter('G6', '=NOSUCH(1)');
  await x.enter('G7', '=B4+1');
  assert.equal(await x.text('G4'), '#DIV/0!');
  assert.equal(await x.text('G5'), '#CIRC!');
  assert.equal(await x.text('G6'), '#NAME?');
  assert.equal(await x.text('G7'), '#VALUE!');
});

test('範囲を選んで入力すると、Enter のたびに範囲の中の次のセルへ進む', async t => {
  const x = await env.open(t, book);
  await x.select('H4:H5');
  await x.type('1');
  await x.press('Enter');
  await x.type('2');
  await x.press('Enter');
  assert.equal(await x.text('H4'), '1');
  assert.equal(await x.text('H5'), '2');
  assert.equal(await x.activeCell(), 'H4', '範囲の終わりから先頭に戻る');
  assert.equal(await x.nameBox(), 'H4:H5');
});

test('編集中の Alt+Enter はセル内で改行し、確定すると折り返しが有効になって行が高くなる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B6');
  await x.press('F2');
  await x.press('Alt+Enter');
  await x.type('2行目');
  assert.equal(await x.editing(), true, 'Alt+Enter では確定しない');
  await x.press('Enter');
  const {book: state, sheet} = await x.state();
  const s = sheet('main');
  assert.equal(s.cells.B6.v, '設定作業\n2行目');
  assert.equal(s.cells.B6.s.wrap, true);
  assert.ok(s.rowHeights[5] > 22, `行の高さが既定の22pxから広がる: ${s.rowHeights[5]}`);
  assert.equal((await x.cell('B6')).style.whiteSpace, 'pre-wrap');
  assert.equal(state.sheets.length, 2);
});

test('改行を打って行が高くなるときも、行の高さはExcelの上限の546pxを超えない', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B20');
  await x.type('1行目');
  for (let i = 2; i <= 40; i++) {
    await x.press('Alt+Enter');
    await x.type(`${i}行目`);
  }
  await x.press('Enter');
  const s = (await x.state()).sheet('main');
  assert.equal(s.cells.B20.v.split('\n').length, 40);
  assert.equal(s.rowHeights[19], 546, '40行分（およそ700px）でも、546px で止まる');
});

test('打ち始めた直後の Alt+Enter も改行になる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('H8');
  await x.type('a');
  await x.press('Alt+Enter');
  await x.type('b');
  await x.press('Enter');
  assert.equal((await x.cells()).H8.v, 'a\nb');
});

test('編集していないときの Alt+Enter は、セルを変えずに、編集してから押すよう案内する', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B6');
  await x.press('Alt+Enter');
  assert.equal(await x.editing(), false);
  assert.equal(await x.text('B6'), '設定作業');
  assert.equal(await x.nameBox(), 'B6', '移動もしない');
  assert.match(await x.status(), /F2 かダブルクリックで編集を始めてから Alt\+Enter/);
});

test('日本語入力の変換中は Enter を押しても確定・移動せず、変換を確定した後の Enter でセルが確定する', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C8');
  await x.page.compose('にほん');
  assert.equal(await x.editing(), true, '変換中の文字が出た時点で編集になる');
  // An IME takes the Enter that confirms a conversion: the page sees keyCode 229 and no character.
  await x.page.send('Input.dispatchKeyEvent', {type: 'rawKeyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 229});
  await x.page.send('Input.dispatchKeyEvent', {type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 229});
  await x.type('日本');
  assert.equal(await x.nameBox(), 'C8', '変換の確定ではセルは確定しない');
  assert.equal(await x.value('cell-editor'), '日本', '最初の文字も欠けない');
  await x.press('Enter');
  assert.equal(await x.text('C8'), '日本');
  assert.equal(await x.nameBox(), 'C9');
});

test('数式バーで編集して Enter で確定し、Esc で取り消す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B5');
  await x.click('#formula');
  await x.evaluate(() => document.getElementById('formula').select());
  await x.type('監視業務');
  await x.press('Enter');
  assert.equal(await x.text('B5'), '監視業務');
  assert.equal(await x.nameBox(), 'B6', 'Enter で下へ移る');
  await x.click('#formula');
  await x.evaluate(() => document.getElementById('formula').select());
  await x.type('取り消す文字');
  await x.press('Escape');
  assert.equal(await x.text('B6'), '設定作業');
  assert.equal(await x.activeElement(), 'cell-editor', 'Esc の後はシートに戻る');
});

test('数式バーで Alt+Enter を押すと改行が入り、改行のあるセルは数式バーに触れても改行を失わない', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B6');
  await x.click('#formula');
  await x.evaluate(() => { const f = document.getElementById('formula'); f.setSelectionRange(f.value.length, f.value.length); });
  await x.press('Alt+Enter');
  await x.type('次の行');
  await x.press('Enter');
  let cells = await x.cells();
  assert.equal(cells.B6.v, '設定作業\n次の行');
  assert.equal(cells.B6.s.wrap, true);
  // Focusing the formula bar and leaving it again must not change the cell (an <input> would drop the line break).
  await x.clickCell('B6');
  await x.click('#formula');
  await x.clickCell('C6');
  cells = await x.cells();
  assert.equal(cells.B6.v, '設定作業\n次の行');
});

test('数式バーの ✓ は確定して移動せず、× は取り消す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.type('確定する');
  await x.click('#commit-edit');
  assert.equal(await x.text('B4'), '確定する');
  assert.equal(await x.nameBox(), 'B4');
  await x.type('取り消す');
  await x.click('#cancel-edit');
  assert.equal(await x.text('B4'), '確定する');
  assert.equal(await x.editing(), false);
});

test('見出しで選んだ列に結合セルの一部だけが入るときは、Excelと同じく Delete で消さずに知らせる', async t => {
  const x = await env.open(t, book);
  await x.click('#grid th[data-col="1"]');
  await x.press('Delete');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, '操作を完了できません');
  assert.equal(dialog?.body, 'この操作は結合したセルには行えません。', 'Excelと同じ知らせ（A11:C12 の結合が列Bに一部だけかかる）');
  await x.answer('ok');
  assert.equal(await x.text('B4'), '保守作業', '列Bの値は消えない');
  assert.equal(await x.text('A11'), '結合セル', '結合セルの値も消えない');
  await x.click('#grid th[data-col="3"]');
  await x.press('Delete');
  assert.equal(await x.dialog(), null, '結合のかからない列Dは消せる');
  assert.equal(await x.text('D4'), '');
});

test('Delete は範囲の値を消して書式を残し、Backspace は消してから編集を始める', async t => {
  const x = await env.open(t, book);
  await x.select('C4:D5');
  await x.press('Delete');
  assert.equal(await x.text('C4'), '');
  assert.equal(await x.text('D5'), '');
  assert.equal(await x.text('E4'), '¥0', '参照していた数式も再計算される');
  const cells = await x.cells();
  assert.deepEqual([cells.D4.v, cells.D4.s.numFmt], ['', '¥#,##0'], '表示形式は残る');
  await x.clickCell('B4');
  await x.press('Backspace');
  assert.equal(await x.editing(), true);
  await x.type('新しい品名');
  await x.press('Enter');
  assert.equal(await x.text('B4'), '新しい品名');
});

test('表示形式を選ぶと、同じ値が桁区切り・パーセント・日付・文字列で表示される', async t => {
  const x = await env.open(t, book);
  await x.clickCell('D4');
  await x.setControl('number-format', '#,##0');
  assert.equal(await x.text('D4'), '12,000');
  await x.enter('F5', '0.125');
  await x.clickCell('F5');
  await x.setControl('number-format', '0.00%');
  assert.equal(await x.text('F5'), '12.50%');
  await x.enter('F6', '46023');
  await x.clickCell('F6');
  await x.setControl('number-format', 'yyyy/mm/dd');
  assert.equal(await x.text('F6'), '2026/01/01');
  await x.clickCell('F7');
  await x.setControl('number-format', '@');
  await x.clickCell('F7');
  await x.type('0012');
  await x.press('Enter');
  const cells = await x.cells();
  assert.deepEqual([cells.F5.t, cells.F5.v, cells.F5.s.numFmt], ['n', 0.125, '0.00%'], '値は変わらず、表示だけが変わる');
  assert.deepEqual([cells.F7.t, cells.F7.v], ['s', '0012']);
  await x.clickCell('F5');
  assert.equal(await x.value('number-format'), '0.00%', 'リボンの表示形式は、選んでいるセルのものを示す');
});
