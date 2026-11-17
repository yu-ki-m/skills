// 表の構造と履歴：行・列の挿入と削除、元に戻す・やり直し、オートフィル、Σ、検索、シート、ズーム、ブック名。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {blankBook, fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

const buttonDisabled = (x, action) => x.evaluate(a => document.querySelector(`.ribbon [data-action="${a}"]`).disabled, action);

test('行の挿入で下の行がずれ、数式の参照・結合・合計が追従する', async t => {
  const x = await env.open(t, book);
  await x.clickCell('A5');
  await x.action('insert-row');
  const sheet = (await x.state()).sheet('main');
  assert.equal(sheet.rows, 31);
  assert.deepEqual(sheet.cells.A5, {t: 's', v: '', s: {numFmt: '@'}}, '挿入した行は空で、上の行の書式を受け継ぐ');
  assert.equal(sheet.cells.A6.v, '00124');
  assert.equal(sheet.cells.E6.f, '=C6*D6');
  assert.equal(sheet.cells.E8.f, '=SUM(E4:E7)', '範囲の終わりが下へ広がる');
  assert.deepEqual(sheet.merges, [{r1: 11, c1: 0, r2: 12, c2: 2}]);
  assert.equal(sheet.rowHeights[0], 30, '上の行の高さは変わらない');
  assert.equal(await x.text('E10'), '¥83,600', '合計の値は変わらない');
});

test('行の削除は確認してから行い、消えたセルを参照する数式は #REF! になる', async t => {
  const x = await env.open(t, book);
  await x.enter('F4', '=C5*2');
  await x.clickCell('A5');
  await x.action('delete-row');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, '選択した行を削除');
  await x.answer('cancel');
  assert.equal(await x.text('A5'), '00124', 'キャンセルでは何もしない');
  await x.action('delete-row');
  await x.answer('ok');
  const sheet = (await x.state()).sheet('main');
  assert.equal(sheet.rows, 29);
  assert.equal(sheet.cells.A5.v, '00230', '下の行が上がる');
  assert.equal(sheet.cells.E6.f, '=SUM(E4:E5)');
  assert.equal(await x.text('E8'), '¥72,600', '残った行で合計を計算し直す');
  assert.equal(sheet.cells.F4.f, '=#REF!*2');
  assert.equal(await x.text('F4'), '#REF!');
});

test('列の挿入と削除で右の列がずれ、数式と列幅が追従する', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.action('insert-col');
  let sheet = (await x.state()).sheet('main');
  assert.equal(sheet.cols, 11);
  assert.equal(sheet.cells.C4.v, '保守作業');
  assert.equal(sheet.cells.F4.f, '=D4*E4');
  assert.equal(sheet.colWidths[1], 96, '挿入した列は左のA列の幅を受け継ぐ');
  assert.equal(sheet.colWidths[2], 120, 'B列の幅は C列へ移る');
  assert.equal(await x.text('F9'), '¥83,600');
  await x.clickCell('B4');
  await x.action('delete-col');
  await x.answer('ok');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.cols, 10);
  assert.equal(sheet.cells.B4.v, '保守作業');
  assert.equal(sheet.cells.E4.f, '=C4*D4');
});

const content = (sheet, addr) => sheet.cells[addr]?.f ?? sheet.cells[addr]?.v ?? '';
const openContextMenu = async (x, selector) => {
  await x.click(selector, {button: 'right'});
  await x.waitFor(() => !document.getElementById('context-menu').hidden, {message: '右クリックメニューの表示'});
};

test('列見出し・行見出しを右クリックして挿入すると、選んでいたセルではなく、右クリックした列・行の位置に入る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('E5');
  await openContextMenu(x, '#grid thead th[data-col="2"]');
  assert.equal(await x.nameBox(), 'C1:C30', 'Excelと同じく、右クリックした列を選ぶ');
  await x.click('#context-menu [data-action="insert-col"]');
  let sheet = (await x.state()).sheet('main');
  assert.deepEqual(['C4', 'D4', 'E4', 'F4'].map(a => content(sheet, a)), ['', 3, 12000, '=D4*E4'], 'C列に空の列が入り、元のC列から右がずれる');
  await x.clickCell('E9');
  await openContextMenu(x, '#grid th[data-row="4"]');
  assert.equal(await x.nameBox(), 'A5:K5', '右クリックした行を選ぶ');
  await x.click('#context-menu [data-action="insert-row"]');
  sheet = (await x.state()).sheet('main');
  assert.deepEqual(['A5', 'A6'].map(a => content(sheet, a)), ['', '00124'], '5行目に空の行が入り、元の5行目から下がずれる');
  await x.click('#grid thead th[data-col="1"]');
  await x.click('#grid thead th[data-col="3"]', {shift: true});
  await openContextMenu(x, '#grid thead th[data-col="2"]');
  assert.equal(await x.nameBox(), 'B1:D31', '選んでいた列の中で右クリックしたときは、選択を変えない');
  await x.click('#context-menu [data-action="insert-col"]');
  assert.equal((await x.state()).sheet('main').cols, 14, '選んでいた3列分が入る');
});

test('列・行を挿入・削除しても、選んでいた範囲は同じ番地のまま残り、結合した題名にかかっても広がらない', async t => {
  const x = await env.open(t, blankBook({sheet: {cells: {A1: {t: 's', v: '題名'}}, merges: [{r1: 0, c1: 0, r2: 0, c2: 4}]}}));
  await x.click('#grid thead th[data-col="2"]');
  await x.action('insert-col');
  assert.equal(await x.nameBox(), 'C1:C30', 'Excelと同じく、挿入した列を選んだまま');
  await x.action('insert-col');
  let sheet = (await x.state()).sheet('main');
  assert.equal(sheet.cols, 14, '続けて押すと、同じ位置にもう1列だけ入る');
  assert.deepEqual(sheet.merges, [{r1: 0, c1: 0, r2: 0, c2: 6}], '題名の結合は、入った列の分だけ広がる');
  await x.action('delete-col');
  await x.answer('ok');
  assert.equal(await x.nameBox(), 'C1:C30', '削除した後も同じ列を選んだまま');
  await x.select('B4:C5');
  await x.action('insert-row');
  assert.equal(await x.nameBox(), 'B4:C5');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.rows, 32);
});

test('全列にかかる題名の結合がある表でも、見出しで選んだ1列の削除は1列ずつ行い、すべての列の削除として止めない', async t => {
  const x = await env.open(t, blankBook({sheet: {cells: {A1: {t: 's', v: '題名'}}, merges: [{r1: 0, c1: 0, r2: 0, c2: 11}]}}));
  await x.click('#grid thead th[data-col="2"]');
  for (const cols of [11, 10]) {
    await x.action('delete-col');
    assert.equal((await x.dialog())?.title, '選択した列を削除', '「すべての行または列は削除できません」ではなく、削除の確認になる');
    await x.answer('ok');
    assert.equal((await x.state()).sheet('main').cols, cols, '1列だけ消える');
    assert.equal(await x.nameBox(), 'C1:C30', '削除した後も列Cを選んだまま');
  }
  assert.deepEqual((await x.state()).sheet('main').merges, [{r1: 0, c1: 0, r2: 0, c2: 9}], '題名の結合は、削除した列の分だけ狭まる');
});

test('挿入した列は、左の列の書式と列幅を受け継ぎ、罫線は両隣に同じ線があるときだけ続き、A列に挿入した列は既定のまま（Excelと同じ）', async t => {
  const x = await env.open(t, book);
  await x.click('#grid thead th[data-col="2"]');
  await x.action('insert-col');
  let sheet = (await x.state()).sheet('main');
  assert.equal(sheet.colWidths[2], 120, '左のB列の幅');
  assert.deepEqual(sheet.cells.C3, {t: 's', v: '', s: {bold: true, fill: '#e2efda', borderBottom: '2px solid #217346'}}, '見出しの太字・塗りと、両隣にある下罫線');
  assert.equal(sheet.cells.C4, undefined, '左のB4に書式がないので、書式なし');
  await x.click('#grid thead th[data-col="6"]');
  await x.action('insert-col');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.colWidths[6], 110, '左の「金額」の列の幅');
  assert.deepEqual(sheet.cells.G3, {t: 's', v: '', s: {bold: true, fill: '#e2efda'}}, '表の右端では、右隣にない下罫線は続けない');
  assert.deepEqual(sheet.cells.G4, {t: 's', v: '', s: {numFmt: '¥#,##0'}}, '値や数式は写さず、表示形式だけ');
  assert.equal(await x.text('G4'), '');
  await x.click('#grid thead th[data-col="0"]');
  await x.action('insert-col');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.colWidths[0], null, 'A列に入れた列は既定の幅');
  assert.equal(sheet.cells.A3, undefined, 'A列に入れた列は書式なし');
  assert.equal(sheet.cells.B3.v, '品番');
});

test('挿入した行は、上の行の書式と行の高さを受け継ぎ、1行目に挿入した行は既定のまま（Excelと同じ）', async t => {
  const x = await env.open(t, book);
  await x.click('#grid th[data-row="3"]');
  await x.action('insert-row');
  let sheet = (await x.state()).sheet('main');
  assert.deepEqual(sheet.cells.A4, {t: 's', v: '', s: {bold: true, fill: '#e2efda'}}, '見出しの太字・塗りを受け継ぎ、見出しの下罫線は見出しの行に残る');
  assert.equal(sheet.cells.A5.v, '00123');
  await x.click('#grid th[data-row="1"]');
  await x.action('insert-row');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.rowHeights[1], 30, '上の1行目の高さ');
  assert.deepEqual(sheet.cells.A2, {t: 's', v: '', s: {bold: true, size: 14}}, '題名の文字は写さず、書式だけ');
  await x.click('#grid th[data-row="0"]');
  await x.action('insert-row');
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.rowHeights[0], null, '1行目に入れた行は既定の高さ');
  assert.equal(sheet.cells.A1, undefined, '1行目に入れた行は書式なし');
  assert.equal(sheet.cells.A2.v, 'テスト用の明細表（右の空いたセルへはみ出す長い題名）');
});

test('結合セルの右や下に挿入した列・行は、結合セルの塗りや太字を受け継ぐが、罫線は受け継がない（Excelと同じ）', async t => {
  const line = '1px solid #000000';
  const merged = {bold: true, fill: '#fff2cc', align: 'center', borderTop: line, borderRight: line, borderBottom: line, borderLeft: line};
  const sheetWithMerge = blankBook({sheet: {cells: {B2: {t: 's', v: '結合', s: merged}}, merges: [{r1: 1, c1: 1, r2: 2, c2: 2}]}});
  const x = await env.open(t, sheetWithMerge);
  await x.click('#grid thead th[data-col="3"]');
  await x.action('insert-col');
  await x.click('#grid th[data-row="3"]');
  await x.action('insert-row');
  const sheet = (await x.state()).sheet('main');
  const inherited = {bold: true, fill: '#fff2cc', align: 'center'};
  assert.deepEqual(['D2', 'D3', 'B4', 'C4', 'D4'].map(a => sheet.cells[a]?.s), Array(5).fill(inherited));
  assert.deepEqual(sheet.merges, [{r1: 1, c1: 1, r2: 2, c2: 2}], '結合の外に入れた列・行は、結合を変えない');
  // Inside the merged area the merged cell grows, and its hidden cells stay empty.
  const y = await env.open(t, sheetWithMerge);
  await y.click('#grid thead th[data-col="2"]');
  await y.action('insert-col');
  const grown = (await y.state()).sheet('main');
  assert.deepEqual(grown.merges, [{r1: 1, c1: 1, r2: 2, c2: 3}]);
  assert.deepEqual(Object.keys(grown.cells), ['B2'], '結合の中に入った列には、隠れたセルを作らない');
  // Between a merged cell with an outline and a bordered row below it: the new row goes on with the outline only.
  const box = {borderTop: line, borderRight: line, borderBottom: line, borderLeft: line};
  const z = await env.open(t, blankBook({sheet: {cells: {B2: {t: 's', v: '結合', s: box}, B4: {t: 's', v: '', s: box}, C4: {t: 's', v: '', s: box}}, merges: [{r1: 1, c1: 1, r2: 2, c2: 2}]}}));
  await z.click('#grid th[data-row="3"]');
  await z.action('insert-row');
  const below = (await z.state()).sheet('main');
  assert.deepEqual([below.cells.B4?.s, below.cells.C4?.s], [{borderLeft: line}, {borderRight: line}], '結合セルの外枠の縦線だけが続き、結合セルの中にあたる境目の線は入らない');
});

test('すべての行の削除と、シートの上限（100列）を超える列の挿入は、理由を知らせて行わない', async t => {
  const x = await env.open(t, book);
  await x.click('#grid th[data-corner]');
  await x.action('delete-row');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, '操作を完了できません');
  assert.match(dialog.body, /すべての行または列は削除できません/);
  await x.answer('ok');
  assert.equal((await x.state()).sheet('main').rows, 30);
  const wide = await env.open(t, blankBook({sheet: {cols: 100}}));
  await wide.clickCell('B2');
  await wide.action('insert-col');
  assert.match((await wide.dialog())?.body ?? '', /シートの上限を超えます/);
  await wide.answer('ok');
  assert.equal((await wide.state()).sheet('main').cols, 100);
});

test('元に戻す・やり直しで、入力・書式・結合・行の挿入を1つずつ行き来できる', async t => {
  const x = await env.open(t, book);
  assert.equal(await buttonDisabled(x, 'undo'), true, '最初は「元に戻す」を押せない');
  await x.enter('B4', '変更1');
  await x.clickCell('B4');
  await x.press('Control+b');
  await x.select('F12:G13');
  await x.action('merge');
  await x.clickCell('A5');
  await x.action('insert-row');
  assert.equal(await buttonDisabled(x, 'undo'), false);
  // Each step must take back exactly one operation: the four operations, newest first.
  const snapshot = async () => {
    const s = (await x.state()).sheet('main');
    return {rows: s.rows, merges: s.merges.length, bold: s.cells.B4.s?.bold === true, b4: s.cells.B4.v};
  };
  const steps = [
    {rows: 31, merges: 2, bold: true, b4: '変更1'},
    {rows: 30, merges: 2, bold: true, b4: '変更1'},
    {rows: 30, merges: 1, bold: true, b4: '変更1'},
    {rows: 30, merges: 1, bold: false, b4: '変更1'},
    {rows: 30, merges: 1, bold: false, b4: '保守作業'},
  ];
  assert.deepEqual(await snapshot(), steps[0]);
  await x.clickCell('H1');
  for (const [i, expected] of steps.slice(1).entries()) {
    await x.press('Control+z');
    assert.deepEqual(await snapshot(), expected, `${i + 1}回目の Ctrl+Z は、${['行の挿入', '結合', '太字', '入力'][i]}だけを戻す`);
  }
  assert.equal(await x.status(), '元に戻しました');
  assert.equal(await buttonDisabled(x, 'undo'), true);
  assert.equal(await buttonDisabled(x, 'redo'), false);
  await x.press('Control+y');
  assert.deepEqual(await snapshot(), steps[3], 'Ctrl+Y は入力だけをやり直す');
  await x.press('Control+Shift+Z');
  assert.deepEqual(await snapshot(), steps[2], 'Ctrl+Shift+Z でもやり直す（太字）');
  assert.equal(await x.status(), 'やり直しました');
  await x.action('undo');
  assert.deepEqual(await snapshot(), steps[3], 'リボンの「元に戻す」');
  await x.action('redo');
  assert.deepEqual(await snapshot(), steps[2], 'リボンの「やり直し」');
});

test('元に戻すと選択範囲も戻り、戻した後に新しく編集すると「やり直し」は消える', async t => {
  const x = await env.open(t, book);
  await x.select('C4:D5');
  await x.press('Delete');
  await x.clickCell('H1');
  await x.press('Control+z');
  assert.equal(await x.nameBox(), 'C4:D5', '操作した時の選択に戻る');
  assert.equal(await x.text('C4'), '3');
  await x.enter('H2', '新しい編集');
  assert.equal(await buttonDisabled(x, 'redo'), true);
});

test('中身の変わらない確定は、元に戻す履歴に残らない', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.press('F2');
  await x.press('Enter');
  await x.clickCell('B4');
  await x.type('保守作業');
  await x.press('Enter');
  assert.equal(await buttonDisabled(x, 'undo'), true);
});

test('フィルハンドルのドラッグで、相対参照の数式・連続データ・書式をコピーする', async t => {
  const x = await env.open(t, book);
  await x.enter('G4', '=C4*2');
  await x.clickCell('G4');
  await x.page.drag(await x.point('#fill-handle'), await x.point('#cell-G6'));
  let cells = await x.cells();
  assert.deepEqual([cells.G5.f, cells.G6.f], ['=C5*2', '=C6*2'], '数式の参照が1行ずつずれる');
  assert.deepEqual([await x.text('G5'), await x.text('G6')], ['4', '2']);
  await x.enter('H4', '1');
  await x.enter('H5', '2');
  await x.select('H4:H5');
  await x.page.drag(await x.point('#fill-handle'), await x.point('#cell-H8'));
  assert.deepEqual([await x.text('H6'), await x.text('H7'), await x.text('H8')], ['3', '4', '5'], '数値2つから連続データ');
  await x.enter('F20', '見出し');
  await x.clickCell('F20');
  await x.press('Control+b');
  await x.page.drag(await x.point('#fill-handle'), await x.point('#cell-H20'));
  cells = await x.cells();
  assert.deepEqual([cells.G20.v, cells.H20.v, cells.H20.s.bold], ['見出し', '見出し', true], '右へも値と書式をコピー');
});

test('結合セルにかかるオートフィルは、理由を知らせて行わない', async t => {
  const x = await env.open(t, book);
  await x.enter('A10', 'x');
  await x.clickCell('A10');
  await x.page.drag(await x.point('#fill-handle'), await x.point('#cell-A11'));
  const dialog = await x.dialog();
  assert.match(dialog?.body ?? '', /結合セルを含むオートフィル/);
  await x.answer('ok');
  assert.equal(await x.text('A11'), '結合セル');
});

test('Ctrl+D は上のセルを下へ、Ctrl+R は左のセルを右へコピーする', async t => {
  const x = await env.open(t, book);
  await x.enter('G4', '=C4*2');
  await x.select('G4:G6');
  await x.press('Control+d');
  await x.enter('F20', 'abc');
  await x.select('F20:H20');
  await x.press('Control+r');
  const cells = await x.cells();
  assert.deepEqual([cells.G5.f, cells.G6.f], ['=C5*2', '=C6*2']);
  assert.deepEqual([cells.G20.v, cells.H20.v], ['abc', 'abc']);
});

test('Σ は、上に続く数値の範囲を SUM する数式を入れる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C7');
  await x.action('sum');
  assert.equal((await x.cells()).C7.f, '=SUM(C4:C6)');
  assert.equal(await x.text('C7'), '6');
  await x.clickCell('A2');
  await x.action('sum');
  assert.equal(await x.status(), '上に連続した数値がありません');
});

test('検索（Ctrl+F）は、表示されている値と数式の中から次に一致するセルを選ぶ', async t => {
  const x = await env.open(t, book);
  await x.press('Control+f');
  assert.equal((await x.dialog())?.title, '検索');
  await x.answer('ok', '監視');
  assert.equal(await x.nameBox(), 'B5');
  assert.match(await x.status(), /「監視」を B5 に見つけました/);
  await x.action('find');
  await x.answer('ok', 'rounddown');
  assert.equal(await x.nameBox(), 'E8', '数式の中の関数名も、大文字小文字を区別せずに探す');
  await x.action('find');
  await x.answer('ok', '¥36,000');
  assert.equal(await x.nameBox(), 'E4', '表示形式を付けた表示のままでも探せる');
  await x.action('find');
  await x.answer('ok', '存在しない文字');
  assert.equal(await x.status(), '一致するセルがありません');
  assert.equal(await x.nameBox(), 'E4');
});

test('シートタブで切り替え、ダブルクリックで名前を変え、＋で追加し、追加は元に戻せる', async t => {
  const x = await env.open(t, book);
  await x.click('#sheet-tabs [data-sheet="sub"]');
  assert.equal(await x.text('A1'), '予備シートの値');
  assert.equal(await x.nameBox(), 'A1');
  assert.equal(await x.evaluate(() => document.querySelector('#sheet-tabs [data-sheet="sub"]').getAttribute('aria-selected')), 'true');
  await x.page.dblclick(await x.point('#sheet-tabs [data-sheet="sub"]'));
  assert.equal((await x.dialog())?.input, '予備');
  await x.answer('ok', '集計');
  assert.equal(await x.textOf('#sheet-tabs [data-sheet="sub"]'), '集計');
  await x.click('#add-sheet');
  let state = await x.state();
  assert.equal(state.book.sheets.length, 3);
  assert.equal(state.book.sheets[2].name, 'Sheet1');
  assert.equal(state.book.activeSheet, state.book.sheets[2].id, '追加したシートが開く');
  await x.click('#sheet-tabs [data-sheet="main"]');
  assert.equal(await x.text('E9'), '¥83,600', '元のシートの内容はそのまま');
  await x.clickCell('A1');
  await x.press('Control+z');
  state = await x.state();
  assert.equal(state.book.sheets.length, 2, 'シートの追加を元に戻す');
  assert.deepEqual(state.book.sheets.map(s => s.name), ['明細', '集計']);
});

test('シート名に使えない文字や重複した名前は、理由を知らせて変えない', async t => {
  const x = await env.open(t, book);
  for (const name of ['a/b', '明細', '']) {
    await x.page.dblclick(await x.point('#sheet-tabs [data-sheet="sub"]'));
    await x.answer('ok', name);
    const dialog = await x.dialog();
    assert.match(dialog?.body ?? '', /シート名/, `「${name}」は受け付けない`);
    await x.answer('ok');
  }
  assert.deepEqual((await x.state()).book.sheets.map(s => s.name), ['明細', '予備']);
});

test('ズームのボタンとスライダーで50%〜200%に拡大・縮小し、選択枠は拡大後もセルに重なる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C4');
  await x.click('#zoom-in');
  assert.equal(await x.textOf('#zoom-label'), '110%');
  assert.equal(await x.evaluate(() => document.getElementById('sheet-surface').style.zoom), '1.1');
  assert.ok(await x.selectionFrameGap('C4', 'C4') <= 2, '110% での選択枠');
  await x.setControl('zoom', '200');
  await x.click('#zoom-in');
  assert.equal(await x.textOf('#zoom-label'), '200%', '200% より大きくならない');
  await x.setControl('zoom', '50');
  await x.click('#zoom-out');
  assert.equal(await x.textOf('#zoom-label'), '50%', '50% より小さくならない');
  assert.ok(await x.selectionFrameGap('C4', 'C4') <= 2, '50% での選択枠');
  await x.clickCell('D5');
  assert.equal(await x.nameBox(), 'D5', '縮小した表示でも、クリックしたセルが選ばれる');
  assert.equal((await x.state()).book.view.zoom, 50);
});

test('リボンと右クリックメニューのボタンは、どれを押してもエラーにならない（動かないボタンを作らない）', async t => {
  const x = await env.open(t, book);
  const ribbon = await x.evaluate(() => [...new Set([...document.querySelectorAll('.ribbon [data-action]')].map(b => b.dataset.action))]);
  const menu = await x.evaluate(() => [...document.querySelectorAll('#context-menu [data-action]')].map(b => b.dataset.action));
  assert.ok(ribbon.length >= 10 && menu.length >= 5, `ボタンが見つかっている: リボン ${ribbon.length}、メニュー ${menu.length}`);
  const settle = async () => {
    await x.evaluate(() => new Promise(resolve => setTimeout(resolve, 50)));
    for (let dialog = await x.dialog(); dialog; dialog = await x.dialog()) {
      if (dialog.id === 'message-dialog') await x.answer('cancel');
      else await x.evaluate(id => document.getElementById(id).close(), dialog.id);
    }
  };
  for (const action of ribbon) { await x.clickCell('B4'); await x.action(action); await settle(); }
  for (const action of menu) { await x.contextAction('B4', action); await settle(); }
  await x.click('[data-tab="help"]');
  await x.click('#help-open');
  assert.equal((await x.dialog())?.id, 'help-dialog', '「操作・貼り付け条件」で使い方を開く');
  await settle();
  // Errors thrown by any button are collected by the page and fail this test when the tab closes.
});

test('ブック名を変えると、ブラウザーのタブの題名と状態テキストの題名が変わる', async t => {
  const x = await env.open(t, book);
  await x.click('#book-title');
  await x.evaluate(() => document.getElementById('book-title').select());
  await x.type('新しいブック名');
  await x.press('Enter');
  assert.equal(await x.evaluate(() => document.title), '新しいブック名');
  assert.equal((await x.state()).book.title, '新しいブック名');
});
