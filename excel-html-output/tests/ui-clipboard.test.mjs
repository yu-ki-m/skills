// コピーと貼り付け：Ctrl+C の中身、シート内・TSV・Excel の HTML の貼り付け、切り取り、Markdown、右クリックメニュー、リボン。
// 利用者のクリップボードには触れない（tests/lib/app.mjs の説明を参照）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {Core, fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

/** What core.js (the code measured against desktop Excel) writes for range, from the page's current state. */
async function expectedClipboard(x, range) {
  const {book: state, sheet} = await x.state();
  return Core.clipboard(state, sheet(), Core.range(range));
}

// The clipboard HTML Excel 2021 writes: formats as classes in <style>, values as displayed text.
const EXCEL_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=ProgId content=Excel.Sheet>
<style><!--table {mso-displayed-decimal-separator:"\\.";}
td {padding-top:1px; color:black; font-size:11.0pt; font-family:游ゴシック, monospace; mso-font-charset:128; white-space:nowrap;}
.xl65 {mso-number-format:"\\0022¥\\0022\\#\\,\\#\\#0"; font-weight:700; color:#C00000; background:#FCE4D6; mso-pattern:black none; border-bottom:1.0pt solid windowtext;}
.xl66 {mso-number-format:"\\@"; text-align:center;}
.xl67 {mso-number-format:Percent;}
.xl68 {mso-number-format:"yyyy\\/mm\\/dd";}
--></style></head><body link="#0563C1" vlink="#954F72">
<table border=0 cellpadding=0 cellspacing=0 width=144 style='border-collapse:collapse;width:108pt'>
<!--StartFragment-->
<tr height=18 style='height:13.5pt'><td height=18 class=xl65 width=72 style='height:13.5pt;width:54pt'>¥1,200</td><td class=xl66 width=72 style='width:54pt'>00123</td></tr>
<tr height=18 style='height:13.5pt'><td height=18 class=xl67 style='height:13.5pt'>12.35%</td><td class=xl68>2025/12/09</td></tr>
<!--EndFragment-->
</table></body></html>`;

test('Ctrl+C で書き出す中身は、Excelで実測済みの生成処理（core.js）の出力と完全に一致する', async t => {
  const x = await env.open(t, book);
  await x.select('C4:E6');
  const copied = await x.copy();
  const expected = await expectedClipboard(x, 'C4:E6');
  assert.equal(copied.prevented, true, 'ページが中身を用意する');
  assert.equal(copied.html, expected.html);
  assert.equal(copied.text, expected.text);
  assert.equal(copied.text, '3\t¥12,000\t¥36,000\r\n2\t¥5,000\t¥10,000\r\n1\t¥30,000\t¥30,000', '表示どおりのタブ区切り');
  assert.match(copied.html, /x:fmla="=A1\*B1"/, '範囲内を参照する数式は、範囲の左上を基準にした数式で渡す');
  assert.match(copied.html, /x:num="36000"/);
  assert.equal(await x.evaluate(() => document.getElementById('selection').classList.contains('copying')), true, 'コピーした範囲を点線で示す');
  assert.match(await x.status(), /選択範囲をコピーしました/);
  await x.press('Escape');
  assert.equal(await x.evaluate(() => document.getElementById('selection').classList.contains('copying')), false, 'Esc で点線を消す');
});

test('結合セルを一部だけ選んでコピーしても、結合全体の中身を渡す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('A10');
  await x.clickCell('B11', {shift: true});
  const copied = await x.copy();
  const expected = await expectedClipboard(x, 'A10:C12');
  assert.equal(copied.html, expected.html);
  assert.match(copied.html, /colspan="3" rowspan="2"/);
});

test('Ctrl+C・Ctrl+V・Ctrl+X は、ページが止めずにブラウザーのコピー・貼り付け・切り取りに任せる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  for (const key of ['c', 'v', 'x']) await x.press('Control+' + key);
  assert.deepEqual(await x.evaluate(() => window.__clipboardKeys), [
    {key: 'c', leftToBrowser: true}, {key: 'v', leftToBrowser: true}, {key: 'x', leftToBrowser: true},
  ]);
});

test('このシートからコピーした範囲を貼ると、値・相対参照の数式・書式・結合が移る', async t => {
  const x = await env.open(t, book);
  await x.select('E4:E6');
  const copied = await x.copy();
  await x.clickCell('G4');
  const pasted = await x.paste(copied);
  assert.equal(pasted.prevented, true);
  let cells = await x.cells();
  assert.deepEqual([cells.G4.f, cells.G6.f], ['=E4*F4', '=E6*F6'], 'Excelと同じく、参照が貼った位置に合わせてずれる');
  assert.equal(cells.G4.s.numFmt, '¥#,##0');
  assert.equal(await x.status(), '値・数式・書式・結合を貼り付けました');
  await x.clickCell('A11');
  const merged = await x.copy();
  await x.clickCell('F14');
  await x.paste(merged);
  const f14 = await x.cell('F14');
  assert.deepEqual([f14.text, f14.colSpan, f14.rowSpan], ['結合セル', 3, 2]);
  cells = await x.cells();
  assert.equal(cells.F14.s.align, 'center');
});

test('結合セルの一部だけに重なる貼り付けは、理由を知らせて行わない', async t => {
  const x = await env.open(t, book);
  await x.select('B4:B5');
  const copied = await x.copy();
  await x.clickCell('A11');
  await x.paste(copied);
  assert.match((await x.dialog())?.body ?? '', /結合セルの一部に重なります/);
  await x.answer('ok');
  assert.equal(await x.text('A11'), '結合セル');
});

test('タブ区切りの文字を貼ると、型を推定して入り、引用符で囲んだタブや改行はセルの中に残る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('G20');
  await x.paste({text: '品番\t数量\r\n00777\t5\r\nTRUE\t1/2\r\n"a\tb"\t"1行目\n2行目"'});
  const cells = await x.cells();
  assert.deepEqual(['G20', 'H20', 'G21', 'H21', 'G22', 'H22', 'G23', 'H23'].map(a => [cells[a].t, cells[a].v]), [
    ['s', '品番'], ['s', '数量'], ['s', '00777'], ['n', 5], ['b', true], ['s', '1/2'], ['s', 'a\tb'], ['s', '1行目\n2行目'],
  ]);
  assert.match(await x.status(), /貼り付けました。Excelなど外部からの値は/);
});

test('ExcelでコピーしたHTMLを貼ると、表示形式・太字・色・塗り・罫線・配置と、表示されていた値が入る', async t => {
  const x = await env.open(t, book);
  await x.clickCell('G20');
  await x.paste({html: EXCEL_HTML, text: '¥1,200\t00123\r\n12.35%\t2025/12/09'});
  const cells = await x.cells();
  assert.deepEqual([cells.G20.t, cells.G20.v], ['n', 1200]);
  const s = cells.G20.s;
  assert.deepEqual([s.numFmt, s.bold, s.color, s.fill, s.borderBottom, s.font], ['¥#,##0', true, '#c00000', '#fce4d6', '2px solid #000000', 'Yu Gothic']);
  assert.deepEqual([cells.H20.t, cells.H20.v, cells.H20.s.numFmt, cells.H20.s.align], ['s', '00123', '@', 'center']);
  assert.deepEqual([cells.G21.t, cells.G21.v, cells.G21.s.numFmt], ['n', 0.1235, '0.00%']);
  assert.deepEqual([cells.H21.t, cells.H21.v, cells.H21.s.numFmt], ['n', 46000, 'yyyy/mm/dd']);
  assert.equal(await x.text('G20'), '¥1,200');
});

test('1つのセルの値を範囲に貼ると、範囲のすべてのセルに入る', async t => {
  const x = await env.open(t, book);
  await x.select('G4:H5');
  await x.paste({text: '7'});
  const cells = await x.cells();
  assert.deepEqual(['G4', 'H4', 'G5', 'H5'].map(a => cells[a].v), [7, 7, 7, 7]);
});

test('切り取り（Ctrl+X）は、中身をクリップボードに渡してから値を消す', async t => {
  const x = await env.open(t, book);
  await x.select('B4:C4');
  const expected = await expectedClipboard(x, 'B4:C4');
  const cut = await x.cut();
  assert.equal(cut.html, expected.html);
  assert.equal(await x.text('B4'), '');
  assert.equal(await x.text('C4'), '');
  assert.match(await x.status(), /切り取りました/);
});

test('編集中のコピーと貼り付けは、ページが横取りせず、入力欄の文字の操作になる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.press('F2');
  const copied = await x.copy();
  assert.deepEqual([copied.prevented, copied.html], [false, '']);
  const pasted = await x.paste({text: 'x\ty'});
  assert.equal(pasted.prevented, false);
  await x.press('Escape');
  assert.equal(await x.text('C4'), '3', '隣のセルは変わらない');
});

test('「Markdown」は、表示されている文字の Markdown の表をコピーし、値のない範囲では知らせるだけにする', async t => {
  const x = await env.open(t, book);
  await x.select('A3:E6');
  await x.action('copy-markdown');
  const {book: state, sheet} = await x.state();
  const expected = Core.markdown(state, sheet(), Core.range('A3:E6')).text;
  await x.waitFor(() => window.__testClipboard.writes > 0, {message: 'クリップボードへの書き込み'});
  assert.deepEqual(await x.clipboard(), [{type: 'text/plain', size: new Blob([expected]).size, text: expected}]);
  assert.equal(expected.split('\n').length, 5, '見出し・区切り・3行');
  assert.match(expected.split('\n')[2], /^\| 00123 +\| 保守作業 \| +3 \| +¥12,000 \| +¥36,000 \|$/, '表示どおりの文字で、数値は右寄せ');
  assert.match(await x.status(), /Markdown の表としてコピーしました（4行 × 5列、1行目が見出し）/);
  await x.select('G25:H26');
  await x.action('copy-markdown');
  assert.equal(await x.status(), '選択範囲に値がありません');
});

test('右クリックメニューは、クリックしたセルを選んで開き、選んだ項目を実行して閉じる', async t => {
  const x = await env.open(t, book);
  await x.select('B4:C5');
  await x.rightClickCell('C5');
  assert.equal(await x.nameBox(), 'B4:C5', '選択範囲の中を右クリックしても、選択は変わらない');
  assert.equal(await x.evaluate(() => document.getElementById('context-menu').hidden), false);
  await x.clickCell('H2');
  assert.equal(await x.evaluate(() => document.getElementById('context-menu').hidden), true, 'ほかの場所を押すと閉じる');
  await x.contextAction('D5', 'clear');
  assert.equal(await x.nameBox(), 'D5', '選択範囲の外を右クリックすると、そのセルを選ぶ');
  assert.equal(await x.text('D5'), '');
  assert.equal(await x.evaluate(() => document.getElementById('context-menu').hidden), true);
  await x.rightClickCell('B4');
  await x.press('Escape');
  assert.equal(await x.evaluate(() => document.getElementById('context-menu').hidden), true, 'Esc で閉じる');
});

test('リボンの「コピー」と「貼り付け」は、クリップボードの読み書きの機能を通して同じ内容を受け渡す', async t => {
  const x = await env.open(t, book);
  await x.select('C4:E5');
  await x.action('copy');
  await x.waitFor(() => window.__testClipboard.writes > 0, {message: 'クリップボードへの書き込み'});
  const expected = await expectedClipboard(x, 'C4:E5');
  const stored = await x.clipboard();
  assert.deepEqual(stored.map(e => e.type).sort(), ['text/html', 'text/plain']);
  assert.equal(stored.find(e => e.type === 'text/html').text, expected.html);
  assert.equal(stored.find(e => e.type === 'text/plain').text, expected.text);
  await x.clickCell('G10');
  await x.action('paste');
  await x.waitFor(() => document.getElementById('cell-G10').textContent === '3', {message: '貼り付けの反映'});
  const cells = await x.cells();
  assert.deepEqual([cells.G10.v, cells.I11.f], [3, '=G11*H11']);
});

test('ブラウザーがクリップボードの読み書きを許さないとき、黙って失敗せずに理由を知らせる', async t => {
  const x = await env.open(t, book);
  await x.evaluate(() => { window.__testClipboard.denyRead(true); window.__testClipboard.denyWrite(true); });
  await x.clickCell('G10');
  await x.action('paste');
  await x.waitFor(() => /読み取り許可が必要です/.test(document.getElementById('status').textContent), {message: '貼り付けできない理由の表示'});
  await x.select('C4:E5');
  await x.action('copy');
  await x.waitDialog('message-dialog');
  assert.match((await x.dialog()).body, /書式付きコピーを許可しませんでした/);
  assert.deepEqual(await x.evaluate(() => window.__testClipboard.legacy), ['copy'], '古いコピーの方法も試してから知らせる');
  await x.answer('ok');
});

test('画面の上部に「Excelへコピー」のような専用のボタンを置かない', async t => {
  const x = await env.open(t, book);
  const labels = await x.evaluate(() => [...document.querySelectorAll('header button, nav button, .ribbon button')].map(b => (b.textContent + ' ' + (b.title || '')).trim()));
  assert.deepEqual(labels.filter(l => /Excel/.test(l)), []);
});
