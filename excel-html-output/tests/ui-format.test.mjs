// 書式と見た目：文字の書式、色、配置、罫線、結合、列幅と行の高さ、枠線、枠の固定、はみ出し表示、リンク。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {rgb, useBrowser} from './lib/app.mjs';
import {fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

/** Point inside the visible part (the inner 4px) of a column or row resize grip. */
const grip = (x, selector) => x.evaluate(s => {
  const el = document.querySelector(s), r = el.getBoundingClientRect();
  return el.classList.contains('col-resize') ? {x: r.left + 2, y: r.top + r.height / 2} : {x: r.left + r.width / 2, y: r.top + 2};
}, selector);

/** Selects whole rows (or columns) first..last, 0-based, as a person does: click the first header, Shift+click the last. */
const selectRows = async (x, first, last) => {
  await x.click(`#grid th[data-row="${first}"]`);
  await x.click(`#grid th[data-row="${last}"]`, {shift: true});
};
const selectCols = async (x, first, last) => {
  await x.click(`#grid th[data-col="${first}"]`);
  await x.click(`#grid th[data-col="${last}"]`, {shift: true});
};

/** The fixture with other row heights and column widths (rows 4 to 6 are indexes 3 to 5, columns D to F are 3 to 5). */
const resized = (rowHeights, colWidths = book.sheets[0].colWidths) => {
  const b = fixture('book');
  Object.assign(b.sheets[0], {rowHeights, colWidths});
  return b;
};
// Rows 4, 5 and 6 of three different heights, so that "all the same height" differs from where each row started.
const uneven = resized([30, 22, 22, 24, 30, 40]);
// Rows 4 to 6 and columns D to F much larger than their contents, so that fitting them to the contents shrinks each.
const oversized = resized([30, 22, 22, 60, 60, 60], [96, 120, 80, 300, 300, 300, 96, 96, 40, 40]);

test('太字・斜体・下線・取り消し線のボタンは、選んだ範囲に効き、押されている状態を示す', async t => {
  const x = await env.open(t, book);
  await x.select('B4:C4');
  await x.click('[data-style="bold"]');
  assert.equal((await x.cell('B4')).style.fontWeight, '700');
  assert.equal((await x.cell('C4')).style.fontWeight, '700');
  assert.equal(await x.evaluate(() => document.querySelector('[data-style="bold"]').getAttribute('aria-pressed')), 'true');
  await x.click('[data-style="bold"]');
  assert.equal((await x.cell('B4')).style.fontWeight, '400', 'もう一度押すと解除');
  await x.click('[data-style="italic"]');
  await x.click('[data-style="underline"]');
  await x.click('[data-style="strike"]');
  const style = (await x.cell('B4')).style;
  assert.equal(style.fontStyle, 'italic');
  assert.equal(style.textDecorationLine, 'underline line-through');
  const s = (await x.cells()).C4.s;
  assert.deepEqual([s.italic, s.underline, s.strike], [true, true, true]);
});

test('Ctrl+B・Ctrl+I・Ctrl+U で太字・斜体・下線を切り替える', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B5');
  await x.press('Control+b');
  await x.press('Control+i');
  await x.press('Control+u');
  const s = (await x.cells()).B5.s;
  assert.deepEqual([s.bold, s.italic, s.underline], [true, true, true]);
  await x.press('Control+b');
  assert.equal((await x.cells()).B5.s.bold, false);
  assert.equal(await x.editing(), false, 'ショートカットで編集は始まらない');
});

test('フォント・サイズ・文字色・塗りつぶし・塗りなし・横位置・縦位置・折り返しを変える', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.setControl('font', 'Meiryo');
  await x.setControl('font-size', '16');
  await x.setControl('font-color', '#c00000');
  await x.setControl('fill-color', '#fce4d6');
  await x.setControl('align', 'center');
  await x.setControl('valign', 'top');
  await x.click('[data-style="wrap"]');
  let cell = await x.cell('B4');
  assert.match(cell.style.fontFamily, /Meiryo/);
  assert.equal(cell.style.color, rgb('#c00000'));
  assert.equal(cell.style.backgroundColor, rgb('#fce4d6'));
  assert.equal(cell.style.textAlign, 'center');
  assert.equal(cell.style.whiteSpace, 'pre-wrap');
  const s = (await x.cells()).B4.s;
  assert.deepEqual([s.font, s.size, s.color, s.fill, s.align, s.valign, s.wrap], ['Meiryo', 16, '#c00000', '#fce4d6', 'center', 'top', true]);
  await x.click('#fill-none');
  cell = await x.cell('B4');
  assert.equal(cell.style.backgroundColor, 'rgba(0, 0, 0, 0)', '塗りなしのセルは透明で、後ろのシートの白が見える');
  assert.equal((await x.cells()).B4.s.fill, 'none', '白で塗るのではなく「塗りなし」になる（Excelで枠線が消えないように）');
  assert.equal(await x.value('font'), 'Meiryo', 'リボンは選んでいるセルの書式を示す');
});

test('罫線の種類の一覧は、Excelに貼ったときの線の名前で、すべてExcelで表せる線である', async t => {
  const x = await env.open(t, book);
  const options = await x.evaluate(() => [...document.querySelectorAll('#border-line option')].map(o => [o.value, o.textContent]));
  assert.deepEqual(options, [['1px solid', '細線'], ['2px solid', '中太線'], ['3px solid', '太線'], ['1px dashed', '破線'], ['1px dotted', '点線'], ['3px double', '二重線']]);
  const targets = ['B21', 'C21', 'D21', 'E21', 'F21', 'G21'];
  for (const [i, [line]] of options.entries()) {
    await x.clickCell(targets[i]);
    await x.setControl('border-line', line);
    await x.setControl('border-color', '#000000');
    await x.setControl('border', 'bottom');
    assert.equal(await x.dialog(), null, line + ' はエラーにならない');
  }
  const cells = await x.cells();
  assert.deepEqual(targets.map(a => cells[a].s.borderBottom), options.map(([line]) => line + ' #000000'));
});

test('罫線の格子・外枠・下罫線・罫線なしは、選んだ範囲の決まった辺に線を引く', async t => {
  const x = await env.open(t, book);
  await x.setControl('border-line', '1px solid');
  await x.setControl('border-color', '#000000');
  const line = '1px solid #000000';
  const sides = s => ['borderTop', 'borderRight', 'borderBottom', 'borderLeft'].map(k => s[k] ?? 'none');
  await x.select('B20:C21');
  await x.setControl('border', 'all');
  await x.select('D20:E21');
  await x.setControl('border', 'outer');
  await x.select('F20:G21');
  await x.setControl('border', 'bottom');
  let cells = await x.cells();
  for (const a of ['B20', 'C20', 'B21', 'C21']) assert.deepEqual(sides(cells[a].s), [line, line, line, line], '格子 ' + a);
  assert.deepEqual(sides(cells.D20.s), [line, 'none', 'none', line], '外枠の左上');
  assert.deepEqual(sides(cells.E21.s), ['none', line, line, 'none'], '外枠の右下');
  assert.deepEqual(sides(cells.F20.s ?? {}), ['none', 'none', 'none', 'none'], '下罫線は上の行に引かない');
  assert.deepEqual(sides(cells.G21.s), ['none', 'none', line, 'none'], '下罫線');
  assert.equal(await x.value('border'), '', '選んだ後、罫線の一覧は「罫線」の表示に戻る');
  await x.select('B20:C21');
  await x.setControl('border', 'none');
  cells = await x.cells();
  for (const a of ['B20', 'C20', 'B21', 'C21']) assert.deepEqual(sides(cells[a].s), ['none', 'none', 'none', 'none'], '罫線なし ' + a);
});

test('隣り合うセルの境界の罫線は、画面では片方のセルの線として1本だけ描かれ、データは変わらない', async t => {
  const x = await env.open(t, book);
  // C18 has a left border in the data; the screen draws it as B18's right border.
  assert.equal((await x.cell('B18')).inline.borderRight, '2px solid ' + rgb('#ff0000'));
  assert.equal((await x.cell('C18')).inline.borderLeft, '');
  // A grid of borders: the shared edges are drawn once, not twice.
  await x.setControl('border-line', '1px solid');
  await x.setControl('border-color', '#000000');
  await x.select('B20:C21');
  await x.setControl('border', 'all');
  assert.equal((await x.cell('C20')).inline.borderLeft, '', '右隣の左の線は描かない');
  assert.equal((await x.cell('B21')).inline.borderTop, '', '下隣の上の線は描かない');
  assert.notEqual((await x.cell('B20')).inline.borderRight, '');
  const cells = await x.cells();
  assert.equal(cells.C18.s.borderLeft, '2px solid #ff0000', 'データの4辺の指定はそのまま（Excelへのコピーは変わらない）');
  assert.equal(cells.B18?.s?.borderRight ?? 'none', 'none');
  assert.equal(cells.C20.s.borderLeft, '1px solid #000000');
});

test('値のないセル同士は確認なしで結合し、もう一度押すと解除する', async t => {
  const x = await env.open(t, book);
  await x.select('F12:G13');
  await x.action('merge');
  const merged = await x.cell('F12');
  assert.deepEqual([merged.colSpan, merged.rowSpan], [2, 2]);
  assert.equal(await x.cell('G13'), null, '結合の内側のセルは描かれない');
  assert.deepEqual((await x.state()).sheet('main').merges.at(-1), {r1: 11, c1: 5, r2: 12, c2: 6});
  await x.clickCell('F12');
  await x.action('merge');
  assert.deepEqual([(await x.cell('F12')).colSpan, (await x.cell('F12')).rowSpan], [1, 1]);
  assert.equal((await x.state()).sheet('main').merges.length, 1, '元からあった結合だけが残る');
});

test('左上以外に値がある範囲の結合は、確認で「キャンセル」なら何もせず、「OK」なら左上の値だけを残して結合する', async t => {
  const x = await env.open(t, book);
  await x.select('B4:C4');
  await x.action('merge');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, 'セルを結合');
  assert.match(dialog.body, /左上以外のセルの内容は削除されます/);
  await x.answer('cancel');
  assert.equal((await x.cell('B4')).colSpan, 1);
  assert.equal(await x.text('C4'), '3');
  await x.action('merge');
  await x.answer('ok');
  assert.equal((await x.cell('B4')).colSpan, 2);
  assert.equal(await x.text('B4'), '保守作業');
  const cells = await x.cells();
  assert.equal(cells.C4.v, '', '左上以外の値は消える');
  assert.equal(await x.text('E4'), '¥0', '消えた値を参照する数式も再計算される');
});

test('列幅・行の高さは、数値の入力でも、見出しの境界のドラッグでも変えられる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C4');
  await x.click('[data-tab="cells"]');
  await x.setControl('column-width', '150');
  await x.setControl('row-height', '40');
  let sheet = (await x.state()).sheet('main');
  assert.equal(sheet.colWidths[2], 150);
  assert.equal(sheet.rowHeights[3], 40);
  assert.ok(Math.abs((await x.cell('C4')).rect.width - 150) <= 1, '画面の列幅');
  const from = await grip(x, '[data-resize-col="1"]');
  await x.page.drag(from, {x: from.x + 50, y: from.y});
  const rowFrom = await grip(x, '[data-resize-row="5"]');
  await x.page.drag(rowFrom, {x: rowFrom.x, y: rowFrom.y + 20});
  sheet = (await x.state()).sheet('main');
  assert.equal(sheet.colWidths[1], 170, 'B列 120px + 50px');
  assert.equal(sheet.rowHeights[5], 42, '6行目 22px + 20px');
  await x.clickCell('H2');
  await x.press('Control+z');
  await x.press('Control+z');
  sheet = (await x.state()).sheet('main');
  assert.deepEqual([sheet.colWidths[1], sheet.rowHeights[5] ?? null], [120, null], 'ドラッグで変えた幅と高さも、元に戻せる');
});

test('「列幅の自動調整」は、選んだ列の幅を中身に合わせる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B4');
  await x.action('autofit-col');
  const width = (await x.state()).sheet('main').colWidths[1];
  assert.ok(width > 50 && width < 120, `B列（最長は「設定作業」）が中身に合わせて縮む: ${width}px`);
  await x.enter('G4', 'とても長い文字列をこの列に入れて幅を広げる');
  await x.clickCell('G4');
  await x.action('autofit-col');
  const g = (await x.state()).sheet('main').colWidths[6];
  assert.ok(g > 200, `G列が長い文字に合わせて広がる: ${g}px`);
});

test('列と行の見出しの境界をダブルクリックすると、中身に合わせて幅と高さを変える', async t => {
  const x = await env.open(t, book);
  await x.page.dblclick(await grip(x, '[data-resize-col="1"]'));
  const width = (await x.state()).sheet('main').colWidths[1];
  assert.ok(width > 50 && width < 120, `B列が中身に合わせて縮む: ${width}px`);
  await x.page.dblclick(await grip(x, '[data-resize-row="0"]'));
  assert.notEqual((await x.state()).sheet('main').rowHeights[0], 30, '1行目（30px）の高さが、14ptの題名1行に合わせて変わる');
});

test('「行高の自動調整」は、折り返したセルの行数に合わせて行を高くする', async t => {
  const x = await env.open(t, book);
  await x.clickCell('B6');
  await x.click('[data-style="wrap"]');
  await x.enter('B6', 'とても長い品名なので列の幅では収まらずに折り返して何行にもなる');
  await x.clickCell('B6');
  await x.action('autofit-row');
  const h = (await x.state()).sheet('main').rowHeights[5];
  assert.ok(h > 44, `3行以上の高さになる: ${h}px`);
});

// Several rows or columns at once, as in Excel: select whole rows (columns) with their headers, then drag or
// double-click the boundary of any one of the selected headers. Every selected row (column) changes.
test('行見出しで複数の行を選び、そのうち1行の見出しの境界をドラッグすると、選んだ行すべてが同じ高さになる', async t => {
  const x = await env.open(t, uneven);
  await selectRows(x, 3, 5);
  assert.equal(await x.nameBox(), 'A4:J6', '4〜6行目を行全体で選んでいる');
  const from = await grip(x, '[data-resize-row="4"]');
  await x.page.drag(from, {x: from.x, y: from.y + 20});
  const heights = (await x.state()).sheet('main').rowHeights;
  assert.deepEqual(heights.slice(3, 6), [50, 50, 50], '5行目（30px）を 20px 広げた 50px に、4行目（24px）と6行目（40px）もそろう');
  assert.deepEqual([heights[2], heights[6] ?? null], [22, null], '選んでいない3行目と7行目は変わらない');
  for (const addr of ['A4', 'A5', 'A6']) assert.ok(Math.abs((await x.cell(addr)).rect.height - 50) <= 1, `画面でも ${addr} の行が 50px`);
});

test('列見出しで複数の列を選び、そのうち1列の見出しの境界をドラッグすると、選んだ列すべてが同じ幅になる', async t => {
  const x = await env.open(t, book);
  await selectCols(x, 3, 5);
  assert.equal(await x.nameBox(), 'D1:F30', 'D〜F列を列全体で選んでいる');
  const from = await grip(x, '[data-resize-col="4"]');
  await x.page.drag(from, {x: from.x + 30, y: from.y});
  const widths = (await x.state()).sheet('main').colWidths;
  assert.deepEqual(widths.slice(3, 6), [140, 140, 140], 'E列（110px）を 30px 広げた 140px に、D列とF列（96px）もそろう');
  assert.deepEqual([widths[2], widths[6]], [80, 96], '選んでいないC列とG列は変わらない');
  for (const addr of ['D4', 'E4', 'F4']) assert.ok(Math.abs((await x.cell(addr)).rect.width - 140) <= 1, `画面でも ${addr} の列が 140px`);
});

test('行見出しで複数の行を選び、そのうち1行の見出しの境界をダブルクリックすると、選んだ行すべてが中身に合わせた高さになる', async t => {
  const x = await env.open(t, oversized);
  await selectRows(x, 3, 5);
  await x.page.dblclick(await grip(x, '[data-resize-row="4"]'));
  const heights = (await x.state()).sheet('main').rowHeights.slice(3, 6);
  assert.ok(heights.every(h => h < 60), `選んだ3行（どれも 60px）とも、1行の文字に合わせて低くなる: ${heights.join(', ')}`);
  await x.action('autofit-row');
  assert.deepEqual((await x.state()).sheet('main').rowHeights.slice(3, 6), heights, 'リボンの「行高の自動調整」と同じ高さ（自動調整し直しても変わらない）');
});

test('列見出しで複数の列を選び、そのうち1列の見出しの境界をダブルクリックすると、選んだ列すべてが中身に合わせた幅になる', async t => {
  const x = await env.open(t, oversized);
  await selectCols(x, 3, 5);
  await x.page.dblclick(await grip(x, '[data-resize-col="4"]'));
  const widths = (await x.state()).sheet('main').colWidths.slice(3, 6);
  assert.ok(widths.every(w => w < 300), `選んだ3列（どれも 300px）とも、中身に合わせて狭くなる: ${widths.join(', ')}`);
  await x.action('autofit-col');
  assert.deepEqual((await x.state()).sheet('main').colWidths.slice(3, 6), widths, 'リボンの「列幅の自動調整」と同じ幅（自動調整し直しても変わらない）');
});

test('リボンの「列幅の自動調整」「行高の自動調整」は、見出しで選んだ列・行のすべてを中身に合わせる', async t => {
  const x = await env.open(t, oversized);
  await selectCols(x, 3, 5);
  await x.action('autofit-col');
  const widths = (await x.state()).sheet('main').colWidths.slice(3, 6);
  assert.ok(widths.every(w => w < 300), `選んだ3列（どれも 300px）とも狭くなる: ${widths.join(', ')}`);
  assert.deepEqual([await x.text('D6'), await x.text('E9')], ['¥30,000', '¥83,600'], '狭くしても、数値は #### にならずに全部見える');
  await selectRows(x, 3, 5);
  await x.action('autofit-row');
  const heights = (await x.state()).sheet('main').rowHeights.slice(3, 6);
  assert.ok(heights.every(h => h < 60), `選んだ3行（どれも 60px）とも低くなる: ${heights.join(', ')}`);
});

test('ドラッグした境界の行や列を見出しで丸ごと選んでいないときは、その1行・1列だけが変わる', async t => {
  const x = await env.open(t, uneven);
  await selectRows(x, 3, 5);
  const row = await grip(x, '[data-resize-row="7"]');
  await x.page.drag(row, {x: row.x, y: row.y + 20});
  const heights = (await x.state()).sheet('main').rowHeights;
  assert.deepEqual([...heights.slice(3, 6), heights[7]], [24, 30, 40, 42], '選んだ4〜6行目の外にある8行目（22px）だけが 20px 広がる');
  await x.select('D4:F6');
  const col = await grip(x, '[data-resize-col="4"]');
  await x.page.drag(col, {x: col.x + 30, y: col.y});
  assert.deepEqual((await x.state()).sheet('main').colWidths.slice(3, 6), [96, 140, 96], 'セルの範囲を選んでいるだけなら、Excelと同じく境界のE列だけが変わる');
});

test('「枠線の表示」で補助の枠線を消し、もう一度押すと戻す', async t => {
  const x = await env.open(t, book);
  await x.action('gridlines');
  assert.equal(await x.evaluate(() => document.getElementById('grid').classList.contains('no-grid')), true);
  assert.equal((await x.state()).sheet('main').gridlines, false);
  await x.action('gridlines');
  assert.equal(await x.evaluate(() => document.getElementById('grid').classList.contains('no-grid')), false);
});

test('枠の固定：選んだセルより上の行と左の列が、スクロールしても動かない', async t => {
  const large = fixture('book');
  Object.assign(large.sheets[0], {rows: 120, cols: 40});
  const x = await env.open(t, large);
  await x.clickCell('D3');
  await x.action('freeze');
  assert.deepEqual((await x.state()).sheet('main').freeze, {rows: 2, cols: 3});
  const before = {fixed: await x.cell('B2'), row: await x.cell('E1'), col: await x.cell('A6'), free: await x.cell('E6')};
  assert.deepEqual([before.fixed.style.position, before.fixed.style.zIndex], ['sticky', '6'], '行と列の両方で固定');
  assert.deepEqual([before.row.style.position, before.row.style.zIndex], ['sticky', '4'], '行だけで固定');
  assert.equal(before.col.style.zIndex, '4', '列だけで固定');
  assert.equal(before.free.style.position, 'relative');
  assert.equal(before.fixed.style.backgroundColor, 'rgb(255, 255, 255)', '固定セルは白で塗り、後ろを流れる値を透かさない');
  await x.evaluate(() => { const v = document.getElementById('viewport'); v.scrollTop = 120; v.scrollLeft = 150; v.dispatchEvent(new Event('scroll')); });
  const after = {fixed: await x.cell('B2'), row: await x.cell('E1'), col: await x.cell('A6'), free: await x.cell('E6')};
  assert.deepEqual([after.fixed.rect.left, after.fixed.rect.top], [before.fixed.rect.left, before.fixed.rect.top], '両方で固定したセルは動かない');
  assert.equal(after.row.rect.top, before.row.rect.top, '固定した行は上下に動かない');
  assert.equal(after.col.rect.left, before.col.rect.left, '固定した列は左右に動かない');
  assert.equal(after.free.rect.top, before.free.rect.top - 120, '固定していないセルはスクロールに合わせて動く');
  await x.action('unfreeze');
  assert.deepEqual((await x.state()).sheet('main').freeze, {rows: 0, cols: 0});
  assert.equal((await x.cell('B2')).style.position, 'relative');
});

test('右のセルが空いていれば文字ははみ出して表示され、はみ出した上をクリックすると下のセルが選ばれる', async t => {
  const x = await env.open(t, book);
  const a1 = await x.cell('A1');
  assert.ok(a1.classes.includes('ov'), 'A1 の文字がはみ出している');
  assert.ok(a1.textRect.width > a1.rect.width + 100, '題名が右の列まで描かれる');
  const a14 = await x.cell('A14');
  assert.ok(!a14.classes.includes('ov'), '右に値（B14）があると、はみ出さない');
  assert.ok(a14.textRect.width <= a14.rect.width, 'A14 はセルの幅で切れる');
  await x.clickCell('B1');
  assert.equal(await x.nameBox(), 'B1', 'はみ出した文字はクリックを受け取らない');
});

test('結合セルを横切る位置では枠を固定せず、理由を知らせる', async t => {
  const x = await env.open(t, book);
  await x.clickCell('C3');
  await x.action('freeze');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, '操作を完了できません');
  assert.match(dialog.body, /結合をまたぐ固定は未対応/, 'A11:C12 の結合が、A〜B列の固定の境目を横切る');
  await x.answer('ok');
  assert.equal((await x.state()).sheet('main').freeze, undefined);
});

test('固定した列の文字は、Excelと同じく固定の境目を越えて右の空きセルへはみ出す', async t => {
  const x = await env.open(t, book);
  await x.clickCell('D3');
  await x.action('freeze');
  const a1 = await x.cell('A1'), d1 = await x.cell('D1');
  assert.equal(d1.text, a1.title, '境目より右の部分は、固定していない最初の空きセル D1 が描く');
  assert.ok(d1.classes.includes('ov'));
  assert.ok(a1.textRect.left + a1.textRect.width <= d1.rect.left + 1, '元のセルの文字は固定の境目で切れる');
  assert.equal(a1.style.zIndex, '7', 'はみ出す固定セルは、白く塗った隣の固定セルより手前（6 + 1）');
  assert.equal(d1.style.zIndex, '5', '続きを描く行だけ固定のセルも 1 つ手前（4 + 1）');
});

test('狭い列の数値は、表示形式があれば #### になり、標準なら桁を減らして表示する', async t => {
  const x = await env.open(t, book);
  assert.match(await x.text('I2'), /^#+$/);
  assert.equal((await x.cell('I2')).title, '1,234,567,890', 'セルの説明には全体の値が出る');
  const j2 = await x.text('J2');
  assert.notEqual(j2, '3.14159265358979');
  assert.match(j2, /^3\.1\d*$/);
});

test('リンクのセルは Ctrl+クリックで開き、「リンク」で付け外しができ、http(s) 以外は受け付けない', async t => {
  const x = await env.open(t, book);
  assert.equal((await x.cell('B16')).link, 'https://example.com/');
  await x.clickCell('B16');
  assert.deepEqual(await x.evaluate(() => window.__opened), [], '普通のクリックでは開かない');
  await x.clickCell('B16', {ctrl: true});
  assert.deepEqual(await x.evaluate(() => window.__opened), [['https://example.com/', '_blank', 'noopener,noreferrer']]);
  await x.clickCell('B17');
  await x.action('link');
  await x.answer('ok', 'https://example.org/page');
  assert.equal((await x.cells()).B17.link, 'https://example.org/page');
  await x.clickCell('B16');
  await x.action('link');
  assert.equal((await x.dialog()).input, 'https://example.com/', '今のリンクが入力欄に出る');
  await x.answer('ok', '');
  assert.equal((await x.cells()).B16.link, undefined, '空欄で解除');
  await x.clickCell('B17');
  await x.action('link');
  await x.answer('ok', 'javascript:alert(1)');
  const dialog = await x.dialog();
  assert.equal(dialog?.title, '操作を完了できません');
  assert.match(dialog.body, /リンクはhttp\(s\)のみ/);
  await x.answer('ok');
  assert.equal((await x.cells()).B17.link, 'https://example.org/page', '受け付けなかったときは元のまま');
});
