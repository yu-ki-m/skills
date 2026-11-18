// scripts/xlsx.cjs：ブックを、Excelで開けるxlsxファイルとして書き出す。
// 書き方は、同じ表をデスクトップ版Excel（16.0、ビルド20326）の中で作って保存したxlsxに合わせてある。
// ここでは、書き出したファイルを tests/lib/xlsx.mjs で読み、セル・書式・寸法・表示・リンク・画像を確かめる。
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {test} from 'node:test';
import {blankBook, SKILL_ROOT} from './lib/workbook.mjs';
import {readXlsx} from './lib/xlsx.mjs';

const {toXlsx} = createRequire(import.meta.url)(join(SKILL_ROOT, 'scripts', 'xlsx.cjs'));
const read = book => readXlsx(toXlsx(book).bytes);
const sheetOf = book => read(book).sheets[0];
const valueOf = (sheet, a) => ({type: sheet.cells[a]?.type, value: sheet.cells[a]?.value});
const png = Buffer.from('png の中身のかわり');

test('セルの値を、Excelと同じ型で書く（文字列・数値・真偽値・書式だけの空のセル）', () => {
  const x = read(blankBook({sheet: {cells: {
    A1: {t: 's', v: '見出し'}, B1: {t: 'n', v: 1234567.891}, C1: {t: 'b', v: true}, D1: {t: 'b', v: false},
    A2: {t: 's', v: '', s: {fill: '#ffff00'}}, B2: {t: 'n', v: 0.1}, C2: {t: 'n', v: 123456789012345}, D2: {t: 's', v: '1行目\n2行目', s: {wrap: true}},
    E2: {t: 's', v: '  前後の空白  '}, F2: {t: 's', v: '<&>"\''}, G2: {t: 's', v: '見出し'}, H2: {t: 's', v: ''},
  }}}));
  const s = x.sheets[0];
  assert.deepEqual(valueOf(s, 'A1'), {type: 's', value: '見出し'});
  assert.deepEqual(valueOf(s, 'B1'), {type: 'n', value: 1234567.891});
  assert.deepEqual(valueOf(s, 'C1'), {type: 'b', value: true});
  assert.deepEqual(valueOf(s, 'D1'), {type: 'b', value: false});
  assert.equal(s.cells.A2.value, undefined, '値のない書式だけのセルは、値を持たない');
  assert.equal(s.cells.A2.style.fill, '#ffff00');
  assert.deepEqual(valueOf(s, 'B2'), {type: 'n', value: 0.1});
  assert.deepEqual(valueOf(s, 'C2'), {type: 'n', value: 123456789012345}, '15桁の数値も全桁');
  assert.deepEqual(valueOf(s, 'D2'), {type: 's', value: '1行目\n2行目'});
  assert.deepEqual(valueOf(s, 'E2'), {type: 's', value: '  前後の空白  '});
  assert.deepEqual(valueOf(s, 'F2'), {type: 's', value: '<&>"\''});
  assert.equal(s.cells.H2, undefined, '値も書式もないセルは書かない');
  assert.equal(x.strings.filter(t => t === '見出し').length, 1, '同じ文字列は共有文字列の表に1回だけ入る');
});

test('数式は数式のまま書き、HTMLが計算した値も一緒に入れて、Excelが開くときに計算し直すようにする', () => {
  const x = read(blankBook({sheet: {cells: {
    A1: {t: 'n', v: 2}, A2: {t: 'n', v: 3},
    B1: {t: 'f', f: '=SUM(A1:A2)*2'}, B2: {t: 'f', f: '=round(a1/a2,2)&"円"'}, B3: {t: 'f', f: '=A1>A2'}, B4: {t: 'f', f: '=1/0'},
    B5: {t: 'f', f: '=B5'}, B6: {t: 'f', f: '=$a$1+a$2'}, B7: {t: 'f', f: '=IF(A1=2,"はい","いいえ")'}, B8: {t: 'f', f: '="total: "&a1'},
    B9: {t: 'f', f: '=vlookup(A1,A1:A2,1,FALSE)', s: {numFmt: '#,##0'}},
  }}}));
  const s = x.sheets[0], f = a => ({formula: s.cells[a].formula, ...valueOf(s, a)});
  assert.deepEqual(f('B1'), {formula: 'SUM(A1:A2)*2', type: 'n', value: 10});
  assert.deepEqual(f('B2'), {formula: 'ROUND(A1/A2,2)&"円"', type: 'str', value: '0.67円'}, '関数名と番地は大文字にし、文字列はそのまま');
  assert.deepEqual(f('B3'), {formula: 'A1>A2', type: 'b', value: false});
  assert.deepEqual(f('B4'), {formula: '1/0', type: 'e', value: '#DIV/0!'});
  assert.deepEqual(f('B5'), {formula: 'B5', type: 'n', value: undefined}, '循環参照はHTMLだけの #CIRC! なので、計算結果を入れない');
  assert.deepEqual(f('B6'), {formula: '$A$1+A$2', type: 'n', value: 5});
  assert.deepEqual(f('B7'), {formula: 'IF(A1=2,"はい","いいえ")', type: 'str', value: 'はい'});
  assert.deepEqual(f('B8'), {formula: '"total: "&A1', type: 'str', value: 'total: 2'});
  assert.equal(s.cells.B9.formula, 'VLOOKUP(A1,A1:A2,1,FALSE)', 'HTMLで計算できない関数も、Excelでは計算できるので数式のまま書く');
  assert.equal(s.cells.B9.style.numFmt, '#,##0');
  assert.equal(x.calc.fullCalcOnLoad, '1', 'Excelが開くときに全部の数式を計算し直す');
});

test('HTMLの数式として読めない式は、Excelで壊れたファイルにしないよう、その文字のまま文字列で書き、そのことを知らせる', () => {
  const book = blankBook({sheet: {cells: {A1: {t: 'f', f: '=SUM(A2:A3'}, A2: {t: 'f', f: '=Sheet2!A1'}, A3: {t: 'f', f: '=A1+'}}}});
  const {bytes, warnings} = toXlsx(book), s = readXlsx(bytes).sheets[0];
  for (const [a, text] of [['A1', '=SUM(A2:A3'], ['A2', '=Sheet2!A1'], ['A3', '=A1+']]) {
    assert.deepEqual({...valueOf(s, a), formula: s.cells[a].formula}, {type: 's', value: text, formula: undefined}, a);
    assert.equal(s.cells[a].style.quotePrefix, true, 'Excelで「\'」を付けて入力した文字列と同じにする');
  }
  assert.equal(warnings.length, 3);
  assert.match(warnings[0], /Sheet1.*A1.*=SUM\(A2:A3/);
});

test('Excelが数値・日付・数式として読む文字列は、Excelで先頭に「\'」を付けて入力したときと同じ状態にする', () => {
  const s = sheetOf(blankBook({sheet: {cells: {
    A1: {t: 's', v: '00123'}, A2: {t: 's', v: '00123', s: {numFmt: '@'}}, A3: {t: 's', v: '1/2'}, A4: {t: 's', v: 'TRUE'},
    A5: {t: 's', v: '=式ではない'}, A6: {t: 's', v: ' 前後に空白 '}, A7: {t: 's', v: '普通の文字'},
  }}}));
  const prefixed = Object.fromEntries(['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7'].map(a => [a, s.cells[a].style.quotePrefix]));
  assert.deepEqual(prefixed, {A1: true, A2: false, A3: true, A4: true, A5: true, A6: false, A7: false});
  assert.equal(s.cells.A2.style.numFmt, '@', '表示形式「文字列」のセルには付けない（Excelも付けない）');
  assert.deepEqual(valueOf(s, 'A1'), {type: 's', value: '00123'});
});

test('表示形式は、Excelの組み込みの形式があればその番号で、なければExcelと同じ書き方で書く', () => {
  const formats = ['General', '@', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '¥#,##0', 'yyyy/mm/dd'];
  const cells = Object.fromEntries(formats.map((numFmt, i) => ['A' + (i + 1), {t: 'n', v: 1, s: {numFmt}}]));
  const x = read(blankBook({sheet: {cells}}));
  const got = formats.map((_, i) => x.sheets[0].cells['A' + (i + 1)].style.numFmt);
  assert.deepEqual(got, ['General', '@', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '"¥"#,##0', 'yyyy/mm/dd']);
  const custom = Object.entries(x.styles.formats).filter(([id]) => +id >= 164).map(([, code]) => code).sort();
  assert.deepEqual(custom, ['"¥"#,##0', 'yyyy/mm/dd'], '組み込みにない2つだけを独自の形式として定義する');
});

test('フォント・文字色・塗り・配置・折り返し・インデントを書き、既定の書式はブックの defaultStyle から受け継ぐ', () => {
  const s = sheetOf(blankBook({defaultStyle: {size: 10}, sheet: {cells: {
    A1: {t: 's', v: 'a', s: {bold: true, italic: true, underline: true, strike: true, color: '#c00000', size: 14, font: 'Meiryo'}},
    A2: {t: 's', v: 'b', s: {fill: '#e2efda', align: 'center', valign: 'top', wrap: true}},
    A3: {t: 'n', v: 1, s: {indent: 2}}, A4: {t: 's', v: 'c', s: {align: 'right', indent: 1, valign: 'bottom'}},
    A5: {t: 's', v: 'd', s: {font: 'Arial'}}, A6: {t: 's', v: 'e'},
  }}}));
  const font = a => { const {name, size, bold, italic, strike, underline, color} = s.cells[a].style.font; return {name, size, bold, italic, strike, underline, color}; };
  assert.deepEqual(font('A1'), {name: 'メイリオ', size: 14, bold: true, italic: true, strike: true, underline: 'single', color: '#c00000'}, '日本語版Excelと同じく、フォント名は日本語の名前で書く');
  assert.equal(s.cells.A1.style.font.charset, 128, '日本語のフォントは、文字セットを日本語にする');
  assert.equal(s.cells.A2.style.fill, '#e2efda');
  assert.deepEqual(s.cells.A2.style.alignment, {horizontal: 'center', vertical: 'top', wrapText: true, indent: 0});
  assert.deepEqual(s.cells.A3.style.alignment, {horizontal: 'left', vertical: 'center', wrapText: false, indent: 2}, '「標準」の配置で字下げすると、Excelと同じく左揃えになる');
  assert.deepEqual(s.cells.A4.style.alignment, {horizontal: 'right', vertical: 'bottom', wrapText: false, indent: 1});
  assert.equal(s.cells.A5.style.font.name, 'Arial');
  assert.notEqual(s.cells.A5.style.font.charset, 128);
  assert.deepEqual(font('A6'), {name: '游ゴシック', size: 10, bold: false, italic: false, strike: false, underline: 'none', color: '#000000'});
  assert.deepEqual({fill: s.cells.A6.style.fill, vertical: s.cells.A6.style.alignment.vertical}, {fill: 'none', vertical: 'center'}, '日本語版Excelの既定と同じく上下中央');
});

test('罫線の種類と色を、Excelの線に対応させる', () => {
  const kinds = {'0.5px solid': 'hair', '1px solid': 'thin', '2px solid': 'medium', '3px solid': 'thick', '1px dashed': 'dashed',
    '2px dashed': 'mediumDashed', '1px dotted': 'dotted', '3px double': 'double'};
  const cells = Object.fromEntries(Object.keys(kinds).map((k, i) => ['B' + (2 * i + 2), {t: 's', v: k, s: {borderTop: k + ' #12ab' + String(i).padStart(2, '0')}}]));
  cells.D2 = {t: 's', v: '4辺', s: {borderTop: '1px solid #ff0000', borderRight: '2px solid #00ff00', borderBottom: '3px double #0000ff', borderLeft: '1px dotted #000000'}};
  const s = sheetOf(blankBook({sheet: {cells}}));
  Object.values(kinds).forEach((style, i) => assert.deepEqual(s.cells['B' + (2 * i + 2)].style.borders.top, {style, color: '#12ab' + String(i).padStart(2, '0')}, style));
  assert.deepEqual(s.cells.D2.style.borders, {top: {style: 'thin', color: '#ff0000'}, right: {style: 'medium', color: '#00ff00'},
    bottom: {style: 'double', color: '#0000ff'}, left: {style: 'dotted', color: '#000000'}});
});

test('結合セルは、左上のセルの書式で全体を書き、罫線は外枠の辺に接するセルにだけ持たせる', () => {
  const s = sheetOf(blankBook({sheet: {merges: [{r1: 0, c1: 0, r2: 1, c2: 2}], cells: {
    A1: {t: 's', v: '題名', s: {bold: true, fill: '#e2efda', align: 'center', borderTop: '2px solid #217346', borderBottom: '1px solid #000000',
      borderLeft: '1px dashed #ff0000', borderRight: '3px solid #0000ff'}},
    B2: {t: 's', v: '', s: {fill: '#ff0000', borderTop: '1px solid #000000'}},
  }}}));
  assert.deepEqual(s.merges, ['A1:C2']);
  const edges = a => Object.fromEntries(Object.entries(s.cells[a].style.borders).filter(([, v]) => v !== 'none').map(([k, v]) => [k, v.style]));
  assert.deepEqual(edges('A1'), {top: 'medium', left: 'dashed'});
  assert.deepEqual(edges('B1'), {top: 'medium'});
  assert.deepEqual(edges('C1'), {top: 'medium', right: 'thick'});
  assert.deepEqual(edges('A2'), {bottom: 'thin', left: 'dashed'});
  assert.deepEqual(edges('B2'), {bottom: 'thin'}, '結合で隠れたセル自身の書式は使わない');
  assert.deepEqual(edges('C2'), {right: 'thick', bottom: 'thin'});
  for (const a of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2']) {
    assert.equal(s.cells[a].style.fill, '#e2efda', a);
    assert.equal(s.cells[a].style.font.bold, true, a);
    assert.equal(s.cells[a].style.alignment.horizontal, 'center', a);
  }
  assert.deepEqual(valueOf(s, 'A1'), {type: 's', value: '題名'});
  assert.equal(s.cells.B1.value, undefined);
});

test('列幅と行の高さは、指定のない列・行も含めて、HTMLと同じ大きさになる値で書く', () => {
  const x = read(blankBook({sheet: {rows: 6, cols: 5, colWidths: [24, 150, null, 2044], rowHeights: [40, null, 16, 546], cells: {B2: {t: 's', v: 'x'}}}}));
  const s = x.sheets[0];
  // 列幅は、標準のフォント（游ゴシック11pt）の数字1文字の幅を単位に書く。96dpiで8px、120dpiで10pxなので、HTMLのpx÷8になる。
  assert.deepEqual([0, 1, 2, 3, 4, 5].map(c => s.widthOf(c)), [3, 18.75, 12, 255.5, 12, undefined], '指定のない列はHTMLの既定の96px、表の外の列はExcelの既定');
  assert.ok(s.cols.every(c => c.customWidth));
  const font = x.styles.cellXfs[0].font;
  assert.deepEqual({name: font.name, size: font.size}, {name: '游ゴシック', size: 11}, '列幅の単位を決める標準のフォントは、日本語版Excelと同じ');
  assert.deepEqual([1, 3, 4].map(r => s.rows[r]), [{ht: 30, customHeight: true}, {ht: 12, customHeight: true}, {ht: 409.5, customHeight: true}], '行の高さはpt（px×0.75）');
  // 高さを書かない行は、Excelが中の文字に合わせて高くする（游ゴシック11ptで18pt、デスクトップ版Excel 16.0で確認）。
  assert.deepEqual(s.rows[2], {ht: 16.5, customHeight: true}, 'セルのある行は、指定がなくてもHTMLの既定の22pxを、手で決めた高さとして書く');
  assert.equal(s.rows[5], undefined, 'セルのない行は書かない');
  assert.deepEqual({height: s.defaultRowHeight, custom: s.customDefaultHeight}, {height: 16.5, custom: true}, '書かない行の高さも、HTMLの既定の22px');
});

test('枠の固定・枠線の表示・倍率・最初に開くシート・選択している範囲を書く', () => {
  const book = blankBook({activeSheet: 's3', view: {zoom: 120, active: [3, 2], selection: {r1: 3, c1: 2, r2: 4, c2: 3}}});
  const base = book.sheets[0];
  book.sheets = [{...base, id: 's1', name: '固定なし'}, {...base, id: 's2', name: '行だけ', freeze: {rows: 3, cols: 0}},
    {...base, id: 's3', name: '行と列', freeze: {rows: 2, cols: 1}, gridlines: false}, {...base, id: 's4', name: '列だけ', freeze: {rows: 0, cols: 2}}];
  const x = read(book), [s1, s2, s3, s4] = x.sheets;
  assert.equal(x.activeTab, 2);
  assert.deepEqual(x.sheets.map(s => s.view.tabSelected), [false, false, true, false]);
  assert.deepEqual(x.sheets.map(s => s.view.showGridLines), [true, true, false, true]);
  assert.deepEqual(x.sheets.map(s => s.view.zoomScale), [120, 120, 120, 120], 'HTMLの倍率はブック全体なので、全シートに入れる');
  assert.equal(s1.view.pane, null);
  assert.deepEqual(s2.view.pane, {xSplit: 0, ySplit: 3, topLeftCell: 'A4', activePane: 'bottomLeft', state: 'frozen'});
  assert.deepEqual(s3.view.pane, {xSplit: 1, ySplit: 2, topLeftCell: 'B3', activePane: 'bottomRight', state: 'frozen'});
  assert.deepEqual(s4.view.pane, {xSplit: 2, ySplit: 0, topLeftCell: 'C1', activePane: 'topRight', state: 'frozen'});
  assert.deepEqual(s3.view.selections, [{pane: 'bottomRight', activeCell: 'C4', sqref: 'C4:D5'}], '開いているシートの選択範囲');
});

test('リンクは、Excelでセルにハイパーリンクを付けたときと同じく、青の下線の文字にする（フォント名と大きさは元のまま）', () => {
  const s = sheetOf(blankBook({sheet: {cells: {
    A1: {t: 's', v: 'リンク', link: 'https://example.com/a?b=1&c=2'},
    A2: {t: 's', v: '太字の赤', link: 'https://example.com/', s: {bold: true, color: '#c00000', size: 14, font: 'Meiryo', fill: '#ffff00'}},
  }}}));
  assert.deepEqual(s.links, {A1: 'https://example.com/a?b=1&c=2', A2: 'https://example.com/'});
  const font = a => { const {name, size, bold, underline, color} = s.cells[a].style.font; return {name, size, bold, underline, color}; };
  assert.deepEqual(font('A1'), {name: '游ゴシック', size: 11, bold: false, underline: 'single', color: '#0563c1'});
  assert.deepEqual(font('A2'), {name: 'メイリオ', size: 14, bold: false, underline: 'single', color: '#0563c1'});
  assert.equal(s.cells.A2.style.fill, '#ffff00', 'フォント以外の書式は残る');
  assert.equal(s.cells.A1.style.cellStyle, 8, 'セルのスタイルは、Excelの組み込みの「ハイパーリンク」');
});

test('画像は、置いたセルからのずれと大きさのまま、セルと一緒に動く画像として書き、同じ画像の中身は1つにまとめる', () => {
  const book = blankBook({assets: {pic: 'data:image/png;base64,' + png.toString('base64'), photo: 'data:image/jpeg;base64,' + Buffer.from('jpeg').toString('base64')}});
  const base = book.sheets[0];
  book.sheets = [
    {...base, colWidths: [80, 120], rowHeights: [30, 30], images: [
      {id: 'p1', asset: 'pic', r: 3, c: 1, x: 10, y: 5, w: 120, h: 60},
      {id: 'p2', asset: 'pic', r: 0, c: 0, x: 100, y: 40, w: 30, h: 20},
    ]},
    {...base, id: 'other', name: '別のシート', images: [{id: 'p3', asset: 'photo', r: 1, c: 1, x: 0, y: 0, w: 50, h: 40}]},
  ];
  const x = read(book), [p1, p2] = x.sheets[0].pictures, [p3] = x.sheets[1].pictures;
  const EMU = 9525; // CSSの1px（0.75pt）
  assert.deepEqual({anchor: p1.anchor, from: p1.from, ext: p1.ext}, {anchor: 'oneCellAnchor', from: {col: 1, colOff: 10 * EMU, row: 3, rowOff: 5 * EMU}, ext: {cx: 120 * EMU, cy: 60 * EMU}});
  assert.deepEqual(p2.from, {col: 1, colOff: 20 * EMU, row: 1, rowOff: 10 * EMU}, 'セルの幅や高さを超えるずれは、その先のセルからのずれにする');
  assert.deepEqual(p1.bytes, png);
  assert.equal(p1.media, p2.media, '同じ画像の中身は1つのファイルにする');
  assert.equal(x.typeOf(p1.media), 'image/png');
  assert.equal(x.typeOf(p3.media), 'image/jpeg');
  assert.deepEqual(p3.bytes, Buffer.from('jpeg'));
  assert.equal(x.sheets[1].pictures.length, 1, '画像はシートごとに置く');
});

test('シートの名前と順番、ブックの題名を保つ', () => {
  const book = blankBook({title: '見積 & "明細"', activeSheet: 'a'});
  const base = book.sheets[0];
  book.sheets = [{...base, id: 'a', name: '売上 & 原価'}, {...base, id: 'b', name: '集計<2026>'}, {...base, id: 'c', name: "O'Brien"}];
  const x = read(book);
  assert.deepEqual(x.sheets.map(s => s.name), ['売上 & 原価', '集計<2026>', "O'Brien"]);
  assert.equal(x.title, '見積 & "明細"');
});

test('XMLに書けない制御文字や、_x0041_ のような文字も、元の文字のまま読めるように書く', () => {
  const s = sheetOf(blankBook({sheet: {cells: {A1: {t: 's', v: '制御\u0001文字'}, A2: {t: 's', v: '_x0041_ はそのまま'}, A3: {t: 's', v: 'CR\rLF\n'}}}}));
  assert.equal(s.cells.A1.value, '制御\u0001文字');
  assert.equal(s.cells.A2.value, '_x0041_ はそのまま');
  assert.equal(s.cells.A3.value, 'CR\rLF\n');
});

test('同じブックからは、毎回同じ内容のファイルを作る（作った日時や利用者の情報を入れない）', () => {
  const book = blankBook({sheet: {cells: {A1: {t: 's', v: '同じ'}}}});
  assert.deepEqual(toXlsx(book).bytes, toXlsx(structuredClone(book)).bytes);
  const core = read(book).parts.get('docProps/core.xml').toString('utf8');
  assert.ok(!/creator|lastModifiedBy|created|modified/.test(core), core);
});
