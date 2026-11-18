// scripts/read_xlsx.cjs：Excelのファイル（xlsx）を、HTMLの表のデータ（ブック）として読み込む。
// tests/fixtures/excel-made.xlsx は、デスクトップ版Excel 16.0（ビルド20326）の中で作って保存した見本である
// （作り方は tests/fixtures/make-excel-made.ps1）。期待する色・大きさは、そのファイルをExcelで開いて読み取った値である。
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {test} from 'node:test';
import {blankBook, Core, SKILL_ROOT} from './lib/workbook.mjs';
import {miniXlsx, readXlsx} from './lib/xlsx.mjs';

const require = createRequire(import.meta.url);
const {fromXlsx, tint} = require(join(SKILL_ROOT, 'scripts', 'read_xlsx.cjs'));
const {toXlsx} = require(join(SKILL_ROOT, 'scripts', 'xlsx.cjs'));
const fixture = () => fromXlsx(readFileSync(join(SKILL_ROOT, 'tests', 'fixtures', 'excel-made.xlsx')), {name: 'excel-made.xlsx'});
const warned = (result, pattern) => result.warnings.some(w => pattern.test(w));

test('Excelで作ったxlsxを読み込むと、値・型・書式・結合がExcelの表示どおりに入る', () => {
  const {workbook: w} = fixture(), s = w.sheets[0];
  assert.deepEqual(w.sheets.map(x => x.name), ['見積', '単価表']);
  assert.deepEqual(w.defaultStyle, {font: 'Yu Gothic', size: 11, valign: 'middle'}, '標準のスタイルのフォントと縦位置が既定になる（フォント名は、リボンの一覧と同じ英語名）');
  assert.deepEqual(s.merges, [{r1: 0, c1: 0, r2: 0, c2: 5}]);
  const border = '2px solid #000000';
  assert.deepEqual(s.cells.A1, {t: 's', v: '御見積書', s: {size: 16, bold: true, fill: '#d9e1f2', align: 'center', borderTop: border, borderRight: border, borderBottom: border, borderLeft: border}},
    'テーマの色（アクセント1、明るさ80%）は #d9e1f2、結合セルの外枠は端のセルの罫線から取る');
  assert.deepEqual(s.cells.A3, {t: 's', v: '品番', s: {bold: true, color: '#ffffff', fill: '#217346', align: 'center', borderBottom: '1px solid #000000'}});
  assert.deepEqual(s.cells.A4, {t: 's', v: '00123', s: {numFmt: '@'}});
  assert.deepEqual(s.cells.A5, {t: 's', v: '2行の\n品名', s: {wrap: true, numFmt: '@'}}, 'セル内の改行は LF');
  assert.deepEqual(s.cells.C4, {t: 'n', v: 12, s: {numFmt: '#,##0'}});
  assert.deepEqual(s.cells.D4, {t: 'n', v: 30000, s: {numFmt: '¥#,##0'}}, '"¥"#,##0 はHTMLの ¥#,##0');
  assert.deepEqual(s.cells.A12, {t: 's', v: 'サイト', link: 'https://example.com/', s: {underline: true, color: '#0563c1'}});
  assert.deepEqual(w.sheets[1].cells.A2, {t: 's', v: '00123', s: {numFmt: '@'}});
  assert.deepEqual(w.sheets[1].cells.B2, {t: 'n', v: 1200});
});

test('Excelの列幅・行の高さ・枠の固定・枠線・倍率・選択・画像を、HTMLの大きさ（px）で移す', () => {
  const {workbook: w} = fixture(), s = w.sheets[0];
  // 列幅は「標準のフォントの数字1文字の幅」の単位。游ゴシック11ptは8px（96dpi）なので、xlsxの幅×8。
  // 幅を指定していない列は、Excelの既定の幅（数字8文字分と余白を、8pxの倍数に切り上げた72px）。
  assert.deepEqual(s.colWidths.slice(0, 9), [74, 150, 62, 86, 102, 102, 94, 72, 72], 'H列は非表示なので、既定の幅で表示する');
  assert.deepEqual([s.rowHeights[0], s.rowHeights[1], s.rowHeights[4], s.rowHeights[11]], [40, 25, 48, 42],
    '行の高さはpt÷0.75。指定のない行は、Excelと同じく標準のフォント（游ゴシック11pt）に合う18.75pt（ファイルの18ptは、表示の倍率125%で保存した値）');
  assert.deepEqual(s.freeze, {rows: 3, cols: 0});
  assert.equal(w.sheets[1].gridlines, false);
  assert.equal(w.activeSheet, 's1');
  assert.deepEqual(w.view, {zoom: 120, active: [4, 1], selection: {r1: 4, c1: 1, r2: 4, c2: 1}});
  assert.deepEqual(s.images, [{id: 'pic-1', asset: 'img-1', r: 13, c: 1, x: 8, y: 4, w: 120, h: 60}]);
  assert.match(w.assets['img-1'], /^data:image\/png;base64,iVBOR/);
  assert.ok(s.rows >= 16 && s.cols >= 8, '値のある範囲より少し広く取る');
});

test('数式は、HTMLで計算できれば数式のまま、できなければExcelが保存していた値にして知らせる', () => {
  const result = fixture(), s = result.workbook.sheets[0];
  assert.deepEqual(['E4', 'E5', 'E6', 'E7', 'E8', 'E9'].map(a => s.cells[a].f), ['=C4*D4', '=C5*D5', '=C6*D6', '=SUM(E4:E6)', '=ROUNDDOWN(E7*0.1,0)', '=E7+E8'],
    'E5・E6 は共有数式（E4の式をずらしたもの）');
  assert.deepEqual([s.cells.F4, s.cells.F5], [{t: 'n', v: 1200, s: {numFmt: '¥#,##0'}}, {t: 'n', v: 800}], '別のシートを参照する数式は値にする（VLOOKUP の表も、別のシートの列全体）');
  assert.ok(warned(result, /F4（=VLOOKUP\(A4,単価表!A:B,2,FALSE\)/));
  assert.ok(warned(result, /F5（=単価表!B3/));
  assert.deepEqual(result.notes['見積'].cells.F4, [{fields: ['hasFormula', 'formula'], reason: 'HTMLで計算できない数式は値にした'}]);
});

test('計算できるかは core.js に問い合わせて決めるので、core.js にある関数なら数式のまま残る', () => {
  const book = miniXlsx({sheet: '<sheetData><row r="1"><c r="A1"><f>ABS(-3)+MOD(7,4)</f><v>6</v></c><c r="B1"><f>SUMSQ(2,4)</f><v>20</v></c>' +
    '<c r="C1"><f>_xlfn.BITAND(6,3)</f><v>2</v></c><c r="D1"><f>SUM(A1:A600)</f><v>6</v></c><c r="E1"><f>SUMIF(A1:A1,"&gt;1")</f><v>6</v></c>' +
    '<c r="F1"><f>_xlfn.IFS(TRUE,1)</f><v>1</v></c></row></sheetData>'});
  const result = fromXlsx(book), cells = result.workbook.sheets[0].cells;
  assert.deepEqual([cells.A1.f, cells.E1.f, cells.F1.f], ['=ABS(-3)+MOD(7,4)', '=SUMIF(A1:A1,">1")', '=IFS(TRUE,1)']);
  assert.deepEqual([cells.B1, cells.C1, cells.D1], [{t: 'n', v: 20}, {t: 'n', v: 2}, {t: 'n', v: 6}], 'core.js にない関数と、上限の外を参照する式は値');
  assert.ok(warned(result, /SUMSQ/));
  assert.ok(warned(result, /C1（=BITAND\(6,3\)：HTMLで計算できない関数（BITAND）/), 'ファイルの中で関数名の前に付く _xlfn. は外して、Excelの画面と同じ名前で知らせる');
  assert.ok(warned(result, /上限（500行・100列）の外を参照/));
  assert.ok(!warned(result, /HTMLで計算した値が/), '数式のまま残した SUMIF・IFS も、Excelが保存していた値と同じになる');
});

// Excel saves the functions it added in 2010 and later with _xlfn. before their names (_xlfn.CONCAT(, _xlfn.TEXTJOIN(); a
// few have _xlfn._xlws. (_xlfn._xlws.SORT(). The page names them as Excel shows them, and computes the ones core.js knows.
test('Excelが関数名の前に _xlfn. を付けて保存した関数も、計算できれば数式のまま移す（共有数式も、文字列の中はそのまま）', () => {
  const sheet = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>' +
    '<c r="B1" t="str"><f t="shared" ref="B1:B2" si="0">_xlfn.CONCAT(A1,"_xlfn.")</f><v>a_xlfn.</v></c><c r="C1" t="str"><f>_xlfn.TEXTJOIN(",",TRUE,A1:A2)</f><v>a,b</v></c>' +
    '<c r="D1"><f>_xlfn._xlws.SORT(A1:A2)</f><v>0</v></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t>b</t></is></c><c r="B2" t="str"><f t="shared" si="0"/><v>b_xlfn.</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet})), cells = result.workbook.sheets[0].cells;
  assert.deepEqual([cells.B1.f, cells.B2.f, cells.C1.f], ['=CONCAT(A1,"_xlfn.")', '=CONCAT(A2,"_xlfn.")', '=TEXTJOIN(",",TRUE,A1:A2)']);
  assert.deepEqual(cells.D1, {t: 'n', v: 0});
  assert.ok(warned(result, /D1（=SORT\(A1:A2\)：HTMLで計算できない関数（SORT）/));
  assert.ok(!warned(result, /HTMLで計算した値が/), 'HTMLの計算も、Excelが保存していた値と同じ');
});

// tests/fixtures/excel-text.xlsx holds formulas that join text, made and saved inside desktop Excel 16.0 (build 20326) with
// the values Excel computed (tests/fixtures/make-excel-text.ps1). They are read here with tests/lib/xlsx.mjs, apart from the
// importer, and compared with what core.js computes on the book read from the file.
test('文字をつなぐ数式（&・CONCATENATE・CONCAT・TEXTJOIN）は、Excelで作ったファイルから数式のまま移り、HTMLの計算がExcelの保存した値と同じになる', () => {
  const bytes = readFileSync(join(SKILL_ROOT, 'tests', 'fixtures', 'excel-text.xlsx'));
  const result = fromXlsx(bytes, {name: 'excel-text.xlsx'}), w = result.workbook, s = w.sheets[0];
  assert.deepEqual(result.warnings, [], '値にした数式も、HTMLとExcelで値の違う数式もない');
  const saved = Object.entries(readXlsx(bytes).sheets[0].cells).filter(([, cell]) => cell.formula !== undefined);
  assert.ok(saved.length > 150 && saved.some(([a]) => a === 'N44'), 'ファイルの数式を読めている');
  assert.deepEqual(saved.filter(([a]) => s.cells[a]?.t !== 'f').map(([a]) => a), [], 'どの数式も数式のまま');
  const cache = new Map(), differ = [];
  for (const [a, cell] of saved) {
    const [r, c] = Core.coord(a), got = Core.calculate(w, s, r, c, new Set(), cache), want = cell.value ?? '';
    if (got !== want) differ.push(`${a} ${cell.formula}: HTML ${JSON.stringify(got)}、Excel ${JSON.stringify(want)}`);
  }
  assert.deepEqual(differ, [], 'HTMLの計算がExcelの値と違う数式');
  assert.equal(s.cells.I1.f, '=TEXTJOIN(",",D1:D4,"a","","b")');
});

test('HTMLにない表示形式・罫線・網かけや、移せない回転・非表示の列・条件付き書式などは、近いものにするか外して知らせる', () => {
  const result = fixture(), s = result.workbook.sheets[0];
  assert.deepEqual(s.cells.G4, {t: 'n', v: 46346, s: {numFmt: 'yyyy/mm/dd'}}, 'yyyy/m/d は yyyy/mm/dd');
  assert.deepEqual(s.cells.G5, {t: 'n', v: 46346, s: {numFmt: 'yyyy/mm/dd'}});
  assert.deepEqual(s.cells.G6, {t: 'n', v: 0.5208333333333334}, '時刻の表示形式はHTMLにないので、数値のまま');
  assert.deepEqual(s.cells.D6.s, {numFmt: '¥#,##0'}, '負の数を赤くする表示形式は ¥#,##0');
  assert.deepEqual(s.cells.C12.s, {borderBottom: '1px dashed #000000'}, '一点鎖線は破線');
  assert.deepEqual(s.cells.D12.s, {fill: '#ffbfbf'}, '網かけ（赤の25%）は、白と混ぜた1色');
  assert.deepEqual(s.cells.B10, {t: 's', v: '太字と普通', s: {bold: true}}, '一部だけ太字の文字は、セル全体の書式（太字）');
  for (const pattern of [/表示形式「yyyy\/m\/d」.*G4/, /表示形式「h:mm」.*G6/, /\[Red\].*D6/, /一点鎖線|dashDot/, /網かけ/, /回転/, /非表示の列（H）/,
    /条件付き書式/, /入力規則/, /コメント/, /文字ごとに違う書式/]) assert.ok(warned(result, pattern), String(pattern));
  assert.ok(!warned(result, /G5/), 'HTMLにある表示形式（yyyy/mm/dd）は知らせない');
  assert.ok(result.notes['見積'].sheet.some(n => n.item === 'colWidth' && n.cols.includes('H') && n.reason === '非表示の列'));
  assert.ok(warned(result, /シート「単価表」：倍率（100%）は、ブック全体の倍率（120%）にした/), 'HTMLの倍率はブック全体で1つ');
  assert.ok(result.notes['単価表'].sheet.some(n => n.item === 'zoom'));
});

test('テーマの色の明るさの調整は、Excelと同じ色にする（Excelが表示したテーマの12色×11段階の132組すべて）', () => {
  // Office 2013-2022 のテーマの色（lt1, dk1, lt2, dk2, アクセント1〜6, リンク, 表示済みのリンク）と、色の選択肢の11段階の明るさ。
  // 色は、その組み合わせで塗ったセルを、デスクトップ版Excel 16.0 で開いて読み取ったもの。
  const theme = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  const tints = [0.7999816888943144, 0.5999938962981048, 0.3999755851924192, 0.249977111117893, 0.1499984740745262, 0.0499893185216834,
    -0.0499893185216834, -0.1499984740745262, -0.249977111117893, -0.3499862666707358, -0.499984740745262];
  const shown = [
    'ffffff ffffff ffffff ffffff ffffff ffffff f2f2f2 d9d9d9 bfbfbf a6a6a6 808080', 'cccccc 999999 666666 404040 262626 0d0d0d 000000 000000 000000 000000 000000',
    'fafafa f4f4f4 f0f0f0 ececec eaeaea e8e8e8 dcdada c5c2c2 aeaaaa 989494 757171', 'd6dce4 acb9ca 8497b0 657c9c 576c88 4a5b73 404f64 394759 333f4f 2d3644 222b35',
    'd9e1f2 b4c6e7 8ea9db 7395d2 5f85cd 4b77c7 3c69bd 355faa 305496 284782 203764', 'fce4d6 f8cbad f4b084 f19d65 ef8f50 ee833c ec7424 de6514 c65911 aa4d0f 833c0c',
    'ededed dbdbdb c9c9c9 bbbbbb b1b1b1 a9a9a9 9c9c9c 8b8b8b 7b7b7b 6a6a6a 525252', 'fff2cc ffe699 ffd966 ffcf40 ffc926 ffc20d f2b600 d9a300 bf8f00 a67c00 806000',
    'ddebf7 bdd7ee 9bc2e6 84b4df 73a9db 629fd7 4e93d1 3685cb 2f75b5 28659b 1f4e78', 'e2efda c6e0b4 a9d08e 93c572 84bd5e 76b64b 6aa444 5f923d 548235 48702e 375623',
    'c1dffd 85c0fc 46a0fb 1a89f9 067aee 056acf 045eb7 0454a4 03498f 03407c 023160', 'ebdae2 d8b4c6 c58faa b67495 ab6186 9d5378 8c4a6b 7d4260 703a55 603349 492738'];
  const wrong = [];
  theme.forEach((base, i) => shown[i].split(' ').forEach((want, j) => { if (tint(base, tints[j]) !== want.toUpperCase()) wrong.push(`${base} ${tints[j]}: ${tint(base, tints[j])}（Excel ${want}）`); }));
  assert.deepEqual(wrong, []);
});

test('このスキルで書き出したxlsxを読み込むと、元の表に戻る', () => {
  const png = 'data:image/png;base64,iVBORw0KGgo=';
  const book = blankBook({activeSheet: 'b', view: {zoom: 110, active: [2, 1], selection: {r1: 2, c1: 1, r2: 3, c2: 2}}, assets: {logo: png}});
  const base = book.sheets[0];
  book.sheets = [
    {...base, id: 'a', name: '明細', colWidths: [80, 200, null, 64], rowHeights: [36, null, 44], freeze: {rows: 2, cols: 1}, merges: [{r1: 0, c1: 1, r2: 0, c2: 3}],
      images: [{id: 'p', asset: 'logo', r: 5, c: 2, x: 10, y: 6, w: 90, h: 40}], cells: {
        B1: {t: 's', v: '題名', s: {bold: true, size: 14, fill: '#e2efda', align: 'center', borderBottom: '2px solid #217346'}},
        A3: {t: 's', v: '00123', s: {numFmt: '@'}}, B3: {t: 's', v: '1行目\n2行目', s: {wrap: true, valign: 'top'}}, C3: {t: 'n', v: 1234.5, s: {numFmt: '#,##0.00', italic: true}},
        D3: {t: 'f', f: '=C3*2', s: {numFmt: '¥#,##0', borderTop: '0.5px solid #808080', borderLeft: '1px dotted #ff0000'}}, A4: {t: 'b', v: true},
        B4: {t: 'n', v: 46023, s: {numFmt: 'yyyy/mm/dd', align: 'right', indent: 2}}, C4: {t: 'n', v: 0.125, s: {numFmt: '0.00%', strike: true, color: '#c00000'}},
        D4: {t: 's', v: '3px double', s: {borderBottom: '3px double #000000', underline: true, font: 'Meiryo'}},
      }},
    {...base, id: 'b', name: '集計', gridlines: false, cells: {A1: {t: 'f', f: '=SUM(B1:B3)'}, B1: {t: 'n', v: 1}}},
  ];
  Core.validate(book);
  const back = fromXlsx(toXlsx(book).bytes).workbook;
  assert.deepEqual(back.sheets.map(s => s.name), ['明細', '集計']);
  assert.equal(back.activeSheet, 's2');
  assert.deepEqual(back.view, {zoom: 110, active: [2, 1], selection: {r1: 2, c1: 1, r2: 3, c2: 2}});
  const [a, b] = back.sheets, alias = {'Yu Gothic': '游ゴシック', 'Meiryo': 'メイリオ'};
  for (const key of Object.keys(book.sheets[0].cells)) {
    const want = Core.style(book, book.sheets[0], book.sheets[0].cells[key]), got = Core.style(back, a, a.cells[key]);
    assert.deepEqual({...got, font: alias[got.font] || got.font}, {...want, font: alias[want.font] || want.font}, key + ' の書式');
    const {s: _1, ...value} = book.sheets[0].cells[key], {s: _2, ...gotValue} = a.cells[key];
    assert.deepEqual(gotValue, value, key + ' の値');
  }
  assert.deepEqual(a.colWidths.slice(0, 5), [80, 200, 96, 64, 96], '指定のない列はHTMLの既定の96px');
  assert.deepEqual(a.rowHeights.slice(0, 4), [36, 22, 44, 22], '指定のない行はHTMLの既定の22px');
  assert.deepEqual([a.freeze, a.merges, b.gridlines], [{rows: 2, cols: 1}, [{r1: 0, c1: 1, r2: 0, c2: 3}], false]);
  assert.deepEqual(a.images.map(({asset, ...p}) => p), [{id: 'pic-1', r: 5, c: 2, x: 10, y: 6, w: 90, h: 40}]);
  assert.equal(back.assets[a.images[0].asset], png);
  assert.deepEqual(b.cells.A1, {t: 'f', f: '=SUM(B1:B3)'});
});

test('値のある範囲がHTMLの上限を超えるxlsxは読み込まずに止め、どのシートが何行・何列かを伝える', () => {
  const big = miniXlsx({sheet: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="501"><c r="CW501"><v>2</v></c></row></sheetData>'});
  assert.throws(() => fromXlsx(big), /シート「S」は 501 行・101 列/);
  const styledOnly = miniXlsx({sheet: '<sheetData><row r="2"><c r="A2"><v>5</v></c></row><row r="900"><c r="A900" s="1"/></row></sheetData>'});
  const result = fromXlsx(styledOnly);
  assert.equal(result.workbook.sheets[0].rows, 500, '書式だけのセルは上限で切る');
  assert.ok(warned(result, /書式だけのセル/));
});

test('ふりがな・改行・エラーの値・日付の書き方・1904年の日付も、Excelの表示どおりに読む', () => {
  const strings = '<si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="0"/></si><si><t xml:space="preserve">1行目\r\n2行目</t></si>' +
    '<si><t>_x005F_x0041_ と _x0001_</t></si>';
  const sheet = '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c>' +
    '<c r="D1" t="e"><v>#N/A</v></c><c r="E1" t="d" s="2"><v>2026-01-01</v></c><c r="F1" t="inlineStr"><is><t>直接の文字</t></is></c>' +
    '<c r="G1"><f>A9</f><v>0</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet, strings})), cells = result.workbook.sheets[0].cells;
  assert.deepEqual([cells.A1.v, cells.B1.v, cells.C1.v], ['漢字', '1行目\n2行目', '_x0041_ と \u0001'], 'ふりがなは文字に入れない');
  assert.deepEqual(cells.D1, {t: 's', v: '#N/A'}, 'エラーの値は、その文字');
  assert.ok(warned(result, /エラーの値/));
  assert.deepEqual(cells.E1, {t: 'n', v: 46023, s: {numFmt: 'yyyy/mm/dd'}}, '日付で書かれた値（t="d"）はシリアル値');
  assert.equal(cells.F1.v, '直接の文字');
  assert.deepEqual(cells.G1, {t: 'f', f: '=A9'});
  assert.ok(!warned(result, /HTMLで計算した値が/), '空のセルの参照は、HTMLでもExcelと同じく0');
  const date1904 = fromXlsx(miniXlsx({workbookPr: '<workbookPr date1904="1"/>', sheet: '<sheetData><row r="1"><c r="A1" s="2"><v>44561</v></c></row></sheetData>'}));
  assert.equal(date1904.workbook.sheets[0].cells.A1.v, 46023, '1904年から数えた日付は、1900年からのシリアル値にする');
});

test('xlsx でないファイルや、ワークシートのないファイルは、理由を示して止める', () => {
  assert.throws(() => fromXlsx(Buffer.from('not a zip at all')), /xlsx（ZIP）として読めません/);
});
