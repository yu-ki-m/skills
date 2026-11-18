// HTMLのブックの機能の1つ1つについて、xlsxへの書き出し（HTML→Excel）と、xlsxからの読み込み（Excel→HTML）を確かめる。
//
// 機能の一覧は、HTMLが実際に使っている一覧から読む。
// core.js の書式の項目（DEFAULT）・配置（CHOICES）・表示形式（FORMATS）・罫線（BORDER_XL）・日本語の名前があるフォント（FONT_ALIAS）・
// セルの型（CELL_TYPES）・数式の関数（FUNCTIONS）・ブックやシートなどの項目（FIELDS）と、リボンのフォントの一覧（assets/shell.html）である。
// ブックの検査（validate）はこれらの一覧にないものを受け付けず、数式は FUNCTIONS にない関数を計算しないので、
// HTMLに機能を足すときは、どれかの一覧にも足すことになる。
// そのとき、下の SAMPLES にその機能の見本がないと、最初のテストが失敗する（SKILL.md の「実装を変えるとき」）。
//
// 見本は、HTMLのブック（book）と、書き出したxlsxの中でその機能がExcelの書き方になっていること（excel）と、
// そのxlsxを読み込み直した結果（back）を確かめる。
// back を省いた見本は、読み込み直したセルの値と、画面に出る書式が、元と同じであることを確かめる。
// Excelに移さない機能は、xlsxに書かないことと、読み込むとどうなるかを確かめ、その理由をコメントに書く。
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {test} from 'node:test';
import {blankBook, Core, SKILL_ROOT} from './lib/workbook.mjs';
import {kid, kids, parseXml, readXlsx} from './lib/xlsx.mjs';

const require = createRequire(import.meta.url);
const {toXlsx} = require(join(SKILL_ROOT, 'scripts', 'xlsx.cjs'));
const {fromXlsx} = require(join(SKILL_ROOT, 'scripts', 'read_xlsx.cjs'));
// The fonts the page's ribbon offers (assets/shell.html).
const ribbonFonts = [...readFileSync(join(SKILL_ROOT, 'assets', 'shell.html'), 'utf8').match(/<select id="font"[^>]*>(.*?)<\/select>/s)[1]
  .matchAll(/<option>([^<]+)<\/option>/g)].map(m => m[1]);

/** Every feature of the page, named as the samples' covers name them. */
function features() {
  const fonts = [...new Set([...Object.values(Core.FONT_ALIAS), ...ribbonFonts])];
  return [
    ...Object.keys(Core.DEFAULT).map(k => 'style.' + k),
    ...Object.entries(Core.CHOICES).flatMap(([key, values]) => values.map(v => `${key}.${v}`)),
    ...Core.FORMATS.map(f => 'numFmt.' + f),
    ...Object.keys(Core.BORDER_XL).map(b => 'border.' + b),
    ...fonts.map(f => 'font.' + f),
    ...Core.CELL_TYPES.map(t => 'type.' + t),
    ...Core.FUNCTIONS.map(f => 'function.' + f),
    ...Object.entries(Core.FIELDS).flatMap(([kind, list]) => list.map(f => `${kind}.${f}`)),
  ];
}

const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const EMU = 9525; // a CSS pixel (0.75pt) in EMU
/** A book whose one sheet holds cells; sheet adds to the sheet, and the rest to the book. */
const bookOf = (cells, {sheet = {}, ...rest} = {}) => blankBook({sheet: {cells, ...sheet}, ...rest});
/** The cell a of the written file's first (or given) sheet, as tests/lib/xlsx.mjs reads it. */
const cellOf = (x, a, sheet = 0) => x.sheets[sheet].cells[a];
/** The attributes of a written cell's format (its <xf> in cellXfs). */
const xfOf = (x, cell) => kids(kid(parseXml(x.parts.get('xl/styles.xml').toString('utf8')), 'cellXfs'), 'xf')[cell.styleIndex].attrs;
const sheetView = sheet => kid(kid(sheet.root, 'sheetViews'), 'sheetView');

/** The cells read back hold the values written and show the same format (the book's, the sheet's and the cell's together). */
function sameCells(back, w) {
  w.sheets.forEach((s, i) => {
    const b = back.sheets[i];
    for (const [a, cell] of Object.entries(s.cells)) {
      const {s: _written, ...value} = cell, {s: _read, ...got} = b.cells[a] || {t: 's', v: ''};
      assert.deepEqual(got, value, `${s.name}!${a} の値`);
      assert.deepEqual(Core.style(back, b, b.cells[a] || {}), Core.style(w, s, cell), `${s.name}!${a} の書式`);
    }
  });
}
/** As sameCells, and every cell of each sheet's table shows the same format, whether it holds a value or not. */
function sameTable(back, w) {
  w.sheets.forEach((s, i) => {
    const b = back.sheets[i];
    assert.ok(b.rows >= s.rows && b.cols >= s.cols, `${s.name} の表の大きさ`);
    Core.each({r1: 0, c1: 0, r2: s.rows - 1, c2: s.cols - 1}, (r, c) => {
      const a = Core.addr(r, c);
      assert.deepEqual(Core.style(back, b, b.cells[a] || {}), Core.style(w, s, s.cells[a] || {}), `${s.name}!${a} の書式`);
    });
  });
  sameCells(back, w);
}

// How Japanese Excel names each font, and its character set (128 for a Japanese font).
const FONT_IN_EXCEL = {'Yu Gothic': ['游ゴシック', 128], 'Yu Gothic UI': ['Yu Gothic UI', 128], 'Yu Mincho': ['游明朝', 128], 'Meiryo': ['メイリオ', 128],
  'MS Gothic': ['ＭＳ ゴシック', 128], 'MS PGothic': ['ＭＳ Ｐゴシック', 128], 'MS Mincho': ['ＭＳ 明朝', 128], 'MS PMincho': ['ＭＳ Ｐ明朝', 128],
  'Aptos': ['Aptos', undefined], 'Calibri': ['Calibri', undefined], 'Arial': ['Arial', undefined]};
// The alignment Excel stores for each of the page's (no attribute reads as general and bottom).
const HORIZONTAL_IN_EXCEL = {general: 'general', left: 'left', center: 'center', right: 'right'};
const VERTICAL_IN_EXCEL = {top: 'top', middle: 'center', bottom: 'bottom'};
// Each number format: Excel's built-in format of that number, or the format code the file defines (a string); and a value.
const FORMAT_IN_EXCEL = {'General': [0, 1234.5], '@': [49, '00123'], '0': [1, 1234.5], '0.00': [2, 1234.5], '#,##0': [3, 1234.5], '#,##0.00': [4, 1234.5],
  '0%': [9, 0.125], '0.00%': [10, 0.125], '¥#,##0': ['"¥"#,##0', 30000], 'yyyy/mm/dd': ['yyyy/mm/dd', 46023]};
const BORDER_IN_EXCEL = {'0.5px solid': 'hair', '1px solid': 'thin', '2px solid': 'medium', '3px solid': 'thick', '1px dashed': 'dashed',
  '2px dashed': 'mediumDashed', '1px dotted': 'dotted', '3px double': 'double'};
// A formula with each function, and its value (the page's, which desktop Excel 16.0 gives too: each was computed there on
// the cells of FUNCTION_DATA); a third item is the formula as Excel stores it in the file where it differs (Excel saves the
// functions it added in 2010 and later, such as CONCAT and TEXTJOIN of Excel 2016, with _xlfn. before their names).
const FUNCTION_DATA = {B1: 1, B2: 2, B3: 3, C1: 10, C2: 20, C3: 30, E1: 'a', F1: 'b', G1: 'c', E2: 1, F2: 2, G2: 3};
const dataCells = () => Object.fromEntries(Object.entries(FUNCTION_DATA).map(([a, v]) => [a, typeof v === 'number' ? {t: 'n', v} : {t: 's', v}]));
const FUNCTION_IN_EXCEL = {IF: ['=IF(TRUE,1,2)', 1], IFS: ['=IFS(1>2,"a",2>1,"b")', 'b', '_xlfn.IFS(1>2,"a",2>1,"b")'],
  IFERROR: ['=IFERROR(1/0,5)', 5], IFNA: ['=IFNA(NA(),"なし")', 'なし', '_xlfn.IFNA(NA(),"なし")'], AND: ['=AND(TRUE,FALSE)', false],
  OR: ['=OR(FALSE,TRUE)', true], XOR: ['=XOR(TRUE,FALSE,TRUE)', false, '_xlfn.XOR(TRUE,FALSE,TRUE)'], NOT: ['=NOT(FALSE)', true],
  TRUE: ['=TRUE()', true], FALSE: ['=FALSE()', false], NA: ['=ISNA(NA())', true],
  SWITCH: ['=SWITCH(2,1,"一",2,"二","他")', '二', '_xlfn.SWITCH(2,1,"一",2,"二","他")'], CHOOSE: ['=CHOOSE(2,"a","b","c")', 'b'],
  ISBLANK: ['=ISBLANK(D1)', true], ISNUMBER: ['=ISNUMBER(1.5)', true], ISTEXT: ['=ISTEXT("a")', true], ISNONTEXT: ['=ISNONTEXT(1)', true],
  ISLOGICAL: ['=ISLOGICAL(FALSE)', true], ISERROR: ['=ISERROR(1/0)', true], ISERR: ['=ISERR(NA())', false], ISNA: ['=ISNA(1)', false],
  ISEVEN: ['=ISEVEN(-3)', false], ISODD: ['=ISODD(2.5)', false], N: ['=N(TRUE)', 1], T: ['=T("a")&T(1)', 'a'], SUM: ['=SUM(1,2,3)', 6],
  PRODUCT: ['=PRODUCT(2,3,4)', 24], ROUND: ['=ROUND(1.25,1)', 1.3], ROUNDUP: ['=ROUNDUP(1.21,1)', 1.3], ROUNDDOWN: ['=ROUNDDOWN(1.29,1)', 1.2],
  TRUNC: ['=TRUNC(-1.59,1)', -1.5], INT: ['=INT(-1.5)', -2], ABS: ['=ABS(-3)', 3], MOD: ['=MOD(7,4)', 3], QUOTIENT: ['=QUOTIENT(7,2)', 3],
  SIGN: ['=SIGN(-2)', -1], SQRT: ['=SQRT(16)', 4], POWER: ['=POWER(2,10)', 1024], EXP: ['=EXP(0)', 1], LN: ['=LN(1)', 0], 'LOG10': ['=LOG10(1000)', 3],
  LOG: ['=LOG(8,2)', 3], PI: ['=ROUND(PI(),4)', 3.1416], CEILING: ['=CEILING(2.1,0.5)', 2.5], FLOOR: ['=FLOOR(2.9,0.5)', 2.5],
  'CEILING.MATH': ['=CEILING.MATH(-2.5)', -2, '_xlfn.CEILING.MATH(-2.5)'], 'FLOOR.MATH': ['=FLOOR.MATH(-2.5)', -3, '_xlfn.FLOOR.MATH(-2.5)'],
  MROUND: ['=MROUND(10,3)', 9], SUMPRODUCT: ['=SUMPRODUCT(B1:B3,C1:C3)', 140], AVERAGE: ['=AVERAGE(2,4)', 3], MIN: ['=MIN(5,2,9)', 2],
  MAX: ['=MAX(5,2,9)', 9], COUNT: ['=COUNT(1,2,"a")', 2], COUNTA: ['=COUNTA(1,"a")', 2], COUNTBLANK: ['=COUNTBLANK(B1:D3)', 3],
  MEDIAN: ['=MEDIAN(3,1,2,10)', 2.5], MODE: ['=MODE(1,2,2,3)', 2], 'MODE.SNGL': ['=MODE.SNGL(1,2,2,3)', 2, '_xlfn.MODE.SNGL(1,2,2,3)'],
  LARGE: ['=LARGE(B1:B3,2)', 2], SMALL: ['=SMALL(B1:B3,1)', 1], RANK: ['=RANK(2,B1:B3)', 2],
  'RANK.EQ': ['=RANK.EQ(3,B1:B3,1)', 3, '_xlfn.RANK.EQ(3,B1:B3,1)'], STDEV: ['=STDEV(2,4,6)', 2],
  'STDEV.S': ['=STDEV.S(2,4,6)', 2, '_xlfn.STDEV.S(2,4,6)'], STDEVP: ['=STDEVP(1,3)', 1], 'STDEV.P': ['=STDEV.P(1,3)', 1, '_xlfn.STDEV.P(1,3)'],
  VAR: ['=VAR(2,4,6)', 4], 'VAR.S': ['=VAR.S(2,4,6)', 4, '_xlfn.VAR.S(2,4,6)'], VARP: ['=VARP(1,3)', 1],
  'VAR.P': ['=VAR.P(1,3)', 1, '_xlfn.VAR.P(1,3)'], SUBTOTAL: ['=SUBTOTAL(9,B1:B3)', 6], COUNTIF: ['=COUNTIF(B1:B3,">1")', 2],
  COUNTIFS: ['=COUNTIFS(B1:B3,">1",C1:C3,"<30")', 1], SUMIF: ['=SUMIF(B1:B3,">1",C1:C3)', 50], AVERAGEIF: ['=AVERAGEIF(B1:B3,">1",C1:C3)', 25],
  SUMIFS: ['=SUMIFS(C1:C3,B1:B3,">=2")', 50], AVERAGEIFS: ['=AVERAGEIFS(C1:C3,B1:B3,"<3")', 15],
  MAXIFS: ['=MAXIFS(C1:C3,B1:B3,"<3")', 20, '_xlfn.MAXIFS(C1:C3,B1:B3,"<3")'],
  MINIFS: ['=MINIFS(C1:C3,B1:B3,">1")', 20, '_xlfn.MINIFS(C1:C3,B1:B3,">1")'], VLOOKUP: ['=VLOOKUP(2,B1:C3,2,FALSE)', 20],
  HLOOKUP: ['=HLOOKUP("b",E1:G2,2,FALSE)', 2], MATCH: ['=MATCH(3,B1:B3,0)', 3], INDEX: ['=INDEX(C1:C3,2)', 20],
  XLOOKUP: ['=XLOOKUP(3,B1:B3,C1:C3)', 30, '_xlfn.XLOOKUP(3,B1:B3,C1:C3)'], XMATCH: ['=XMATCH(2,B1:B3)', 2, '_xlfn.XMATCH(2,B1:B3)'],
  LOOKUP: ['=LOOKUP(2.5,B1:B3,C1:C3)', 20], ROW: ['=ROW(C5)', 5], COLUMN: ['=COLUMN(C5)', 3], ROWS: ['=ROWS(B1:C4)', 4],
  COLUMNS: ['=COLUMNS(B1:C4)', 2], LEN: ['=LEN("Excel 関数")', 8], LEFT: ['=LEFT("Excel",2)', 'Ex'], RIGHT: ['=RIGHT("Excel",3)', 'cel'],
  MID: ['=MID("Excel",2,3)', 'xce'], FIND: ['=FIND("c","Excel")', 3], SEARCH: ['=SEARCH("C?L","Excel")', 3],
  SUBSTITUTE: ['=SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'], REPLACE: ['=REPLACE("abcdef",2,3,"X")', 'aXef'], TRIM: ['=TRIM("  a   b  ")', 'a b'],
  CLEAN: ['=CLEAN("a"&CHAR(10)&"b")', 'ab'], UPPER: ['=UPPER("abc")', 'ABC'], LOWER: ['=LOWER("ABC")', 'abc'],
  PROPER: ['=PROPER("hello world")', 'Hello World'], REPT: ['=REPT("ab",3)', 'ababab'], EXACT: ['=EXACT("a","A")', false], CHAR: ['=CHAR(65)', 'A'],
  CODE: ['=CODE("A")', 65], UNICHAR: ['=UNICHAR(9731)', '☃', '_xlfn.UNICHAR(9731)'], UNICODE: ['=UNICODE("あ")', 12354, '_xlfn.UNICODE("あ")'],
  ASC: ['=ASC("ＡＢＣ")', 'ABC'], DBCS: ['=DBCS("ABC")', 'ＡＢＣ'], VALUE: ['=VALUE("1,234")', 1234], TEXT: ['=TEXT(46024,"yyyy年m月d日")', '2026年1月2日'],
  CONCATENATE: ['=CONCATENATE("a",1/4,TRUE)', 'a0.25TRUE'], CONCAT: ['=CONCAT("a",1/3)', 'a0.333333333333333', '_xlfn.CONCAT("a",1/3)'],
  TEXTJOIN: ['=TEXTJOIN("-",TRUE,"a","","b")', 'a-b', '_xlfn.TEXTJOIN("-",TRUE,"a","","b")'], DATE: ['=DATE(2026,1,2)', 46024],
  YEAR: ['=YEAR(46024)', 2026], MONTH: ['=MONTH(46024)', 1], DAY: ['=DAY(46024)', 2], WEEKDAY: ['=WEEKDAY(46024)', 6], WEEKNUM: ['=WEEKNUM(46024)', 1],
  ISOWEEKNUM: ['=ISOWEEKNUM(46024)', 1, '_xlfn.ISOWEEKNUM(46024)'], EDATE: ['=EDATE(46053,1)', 46081], EOMONTH: ['=EOMONTH(46024,0)', 46053],
  DATEDIF: ['=DATEDIF(DATE(2020,1,15),DATE(2026,1,14),"Y")', 5], DAYS: ['=DAYS(46053,46024)', 29, '_xlfn.DAYS(46053,46024)'],
  NETWORKDAYS: ['=NETWORKDAYS(46024,46031)', 6], 'NETWORKDAYS.INTL': ['=NETWORKDAYS.INTL(46024,46031,11)', 7], WORKDAY: ['=WORKDAY(46024,1)', 46027],
  'WORKDAY.INTL': ['=WORKDAY.INTL(46024,1,11)', 46025], TIME: ['=TIME(12,0,0)', 0.5], HOUR: ['=HOUR(0.75)', 18], MINUTE: ['=MINUTE("12:34")', 34],
  SECOND: ['=SECOND("12:34:56")', 56], DATEVALUE: ['=DATEVALUE("2026/1/2")', 46024], TIMEVALUE: ['=TIMEVALUE("18:00")', 0.75],
  TODAY: ['=TODAY()>DATE(2000,1,1)', true], NOW: ['=NOW()-TODAY()<1', true]};
const TYPE_IN_EXCEL = {boolean: 'b', number: 'n', string: 'str'};

/** A font flag: the cell with it, and a cell without it. */
const fontFlag = (key, name, want) => ({
  name, covers: ['style.' + key],
  book: () => bookOf({A1: {t: 's', v: name, s: {[key]: true}}, A2: {t: 's', v: 'なし'}}),
  excel: x => assert.deepEqual([cellOf(x, 'A1').style.font[key], cellOf(x, 'A2').style.font[key]], want),
});
const valueOf = (x, a) => ({type: cellOf(x, a).type, value: cellOf(x, a).value});

const SAMPLES = [
  {name: 'フォント', covers: ['style.font', ...Object.keys(FONT_IN_EXCEL).map(f => 'font.' + f)],
    book: () => bookOf(Object.fromEntries(Object.keys(FONT_IN_EXCEL).map((font, i) => ['A' + (i + 1), {t: 's', v: font, s: {font}}]))),
    excel: x => Object.entries(FONT_IN_EXCEL).forEach(([font, want], i) => { const f = cellOf(x, 'A' + (i + 1)).style.font; assert.deepEqual([f.name, f.charset], want, font); })},
  {name: '文字の大きさ', covers: ['style.size'],
    book: () => bookOf({A1: {t: 's', v: '6pt', s: {size: 6}}, A2: {t: 's', v: '14pt', s: {size: 14}}, A3: {t: 's', v: '96pt', s: {size: 96}}}),
    excel: x => assert.deepEqual(['A1', 'A2', 'A3'].map(a => cellOf(x, a).style.font.size), [6, 14, 96])},
  fontFlag('bold', '太字', [true, false]), fontFlag('italic', '斜体', [true, false]), fontFlag('strike', '取り消し線', [true, false]),
  fontFlag('underline', '下線', ['single', 'none']),
  {name: '文字の色', covers: ['style.color'],
    book: () => bookOf({A1: {t: 's', v: '赤', s: {color: '#c00000'}}, A2: {t: 's', v: '黒'}}),
    excel: x => assert.deepEqual([cellOf(x, 'A1').style.font.color, cellOf(x, 'A2').style.font.color], ['#c00000', '#000000'])},
  {name: '塗り', covers: ['style.fill'],
    book: () => bookOf({A1: {t: 's', v: '緑', s: {fill: '#e2efda'}}, A2: {t: 's', v: '塗りなし', s: {fill: 'none'}}}),
    excel: x => assert.deepEqual([cellOf(x, 'A1').style.fill, cellOf(x, 'A2').style.fill], ['#e2efda', 'none'])},
  {name: '横位置', covers: ['style.align', ...Object.keys(HORIZONTAL_IN_EXCEL).map(v => 'align.' + v)],
    book: () => bookOf(Object.fromEntries(Object.keys(HORIZONTAL_IN_EXCEL).map((align, i) => ['A' + (i + 1), {t: 's', v: align, s: {align}}]))),
    excel: x => Object.values(HORIZONTAL_IN_EXCEL).forEach((want, i) => assert.equal(cellOf(x, 'A' + (i + 1)).style.alignment.horizontal, want, want))},
  {name: '縦位置', covers: ['style.valign', ...Object.keys(VERTICAL_IN_EXCEL).map(v => 'valign.' + v)],
    book: () => bookOf(Object.fromEntries(Object.keys(VERTICAL_IN_EXCEL).map((valign, i) => ['A' + (i + 1), {t: 's', v: valign, s: {valign}}]))),
    excel: x => Object.values(VERTICAL_IN_EXCEL).forEach((want, i) => assert.equal(cellOf(x, 'A' + (i + 1)).style.alignment.vertical, want, want))},
  {name: '折り返し', covers: ['style.wrap'],
    book: () => bookOf({A1: {t: 's', v: '1行目\n2行目', s: {wrap: true}}, A2: {t: 's', v: '折り返さない'}}),
    excel: x => assert.deepEqual([cellOf(x, 'A1').style.alignment.wrapText, cellOf(x, 'A2').style.alignment.wrapText], [true, false])},
  {name: 'インデント', covers: ['style.indent'],
    book: () => bookOf({A1: {t: 's', v: '左', s: {align: 'left', indent: 2}}, A2: {t: 'n', v: 5, s: {align: 'right', indent: 1}}}),
    excel: x => assert.deepEqual(['A1', 'A2'].map(a => [cellOf(x, a).style.alignment.horizontal, cellOf(x, a).style.alignment.indent]), [['left', 2], ['right', 1]])},
  // Indenting a cell aligned "general" in Excel aligns it left, so such a cell is written, and read back, as aligned left
  // (the page's screen still shows a number in it aligned right: a todo test in core.test.cjs).
  {name: 'インデント（標準の横位置）', covers: ['style.indent'],
    book: () => bookOf({A1: {t: 'n', v: 5, s: {indent: 2}}}),
    excel: x => assert.deepEqual(cellOf(x, 'A1').style.alignment, {horizontal: 'left', vertical: 'center', wrapText: false, indent: 2}),
    back: w => assert.deepEqual(w.sheets[0].cells.A1, {t: 'n', v: 5, s: {align: 'left', indent: 2}})},
  {name: '表示形式', covers: ['style.numFmt', ...Object.keys(FORMAT_IN_EXCEL).map(f => 'numFmt.' + f)],
    book: () => bookOf(Object.fromEntries(Object.entries(FORMAT_IN_EXCEL).map(([numFmt, [, v]], i) => ['A' + (i + 1), {t: typeof v === 'string' ? 's' : 'n', v, s: {numFmt}}]))),
    excel: x => Object.entries(FORMAT_IN_EXCEL).forEach(([numFmt, [want]], i) => {
      const id = +xfOf(x, cellOf(x, 'A' + (i + 1))).numFmtId;
      if (typeof want === 'number') assert.equal(id, want, numFmt + ' はExcelの組み込みの形式');
      else assert.deepEqual([id >= 164, x.styles.formats[id]], [true, want], numFmt + ' はファイルで定義する形式');
    })},
  {name: '罫線', covers: ['style.borderTop', 'style.borderRight', 'style.borderBottom', 'style.borderLeft', ...Object.keys(BORDER_IN_EXCEL).map(b => 'border.' + b)],
    book: () => bookOf({
      ...Object.fromEntries(Object.keys(BORDER_IN_EXCEL).map((kind, i) => ['B' + (2 * i + 2), {t: 's', v: kind, s: {borderTop: `${kind} #12ab${String(i).padStart(2, '0')}`}}])),
      D2: {t: 's', v: '4辺', s: {borderTop: '1px solid #ff0000', borderRight: '2px solid #00ff00', borderBottom: '3px double #0000ff', borderLeft: '1px dotted #000000'}}}),
    excel: x => {
      Object.values(BORDER_IN_EXCEL).forEach((style, i) => assert.deepEqual(cellOf(x, 'B' + (2 * i + 2)).style.borders.top, {style, color: '#12ab' + String(i).padStart(2, '0')}, style));
      assert.deepEqual(cellOf(x, 'D2').style.borders, {top: {style: 'thin', color: '#ff0000'}, right: {style: 'medium', color: '#00ff00'},
        bottom: {style: 'double', color: '#0000ff'}, left: {style: 'dotted', color: '#000000'}});
    }},
  {name: '文字列', covers: ['type.s', 'cell.t', 'cell.v', 'sheet.cells'],
    book: () => bookOf({A1: {t: 's', v: '文字'}, A2: {t: 's', v: '00123'}}),
    excel: x => assert.deepEqual([valueOf(x, 'A1'), valueOf(x, 'A2')], [{type: 's', value: '文字'}, {type: 's', value: '00123'}])},
  {name: '数値', covers: ['type.n'],
    book: () => bookOf({A1: {t: 'n', v: 1234.5}, A2: {t: 'n', v: -0.001}}),
    excel: x => assert.deepEqual([valueOf(x, 'A1'), valueOf(x, 'A2')], [{type: 'n', value: 1234.5}, {type: 'n', value: -0.001}])},
  {name: '真偽値', covers: ['type.b'],
    book: () => bookOf({A1: {t: 'b', v: true}, A2: {t: 'b', v: false}}),
    excel: x => assert.deepEqual([valueOf(x, 'A1'), valueOf(x, 'A2')], [{type: 'b', value: true}, {type: 'b', value: false}])},
  {name: '数式', covers: ['type.f', 'cell.f'],
    book: () => bookOf({A1: {t: 'n', v: 2}, A2: {t: 'f', f: '=A1*3'}}),
    excel: x => assert.deepEqual({formula: cellOf(x, 'A2').formula, ...valueOf(x, 'A2')}, {formula: 'A1*3', type: 'n', value: 6})},
  {name: '書式を重ねたセル', covers: ['cell.s'],
    book: () => bookOf({B2: {t: 'n', v: 1200, s: {bold: true, color: '#ffffff', fill: '#217346', align: 'center', numFmt: '#,##0', borderBottom: '2px solid #000000'}}}),
    excel: x => { const st = cellOf(x, 'B2').style;
      assert.deepEqual([st.font.bold, st.font.color, st.fill, st.alignment.horizontal, st.numFmt, st.borders.bottom.style], [true, '#ffffff', '#217346', 'center', '#,##0', 'medium']); }},
  {name: 'リンク', covers: ['cell.link'],
    book: () => bookOf({A1: {t: 's', v: 'サイト', link: 'https://example.com/a?b=1&c=2'}}),
    excel: x => { const s = x.sheets[0], st = s.cells.A1.style;
      assert.deepEqual([s.links.A1, st.font.underline, st.font.color, st.cellStyle], ['https://example.com/a?b=1&c=2', 'single', '#0563c1', 8]); },
    // In Excel a hyperlink is the Hyperlink cell style's blue underlined text, so the text read back has that format.
    back: w => assert.deepEqual(w.sheets[0].cells.A1, {t: 's', v: 'サイト', link: 'https://example.com/a?b=1&c=2', s: {underline: true, color: '#0563c1'}})},
  ...Object.entries(FUNCTION_IN_EXCEL).map(([name, [f, value, stored = f.slice(1)]]) => ({
    name: '関数 ' + name, covers: ['function.' + name],
    book: () => bookOf({A1: {t: 'f', f}, ...dataCells()}),
    excel: x => assert.deepEqual({formula: cellOf(x, 'A1').formula, ...valueOf(x, 'A1')}, {formula: stored, type: TYPE_IN_EXCEL[typeof value], value}),
  })),
  // The Japanese Excel names DBCS JIS, and the page takes either name; the file has DBCS, the name every Excel reads.
  {name: '関数 JIS（DBCS の日本語の名前）', covers: [],
    book: () => bookOf({A1: {t: 'f', f: '=JIS("ABC")&jis("1")'}}),
    excel: x => assert.deepEqual({formula: cellOf(x, 'A1').formula, ...valueOf(x, 'A1')}, {formula: 'DBCS("ABC")&DBCS("1")', type: 'str', value: 'ＡＢＣ１'}),
    back: w => assert.deepEqual(w.sheets[0].cells.A1, {t: 'f', f: '=DBCS("ABC")&DBCS("1")'})},
  // The ids and the revision let the page match an edit state to its book; Excel has nothing like them. They are not
  // written, and a book read from xlsx gets new ones (from the file's name and contents).
  {name: 'ブックとシートと画像のID（xlsxに書かない）', covers: ['book.schema', 'book.id', 'book.revision', 'sheet.id', 'image.id', 'image.asset'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {id: 'book-id-q7', revision: 'revision-q7', activeSheet: 'sheet-id-q7', assets: {'asset-q7': PNG},
      sheet: {id: 'sheet-id-q7', images: [{id: 'picture-q7', asset: 'asset-q7', r: 2, c: 1, x: 0, y: 0, w: 40, h: 30}]}}),
    excel: x => { for (const [name, data] of x.parts) assert.ok(!data.toString('latin1').includes('q7'), name + ' にIDがない'); },
    back: w => assert.deepEqual([w.schema, /^xlsx-book-[0-9a-f]{8}$/.test(w.id), w.revision, w.sheets[0].id, w.sheets[0].images[0].id, Object.keys(w.assets)],
      [Core.SCHEMA, true, 'draft', 's1', 'pic-1', ['img-1']])},
  {name: 'ブックの題名', covers: ['book.title'],
    book: () => bookOf({}, {title: '見積 & "明細" <2026>'}),
    excel: x => assert.equal(x.title, '見積 & "明細" <2026>'),
    back: w => assert.equal(w.title, '見積 & "明細" <2026>')},
  {name: 'シートの順と名前', covers: ['book.sheets', 'sheet.name'],
    book: () => { const w = bookOf({A1: {t: 's', v: '値'}}), base = w.sheets[0];
      w.sheets = [{...base, id: 'a', name: '売上 & 原価'}, {...base, id: 'b', name: '集計<2026>'}, {...base, id: 'c', name: "O'Brien"}]; w.activeSheet = 'a'; return w; },
    excel: x => assert.deepEqual(x.sheets.map(s => s.name), ['売上 & 原価', '集計<2026>', "O'Brien"]),
    back: (w, original) => { assert.deepEqual(w.sheets.map(s => s.name), ['売上 & 原価', '集計<2026>', "O'Brien"]); sameCells(w, original); }},
  {name: '最初に開くシート', covers: ['book.activeSheet'],
    book: () => { const w = bookOf({}), base = w.sheets[0]; w.sheets = [{...base, id: 'a', name: 'A'}, {...base, id: 'b', name: 'B'}]; w.activeSheet = 'b'; return w; },
    excel: x => assert.deepEqual([x.activeTab, x.sheets.map(s => s.view.tabSelected)], [1, [false, true]]),
    back: w => assert.equal(w.activeSheet, w.sheets[1].id)},
  // The page shows every cell of the table in the default format (defaultStyle), so every cell of the table is written in
  // it; the table read back is at least as large, and every cell of it shows the same format.
  {name: 'ブックの既定の書式', covers: ['book.defaultStyle'],
    book: () => bookOf({A1: {t: 's', v: '値'}}, {defaultStyle: {font: 'Meiryo', size: 10, fill: '#fff2cc'}, sheet: {rows: 3, cols: 2}}),
    excel: x => { for (const a of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3']) { const st = cellOf(x, a).style; assert.deepEqual([st.font.name, st.font.size, st.fill], ['メイリオ', 10, '#fff2cc'], a); } },
    back: sameTable},
  {name: 'シートの既定の書式', covers: ['sheet.defaultStyle'],
    book: () => bookOf({A1: {t: 's', v: '値'}}, {sheet: {rows: 2, cols: 2, defaultStyle: {borderBottom: '1px solid #000000', valign: 'top'}}}),
    excel: x => { for (const a of ['A1', 'B1', 'A2', 'B2']) { const st = cellOf(x, a).style; assert.deepEqual([st.borders.bottom.style, st.alignment.vertical], ['thin', 'top'], a); } },
    back: sameTable},
  // Excel has no table size. The file says which cells hold something (dimension), and the table read back is a little
  // larger than what it holds: 30 rows, or 20 rows past the last; 10 columns, or 3 columns past the last.
  {name: '表の大きさ', covers: ['sheet.rows', 'sheet.cols'],
    book: () => bookOf({A1: {t: 's', v: '左上'}, C40: {t: 's', v: '右下'}}, {sheet: {rows: 45, cols: 12}}),
    excel: x => assert.equal(x.sheets[0].dimension, 'A1:C40'),
    back: (w, original) => { assert.deepEqual([w.sheets[0].rows, w.sheets[0].cols], [60, 10]); sameCells(w, original); }},
  {name: '結合', covers: ['sheet.merges', 'range.r1', 'range.c1', 'range.r2', 'range.c2'],
    book: () => bookOf({B1: {t: 's', v: '題名', s: {bold: true, align: 'center', fill: '#e2efda'}}}, {sheet: {merges: [{r1: 0, c1: 1, r2: 1, c2: 3}]}}),
    excel: x => assert.deepEqual(x.sheets[0].merges, ['B1:D2']),
    back: (w, original) => { assert.deepEqual(w.sheets[0].merges, [{r1: 0, c1: 1, r2: 1, c2: 3}]); sameCells(w, original); }},
  {name: '列幅', covers: ['sheet.colWidths'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {sheet: {cols: 5, colWidths: [80, 200, null, 64]}}),
    excel: x => assert.deepEqual([0, 1, 2, 3, 4].map(c => x.sheets[0].widthOf(c)), [10, 25, 12, 8, 12], 'HTMLのpx÷8（指定のない列はHTMLの既定の96px）'),
    back: w => assert.deepEqual(w.sheets[0].colWidths.slice(0, 5), [80, 200, 96, 64, 96])},
  {name: '行の高さ', covers: ['sheet.rowHeights'],
    book: () => bookOf({A1: {t: 's', v: 'a'}, A2: {t: 's', v: 'b'}, A3: {t: 's', v: 'c'}}, {sheet: {rowHeights: [36, null, 44]}}),
    excel: x => assert.deepEqual([1, 2, 3].map(r => x.sheets[0].rows[r]), [{ht: 27, customHeight: true}, {ht: 16.5, customHeight: true}, {ht: 33, customHeight: true}],
      'HTMLのpx×0.75（指定のない行はHTMLの既定の22px）'),
    back: w => assert.deepEqual(w.sheets[0].rowHeights.slice(0, 4), [36, 22, 44, 22])},
  {name: '枠の固定', covers: ['sheet.freeze', 'freeze.rows', 'freeze.cols'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {sheet: {freeze: {rows: 2, cols: 1}}}),
    excel: x => assert.deepEqual(x.sheets[0].view.pane, {xSplit: 1, ySplit: 2, topLeftCell: 'B3', activePane: 'bottomRight', state: 'frozen'}),
    back: w => assert.deepEqual(w.sheets[0].freeze, {rows: 2, cols: 1})},
  // Gridlines shown (true) is the default: it is not written, and the sheet read back has no gridlines field.
  {name: '枠線の表示', covers: ['sheet.gridlines'],
    book: () => { const w = bookOf({A1: {t: 's', v: 'x'}}), base = w.sheets[0]; w.sheets = [{...base, gridlines: false}, {...base, id: 'b', name: '枠線あり', gridlines: true}]; return w; },
    excel: x => assert.deepEqual(x.sheets.map(s => s.view.showGridLines), [false, true]),
    back: w => assert.deepEqual(w.sheets.map(s => s.gridlines), [false, undefined])},
  {name: '画像', covers: ['sheet.images', 'book.assets', 'image.r', 'image.c', 'image.x', 'image.y', 'image.w', 'image.h'],
    book: () => bookOf({}, {assets: {logo: PNG}, sheet: {images: [{id: 'p', asset: 'logo', r: 5, c: 2, x: 10, y: 6, w: 90, h: 40}]}}),
    excel: x => { const [p] = x.sheets[0].pictures;
      assert.deepEqual({anchor: p.anchor, from: p.from, ext: p.ext, bytes: p.bytes.toString('base64')},
        {anchor: 'oneCellAnchor', from: {col: 2, colOff: 10 * EMU, row: 5, rowOff: 6 * EMU}, ext: {cx: 90 * EMU, cy: 40 * EMU}, bytes: PNG.split(',')[1]}); },
    back: w => { assert.deepEqual(w.sheets[0].images, [{id: 'pic-1', asset: 'img-1', r: 5, c: 2, x: 10, y: 6, w: 90, h: 40}]); assert.equal(w.assets['img-1'], PNG); }},
  {name: '倍率', covers: ['book.view', 'view.zoom'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {view: {zoom: 120}}),
    excel: x => assert.equal(x.sheets[0].view.zoomScale, 120),
    back: w => assert.equal(w.view.zoom, 120)},
  {name: '選択', covers: ['view.active', 'view.selection', 'range.r1', 'range.c1', 'range.r2', 'range.c2'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {view: {zoom: 100, active: [2, 1], selection: {r1: 2, c1: 1, r2: 3, c2: 2}}}),
    excel: x => assert.deepEqual(x.sheets[0].view.selections, [{pane: undefined, activeCell: 'B3', sqref: 'B3:C4'}]),
    back: w => assert.deepEqual(w.view, {zoom: 100, active: [2, 1], selection: {r1: 2, c1: 1, r2: 3, c2: 2}})},
  // The scroll position is where the page was when its state was saved, not part of the table. It is not written, so
  // Excel opens the sheet at its top-left cell.
  {name: 'スクロールの位置（xlsxに書かない）', covers: ['view.scrollTop', 'view.scrollLeft'],
    book: () => bookOf({A1: {t: 's', v: 'x'}}, {view: {zoom: 100, scrollTop: 300, scrollLeft: 50}}),
    excel: x => assert.equal(sheetView(x.sheets[0]).attrs.topLeftCell, undefined),
    back: w => assert.deepEqual(w.view, {zoom: 100})},
];

test('core.js とリボンにある機能には、どれもxlsxの書き出しと読み込みの見本がある', () => {
  assert.ok(ribbonFonts.length >= 7, 'リボンのフォントの一覧を読めている');
  const list = features(), covered = new Set(SAMPLES.flatMap(s => s.covers));
  assert.deepEqual(list.filter(f => !covered.has(f)), [], '見本のない機能がある。SAMPLES に、書き出しと読み込みを確かめる見本を足す（SKILL.md の「実装を変えるとき」）');
  assert.deepEqual([...covered].filter(f => !list.includes(f)), [], 'HTMLにない機能の見本がある。機能を外したか名前を変えたなら、見本も直す');
  assert.equal(new Set(SAMPLES.map(s => s.name)).size, SAMPLES.length, '見本の名前は重ならない');
});

for (const sample of SAMPLES) {
  test(`${sample.name}：xlsxに書き出し、読み込み直す`, () => {
    const w = sample.book();
    Core.validate(w);
    const {bytes, warnings} = toXlsx(w);
    assert.deepEqual(warnings, [], '書き出しの知らせ');
    sample.excel(readXlsx(bytes), w);
    const read = fromXlsx(bytes);
    assert.deepEqual(read.warnings, [], '読み込みの知らせ');
    (sample.back || sameCells)(read.workbook, w);
  });
}
