// scripts/verify_xlsx.py：Excelで開いたxlsxとHTMLの表を比べた差を、想定外の差と、仕様どおりの差（読み込みで置き換えた所）に
// 分ける部分（check_sheet）。Excelは使わず、Excelから読み取る値（scripts/excel_check/excel.ps1 の xlsx モード）のかわりの辞書を渡す。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {join} from 'node:path';
import {test} from 'node:test';
import {blankBook, findPython, SKILL_ROOT} from './lib/workbook.mjs';

/** The result of call (Python, with the arguments in a) on scripts/verify_xlsx.py imported as v, read back from JSON. */
function verify(call, a) {
  const py = findPython();
  const code = 'import json, sys\nsys.dont_write_bytecode = True\nsys.path.insert(0, sys.argv[1])\nimport verify_xlsx as v\na = json.load(sys.stdin)\n' +
    `print(json.dumps(${call}, ensure_ascii=False))`;
  const run = spawnSync(py.command, [...py.args, '-c', code, join(SKILL_ROOT, 'scripts')], {input: JSON.stringify(a), encoding: 'utf8', timeout: 60000,
    env: {...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1'}});
  if (run.status !== 0) throw new Error(run.stderr);
  return JSON.parse(run.stdout);
}
/** check_sheet for the first sheet of workbook, as {problems, expected}. */
const checkSheet = (workbook, info, notes = null) => {
  const [problems, expected] = verify('v.check_sheet(a["workbook"], a["workbook"]["sheets"][0], a["info"], a["notes"])', {workbook, info, notes});
  return {problems, expected};
};
const items = list => list.map(d => d.reason ? [d.item, d.reason] : d.item);

// The table: column A 80px (60pt), B to D the page's 96px (72pt); row 1 30px (22.5pt), rows 2 and 3 22px (16.5pt); one picture.
const workbook = blankBook({view: {zoom: 120}, assets: {p: 'data:image/png;base64,AAAA'}, sheet: {name: 'S', rows: 3, cols: 4, colWidths: [80], rowHeights: [30],
  images: [{id: 'p1', asset: 'p', r: 0, c: 0, x: 0, y: 0, w: 40, h: 20}]}});
// What Excel shows of the file: column B hidden, C and D at Excel's default width, row 2 hidden, the panes frozen, a
// zoom of 100%, and two pictures (one more than the page has; the first one wider: 36pt where the page has 30pt).
const excel = {repairs: [], sheetNames: ['S'], activeSheet: 'S', colWidths: [60, 0, 54, 54], rowHeights: [22.5, 0, 16.5], freeze: true, splitRow: 1, splitColumn: 0,
  gridlines: true, zoom: 100, page: {paperSize: 9, orientation: 1, zoom: false, fitWide: 1, fitTall: false},
  pictures: [{topLeft: 'A1', left: 0, top: 0, width: 36, height: 15}, {topLeft: 'B2', left: 60, top: 22.5, width: 10, height: 10}]};
const DEFAULT_WIDTH = 'Excelの既定の列幅は、表示の倍率で少し変わる（96dpiで72px、120dpiで70.4px相当）';

test('元のxlsxとの照合では、読み込みが知らせた列・行・固定・倍率・画像の差を、その列や行に付いた理由ごとに仕様どおりの差にする', () => {
  const notes = {cells: {}, sheet: [{item: 'colWidth', cols: ['B'], reason: '非表示の列'}, {item: 'colWidth', cols: ['C', 'D'], reason: DEFAULT_WIDTH},
    {item: 'rowHeight', rows: [2], reason: '非表示の行'}, {item: 'freeze', reason: '固定の境目をまたぐ結合'}, {item: 'zoom', reason: 'HTMLの倍率はブック全体で1つ'},
    {item: 'pictures', reason: '移せない画像'}]};
  const {problems, expected} = checkSheet(workbook, excel, notes);
  assert.deepEqual(problems, []);
  assert.deepEqual(items(expected), [['列 B の幅（pt）', '非表示の列'], ['列 C の幅（pt）', DEFAULT_WIDTH], ['列 D の幅（pt）', DEFAULT_WIDTH], ['2 行目の高さ（pt）', '非表示の行'],
    ['枠の固定', '固定の境目をまたぐ結合'], ['倍率（%）', 'HTMLの倍率はブック全体で1つ'], ['画像の数', '移せない画像']]);
});

test('元のxlsxとの照合でも、理由のない差・修復・シートの名前・開くシート・枠線の差は想定外の差にし、印刷の設定は比べない', () => {
  const notes = {cells: {}, sheet: [{item: 'colWidth', cols: ['B'], reason: '非表示の列'}, {item: 'colWidthAll', reason: '標準のフォントが游ゴシック11ptでないので、列幅は近い値'}]};
  const info = {...excel, repairs: ['修復したレコード: セルの情報'], sheetNames: ['S', '別'], activeSheet: '別', gridlines: false, page: {paperSize: 1, orientation: 2, zoom: 100, fitWide: false, fitTall: false}};
  const {problems, expected} = checkSheet(workbook, info, notes);
  assert.deepEqual(items(problems), ['Excelによる修復', 'シートの名前と順番', '開いたときのシート', '2 行目の高さ（pt）', '枠の固定', '枠線の表示', '倍率（%）', '画像の数',
    '画像 1（p1）の幅'], '印刷の設定は、元のファイルのものなので比べない');
  assert.deepEqual(items(expected), [['列 B の幅（pt）', '非表示の列'], ['列 C の幅（pt）', '標準のフォントが游ゴシック11ptでないので、列幅は近い値'],
    ['列 D の幅（pt）', '標準のフォントが游ゴシック11ptでないので、列幅は近い値']], '列ごとの理由がない列は、ブック全体の列幅の理由');
  const rows = checkSheet(workbook, info, {cells: {}, sheet: [...notes.sheet, {item: 'rowHeightAll', reason: '標準のフォントに合う行の高さは近い値'}]});
  assert.deepEqual(items(rows.expected).slice(-1), [['2 行目の高さ（pt）', '標準のフォントに合う行の高さは近い値']], '行ごとの理由がない行は、ブック全体の行の高さの理由');
});

test('セルの差は、読み込みが置き換えたと伝えた項目なら仕様どおりの差にし、日本語の文字に英語のフォントを付けたセルでExcelがフォント名を答えないことは差にしない', () => {
  // Excelの中で、日本語の文字のセルにArialを付けると、日本語の文字は別のフォントで描かれ、フォント名は答えない（None）。
  const book = blankBook({sheet: {cells: {A1: {t: 's', v: '文字', s: {font: 'Arial'}}, A2: {t: 's', v: '文字', s: {font: 'Arial'}}, B1: {t: 'n', v: 1}}}});
  const diffs = [{cell: 'A1', field: 'font', label: 'フォント', excel: null, pasted: 'Arial'}, {cell: 'A2', field: 'font', label: 'フォント', excel: null, pasted: 'Calibri'},
    {cell: 'B1', field: 'numfmt', label: '表示形式', excel: '0', pasted: '0_ '}, {cell: 'B1', field: 'bTop', label: '上罫線', excel: 'none', pasted: '1/2/0', cause: 'border-conflict'}];
  const notes = {cells: {B1: [{fields: ['numfmt'], reason: '表示形式「0_ 」は、同じ表示の「0」にした'}]}, sheet: []};
  const [cells, conflicts, expected] = verify('v.split_cells(a["diffs"], a["workbook"], a["workbook"]["sheets"][0], a["notes"])', {diffs, workbook: book, notes});
  assert.deepEqual(cells.map(d => d.cell), ['A2'], 'ファイルのフォントがHTMLと違えば、想定外の差');
  assert.deepEqual(expected.map(d => [d.cell, d.field, d.reason]), [['B1', 'numfmt', '表示形式「0_ 」は、同じ表示の「0」にした']]);
  assert.deepEqual(conflicts.map(d => d.cell), ['B1']);
  const written = verify('v.split_cells(a["diffs"], a["workbook"], a["workbook"]["sheets"][0], None)', {diffs, workbook: book});
  assert.deepEqual(written[0].map(d => d.cell + ' ' + d.field), ['A2 font', 'B1 numfmt'], '書き出したxlsxの照合では、読み込みの理由を使わない');
});

test('書き出したxlsxの照合では、差はすべて想定外の差にし、印刷の設定（A4縦・すべての列を1ページ）と画像の位置も比べる', () => {
  const {problems, expected} = checkSheet(workbook, excel);
  assert.deepEqual(expected, []);
  assert.deepEqual(items(problems), ['列 B の幅（pt）', '列 C の幅（pt）', '列 D の幅（pt）', '2 行目の高さ（pt）', '枠の固定', '倍率（%）', '画像の数', '画像 1（p1）の幅']);
  const same = {...excel, colWidths: [60, 72, 72, 72], rowHeights: [22.5, 16.5, 16.5], freeze: false, zoom: 120, pictures: [{topLeft: 'A1', left: 0.5, top: 0, width: 30.4, height: 15}]};
  assert.deepEqual(checkSheet(workbook, same).problems, [], '大きさは0.8pt、位置は3ptまでの差を同じとする');
  const letter = checkSheet(workbook, {...same, page: {paperSize: 1, orientation: 1, zoom: false, fitWide: 1, fitTall: false}});
  assert.deepEqual(items(letter.problems), ['印刷の設定（A4縦、すべての列を1ページに印刷）']);
});
