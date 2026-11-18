// 関数の計算を、デスクトップ版Excelが計算して保存した値と比べる。
//
// tests/fixtures/excel-functions.xlsx は、tests/fixtures/excel-functions.mjs の数式を desktop Excel 16.0（build 20326）に
// 入れて計算させ、その値ごと保存したファイルである。HTMLが計算する関数（core.js の FUNCTIONS）のすべてを、引数の型
// （数・文字・TRUE/FALSE・空のセル・エラー・範囲）と境目の値を変えて使っている。
// ここでは、このファイルをHTMLの読み込み（scripts/read_xlsx.cjs）で読み、どの数式も数式のまま移ることと、
// HTMLで計算した値がExcelの保存した値と最後の桁まで同じになることを確かめる。
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {test} from 'node:test';
import {Core, SKILL_ROOT} from './lib/workbook.mjs';
import {kid, parseXml, readXlsx, textOf} from './lib/xlsx.mjs';
import {bookSpec} from './fixtures/excel-functions.mjs';

const require = createRequire(import.meta.url);
const F = require(join(SKILL_ROOT, 'assets', 'formulas.js'));
const {fromXlsx} = require(join(SKILL_ROOT, 'scripts', 'read_xlsx.cjs'));
const {NEWER_FUNCTIONS} = require(join(SKILL_ROOT, 'scripts', 'xlsx.cjs'));
const bytes = readFileSync(join(SKILL_ROOT, 'tests', 'fixtures', 'excel-functions.xlsx'));
const file = readXlsx(bytes);
// The text functions (&, CONCATENATE, CONCAT, TEXTJOIN) have their own file, which tests/xlsx-import.test.mjs compares.
const textFile = readXlsx(readFileSync(join(SKILL_ROOT, 'tests', 'fixtures', 'excel-text.xlsx')));
// The formula as Excel shows it: the file writes _xlfn. before the names of newer functions.
const shown = f => f.replace(/"(?:[^"]|"")*"|(?<![A-Za-z0-9_.])_xlfn\.(?:_xlws\.)?/gi, t => t[0] === '"' ? t : '');
const formulas = (x = file) => x.sheets.flatMap(s => Object.entries(s.cells).filter(([, c]) => c.formula !== undefined).map(([a, c]) => ({sheet: s.name, a, ...c})));
const bothFiles = () => [...formulas(), ...formulas(textFile)];
// The time Excel made the file: a date written without its year ("1/2") is in that year.
const made = Date.parse(textOf(kid(parseXml(file.parts.get('docProps/core.xml').toString('utf8')), 'created')));
/** Runs fn with TODAY and NOW reading the clock at time t. */
function at(t, fn) { const now = F.clock.now; F.clock.now = () => t; try { return fn(); } finally { F.clock.now = now; } }
// Excel's own power differs from the page's in the last binary digit for a few fractions (formulas.js, power); the 15
// digits Excel shows are the same.
const LAST_DIGIT_ONLY = new Set(['数学!H352']);

test('Excelで計算させた2500の数式は、xlsxから数式のまま移り、HTMLの計算がExcelの保存した値と最後の桁まで同じになる', () => {
  at(made, () => {
    const result = fromXlsx(bytes, {name: 'excel-functions.xlsx'}), w = result.workbook;
    assert.deepEqual(result.warnings.filter(text => !/表示形式/.test(text)), [], '知らせは、HTMLにない日付と時刻の表示形式だけ');
    const all = formulas(), caches = new Map(), kept = [], differ = [];
    assert.ok(all.length > 2400, 'ファイルの数式を読めている');
    for (const {sheet, a, value} of all) {
      const s = w.sheets.find(x => x.name === sheet);
      if (s.cells[a]?.t !== 'f') { kept.push(sheet + '!' + a); continue; }
      if (!caches.has(sheet)) caches.set(sheet, new Map());
      const [r, c] = Core.coord(a), got = Core.calculate(w, s, r, c, new Set(), caches.get(sheet)), want = value ?? '';
      const same = got === want || LAST_DIGIT_ONLY.has(sheet + '!' + a) && typeof got === 'number' && F.to15(got) === F.to15(want);
      if (!same) differ.push(`${sheet}!${a} =${shown(s.cells[a].f.slice(1))}: HTML ${JSON.stringify(got)}、Excel ${JSON.stringify(want)}`);
    }
    assert.deepEqual(kept, [], '値になった数式');
    assert.deepEqual(differ, [], 'HTMLの計算がExcelの値と違う数式');
  });
});

test('HTMLが計算する関数は、どれもExcelで作ったファイルでExcelの値と比べてある', () => {
  const used = new Set(bothFiles().flatMap(({formula}) => [...shown(formula).replace(/"(?:[^"]|"")*"/g, '').matchAll(/([A-Z][A-Z0-9.]*)\s*\(/g)].map(m => m[1])));
  assert.deepEqual(Core.FUNCTIONS.filter(name => !used.has(name)), [], 'ファイルにない関数');
  assert.deepEqual([...used].filter(name => !Core.FUNCTIONS.includes(name)), [], 'HTMLが計算しない関数（ファイルに入れたなら FUNCTIONS にも要る）');
});

test('xlsxに _xlfn. を付けて書く関数は、Excelが付けて保存した関数とちょうど同じ', () => {
  const saved = new Set(bothFiles().flatMap(({formula}) => [...formula.replace(/"(?:[^"]|"")*"/g, '').matchAll(/_xlfn\.([A-Z][A-Z0-9.]*)\(/g)].map(m => m[1])));
  assert.deepEqual([...NEWER_FUNCTIONS].sort(), [...saved].sort());
});

// A number in a formula as Excel saves it: to 15 significant digits, the rest cut off (2.9999999999999996 becomes
// 2.99999999999999, and 1E+15 1000000000000000), to compare the file with the list it was made from.
const sameNumbers = f => f.replace(/"(?:[^"]|"")*"|(?<![A-Za-z0-9_.$])(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?(?![A-Za-z0-9_(])/gi, t => t[0] === '"' ? t : String(F.numberLiteral(t)));
test('ファイルの数式は excel-functions.mjs の数式と同じ（数式を変えたら、ファイルを作り直してある）', () => {
  const want = bookSpec().sheets.flatMap(s => Object.entries(s.cells).filter(([, c]) => c?.f).map(([a, c]) => `${s.name}!${a} ${sameNumbers(c.f.slice(1))}`));
  const got = formulas().map(({sheet, a, formula}) => `${sheet}!${a} ${sameNumbers(shown(formula))}`);
  assert.deepEqual(got.sort(), want.sort());
});

test('TODAY と NOW は、パソコンの時計のその時の日付と時刻を、Excelと同じ数（1900年1月1日が1、1日が1）で返す', () => {
  const local = new Date(2026, 0, 2, 18, 0, 0).getTime(); // 2026/01/02 18:00, the computer's own time
  at(local, () => {
    assert.equal(F.evalFormula('=TODAY()', null), 46024);
    assert.equal(F.evalFormula('=NOW()', null), 46024.75);
    assert.equal(F.evalFormula('=TEXT(NOW(),"yyyy/mm/dd hh:mm")', null), '2026/01/02 18:00');
    assert.equal(F.evalFormula('=DATEVALUE("1/2")', null), 46024, '年のない日付はその年');
  });
});
