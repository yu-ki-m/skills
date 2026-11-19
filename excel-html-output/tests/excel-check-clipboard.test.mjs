// scripts/excel_check/clipboard.cjs：照合のスクリプト（verify_excel_paste.py・verify_xlsx.py）が使う、範囲のクリップボードHTMLと、
// セルごとの注記。注記の border-conflict は、Excelとの罫線の差を「データの食い違いによる差」として分けるのに使う。
// Excel もクリップボードも使わない。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {existsSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {after, before, test} from 'node:test';
import {blankBook, removeDir, SKILL_ROOT, tempDir} from './lib/workbook.mjs';

let dir;
before(() => { dir = tempDir(); });
after(() => removeDir(dir));

/** Runs clipboard.cjs on the sheet main of book for range, and returns what it wrote in <name>.notes.json. */
function notesOf(name, book, range) {
  const wb = join(dir, name + '.json'), out = join(dir, name);
  writeFileSync(wb, JSON.stringify(book), 'utf8');
  const result = spawnSync(process.execPath, [join(SKILL_ROOT, 'scripts', 'excel_check', 'clipboard.cjs'), wb, 'main', range, out],
    {encoding: 'utf8', timeout: 60000});
  assert.equal(result.status, 0, result.stderr);
  assert.ok(existsSync(out + '.html') && existsSync(out + '.txt'), 'クリップボードのHTMLと文字も書き出す');
  return JSON.parse(readFileSync(out + '.notes.json', 'utf8'));
}

test('隣り合うセルの境界の違う線は、両方のセルに border-conflict の注記を付け、食い違いとして書き出す', () => {
  const L = '1px solid #808080', T = '2px solid #217346';
  const one = notesOf('one', blankBook({sheet: {cells: {A1: {t: 's', v: '上', s: {borderBottom: L}}, A2: {t: 's', v: '下', s: {borderTop: T}}}}}), 'A1:B2');
  assert.deepEqual(one.notes, {A1: 'border-conflict', A2: 'border-conflict'});
  assert.deepEqual(one.conflicts, [{a: 'A1', b: 'A2', edge: '下/上', lines: [L, T], cells: ['A1', 'A2']}]);
});

test('結合セルは1つの塊として、左上のセルの線で比べる', () => {
  const L = '1px solid #808080', T = '2px solid #217346';
  const merged = notesOf('merged', blankBook({sheet: {merges: [{r1: 1, c1: 1, r2: 2, c2: 2}],
    cells: {B2: {t: 's', v: '結合', s: {borderRight: T}}, D3: {t: 's', v: '', s: {borderLeft: L}}}}}), 'A1:E5');
  assert.deepEqual(merged.conflicts, [{a: 'C3', b: 'D3', edge: '右/左', lines: [T, L], cells: ['B2:C3', 'D3']}]);
  assert.deepEqual(merged.notes, {C3: 'border-conflict', D3: 'border-conflict'});
});

test('大文字と小文字だけ違う色は同じ線なので、注記を付けない', () => {
  const same = notesOf('case', blankBook({sheet: {cells: {A1: {t: 's', v: '', s: {borderRight: '1px solid #ABCDEF'}},
    B1: {t: 's', v: '', s: {borderLeft: '1px solid #abcdef'}}}}}), 'A1:B1');
  assert.deepEqual([same.notes, same.conflicts], [{}, []]);
});
