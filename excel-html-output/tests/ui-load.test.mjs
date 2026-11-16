// 生成したHTMLが開けること：見本とテスト用のブックが、エラーなしで描画される。
import assert from 'node:assert/strict';
import {join} from 'node:path';
import {readFileSync} from 'node:fs';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {blankBook, fixture, SKILL_ROOT} from './lib/workbook.mjs';

const env = useBrowser();

test('見本の assets/example.json が、エラーなしで全シートを描画する', async t => {
  const example = JSON.parse(readFileSync(join(SKILL_ROOT, 'assets', 'example.json'), 'utf8'));
  const x = await env.open(t, example);
  assert.equal(await x.evaluate(() => document.title), example.title);
  const tabs = await x.evaluate(() => [...document.querySelectorAll('#sheet-tabs [data-sheet]')].map(b => b.textContent));
  assert.deepEqual(tabs, example.sheets.map(s => s.name));
  for (const sheet of example.sheets) {
    await x.click(`#sheet-tabs [data-sheet="${sheet.id}"]`);
    const size = await x.evaluate(() => ({rows: document.querySelectorAll('#grid tbody tr').length, cols: document.querySelectorAll('#grid thead th[data-col]').length}));
    assert.deepEqual(size, {rows: sheet.rows, cols: sheet.cols}, sheet.name + ' の行数と列数');
  }
});

test('テスト用のブックを開くと、題名・シートタブ・A1の選択・準備完了が表示される', async t => {
  const x = await env.open(t, fixture('book'));
  assert.equal(await x.evaluate(() => document.title), 'テスト用ブック');
  assert.equal(await x.value('book-title'), 'テスト用ブック');
  assert.deepEqual(await x.evaluate(() => [...document.querySelectorAll('#sheet-tabs [data-sheet]')].map(b => [b.textContent, b.getAttribute('aria-selected')])),
    [['明細', 'true'], ['予備', 'false']]);
  assert.equal(await x.nameBox(), 'A1');
  assert.equal(await x.formula(), 'テスト用の明細表（右の空いたセルへはみ出す長い題名）');
  assert.equal(await x.status(), '準備完了');
  assert.equal(await x.textOf('#dirty-state'), '書き出し済み / 元の状態');
  assert.equal(await x.text('E9'), '¥83,600', '数式は開いた時点で計算されている');
});

test('セルの文字は、HTMLやスクリプトとして解釈されずに、そのまま表示される', async t => {
  const book = blankBook({sheet: {cells: {
    A1: {t: 's', v: '<img src=x onerror="window.__xss=1">'},
    A2: {t: 's', v: '</script><script>window.__xss=2</script>'},
    A3: {t: 's', v: '__TITLE__ /*__DATA__*/'},
  }}});
  const x = await env.open(t, book);
  assert.equal(await x.text('A1'), '<img src=x onerror="window.__xss=1">');
  assert.equal(await x.text('A2'), '</script><script>window.__xss=2</script>');
  assert.equal(await x.text('A3'), '__TITLE__ /*__DATA__*/', 'ビルドの置換の印に似た文字も、そのまま残る');
  assert.equal(await x.evaluate(() => window.__xss ?? null), null);
  assert.equal(await x.evaluate(() => document.querySelectorAll('#grid img').length), 0);
});

test('保存されていた選択位置・ズーム・スクロール位置で開く', async t => {
  const book = fixture('book');
  book.view = {zoom: 125, active: [3, 2], selection: {r1: 3, c1: 2, r2: 5, c2: 4}, scrollTop: 40, scrollLeft: 0};
  const x = await env.open(t, book);
  assert.equal(await x.nameBox(), 'C4:E6');
  assert.equal(await x.activeCell(), 'C4');
  assert.equal(await x.textOf('#zoom-label'), '125%');
  assert.equal(await x.evaluate(() => document.getElementById('sheet-surface').style.zoom), '1.25');
  assert.equal(await x.evaluate(() => document.getElementById('viewport').scrollTop), 40);
  assert.ok(await x.selectionFrameGap('C4', 'E6') <= 2, '緑の選択枠が、拡大した表示の上でも選択範囲に重なる');
});
