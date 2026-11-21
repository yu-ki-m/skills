// scripts/build.py の状態テキストの反映（--state）：形式、ブックIDとリビジョンの照合（反映済みの状態、以前の版から
// 書き出した状態）、書き換えられたHTML、反映する先。
import assert from 'node:assert/strict';
import {readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {test} from 'node:test';
import {listHtml, newFolder, olderWith, stateText, useTempDir} from './lib/build.mjs';
import {blankBook, buildBook, embeddedBook, fixture, runBuild} from './lib/workbook.mjs';

useTempDir();

for (const [format, label] of [['toon', 'TOON'], ['json', 'JSON']]) {
  test(`状態テキスト（${label}）の反映：ID違い・途中で切れたテキスト・見出しのないテキストは拒み、元のファイルを変えない`, () => {
    const folder = newFolder('state-' + format);
    const html = buildBook(fixture('book'), folder, '表.html');
    const original = readFileSync(html, 'utf8');
    const good = stateText(html, w => { w.sheets[0].cells.B4.v = '反映'; }, format);
    const lines = good.split('\n'), cut = format === 'toon' ? /EXCEL_HTML_STATE_END/ : /JSON/;
    const cases = {
      'ブックID違い': [stateText(html, (w, state) => { state.workbookId = 'other'; }, format), /この編集状態は別の表のものです（状態の題名: テスト用ブック、ブックID: other。元のHTMLのブックID: /],
      '途中で切れた': [good.slice(0, Math.floor(good.length / 2)), cut],
      '行の境目で切れた': [lines.slice(0, Math.floor(lines.length / 2)).join('\n'), cut],
      '最後の行が欠けた': [lines.slice(0, -1).join('\n'), cut],
      '見出しなし': [lines.slice(1).join('\n'), /EXCEL_HTML_STATE_V1で始まる/],
    };
    for (const [name, [text, reason]] of Object.entries(cases)) {
      writeFileSync(join(folder, 'state.txt'), text, 'utf8');
      const result = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
      assert.equal(result.status, 1, name);
      assert.match(result.stderr, /^エラー: /, name);
      assert.match(result.stderr, reason, name);
      assert.deepEqual(listHtml(folder), ['表.html'], name + 'では以前の内容のファイルも作らない');
      assert.equal(readFileSync(html, 'utf8'), original, name);
    }
    writeFileSync(join(folder, 'state.txt'), good, 'utf8');
    const applied = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
    assert.equal(applied.status, 0, applied.stderr);
    assert.equal(embeddedBook(html).sheets[0].cells.B4.v, '反映');
  });
}

test('作り直した後のHTMLに、作り直す前の版から書き出した状態を渡すと、以前の版を元にするよう伝えて止まる', () => {
  const folder = newFolder('older-base');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  writeFileSync(join(folder, 'second.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  const result = runBuild(['--base', html, '--state', join(folder, 'second.txt')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /この編集状態は、--base のHTMLの版ではなく、別の版（以前の版など）から書き出されたものです/);
  assert.match(result.stderr, /以前の版（表_<年月日時分秒>\.html）から、リビジョンが一致するものを --base に/);
  const older = listHtml(folder).find(n => /^表_\d{14}\.html$/.test(n));
  const recovered = runBuild(['--base', join(folder, older), '--state', join(folder, 'second.txt'), '--output', html]);
  assert.equal(recovered.status, 0, recovered.stderr);
  assert.deepEqual([embeddedBook(html).sheets[0].cells.B4.v, embeddedBook(html).sheets[0].cells.B5.v], ['1回目', '2回目']);
});

test('反映済みの編集状態をもう一度渡すと（コピーし直し忘れ）、反映済みと伝えて止め、HTMLを書き換えない', () => {
  const folder = newFolder('stale');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '反映した'; w.sheets[0].cells.H1 = {t: 'f', f: '=SUMSQ(1,2)'}; }), 'utf8');
  const first = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(first.status, 0, first.stderr);
  assert.match(first.stderr, /^注意: HTMLでは計算できない数式があります/m);
  const applied = readFileSync(html, 'utf8'), before = listHtml(folder);
  const older = before.find(n => /^表_\d{14}\.html$/.test(n));
  // 以前の版を元にする回復の手順でも、書き出し先のHTMLと同じ内容になるなら、新しい編集はない。
  const result = runBuild(['--base', join(folder, older), '--state', join(folder, 'state.txt'), '--output', html]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /この編集状態は、すでに .*表\.html に反映されています/);
  assert.match(result.stderr, /もう一度「テキストをコピー」を押してから/);
  assert.ok(!/注意:/.test(result.stderr), '反映しないので、数式の注意も出さない:\n' + result.stderr);
  assert.equal(readFileSync(html, 'utf8'), applied);
  assert.deepEqual(listHtml(folder), before, '以前の内容のファイルも増やさない');
});

test('リビジョンの合わない状態は、すでに反映したものなら反映済みと伝え、そうでなければ一致する以前の版の名前を示す', () => {
  const folder = newFolder('mismatch-names');
  const html = buildBook(fixture('book'), folder, '表.html');
  const base = embeddedBook(html).revision;
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  writeFileSync(join(folder, 'second.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  const again = runBuild(['--base', html, '--state', join(folder, 'first.txt'), '--xlsx', join(folder, '表.xlsx')]);
  assert.equal(again.status, 1);
  assert.match(again.stderr, /この編集状態は、すでに --base の .*表\.html に反映されています。--state を付けずに、--base だけで作ってください（同じ表になります）。/);
  const [older] = olderWith(folder, base);
  const named = runBuild(['--base', html, '--state', join(folder, 'second.txt')]);
  assert.equal(named.status, 1);
  assert.match(named.stderr, new RegExp('リビジョンが一致するのは ' + older.split(/[\\/]/).pop().replace('.', '\\.') + ' です'));
  const fromOlder = runBuild(['--base', older, '--state', join(folder, 'first.txt'), '--output', join(folder, '別名.html')]);
  assert.equal(fromOlder.status, 0, '一致する以前の版からなら、別の名前に書き出せる: ' + fromOlder.stderr);
});

test('生成後に中のデータを書き換えたHTMLは、反映の元として受け付けない', () => {
  const folder = newFolder('tampered');
  const html = buildBook(fixture('book'), folder, '表.html');
  const text = stateText(html, () => {});
  writeFileSync(html, readFileSync(html, 'utf8').replace('"保守作業"', '"改ざん"'), 'utf8');
  writeFileSync(join(folder, 'state.txt'), text, 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /リビジョン生成後に変更されています/);
});

test('反映する先（--output）が別の表のHTMLなら、置き換えずに止まる', () => {
  const folder = newFolder('other-book');
  const html = buildBook(fixture('book'), folder, '表.html');
  const other = buildBook(blankBook({id: 'other-book', title: '別の表'}), folder, '別.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '反映'; }), 'utf8');
  const before = readFileSync(other, 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--output', other]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /--output の 別\.html は、別の表のHTMLです（ブックID: other-book、題名: 別の表）。反映すると、その表が置き換わります。/);
  assert.equal(readFileSync(other, 'utf8'), before);
});
