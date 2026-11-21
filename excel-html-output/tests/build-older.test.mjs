// scripts/build.py で、以前の版を元に反映するとき：止まったときの直し方、--output の指定、以前の版の見分け方、
// 反映すると元に戻る変更と、xlsx・JSONに入らない変更の注意。
import assert from 'node:assert/strict';
import {readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {test} from 'node:test';
import {listHtml, newFolder, olderWith, stateText, useTempDir, xlsxOf} from './lib/build.mjs';
import {blankBook, buildBook, embeddedBook, fixture, runBuild} from './lib/workbook.mjs';

useTempDir();

/** A folder with 表.html applied once from a first state, then built again by Claude with change(json) (the recovery
 *  case): returns {folder, html, older (the version the states were copied from), second (a state from it)}. */
function rebuiltAfterCopy(name, change, second = w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }) {
  const folder = newFolder(name);
  const html = buildBook(fixture('book'), folder, '表.html');
  const base = embeddedBook(html).revision;
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  writeFileSync(join(folder, 'second.txt'), stateText(html, second), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  const json = join(folder, 'edit.json');
  assert.equal(runBuild(['--base', html, '--output', json, '--extract-json']).status, 0);
  const edited = JSON.parse(readFileSync(json, 'utf8'));
  change(edited);
  writeFileSync(json, JSON.stringify(edited), 'utf8');
  const rebuilt = runBuild(['--input', json, '--output', html]);
  assert.equal(rebuilt.status, 0, rebuilt.stderr);
  const [older] = olderWith(folder, base);
  return {folder, html, older, second: join(folder, 'second.txt')};
}

test('以前の版から書き出した状態で止まったときの直し方は、書き出すもので変わる（HTMLは --output に元の名前、xlsx とJSONは --base だけを替える）', () => {
  const folder = newFolder('older-modes');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  writeFileSync(join(folder, 'second.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  const state = join(folder, 'second.txt');
  const xlsx = runBuild(['--base', html, '--state', state, '--xlsx', join(folder, '表.xlsx')]);
  assert.equal(xlsx.status, 1);
  assert.match(xlsx.stderr, /リビジョンが一致するものを --base に指定してください（--xlsx はそのまま）/);
  assert.match(xlsx.stderr, /この編集状態をすでにHTMLに反映したなら、--state を付けずに、反映した後のHTMLを --base にしてください/);
  assert.ok(!/元のファイル名を --output に指定してください/.test(xlsx.stderr), '--xlsx は --output と一緒に使えないので、--xlsx の実行に --output を指定させない');
  assert.match(xlsx.stderr, /作り直した後の変更はxlsxに入らない（「注意:」に出ます）ので、先にこの編集状態をHTMLに反映し（--base に以前の版、--output に元のファイル名）、反映した後のHTMLから作るほうが確かです/,
    '--output を使うのは、先にHTMLに反映するとき');
  const json = runBuild(['--base', html, '--state', state, '--output', join(folder, 'edit.json'), '--extract-json']);
  assert.equal(json.status, 1);
  assert.match(json.stderr, /リビジョンが一致するものを --base に指定してください（--output のJSONはそのまま）/);
  assert.ok(!/元のファイル名を --output/.test(json.stderr), 'JSON を元のHTMLの名前に書き出させない');
  const extracted = runBuild(['--base', html, '--output', join(folder, 'edit.json'), '--extract-json']);
  assert.equal(extracted.status, 0, '反映した後のHTMLからは、--state なしで取り出せる: ' + extracted.stderr);
  assert.equal(JSON.parse(readFileSync(join(folder, 'edit.json'), 'utf8')).sheets[0].cells.B4.v, '1回目');
  const made = runBuild(['--base', html, '--xlsx', join(folder, '表.xlsx')]);
  assert.equal(made.status, 0, made.stderr);
  assert.equal(xlsxOf(join(folder, '表.xlsx')).sheets[0].cells.B4.value, '1回目');
});

test('以前の版を --base にして --output を省くと、以前の版に書き出してしまうので、元のファイル名を --output に指定するよう伝えて止まる', () => {
  const folder = newFolder('older-without-output');
  const html = buildBook(fixture('book'), folder, '表.html');
  const base = embeddedBook(html).revision;
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  writeFileSync(join(folder, 'second.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  const [older] = olderWith(folder, base), before = listHtml(folder);
  const result = runBuild(['--base', older, '--state', join(folder, 'second.txt')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /--base の 表_\d{14}\.html は、表\.html の以前の版です。--output を省くと、この以前の版に書き出すので、ユーザーが開く 表\.html は変わりません。--output に .*表\.html を指定してください/);
  assert.deepEqual(listHtml(folder), before);
  const into = runBuild(['--base', older, '--state', join(folder, 'second.txt'), '--output', older]);
  assert.equal(into.status, 0, '以前の版に書き出すなら、--output にそのファイルを指定する: ' + into.stderr);
});

test('以前の版を元に反映するとき、書き出し先のHTMLにだけある変更（作り直したときに加えたもの）が元に戻るなら、そのセルを注意で示す', () => {
  const folder = newFolder('undone');
  const html = buildBook(fixture('book'), folder, '表.html');
  const base = embeddedBook(html).revision;
  writeFileSync(join(folder, 'first.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; }), 'utf8');
  // ユーザーは同じ画面で編集を続ける（以前の版から書き出した状態に、1回目の編集も含まれる）。
  writeFileSync(join(folder, 'second.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].cells.B5.v = '2回目'; }), 'utf8');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'first.txt')]).status, 0);
  // その間に、Claude が反映した後のHTMLを取り出して直し、作り直す。
  const json = join(folder, 'edit.json');
  assert.equal(runBuild(['--base', html, '--output', json, '--extract-json']).status, 0);
  const edited = JSON.parse(readFileSync(json, 'utf8'));
  edited.sheets[0].cells.C4.v = 99;
  edited.sheets[0].colWidths[1] = 200;
  writeFileSync(json, JSON.stringify(edited), 'utf8');
  assert.equal(runBuild(['--input', json, '--output', html]).status, 0);
  const [older] = olderWith(folder, base);
  const result = runBuild(['--base', older, '--state', join(folder, 'second.txt'), '--output', html]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: この編集状態の元になった版（ユーザーが画面を開いたときの版）の後に、表\.html が作り直されています。その作り直しで変わったところは、反映すると元に戻ります: 明細!C4、シート「明細」の colWidths\[1\]。/m);
  assert.match(result.stderr, /自分（Claude）が加えた変更なら、反映した後にもう一度加えてください。前に反映した編集状態の変更なら、ユーザーが画面で戻したものなので、そのままにしてください/);
  assert.ok(!/明細!B4/.test(result.stderr), 'ユーザーが同じ画面で続けた編集は、元に戻るものではない');
  const cells = embeddedBook(html).sheets[0].cells;
  assert.deepEqual([cells.B4.v, cells.B5.v], ['1回目', '2回目']);
});

test('以前の版から書き出した状態をxlsxやJSONにするときも、作り直した後のHTMLにだけある変更が入らないことを注意で示す', () => {
  const {folder, older, second} = rebuiltAfterCopy('undone-outputs', w => { w.sheets[0].cells.C4.v = 99; });
  const json = runBuild(['--base', older, '--state', second, '--output', join(folder, '取り出し.json'), '--extract-json']);
  assert.equal(json.status, 0, json.stderr);
  assert.match(json.stderr, /^注意: この編集状態の元になった版の後に、表\.html が作り直されています。その作り直しで変わったところは、このJSONには入っていません: 明細!C4。/m);
  const xlsx = runBuild(['--base', older, '--state', second, '--xlsx', join(folder, '表.xlsx')]);
  assert.equal(xlsx.status, 0, xlsx.stderr);
  assert.match(xlsx.stderr, /このxlsxには入っていません: 明細!C4。/);
});

test('作り直しで消したシート・並べ替え・開くシート・列幅の一部・結合と、画面の編集と作り直しの両方で変わったセルも、注意で示す', () => {
  const {html, older, second} = rebuiltAfterCopy('undone-kinds', w => {
    w.sheets = [w.sheets[0]];  // 予備のシートを消す
    w.activeSheet = 'main';
    w.sheets[0].colWidths[1] = 200;
    w.sheets[0].merges.push({r1: 20, c1: 5, r2: 20, c2: 6});
    w.sheets[0].cells.B6 = {t: 's', v: 'AIの値'};
  }, w => { w.sheets[0].cells.B4.v = '1回目'; w.sheets[0].colWidths[3] = 150; w.sheets[0].cells.B6 = {t: 's', v: 'ユーザーの値'}; });
  const result = runBuild(['--base', older, '--state', second, '--output', html]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /反映すると元に戻ります: .*シート「予備」（作り直したときに消したもの）/);
  assert.match(result.stderr, /シート「明細」の colWidths\[1\]/, 'ユーザーが別の列の幅を変えていても、作り直しで変えた列の幅は元に戻る');
  assert.match(result.stderr, /シート「明細」の結合 F21:G21（作り直したときに足したもの）/);
  assert.match(result.stderr, /^注意: 画面の編集と作り直しの両方で変わったセルがあり、ユーザーの値にしました: 明細!B6。/m);
  assert.equal(embeddedBook(html).sheets[0].cells.B6.v, 'ユーザーの値');
});

test('名前に年月日時分秒が付いていても、元の名前のファイルが別の表なら、以前の版とはみなさない', () => {
  const folder = newFolder('not-older');
  const dated = buildBook(fixture('book'), folder, '表_20260101000000.html');
  buildBook(blankBook({id: 'other-book', title: '別の表'}), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(dated, w => { w.sheets[0].cells.B4.v = '反映'; }), 'utf8');
  const result = runBuild(['--base', dated, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(embeddedBook(dated).sheets[0].cells.B4.v, '反映');
});
