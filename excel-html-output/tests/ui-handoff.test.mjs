// 編集内容をAIに渡す：状態テキスト、コピーとエクスポート、未書き出しの警告、build.py での反映の往復。
import assert from 'node:assert/strict';
import {copyFileSync, mkdirSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {basename, join} from 'node:path';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {Core, embeddedBook, fixture, removeDir, runBuild, tempDir} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');

const beforeUnloadBlocked = x => x.evaluate(() => {
  const event = new Event('beforeunload', {cancelable: true});
  window.dispatchEvent(event);
  return event.defaultPrevented;
});

const chosenFormat = x => x.evaluate(() => document.querySelector('input[name="handoff-format"]:checked')?.value);

test('Ctrl+S で、ブック全体と表示位置を含む状態テキストを、既定ではTOONで、元のブックIDとリビジョン付きで出す', async t => {
  const x = await env.open(t, book);
  const embedded = embeddedBook(x.html);
  await x.enter('B4', '点検作業');
  await x.select('C4:D5');
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  assert.equal(await chosenFormat(x), 'toon', '形式は TOON が選ばれている');
  assert.equal(await x.evaluate(() => document.activeElement?.id), 'handoff-text', '開いたときは、テキストの欄にフォーカスがある');
  const text = await x.value('handoff-text');
  assert.ok(text.startsWith('EXCEL_HTML_STATE_V1 TOON\n'), '1行目で形式を示す');
  assert.ok(text.endsWith('\nEXCEL_HTML_STATE_END'), '最後の行で、テキストが切れていないことを示す');
  const state = Core.readStateText(text);
  assert.equal(state.schema, 'excel-html-state/v1');
  assert.equal(state.workbookId, 'excel-html-test-fixture');
  assert.equal(state.baseRevision, embedded.revision);
  assert.match(state.baseRevision, /^sha256:[0-9a-f]{64}$/);
  assert.ok(!Number.isNaN(Date.parse(state.exportedAt)));
  assert.doesNotThrow(() => Core.validate(state.workbook), 'そのままブックとして検査を通る');
  assert.equal(state.workbook.sheets[0].cells.B4.v, '点検作業');
  assert.deepEqual(state.workbook.view.selection, {r1: 3, c1: 2, r2: 4, c2: 3}, '選択範囲も渡す');
  assert.deepEqual(state.workbook.view.active, [3, 2]);
  assert.equal(state.workbook.sheets.length, 2, '全シートを渡す');
});

test('ダイアログには、渡し方の指示と形式の名前だけを出す', async t => {
  const x = await env.open(t, book);
  await x.click('#handoff');
  await x.waitDialog('handoff-dialog');
  const shown = await x.evaluate(() => ({
    guide: document.querySelector('#handoff-dialog p').textContent,
    formats: [...document.querySelectorAll('#handoff-dialog .handoff-format label')].map(label => label.textContent),
    result: document.querySelector('#handoff-result').textContent,
  }));
  assert.equal(shown.guide, '下のテキストと元のHTMLをClaude Codeに渡し、「この編集状態を反映してHTMLを再生成して」と伝えてください。');
  assert.deepEqual(shown.formats, ['TOON', 'JSON'], '形式の名前に、良し悪しや既定かどうかを書き添えない');
  assert.equal(shown.result, '', '開いただけでは、テキストの下に何も出さない');
});

test('形式で JSON を選ぶと同じ状態をJSONで出し、コピーもJSONになる。TOON に戻すと元のテキストに戻る', async t => {
  const x = await env.open(t, book);
  await x.enter('B4', '点検作業');
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  const toon = await x.value('handoff-text');
  await x.click('#handoff-json');
  assert.equal(await chosenFormat(x), 'json');
  const json = await x.value('handoff-text');
  assert.ok(json.startsWith('EXCEL_HTML_STATE_V1\n{'), '1行目の後はJSON');
  assert.deepEqual(JSON.parse(json.slice(json.indexOf('\n') + 1)), Core.readStateText(toon), 'TOON と同じ状態（書き出した時刻も同じ）');
  await x.click('#copy-state');
  await x.waitFor(() => window.__testClipboard.writes > 0, {message: '状態テキストのコピー'});
  const [entry] = await x.clipboard();
  assert.equal(entry.text, json, '選んだ形式のテキストをコピーする');
  await x.click('#handoff-toon');
  assert.equal(await x.value('handoff-text'), toon, 'TOON に戻すと、同じ時点の TOON に戻る');
});

test('JSON で書き出した状態テキストも、build.py で反映できる', async t => {
  const dir = tempDir();
  t.after(() => removeDir(dir));
  const html = join(dir, '明細表.html');
  copyFileSync(env.build(book), html);
  const x = await env.open(t, html);
  await x.enter('B4', 'JSONで渡した編集');
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  await x.click('#handoff-json');
  const edits = join(dir, 'edits.txt');
  writeFileSync(edits, await x.value('handoff-text'), 'utf8');
  const result = runBuild(['--base', html, '--state', edits]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(embeddedBook(html).sheets[0].cells.B4.v, 'JSONで渡した編集');
});

test('状態テキストを書き出せないときは、理由をダイアログに出し、空のテキストをコピーしたことにしない', async t => {
  const x = await env.open(t, book);
  await x.enter('B4', '点検作業');
  // A fault the page cannot produce by itself today: the TOON encoder refusing the state.
  await x.evaluate(() => { window.ExcelHTMLToon.encode = () => { throw new Error('TOONに書き出せない文字（試験用）があります'); }; });
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  assert.equal(await x.value('handoff-text'), '');
  assert.match(await x.textOf('#handoff-result'), /TOONに書き出せない文字（試験用）があります/);
  await x.click('#copy-state');
  assert.equal(await x.evaluate(() => window.__testClipboard.writes), 0, 'クリップボードに何も書かない');
  assert.equal(await x.textOf('#dirty-state'), '編集あり · AIへ書き出し', '書き出し済みにしない');
  await x.click('#handoff-json');
  assert.ok((await x.value('handoff-text')).startsWith('EXCEL_HTML_STATE_V1\n{'), 'JSON なら書き出せる');
});

test('「テキストをコピー」は状態テキストをクリップボードへ入れ、未書き出しの印を消す', async t => {
  const x = await env.open(t, book);
  assert.equal(await x.textOf('#dirty-state'), '書き出し済み / 元の状態');
  assert.equal(await beforeUnloadBlocked(x), false, '編集していなければ、閉じるときに引き止めない');
  await x.enter('B4', '点検作業');
  assert.equal(await x.textOf('#dirty-state'), '編集あり · AIへ書き出し');
  assert.equal(await beforeUnloadBlocked(x), true, '書き出していない編集があると、閉じる前に確認する');
  await x.click('#handoff');
  await x.click('#copy-state');
  await x.waitFor(() => window.__testClipboard.writes > 0, {message: '状態テキストのコピー'});
  const [entry] = await x.clipboard();
  assert.equal(entry.text, await x.value('handoff-text'));
  assert.equal(await x.textOf('#dirty-state'), '編集状態をコピー済み');
  assert.equal(await x.textOf('#handoff-result'), '編集状態をコピーしました');
  assert.equal(await beforeUnloadBlocked(x), false);
});

test('「テキストをエクスポート」は、ブック名-edits.txt という名前で同じ状態テキストを保存させる', async t => {
  const x = await env.open(t, book);
  // Inside the file's own temp folder, which is removed after the browser has closed and released the file.
  const dir = join(env.dir, 'downloads');
  mkdirSync(dir, {recursive: true});
  await x.enter('B4', '点検作業');
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  const file = await x.page.download(dir, () => x.click('#download-state'));
  assert.equal(file.suggestedFilename, 'テスト用ブック-edits.txt');
  assert.equal(readFileSync(file.path, 'utf8'), await x.value('handoff-text'));
  assert.equal(await x.textOf('#dirty-state'), '編集状態を書き出し済み');
  assert.equal(await x.textOf('#handoff-result'), 'テキストのダウンロードを開始しました');
  assert.equal(await beforeUnloadBlocked(x), false);
});

test('画面で編集した状態テキストを build.py で反映すると、同じ名前のHTMLに編集がすべて入り、以前の内容も残る', async t => {
  const dir = tempDir();
  t.after(() => removeDir(dir));
  const html = join(dir, '明細表.html');
  copyFileSync(env.build(book), html);
  const original = readFileSync(html, 'utf8');
  const x = await env.open(t, html);
  await x.enter('B4', '点検作業');
  await x.clickCell('C4');
  await x.press('Control+b');
  await x.select('F12:G13');
  await x.action('merge');
  await x.clickCell('D4');
  await x.click('[data-tab="cells"]');
  await x.setControl('column-width', '130');
  await x.click('#add-sheet');
  const edits = join(dir, 'edits.txt'), {raw} = await x.state();
  assert.ok(raw.startsWith('EXCEL_HTML_STATE_V1 TOON\n'), '既定の TOON で渡す');
  writeFileSync(edits, raw, 'utf8');

  const result = runBuild(['--base', html, '--state', edits]);
  assert.equal(result.status, 0, result.stderr);
  const previous = readdirSync(dir).filter(name => /^明細表_\d{14}\.html$/.test(name));
  assert.equal(previous.length, 1, '書き換える前の内容を、年月日時分秒を付けた名前で残す');
  assert.equal(readFileSync(join(dir, previous[0]), 'utf8'), original, '残したファイルは元のまま');
  assert.match(result.stdout, new RegExp('以前の内容: .*' + previous[0]));

  const rebuilt = embeddedBook(html);
  const main = rebuilt.sheets[0];
  assert.equal(rebuilt.id, 'excel-html-test-fixture', 'ブックIDは変えない');
  assert.notEqual(rebuilt.revision, embeddedBook(join(dir, previous[0])).revision, 'リビジョンは新しくなる');
  assert.equal(main.cells.B4.v, '点検作業');
  assert.equal(main.cells.C4.s.bold, true);
  assert.deepEqual(main.merges.at(-1), {r1: 11, c1: 5, r2: 12, c2: 6});
  assert.equal(main.colWidths[3], 130);
  assert.equal(rebuilt.sheets.length, 3);
  assert.equal(main.cells.E9.f, '=E7+E8', '編集していないセルはそのまま');

  const reopened = await env.open(t, html);
  assert.equal(await reopened.textOf('#sheet-tabs [aria-selected="true"]'), 'Sheet1', '編集を終えたときに開いていたシートで開く');
  await reopened.click('#sheet-tabs [data-sheet="main"]');
  assert.equal(await reopened.text('B4'), '点検作業', '作り直したHTMLを開くと編集が見える');
  assert.equal(await reopened.text('E9'), '¥83,600');
});

test('元のリビジョンと合わない状態テキストは反映を拒み、ファイルの名前も中身も変えない', async t => {
  const dir = tempDir();
  t.after(() => removeDir(dir));
  const html = join(dir, '明細表.html');
  copyFileSync(env.build(book), html);
  const x = await env.open(t, html);
  await x.enter('B4', '1回目の編集');
  const edits = join(dir, 'edits.txt');
  writeFileSync(edits, (await x.state()).raw, 'utf8');
  assert.equal(runBuild(['--base', html, '--state', edits]).status, 0);
  const files = readdirSync(dir).sort();
  const rebuilt = readFileSync(html, 'utf8');
  // The same state text no longer matches: the HTML now has a new revision.
  const again = runBuild(['--base', html, '--state', edits]);
  assert.equal(again.status, 1);
  assert.match(again.stderr, /リビジョンが一致しません/);
  assert.deepEqual(readdirSync(dir).sort(), files, '新しいファイルを作らない');
  assert.equal(readFileSync(html, 'utf8'), rebuilt, '中身も変えない');
  // The documented recovery: use the kept previous version as --base and the original name as --output.
  const previous = files.find(name => /^明細表_\d{14}\.html$/.test(name));
  const recovered = runBuild(['--base', join(dir, previous), '--state', edits, '--output', html]);
  assert.equal(recovered.status, 0, recovered.stderr);
  assert.equal(embeddedBook(html).sheets[0].cells.B4.v, '1回目の編集');
  assert.ok(basename(previous).startsWith('明細表_'));
});
