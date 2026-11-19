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
  assert.equal(shown.guide, '「テキストをコピー」を押してから、Claude Codeに元のHTMLのファイル名と「コピーした編集状態を反映して」を伝えてください。' +
    'Claude Codeをこのパソコンで動かしているときは、クリップボードから読み取るので、貼り付けは要りません' +
    '（コピーした後に表を編集したときや、ほかのものをコピーしたときは、もう一度「テキストをコピー」を押してください）。' +
    'SSH・リモート開発・コンテナ・クラウドで動かしているときは、「テキストをエクスポート」で保存したファイルを渡すか、テキストを貼り付けてください。',
    'クリップボードから渡せるので、長いテキストをチャットに貼らなくてよいことと、そうできない場合を伝える');
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

test('状態テキストが大きいときは、チャットに貼らずに渡す方法を、画像のあるなしにかかわらず示す', async t => {
  const x = await env.open(t, book);
  // 画像のない表で、大きなテキストを作る（TOON の書き出しに、長い文字の項目を足す）。
  await x.evaluate(() => { const encode = window.ExcelHTMLToon.encode; window.ExcelHTMLToon.encode = v => encode(v) + '\nlong: "' + 'あ'.repeat(600000) + '"'; });
  await x.press('Control+s');
  await x.waitDialog('handoff-dialog');
  const result = await x.textOf('#handoff-result');
  assert.match(result, /^テキストが大きい（約\d+\.\dMB）ため、チャットに貼ると途中で切れることがあります。/);
  assert.match(result, /Claude Codeをこのパソコンで動かしているときは、「テキストをコピー」を押してから伝えれば、貼り付けずに渡せます。そうでなければ、「テキストをエクスポート」で保存したファイルを渡してください。$/);
  assert.ok(!/画像/.test(result), '画像のない表でも出るので、画像のことは書かない: ' + result);
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

test('元のリビジョンと合わない状態テキストは反映を拒み、以前の版を元にすれば反映でき、反映済みの状態の当て直しは止める', async t => {
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
  const previous = files.find(name => /^明細表_\d{14}\.html$/.test(name));
  assert.ok(basename(previous).startsWith('明細表_'));
  // The same state text again (copied before more edits, and not copied since): reflecting it changes nothing, while
  // the page still holds the later edits. It is refused as applied already, from the rebuilt HTML at once (whose book
  // it is) and from the kept previous version (whose version it was copied from), not reported done.
  for (const base of [html, join(dir, previous)]) {
    const again = runBuild(['--base', base, '--state', edits, '--output', html]);
    assert.equal(again.status, 1, base);
    assert.match(again.stderr, /すでに .*明細表\.html に反映されています/, base);
    assert.deepEqual(readdirSync(dir).sort(), files, '新しいファイルを作らない');
    assert.equal(readFileSync(html, 'utf8'), rebuilt, '中身も変えない');
  }
  // More edits in the same page, which still holds the book from before the rebuild: its state is of that version.
  await x.enter('B5', '2回目の編集');
  const more = join(dir, 'edits2.txt');
  writeFileSync(more, (await x.state()).raw, 'utf8');
  const mismatch = runBuild(['--base', html, '--state', more]);
  assert.equal(mismatch.status, 1);
  assert.match(mismatch.stderr, /別の版（以前の版など）から書き出されたものです/);
  assert.match(mismatch.stderr, new RegExp('リビジョンが一致するのは ' + previous.replace('.', '\\.') + ' です'), '一致する以前の版を名指しする');
  // The documented recovery: use the kept previous version as --base and the original name as --output.
  const recovered = runBuild(['--base', join(dir, previous), '--state', more, '--output', html]);
  assert.equal(recovered.status, 0, recovered.stderr);
  assert.deepEqual([embeddedBook(html).sheets[0].cells.B4.v, embeddedBook(html).sheets[0].cells.B5.v], ['1回目の編集', '2回目の編集']);
});
