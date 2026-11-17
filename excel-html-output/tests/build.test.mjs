// scripts/build.py：検査、埋め込み、リビジョン、以前の内容を残す書き出し、状態テキストの反映、--extract-json。
import assert from 'node:assert/strict';
import {existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {after, before, test} from 'node:test';
import {blankBook, buildBook, Core, embeddedBook, fixture, removeDir, runBuild, tempDir} from './lib/workbook.mjs';

let dir;
before(() => { dir = tempDir(); });
after(() => removeDir(dir));

/**
 * A state text of 「AIに渡す」 for the HTML at htmlPath, in format 'toon' (what the page gives by default) or 'json',
 * after edit(workbook, state) changes a copy of its workbook or the state around it.
 */
function stateText(htmlPath, edit, format = 'toon') {
  const base = embeddedBook(htmlPath), workbook = structuredClone(base);
  const state = {schema: 'excel-html-state/v1', workbookId: base.id, baseRevision: base.revision, exportedAt: new Date().toISOString(), workbook};
  edit(workbook, state);
  return Core.stateText(state, format);
}
const newFolder = name => { const folder = join(dir, name); mkdirSync(folder, {recursive: true}); return folder; };
const listHtml = folder => readdirSync(folder).filter(n => n.endsWith('.html')).sort();

test('ブックJSONから、データを埋め込んだ単一のHTMLを作り、リビジョンを内容から決める', () => {
  const html = buildBook(fixture('book'), dir, 'revision.html');
  const source = readFileSync(html, 'utf8');
  const embedded = embeddedBook(html);
  assert.match(embedded.revision, /^sha256:[0-9a-f]{64}$/);
  assert.equal(embedded.sheets[0].cells.E9.f, '=E7+E8');
  assert.ok(!/<script[^>]+src=|<link[^>]+href=/i.test(source), '外部のファイルや通信を使わない');
  assert.ok(source.includes('window.ExcelHTMLCore'), 'core.js と app.js を埋め込む');
  const same = buildBook(fixture('book'), dir, 'revision-same.html');
  assert.equal(embeddedBook(same).revision, embedded.revision, '同じ内容なら同じリビジョン');
  const changed = fixture('book');
  changed.sheets[0].cells.B4.v = '変更';
  assert.notEqual(embeddedBook(buildBook(changed, dir, 'revision-changed.html')).revision, embedded.revision, '内容が変わればリビジョンも変わる');
});

test('セルの文字や題名は、埋め込み先のスクリプトやHTMLを壊さない', () => {
  const book = blankBook({title: '<b>題名</b> & "x"', sheet: {cells: {
    A1: {t: 's', v: '</script><script>alert(1)</script>'}, A2: {t: 's', v: '/*__APP__*/ /*__CORE__*/ /*__TOON__*/ __TITLE__'}, A3: {t: 's', v: '行区切り 段落区切り '},
  }}});
  const html = buildBook(book, dir, 'escape.html');
  const source = readFileSync(html, 'utf8');
  assert.ok(source.includes('<title>&lt;b&gt;題名&lt;/b&gt; &amp; &quot;x&quot;</title>'));
  assert.ok(!source.includes('</script><script>alert(1)'), '値の </script> で埋め込みが終わらない');
  const cells = embeddedBook(html).sheets[0].cells;
  assert.equal(cells.A1.v, '</script><script>alert(1)</script>');
  assert.equal(cells.A2.v, '/*__APP__*/ /*__CORE__*/ /*__TOON__*/ __TITLE__', '置換の印に似た文字は置き換えない');
  assert.equal(cells.A3.v, '行区切り 段落区切り ');
  assert.equal(source.split('window.ExcelHTMLCore,$=').length, 2, 'app.js は1回だけ埋め込む');
});

test('同じ名前の出力先があると、年月日時分秒を付けた名前で残してから書き出し、続けて作っても上書きしない', () => {
  const folder = newFolder('keep');
  buildBook(blankBook(), folder, '一覧.html');
  for (const v of ['2回目', '3回目']) buildBook(blankBook({sheet: {cells: {A1: {t: 's', v}}}}), folder, '一覧.html');
  const names = listHtml(folder);
  assert.equal(names.length, 3, '最新の1つと、以前の2つ');
  assert.ok(names.includes('一覧.html'));
  assert.equal(names.filter(n => /^一覧_\d{14}\.html$/.test(n)).length, 2, '同じ秒に作っても名前がぶつからない');
  assert.equal(embeddedBook(join(folder, '一覧.html')).sheets[0].cells.A1.v, '3回目');
});

test('検査を通らないブックは、エラーを出して何も書き出さない', () => {
  const folder = newFolder('invalid');
  const book = blankBook({sheet: {cells: {A1: {t: 's', v: 'x', s: {borderTop: '1px groove #000000'}}}}});
  writeFileSync(join(folder, 'bad.json'), JSON.stringify(book), 'utf8');
  const result = runBuild(['--input', join(folder, 'bad.json'), '--output', join(folder, 'bad.html')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /ブック形式エラー: borderTop/);
  assert.deepEqual(listHtml(folder), []);
});

for (const [format, label] of [['toon', 'TOON'], ['json', 'JSON']]) {
  test(`状態テキスト（${label}）の反映：ID違い・途中で切れたテキスト・見出しのないテキストは拒み、元のファイルを変えない`, () => {
    const folder = newFolder('state-' + format);
    const html = buildBook(fixture('book'), folder, '表.html');
    const original = readFileSync(html, 'utf8');
    const good = stateText(html, w => { w.sheets[0].cells.B4.v = '反映'; }, format);
    const lines = good.split('\n'), cut = format === 'toon' ? /EXCEL_HTML_STATE_END/ : /JSON/;
    const cases = {
      'ブックID違い': [stateText(html, (w, state) => { state.workbookId = 'other'; }, format), /ブックIDまたはリビジョンが一致しません/],
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

test('--extract-json は反映後のブックをJSONで出し、そのJSONから同じ表を作り直せる', () => {
  const folder = newFolder('extract');
  const html = buildBook(fixture('book'), folder, '表.html');
  const original = readFileSync(html, 'utf8');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.C4.v = 9; }), 'utf8');
  const json = join(folder, 'edit.json');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--output', json, '--extract-json']);
  assert.equal(result.status, 0, result.stderr);
  const workbook = JSON.parse(readFileSync(json, 'utf8'));
  assert.equal(workbook.sheets[0].cells.C4.v, 9);
  assert.equal(readFileSync(html, 'utf8'), original, '--extract-json では元のHTMLを書き換えない');
  assert.deepEqual(listHtml(folder), ['表.html']);
  workbook.sheets[0].cells.B4.v = 'AIが直した値';
  writeFileSync(json, JSON.stringify(workbook), 'utf8');
  const rebuilt = runBuild(['--input', json, '--output', html]);
  assert.equal(rebuilt.status, 0, rebuilt.stderr);
  const cells = embeddedBook(html).sheets[0].cells;
  assert.deepEqual([cells.C4.v, cells.B4.v], [9, 'AIが直した値']);
});

test('引数の組み合わせが足りない・出力先がフォルダーのときは、理由を出して止まる', () => {
  const missing = runBuild([]);
  assert.equal(missing.status, 2);
  assert.match(missing.stderr, /--input、または--baseと--stateの組を指定してください/);
  const folder = newFolder('outdir');
  writeFileSync(join(folder, 'b.json'), JSON.stringify(blankBook()), 'utf8');
  const toDir = runBuild(['--input', join(folder, 'b.json'), '--output', dir]);
  assert.equal(toDir.status, 1);
  assert.match(toDir.stderr, /出力先がフォルダーです/);
  assert.ok(existsSync(dir));
});
