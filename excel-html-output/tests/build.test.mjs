// scripts/build.py：ブックJSONからHTMLを作る、データの埋め込み、リビジョン、以前の内容を残す書き出し、
// 引数とヘルプ、出力の文字コード。ほかのテストは、検査と注意が build-checks、状態テキストの反映と
// ブックID・リビジョンの照合が build-state、以前の版を元に反映するときが build-older、
// --extract-json・--xlsx・--from-xlsx での書き出しが build-outputs にある（共通の部品は lib/build.mjs）。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {existsSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {test} from 'node:test';
import {dir, listHtml, newFolder, stateText, useTempDir} from './lib/build.mjs';
import {blankBook, buildBook, embeddedBook, findPython, fixture, runBuild, SKILL_ROOT} from './lib/workbook.mjs';

useTempDir();

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

test('引数の組み合わせが足りない・出力先がフォルダーのときは、理由を出して止まる', () => {
  const missing = runBuild([]);
  assert.equal(missing.status, 2);
  assert.match(missing.stderr, /--input、または --base と --state の組を指定してください（--extract-json なら --base だけでもよい）/);
  const folder = newFolder('outdir');
  writeFileSync(join(folder, 'b.json'), JSON.stringify(blankBook()), 'utf8');
  const toDir = runBuild(['--input', join(folder, 'b.json'), '--output', dir]);
  assert.equal(toDir.status, 1);
  assert.match(toDir.stderr, /出力先がフォルダーです/);
  assert.ok(existsSync(dir));
});

test('書き出し先にある、このスキルのものでないファイル（UTF-8 でないHTMLなど）も、年月日時分秒を付けた名前で残してから書き出す', () => {
  const folder = newFolder('foreign-output');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '反映'; }), 'utf8');
  const sjis = Buffer.from([0x95, 0x5c, 0x0d, 0x0a]);  // 「表」と改行の Shift_JIS
  for (const [name, content] of [['sjis.html', sjis], ['array.html', '<script id="workbook-data" type="application/json">[1,2]</script>']]) {
    const output = join(folder, name);
    writeFileSync(output, content);
    const result = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--output', output]);
    assert.equal(result.status, 0, name + ': ' + result.stderr);
    assert.equal(embeddedBook(output).sheets[0].cells.B4.v, '反映', name);
    const kept = readdirSync(folder).find(n => new RegExp('^' + name.replace('.html', '') + '_\\d{14}\\.html$').test(n));
    assert.ok(kept, name + ' を残す');
    assert.deepEqual(readFileSync(join(folder, kept)), Buffer.from(content), name);
  }
});

test('build.py の出力は、端末の文字コードにかかわらず UTF-8 で書く（Claude Code はコマンドの出力を UTF-8 として読む）', () => {
  const folder = newFolder('utf8-output'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {values: [{at: 'A1', rows: [['数量', '2']]}]}})), 'utf8');
  const py = findPython(), env = {...process.env};
  delete env.PYTHONIOENCODING;
  delete env.PYTHONUTF8;
  const result = spawnSync(py.command, [...py.args, join(SKILL_ROOT, 'scripts', 'build.py'), '--input', json, '--output', join(folder, '表—é.html')], {timeout: 60000, env});
  assert.equal(result.status, 0, result.stderr.toString('utf8'));
  assert.match(result.stderr.toString('utf8'), /^注意: 数値に見える文字列が、文字のセルになります/m);
  assert.match(result.stdout.toString('utf8'), /表—é\.html/, 'cp932 にない文字のパスも書ける');
});

test('--extract-json のヘルプは、どの引数と組み合わせて何を書き出すかを示す', () => {
  const help = runBuild(['--help']);
  assert.equal(help.status, 0);
  assert.match(help.stdout.replace(/\s+/g, ''), /--inputなら省略形を広げた後、--baseなら--stateを反映した後（--stateがなければそのHTMLの表）、--from-xlsxなら読み込んだもの/);
  assert.match(help.stdout, /--keep-user-content/);
});
