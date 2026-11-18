// scripts/build.py：検査、埋め込み、リビジョン、以前の内容を残す書き出し、状態テキストの反映、--extract-json、--xlsx。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {after, before, test} from 'node:test';
import {blankBook, buildBook, Core, embeddedBook, findPython, fixture, removeDir, runBuild, SKILL_ROOT, tempDir} from './lib/workbook.mjs';
import {readXlsx} from './lib/xlsx.mjs';

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
const listXlsx = folder => readdirSync(folder).filter(n => n.endsWith('.xlsx')).sort();
const xlsxOf = path => readXlsx(readFileSync(path));

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

test('--xlsx は、生成したHTMLに入っている表からxlsxを作り、HTMLは変えない', () => {
  const folder = newFolder('xlsx-base');
  const html = buildBook(fixture('book'), folder, '表.html');
  const original = readFileSync(html, 'utf8');
  const xlsx = join(folder, '表.xlsx');
  const result = runBuild(['--base', html, '--xlsx', xlsx]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /表\.xlsx/);
  const x = xlsxOf(xlsx);
  assert.deepEqual(x.sheets.map(s => s.name), ['明細', '予備']);
  assert.deepEqual([x.sheets[0].cells.B4.value, x.sheets[0].cells.E9.formula], ['保守作業', 'E7+E8']);
  assert.equal(readFileSync(html, 'utf8'), original);
  assert.deepEqual(listHtml(folder), ['表.html'], 'HTMLの以前の内容のファイルも作らない');
});

test('--xlsx に --state を足すと、ブラウザーで編集した後の表でxlsxを作り、HTMLは変えない', () => {
  const folder = newFolder('xlsx-state');
  const html = buildBook(fixture('book'), folder, '表.html');
  const original = readFileSync(html, 'utf8');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = 'ブラウザーで直した'; w.sheets[0].cells.C4.v = 5; }), 'utf8');
  const xlsx = join(folder, '表.xlsx');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--xlsx', xlsx]);
  assert.equal(result.status, 0, result.stderr);
  const cells = xlsxOf(xlsx).sheets[0].cells;
  assert.deepEqual([cells.B4.value, cells.C4.value, cells.E4.value], ['ブラウザーで直した', 5, 60000], '数式の値も、編集した後の値で計算する');
  assert.equal(readFileSync(html, 'utf8'), original);
  assert.deepEqual(listHtml(folder), ['表.html']);
});

test('--xlsx は、ブックJSON（--input）からも作れ、同じ名前のxlsxがあると年月日時分秒を付けた名前で残す', () => {
  const folder = newFolder('xlsx-input');
  const json = join(folder, 'b.json'), xlsx = join(folder, '一覧.xlsx');
  const outputs = [];
  for (const v of ['1回目', '2回目']) {
    writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 's', v}}}})), 'utf8');
    const result = runBuild(['--input', json, '--xlsx', xlsx]);
    assert.equal(result.status, 0, result.stderr);
    outputs.push(result.stdout);
  }
  const names = listXlsx(folder);
  assert.equal(names.length, 2, '最新の1つと、以前の1つ');
  const previous = names.find(n => /^一覧_\d{14}\.xlsx$/.test(n));
  assert.ok(previous, names.join(', '));
  assert.match(outputs[1], new RegExp('以前の内容: .*' + previous.replace('.', '\.')));
  assert.equal(xlsxOf(xlsx).sheets[0].cells.A1.value, '2回目');
  assert.equal(xlsxOf(join(folder, previous)).sheets[0].cells.A1.value, '1回目');
});

test('--xlsx でも、ブックIDやリビジョンが合わない状態テキストでは、何も書き出さずに止まる', () => {
  const folder = newFolder('xlsx-mismatch');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, (w, state) => { state.workbookId = 'other'; }), 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--xlsx', join(folder, '表.xlsx')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /ブックIDまたはリビジョンが一致しません/);
  assert.deepEqual(listXlsx(folder), []);
});

test('--xlsx は --output・--extract-json と一緒には使えず、数式として書けない式があると注意を出してから書き出す', () => {
  const folder = newFolder('xlsx-args');
  const json = join(folder, 'b.json'), xlsx = join(folder, 'b.xlsx');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 'f', f: '=SUM(A2:A3'}}}})), 'utf8');
  for (const extra of [['--output', join(folder, 'b.html')], ['--extract-json']]) {
    const result = runBuild(['--input', json, '--xlsx', xlsx, ...extra]);
    assert.equal(result.status, 2, extra.join(' '));
    assert.match(result.stderr, /--xlsx/);
  }
  assert.deepEqual(listXlsx(folder), []);
  const html = buildBook(blankBook(), folder, 'base.html');
  const alone = runBuild(['--base', html]);
  assert.equal(alone.status, 2, 'xlsxを作らないときは、--base に --state が要る');
  const result = runBuild(['--input', json, '--xlsx', xlsx]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: .*A1.*=SUM\(A2:A3/m);
  assert.equal(xlsxOf(xlsx).sheets[0].cells.A1.value, '=SUM(A2:A3');
});

test('書き出し先のxlsxがExcelで開かれていて書き換えられないときは、閉じるよう伝えて止まり、元のファイルを残す', () => {
  // Windowsは、Excelが開いているファイルの名前の変更や置き換えを許さない（PermissionError）。その状態を、
  // 書き出しの2つの段階（元のファイルの名前を変える・新しい中身に置き換える）のそれぞれで失敗させて作る。
  for (const [step, patch] of [['名前の変更', 'rename'], ['置き換え', 'replace']]) {
    const folder = newFolder('xlsx-locked-' + patch), json = join(folder, 'b.json'), xlsx = join(folder, '一覧.xlsx');
    writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 's', v: '新しい表'}}}})), 'utf8');
    writeFileSync(xlsx, 'Excelで開いている元のファイル', 'utf8');
    const code = 'import pathlib, sys\nsys.dont_write_bytecode = True\nsys.path.insert(0, sys.argv[1])\nimport build\n' +
      `real = pathlib.Path.${patch}\ndef locked(self, target):\n    if str(self).endswith(".xlsx") or str(self).endswith(".tmp"):\n        raise PermissionError(13, "使用中", str(self))\n    return real(self, target)\n` +
      `pathlib.Path.${patch} = locked\nsys.argv = ["build.py", "--input", sys.argv[2], "--xlsx", sys.argv[3]]\nsys.exit(build.main())`;
    const py = findPython();
    const run = spawnSync(py.command, [...py.args, '-c', code, join(SKILL_ROOT, 'scripts'), json, xlsx], {encoding: 'utf8', timeout: 60000,
      env: {...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1'}});
    assert.equal(run.status, 1, step + ': ' + run.stderr);
    assert.match(run.stderr, /書き出し先のファイルを書き換えられません。Excelなどで開いているなら、閉じてからやり直してください/, step);
    assert.equal(readFileSync(xlsx, 'utf8'), 'Excelで開いている元のファイル', step + 'で失敗しても、元のファイルはその名前のまま');
    assert.deepEqual(readdirSync(folder).sort(), ['b.json', '一覧.xlsx'], step + 'で失敗しても、書きかけのファイルや名前を変えた元のファイルを残さない');
  }
});

test('--from-xlsx は、Excelのファイルから表のHTMLを作り、移せなかったものを「注意:」で知らせる', () => {
  const folder = newFolder('from-xlsx');
  const html = join(folder, '見積.html');
  const result = runBuild(['--from-xlsx', join(SKILL_ROOT, 'tests', 'fixtures', 'excel-made.xlsx'), '--output', html]);
  assert.equal(result.status, 0, result.stderr);
  const book = embeddedBook(html);
  assert.deepEqual(book.sheets.map(s => s.name), ['見積', '単価表']);
  assert.equal(book.sheets[0].cells.E9.f, '=E7+E8');
  assert.match(result.stderr, /^注意: シート「見積」：次の数式は、HTMLで計算できないので/m);
  assert.match(result.stderr, /^注意: シート「見積」：条件付き書式は移せない/m);
  const json = join(folder, 'book.json');
  const extracted = runBuild(['--from-xlsx', join(SKILL_ROOT, 'tests', 'fixtures', 'excel-made.xlsx'), '--output', json, '--extract-json']);
  assert.equal(extracted.status, 0, extracted.stderr);
  assert.equal(JSON.parse(readFileSync(json, 'utf8')).sheets[0].cells.A1.v, '御見積書', '--extract-json なら読み込んだ表をJSONで書く');
});

test('--from-xlsx は他の入力と一緒には使えず、xlsx として読めないファイルでは理由を出して止まる', () => {
  const folder = newFolder('from-xlsx-bad');
  writeFileSync(join(folder, 'b.json'), JSON.stringify(blankBook()), 'utf8');
  const both = runBuild(['--from-xlsx', join(SKILL_ROOT, 'tests', 'fixtures', 'excel-made.xlsx'), '--input', join(folder, 'b.json'), '--output', join(folder, 'x.html')]);
  assert.equal(both.status, 2);
  assert.match(both.stderr, /--from-xlsx/);
  const noOutput = runBuild(['--from-xlsx', join(SKILL_ROOT, 'tests', 'fixtures', 'excel-made.xlsx')]);
  assert.equal(noOutput.status, 2);
  assert.match(noOutput.stderr, /--from-xlsx には --output を指定してください/);
  writeFileSync(join(folder, 'not.xlsx'), 'これはxlsxではない', 'utf8');
  const bad = runBuild(['--from-xlsx', join(folder, 'not.xlsx'), '--output', join(folder, 'x.html')]);
  assert.equal(bad.status, 1);
  assert.match(bad.stderr, /xlsx（ZIP）として読めません/);
  assert.deepEqual(listHtml(folder), []);
});
