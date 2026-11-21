// scripts/build.py の --extract-json・--xlsx・--from-xlsx。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {test} from 'node:test';
import {listHtml, newFolder, stateText, useTempDir, xlsxOf} from './lib/build.mjs';
import {blankBook, buildBook, embeddedBook, findPython, fixture, runBuild, SKILL_ROOT} from './lib/workbook.mjs';

useTempDir();

const listXlsx = folder => readdirSync(folder).filter(n => n.endsWith('.xlsx')).sort();

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
  assert.match(result.stderr, /この編集状態は別の表のものです/);
  assert.deepEqual(listXlsx(folder), []);
});

test('--xlsx は --output・--extract-json と一緒には使えず、数式として書けない式は、自分で書いたブックなら止まり、ユーザーの内容なら注意を出してから文字列として書き出す', () => {
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
  const own = runBuild(['--input', json, '--xlsx', xlsx]);
  assert.equal(own.status, 1, '自分で書いたブックの読めない式は直す');
  assert.match(own.stderr, /HTMLの数式として読めない式があります: Sheet1!A1 =SUM\(A2:A3/);
  assert.deepEqual(listXlsx(folder), []);
  const result = runBuild(['--input', json, '--xlsx', xlsx, '--keep-user-content']);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: .*A1.*=SUM\(A2:A3/m);
  assert.equal(result.stderr.match(/=SUM\(A2:A3/g).length, 1, '同じ式を2回知らせない');
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

test('--extract-json の --output にこのスキルのHTMLを指定すると、HTMLをJSONで置き換えずに止まる', () => {
  const folder = newFolder('json-over-html');
  const html = buildBook(fixture('book'), folder, '表.html'), before = readFileSync(html, 'utf8');
  const result = runBuild(['--base', html, '--output', html, '--extract-json']);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /--output の 表\.html は、このスキルのHTMLです。--extract-json はJSONを書き出すので、HTMLがJSONに置き換わります。--output にはJSONのファイル（例 .*表-edit\.json）を指定してください/);
  assert.equal(readFileSync(html, 'utf8'), before);
});
