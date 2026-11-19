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
  assert.match(result.stderr, /ブック形式エラー: sheets\[0\]\.cells\.A1\.s\.borderTop "1px groove #000000"（none か、線と色/);
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

test('省略形（values・formulas・styles）で書いたブックJSONは、セルごとの形に広げてから、HTML・xlsx・JSONにする', () => {
  const folder = newFolder('shorthand');
  const line = '1px solid #808080', json = join(folder, '見積.json');
  const book = blankBook({sheet: {rows: 10, cols: 5, cells: {A1: {t: 's', v: '見積', s: {bold: true}}},
    values: [{at: 'A2', rows: [['品名', '数量', '単価', '金額'], ['保守', 2, 1000], ['部品', 3, 500]]}],
    formulas: [{range: 'D3:D4', f: '=B3*C3'}, {range: 'D5', f: '=SUM(D3:D4)'}],
    styles: [{range: 'A2:D5', s: {borderBottom: line}}, {range: 'C3:D5', s: {numFmt: '¥#,##0'}}]}});
  delete book.sheets[0].merges;
  delete book.sheets[0].rowHeights;
  writeFileSync(json, JSON.stringify(book), 'utf8');
  const html = join(folder, '見積.html');
  const built = runBuild(['--input', json, '--output', html]);
  assert.equal(built.status, 0, built.stderr);
  const sheet = embeddedBook(html).sheets[0];
  assert.deepEqual(sheet.cells.D4, {t: 'f', f: '=B4*C4', s: {borderBottom: line, numFmt: '¥#,##0'}});
  assert.deepEqual(sheet.cells.A5, {t: 's', v: '', s: {borderBottom: line}});
  assert.deepEqual([sheet.merges, sheet.rowHeights], [[], []], '省いた項目は空になる');
  assert.deepEqual(['values', 'formulas', 'styles'].filter(k => k in sheet), [], '埋め込んだブックに省略形は残らない');
  const xlsx = join(folder, '見積.xlsx');
  const written = runBuild(['--input', json, '--xlsx', xlsx]);
  assert.equal(written.status, 0, written.stderr);
  const cells = xlsxOf(xlsx).sheets[0].cells;
  assert.deepEqual([cells.D4.formula, cells.D4.value, cells.D5.formula, cells.D5.value], ['B4*C4', 1500, 'SUM(D3:D4)', 3500]);
  const extracted = join(folder, '広げた.json');
  const extract = runBuild(['--input', json, '--output', extracted, '--extract-json']);
  assert.equal(extract.status, 0, extract.stderr);
  assert.deepEqual(JSON.parse(readFileSync(extracted, 'utf8')).sheets[0].cells, sheet.cells, '--extract-json で、広げた形を見られる');
});

test('--input のブックに罫線の食い違いがあると、場所と2本の線を示して止まり、何も書き出さない', () => {
  const folder = newFolder('conflict-input'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {
    A1: {t: 's', v: '明細', s: {borderBottom: '1px solid #808080'}}, A2: {t: 's', v: '合計', s: {borderTop: '2px solid #217346'}}}}})), 'utf8');
  for (const output of [['--output', join(folder, 'b.html')], ['--xlsx', join(folder, 'b.xlsx')]]) {
    const result = runBuild(['--input', json, ...output]);
    assert.equal(result.status, 1, output[0]);
    assert.match(result.stderr, /^エラー: ブック形式エラー: 罫線の食い違いが1か所あります/);
    assert.match(result.stderr, /Sheet1!A1 の下と A2 の上（1px solid #808080 \/ 2px solid #217346）/);
    assert.match(result.stderr, /片方のセルにだけ書くか、同じ線にそろえてください/);
  }
  assert.match(runBuild(['--input', json, '--output', join(folder, 'b.html')]).stderr, /--keep-user-content/, 'ユーザーが付けた線なら残せることも伝える');
  assert.deepEqual(readdirSync(folder), ['b.json']);
});

test('画面で編集した後の表やExcelのファイルから取り出したJSONは、--keep-user-content を付けると、ユーザーの内容（罫線の食い違い・Excelに入らない長さ・読めない数式）を注意にして作れる', () => {
  const folder = newFolder('keep-user-content');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => {
    w.sheets[0].cells.A20 = {t: 's', v: '上', s: {borderBottom: '1px solid #808080'}};
    w.sheets[0].cells.A21 = {t: 's', v: '下', s: {borderTop: '2px solid #000000'}};
    w.sheets[0].cells.A22 = {t: 's', v: 'あ'.repeat(32768)};
    w.sheets[0].cells.A23 = {t: 'f', f: '=A1*税率'};
  }), 'utf8');
  const json = join(folder, 'edit.json');
  assert.equal(runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--output', json, '--extract-json']).status, 0);
  const edited = JSON.parse(readFileSync(json, 'utf8'));
  edited.sheets[0].cells.B4.v = 'AIが直した値';
  writeFileSync(json, JSON.stringify(edited), 'utf8');
  const stopped = runBuild(['--input', json, '--output', html]);
  assert.equal(stopped.status, 1, '付けないと止まる');
  for (const reason of [/罫線の食い違いが1か所あります/, /Excelのセルに入らない長さがあります: 明細!A22/, /HTMLの数式として読めない式があります: 明細!A23 =A1\*税率/])
    assert.match(stopped.stderr, reason, '止まる理由を1回でまとめて示す');
  assert.equal(stopped.stderr.match(/^エラー:/gm).length, 1);
  const result = runBuild(['--input', json, '--output', html, '--keep-user-content']);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: 罫線の食い違いが1か所あります.*A20 の下と A21 の上/m);
  assert.match(result.stderr, /^注意: Excelのセルに入らない長さがあります: 明細!A22/m);
  assert.match(result.stderr, /^注意: HTMLの数式として読めない式があります: 明細!A23 =A1\*税率/m);
  assert.match(result.stderr, /ユーザーが付けた線（画面での編集や、元のxlsxの線）なら、そのまま残しました。取り出した後に自分で加えた線なら、直してください。/,
    '取り出した後に自分で加えたものもありうるので、そのときは直すよう示す');
  assert.match(result.stderr, /取り出した後に自分で書いた式なら、書き直してください。/);
  assert.deepEqual([embeddedBook(html).sheets[0].cells.B4.v, embeddedBook(html).sheets[0].cells.A21.s.borderTop], ['AIが直した値', '2px solid #000000']);
});

test('省略形の values に、数値や日付に見える文字列があると、文字のセルになることを注意で知らせる', () => {
  const folder = newFolder('number-like'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {values: [{at: 'A1', rows: [['品名', '数量', '単価'], ['保守', '2', '30,000']]}]}})), 'utf8');
  const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: 数値に見える文字列が、文字のセルになります（SUM などの計算に入りません）: Sheet1!B2 "2"、Sheet1!C2 "30,000"/m);
});

test('照合のスクリプトも、省略形で書いたブックJSON（--input）を、セルごとの形に広げてから使う', () => {
  const folder = newFolder('verify-input'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {values: [{at: 'A1', rows: [['数量', 2]]}], formulas: [{range: 'C1', f: '=B1*3'}],
    styles: [{range: 'A1:C1', s: {bold: true}}]}})), 'utf8');
  const code = 'import argparse, json, sys\nsys.dont_write_bytecode = True\nsys.path.insert(0, sys.argv[1])\nimport build, verify_excel_paste, verify_xlsx\n' +
    'a = build.read_input(sys.argv[2])\nb = verify_excel_paste.load_workbook(argparse.Namespace(input=sys.argv[2], html=None))\n' +
    'c = verify_xlsx.load_workbook(argparse.Namespace(input=sys.argv[2], base=None, state=None))\n' +
    'print(json.dumps([a["sheets"][0]["cells"], b["sheets"][0]["cells"], c["sheets"][0]["cells"]], ensure_ascii=False))';
  const py = findPython();
  const run = spawnSync(py.command, [...py.args, '-c', code, join(SKILL_ROOT, 'scripts'), json], {encoding: 'utf8', timeout: 60000,
    env: {...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1'}});
  assert.equal(run.status, 0, run.stderr);
  const expected = {A1: {t: 's', v: '数量', s: {bold: true}}, B1: {t: 'n', v: 2, s: {bold: true}}, C1: {t: 'f', f: '=B1*3', s: {bold: true}}};
  assert.deepEqual(JSON.parse(run.stdout), [expected, expected, expected], 'build.read_input と、verify_excel_paste.py・verify_xlsx.py の読み込み');
});

test('ブラウザーで編集した後の表（--state）の罫線の食い違いは、注意を出してから反映する', () => {
  const folder = newFolder('conflict-state');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => {
    w.sheets[0].cells.A20 = {t: 's', v: '上', s: {borderBottom: '1px solid #808080'}};
    w.sheets[0].cells.A21 = {t: 's', v: '下', s: {borderTop: '3px double #000000'}};
  }), 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: 罫線の食い違いが1か所あります.*A20 の下と A21 の上（1px solid #808080 \/ 3px double #000000）/m);
  assert.equal(embeddedBook(html).sheets[0].cells.A21.v, '下', 'ユーザーの編集は反映する');
});

test('状態テキストのブックは省略形を持てず、未対応の項目として止まる', () => {
  const folder = newFolder('shorthand-state');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].styles = [{range: 'A1', s: {bold: true}}]; }), 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /未対応の項目 sheets\[0\]\.styles/);
});

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

test('罫線の食い違いの文は、画面が描いている線と、見た目を変えずにそろえる方法も示す', () => {
  const folder = newFolder('conflict-wording'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 's', v: '上', s: {borderBottom: '1px solid #808080'}}, A2: {t: 's', v: '下', s: {borderTop: '2px solid #217346'}}}}})), 'utf8');
  const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.match(result.stderr, /画面では、上または左のセルの線（各組の1本目）を描いています/);
  assert.match(result.stderr, /見た目を変えずにそろえるなら、下または右のセルの向かい合う辺の線を消してください/);
});

test('HTMLで計算できない数式（計算できない関数・シートの外の参照・循環参照）は理由と直し方を、それを参照するだけの数式とExcelでもエラーになる数式は分けて、表示形式が文字列の数式とともに注意で知らせる', () => {
  const folder = newFolder('formula-cautions'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {rows: 30, cells: {
    A1: {t: 'f', f: '=SUMSQ(1,2)'}, A2: {t: 'f', f: '=SUM(B1:B40)'}, A3: {t: 'f', f: '=A3+1'}, A4: {t: 'f', f: '=IFERROR(1/0,0)'},
    A5: {t: 'f', f: '=1+1', s: {numFmt: '@'}}, A6: {t: 'f', f: '=A1*2'}, A7: {t: 'f', f: '=INDEX(B1:B3,5)'}}}})), 'utf8');
  const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: HTMLでは計算できない数式があります（Excelでは計算されることがあるので、HTMLとExcelで表示が変わります）: Sheet1!A1 #NAME\?（SUMSQ はHTMLで計算できない関数）、Sheet1!A2 #REF!（B40 はシートの外。シートは30行・12列）、Sheet1!A3 #CIRC!（循環参照）。/m);
  assert.match(result.stderr, /関数は表にあるもので書き換えるか値を数値で入れ、シートの外の参照は rows・cols を増やすか参照を直し、循環参照はなくしてください。/);
  assert.match(result.stderr, /^注意: エラーのセルを参照していて、同じエラーになっている数式があります: Sheet1!A6 #NAME\?。参照先のセルを直すと直ります。/m);
  assert.match(result.stderr, /^注意: 式の結果がエラーになる数式があります（Excelでも同じエラーになります）: Sheet1!A7 #REF!。/m);
  assert.ok(!/Sheet1!A4/.test(result.stderr), 'IFERROR で受けたエラーは数えない');
  assert.match(result.stderr, /^注意: 表示形式が「@」（文字列）の数式のセルがあります: Sheet1!A5/m);
});

test('自分で書いたブック（--input）の、HTMLの数式として読めない式は、場所と式を示して止まる。画面で編集した後の表では注意にとどめ、xlsx では1回だけ知らせる', () => {
  const folder = newFolder('unreadable'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 'f', f: '=A2*Tax'}, A2: {t: 'f', f: '=SUM(B1'}}}})), 'utf8');
  const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /^エラー: ブック形式エラー: HTMLの数式として読めない式があります: Sheet1!A1 =A2\*Tax、Sheet1!A2 =SUM\(B1。/);
  assert.match(result.stderr, /画面では #NAME\? か #VALUE! になり、xlsx には文字列として入ります。同じシートの番地と、HTMLで計算できる関数で書き直してください。/);
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.H1 = {t: 'f', f: '=Sheet2!A1'}; }), 'utf8');
  const state = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--output', join(folder, '反映.html')]);
  assert.equal(state.status, 0, state.stderr);
  assert.match(state.stderr, /^注意: HTMLの数式として読めない式があります: 明細!H1 =Sheet2!A1。/m);
  assert.ok(!/書き直してください/.test(state.stderr), 'ユーザーの式なので、直し方は示さない');
  const xlsx = runBuild(['--base', html, '--state', join(folder, 'state.txt'), '--xlsx', join(folder, '反映.xlsx')]);
  assert.equal(xlsx.status, 0, xlsx.stderr);
  assert.equal(xlsx.stderr.match(/H1/g).length, 1, 'xlsx.cjs が文字列にしたと知らせるので、同じセルを2回知らせない:\n' + xlsx.stderr);
});

test('Excelのセルに入らない長さの文字や数式は、--input では場所を示して止まり、編集状態の反映では注意にとどめる', () => {
  const folder = newFolder('excel-limits'), json = join(folder, 'b.json');
  const cases = [
    [{A1: {t: 's', v: 'あ'.repeat(32768)}}, /Sheet1!A1 の文字（32,768文字。Excelのセルは32,767文字まで）/],
    [{A1: {t: 'f', f: '="' + 'x'.repeat(256) + '"'}}, /Sheet1!A1 の数式の中の文字（256文字。Excelの数式の中の文字は255文字まで）/],
  ];
  for (const [cells, message] of cases) {
    writeFileSync(json, JSON.stringify(blankBook({sheet: {cells}})), 'utf8');
    const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
    assert.equal(result.status, 1, String(message));
    assert.match(result.stderr, /^エラー: ブック形式エラー: Excelのセルに入らない長さがあります/);
    assert.match(result.stderr, message);
  }
  const fits = [{t: 's', v: 'あ'.repeat(32767)}, {t: 'f', f: '=' + '1+'.repeat(4095) + '1'}, {t: 'f', f: '="' + 'x'.repeat(255) + '"'}, {t: 'f', f: '="' + '""'.repeat(200) + '"'}];
  assert.equal(fits[1].f.length, 8192);
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: Object.fromEntries(fits.map((cell, i) => ['A' + (i + 1), cell]))}})), 'utf8');
  const exact = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(exact.status, 0, exact.stderr);
  assert.ok(!/Excelのセルに入らない長さ/.test(exact.stderr), 'ちょうど上限の長さはExcelに入る（"" は1文字）:\n' + exact.stderr);
  const long = '=' + Array.from({length: 4200}, (_, i) => i % 10).join('+');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {A1: {t: 'f', f: long}}}})), 'utf8');
  const formula = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(formula.status, 1);
  assert.match(formula.stderr, new RegExp('Sheet1!A1 の数式（' + long.length.toLocaleString('en-US') + '文字。Excelの数式は8,192文字まで）'));
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => { w.sheets[0].cells.A20 = {t: 's', v: 'い'.repeat(33000)}; }), 'utf8');
  const state = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(state.status, 0, state.stderr);
  assert.match(state.stderr, /^注意: Excelのセルに入らない長さがあります: 明細!A20 の文字（33,000文字/m);
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
  assert.match(missing.stderr, /--input、または --base と --state の組を指定してください（--extract-json なら --base だけでもよい）/);
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

/** The kept older versions in folder (<stem>_<YYYYMMDDHHMMSS>.html) whose revision is revision. */
const olderWith = (folder, revision) => listHtml(folder).filter(n => /_\d{14}\.html$/.test(n)).map(n => join(folder, n)).filter(p => embeddedBook(p).revision === revision);

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

test('画面で編集した後の表の罫線の食い違いは、ユーザーが付けた線として知らせ、直し方は示さない', () => {
  const folder = newFolder('conflict-note');
  const html = buildBook(fixture('book'), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(html, w => {
    w.sheets[0].cells.A20 = {t: 's', v: '上', s: {borderBottom: '1px solid #808080'}};
    w.sheets[0].cells.A21 = {t: 's', v: '下', s: {borderTop: '3px double #000000'}};
  }), 'utf8');
  const result = runBuild(['--base', html, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: 罫線の食い違いが1か所あります.*ユーザーが付けた線（画面での編集や、元のxlsxの線）なので、そのまま残しました。/m);
  assert.ok(!/片方のセルにだけ書くか|省略形の styles/.test(result.stderr), result.stderr);
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

test('Excelのファイルから読み込んだ表の罫線の食い違いは、止めずに注意で知らせる', () => {
  const folder = newFolder('conflict-xlsx'), json = join(folder, 'b.json'), xlsx = join(folder, '元.xlsx');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {cells: {
    A1: {t: 's', v: '上', s: {borderBottom: '1px solid #808080'}}, A2: {t: 's', v: '下', s: {borderTop: '2px solid #217346'}}}}})), 'utf8');
  const made = runBuild(['--input', json, '--xlsx', xlsx, '--keep-user-content']);
  assert.equal(made.status, 0, made.stderr);
  const result = runBuild(['--from-xlsx', xlsx, '--output', join(folder, '読み込み.html')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^注意: 罫線の食い違いが1か所あります.*A1 の下と A2 の上/m);
  assert.ok(existsSync(join(folder, '読み込み.html')));
});

test('結合セル同士が複数のセルで接する辺の食い違いは、1か所として1回だけ示す', () => {
  const folder = newFolder('conflict-merged'), json = join(folder, 'b.json');
  writeFileSync(json, JSON.stringify(blankBook({sheet: {merges: [{r1: 1, c1: 1, r2: 2, c2: 1}, {r1: 1, c1: 2, r2: 2, c2: 2}], cells: {
    B2: {t: 's', v: '左', s: {borderRight: '1px solid #808080'}}, C2: {t: 's', v: '右', s: {borderLeft: '2px solid #217346'}}}}})), 'utf8');
  const result = runBuild(['--input', json, '--output', join(folder, 'b.html')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /罫線の食い違いが1か所あります/);
  assert.equal(result.stderr.match(/B2:B3 の右と C2:C3 の左/g).length, 1, result.stderr);
});

test('--extract-json のヘルプは、どの引数と組み合わせて何を書き出すかを示す', () => {
  const help = runBuild(['--help']);
  assert.equal(help.status, 0);
  assert.match(help.stdout.replace(/\s+/g, ''), /--inputなら省略形を広げた後、--baseなら--stateを反映した後（--stateがなければそのHTMLの表）、--from-xlsxなら読み込んだもの/);
  assert.match(help.stdout, /--keep-user-content/);
});

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

test('名前に年月日時分秒が付いていても、元の名前のファイルが別の表なら、以前の版とはみなさない', () => {
  const folder = newFolder('not-older');
  const dated = buildBook(fixture('book'), folder, '表_20260101000000.html');
  buildBook(blankBook({id: 'other-book', title: '別の表'}), folder, '表.html');
  writeFileSync(join(folder, 'state.txt'), stateText(dated, w => { w.sheets[0].cells.B4.v = '反映'; }), 'utf8');
  const result = runBuild(['--base', dated, '--state', join(folder, 'state.txt')]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(embeddedBook(dated).sheets[0].cells.B4.v, '反映');
});

test('--extract-json の --output にこのスキルのHTMLを指定すると、HTMLをJSONで置き換えずに止まる', () => {
  const folder = newFolder('json-over-html');
  const html = buildBook(fixture('book'), folder, '表.html'), before = readFileSync(html, 'utf8');
  const result = runBuild(['--base', html, '--output', html, '--extract-json']);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /--output の 表\.html は、このスキルのHTMLです。--extract-json はJSONを書き出すので、HTMLがJSONに置き換わります。--output にはJSONのファイル（例 .*表-edit\.json）を指定してください/);
  assert.equal(readFileSync(html, 'utf8'), before);
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
