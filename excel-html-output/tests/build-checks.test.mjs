// scripts/build.py の検査と注意：ブックの形式の検査、省略形（values・formulas・styles）、罫線の食い違い、
// HTMLで計算できない数式、Excelの上限。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {existsSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {test} from 'node:test';
import {listHtml, newFolder, stateText, useTempDir, xlsxOf} from './lib/build.mjs';
import {blankBook, buildBook, embeddedBook, findPython, fixture, runBuild, SKILL_ROOT} from './lib/workbook.mjs';

useTempDir();

test('検査を通らないブックは、エラーを出して何も書き出さない', () => {
  const folder = newFolder('invalid');
  const book = blankBook({sheet: {cells: {A1: {t: 's', v: 'x', s: {borderTop: '1px groove #000000'}}}}});
  writeFileSync(join(folder, 'bad.json'), JSON.stringify(book), 'utf8');
  const result = runBuild(['--input', join(folder, 'bad.json'), '--output', join(folder, 'bad.html')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /ブック形式エラー: sheets\[0\]\.cells\.A1\.s\.borderTop "1px groove #000000"（none か、線と色/);
  assert.deepEqual(listHtml(folder), []);
});

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
