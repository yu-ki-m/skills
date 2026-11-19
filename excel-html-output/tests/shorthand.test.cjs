// ブックJSONの省略形（values・formulas・styles）を、セルごとの形に広げる core.js の expand。
// build.py --input が検査の前に使うので、ここで広げた形が、埋め込まれるブック・状態テキスト・xlsx の中身になる。
const assert = require('node:assert/strict');
const {test} = require('node:test');
const C = require('../assets/core.js');

const B = '1px solid #808080', THICK = '2px solid #217346', DOUBLE = '3px double #000000';
const BOX = {borderTop: B, borderRight: B, borderBottom: B, borderLeft: B};
const book = (sheet = {}) => ({schema: C.SCHEMA, id: 'shorthand', revision: 'draft', title: '省略形', activeSheet: 's',
  sheets: [{id: 's', name: 'Sheet1', rows: 20, cols: 8, ...sheet}]});
const expanded = (sheet, notes = []) => C.validate(C.expand(book(sheet), note => notes.push(note))).sheets[0];
const cellsOf = sheet => expanded(sheet).cells;

test('values writes its rows from the cell at: text, a formula for text starting with =, numbers, TRUE / FALSE and whole cells; null and "" write nothing', () => {
  const cells = cellsOf({values: [{at: 'B2', rows: [
    ['品番', '=C3*2', 3, true],
    ['00123', null, -1.5, false, ''],
    [{t: 's', v: '=で始まる文字'}, {t: 's', v: 'サイト', link: 'https://example.com/'}, {t: 'n', v: 46023, s: {numFmt: 'yyyy/mm/dd'}}],
  ]}], styles: [{range: 'B3', s: {numFmt: '@'}}]});
  assert.deepEqual(cells, {
    B2: {t: 's', v: '品番'}, C2: {t: 'f', f: '=C3*2'}, D2: {t: 'n', v: 3}, E2: {t: 'b', v: true},
    B3: {t: 's', v: '00123', s: {numFmt: '@'}}, D3: {t: 'n', v: -1.5}, E3: {t: 'b', v: false},
    B4: {t: 's', v: '=で始まる文字'}, C4: {t: 's', v: 'サイト', link: 'https://example.com/'}, D4: {t: 'n', v: 46023, s: {numFmt: 'yyyy/mm/dd'}},
  });
});

test('formulas fills its range with the formula of its top-left cell, moving the relative references as Excel copies a formula', () => {
  const cells = cellsOf({formulas: [
    {range: 'E2:E4', f: '=C2*D2+$A$1+C$1+$B2'},
    {range: 'F2:G2', f: '=SUM(A2:B2)&"A2"'},
    {range: 'H5:H6', f: '=LOG10(A5*1E2)'},
  ]});
  assert.deepEqual(cells, {
    E2: {t: 'f', f: '=C2*D2+$A$1+C$1+$B2'}, E3: {t: 'f', f: '=C3*D3+$A$1+C$1+$B3'}, E4: {t: 'f', f: '=C4*D4+$A$1+C$1+$B4'},
    F2: {t: 'f', f: '=SUM(A2:B2)&"A2"'}, G2: {t: 'f', f: '=SUM(B2:C2)&"A2"'},
    H5: {t: 'f', f: '=LOG10(A5*1E2)'}, H6: {t: 'f', f: '=LOG10(A6*1E2)'},
  });
});

test('styles lie over the style the cells give themselves, in the order they are written, and make styled empty cells', () => {
  const cells = cellsOf({
    cells: {A1: {t: 's', v: '題名', s: {bold: true, numFmt: '@'}}},
    values: [{at: 'A2', rows: [['見出し', {t: 'n', v: 5, s: {fill: '#ffff00', italic: true}}]]}],
    styles: [{range: 'A1:B2', s: {fill: '#eeeeee'}}, {range: 'A1:B1', s: {numFmt: '0', fill: '#e2efda'}}, {range: 'B2', s: {fill: '#dddddd'}}],
  });
  assert.deepEqual(cells.A1, {t: 's', v: '題名', s: {bold: true, numFmt: '0', fill: '#e2efda'}}, '後に書いた styles ほど強く、セル自身の書式にも重なる');
  assert.deepEqual(cells.B1, {t: 's', v: '', s: {numFmt: '0', fill: '#e2efda'}}, '値のないセルにも書式だけのセルを作る');
  assert.deepEqual(cells.A2, {t: 's', v: '見出し', s: {fill: '#eeeeee'}});
  assert.deepEqual(cells.B2, {t: 'n', v: 5, s: {fill: '#dddddd', italic: true}});
});

test('a line laid on one side of a cell edge is laid on the other side too, as Excel keeps one line per edge: a grid with a thick line under the heading and a double line over the total has no conflict', () => {
  const sheet = expanded({styles: [
    {range: 'A1:C5', s: BOX},
    {range: 'A1:C1', s: {borderBottom: THICK}},
    {range: 'A5:C5', s: {borderTop: DOUBLE}},
    {range: 'B3', s: {borderRight: 'none'}},
  ]});
  const s = sheet.cells;
  assert.deepEqual([s.A1.s.borderBottom, s.A2.s.borderTop], [THICK, THICK], '見出しの下の線は、2行目の上の線でもある');
  assert.deepEqual([s.B4.s.borderBottom, s.B5.s.borderTop], [DOUBLE, DOUBLE], '合計の上の線は、4行目の下の線でもある');
  assert.deepEqual([s.B3.s.borderRight, s.C3.s.borderLeft], ['none', 'none'], '線を消すと、隣のセルの側の線も消える');
  assert.deepEqual([s.A2.s.borderBottom, s.A3.s.borderTop], [B, B], 'ほかの辺は格子のまま');
  assert.equal(s.D1, undefined, '範囲の外の、線のない隣のセルには書式を作らない');
  assert.deepEqual(C.borderConflicts(book(), sheet), []);
});

test('a line laid on the edge of a merged cell meets the merged cell as one block', () => {
  const cells = cellsOf({merges: [C.range('B2:C3')], values: [{at: 'B2', rows: [['結合']]}],
    styles: [{range: 'A1:D4', s: BOX}, {range: 'B1:C1', s: {borderBottom: THICK}}, {range: 'D2:D3', s: {borderLeft: DOUBLE}}]});
  assert.deepEqual([cells.B2.s.borderTop, cells.B2.s.borderRight], [THICK, DOUBLE], '結合セルの上と右の辺は、左上のセルが持つ');
  for (const a of ['C2', 'B3', 'C3']) assert.equal(cells[a], undefined, a + ' は結合で隠れるので書式を付けない');
});

test('a range that holds only part of a merged cell is refused, since it would put the style on the whole cell or on none of it', () => {
  for (const range of ['A4:H4', 'A3:H3', 'A4']) {
    assert.throws(() => C.expand(book({merges: [C.range('A3:A4')], styles: [{range, s: {fill: '#e2efda'}}]})),
      new RegExp('sheets\\[0\\]\\.styles\\[0\\]\\.range ' + range + ' は結合セル A3:A4 の一部だけを含みます'), range);
  }
  assert.deepEqual(cellsOf({merges: [C.range('A3:A4')], styles: [{range: 'A3:H4', s: {fill: '#e2efda'}}]}).A3, {t: 's', v: '', s: {fill: '#e2efda'}});
});

test('a cell of cells that holds only a style takes a value from values or formulas, and keeps its style', () => {
  const cells = cellsOf({cells: {A1: {t: 's', v: '', s: {bold: true}}, A2: {t: 's', v: '', s: {italic: true}}},
    values: [{at: 'A1', rows: [['合計']]}], formulas: [{range: 'A2', f: '=1+1'}]});
  assert.deepEqual(cells, {A1: {t: 's', v: '合計', s: {bold: true}}, A2: {t: 'f', f: '=1+1', s: {italic: true}}});
});

test('a value written twice, a range or cell outside the sheet, a whole column, a wrong element and an unknown key stop with where they are', () => {
  const cases = [
    [{cells: {A1: {t: 's', v: 'x'}}, values: [{at: 'A1', rows: [['y']]}]}, /sheets\[0\] の A1 に値を2回書いています（cells と values\[0\]）/],
    [{values: [{at: 'A1', rows: [['x']]}, {at: 'A1', rows: [['y']]}]}, /sheets\[0\] の A1 に値を2回書いています（values\[0\] と values\[1\]）/],
    [{values: [{at: 'A1', rows: [['x']]}], formulas: [{range: 'A1:A2', f: '=1'}]}, /sheets\[0\] の A1 に値を2回書いています（values\[0\] と formulas\[0\]）/],
    [{formulas: [{range: 'A1:A21', f: '=1'}]}, /sheets\[0\]\.formulas\[0\]\.range A1:A21 がシートの外です（20行・8列）/],
    [{values: [{at: 'H1', rows: [[1, 2]]}]}, /sheets\[0\]\.values\[0\]\.rows\[0\]\[1\] I1 がシートの外です/],
    [{values: [{at: 'Z99', rows: [[1]]}]}, /sheets\[0\]\.values\[0\]\.at Z99 がシートの外です/],
    [{styles: [{range: 'B:B', s: {bold: true}}]}, /sheets\[0\]\.styles\[0\]\.range（"A1:C3" のような範囲）/],
    [{values: [{at: 'A1', rows: [[[1]]]}]}, /sheets\[0\]\.values\[0\]\.rows\[0\]\[0\]（文字・数値・true \/ false・null・セルのオブジェクト）/],
    [{values: [{at: 'A1', rows: [[{v: '合計', s: {bold: true}}]]}]}, /sheets\[0\]\.values\[0\]\.rows\[0\]\[0\]（セルのオブジェクトには t（s \/ n \/ b \/ f）を書く）/],
    [{values: [{at: 'A1', rows: ['x']}]}, /sheets\[0\]\.values\[0\]\.rows\[0\]（値の配列）/],
    [{formulas: [{range: 'A1', f: 'SUM(A2)'}]}, /sheets\[0\]\.formulas\[0\]\.f（= で始まる式）/],
    [{styles: [{range: 'A1', s: 'bold'}]}, /sheets\[0\]\.styles\[0\]\.s（書式のオブジェクト）/],
    [{styles: [{range: 'A1', style: {bold: true}}]}, /未対応の項目 sheets\[0\]\.styles\[0\]\.style/],
    [{values: [{at: 'A1', rows: [['x']], s: {bold: true}}]}, /未対応の項目 sheets\[0\]\.values\[0\]\.s/],
    [{styles: {range: 'A1', s: {}}}, /sheets\[0\]\.styles（配列）/],
    [{rows: [['品名', 1]], values: []}, /sheets\[0\]\.rows は行の数（値の行は values の rows に書く）/],
    [{styles: [{range: 'A1', s: {bold: 'yes'}}]}, /ブック形式エラー: sheets\[0\]\.styles\[0\]\.s\.bold "yes"（true \/ false）/],
    [{merges: [C.range('A1:B1')], values: [{at: 'A1', rows: [['x', 'y']]}]}, /結合セル内の隠れた値 B1/],
  ];
  for (const [sheet, message] of cases) assert.throws(() => C.validate(C.expand(book(sheet))), message, JSON.stringify(sheet));
});

test('a text of values that reads as a number or a date becomes a text cell, which SUM leaves out, so it is told unless the cell is formatted as text', () => {
  const notes = [];
  expanded({values: [{at: 'A1', rows: [['00123', '30,000', '2', '2026/1/2', '品名', '-5', '12:30', '50%'], ['00124', 30000, 2]]}], styles: [{range: 'A1:A2', s: {numFmt: '@'}}]}, notes);
  assert.equal(notes.length, 2, notes.join('\n'));
  assert.match(notes[0], /^数値に見える文字列が、文字のセルになります（SUM などの計算に入りません）: Sheet1!B1 "30,000"、Sheet1!C1 "2"、Sheet1!F1 "-5"、Sheet1!H1 "50%"。/);
  assert.match(notes[0], /計算する値なら引用符を外して数値で書き、番号のように文字のままにするなら numFmt "@" を付けてください。$/);
  assert.match(notes[1], /^日付や時刻に見える文字列が、文字のセルになります: Sheet1!D1 "2026\/1\/2"、Sheet1!G1 "12:30"。/);
  assert.match(notes[1], /日付として計算するなら、シリアル値（2026\/1\/2 は 46024）に numFmt "yyyy\/mm\/dd" を付けて書いてください。文字のままにするなら numFmt "@" を付けてください。$/);
  const quiet = [];
  expanded({values: [{at: 'A1', rows: [['品名', 1]]}]}, quiet);
  assert.deepEqual(quiet, []);
});

test('ordinary text that Excel would take for a formula or a value of its own is not told as a number: it goes to Excel as text, and the number to mend is not hidden behind it', () => {
  const notes = [];
  expanded({values: [{at: 'A1', rows: [['-なし'], ['- 交通費は別途'], ['@担当'], ['+81 90 1234 5678'], ["'00123"], ['TRUE'], ['#N/A'], ['30,000']]}]}, notes);
  assert.equal(notes.length, 1, notes.join('\n'));
  assert.match(notes[0], /^数値に見える文字列が、文字のセルになります（SUM などの計算に入りません）: Sheet1!A8 "30,000"。/);
});

test('a book without the shorthand comes out the same; cells, merges, rowHeights and colWidths may be left out; the book given is not changed', () => {
  const plain = book({cells: {A1: {t: 'n', v: 1, s: {bold: true}}}, merges: [], rowHeights: [30], colWidths: [100]});
  assert.deepEqual(C.expand(plain), plain);
  assert.deepEqual(C.expand(book()).sheets[0], {id: 's', name: 'Sheet1', rows: 20, cols: 8, cells: {}, merges: [], rowHeights: [], colWidths: []});
  const given = book({cells: {A1: {t: 's', v: '', s: {bold: true}}, A2: {t: 's', v: 'x', s: {borderTop: B}}},
    values: [{at: 'A3', rows: [['y']]}], styles: [{range: 'A1:B1', s: BOX}]});
  const before = JSON.stringify(given);
  C.expand(given);
  assert.equal(JSON.stringify(given), before);
});

test('an estimate written with the shorthand is the same book as written cell by cell, at a fraction of the length', () => {
  const head = ['品番', '品名', '数量', '単価', '金額', '消費税', '税込'];
  const items = Array.from({length: 50}, (_, i) => [String(1000 + i).padStart(5, '0'), '作業項目' + i, (i % 7) + 1, 1000 * ((i % 13) + 1)]);
  const long = {A1: {t: 's', v: '見積明細', s: {bold: true, size: 16}}};
  head.forEach((h, c) => { long[C.addr(2, c)] = {t: 's', v: h, s: {...BOX, borderBottom: THICK, bold: true, fill: '#e2efda'}}; });
  items.forEach((row, i) => {
    const r = i + 4, top = r === 4 ? {borderTop: THICK} : {};
    long['A' + r] = {t: 's', v: row[0], s: {...BOX, ...top, numFmt: '@'}};
    long['B' + r] = {t: 's', v: row[1], s: {...BOX, ...top}};
    long['C' + r] = {t: 'n', v: row[2], s: {...BOX, ...top}};
    long['D' + r] = {t: 'n', v: row[3], s: {...BOX, ...top, numFmt: '¥#,##0'}};
    long['E' + r] = {t: 'f', f: `=C${r}*D${r}`, s: {...BOX, ...top, numFmt: '¥#,##0'}};
    long['F' + r] = {t: 'f', f: `=ROUNDDOWN(E${r}*0.1,0)`, s: {...BOX, ...top, numFmt: '¥#,##0'}};
    long['G' + r] = {t: 'f', f: `=E${r}+F${r}`, s: {...BOX, ...top, numFmt: '¥#,##0'}};
  });
  long.D54 = {t: 's', v: '合計', s: {...BOX, bold: true, numFmt: '¥#,##0'}};
  for (const c of 'EFG') long[c + '54'] = {t: 'f', f: `=SUM(${c}4:${c}53)`, s: {...BOX, bold: true, numFmt: '¥#,##0'}};
  for (const c of 'ABC') long[c + '54'] = {t: 's', v: '', s: BOX};
  const sheet = {rows: 60, cols: 8};
  const short = {...sheet,
    cells: {A1: {t: 's', v: '見積明細', s: {bold: true, size: 16}}},
    values: [{at: 'A3', rows: [head, ...items]}, {at: 'D54', rows: [['合計']]}],
    formulas: [{range: 'E4:E53', f: '=C4*D4'}, {range: 'F4:F53', f: '=ROUNDDOWN(E4*0.1,0)'}, {range: 'G4:G53', f: '=E4+F4'}, {range: 'E54:G54', f: '=SUM(E4:E53)'}],
    styles: [{range: 'A3:G54', s: BOX}, {range: 'A3:G3', s: {bold: true, fill: '#e2efda', borderBottom: THICK}}, {range: 'A4:A53', s: {numFmt: '@'}},
      {range: 'D4:G54', s: {numFmt: '¥#,##0'}}, {range: 'D54:G54', s: {bold: true}}]};
  const notes = [];
  const result = C.validate(C.expand(book(short), note => notes.push(note)));
  assert.deepEqual(result, C.validate(book({...sheet, cells: long, merges: [], rowHeights: [], colWidths: []})));
  assert.deepEqual(notes, [], '品番は numFmt "@" の文字なので、注意を出さない');
  assert.deepEqual(C.borderConflicts(result, result.sheets[0]), []);
  const length = cells => Object.entries(cells).map(([a, cell]) => JSON.stringify(a) + ':' + JSON.stringify(cell)).join(',\n').length;
  const written = JSON.stringify(short).length;
  assert.ok(written * 10 < length(long), `省略形 ${written} 字、セルごと ${length(long)} 字`);
});

test('the first example of references/workbook-schema.md is a book that builds without a caution or a border conflict', () => {
  const md = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'references', 'workbook-schema.md'), 'utf8');
  const notes = [], w = C.validate(C.expand(JSON.parse(/```json\n([\s\S]*?)\n```/.exec(md)[1]), note => notes.push(note)));
  assert.deepEqual(notes, []);
  for (const s of w.sheets) assert.deepEqual(C.borderConflicts(w, s), [], s.name);
  const cells = w.sheets[0].cells;
  assert.deepEqual([cells.E5.f, cells.E6.f, cells.A4.s.numFmt, cells.A4.s.borderTop, cells.E5.s.numFmt], ['=C5*D5', '=SUM(E4:E5)', '@', '2px solid #217346', '¥#,##0'],
    '資料の本文が説明しているとおりに広がる');
  // 編集状態の節の TOON の例は、この例から作ったHTMLの状態で、セルの一部を示している。
  const toon = C.readStateText(/```text\n(EXCEL_HTML_STATE_V1 TOON[\s\S]*?EXCEL_HTML_STATE_END)\n```/.exec(md)[1]).workbook;
  const {cells: shown, ...sheet} = toon.sheets[0], {cells: all, ...built} = w.sheets[0];
  assert.deepEqual(sheet, built, 'TOON の例のシートの項目は、冒頭の例を広げたものと同じ');
  for (const a of Object.keys(shown)) assert.deepEqual(shown[a], all[a], 'TOON の例の ' + a);
  assert.deepEqual([toon.id, toon.title, toon.defaultStyle, toon.view], [w.id, w.title, w.defaultStyle, w.view]);
});

test('a line laid on only a part of an edge of a merged cell next to the range stops with where it is, when the merged cell has another line there: its edge is one line, so the whole edge would change', () => {
  const frame = {borderTop: THICK, borderRight: THICK, borderBottom: THICK, borderLeft: THICK};
  const sheet = extra => ({merges: [C.range('A5:D6')], styles: [{range: 'A5:D6', s: frame}, ...extra]});
  for (const line of [DOUBLE, 'none']) {
    assert.throws(() => C.expand(book(sheet([{range: 'C4', s: {borderBottom: line}}]))),
      /sheets\[0\]\.styles\[1\]\.range C4 の下の線は、結合セル A5:D6 の上の辺の一部にしか接しません。結合セルの辺は1本の線なので、A4:D4 のように辺全体に引いてください/, line);
  }
  assert.throws(() => C.expand(book({merges: [C.range('A5:D6')], styles: [{range: 'A1:E8', s: BOX}, {range: 'B4:C4', s: {borderBottom: THICK}}]})),
    /sheets\[0\]\.styles\[1\]\.range B4:C4 の下の線は、結合セル A5:D6 の上の辺の一部にしか接しません/, '格子の上に太線を重ねたときも');
  assert.throws(() => C.expand(book({merges: [C.range('B2:B4')], styles: [{range: 'B2:B4', s: frame}, {range: 'C3', s: {borderLeft: DOUBLE}}]})),
    /styles\[1\]\.range C3 の左の線は、結合セル B2:B4 の右の辺の一部にしか接しません。結合セルの辺は1本の線なので、C2:C4 のように/);
  assert.throws(() => C.expand(book({merges: [C.range('A1:D2')], styles: [{range: 'A1:D2', s: frame}, {range: 'C3', s: {borderTop: 'none'}}]})),
    /styles\[1\]\.range C3 の上の線は、結合セル A1:D2 の下の辺の一部にしか接しません。結合セルの辺は1本の線なので、A3:D3 のように/);
  const whole = expanded(sheet([{range: 'A4:D4', s: {borderBottom: DOUBLE}}]));
  assert.deepEqual([whole.cells.A5.s.borderTop, whole.cells.C4.s.borderBottom], [DOUBLE, DOUBLE], '辺全体に接する範囲なら、結合セルの辺もそろえる');
  assert.deepEqual(C.borderConflicts(book(), whole), []);
  const same = expanded(sheet([{range: 'C4', s: {borderBottom: THICK}}]));
  assert.equal(same.cells.A5.s.borderTop, THICK, '結合セルと同じ線なら、変えるものがないので止めない');
  const plain = expanded({merges: [C.range('A5:D6')], styles: [{range: 'C4', s: {borderBottom: DOUBLE}}]});
  assert.equal(plain.cells.C4.s.borderBottom, DOUBLE);
  assert.equal(plain.cells.A5, undefined, '線のない結合セルには何も足さない');
});

test('the lines of the default styles of the book and the sheet are lines of the neighbouring cells too: a grid in defaultStyle with a thick line laid under the heading has no conflict', () => {
  for (const where of ['sheet', 'book']) {
    const w = book({styles: [{range: 'A1:C1', s: {borderBottom: THICK}}]});
    if (where === 'sheet') w.sheets[0].defaultStyle = BOX; else w.defaultStyle = BOX;
    const x = C.validate(C.expand(w)), s = x.sheets[0];
    assert.deepEqual([s.cells.A1.s.borderBottom, s.cells.A2.s.borderTop, s.cells.C2.s.borderTop], [THICK, THICK, THICK], where);
    assert.equal(s.cells.A3, undefined, where + ': 線を変えないセルには書式を作らない');
    assert.deepEqual(C.borderConflicts(x, s), [], where);
  }
});

test('a cell of cells that holds only a style keeps its link under a value of values, and the style the value cell gives itself lies over it', () => {
  const cells = cellsOf({cells: {A1: {t: 's', v: '', s: {bold: true, fill: '#eeeeee'}, link: 'https://example.com/'}},
    values: [{at: 'A1', rows: [[{t: 's', v: 'サイト', s: {fill: '#ffff00'}}]]}]});
  assert.deepEqual(cells.A1, {t: 's', v: 'サイト', s: {bold: true, fill: '#ffff00'}, link: 'https://example.com/'});
});

test('names that every object has (constructor, toString) are unknown style keys, not a failure inside', () => {
  for (const k of ['constructor', 'toString', '__proto__x']) {
    assert.throws(() => C.expand(book({styles: [{range: 'A1', s: {[k]: 1}}]})), new RegExp('未対応の書式 sheets\\[0\\]\\.styles\\[0\\]\\.s\\.' + k), k);
    assert.throws(() => C.validate(C.expand(book({cells: {A1: {t: 's', v: 'x', s: {[k]: 1}}}}))), new RegExp('未対応の書式 sheets\\[0\\]\\.cells\\.A1\\.s\\.' + k), k);
  }
});

test('merges may be written as "A1:E1" ranges in a book written by hand, and a wrong merge says how one is written', () => {
  const s = expanded({merges: ['A1:E1', 'b3:c4'], values: [{at: 'A1', rows: [['題名']]}]});
  assert.deepEqual(s.merges, [{r1: 0, c1: 0, r2: 0, c2: 4}, {r1: 2, c1: 1, r2: 3, c2: 2}]);
  assert.throws(() => C.expand(book({merges: ['A1-E1']})), /sheets\[0\]\.merges\[0\] "A1-E1"（"A1:E1" のような範囲）/);
  assert.throws(() => C.expand(book({merges: ['A1:Z1']})), /sheets\[0\]\.merges\[0\] A1:Z1 がシートの外です（20行・8列）/);
  assert.throws(() => C.validate(C.expand(book({merges: [{r1: 0, c1: 0, r2: 0}]}))),
    /sheets\[0\]\.merges\[0\]（\{"r1":0,"c1":0,"r2":0,"c2":4\} のような、0始まりで両端を含む行と列の番号。シートの中に収める）/);
});

test('a wrong style value stops with where it is and what is taken: the styles entry it was written in, or the cell', () => {
  const cases = [
    [{styles: [{range: 'A1', s: {valign: 'center'}}]}, /sheets\[0\]\.styles\[0\]\.s\.valign "center"（top \/ middle \/ bottom）/],
    [{styles: [{range: 'A1:B2', s: {borderBottom: '3px solid'}}]}, /sheets\[0\]\.styles\[0\]\.s\.borderBottom "3px solid"（none か、線と色（例 1px solid #808080）。線は 0\.5px solid \/ 1px solid \/ 2px solid \/ 3px solid \/ 1px dashed \/ 2px dashed \/ 1px dotted \/ 3px double、色は #rrggbb）/],
    [{styles: [{range: 'A1', s: {size: '16'}}]}, /sheets\[0\]\.styles\[0\]\.s\.size "16"（6〜96 の数）/],
    [{styles: [{range: 'A1', s: {numFmt: 'yyyy/m/d'}}]}, /sheets\[0\]\.styles\[0\]\.s\.numFmt "yyyy\/m\/d"（General \/ @ \/ 0 \/ 0\.00 \/ #,##0 \/ #,##0\.00 \/ 0% \/ 0\.00% \/ ¥#,##0 \/ yyyy\/mm\/dd）/],
    [{cells: {A3: {t: 's', v: 'x', s: {fill: '#ddd'}}}}, /sheets\[0\]\.cells\.A3\.s\.fill "#ddd"（#rrggbb の6桁の色か none）/],
    [{cells: {A3: {t: 's', v: 'x', s: {color: 'black'}}}}, /sheets\[0\]\.cells\.A3\.s\.color "black"（#rrggbb の6桁の色）/],
    [{defaultStyle: {indent: 1.5}}, /sheets\[0\]\.defaultStyle\.indent 1\.5（0〜15 の整数）/],
  ];
  for (const [sheet, message] of cases) assert.throws(() => C.validate(C.expand(book(sheet))), message, JSON.stringify(sheet));
  const w = C.expand(book());
  w.defaultStyle = {font: 'A<B'};
  assert.throws(() => C.validate(w), /defaultStyle\.font "A<B"（フォント名。100文字まで。< > ; \{ \} " \\ を含まない）/);
});

test('a freeze needs both rows and cols, and a freeze across a merged cell names the merged cell', () => {
  assert.throws(() => C.validate(C.expand(book({freeze: {rows: 3}}))), /sheets\[0\]\.freeze（rows と cols の両方を、0 から行の数・列の数までの整数で書く。固定しない方は 0）/);
  assert.throws(() => C.validate(C.expand(book({freeze: {rows: 0, cols: 1}, merges: ['A1:G1']}))), /結合をまたぐ固定は未対応（シート「Sheet1」の結合 A1:G1 が、固定した列（1列）の境目をまたぐ）/);
  assert.throws(() => C.validate(C.expand(book({freeze: {rows: 1, cols: 0}, merges: ['A1:A2']}))), /結合をまたぐ固定は未対応（シート「Sheet1」の結合 A1:A2 が、固定した行（1行）の境目をまたぐ）/);
});

test('a book without its schema, id or revision says which one is missing', () => {
  for (const [patch, message] of [[{schema: 'v1'}, /schema（"excel-html-workbook\/v1"）/], [{id: ''}, /id（空でない文字）/], [{revision: undefined}, /revision（空でない文字。"draft" でよい）/]]) {
    const w = {...book(), ...patch};
    if ('revision' in patch) delete w.revision;
    assert.throws(() => C.validate(w), message, JSON.stringify(patch));
  }
});

test('the whole edge of a merged cell may be laid in two or more ranges, and the line of the merged cell itself may be laid later: only an edge left partly changed stops', () => {
  const frame = {borderTop: THICK, borderRight: THICK, borderBottom: THICK, borderLeft: THICK};
  const two = expanded({merges: [C.range('A5:D6')], styles: [{range: 'A5:D6', s: frame}, {range: 'A4:B4', s: {borderBottom: DOUBLE}}, {range: 'C4:D4', s: {borderBottom: DOUBLE}}]});
  assert.equal(two.cells.A5.s.borderTop, DOUBLE, '2つの範囲で辺全体を覆えば、結合セルの辺もそろう');
  assert.deepEqual(C.borderConflicts(book(), two), []);
  const later = expanded({merges: [C.range('A5:D6')], styles: [{range: 'A5:D6', s: frame}, {range: 'C4', s: {borderBottom: DOUBLE}}, {range: 'A5:D6', s: {borderTop: DOUBLE}}]});
  assert.equal(later.cells.A5.s.borderTop, DOUBLE, '結合セル自身の線として後から書けば、辺の線が決まる');
  assert.deepEqual(C.borderConflicts(book(), later), []);
  assert.throws(() => C.expand(book({merges: [C.range('A5:D6')], styles: [{range: 'A5:D6', s: frame}, {range: 'A4:B4', s: {borderBottom: DOUBLE}}]})),
    /sheets\[0\]\.styles\[1\]\.range A4:B4 の下の線は、結合セル A5:D6 の上の辺の一部にしか接しません/, '辺の一部が残れば止まる');
});
