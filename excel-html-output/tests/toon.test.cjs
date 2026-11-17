// assets/toon.js：TOON（Token-Oriented Object Notation、仕様4.4）の書き出しと読み取り。
// 「AIに渡す」の状態テキストは既定でTOONで書き出し、build.py がそれを読み戻してHTMLを作り直す。
// ここでは、書き出す形が仕様どおりであること（AIや他のTOONの道具が同じ意味に読む）、読み戻すと元と同じデータになること、
// 途中で切れたり崩れたりしたテキストを黙って読まずに止めることを確かめる。期待する形は、仕様書の付録Aの例に合わせている。
const assert = require('node:assert/strict');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const {test} = require('node:test');
const T = require('../assets/toon.js');

const lines = (...xs) => xs.join('\n');
/** The value as the JSON data model sees it (what TOON carries): undefined dropped, NaN to null, -0 to 0. */
const asJson = value => JSON.parse(JSON.stringify(value));

test('オブジェクト・入れ子・値の配列・空の配列とオブジェクトを、仕様の例と同じ形で書き出す', () => {
  assert.equal(T.encode({id: 123, name: 'Ada', active: true}), lines('id: 123', 'name: Ada', 'active: true'));
  assert.equal(T.encode({user: {id: 123, name: 'Ada'}}), lines('user:', '  id: 123', '  name: Ada'));
  assert.equal(T.encode({tags: ['admin', 'ops', 'dev']}), 'tags[3]: admin,ops,dev');
  assert.equal(T.encode({pairs: [[1, 2], [3, 4]]}), lines('pairs[2]:', '  - [2]: 1,2', '  - [2]: 3,4'));
  assert.equal(T.encode({tags: [], empty: {}}), lines('tags: []', 'empty:'));
  assert.equal(T.encode({}), '', '空のオブジェクトは空の文書');
  assert.equal(T.encode([1, 'a']), '[2]: 1,a');
  assert.equal(T.encode([]), '[]');
  assert.equal(T.encode('hello'), 'hello');
  assert.equal(T.encode(''), '""');
});

test('形のそろったオブジェクトの配列は表の形にし、入れ子の列は見出しにまとめ、同じ形の値を持つオブジェクトはキー付きの表にする', () => {
  assert.equal(T.encode({items: [{sku: 'A1', qty: 2, price: 9.99}, {sku: 'B2', qty: 1, price: 14.5}]}),
    lines('items[2]{sku,qty,price}:', '  A1,2,9.99', '  B2,1,14.5'));
  assert.equal(T.encode({orders: [{id: 1, customer: {name: 'Ada', country: 'DK'}, total: 99}, {id: 2, customer: {name: 'Bob', country: 'UK'}, total: 149}]}),
    lines('orders[2]{id,customer{name,country},total}:', '  1,Ada,DK,99', '  2,Bob,UK,149'));
  assert.equal(T.encode({users: {alice: {age: 30, city: 'Berlin'}, bob: {age: 25, city: 'Oslo'}}}),
    lines('users[2:]{age,city}:', '  alice: 30,Berlin', '  bob: 25,Oslo'));
  assert.equal(T.encode({users: {alice: {age: 30}}}), lines('users:', '  alice:', '    age: 30'), '値が1つだけならキー付きの表にしない');
  assert.equal(T.encode({items: [{a: 1}, {}]}), lines('items[2]:', '  - a: 1', '  -'), '空のオブジェクトを含む配列は表にしない');
  assert.equal(T.encode({items: [{a: null}, {a: {b: 1}}]}), lines('items[2]:', '  - a: null', '  - a:', '      b: 1'), 'null とオブジェクトが混じる列は表にしない');
});

test('形のそろわない配列は1行に1要素の並びにし、要素のオブジェクトは最初の項目をハイフンの行に置く', () => {
  assert.equal(T.encode({items: [1, {a: 1}, 'text']}), lines('items[3]:', '  - 1', '  - a: 1', '  - text'));
  assert.equal(T.encode({items: [{id: 1, name: 'First'}, {id: 2, name: 'Second', extra: true}]}),
    lines('items[2]:', '  - id: 1', '    name: First', '  - id: 2', '    name: Second', '    extra: true'));
  assert.equal(T.encode({items: [{users: [{id: 1, name: 'Ada'}, {id: 2, name: 'Bob'}], status: 'active'}]}),
    lines('items[1]:', '  - users[2]{id,name}:', '      1,Ada', '      2,Bob', '    status: active'));
  assert.equal(T.encode({x: [[{a: 1}, {a: 2}], []]}), lines('x[2]:', '  - [2]:', '    - a: 1', '    - a: 2', '  - [0]:'),
    '配列の中の配列は表にせず、並びにする');
});

test('文字列は、数・真偽値・null と取り違えるもの、区切りや記号を含むもの、前後に空白があるものだけを引用符で囲む', () => {
  const cases = [
    ['', '""'], [' ', '" "'], [' x', '" x"'], ['x ', '"x "'], ['true', '"true"'], ['false', '"false"'], ['null', '"null"'],
    ['00123', '"00123"'], ['05', '"05"'], ['1e5', '"1e5"'], ['-3.14', '"-3.14"'], ['+81', '"+81"'],
    ['=SUM(A1:A2)', '"=SUM(A1:A2)"'], ['a,b', '"a,b"'], ['say "hi"', '"say \\"hi\\""'], ['C:\\path', '"C:\\\\path"'],
    ['[x]', '"[x]"'], ['{x}', '"{x}"'], ['-x', '"-x"'], ['#e2efda', '"#e2efda"'],
    ['行1\n行2', '"行1\\n行2"'], ['a\tb', '"a\\tb"'], ['a\rb', '"a\\rb"'], ['\u0001', '"\\u0001"'],
    ['TRUE', 'TRUE'], ['見積明細', '見積明細'], ['保守作業 （月額）', '保守作業 （月額）'], ['2px solid #217346', '2px solid #217346'],
    ['0%', '0%'], ['=C4*D4', '=C4*D4'], ['1/2', '1/2'], ['1.', '1.'], ['a|b', 'a|b'], ['x#y', 'x#y'], ['🎉', '🎉'],
  ];
  for (const [value, expected] of cases) assert.equal(T.encode({v: value}), 'v: ' + expected, JSON.stringify(value));
  assert.equal(T.encode({v: ['a,b', 'c', '']}), 'v[3]: "a,b",c,""', '配列の中でも同じ');
});

test('識別子の形でない名前は、キーでも表の項目名でも引用符で囲む', () => {
  assert.equal(T.encode({'my-key': [1, 2, 3]}), '"my-key"[3]: 1,2,3');
  assert.equal(T.encode({'img-1': 'data:image/png;base64,AAAA', A1: 1, 'user.name': 'x', '備考': 'y', '': 0, '1a': 2}),
    lines('"img-1": "data:image/png;base64,AAAA"', 'A1: 1', 'user.name: x', '"備考": y', '"": 0', '"1a": 2'));
  assert.equal(T.encode({rows: [{'a-b': 1}, {'a-b': 2}]}), lines('rows[2]{"a-b"}:', '  1', '  2'));
});

test('数は指数を使わない最短の十進で書き、JSONと同じく NaN と無限大は null、undefined の項目は書かない', () => {
  assert.equal(T.encode({a: 1e6, b: 1.5, c: -0, d: 0.000001, e: 1e-7, f: 1e21, g: 0.1 + 0.2, h: NaN, i: Infinity, j: 1234567890}),
    lines('a: 1000000', 'b: 1.5', 'c: 0', 'd: 0.000001', 'e: 1e-7', 'f: 1e+21', 'g: 0.30000000000000004', 'h: null', 'i: null', 'j: 1234567890'));
  assert.equal(T.encode({a: 1, b: undefined, c: [undefined], d: new Date(0)}), lines('a: 1', 'c[1]: null', 'd: "1970-01-01T00:00:00.000Z"'));
});

test('書き出したテキストを読み戻すと、元と同じデータになる', () => {
  const book = JSON.parse(readFileSync(join(__dirname, 'fixtures', 'book.json'), 'utf8'));
  const odd = JSON.parse('{"__proto__":1,"constructor":2,"":3,"a.b":4,"#k":5,"-k":6,"k:v":7,"[k]":8,"備考":9}');
  const tricky = {
    strings: ['', ' ', ' x ', 'true', 'false', 'null', 'TRUE', '00123', '1e5', '-1', '+1', '1.', '.5', 'a,b', 'a|b', 'a\tb', '行1\n行2', 'cr\rlf',
      'q"uote', 'back\\slash', '[x]', '{x}', 'a:b', '-x', '#x', 'x#', '\u0000\u001f', '🎉👨‍👩‍👧', '\ufeffbom', '- item', '[]', '"', 'a: b, c'],
    numbers: [0, 1, -1, 1.5, 0.1 + 0.2, 1e21, 1e-7, 9007199254740991, -3.5e-10],
    nested: {a: {b: {c: [1, [2, [3, {d: 'e'}]]]}}},
    table: [{x: 1, y: {p: 'a', q: null}}, {x: 2, y: {p: 'b,c', q: true}}],
    keyed: {k1: {v: 1, w: 'x'}, k2: {v: 2, w: 'y'}},
    mixed: [1, 'a', null, [], [1, 2], [{a: 1}], {}, {a: [1, 2], b: {c: 1}}, {t: [{u: 1}, {u: 2}], v: 3}, {k: {p: {q: 1}, r: {q: 2}}}],
    odd, emptyArr: [], emptyObj: {},
  };
  for (const value of [book, tricky, [tricky, book], 'text', 0, null, [], {}]) assert.deepEqual(T.decode(T.encode(value)), asJson(value));
});

test('仕様が認める書き方の幅（コメント行、CRLF、行末の空白、区切りの前後の空白、タブやパイプの区切り）も読める', () => {
  assert.deepEqual(T.decode('# メモ\r\nname: Ada  \r\ntags[2]: a , b\r\n  # 字下げしたコメント\r\n"x-y": "1"'), {name: 'Ada', tags: ['a', 'b'], 'x-y': '1'});
  assert.deepEqual(T.decode('items[2|]{a|b}:\n  1|x,y\n  2|z'), {items: [{a: 1, b: 'x,y'}, {a: 2, b: 'z'}]});
  assert.deepEqual(T.decode('tags[2\t]: a\tb'), {tags: ['a', 'b']});
  assert.deepEqual(T.decode('[2]{a,b}:\n  1,2\n  3,4'), [{a: 1, b: 2}, {a: 3, b: 4}]);
  assert.deepEqual(T.decode('[2:]{a}:\n  x: 1\n  y: 2'), {x: {a: 1}, y: {a: 2}});
  assert.equal(T.decode('42'), 42);
  assert.equal(T.decode('"42"'), '42');
  assert.deepEqual(T.decode(''), {});
  assert.deepEqual(T.decode('a: []\nb[0]:\nc:'), {a: [], b: [], c: {}});
  assert.deepEqual(T.decode('n: 1.5000\nm: -1E+03\nz: -0\ns: 05\nt: 1.'), {n: 1.5, m: -1000, z: 0, s: '05', t: '1.'});
  assert.deepEqual(T.decode('items[1]{ a , n{x} , b }:\n  1,2,3'), {items: [{a: 1, n: {x: 2}, b: 3}]}, '項目名の前後の空白');
  const proto = T.decode('__proto__: 1');
  assert.equal(Object.getPrototypeOf(proto), Object.prototype, '__proto__ という名前で、オブジェクトの仕組みを変えない');
  assert.deepEqual(Object.getOwnPropertyDescriptor(proto, '__proto__')?.value, 1, '__proto__ も普通の項目として読む');
});

test('数や形が宣言と合わない・途中で切れた・字下げが崩れたテキストは、読み取らずに止める', () => {
  const bad = {
    '値の数が足りない': 'tags[3]: a,b',
    '表の行が足りない': 'items[3]{id,name}:\n  1,Ada\n  2,Bob',
    '表の行の値が足りない': 'items[2]{id,name}:\n  1,Ada\n  2',
    '並びの要素が足りない': 'items[2]:\n  - 1',
    'キー付きの表の行が足りない': 'users[3:]{age}:\n  a: 1\n  b: 2',
    '引用符が閉じていない': 'name: "Ada',
    '知らないエスケープ': 'name: "bad\\xescape"',
    'サロゲートのエスケープ': 'name: "\\ud800"',
    '閉じた引用符の後に文字': 'name: "x" y',
    '字下げが2の倍数でない': 'user:\n   id: 1',
    '字下げにタブ': 'user:\n\tid: 1',
    '配列の途中に空行': 'items[2]:\n  - 1\n\n  - 2',
    '名前の重複': 'a: 1\na: 2',
    'コロンのない行': 'user:\n  key value',
    '深すぎる字下げ': 'a:\n    b: 1',
    '値の下に字下げした行': 'a: 1\n  b: 2',
    '表の見出しの後に値': 'items[2]{a,b}: 1,2',
    '表の項目名の重複': 'a[1]{x,x}:\n  1,2',
    '要素の数の書き方が不正': 'k[03]: a',
    '名前と [ の間の空白': 'k [2]: a,b',
    '並びの要素に名前のない表の見出し': 'x[1]:\n  - [1]{a}:\n    1',
    '配列の後に続く行': '[2]: a,b\nextra: 1',
  };
  for (const [name, text] of Object.entries(bad)) assert.throws(() => T.decode(text), /TOON/, name);
});

test('対になっていないサロゲートを含む文字列は、仕様どおり書き出さずに止める', () => {
  assert.throws(() => T.encode({v: 'a\ud800b'}), /TOON/);
  assert.throws(() => T.encode({['k\udc00']: 1}), /TOON/);
});
