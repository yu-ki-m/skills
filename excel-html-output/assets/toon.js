/* TOON (Token-Oriented Object Notation): encoder and strict decoder for the JSON data model, no dependencies.
 * Follows the TOON specification version 4.4 (https://github.com/toon-format/spec, SPEC.md); section numbers below
 * refer to it. encode() always uses the comma delimiter and two-space indentation. decode() reads every delimiter
 * (comma, tab, pipe) with two-space indentation, in strict mode only: §14.4's non-strict recoveries are not offered.
 * Numbers decode to JavaScript numbers; a numeric token beyond their range is an error (the policy §4 asks for). */
(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ExcelHTMLToon = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';
  const SPEC_VERSION = '4.4', INDENT = 2;
  const hasOwn = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const isObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);
  const isPrimitive = v => v === null || typeof v !== 'object';

  // ---- Encoding ----

  const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
  const ESCAPE = {'\\': '\\\\', '"': '\\"', '\n': '\\n', '\r': '\\r', '\t': '\\t'};
  const NUMERIC_LIKE = /^[+-]?[0-9]+(?:\.[0-9]+)?(?:e[+-]?[0-9]+)?$/i;

  // §3: a string with an unpaired surrogate cannot be written as UTF-8, so it is refused rather than altered.
  function checked(s) {
    if (LONE_SURROGATE.test(s)) throw Error('TOONに書き出せない文字（対になっていないサロゲート）があります');
    return s;
  }
  // §7.1
  const quote = s => '"' + checked(s).replace(/[\\"\u0000-\u001f]/g,
    c => ESCAPE[c] || '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')) + '"';
  // §7.2 with the comma as the document delimiter: quoted only when a decoder could take it for something else.
  const mustQuote = s => s === '' || /^[ \t]|[ \t]$/.test(s) || s === 'true' || s === 'false' || s === 'null'
    || NUMERIC_LIKE.test(s) || /[:"\\\[\]{}\u0000-\u001f,]/.test(s) || s[0] === '-' || s[0] === '#';
  // A primitive value; root marks a whole-document string, which must not start with what a decoder takes for a BOM.
  const primitive = (v, root = false) => typeof v !== 'string' ? String(v)
    : mustQuote(v) || root && v[0] === '﻿' ? quote(v) : checked(v);
  // §7.3
  const key = k => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(k) ? k : quote(k);

  // §9.3: the fields that all the objects share – the same set of keys, and each key's values either all primitives or
  // all objects that again share fields (a nested field group) – in the first object's key order. null if none.
  function sharedFields(objects) {
    const keys = Object.keys(objects[0]);
    if (!keys.length) return null;
    for (const o of objects) {
      const own = Object.keys(o);
      if (own.length !== keys.length || !keys.every(k => hasOwn(o, k))) return null;
    }
    const fields = [];
    for (const name of keys) {
      const column = objects.map(o => o[name]);
      if (column.every(isPrimitive)) fields.push({name, sub: null});
      else if (column.every(isObject)) {
        const sub = sharedFields(column);
        if (!sub) return null;
        fields.push({name, sub});
      } else return null;
    }
    return fields;
  }
  // §9.3: an array of objects that share fields is a table; §9.5: so is an object of two or more such objects.
  const tableFields = list => list.length && list.every(isObject) ? sharedFields(list) : null;
  const keyedFields = obj => {
    const values = Object.values(obj);
    return values.length >= 2 && values.every(isObject) ? sharedFields(values) : null;
  };
  const fieldList = fields => '{' + fields.map(f => key(f.name) + (f.sub ? fieldList(f.sub) : '')).join(',') + '}';
  // The leaf values of one table row, depth first in the order of the field list.
  function cells(o, fields, out = []) {
    for (const f of fields) if (f.sub) cells(o[f.name], f.sub, out); else out.push(primitive(o[f.name]));
    return out;
  }
  const pad = depth => ' '.repeat(depth * INDENT);

  /** The TOON text of value, taken as JSON.stringify sees it (toJSON, undefined left out, NaN and ±Infinity → null). */
  function encode(value) {
    const json = JSON.stringify(value), data = json === undefined ? null : JSON.parse(json);
    if (isPrimitive(data)) return primitive(data, true);
    const out = [];
    if (Array.isArray(data)) array('', data, 0, out);
    else if (keyedFields(data)) member('', data, 0, out);
    else members(data, 0, out);
    return out.join('\n');
  }
  function members(obj, depth, out) {
    for (const k of Object.keys(obj)) member(key(k), obj[k], depth, out);
  }
  // One member of an object at depth (§8); name is the encoded key, or '' for the root of a keyed table.
  function member(name, v, depth, out) {
    if (Array.isArray(v)) return array(name, v, depth, out);
    if (isPrimitive(v)) return void out.push(pad(depth) + name + ': ' + primitive(v));
    const fields = keyedFields(v);
    if (fields) {
      out.push(pad(depth) + name + '[' + Object.keys(v).length + ':]' + fieldList(fields) + ':');
      for (const k of Object.keys(v)) out.push(pad(depth + 1) + key(k) + ': ' + cells(v[k], fields).join(','));
    } else {
      out.push(pad(depth) + name + ':');
      members(v, depth + 1, out);
    }
  }
  // An array in an object member or at the root (name ''): inline, table, or list (§9).
  function array(name, list, depth, out) {
    const at = pad(depth) + name;
    if (!list.length) return void out.push(name ? at + ': []' : at + '[]');
    if (list.every(isPrimitive)) return void out.push(at + '[' + list.length + ']: ' + list.map(v => primitive(v)).join(','));
    const fields = tableFields(list);
    if (fields) {
      out.push(at + '[' + list.length + ']' + fieldList(fields) + ':');
      for (const o of list) out.push(pad(depth + 1) + cells(o, fields).join(','));
      return;
    }
    out.push(at + '[' + list.length + ']:');
    for (const v of list) item(v, depth + 1, out);
  }
  // One element of a list-form array, with its hyphen at depth (§9.4, §10).
  function item(v, depth, out) {
    const at = pad(depth) + '-';
    if (isPrimitive(v)) out.push(at + ' ' + primitive(v));
    else if (Array.isArray(v)) {
      if (v.every(isPrimitive)) out.push(at + ' [' + v.length + ']:' + (v.length ? ' ' + v.map(x => primitive(x)).join(',') : ''));
      else {
        out.push(at + ' [' + v.length + ']:');
        for (const x of v) item(x, depth + 1, out);
      }
    } else if (!Object.keys(v).length) out.push(at);
    else {
      // The object's members stand one level deeper than the hyphen, and the first of them goes on the hyphen line.
      const first = out.length;
      members(v, depth + 1, out);
      out[first] = at + ' ' + out[first].slice(pad(depth + 1).length);
    }
  }

  // ---- Decoding ----

  const UNESCAPE = {'\\': '\\', '"': '"', n: '\n', r: '\r', t: '\t'};
  const NUMBER = /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?$/;
  const HEADER = /\[(0|[1-9][0-9]*)(:?)([\t|]?)\]/y;
  const fail = (no, why) => { throw Error('TOONを読み取れません（' + no + '行目）: ' + why); };
  // Tokens are trimmed of spaces (U+0020) only (§12).
  function trim(s) {
    let a = 0, b = s.length;
    while (a < b && s.charCodeAt(a) === 32) a++;
    while (b > a && s.charCodeAt(b - 1) === 32) b--;
    return a === 0 && b === s.length ? s : s.slice(a, b);
  }
  // In every scan, a '"' opens a quoted span that the next '"' outside an escape closes (§5.2).
  // classify: the first unquoted ':' and the first unquoted '[' before it (an array header when both exist).
  function classify(s) {
    let bracket = -1;
    for (let k = 0, quoted = false; k < s.length; k++) {
      const c = s.charCodeAt(k);
      if (quoted) { if (c === 92) k++; else if (c === 34) quoted = false; }
      else if (c === 34) quoted = true;
      else if (c === 58) return {colon: k, bracket};
      else if (c === 91 && bracket < 0) bracket = k;
    }
    return {colon: -1, bracket};
  }
  // §9.3: a line under a table header is a row unless an unquoted ':' comes before the first unquoted delimiter.
  function isRow(s, delim) {
    const d = delim.charCodeAt(0);
    for (let k = 0, quoted = false; k < s.length; k++) {
      const c = s.charCodeAt(k);
      if (quoted) { if (c === 92) k++; else if (c === 34) quoted = false; }
      else if (c === 34) quoted = true;
      else if (c === d) return true;
      else if (c === 58) return false;
    }
    return true;
  }
  // §11.2: split on the unquoted delimiter, keeping empty tokens.
  function split(s, delim) {
    const d = delim.charCodeAt(0), out = [];
    let from = 0;
    for (let k = 0, quoted = false; k < s.length; k++) {
      const c = s.charCodeAt(k);
      if (quoted) { if (c === 92) k++; else if (c === 34) quoted = false; }
      else if (c === 34) quoted = true;
      else if (c === d) { out.push(trim(s.slice(from, k))); from = k + 1; }
    }
    out.push(trim(s.slice(from)));
    return out;
  }
  // §7.1, §7.4: a token starting with '"' must be one complete quoted string.
  function unquote(tok, no) {
    let out = '', from = 1;
    for (let k = 1; k < tok.length; k++) {
      const c = tok.charCodeAt(k);
      if (c === 34) {
        if (k !== tok.length - 1) fail(no, '閉じた引用符の後に文字があります');
        return out + tok.slice(from, k);
      }
      if (c !== 92) continue;
      out += tok.slice(from, k);
      const e = tok[k + 1];
      if (e === 'u') {
        const hex = tok.slice(k + 2, k + 6);
        if (!/^[0-9a-fA-F]{4}$/.test(hex)) fail(no, '\\u の後には16進数4桁が要ります');
        const code = parseInt(hex, 16);
        if (code >= 0xd800 && code <= 0xdfff) fail(no, 'サロゲートを表す \\u は使えません');
        out += String.fromCharCode(code);
        k += 5;
      } else if (e !== undefined && hasOwn(UNESCAPE, e)) {
        out += UNESCAPE[e];
        k += 1;
      } else fail(no, '使えないエスケープです: \\' + (e ?? ''));
      from = k + 1;
    }
    fail(no, '引用符が閉じていません');
  }
  // §4
  function value(tok, no) {
    if (tok[0] === '"') return unquote(tok, no);
    if (tok === 'true') return true;
    if (tok === 'false') return false;
    if (tok === 'null') return null;
    if (NUMBER.test(tok)) {
      const n = Number(tok);
      if (!Number.isFinite(n)) fail(no, '数が大きすぎて扱えません: ' + tok);
      return n === 0 ? 0 : n;
    }
    return tok;
  }
  const keyOf = (tok, no) => tok[0] === '"' ? unquote(tok, no) : tok;
  // Every key becomes an own property, __proto__ included (§15); a repeated key is an error (§14.3).
  function put(obj, name, v, no) {
    if (hasOwn(obj, name)) fail(no, '同じ名前が2回あります: ' + name);
    Object.defineProperty(obj, name, {value: v, writable: true, enumerable: true, configurable: true});
  }
  // §6: the field list starting at s[k] === '{'; returns [fields, the index after its '}'].
  function parseFields(s, k, delim, no) {
    const fields = [], names = new Set();
    for (k++; ;) {
      const start = k;
      for (let quoted = false; k < s.length; k++) {
        const c = s[k];
        if (quoted) { if (c === '\\') k++; else if (c === '"') quoted = false; }
        else if (c === '"') quoted = true;
        else if (c === '{' || c === '}' || c === delim) break;
        else if (c === ',' || c === '|' || c === '\t') fail(no, '項目名の区切りが、[ ] で決めた区切りと違います');
      }
      if (k >= s.length) fail(no, '{ が閉じていません');
      const raw = s.slice(start, k), tok = trim(raw);
      if (tok === '') fail(no, '名前のない項目があります');
      if (s[k] === '{' && raw.endsWith(' ')) fail(no, '項目名と { の間に空白は置けません');
      const name = keyOf(tok, no);
      if (names.has(name)) fail(no, '同じ項目名が2回あります: ' + name);
      names.add(name);
      let sub = null;
      if (s[k] === '{') {
        [sub, k] = parseFields(s, k, delim, no);
        while (s[k] === ' ') k++;
      }
      fields.push({name, sub});
      if (s[k] === '}') return [fields, k + 1];
      if (s[k] !== delim) fail(no, '項目名の並びが正しくありません');
      k++;
    }
  }
  // §6: the array header in line s whose bracket segment starts at s[bracket].
  function header(s, bracket, no) {
    const head = s.slice(0, bracket);
    if (head.endsWith(' ') || head.endsWith('\t')) fail(no, '名前と [ の間に空白は置けません');
    HEADER.lastIndex = bracket;
    const m = HEADER.exec(s);
    if (!m) fail(no, '配列の見出しの [ ] の書き方が正しくありません');
    const delim = m[3] || ',', keyed = m[2] === ':';
    let k = bracket + m[0].length, fields = null;
    if (s[k] === '{') [fields, k] = parseFields(s, k, delim, no);
    if (s[k] !== ':') fail(no, '配列の見出しの後に : がありません');
    if (keyed && !fields) fail(no, 'キー付きの表の見出しに { } の項目名がありません');
    const rest = trim(s.slice(k + 1));
    if (fields && rest !== '') fail(no, '表の見出しの : の後には何も書けません');
    return {name: head === '' ? null : keyOf(head, no), length: Number(m[1]), keyed, delim, fields, rest};
  }
  const width = fields => fields.reduce((n, f) => n + (f.sub ? width(f.sub) : 1), 0);
  function record(fields, values, pos, no) {
    const o = {};
    for (const f of fields) put(o, f.name, f.sub ? record(f.sub, values, pos, no) : value(values[pos.k++], no), no);
    return o;
  }

  /**
   * The value of a TOON text. Throws an Error naming the line on anything the specification makes an error in strict
   * mode – among them a count that differs from the declared [N], which is how a text cut short inside an array shows.
   * firstLine: the number to give the text's first line in error messages (when it was taken from a larger text).
   */
  function decode(text, {firstLine = 1} = {}) {
    let source = String(text);
    if (source.charCodeAt(0) === 0xfeff) source = source.slice(1);
    // §12 line preparation; comment lines (§5.1) are left out. Each line is {no, depth, s} or {no, blank: true}.
    const L = [], raw = source.split('\n');
    for (let k = 0; k < raw.length; k++) {
      let s = raw[k], end = s.length;
      if (s.charCodeAt(end - 1) === 13) end--;
      while (end > 0 && s.charCodeAt(end - 1) === 32) end--;
      s = s.slice(0, end);
      const no = firstLine + k;
      let n = 0;
      while (n < s.length && s.charCodeAt(n) === 32) n++;
      if (s[n] === '#') continue;
      if (!s.length) { L.push({no, blank: true}); continue; }
      let m = n;
      while (m < s.length && (s[m] === ' ' || s[m] === '\t')) m++;
      if (m > n) fail(no, '字下げにタブは使えません');
      if (n % INDENT) fail(no, '字下げは空白' + INDENT + 'つを1段とし、その倍数にしてください');
      L.push({no, depth: n / INDENT, s: s.slice(n)});
    }
    let i = 0, last = -1;
    const peek = () => { while (i < L.length && L[i].blank) i++; return L[i]; };
    const take = () => { const line = peek(); last = i++; return line; };
    // §12: no blank line between the first item, row or entry of a header and the last line of its content.
    const unbroken = from => {
      if (from >= 0) for (let k = from; k <= last; k++) if (L[k].blank) fail(L[k].no, '配列や表の途中に空行があります');
    };
    const counted = (n, declared, no) => {
      if (n !== declared) fail(no, '見出しの [' + declared + '] と数が合いません（' + n + '）。テキストが途中で切れているか、書き換えられています');
    };

    // §8: the members of an object whose content depth is depth, up to the first line above that depth.
    function members(depth, obj) {
      for (let line = peek(); line && line.depth >= depth; line = peek()) {
        if (line.depth > depth) fail(line.no, '字下げが深すぎます');
        take();
        member(line.s, depth, line.no, obj);
      }
      return obj;
    }
    // One member line s standing at depth (for the first field of a list item, the depth §10 gives it).
    function member(s, depth, no, obj) {
      const {colon, bracket} = classify(s);
      if (colon < 0) fail(no, '「名前: 値」の形になっていない行があります');
      if (bracket >= 0) {
        const h = header(s, bracket, no);
        if (h.name === null) fail(no, '名前のない配列の見出しは、ここには置けません');
        return put(obj, h.name, contents(h, depth, no), no);
      }
      const rest = trim(s.slice(colon + 1));
      put(obj, keyOf(trim(s.slice(0, colon)), no), rest === '' ? nested(depth) : rest === '[]' ? [] : value(rest, no), no);
    }
    // The object that a 'key:' line at depth opens: its content one level deeper, or {} when nothing is there.
    function nested(depth) {
      const line = peek();
      if (!line || line.depth <= depth) return {};
      if (line.depth > depth + 1) fail(line.no, '字下げが深すぎます（1段ずつ下げてください）');
      return members(depth + 1, {});
    }
    // What a header line at depth declares: inline values, table rows, keyed entry rows, or list items (§9).
    function contents(h, depth, no) {
      if (h.keyed) return entries(h, depth, no);
      if (h.fields) return rows(h, depth, no);
      if (h.rest !== '') {
        const list = split(h.rest, h.delim).map(t => value(t, no));
        counted(list.length, h.length, no);
        return list;
      }
      return items(h, depth, no);
    }
    function rows(h, depth, no) {
      const n = width(h.fields), list = [];
      let from = -1;
      for (let line = peek(); line && line.depth === depth + 1 && isRow(line.s, h.delim); line = peek()) {
        if (from < 0) from = i;
        take();
        const values = split(line.s, h.delim);
        if (values.length !== n) fail(line.no, '表の行の値の数（' + values.length + '）が、見出しの項目の数（' + n + '）と合いません');
        list.push(record(h.fields, values, {k: 0}, line.no));
      }
      unbroken(from);
      counted(list.length, h.length, no);
      return list;
    }
    // §9.5: every line one level deeper is an entry row 'key: values'.
    function entries(h, depth, no) {
      const n = width(h.fields), obj = {};
      let from = -1, count = 0;
      for (let line = peek(); line && line.depth > depth; line = peek()) {
        if (line.depth > depth + 1) fail(line.no, '字下げが深すぎます');
        if (from < 0) from = i;
        take();
        const {colon} = classify(line.s);
        if (colon < 0) fail(line.no, 'キー付きの表の行に : がありません');
        const rest = trim(line.s.slice(colon + 1)), values = rest === '' ? [] : split(rest, h.delim);
        if (values.length !== n) fail(line.no, '表の行の値の数（' + values.length + '）が、見出しの項目の数（' + n + '）と合いません');
        put(obj, keyOf(trim(line.s.slice(0, colon)), line.no), record(h.fields, values, {k: 0}, line.no), line.no);
        count++;
      }
      unbroken(from);
      counted(count, h.length, no);
      return obj;
    }
    function items(h, depth, no) {
      const list = [];
      let from = -1;
      for (let line = peek(); line && line.depth === depth + 1 && (line.s === '-' || line.s.startsWith('- ')); line = peek()) {
        if (from < 0) from = i;
        take();
        list.push(element(line, depth + 1));
      }
      unbroken(from);
      counted(list.length, h.length, no);
      return list;
    }
    // One list item whose hyphen stands at depth (§9.4, §10).
    function element(line, depth) {
      if (line.s === '-') return {};
      const s = trim(line.s.slice(1)), no = line.no, {colon, bracket} = classify(s);
      if (s === '[]') return [];
      if (bracket === 0 && colon >= 0) {
        const h = header(s, 0, no);
        if (h.fields) fail(no, '並びの要素には、名前のない表の見出しは置けません');
        return contents(h, depth, no);
      }
      if (colon < 0) return value(s, no);
      const obj = {};
      member(s, depth + 1, no, obj);
      return members(depth + 1, obj);
    }
    const end = () => { const line = peek(); if (line) fail(line.no, '配列の後に、続きの行があります'); };

    // §5: the root form.
    const first = peek();
    if (!first) return {};
    if (first.depth > 0) fail(first.no, '最初の行に字下げがあります');
    const {colon, bracket} = classify(first.s);
    if (bracket === 0 && colon >= 0) {
      take();
      const result = contents(header(first.s, 0, first.no), 0, first.no);
      end();
      return result;
    }
    if (first.s === '[]') { take(); end(); return []; }
    if (colon < 0 && L.filter(line => !line.blank).length === 1) return value(take().s, first.no);
    return members(0, {});
  }

  return {SPEC_VERSION, encode, decode};
});
