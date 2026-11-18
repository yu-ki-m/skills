/* Excel HTML Output — formulas: how a formula is read and computed, as desktop Excel 16.0 (Japanese) computes it.
   No runtime dependencies. The rules were measured on desktop Excel 16.0 (build 20326): tests/fixtures/excel-functions.mjs
   lists the formulas, and tests/xlsx-functions.test.mjs compares every value Excel saved with the one computed here. */
(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ExcelHTMLFormulas = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';

  /* ---- Values ----------------------------------------------------------------------------------------------------
     A formula works with numbers, texts, TRUE and FALSE, errors (their texts, "#N/A"), BLANK (an empty cell as a formula
     reads it: 0, "" or FALSE, by what it is used for; a formula that gives one shows 0) and ranges (the values of their
     cells row by row, in an array that knows its range). */
  const BLANK = Object.freeze({blank: true});
  const ERROR = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CIRC!)$/;
  const isErr = x => typeof x === 'string' && ERROR.test(x);
  const fail = e => { throw Error(e); };
  /** The error a thrown Error stands for, or the Error again when it is no Excel error (a bug). */
  const errorOf = e => { if (e && isErr(e.message)) return e.message; throw e; };
  /** The longest text an Excel cell holds. & and CONCATENATE cut a longer text there; the other functions give #VALUE!
      (Excel 16.0 shows #CALC! for CONCAT and TEXTJOIN, but saves it as #VALUE! and takes #CALC! from the clipboard only as
      text, so the page keeps to #VALUE!). */
  const MAX_TEXT = 32767;
  const arrayOf = (values, rows, cols) => Object.assign(values, {rows, cols});
  const dims = v => v.range ? [v.range.r2 - v.range.r1 + 1, v.range.c2 - v.range.c1 + 1] : [v.rows ?? v.length, v.cols ?? 1];

  /* ---- Numbers and their digits ----------------------------------------------------------------------------------- */
  /** The exact digits of |x| (a finite double): {d, p}, with |x| = 0.d × 10^p (d without trailing zeros; d is '' for 0). */
  function exactDigits(x) {
    if (x === 0) return {d: '', p: 0};
    const bits = new DataView(new ArrayBuffer(8)); bits.setFloat64(0, Math.abs(x));
    const hi = bits.getUint32(0), e = hi >>> 20, q = e ? e - 1075 : -1074;
    const m = BigInt(hi & 0xfffff) * 4294967296n + BigInt(bits.getUint32(4)) + (e ? 1n << 52n : 0n);
    const s = q >= 0 ? (m << BigInt(q)).toString() : (m * 5n ** BigInt(-q)).toString();
    return {d: s.replace(/0+$/, ''), p: s.length + Math.min(q, 0)};
  }
  /** digits rounded to n significant digits; half: 'down' (an exact half toward zero, as Excel rounds to 15 digits) or
      'up' (half away from zero, as ROUND and TEXT round). */
  function roundDigits({d, p}, n, half) {
    if (d.length <= n) return {d, p};
    if (n < 0) return {d: '', p};
    let head = d.slice(0, n); const rest = d.slice(n);
    if (half === 'down' ? rest > '5' : rest >= '5') { head = n ? String(BigInt(head) + 1n) : '1'; if (head.length > n) p++; }
    head = head.replace(/0+$/, '');
    return {d: head, p: head ? p : 0};
  }
  /** x as Excel holds it for showing and comparing: rounded to 15 significant digits (to 14 where 15 would overflow). */
  function digits15(x) {
    let r = roundDigits(exactDigits(x), 15, 'down');
    if (!Number.isFinite(+('0.' + r.d + 'e' + r.p))) r = roundDigits(exactDigits(x), 14, 'down');
    return r;
  }
  const valueOfDigits = ({d, p}, neg) => d ? (neg ? -1 : 1) * Number('0.' + d + 'e' + p) : 0;
  /** x rounded to 15 significant digits. */
  const to15 = x => Number.isFinite(x) ? valueOfDigits(digits15(x), x < 0) : x;
  /** The digits of x (15 significant) written out in full: integer part and fraction ('' when none). */
  function fixedParts({d, p}) {
    if (!d) return ['0', ''];
    if (p <= 0) return ['0', '0'.repeat(-p) + d];
    return d.length <= p ? [d + '0'.repeat(p - d.length), ''] : [d.slice(0, p), d.slice(p)];
  }
  /** digits rounded half away from zero to k places after the point (k may be negative: tens, hundreds, ...). */
  const roundPlaces = (r, k) => r.d ? roundDigits(r, r.p + k, 'up') : r;
  /** digits cut to k places after the point (toward zero). */
  const truncPlaces = (r, k) => !r.d ? r : r.p + k <= 0 ? {d: '', p: 0} : {d: r.d.slice(0, r.p + k).replace(/0+$/, ''), p: r.d.slice(0, r.p + k).replace(/0+$/, '') ? r.p : 0};

  // A number as Excel writes it when it joins text (=A1&"", CONCAT, TEXTJOIN): rounded to 15 significant digits, an exact
  // half toward zero, and written out in full when that takes 20 characters at most (the sign aside), else as 1.5E-18.
  // The rounding looks at the exact value of the double, as Excel's does: 1234567890.000005 is a little over
  // 1234567890.0000050 and becomes 1234567890.00001. Where 15 digits would round past the largest number, Excel keeps 14.
  function numberText(x) {
    if (Number.isInteger(x) && Math.abs(x) < 1e15) return String(x); // 15 digits at most, nothing to round (-0 too is "0")
    const r = digits15(x), [int, frac] = fixedParts(r), full = int + (frac ? '.' + frac : '');
    if (!r.d) return '0';
    const sign = x < 0 ? '-' : '';
    if (full.length <= 20) return sign + full;
    return sign + r.d[0] + (r.d.length > 1 ? '.' + r.d.slice(1) : '') + 'E' + (r.p > 0 ? '+' : '-') + String(Math.abs(r.p - 1)).padStart(2, '0');
  }

  /* ---- Dates ------------------------------------------------------------------------------------------------------
     Excel's serial numbers: 1 is 1900/01/01 and 60 is 1900/02/29, a day that never was (Lotus 1-2-3's leap year, which
     Excel keeps), so the days from 1900/03/01 on are counted from 1899/12/30. 2958465 is 9999/12/31, the last day. */
  const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30), LAST_DAY = 2958465;
  const utc = (y, m, d) => { const t = new Date(0); t.setUTCFullYear(y, m - 1, d); return t.getTime(); };
  /** The serial number of day d of month m of year y, as DATE counts it (months past 12 or below 1 carry into the year). */
  function serialOf(y, m, d) {
    y += Math.floor((m - 1) / 12); m = ((m - 1) % 12 + 12) % 12 + 1;
    const start = y === 1900 && m <= 2 ? (m === 1 ? 1 : 32) : (utc(y, m, 1) - EPOCH) / DAY;
    return start + d - 1;
  }
  /** The year, month and day of serial number n (0 is 1900/01/00, as Excel shows it). */
  function civil(n) {
    n = Math.floor(n);
    if (n === 60) return {y: 1900, m: 2, d: 29};
    if (n === 0) return {y: 1900, m: 1, d: 0};
    const t = new Date(n < 60 ? utc(1899, 12, 31) + n * DAY : EPOCH + n * DAY);
    return {y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate()};
  }
  const weekday = n => (Math.floor(n) + 6) % 7; // 0 Sunday ... 6 Saturday, as Excel counts from its 1900/01/01, a Sunday
  /** The days of a month as Excel's serial numbers have them (1900/02 has 29), and as the calendar has them (EDATE and
      EOMONTH keep to it: EOMONTH(1,1) is 1900/02/28). */
  const daysIn = (y, m) => y === 1900 && m === 2 ? 29 : new Date(utc(y, m + 1, 0)).getUTCDate();
  const realDaysIn = (y, m) => new Date(utc(y, m + 1, 0)).getUTCDate();
  const validDate = (y, m, d) => y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysIn(y, m);
  /** The clock that TODAY and NOW read (milliseconds since 1970, as Date.now); tests set it. */
  const clock = {now: () => Date.now()};
  /** Now, as a serial number of the computer's local time. */
  function nowSerial() { const t = new Date(clock.now()); return (t.getTime() - t.getTimezoneOffset() * 60000 - EPOCH) / DAY; }
  // Japanese eras, newest first: [name, letter, first day (y, m, d)]. A year of an era is read and written from its first
  // year (令和1 is 2019); Excel reads an era year without checking that the era had begun or not yet ended (R1.4.30 is
  // 2019/04/30), and writes the era of the day.
  const ERAS = [['令和', 'R', 2019, 5, 1], ['平成', 'H', 1989, 1, 8], ['昭和', 'S', 1926, 12, 25], ['大正', 'T', 1912, 7, 30], ['明治', 'M', 1868, 10, 23]];
  const eraOf = ({y, m, d}) => ERAS.find(([, , ey, em, ed]) => y > ey || y === ey && (m > em || m === em && d >= ed)) || ERAS[ERAS.length - 1];
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const monthByName = s => { s = s.toLowerCase(); const i = MONTHS.findIndex(m => m === s || s.length === 3 && m.startsWith(s)); return i < 0 ? 0 : i + 1; };
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], DAYS_JA = ['日', '月', '火', '水', '木', '金', '土'];

  /* ---- Text read as a number (VALUE, and a text in arithmetic) ----------------------------------------------------- */
  // Full-width letters, digits and signs are read as their half-width forms (１２３, ％, ￥); the ideographic space as a space.
  const halfWidth = s => s.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ').replace(/￥/g, '¥');
  /** A plain number: digits with 3-digit groups, a point, an exponent; a sign or parentheses; a currency sign; %. */
  function plainNumber(s) {
    let percent = false;
    if (s.endsWith('%')) { percent = true; s = s.slice(0, -1).trim(); } else if (s.startsWith('%')) { percent = true; s = s.slice(1).trim(); }
    const m = /^([+-]?)\s*([¥$€]?)\s*([+-]?)\s*(\(?)\s*((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:e([+-]?\d+))?\s*(\)?)$/i.exec(s);
    if (!m || m[1] && m[3] || !m[4] !== !m[7] || m[4] && (m[1] || m[3]) || /%/.test(s)) return null;
    const n = Number(m[5].replace(/,/g, '') + (m[6] !== undefined ? 'e' + m[6] : ''));
    if (!Number.isFinite(n)) return null;
    return (m[1] === '-' || m[3] === '-' || m[4] ? -n : n) / (percent ? 100 : 1);
  }
  /** A time of day: h:mm[:ss[.f]] with AM/PM after a space, m:ss.f, or h時m分[s秒]; null when it is none. */
  function timeOf(s) {
    let m = /^(\d+)時(\d+)分(?:(\d+)秒)?$/.exec(s);
    if (m) return (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
    m = /^(\d+):(\d+)\.(\d+)$/.exec(s); // minutes and seconds with a fraction
    if (m) return (+m[1] * 60 + Number(m[2] + '.' + m[3])) / 86400;
    m = /^(\d+):(\d*)(?::(\d*)(?:\.(\d+))?)?(?:\s+(am|pm|a|p))?$/i.exec(s);
    if (!m || m[1].length > 4) return null;
    let h = +m[1];
    if (m[5]) { if (h > 12) return null; h = h % 12 + (/^p/i.test(m[5]) ? 12 : 0); }
    const seconds = Math.round((h * 3600 + +(m[2] || 0) * 60 + Number((m[3] || '0') + '.' + (m[4] || '0'))) * 1000) / 1000;
    return seconds / 86400;
  }
  const fullYear = y => y.length <= 2 ? (+y < 30 ? 2000 + +y : 1900 + +y) : +y;
  /** A day as text: y/m/d, y-m-d, m/d, d/m, y/m, the Japanese y年m月d日 and its eras, R8.1.2, 2-Jan-2026, Jan 2, 2026. */
  function dateOf(s, now) {
    const year = () => civil(now ?? nowSerial()).y;
    const day = (y, mo, d) => validDate(y, mo, d) ? serialOf(y, mo, d) : null;
    let m = /^(\d{1,4})([/-])(\d{1,2})\2(\d{1,4})$/.exec(s);
    if (m) return day(fullYear(m[1]), +m[3], +m[4]);
    m = /^(\d{1,4})[/-](\d{1,4})$/.exec(s);
    if (m) {
      const a = +m[1], b = +m[2];
      if (m[1].length <= 2 && a <= 12 && b <= 31 && m[2].length <= 2) return day(year(), a, b);
      if (m[1].length <= 2 && a <= 31 && b <= 12 && m[2].length <= 2) return day(year(), b, a);
      if (m[1].length >= 3 && b <= 12 && m[2].length <= 2) return day(+m[1], b, 1);
      return null;
    }
    m = /^(明治|大正|昭和|平成|令和)(\d{1,2})年(\d{1,2})月(\d{1,2})日$/.exec(s) || /^([MTSHR])(\d{1,2})[./-](\d{1,2})[./-](\d{1,2})$/i.exec(s);
    if (m) { const era = ERAS.find(e => e[0] === m[1] || e[1] === m[1].toUpperCase()); return day(era[2] + +m[2] - 1, +m[3], +m[4]); }
    m = /^(\d{1,4})年(\d{1,2})月(?:(\d{1,2})日)?$/.exec(s);
    if (m) return day(fullYear(m[1]), +m[2], m[3] ? +m[3] : 1);
    m = /^(\d{1,2})月(\d{1,2})日$/.exec(s);
    if (m) return day(year(), +m[1], +m[2]);
    m = /^(\d{1,2})[-\s]([a-z]+)(?:[-\s](\d{2,4}))?$/i.exec(s); // 2-Jan-2026, 2 January 2026, 2-Jan
    if (m && monthByName(m[2])) return day(m[3] ? fullYear(m[3]) : year(), monthByName(m[2]), +m[1]);
    m = /^([a-z]+)\s+(\d{1,2}),?\s+(\d{2,4})$/i.exec(s) || /^([a-z]+)\s+()(\d{4})$/i.exec(s); // Jan 2, 2026; January 2026
    if (m && monthByName(m[1])) return day(fullYear(m[3]), monthByName(m[1]), m[2] ? +m[2] : 1);
    return null;
  }
  /** A day and a time in either order, or one of them: {day, time} (a part that is not there is null), or null. */
  function dateTimeOf(s, now) {
    const t = timeOf(s); if (t !== null) return {day: null, time: t};
    const d = dateOf(s, now); if (d !== null) return {day: d, time: null};
    const parts = s.split(/\s+/);
    for (let i = 1; i < parts.length; i++) {
      const a = parts.slice(0, i).join(' '), b = parts.slice(i).join(' ');
      if (/[年月日時分秒]/.test(s)) break; // Excel does not join a Japanese day with a time
      let day = /^\d+[/-]|^[a-z]/i.test(a) ? dateOf(a, now) : null, time = day !== null ? timeOf(b) : null;
      if (day !== null && time !== null) return {day, time};
      time = timeOf(a); day = time !== null ? dateOf(b, now) : null;
      if (day !== null && time !== null) return {day, time};
    }
    return null;
  }
  /** A text as Excel reads it into a number (VALUE, and a text in arithmetic), or null. */
  function textNumber(text, now) {
    const s = halfWidth(String(text)).trim();
    if (!s) return null;
    const n = plainNumber(s); if (n !== null) return n;
    const f = /^([+-]?)(\d+)\s+(\d+)\/(\d+)$/.exec(s); // 1 1/2
    if (f && +f[4] > 0) return (f[1] === '-' ? -1 : 1) * (+f[2] + +f[3] / +f[4]);
    const dt = dateTimeOf(s, now);
    if (!dt) return null;
    const v = (dt.day ?? 0) + (dt.time ?? 0);
    return v <= LAST_DAY + 1 ? v : null;
  }

  /* ---- Coercions ---------------------------------------------------------------------------------------------------- */
  /** A value as a number: TRUE 1, FALSE and an empty cell 0, a text read as VALUE reads it; anything else #VALUE!. */
  function toNumber(v) {
    if (typeof v === 'number') return v;
    if (v === BLANK) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isErr(v)) fail(v);
    if (typeof v === 'string') { const n = textNumber(v); if (n === null) fail('#VALUE!'); return n; }
    return fail('#VALUE!');
  }
  /** A value as text: a number as Excel writes it (numberText), TRUE and FALSE, an empty cell as "". */
  function toText(v) {
    if (typeof v === 'string') { if (isErr(v)) fail(v); return v; }
    if (typeof v === 'number') { if (!Number.isFinite(v)) fail('#NUM!'); return numberText(v); }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v === BLANK) return '';
    return fail('#VALUE!');
  }
  /** A value as TRUE or FALSE: a number that is not 0, the texts TRUE and FALSE (in any case), an empty cell as FALSE. */
  function toBool(v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === BLANK) return false;
    if (isErr(v)) fail(v);
    if (typeof v === 'string' && /^(?:true|false)$/i.test(v)) return /^true$/i.test(v);
    return fail('#VALUE!');
  }

  /* ---- Comparison ---------------------------------------------------------------------------------------------------
     Excel orders numbers before texts before FALSE before TRUE, and an empty cell as 0, "" or FALSE by what it is set
     against. Numbers are compared as they are shown, to 15 significant digits (0.1+0.2=0.3 is TRUE). Texts are equal when
     they differ only in case (not in width or kana: "Ａ"="A" and "ア"="ｱ" are FALSE), and are ordered as Windows orders
     them: case does not count, a hyphen or apostrophe only breaks a tie, katakana come before hiragana and half-width
     before full-width (Intl.Collator, with those ties broken as measured). */
  const COLLATOR = typeof Intl === 'object' ? new Intl.Collator('ja', {sensitivity: 'base', ignorePunctuation: true}) : null;
  const sameText = (a, b) => a === b || a.toLowerCase() === b.toLowerCase();
  function tieRank(s) {
    let punct = 0, kana = 0, wide = 0;
    for (const ch of s) {
      const c = ch.codePointAt(0);
      if (/[\p{P}\p{S}]/u.test(ch)) punct++;
      if (c >= 0x3041 && c <= 0x309f) kana++;
      if (c >= 0xff01 && c <= 0xff5e || c >= 0x3000 && c <= 0x30ff) wide++;
    }
    return [punct, kana, wide];
  }
  function compareText(a, b) {
    if (sameText(a, b)) return 0;
    const primary = COLLATOR ? COLLATOR.compare(a, b) : (a.toLowerCase() < b.toLowerCase() ? -1 : 1);
    if (primary) return primary;
    const x = tieRank(a), y = tieRank(b);
    for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] - y[i];
    return a < b ? -1 : a > b ? 1 : 0;
  }
  const rank = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
  /** Excel's order of two values (no errors): negative, 0 or positive. */
  function compare(a, b) {
    if (a === BLANK && b === BLANK) return 0;
    if (a === BLANK) a = typeof b === 'number' ? 0 : typeof b === 'string' ? '' : false;
    if (b === BLANK) b = typeof a === 'number' ? 0 : typeof a === 'string' ? '' : false;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'number') { const x = to15(a), y = to15(b); return x < y ? -1 : x > y ? 1 : 0; }
    if (typeof a === 'string') return compareText(a, b);
    return (a ? 1 : 0) - (b ? 1 : 0);
  }

  /* ---- Wildcards and criteria ---------------------------------------------------------------------------------------- */
  const wildcards = new Map();
  /** A text with * (any characters), ? (one character) and ~ (the next character as it is; a ~ at the end stands for
      nothing) as a pattern, or null when it has none of them. Case does not count. As in Excel, ? is one UTF-16 unit: half
      of a character such as 𩸽 (SEARCH("?b","a𩸽b") is 3). */
  function wildcard(pattern) {
    if (!/[*?~]/.test(pattern)) return null;
    if (!wildcards.has(pattern)) {
      let re = '';
      for (let i = 0; i < pattern.length; i++) {
        const ch = pattern[i];
        if (ch === '~') { if (i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\/]/g, '\\$&'); }
        else if (ch === '*') re += '[\\s\\S]*';
        else if (ch === '?') re += '[\\s\\S]';
        else re += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
      }
      wildcards.set(pattern, new RegExp('^' + re + '$', 'i'));
    }
    return wildcards.get(pattern);
  }
  const matchText = (pattern, text) => { const re = wildcard(pattern); return re ? re.test(text) : sameText(pattern, text); };
  /** The test that a criterion of COUNTIF, SUMIF and the like makes of a cell's value, as measured. A number (also written
      as text, "5" or "=5.0", spaces around it allowed) finds the numbers equal to it and the texts that read as such a
      number; "<>5" leaves out only the numbers. <, <=, > and >= compare numbers with numbers, texts with texts and TRUE
      or FALSE with TRUE or FALSE. A text finds the texts equal to it, without regard to case and with * ? ~. "" finds
      empty cells and empty texts, "=" only empty cells, "<>" all but empty cells. An empty cell given as the criterion is
      0, and an error (also the text "#N/A") finds the cells with that error. */
  function criterion(c) {
    if (isErr(c)) return v => v === c;
    if (c === BLANK) c = 0;
    const SIGN = {'<': x => x < 0, '<=': x => x <= 0, '>': x => x > 0, '>=': x => x >= 0};
    const numberIs = n => v => typeof v === 'number' && compare(v, n) === 0 || typeof v === 'string' && !isErr(v) && textNumber(v) !== null && compare(textNumber(v), n) === 0;
    if (typeof c === 'number') return numberIs(c);
    if (typeof c === 'boolean') return v => v === c;
    const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(c), op = m[1] || '', operand = m[2];
    if (operand === '') {
      if (op === '') return v => v === BLANK || v === '';
      if (op === '=') return v => v === BLANK;
      if (op === '<>') return v => v !== BLANK;
    }
    const n = textNumber(operand), flag = /^(?:true|false)$/i.test(operand) ? /^true$/i.test(operand) : null;
    if (op === '<>') {
      if (isErr(operand)) return v => v !== operand;
      if (n !== null) return v => !(typeof v === 'number' && compare(v, n) === 0);
      if (flag !== null) return v => v !== flag;
      return v => !(typeof v === 'string' && !isErr(v) && matchText(operand, v));
    }
    if (op === '' || op === '=') {
      if (isErr(operand)) return v => v === operand;
      if (n !== null) return numberIs(n);
      if (flag !== null) return v => v === flag;
      return v => typeof v === 'string' && !isErr(v) && matchText(operand, v);
    }
    if (n !== null) return v => typeof v === 'number' && SIGN[op](compare(v, n));
    if (flag !== null) return v => typeof v === 'boolean' && SIGN[op](compare(v, flag));
    return v => typeof v === 'string' && !isErr(v) && SIGN[op](compareText(v, operand));
  }

  /* ---- Number formats (TEXT) ----------------------------------------------------------------------------------------
     A format has up to four sections (positive; negative; zero; text). A section holds literals ("..." and the characters
     that stand for themselves; ! shows the character after it; in the Japanese Excel \ is the yen sign and does not
     escape the next character), digit
     places (0, # and ?), a point, commas (thousands; at the end, a thousand times smaller), % (a hundred times larger),
     E+ (an exponent), / (a fraction), _x (a space for x), *x (filling, which TEXT leaves out), [conditions], [colours]
     (Japanese names only: [Red] is #VALUE!), [$-411] (a locale), G/標準 (General), @ (the text), or date and time codes. */
  const COLOURS = /^(?:黒|青|水|緑|紫|赤|白|黄|色\d+)$/;
  /** The tokens of one format: {k: 'lit', v}, {k: 'cond', op, n}, {k: 'locale', lang}, {k: 'elapsed', unit, len},
      {k: 'general'}, {k: 'at'}, {k: 'sep'} (;), {k: 'code', v} (one format character or a run of date letters). */
  function formatTokens(fmt) {
    const out = [], re = /"([^"]*)"?|\\|!.?|\[([^\]]*)\]|_.|\*.|G\/標準|AM\/PM|am\/pm|A\/P|a\/p|上午\/下午|午前\/午後|([yY]+|[mM]+|[dD]+|a{3,}|[gG]+|[eE]{2,}|[eE](?![+-])|[rR]+|[hH]+|[sS]+|b+)|E[+-]|e[+-]|./gsu;
    let m;
    while ((m = re.exec(fmt))) {
      const t = m[0];
      if (m[1] !== undefined) out.push({k: 'lit', v: m[1].replace(/\\/g, '¥')});
      else if (t === '\\') out.push({k: 'lit', v: '¥'});
      else if (t[0] === '!') { if (t.length < 2) fail('#VALUE!'); out.push({k: 'lit', v: t[1]}); }
      else if (m[2] !== undefined) {
        const b = m[2], cond = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/.exec(b), elapsed = /^(h+|m+|s+)$/i.exec(b), locale = /^\$([^-]*)(?:-(.*))?$/.exec(b);
        if (cond) out.push({k: 'cond', op: cond[1], n: +cond[2]});
        else if (elapsed) out.push({k: 'elapsed', unit: elapsed[1][0].toLowerCase(), len: elapsed[1].length});
        else if (locale) { if (locale[1]) out.push({k: 'lit', v: locale[1]}); out.push({k: 'locale', lang: (locale[2] || '').toLowerCase()}); }
        else if (!COLOURS.test(b)) fail('#VALUE!');
      }
      else if (t[0] === '_') out.push({k: 'lit', v: ' '});
      else if (t[0] === '*') continue;
      else if (t === 'G/標準') out.push({k: 'general'});
      else if (t === '午前/午後') fail('#VALUE!');
      else if (t === ';') out.push({k: 'sep'});
      else if (t === '@') out.push({k: 'at'});
      // B, n and N are no format characters and no literals: Excel's TEXT gives #VALUE! for them.
      else if (/^[BnN]$/.test(t)) fail('#VALUE!');
      else out.push({k: 'code', v: t});
    }
    return out;
  }
  const DATE_CODE = /^(?:[yYmMdDgGrRhHsS]+|b+|a{3,}|[eE]{1,}|AM\/PM|am\/pm|A\/P|a\/p|上午\/下午)$/;
  // e (and E without a sign) is the year of the era, a date code: TEXT(46024,"e") is 8, and "0.00E" is #VALUE!.
  const isDateSection = tokens => tokens.some(t => t.k === 'elapsed' || t.k === 'code' && DATE_CODE.test(t.v));
  const AMPM = /^(?:AM\/PM|A\/P|上午\/下午)$/i;
  /** x in Excel's General format of TEXT: 11 characters at most (the sign aside), else 1.23457E+11. */
  function generalText(x) {
    if (x === 0) return '0';
    const sign = x < 0 ? '-' : '', r = digits15(Math.abs(x)), e = r.p - 1;
    const science = () => {
      let s = roundDigits(r, Math.abs(e) >= 100 ? 5 : 6, 'up'); const ex = s.p - 1;
      return sign + s.d[0] + (s.d.length > 1 ? '.' + s.d.slice(1) : '') + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0');
    };
    if (e >= 11) return science();
    const intLen = Math.max(e + 1, 1), k = 11 - intLen - 1;
    if (e <= -5 && -(e + 1) + r.d.length > k) return science();
    const s = roundPlaces(r, Math.max(k, 0)), [int, frac] = fixedParts(s);
    if (int.length > 11) return science();
    return sign + int + (frac ? '.' + frac : '');
  }
  /** The sections of a format, and the one that x takes: {tokens, sign} (sign: whether a minus goes in front). */
  function sectionFor(sections, x) {
    const conditional = sections.some(s => s.some(t => t.k === 'cond'));
    const test = s => { const c = s.find(t => t.k === 'cond'); if (!c) return null; return {'<': x < c.n, '<=': x <= c.n, '>': x > c.n, '>=': x >= c.n, '=': x === c.n, '<>': x !== c.n}[c.op]; };
    const numeric = sections.length === 4 ? sections.slice(0, 3) : sections.filter((s, i) => i < 3);
    if (conditional) {
      for (let i = 0; i < numeric.length; i++) {
        const t = test(numeric[i]);
        if (t === true) return {tokens: numeric[i], sign: i === 0 && x < 0 && !numeric[i].some(u => u.k === 'cond' && u.n < 0)};
        if (t === null) return {tokens: numeric[i], sign: false};
      }
      return {tokens: numeric[numeric.length - 1], sign: x < 0};
    }
    if (numeric.length === 1 || x > 0 || x === 0 && numeric.length < 3) return {tokens: numeric[0], sign: x < 0};
    if (x < 0) return {tokens: numeric[1], sign: false};
    return {tokens: numeric[2], sign: false};
  }
  /** x (a number, its sign already taken care of) in a section of digit places. */
  function formatDigits(x, tokens) {
    tokens = tokens.filter(t => t.k !== 'cond' && t.k !== 'locale');
    const isPlace = t => t.k === 'code' && /^[0#?]$/.test(t.v);
    let point = -1, exp = -1, slash = -1, percent = 0;
    tokens.forEach((t, i) => {
      if (t.k !== 'code') return;
      if (t.v === '.' && point < 0 && exp < 0 && slash < 0) point = i;
      else if (/^[eE][+-]$/.test(t.v) && exp < 0) exp = i;
      else if (/^[eE]$/.test(t.v) && exp < 0 && tokens.slice(i + 1).some(isPlace)) fail('#VALUE!');
      else if (t.v === '/' && slash < 0 && tokens.slice(0, i).some(isPlace) && tokens.slice(i + 1).some(u => isPlace(u) || u.k === 'code' && /^\d$/.test(u.v))) slash = i;
      else if (t.v === '%') percent++;
    });
    if (percent > 1) fail('#VALUE!');
    const intEnd = point >= 0 ? point : exp >= 0 ? exp : tokens.length;
    if (slash >= 0) return formatFraction(x * (percent ? 100 : 1), tokens, slash);
    // Commas: between digit places they group thousands; right after the last digit place before the point, each one
    // makes the number a thousand times smaller.
    const lastInt = tokens.slice(0, intEnd).map(isPlace).lastIndexOf(true), firstInt = tokens.slice(0, intEnd).findIndex(isPlace);
    const numEnd = exp >= 0 ? exp : tokens.length, lastPlace = tokens.slice(0, numEnd).map(isPlace).lastIndexOf(true);
    let scale = 0, group = false;
    tokens.forEach((t, i) => {
      if (t.k !== 'code' || t.v !== ',' || i >= numEnd) return;
      if (i > lastPlace) { if (tokens.slice(lastPlace + 1, i).every(u => u.k === 'code' && u.v === ',')) scale++; }
      else if (i > firstInt && i < lastInt) group = true;
    });
    const isComma = (t, i) => t.k === 'code' && t.v === ',' && i < numEnd;
    x = x * 100 ** percent / 1000 ** scale;
    const fracPlaces = point >= 0 ? tokens.slice(point + 1, exp >= 0 ? exp : tokens.length).filter(isPlace) : [];
    const intPlaces = tokens.slice(0, intEnd).filter(isPlace);
    let intDigits, fracDigits, expText = null;
    if (exp >= 0) {
      const ni = Math.max(intPlaces.length, 1), expPlaces = tokens.slice(exp + 1).filter(isPlace);
      let r = digits15(x), e = r.d ? r.p - 1 : 0;
      let shift = ni > 1 && intPlaces.some(t => t.v === '#') ? Math.floor(e / ni) * ni : e - (ni - 1);
      let m = roundPlaces({d: r.d, p: r.p - shift}, fracPlaces.length);
      if (m.d && m.p > (ni > 1 && intPlaces.some(t => t.v === '#') ? ni : ni)) { shift += ni > 1 && intPlaces.some(t => t.v === '#') ? ni : 1; m = roundPlaces({d: r.d, p: r.p - shift}, fracPlaces.length); }
      [intDigits, fracDigits] = fixedParts(m);
      const sign = tokens[exp].v[1], ex = m.d ? shift : 0;
      expText = tokens[exp].v[0] + (ex < 0 ? '-' : sign === '+' ? '+' : '') + String(Math.abs(ex)).padStart(expPlaces.length, '0');
    } else {
      const r = roundPlaces(digits15(x), fracPlaces.length);
      [intDigits, fracDigits] = fixedParts(r);
    }
    if (intDigits === '0') intDigits = '';
    fracDigits = fracDigits.padEnd(fracPlaces.length, '0').slice(0, fracPlaces.length);
    // The fraction: trailing zeros at # places go, at ? places become spaces.
    const frac = fracPlaces.map((t, i) => fracDigits[i]);
    for (let i = frac.length - 1; i >= 0 && frac[i] === '0' && fracPlaces[i].v !== '0'; i--) frac[i] = fracPlaces[i].v === '?' ? ' ' : '';
    // The integer: digits fill the places from the right; the first place takes what is left.
    const literalsInside = tokens.slice(firstInt, lastInt + 1).some(t => t.k === 'lit' || t.k === 'code' && !/^[0#?,]$/.test(t.v));
    let intText;
    if (!literalsInside) {
      const need = intPlaces.length - intPlaces.findIndex(t => t.v !== '#');
      let s = intDigits;
      if (intPlaces.some(t => t.v !== '#') && s.length < need) {
        const pad = intPlaces.slice(intPlaces.length - need).slice(0, need - s.length).map(t => t.v === '?' ? ' ' : t.v === '0' ? '0' : '').join('');
        s = pad + s;
      }
      if (group) s = s.replace(/\d(?=(?:\d{3})+$)/g, '$&,').replace(/^(\s*),/, '$1');
      intText = s;
    }
    const out = []; let placed = intPlaces.length, digits = [...intDigits];
    let fracIndex = 0, inExp = false, expDone = false;
    tokens.forEach((t, i) => {
      if (i === exp) { inExp = true; return; }
      if (inExp) { if (isPlace(t)) { if (!expDone) { out.push(expText); expDone = true; } return; } }
      if (isPlace(t)) {
        if (i < intEnd) {
          if (!literalsInside) { if (i === firstInt) out.push(intText); return; }
          placed--;
          const take = placed === 0 ? digits.splice(0).join('') : digits.length > placed ? digits.splice(digits.length - 1, 1)[0] : digits.length && digits.length === placed + 1 ? digits.splice(digits.length - 1, 1)[0] : null;
          out.push(take ?? (t.v === '0' ? '0' : t.v === '?' ? ' ' : ''));
          return;
        }
        out.push(frac[fracIndex++] ?? ''); return;
      }
      if (isComma(t, i)) return;
      if (t.k === 'code' && t.v === '%') { out.push('%'); return; }
      if (i === point) { out.push('.'); return; }
      if (t.k === 'lit' || t.k === 'code') out.push(t.v);
    });
    if (literalsInside) {
      // With literals among the integer places, the digits are given out from the right, one place at a time.
      const places = tokens.map((t, i) => i < intEnd && isPlace(t) ? i : -1).filter(i => i >= 0), assigned = new Map();
      const ds = [...intDigits];
      for (let k = places.length - 1; k >= 0; k--) {
        const t = tokens[places[k]];
        if (k === 0) assigned.set(places[k], ds.length ? ds.join('') : t.v === '0' ? '0' : t.v === '?' ? ' ' : '');
        else assigned.set(places[k], ds.length ? ds.pop() : t.v === '0' ? '0' : t.v === '?' ? ' ' : '');
      }
      out.length = 0; fracIndex = 0; inExp = false; expDone = false;
      tokens.forEach((t, i) => {
        if (i === exp) { inExp = true; return; }
        if (inExp && isPlace(t)) { if (!expDone) { out.push(expText); expDone = true; } return; }
        if (isPlace(t)) { out.push(i < intEnd ? assigned.get(i) : frac[fracIndex++] ?? ''); return; }
        if (isComma(t, i)) return;
        if (i === point) { out.push('.'); return; }
        if (t.k === 'lit' || t.k === 'code') out.push(t.v);
      });
    }
    return out.join('');
  }
  /** x as a fraction: # ?/?, ?/?, # ?/3, #/4 ... */
  function formatFraction(x, tokens, slash) {
    const isPlace = t => t.k === 'code' && /^[0#?]$/.test(t.v);
    const den = tokens.slice(slash + 1).filter(t => isPlace(t) || t.k === 'code' && /^\d$/.test(t.v));
    const fixed = den.every(t => /^\d$/.test(t.v) && !/^[0#?]$/.test(t.v)) ? +den.map(t => t.v).join('') : 0;
    // The places before the slash: the numerator is the run of places just before it; places before a literal are the whole part.
    let numStart = slash; while (numStart > 0 && isPlace(tokens[numStart - 1])) numStart--;
    const hasWhole = tokens.slice(0, numStart).some(isPlace);
    let whole = hasWhole ? Math.trunc(x) : 0, rest = x - whole;
    let best = [0, 1];
    if (fixed) best = [Math.round(rest * fixed), fixed];
    else {
      const max = 10 ** den.length - 1; let err = Infinity;
      for (let q = 1; q <= max; q++) { const p = Math.round(rest * q), e = Math.abs(rest - p / q); if (e < err - 1e-12) { err = e; best = [p, q]; } }
    }
    if (hasWhole && best[0] === best[1]) { whole += 1; best[0] = 0; }
    const numPlaces = tokens.slice(numStart, slash).filter(isPlace), denPlaces = den;
    const out = [];
    tokens.forEach((t, i) => {
      if (i < numStart) { if (isPlace(t)) { if (tokens.slice(0, i).some(isPlace)) return; out.push(whole ? String(whole) : t.v === '0' ? '0' : ''); } else if (t.k === 'lit' || t.k === 'code') out.push(t.v); return; }
      if (i >= numStart && i < slash) { if (i === numStart) out.push(String(best[0]).padStart(numPlaces.length, ' ')); return; }
      if (i === slash) { out.push('/'); return; }
      if (den.includes(t)) { if (t === den[0]) out.push(fixed ? String(fixed) : String(best[1]).padEnd(denPlaces.length, ' ')); return; }
      if (t.k === 'lit' || t.k === 'code') out.push(t.v);
    });
    return out.join('');
  }
  /** Where a section of date codes shows a fraction of a second (the 0s after s.), as [index of the point, number of 0s];
      another digit place there is #VALUE! ("0.00E" and "0;0y" too, where e and y are date codes). */
  function secondFraction(tokens) {
    const fracTok = tokens.findIndex((t, i) => t.k === 'code' && t.v === '.' && i > 0 && tokens.slice(0, i).some(u => u.k === 'code' && /^[sS]+$/.test(u.v) || u.k === 'elapsed' && u.unit === 's'));
    let fracLen = 0;
    if (fracTok >= 0) for (let i = fracTok + 1; i < tokens.length && tokens[i].k === 'code' && tokens[i].v === '0'; i++) fracLen++;
    if (tokens.some((t, i) => t.k === 'code' && /^[0#?]$/.test(t.v) && !(fracTok >= 0 && i > fracTok && i <= fracTok + fracLen))) fail('#VALUE!');
    return [fracTok, fracLen];
  }
  /** Serial number n in a section of date and time codes. */
  function formatDate(n, tokens) {
    if (n < 0) fail('#VALUE!');
    const english = tokens.some(t => t.k === 'locale' && /^(?:409|en)/.test(t.lang));
    const codes = tokens.map((t, i) => [t, i]).filter(([t]) => t.k === 'elapsed' || t.k === 'code' && DATE_CODE.test(t.v));
    // The smallest unit shown: a fraction of a second after s, or the second; the time is rounded to it first.
    const [fracTok, fracLen] = secondFraction(tokens);
    const unit = 10 ** -Math.min(fracLen, 3), units = Math.round(n * 86400 / unit), total = units * unit;
    const day = Math.floor(total / 86400 + 1e-9), seconds = total - day * 86400;
    if (day > LAST_DAY) fail('#VALUE!');
    const {y, m, d} = civil(day), wd = weekday(day), era = eraOf({y, m, d}), eraYear = y - era[2] + 1;
    const ampm = codes.some(([t]) => t.k === 'code' && AMPM.test(t.v));
    const h = Math.floor(seconds / 3600), mi = Math.floor(seconds % 3600 / 60), s = Math.floor(seconds % 60 + 1e-9);
    const out = [];
    tokens.forEach((t, i) => {
      if (t.k === 'lit') { out.push(t.v); return; }
      if (t.k === 'cond' || t.k === 'locale') return;
      if (t.k === 'elapsed') {
        const v = t.unit === 'h' ? Math.floor(total / 3600 + 1e-9) : t.unit === 'm' ? Math.floor(total / 60 + 1e-9) : Math.floor(total + 1e-9);
        out.push(String(v).padStart(t.len, '0')); return;
      }
      if (t.k !== 'code') return;
      const v = t.v, c = v[0].toLowerCase(), len = v.length;
      if (i > fracTok && fracTok >= 0 && v === '0' && i <= fracTok + fracLen) { out.push(String(Math.round((total % 1) / unit * unit * 10 ** fracLen)).padStart(fracLen, '0')[i - fracTok - 1] ?? '0'); return; }
      if (AMPM.test(v)) {
        const pm = h >= 12;
        out.push(v === '上午/下午' ? (pm ? '下午' : '上午') : /^am\/pm$/i.test(v) ? (pm ? v.slice(3) : v.slice(0, 2)) : (pm ? v[2] : v[0])); return;
      }
      if (!DATE_CODE.test(v)) { out.push(v); return; }
      const near = (dir) => { for (let j = codes.findIndex(([, k]) => k === i) + dir; j >= 0 && j < codes.length; j += dir) { const [u] = codes[j]; if (u.k === 'elapsed') return u.unit; if (AMPM.test(u.v)) continue; return u.v[0].toLowerCase(); } return ''; };
      if (c === 'y') out.push(len <= 2 ? String(y % 100).padStart(2, '0') : String(y));
      else if (c === 'm' && len <= 2 && (near(-1) === 'h' || near(1) === 's')) out.push(String(mi).padStart(len, '0'));
      else if (c === 'm') out.push(len === 1 ? String(m) : len === 2 ? String(m).padStart(2, '0') : len === 3 ? MONTHS[m - 1].slice(0, 3).replace(/^./, ch => ch.toUpperCase()) : len === 5 ? MONTHS[m - 1][0].toUpperCase() : MONTHS[m - 1].replace(/^./, ch => ch.toUpperCase()));
      else if (c === 'd') out.push(len === 1 ? String(d) : len === 2 ? String(d).padStart(2, '0') : len === 3 ? DAYS[wd].slice(0, 3) : DAYS[wd]);
      else if (c === 'a') out.push(english ? (len === 3 ? DAYS[wd].slice(0, 3) : DAYS[wd]) : len === 3 ? DAYS_JA[wd] : DAYS_JA[wd] + '曜日');
      else if (c === 'g') out.push(len === 1 ? era[1] : len === 2 ? era[0][0] : era[0]);
      else if (c === 'b') out.push(len <= 2 ? String((y + 543) % 100).padStart(2, '0') : String(y + 543)); // the Buddhist year
      else if (c === 'e') out.push(len === 1 ? String(eraYear) : String(eraYear).padStart(2, '0'));
      else if (c === 'r') out.push((len >= 2 ? era[0] : '') + String(eraYear).padStart(2, '0'));
      else if (c === 'h') out.push(String(ampm ? (h % 12 || 12) : h).padStart(len >= 2 ? 2 : 1, '0'));
      else if (c === 's') out.push(String(s).padStart(len >= 2 ? 2 : 1, '0'));
    });
    return out.join('');
  }
  /** TEXT(value, format): value is a number, a text, TRUE / FALSE or an empty cell. */
  function formatValue(value, fmt) {
    if (value === BLANK) value = 0;
    if (typeof value === 'string') { const n = textNumber(value); if (n !== null) value = n; }
    if (fmt === '') return '';
    if (/^general$/i.test(fmt.trim())) fail('#VALUE!');
    const tokens = formatTokens(fmt), sections = [[]];
    for (const t of tokens) { if (t.k === 'sep') sections.push([]); else sections[sections.length - 1].push(t); }
    // Every section must be right, also one the value does not use: TEXT(1,"0;0y") is #VALUE!.
    sections.forEach((section, i) => { if (isDateSection(section)) { if (i === 3) fail('#VALUE!'); secondFraction(section); } });
    if (typeof value !== 'number') {
      const text = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value;
      const section = sections.length >= 4 ? sections[3] : sections.length === 1 && sections[0].some(t => t.k === 'at') ? sections[0] : null;
      if (!section) return text;
      return section.map(t => t.k === 'at' ? text : t.k === 'lit' ? t.v : t.k === 'code' ? t.v : '').join('');
    }
    const {tokens: section, sign} = sectionFor(sections, value);
    if (!section.length) return '';
    const x = sign || sections.length === 1 || section !== sections[0] ? Math.abs(value) : value;
    if (section.some(t => t.k === 'general')) {
      if (section.some(t => t.k === 'code' && /^[0#?]$/.test(t.v))) fail('#VALUE!');
      return (sign ? '-' : '') + section.map(t => t.k === 'general' ? generalText(Math.abs(value)) : t.k === 'lit' || t.k === 'code' ? t.v : '').join('');
    }
    if (isDateSection(section)) return formatDate(value, section);
    if (section.some(t => t.k === 'at') && !section.some(t => t.k === 'code' && /^[0#?]$/.test(t.v))) return section.map(t => t.k === 'at' ? generalText(value) : t.k === 'lit' || t.k === 'code' ? t.v : '').join('');
    const body = formatDigits(Math.abs(x), section);
    const zero = !/[1-9]/.test(body.replace(/E[+-]?\d+/i, ''));
    return (sign && !zero ? '-' : '') + body;
  }

  /* ---- Width of characters (ASC and DBCS, which the Japanese Excel shows as JIS) -------------------------------------
     As measured: ASC makes full-width letters, digits, signs, katakana and Japanese punctuation half-width (ガ to ｶﾞ),
     but leaves ヴ ヵ ヶ ヮ ヰ ヱ, hiragana, “ ￣ 〜 and − alone; ’ becomes ', ” becomes ", ‘ becomes `, ￥ becomes \ and ―
     becomes ｰ. DBCS does the reverse, joining ﾞ and ﾟ to the kana before them where a voiced kana exists (ｶﾞ to ガ, but
     ｳﾞ to ウ゛), and makes ' ’, " ”, ` ‘ and \ ￥. */
  const HALF_KANA = 'ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ';
  const FULL_KANA = 'ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜';
  const VOICED = 'カキクケコサシスセソタチツテトハヒフヘホ', VOICED_TO = 'ガギグゲゴザジズゼゾダヂヅデドバビブベボ', SEMI = 'ハヒフヘホ', SEMI_TO = 'パピプペポ';
  const HALF_SIGNS = {'。': '｡', '「': '｢', '」': '｣', '、': '､', '・': '･', '’': "'", '”': '"', '‘': '`', '￥': '\\', '―': 'ｰ', '　': ' '};
  const FULL_SIGNS = {"'": '’', '"': '”', '`': '‘', '\\': '￥', ' ': '　', '｡': '。', '｢': '「', '｣': '」', '､': '、', '･': '・'};
  function toHalf(s) {
    let out = '';
    for (const ch of s) {
      const c = ch.charCodeAt(0), k = FULL_KANA.indexOf(ch);
      if (c >= 0xff01 && c <= 0xff5e) out += String.fromCharCode(c - 0xfee0);
      else if (HALF_SIGNS[ch]) out += HALF_SIGNS[ch];
      else if (k >= 0) out += HALF_KANA[k];
      else if (VOICED_TO.includes(ch)) out += HALF_KANA[FULL_KANA.indexOf(VOICED[VOICED_TO.indexOf(ch)])] + 'ﾞ';
      else if (SEMI_TO.includes(ch)) out += HALF_KANA[FULL_KANA.indexOf(SEMI[SEMI_TO.indexOf(ch)])] + 'ﾟ';
      else out += ch;
    }
    return out;
  }
  function toFull(s) {
    let out = '';
    const chars = [...s];
    for (let i = 0; i < chars.length; i++) {
      const ch = chars[i], c = ch.charCodeAt(0), k = HALF_KANA.indexOf(ch), next = chars[i + 1];
      if (k >= 0) {
        const full = FULL_KANA[k];
        if (next === 'ﾞ' && VOICED.includes(full)) { out += VOICED_TO[VOICED.indexOf(full)]; i++; }
        else if (next === 'ﾟ' && SEMI.includes(full)) { out += SEMI_TO[SEMI.indexOf(full)]; i++; }
        else out += full;
      }
      else if (FULL_SIGNS[ch]) out += FULL_SIGNS[ch];
      else if (c >= 0x21 && c <= 0x7e) out += String.fromCharCode(c + 0xfee0);
      else out += ch;
    }
    return out;
  }

  /* ---- Code page 932 (CHAR and CODE) ----------------------------------------------------------------------------------
     CHAR takes 1-255 as a byte of the Japanese code page (161-223 half-width katakana; a lead byte alone is a space) and
     a larger number as a JIS code (9250 is あ); CODE gives the byte of a one-byte character, the JIS code of a two-byte
     one, and 63 (?) for a character the code page lacks. Shift_JIS is read with the browser's decoder. */
  let sjisDecoder = null, sjisOf = null;
  const decoder = () => sjisDecoder ??= new TextDecoder('shift_jis');
  function charOf(n) {
    if (n >= 1 && n <= 255) {
      if (n === 0x80 || n === 0xa0 || n >= 0xfd) return String.fromCharCode(n);
      if (n >= 0x81 && n <= 0x9f || n >= 0xe0 && n <= 0xfc) return ' ';
      return n < 0x80 ? String.fromCharCode(n) : decoder().decode(new Uint8Array([n]));
    }
    const j1 = n >> 8, j2 = n & 0xff;
    if (j1 < 0x21 || j1 > 0xfc || j2 < 0x21 || j2 > 0x7e) fail('#VALUE!');
    const s1 = ((j1 + 1) >> 1) + (j1 <= 0x5e ? 0x70 : 0xb0), s2 = j1 % 2 ? j2 + (j2 <= 0x5f ? 0x1f : 0x20) : j2 + 0x7e;
    const ch = decoder().decode(new Uint8Array([s1, s2]));
    if (ch === '�' || ch.length !== 1) fail('#VALUE!');
    return ch;
  }
  function codeOf(ch) {
    const c = ch.codePointAt(0);
    if (c < 0x80) return c;
    if (!sjisOf) {
      // Each character's first two-byte code, leaving out the NEC-selected IBM rows (0xED, 0xEE), as Windows encodes.
      sjisOf = new Map(); const dec = decoder();
      for (let b = 0xa1; b <= 0xdf; b++) sjisOf.set(dec.decode(new Uint8Array([b])), b);
      for (let s1 = 0x81; s1 <= 0xfc; s1++) {
        if (s1 >= 0xa0 && s1 <= 0xdf || s1 === 0xed || s1 === 0xee) continue;
        for (let s2 = 0x40; s2 <= 0xfc; s2++) {
          if (s2 === 0x7f) continue;
          const t = dec.decode(new Uint8Array([s1, s2]));
          if (t.length === 1 && t !== '�' && !sjisOf.has(t)) sjisOf.set(t, s1 << 8 | s2);
        }
      }
    }
    const s = sjisOf.get(String.fromCodePoint(c));
    if (s === undefined) return 63;
    if (s < 0x100) return s;
    let s1 = s >> 8, s2 = s & 0xff, j1 = (s1 <= 0x9f ? s1 - 0x71 : s1 - 0xb1) * 2 + 1, j2;
    if (s2 > 0x9e) { j1++; j2 = s2 - 0x7e; } else j2 = s2 - (s2 >= 0x80 ? 0x20 : 0x1f);
    return j1 << 8 | j2;
  }

  /* ---- Reading a formula ------------------------------------------------------------------------------------------
     References, numbers, texts, errors, function calls (an argument may be left empty: =IF(TRUE,) or =TEXTJOIN(,,A1:A3)),
     and operators with Excel's precedence: a minus sign before ^ (=-2^2 is 4), %, ^, * /, + -, &, comparisons; all of
     them from left to right (=2^3^2 is 64). A reference is not a function name (LOG10( is the function). */
  const TOKEN = /\s*(?:(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|("(?:[^"]|"")*")|(\$?[A-Z]{1,3}\$?[1-9]\d*)(?![A-Z0-9_.(])|([A-Z_][A-Z_0-9.]*)|(<=|>=|<>|[+\-*/^%&=<>(),:])|(#(?:REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)))/iy;
  const PREC = {'=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5};
  // The Japanese Excel shows DBCS as JIS; a formula written with JIS is read as DBCS.
  const ALIASES = {JIS: 'DBCS'};
  /** The formula with the functions under the names every Excel reads (=JIS( becomes =DBCS(): Excel takes a formula on
      the clipboard or in a file only so, and a JIS there is #NAME?). Texts in quotes stay as they are. */
  const englishNames = f => f.replace(/"(?:[^"]|"")*"|(?<![A-Za-z0-9_.])[A-Za-z_][A-Za-z0-9_.]*(?=\s*\()/g, t => t[0] === '"' ? t : ALIASES[t.toUpperCase()] || t);
  /** A number written in a formula, as Excel reads it: to 15 significant digits, the rest cut off. Excel saves
      =INT(2.9999999999999996) as INT(2.99999999999999), and its value is 2. */
  function numberLiteral(text) {
    const m = /^(\d*)\.?(\d*)(?:e([+-]?\d+))?$/i.exec(text);
    let digits = m[1] + m[2], point = m[1].length + (m[3] ? +m[3] : 0);
    const lead = digits.match(/^0*/)[0].length;
    digits = digits.slice(lead); point -= lead;
    return digits.length <= 15 ? +text : Number('0.' + digits.slice(0, 15) + 'e' + point);
  }
  function parse(formula) {
    const tok = []; let p = 1, m;
    while (p < formula.length) {
      if (!formula.slice(p).trim()) break;
      TOKEN.lastIndex = p; m = TOKEN.exec(formula);
      if (!m) fail('#NAME?');
      tok.push(m[1] ? {k: 'num', v: numberLiteral(m[1])} : m[2] ? {k: 'str', v: m[2].slice(1, -1).replace(/""/g, '"')} : m[3] ? {k: 'ref', v: m[3].toUpperCase()}
        : m[4] ? {k: 'id', v: m[4].toUpperCase()} : m[5] ? {k: m[5], v: m[5]} : {k: 'err', v: m[6].toUpperCase()});
      p = TOKEN.lastIndex;
    }
    let i = 0; const peek = () => tok[i]?.k, eat = k => { if (peek() !== k) fail('#VALUE!'); return tok[i++]; };
    function expression(min = 0) {
      let left;
      if (peek() === '+' || peek() === '-') { const sign = tok[i++].k; left = {op: 'unary', sign, x: expression(6)}; }
      else if (peek() === '(') { i++; left = expression(); eat(')'); }
      else if (peek() === 'num' || peek() === 'str' || peek() === 'err') left = {op: 'literal', v: tok[i++].v};
      else if (peek() === 'ref') { const a = tok[i++].v; let b; if (peek() === ':') { i++; b = eat('ref').v; } left = {op: 'ref', a, b}; }
      else if (peek() === 'id') {
        const name = tok[i++].v;
        if (['TRUE', 'FALSE'].includes(name) && peek() !== '(') left = {op: 'literal', v: name === 'TRUE'};
        else {
          eat('('); const args = [];
          const arg = () => peek() === ',' || peek() === ')' ? {op: 'missing'} : expression();
          if (peek() !== ')') { args.push(arg()); while (peek() === ',') { i++; args.push(arg()); } }
          eat(')'); left = {op: 'call', name: ALIASES[name] || name, args};
        }
      } else fail('#VALUE!');
      while (true) {
        if (peek() === '%') { i++; left = {op: 'percent', x: left}; continue; }
        const op = peek(), pr = PREC[op];
        if (pr === undefined || pr < min) break;
        i++; left = {op: 'binary', sign: op, l: left, r: expression(pr + 1)};
      }
      return left;
    }
    const ast = expression();
    if (i !== tok.length) fail('#NAME?');
    return ast;
  }

  /* ---- Computing a formula ----------------------------------------------------------------------------------------- */
  const colName = n => { let out = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out; };
  const cellName = (r, c) => colName(c) + (r + 1);
  const coordOf = a => { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a); return [+m[2] - 1, [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1]; };
  /** A range used where one value is taken gives, as in Excel, its cell in the row of the formula (a range in one column)
      or in its column (a range in one row): the implicit intersection. A range across rows and columns, or one that the
      row or column of the formula misses, is #VALUE!. Another array gives its first value. */
  function single(v, env) {
    if (!Array.isArray(v)) return v;
    const q = v.range;
    if (!q) return v.length ? v[0] : fail('#VALUE!');
    const known = env.row !== undefined && env.row !== null;
    const r = q.r1 === q.r2 ? q.r1 : known && env.row >= q.r1 && env.row <= q.r2 ? env.row : -1;
    const c = q.c1 === q.c2 ? q.c1 : known && env.col >= q.c1 && env.col <= q.c2 ? env.col : -1;
    if (r < 0 || c < 0) fail('#VALUE!');
    return v[(r - q.r1) * (q.c2 - q.c1 + 1) + c - q.c1];
  }
  /** The rows of a range or array, each an array of values. */
  const grid = v => { const [rows, cols] = dims(v); return Array.from({length: rows}, (_, r) => v.slice(r * cols, r * cols + cols)); };
  /** A part of a range (rows r1..r2 and columns c1..c2 of it, counted from 0) as a range. */
  function subRange(v, r1, c1, r2, c2) {
    const [, cols] = dims(v), out = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(v[r * cols + c]);
    if (v.range) out.range = {r1: v.range.r1 + r1, c1: v.range.c1 + c1, r2: v.range.r1 + r2, c2: v.range.c1 + c2};
    else arrayOf(out, r2 - r1 + 1, c2 - c1 + 1);
    return out;
  }
  // Numbers are equal when they are the same to 15 significant digits (compare): 1+22*2^-52=1 is TRUE and 1+23*2^-52=1
  // FALSE, 1-10*2^-53=1 FALSE, as measured.
  const sameNumber = (a, b) => a === b || to15(a) === to15(b);
  const order = compare;
  const finite = x => Number.isFinite(x) ? x : fail('#NUM!');
  function binary(op, a, b) {
    switch (op) {
      case '&': return (toText(a) + toText(b)).slice(0, MAX_TEXT);
      case '=': return order(a, b) === 0;
      case '<>': return order(a, b) !== 0;
      case '<': return order(a, b) < 0;
      case '>': return order(a, b) > 0;
      case '<=': return order(a, b) <= 0;
      case '>=': return order(a, b) >= 0;
      case '+': return finite(toNumber(a) + toNumber(b));
      case '-': return finite(toNumber(a) - toNumber(b));
      case '*': return finite(toNumber(a) * toNumber(b));
      case '/': { const x = toNumber(a), y = toNumber(b); if (y === 0) fail('#DIV/0!'); return finite(x / y); }
      case '^': return power(toNumber(a), toNumber(b));
    }
    return fail('#VALUE!');
  }
  /** x^y. A negative x has a root only when 1/y is an odd whole number: (-8)^(1/3) is -2, (-8)^(2/3) and (-4)^0.5 are
      #NUM!. A fraction y goes through logarithms, as Excel's own power seems to (8^(1/3) is 1.9999999999999998 there):
      this gives Excel's value to the last binary digit in 209 of 238 measured powers (Math.pow in 129), and the rest
      differ only there, below the 15 digits Excel shows. */
  function power(x, y) {
    if (x === 0 && y === 0) fail('#NUM!');
    if (x === 0 && y < 0) fail('#DIV/0!');
    if (Number.isInteger(y)) return finite(x ** y);
    if (x < 0) { const k = 1 / y; if (!Number.isInteger(k) || k % 2 === 0) fail('#NUM!'); }
    const a = Math.abs(x), p = y === 0.5 ? Math.sqrt(a) : y < 0 ? 1 / Math.exp(-y * Math.log(a)) : Math.exp(y * Math.log(a));
    return finite(x < 0 ? -p : p);
  }
  /** f over one value, or over every value of an array in array mode (an error in a value stays in its place). */
  function lift1(x, env, arr, f) {
    if (arr && Array.isArray(x)) { const [rows, cols] = dims(x); return arrayOf(x.map(v => { if (isErr(v)) return v; try { return f(v); } catch (e) { return errorOf(e); } }), rows, cols); }
    x = single(x, env); if (isErr(x)) fail(x);
    return f(x);
  }
  /** f over two values, or element by element over arrays in array mode (a single row or column is repeated; elements
      beyond a smaller array are #N/A, as in Excel). */
  function lift2(a, b, env, arr, f) {
    if (arr && (Array.isArray(a) || Array.isArray(b))) {
      const da = Array.isArray(a) ? dims(a) : [1, 1], db = Array.isArray(b) ? dims(b) : [1, 1], rows = Math.max(da[0], db[0]), cols = Math.max(da[1], db[1]);
      const at = (v, [vr, vc], r, c) => !Array.isArray(v) ? v : (vr === 1 || r < vr) && (vc === 1 || c < vc) ? v[(vr === 1 ? 0 : r) * vc + (vc === 1 ? 0 : c)] : '#N/A';
      const out = [];
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
        const x = at(a, da, r, c), y = at(b, db, r, c);
        if (isErr(x)) out.push(x); else if (isErr(y)) out.push(y); else { try { out.push(f(x, y)); } catch (e) { out.push(errorOf(e)); } }
      }
      return arrayOf(out, rows, cols);
    }
    a = single(a, env); if (isErr(a)) fail(a);
    b = single(b, env); if (isErr(b)) fail(b);
    return f(a, b);
  }
  function run(n, env, arr) {
    switch (n.op) {
      case 'literal': return n.v;
      case 'missing': return BLANK;
      case 'ref': return env.ref(n.a, n.b);
      case 'unary': {
        const x = run(n.x, env, arr);
        return n.sign === '+' ? x : lift1(x, env, arr, v => { const r = -toNumber(v); return r === 0 ? 0 : r; }); // =+"5" stays the text 5
      }
      case 'percent': return lift1(run(n.x, env, arr), env, arr, v => toNumber(v) / 100);
      case 'binary': {
        // The left operand is looked at first: its error is the result even when the right one has an error too.
        let a = run(n.l, env, arr);
        if (!arr || !Array.isArray(a)) { a = single(a, env); if (isErr(a)) fail(a); }
        return lift2(a, run(n.r, env, arr), env, arr, (x, y) => binary(n.sign, x, y));
      }
      case 'call': return call(n, env, arr);
    }
    return fail('#VALUE!');
  }

  /** What a function implementation gets: its arguments (evaluated when asked for, in order) and helpers. */
  class Call {
    constructor(nodes, env, arr, values = null) { Object.assign(this, {nodes, env, arr, values, n: nodes.length}); }
    /** Whether argument i is there and not left empty. */
    given(i) { return i < this.n && this.nodes[i].op !== 'missing'; }
    /** Argument i as it is computed: a range stays an array. */
    raw(i) { return this.values ? this.values[i] : run(this.nodes[i], this.env, this.arr); }
    /** Argument i as one value (implicit intersection); an error is returned, not thrown. */
    value(i) { try { return single(this.raw(i), this.env); } catch (e) { return errorOf(e); } }
    /** Argument i as one value; an error is thrown. */
    v(i) { const v = single(this.raw(i), this.env); if (isErr(v)) fail(v); return v; }
    num(i, d) { return i < this.n ? toNumber(this.v(i)) : d; }
    /** Argument i as a number, where TRUE and FALSE are #VALUE!, as in the functions that came from the Analysis ToolPak
        (EDATE, EOMONTH, MROUND, QUOTIENT, WEEKNUM, WORKDAY, NETWORKDAYS). */
    strict(i, d) { if (i >= this.n) return d; const v = this.v(i); if (typeof v === 'boolean') fail('#VALUE!'); return toNumber(v); }
    int(i, d) { return i < this.n ? Math.trunc(toNumber(this.v(i))) : d; }
    text(i, d) { return i < this.n ? toText(this.v(i)) : d; }
    bool(i, d) { return i < this.n ? toBool(this.v(i)) : d; }
    /** Argument i as a range or array (one value is an array of one). */
    range(i) { const v = this.raw(i); if (Array.isArray(v)) return v; if (isErr(v)) fail(v); const out = [v]; return arrayOf(out, 1, 1); }
    /** Argument i as the cells of a range, where an error is a cell's value like any other (COUNTIF, COUNTBLANK). */
    cells(i) { const v = this.raw(i); return Array.isArray(v) ? v : arrayOf([v], 1, 1); }
  }
  // kinds of the arguments of a function that array mode (inside SUMPRODUCT) computes value by value: v one value, r a
  // range (kept whole); the last kind repeats. Functions without kinds take arrays as they are.
  const FN = {};
  const def = (names, min, max, kinds, f) => { for (const name of names.split(' ')) FN[name] = {min, max, kinds, f, name}; };
  function call(n, env, arr) {
    const name = n.name, d = FN[name];
    // A function the page does not compute is #NAME?, whatever its arguments are, as Excel shows a name it does not know
    // (=FOO(1/0) is #NAME?).
    if (!d || !FUNCTIONS.includes(name)) return fail('#NAME?');
    if (n.args.length < d.min || n.args.length > d.max) fail('#VALUE!');
    if (arr && d.kinds) {
      const kind = i => d.kinds[Math.min(i, d.kinds.length - 1)];
      const values = n.args.map((node, i) => { try { return run(node, env, kind(i) === 'v'); } catch (e) { return errorOf(e); } });
      const lifted = values.map((v, i) => kind(i) === 'v' && Array.isArray(v));
      if (lifted.some(Boolean)) {
        const shapes = values.map((v, i) => lifted[i] ? dims(v) : [1, 1]), rows = Math.max(...shapes.map(s => s[0])), cols = Math.max(...shapes.map(s => s[1])), out = [];
        for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
          const args = values.map((v, i) => { if (!lifted[i]) return v; const [vr, vc] = shapes[i]; return (vr === 1 || r < vr) && (vc === 1 || c < vc) ? v[(vr === 1 ? 0 : r) * vc + (vc === 1 ? 0 : c)] : '#N/A'; });
          try { out.push(d.f(new Call(n.args, env, false, args))); } catch (e) { out.push(errorOf(e)); }
        }
        return arrayOf(out, rows, cols);
      }
      return d.f(new Call(n.args, env, false, values));
    }
    return d.f(new Call(n.args, env, arr));
  }
  /** The arguments from index i on, in the order Excel reads them: those given directly first, then references
      (=SUM(A1,1/0) is #DIV/0! though A1 holds #N/A). Each is {v, ref}: ref when it came from cells or is an array. */
  function collect(c, from = 0, to = c.n) {
    const out = [], later = [];
    for (let i = from; i < to; i++) {
      if (c.nodes[i].op === 'ref' && !c.values) { later.push(i); continue; }
      let v; try { v = c.raw(i); } catch (e) { v = errorOf(e); }
      out.push({v, ref: Array.isArray(v) || c.nodes[i].op === 'ref'});
    }
    for (const i of later) out.push({v: c.raw(i), ref: true});
    return out;
  }
  /** The numbers of the arguments, as SUM takes them: from cells only numbers; given directly, a number, TRUE or FALSE,
      an empty argument (0) or a text that reads as a number (another text is #VALUE!). The first error is thrown. */
  function numbersOf(items) {
    const out = [];
    for (const {v, ref} of items) {
      if (ref) { for (const x of [v].flat()) { if (isErr(x)) fail(x); if (typeof x === 'number') out.push(x); } }
      else { if (isErr(v)) fail(v); out.push(toNumber(v)); }
    }
    return out;
  }
  const numbers = (c, from = 0) => numbersOf(collect(c, from));
  /** TRUE and FALSE of the arguments, as AND takes them: numbers too; texts TRUE and FALSE only when given directly. */
  function logicalsOf(c) {
    const out = [];
    for (const {v, ref} of collect(c)) {
      for (const x of [v].flat()) {
        if (isErr(x)) fail(x);
        if (typeof x === 'boolean') out.push(x); else if (typeof x === 'number') out.push(x !== 0);
        else if (!ref && x === BLANK) out.push(false);
        else if (!ref && typeof x === 'string' && /^(?:true|false)$/i.test(x)) out.push(/^true$/i.test(x));
      }
    }
    if (!out.length) fail('#VALUE!');
    return out;
  }
  const sum = xs => { let s = 0; for (const x of xs) s += x; return finite(s); };
  const mean = xs => xs.length ? sum(xs) / xs.length : fail('#DIV/0!');
  /** The variance as Excel computes it, to the last binary digit: of a sample from the sum and the sum of squares
      (VAR(1,2,1) is 0.3333333333333335 there), of a population from the squares of the differences from the mean
      (VAR.P(1,2,1) is 0.2222222222222222). */
  function variance(xs, sample) {
    if (xs.length < (sample ? 2 : 1)) fail('#DIV/0!');
    let s = 0, q = 0;
    if (sample) { for (const x of xs) { s += x; q += x * x; } return Math.max(0, finite((q - s * s / xs.length) / (xs.length - 1))); }
    const m = sum(xs) / xs.length; for (const x of xs) q += (x - m) ** 2;
    return finite(q / xs.length);
  }
  /** x rounded at k places after the point (k < 0: tens, hundreds ...) on its 15 significant digits: half (away from
      zero), up (away from zero) or down (toward zero), as ROUND, ROUNDUP and ROUNDDOWN do. */
  function roundAt(x, k, mode) {
    if (!x) return 0;
    const {d, p} = digits15(Math.abs(x)), shift = d.length - p - k;
    let v;
    if (shift <= 0) v = Number('0.' + d + 'e' + p);
    else {
      const unit = 10n ** BigInt(shift), n = BigInt(d);
      let q = n / unit; const rest = n % unit;
      if (mode === 'half' && rest * 2n >= unit) q += 1n;
      v = Number(q.toString() + 'e' + (-k));
      // ROUNDUP adds the unit to the number cut down, as Excel does: ROUNDUP(1.15,1) is 1.1+0.1, 1.2000000000000002.
      if (mode === 'up' && rest > 0n) v += Number('1e' + (-k));
    }
    return x < 0 ? -v : v;
  }
  // CEILING and FLOOR: the quotient to 15 digits, made whole, times the step (CEILING(0.3,0.1) is 0.30000000000000004).
  const floorTo = (x, s) => Math.floor(to15(x / s)) * s || 0;
  const ceilTo = (x, s) => Math.ceil(to15(x / s)) * s || 0;

  // ---- Logical and information functions
  def('IF', 2, 3, null, c => toBool(c.v(0)) ? c.raw(1) : c.n > 2 ? c.raw(2) : false);
  def('IFS', 2, 254, null, c => {
    if (c.n % 2) fail('#VALUE!');
    for (let i = 0; i < c.n; i += 2) if (toBool(c.v(i))) return c.raw(i + 1);
    return fail('#N/A');
  });
  def('IFERROR', 2, 2, null, c => { const v = c.value(0); return isErr(v) ? c.raw(1) : c.raw(0); });
  def('IFNA', 2, 2, null, c => { const v = c.value(0); return v === '#N/A' ? c.raw(1) : c.raw(0); });
  def('AND', 1, 255, null, c => logicalsOf(c).every(Boolean));
  def('OR', 1, 255, null, c => logicalsOf(c).some(Boolean));
  def('XOR', 1, 254, null, c => logicalsOf(c).filter(Boolean).length % 2 === 1);
  def('NOT', 1, 1, 'v', c => !toBool(c.v(0)));
  def('TRUE', 0, 0, null, () => true);
  def('FALSE', 0, 0, null, () => false);
  def('NA', 0, 0, null, () => '#N/A');
  // SWITCH matches as = does, but an empty cell only as 0 (not as "").
  def('SWITCH', 3, 254, null, c => {
    let x = c.v(0); if (x === BLANK) x = 0;
    for (let i = 1; i + 1 < c.n; i += 2) { const v = c.v(i); if (rank(x) === rank(v === BLANK ? 0 : v) && order(x, v === BLANK ? 0 : v) === 0) return c.raw(i + 1); }
    return c.n % 2 === 0 ? c.raw(c.n - 1) : fail('#N/A');
  });
  def('CHOOSE', 2, 255, null, c => { const i = c.int(0); if (i < 1 || i >= c.n) fail('#VALUE!'); return c.raw(i); });
  def('ISBLANK', 1, 1, 'v', c => c.value(0) === BLANK);
  def('ISNUMBER', 1, 1, 'v', c => typeof c.value(0) === 'number');
  def('ISTEXT', 1, 1, 'v', c => { const v = c.value(0); return typeof v === 'string' && !isErr(v); });
  def('ISNONTEXT', 1, 1, 'v', c => { const v = c.value(0); return !(typeof v === 'string' && !isErr(v)); });
  def('ISLOGICAL', 1, 1, 'v', c => typeof c.value(0) === 'boolean');
  def('ISERROR', 1, 1, 'v', c => isErr(c.value(0)));
  def('ISERR', 1, 1, 'v', c => { const v = c.value(0); return isErr(v) && v !== '#N/A'; });
  def('ISNA', 1, 1, 'v', c => c.value(0) === '#N/A');
  const parity = c => { const v = c.v(0); if (typeof v === 'boolean') fail('#VALUE!'); return Math.abs(Math.trunc(toNumber(v))) % 2; };
  def('ISEVEN', 1, 1, 'v', c => parity(c) === 0);
  def('ISODD', 1, 1, 'v', c => parity(c) === 1);
  def('N', 1, 1, 'v', c => { const v = c.v(0); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; });
  def('T', 1, 1, 'v', c => { const v = c.v(0); return typeof v === 'string' ? v : ''; });

  // ---- Math
  def('SUM', 1, 255, null, c => sum(numbers(c)));
  // PRODUCT leaves out an argument left empty (=PRODUCT(5,) is 5), where SUM counts it as 0.
  def('PRODUCT', 1, 255, null, c => { const xs = numbersOf(collect(c).filter(({v, ref}) => ref || v !== BLANK)); if (!xs.length) return 0; let p = 1; for (const x of xs) p *= x; return finite(p); });
  def('ROUND', 2, 2, 'v', c => roundAt(c.num(0), c.int(1), 'half'));
  def('ROUNDUP', 2, 2, 'v', c => roundAt(c.num(0), c.int(1), 'up'));
  def('ROUNDDOWN', 2, 2, 'v', c => roundAt(c.num(0), c.int(1), 'down'));
  def('TRUNC', 1, 2, 'v', c => roundAt(c.num(0), c.int(1, 0), 'down'));
  def('INT', 1, 1, 'v', c => { const x = c.num(0); return x >= 0 ? roundAt(x, 0, 'down') : roundAt(x, 0, 'up'); });
  def('ABS', 1, 1, 'v', c => Math.abs(c.num(0)));
  // MOD is the exact remainder of the two numbers as they are held (JavaScript's %, as measured: MOD(6,0.1) is
  // 0.09999999999999967), with the sign of the divisor; a quotient beyond 2^50/1000 (about 1.1259E+12) is #NUM!.
  def('MOD', 2, 2, 'v', c => {
    const a = c.num(0), b = c.num(1);
    if (b === 0) fail('#DIV/0!');
    if (Math.abs(a / b) > 2 ** 50 / 1000) fail('#NUM!');
    let r = a % b; if (r !== 0 && (r < 0) !== (b < 0)) r += b;
    return r || 0;
  });
  def('QUOTIENT', 2, 2, 'v', c => { const a = c.strict(0), b = c.strict(1); if (b === 0) fail('#DIV/0!'); return Math.trunc(finite(a / b)) || 0; });
  def('SIGN', 1, 1, 'v', c => Math.sign(c.num(0)) || 0);
  def('SQRT', 1, 1, 'v', c => { const x = c.num(0); if (x < 0) fail('#NUM!'); return Math.sqrt(x); });
  def('POWER', 2, 2, 'v', c => power(c.num(0), c.num(1)));
  def('EXP', 1, 1, 'v', c => finite(Math.exp(c.num(0))));
  def('LN', 1, 1, 'v', c => { const x = c.num(0); if (x <= 0) fail('#NUM!'); return Math.log(x); });
  def('LOG10', 1, 1, 'v', c => { const x = c.num(0); if (x <= 0) fail('#NUM!'); return Math.log10(x); });
  def('LOG', 1, 2, 'v', c => {
    const x = c.num(0), b = c.n > 1 ? c.num(1) : 10;
    if (x <= 0 || b <= 0) fail('#NUM!'); if (b === 1) fail('#DIV/0!');
    // With a base given, ln(x)/ln(base), as Excel computes it: LOG(1000,10) is 2.9999999999999996 there.
    return c.n > 1 ? Math.log(x) / Math.log(b) : Math.log10(x);
  });
  def('PI', 0, 0, null, () => Math.PI);
  def('CEILING', 2, 2, 'v', c => { const x = c.num(0), s = c.num(1); if (s === 0) return 0; if (x > 0 && s < 0) fail('#NUM!'); return ceilTo(x, s); });
  def('FLOOR', 2, 2, 'v', c => { const x = c.num(0), s = c.num(1); if (s === 0) return x === 0 ? 0 : fail('#DIV/0!'); if (x > 0 && s < 0) fail('#NUM!'); return floorTo(x, s); });
  def('CEILING.MATH', 1, 3, 'v', c => {
    const x = c.num(0), s = Math.abs(c.given(1) ? c.num(1) : 1), away = c.n > 2 && c.num(2) !== 0;
    if (s === 0) return 0;
    return x < 0 && away ? -ceilTo(-x, s) : ceilTo(x, s);
  });
  def('FLOOR.MATH', 1, 3, 'v', c => {
    const x = c.num(0), s = Math.abs(c.given(1) ? c.num(1) : 1), toward = c.n > 2 && c.num(2) !== 0;
    if (s === 0) return 0;
    return x < 0 && toward ? -floorTo(-x, s) : floorTo(x, s);
  });
  // MROUND goes up when the remainder (as MOD has it), to 15 digits, is at least half the multiple, as measured:
  // MROUND(0.15,0.1) is 0.2 (the remainder 0.04999999999999999 is 0.05 to 15 digits), MROUND(6.05,0.1) is 6 (the
  // remainder is 0.04999999999999949).
  def('MROUND', 2, 2, 'v', c => {
    const x = c.strict(0), m = c.strict(1);
    if (m === 0) return 0; if (x * m < 0) fail('#NUM!');
    const a = Math.abs(x), b = Math.abs(m), r = a % b;
    return Math.sign(x) * (a - r + (to15(r) >= to15(b / 2) ? b : 0)) || 0;
  });
  def('SUMPRODUCT', 1, 255, null, c => {
    const arrays = c.nodes.map((node, i) => { const v = run(node, c.env, true); return Array.isArray(v) ? v : arrayOf([v], 1, 1); });
    const [rows, cols] = dims(arrays[0]);
    if (arrays.some(a => { const [r, k] = dims(a); return r !== rows || k !== cols; })) fail('#VALUE!');
    let s = 0;
    for (let i = 0; i < rows * cols; i++) {
      let p = 1;
      for (const a of arrays) { const x = a[i]; if (isErr(x)) fail(x); p *= typeof x === 'number' ? x : 0; }
      s += p;
    }
    return finite(s);
  });

  // ---- Statistics
  def('AVERAGE', 1, 255, null, c => mean(numbers(c)));
  def('MIN', 1, 255, null, c => { const xs = numbers(c); return xs.length ? Math.min(...xs) : 0; });
  def('MAX', 1, 255, null, c => { const xs = numbers(c); return xs.length ? Math.max(...xs) : 0; });
  def('COUNT', 1, 255, null, c => {
    let n = 0;
    for (const {v, ref} of collect(c)) for (const x of [v].flat()) {
      if (typeof x === 'number') n++;
      else if (!ref && (typeof x === 'boolean' || x === BLANK || typeof x === 'string' && !isErr(x) && textNumber(x) !== null)) n++;
    }
    return n;
  });
  def('COUNTA', 1, 255, null, c => { let n = 0; for (const {v, ref} of collect(c)) for (const x of [v].flat()) if (!ref || x !== BLANK) n++; return n; });
  def('COUNTBLANK', 1, 1, null, c => c.cells(0).filter(x => x === BLANK || x === '').length);
  def('MEDIAN', 1, 255, null, c => {
    const xs = numbers(c).sort((a, b) => a - b); if (!xs.length) fail('#NUM!');
    const h = xs.length >> 1; return xs.length % 2 ? xs[h] : (xs[h - 1] + xs[h]) / 2;
  });
  // MODE counts numbers only, also of what is given directly (=MODE("1","1",2,2) is 2; TRUE and texts are left out);
  // an argument left empty is #VALUE!.
  def('MODE MODE.SNGL', 1, 255, null, c => {
    const xs = [];
    for (const {v, ref} of collect(c)) for (const x of [v].flat()) {
      if (isErr(x)) fail(x);
      if (typeof x === 'number') xs.push(x); else if (!ref && x === BLANK) fail('#VALUE!');
    }
    const counts = new Map(); let best = null, most = 1;
    for (const x of xs) { const n = (counts.get(x) || 0) + 1; counts.set(x, n); if (n > most || n === most && n > 1 && best !== null && xs.indexOf(x) < xs.indexOf(best)) { most = n; best = x; } }
    return best === null ? fail('#N/A') : best;
  });
  // k of LARGE and SMALL must lie between 1 and the count of numbers; between two whole numbers, LARGE takes the next
  // (LARGE(x,1.5) is the second largest) and SMALL the one before (SMALL(x,1.5) is the smallest), as measured.
  const kth = (c, large) => {
    const xs = numbersOf(collect(c, 0, 1)).sort((a, b) => large ? b - a : a - b), k = c.num(1);
    if (k < 1 || k > xs.length) fail('#NUM!');
    return xs[(large ? Math.ceil(k) : Math.floor(k)) - 1];
  };
  def('LARGE', 2, 2, 'rv', c => kth(c, true));
  def('SMALL', 2, 2, 'rv', c => kth(c, false));
  // RANK: an error among the numbers is the result; the order is a number or TRUE / FALSE (a text, even "1", is #VALUE!).
  def('RANK RANK.EQ', 2, 3, 'vrv', c => {
    const x = c.num(0);
    let ascending = false;
    if (c.n > 2) { const o = c.v(2); if (typeof o === 'string') fail('#VALUE!'); ascending = toNumber(o) !== 0; }
    const xs = [];
    for (const v of c.range(1)) { if (isErr(v)) fail(v); if (typeof v === 'number') xs.push(v); }
    if (!xs.some(v => sameNumber(v, x))) fail('#N/A');
    return 1 + xs.filter(v => !sameNumber(v, x) && (ascending ? v < x : v > x)).length;
  });
  def('STDEV STDEV.S', 1, 255, null, c => Math.sqrt(variance(numbers(c), true)));
  def('STDEVP STDEV.P', 1, 255, null, c => Math.sqrt(variance(numbers(c), false)));
  def('VAR VAR.S', 1, 255, null, c => variance(numbers(c), true));
  def('VARP VAR.P', 1, 255, null, c => variance(numbers(c), false));
  // SUBTOTAL leaves out the cells whose formulas hold SUBTOTAL themselves, as Excel does (the page hides no rows, so
  // 101-111 are the same as 1-11).
  const SUBTOTALS = {1: xs => mean(xs), 2: (xs, all) => xs.length, 3: (xs, all) => all.filter(x => x !== BLANK).length, 4: xs => xs.length ? Math.max(...xs) : 0,
    5: xs => xs.length ? Math.min(...xs) : 0, 6: xs => { let p = 1; for (const x of xs) p *= x; return xs.length ? p : 0; }, 7: xs => Math.sqrt(variance(xs, true)),
    8: xs => Math.sqrt(variance(xs, false)), 9: xs => sum(xs), 10: xs => variance(xs, true), 11: xs => variance(xs, false)};
  def('SUBTOTAL', 2, 255, null, c => {
    const k = c.int(0), f = SUBTOTALS[k > 100 ? k - 100 : k];
    if (!f) fail('#VALUE!');
    const all = [], xs = [];
    for (let i = 1; i < c.n; i++) {
      const v = c.raw(i), q = Array.isArray(v) ? v.range : c.nodes[i].op === 'ref' ? (() => { const [r, cc] = coordOf(c.nodes[i].a); return {r1: r, c1: cc, r2: r, c2: cc}; })() : null;
      if (!q) fail('#VALUE!');
      const values = [v].flat(), cols = q.c2 - q.c1 + 1;
      values.forEach((x, j) => {
        const f2 = c.env.formulaAt(q.r1 + Math.floor(j / cols), q.c1 + j % cols);
        if (f2 && /\b(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(f2)) return;
        if (k % 100 === 2 || k % 100 === 3) { if (k % 100 === 3) all.push(x); else if (typeof x === 'number') xs.push(x); return; }
        if (isErr(x)) fail(x);
        if (typeof x === 'number') xs.push(x);
      });
    }
    return f(xs, all);
  });

  // ---- Criteria (COUNTIF, SUMIF and the like)
  /** Which cells meet every criterion. pairs are [range, index of its criterion], where a range is an argument index or
      the cells themselves; all ranges must be of one size (size, or that of the first range), else #VALUE!. A criterion
      that is an error is no error here: it finds the cells holding that error. */
  function criteriaHits(c, pairs, size) {
    const tests = pairs.map(([r, k]) => [Array.isArray(r) ? r : c.cells(r), criterion(c.value(k))]);
    const [rows, cols] = size ?? dims(tests[0][0]);
    if (tests.some(([r]) => { const [a, b] = dims(r); return a !== rows || b !== cols; })) fail('#VALUE!');
    const hit = []; for (let j = 0; j < rows * cols; j++) hit.push(tests.every(([r, test]) => test(r[j])));
    return hit;
  }
  const pairsFrom = (c, first) => { const out = []; for (let i = first; i + 1 < c.n; i += 2) out.push([i, i + 1]); return out; };
  /** The range of SUMIF and AVERAGEIF to add up: the given one, made as large as the criteria range from its first cell
      (=SUMIF(A1:A14,"東京",B1:B1) adds B1:B14). */
  function sumRange(c, i, like) {
    const [rows, cols] = dims(like), node = c.nodes[i], v = c.raw(i);
    let r1, c1;
    if (Array.isArray(v) && v.range) ({r1, c1} = v.range);
    else if (node.op === 'ref' && !node.b) [r1, c1] = coordOf(node.a);
    else return fail('#VALUE!');
    if (Array.isArray(v) && dims(v)[0] === rows && dims(v)[1] === cols) return v;
    const out = c.env.ref(cellName(r1, c1), cellName(r1 + rows - 1, c1 + cols - 1));
    return Array.isArray(out) ? out : fail('#REF!');
  }
  const pickNumbers = (values, hit) => { const xs = []; values.forEach((x, j) => { if (!hit[j]) return; if (isErr(x)) fail(x); if (typeof x === 'number') xs.push(x); }); return xs; };
  def('COUNTIF', 2, 2, 'rv', c => criteriaHits(c, [[0, 1]]).filter(Boolean).length);
  def('COUNTIFS', 2, 254, 'rv', c => { if (c.n % 2) fail('#VALUE!'); return criteriaHits(c, pairsFrom(c, 0)).filter(Boolean).length; });
  def('SUMIF', 2, 3, 'rvr', c => { const r = c.cells(0), hit = criteriaHits(c, [[r, 1]]); return sum(pickNumbers(c.n > 2 ? sumRange(c, 2, r) : r, hit)); });
  def('AVERAGEIF', 2, 3, 'rvr', c => { const r = c.cells(0), hit = criteriaHits(c, [[r, 1]]); return mean(pickNumbers(c.n > 2 ? sumRange(c, 2, r) : r, hit)); });
  const ifs = (c, f) => { if (c.n % 2 === 0) fail('#VALUE!'); const values = c.cells(0), hit = criteriaHits(c, pairsFrom(c, 1), dims(values)); return f(pickNumbers(values, hit)); };
  def('SUMIFS', 3, 255, 'rrv', c => ifs(c, sum));
  def('AVERAGEIFS', 3, 255, 'rrv', c => ifs(c, mean));
  def('MAXIFS', 3, 255, 'rrv', c => ifs(c, xs => xs.length ? Math.max(...xs) : 0));
  def('MINIFS', 3, 255, 'rrv', c => ifs(c, xs => xs.length ? Math.min(...xs) : 0));

  // ---- Lookup and reference
  /** Whether a cell value equals a lookup value exactly (VLOOKUP FALSE, MATCH 0, XLOOKUP): of the same kind, texts
      without regard to case and, when wild, with * ? ~. */
  function same(x, v, wild) {
    if (x === BLANK) return false;
    if (typeof x !== typeof v || isErr(x) !== isErr(v)) return false;
    if (typeof v === 'string') return wild ? matchText(v, x) : sameText(v, x);
    return typeof v === 'number' ? sameNumber(x, v) : x === v;
  }
  // How a cell value stands to the lookup value in the searches below: an empty cell and an error count as larger.
  const against = (x, v) => x === BLANK || isErr(x) ? 1 : order(x, v);
  const equal = (x, v) => against(x, v) === 0 && rank(x) === rank(v);
  /** Excel's search of values sorted ascending (MATCH 1, VLOOKUP and HLOOKUP TRUE, LOOKUP), as measured on unsorted and
      repeated values: a binary search that, on meeting an equal value, goes on over the equal values right after it and
      gives the last of them (MATCH(2,{1,2,2,2,3}) is 4); else the index of the last value before the lookup value. -1
      when there is none. */
  function bsearch(values, v) {
    let lo = 0, hi = values.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (equal(values[mid], v)) { let i = mid; while (i + 1 < values.length && equal(values[i + 1], v)) i++; return i; }
      if (against(values[mid], v) < 0) lo = mid + 1; else hi = mid - 1;
    }
    return hi;
  }
  /** MATCH with a negative type, for values sorted descending, as measured: a search from the first value on that stops
      at an equal value, or before the first value smaller than the lookup value. */
  function descendingSearch(values, v) {
    let last = -1;
    for (let i = 0; i < values.length; i++) {
      const x = values[i];
      if (x === BLANK || isErr(x) || rank(x) !== rank(v)) continue;
      const o = order(x, v);
      if (o === 0) return i;
      if (o < 0) return last;
      last = i;
    }
    return last;
  }
  /** The binary search of XLOOKUP and XMATCH (search mode 2, or -2 for descending values), as measured: it stops at an
      equal value ({found}); else lo and hi are the indexes the lookup value lies between. */
  function xbinary(values, v, descending) {
    let lo = 0, hi = values.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (equal(values[mid], v)) return {found: mid};
      if (against(values[mid], v) * (descending ? -1 : 1) < 0) lo = mid + 1; else hi = mid - 1;
    }
    return {lo, hi};
  }
  const lookupValue = c => { const v = c.v(0); return v === BLANK ? 0 : v; };
  function tableLookup(c, horizontal) {
    const v = lookupValue(c), table = c.range(1), n = c.int(2), approx = c.n > 3 ? c.bool(3) : true;
    if (n < 1) fail('#VALUE!');
    const rows = horizontal ? grid(table)[0] ? grid(table) : [] : grid(table), lines = horizontal ? rows[0] : rows.map(r => r[0]);
    const [r, k] = dims(table);
    if (n > (horizontal ? r : k)) fail('#REF!');
    const i = approx ? bsearch(lines, v) : lines.findIndex(x => same(x, v, typeof v === 'string'));
    if (i < 0) fail('#N/A');
    return horizontal ? rows[n - 1][i] : rows[i][n - 1];
  }
  def('VLOOKUP', 3, 4, 'vrvv', c => tableLookup(c, false));
  def('HLOOKUP', 3, 4, 'vrvv', c => tableLookup(c, true));
  const line = v => { const [r, k] = dims(v); if (r !== 1 && k !== 1) return null; return [...v]; };
  // The type of MATCH counts by its sign: -0.5 is -1 and 0.5 is 1.
  def('MATCH', 2, 3, 'vrv', c => {
    const v = lookupValue(c), values = line(c.range(1)), type = c.n > 2 ? c.num(2) : 1;
    if (!values) fail('#N/A');
    const i = type === 0 ? values.findIndex(x => same(x, v, typeof v === 'string')) : type > 0 ? bsearch(values, v) : descendingSearch(values, v);
    if (i < 0) fail('#N/A');
    return i + 1;
  });
  def('INDEX', 2, 4, 'rvv', c => {
    const a = c.range(0), [rows, cols] = dims(a);
    let r = c.given(1) ? c.int(1) : 0, k = c.n > 2 && c.given(2) ? c.int(2) : 0;
    if (c.n > 3 && c.int(3) !== 1) fail('#REF!');
    if (r < 0 || k < 0) fail('#VALUE!');
    if (c.n <= 2 && rows === 1 && cols > 1) { k = r; r = 1; }
    else if (c.n <= 2 && cols > 1 && rows > 1 && r) fail('#REF!');
    if (r > rows || k > cols) fail('#REF!');
    if (r && k) return a[(r - 1) * cols + k - 1];
    return subRange(a, r ? r - 1 : 0, k ? k - 1 : 0, r ? r - 1 : rows - 1, k ? k - 1 : cols - 1);
  });
  /** XLOOKUP and XMATCH: the index of the match in a row or column, or -1. */
  function xfind(v, values, mode, search) {
    if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search)) fail('#VALUE!');
    if (search === 2 || search === -2) {
      const r = xbinary(values, v, search === -2);
      if (r.found !== undefined) return r.found;
      // Not found: the next smaller value lies before the place in ascending values (hi) and after it in descending ones.
      const j = mode === -1 ? (search === 2 ? r.hi : r.lo) : mode === 1 ? (search === 2 ? r.lo : r.hi) : -1;
      return j >= 0 && j < values.length ? j : -1;
    }
    const order1 = search === 1 ? values.map((_, i) => i) : values.map((_, i) => values.length - 1 - i);
    for (const i of order1) if (same(values[i], v, mode === 2)) return i;
    if (mode === -1 || mode === 1) {
      let best = -1;
      for (const i of order1) {
        const x = values[i]; if (x === BLANK || isErr(x) || rank(x) !== rank(v)) continue;
        const o = order(x, v);
        if (mode === -1 ? o < 0 && (best < 0 || order(x, values[best]) > 0) : o > 0 && (best < 0 || order(x, values[best]) < 0)) best = i;
      }
      return best;
    }
    return -1;
  }
  def('XLOOKUP', 3, 6, 'vrrvvv', c => {
    const v = c.v(0), look = c.range(1), back = c.range(2), values = line(look);
    if (!values) fail('#VALUE!');
    const [lr, lc] = dims(look), [br, bc] = dims(back), vertical = lc === 1 && lr > 1 || lr === 1 && lc === 1 && bc === 1;
    if (vertical ? br !== lr : bc !== lc) fail('#VALUE!');
    const i = xfind(v === BLANK ? 0 : v, values, c.n > 4 && c.given(4) ? c.int(4) : 0, c.n > 5 && c.given(5) ? c.int(5) : 1);
    if (i < 0) return c.n > 3 && c.given(3) ? c.raw(3) : fail('#N/A');
    return vertical ? subRange(back, i, 0, i, bc - 1).length === 1 ? back[i * bc] : subRange(back, i, 0, i, bc - 1) : br === 1 ? back[i] : subRange(back, 0, i, br - 1, i);
  });
  def('XMATCH', 2, 4, 'vrvv', c => {
    const v = c.v(0), values = line(c.range(1));
    if (!values) fail('#VALUE!');
    const i = xfind(v === BLANK ? 0 : v, values, c.n > 2 && c.given(2) ? c.int(2) : 0, c.n > 3 && c.given(3) ? c.int(3) : 1);
    return i < 0 ? fail('#N/A') : i + 1;
  });
  def('LOOKUP', 2, 3, 'vrr', c => {
    const v = lookupValue(c), a = c.range(1), [rows, cols] = dims(a);
    let keys, results;
    if (c.n > 2) { keys = line(a); results = line(c.range(2)); if (!keys || !results) fail('#N/A'); }
    else if (rows >= cols) { const g = grid(a); keys = g.map(r => r[0]); results = g.map(r => r[cols - 1]); }
    else { const g = grid(a); keys = g[0]; results = g[rows - 1]; }
    const i = bsearch(keys, v);
    if (i < 0) fail('#N/A');
    return i < results.length ? results[i] : fail('#N/A');
  });
  /** The first row or column of the range of argument i (a reference), or of the formula's own cell. */
  function place(c, rows) {
    if (c.n === 0) return (rows ? c.env.row : c.env.col) + 1;
    const node = c.nodes[0];
    if (node.op === 'ref') { const [r, k] = coordOf(node.a); return (rows ? r : k) + 1; }
    const v = c.raw(0); if (Array.isArray(v) && v.range) return (rows ? v.range.r1 : v.range.c1) + 1;
    return fail('#VALUE!');
  }
  def('ROW', 0, 1, null, c => place(c, true));
  def('COLUMN', 0, 1, null, c => place(c, false));
  def('ROWS', 1, 1, null, c => dims(c.range(0))[0]);
  def('COLUMNS', 1, 1, null, c => dims(c.range(0))[1]);

  // ---- Text
  const cut = s => s.length > MAX_TEXT ? fail('#VALUE!') : s;
  // LEN, MID, REPLACE, FIND and SEARCH count UTF-16 units, as Excel does: 𩸽 is two (LEN("𩸽") is 2, and MID can cut it in
  // half). LEFT and RIGHT count characters: LEFT("𩸽の",1) is 𩸽.
  def('LEN', 1, 1, 'v', c => c.text(0).length);
  def('LEFT', 1, 2, 'v', c => { const s = c.text(0), n = c.int(1, 1); if (n < 0) fail('#VALUE!'); return Array.from(s).slice(0, n).join(''); });
  def('RIGHT', 1, 2, 'v', c => { const s = c.text(0), n = c.int(1, 1); if (n < 0) fail('#VALUE!'); return n ? Array.from(s).slice(-n).join('') : ''; });
  def('MID', 3, 3, 'v', c => { const s = c.text(0), start = c.int(1), n = c.int(2); if (start < 1 || n < 0) fail('#VALUE!'); return s.slice(start - 1, start - 1 + n); });
  // FIND may start right after the last character (FIND("","abc",4) is 4); SEARCH may not (SEARCH("","abc",4) is #VALUE!).
  function find(c, insensitive) {
    const what = c.text(0), within = c.text(1), start = c.int(2, 1);
    if (start < 1 || start > within.length + (insensitive ? 0 : 1)) fail('#VALUE!');
    if (what === '') return start;
    if (!insensitive) { const i = within.indexOf(what, start - 1); return i < 0 ? fail('#VALUE!') : i + 1; }
    const re = wildcard(what);
    if (!re) { const i = within.toLowerCase().indexOf(what.toLowerCase(), start - 1); return i < 0 ? fail('#VALUE!') : i + 1; }
    const search = new RegExp(re.source.slice(1, -1), 'iy');
    for (let i = start - 1; i < within.length; i++) { search.lastIndex = i; if (search.test(within)) return i + 1; }
    return fail('#VALUE!');
  }
  def('FIND', 2, 3, 'v', c => find(c, false));
  def('SEARCH', 2, 3, 'v', c => find(c, true));
  def('SUBSTITUTE', 3, 4, 'v', c => {
    const s = c.text(0), old = c.text(1), rep = c.text(2), k = c.n > 3 ? c.int(3) : 0;
    if (c.n > 3 && k < 1) fail('#VALUE!');
    if (!old) return s;
    if (!k) return cut(s.split(old).join(rep));
    // The occurrences are counted overlapping: the second "aa" of "aaa" starts at its second letter (the result is "ab").
    let at = -1; for (let i = 0; i < k; i++) { at = s.indexOf(old, at + 1); if (at < 0) return s; }
    return cut(s.slice(0, at) + rep + s.slice(at + old.length));
  });
  def('REPLACE', 4, 4, 'v', c => { const s = c.text(0), start = c.int(1), n = c.int(2), rep = c.text(3); if (start < 1 || n < 0) fail('#VALUE!'); return cut(s.slice(0, start - 1) + rep + s.slice(start - 1 + n)); });
  // TRIM takes away spaces (also full-width ones) at both ends, and leaves one of each run inside.
  def('TRIM', 1, 1, 'v', c => c.text(0).replace(/^[ 　]+|[ 　]+$/g, '').replace(/ {2,}/g, ' ').replace(/　{2,}/g, '　'));
  def('CLEAN', 1, 1, 'v', c => c.text(0).replace(/[\x00-\x1f\x81\x8d\x8f\x90\x9d]/g, ''));
  def('UPPER', 1, 1, 'v', c => c.text(0).toUpperCase());
  def('LOWER', 1, 1, 'v', c => c.text(0).toLowerCase());
  // PROPER looks at each UTF-16 unit: half of 𩸽 is no letter, so PROPER("𩸽abc") is 𩸽Abc (as 日本abc stays as it is).
  def('PROPER', 1, 1, 'v', c => { let prev = false; return c.text(0).toLowerCase().split('').map(ch => { const letter = /\p{L}/u.test(ch), out = letter && !prev ? ch.toUpperCase() : ch; prev = letter; return out; }).join(''); });
  def('REPT', 2, 2, 'v', c => { const s = c.text(0), n = c.int(1); if (n < 0 || s.length * n > MAX_TEXT) fail('#VALUE!'); return s.repeat(n); });
  def('EXACT', 2, 2, 'v', c => c.text(0) === c.text(1));
  def('CHAR', 1, 1, 'v', c => { const n = c.int(0); if (n < 1) fail('#VALUE!'); return charOf(n); });
  def('CODE', 1, 1, 'v', c => { const s = c.text(0); if (!s) fail('#VALUE!'); return codeOf([...s][0]); });
  def('UNICHAR', 1, 1, 'v', c => { const n = c.int(0); if (n < 1 || n > 0x10ffff) fail('#VALUE!'); if (n >= 0xd800 && n <= 0xdfff) fail('#N/A'); return String.fromCodePoint(n); });
  def('UNICODE', 1, 1, 'v', c => { const s = c.text(0); if (!s) fail('#VALUE!'); return s.codePointAt(0); });
  def('ASC', 1, 1, 'v', c => toHalf(c.text(0)));
  def('DBCS', 1, 1, 'v', c => toFull(c.text(0)));
  def('VALUE', 1, 1, 'v', c => { const v = c.v(0); if (typeof v === 'number') return v; if (v === BLANK) return 0; if (typeof v === 'boolean') fail('#VALUE!'); const n = textNumber(v); return n === null ? fail('#VALUE!') : n; });
  // TEXT gives at most 255 characters (a longer result is #VALUE!). Excel also refuses formats with long runs of literal
  // characters (around 60 or more, by a count the page does not copy).
  def('TEXT', 2, 2, 'v', c => { const v = c.v(0), f = c.text(1), s = formatValue(v, f); return s.length > 255 ? fail('#VALUE!') : s; });
  // CONCATENATE takes each argument as one value, in order, and cuts the text at MAX_TEXT. CONCAT and TEXTJOIN read the
  // references given to a text or to the delimiter only after the other arguments, so that an error given directly comes
  // before an error in a referenced cell (=CONCAT(E1,NA()) is #N/A though E1 holds #DIV/0!); a text longer than MAX_TEXT
  // is #VALUE!. The ignore_empty of TEXTJOIN is read in its turn.
  def('CONCATENATE', 1, 255, 'v', c => { let out = ''; for (let i = 0; i < c.n; i++) out += toText(c.v(i)); return out.slice(0, MAX_TEXT); });
  const textsOf = v => [v].flat().map(x => { if (isErr(x)) fail(x); return toText(x); });
  function joinParts(c, from) {
    const now = [], later = [];
    for (let i = from; i < c.n; i++) { if (c.nodes[i].op === 'ref' && !c.values) later.push(i); else now.push([i, textsOf(c.raw(i))]); }
    const parts = new Map(now); for (const i of later) parts.set(i, textsOf(c.raw(i)));
    return parts;
  }
  def('CONCAT', 1, 254, null, c => { const parts = joinParts(c, 0), out = [...Array(c.n).keys()].flatMap(i => parts.get(i)).join(''); return out.length > MAX_TEXT ? fail('#VALUE!') : out; });
  def('TEXTJOIN', 3, 254, null, c => {
    const delimNode = c.nodes[0], deferred = delimNode.op === 'ref' && !c.values;
    let delims = deferred ? null : textsOf(c.raw(0)), skip = true;
    if (c.given(1)) {
      // An empty cell is FALSE; the text "" is not TRUE or FALSE (#VALUE!).
      if (c.nodes[1].op === 'literal' && c.nodes[1].v === '') fail('#VALUE!');
      skip = toBool(c.v(1));
    }
    const now = [], later = [];
    for (let i = 2; i < c.n; i++) { if (c.nodes[i].op === 'ref' && !c.values) later.push(i); else now.push([i, textsOf(c.raw(i))]); }
    if (deferred) delims = textsOf(c.raw(0));
    const parts = new Map(now); for (const i of later) parts.set(i, textsOf(c.raw(i)));
    // The next delimiter is taken only when a text is put in. Excel counts it in front of an empty text before it leaves
    // the text out, so such a text can still make the result too long.
    let out = '', count = 0;
    for (let i = 2; i < c.n; i++) for (const t of parts.get(i)) {
      if (count) { const sep = delims[(count - 1) % delims.length]; if (out.length + sep.length + t.length > MAX_TEXT) fail('#VALUE!'); if (skip && t === '') continue; out += sep + t; count++; }
      else if (!skip || t !== '') { if (t.length > MAX_TEXT) fail('#VALUE!'); out = t; count = 1; }
    }
    return out;
  });

  // ---- Dates and times
  const inDays = n => n < 0 || n >= LAST_DAY + 1 ? fail('#NUM!') : n;
  const serialArg = (c, i) => inDays(c.num(i));
  const strictSerial = (c, i) => inDays(c.strict(i));
  def('DATE', 3, 3, 'v', c => {
    let y = c.int(0); const m = c.int(1), d = c.int(2);
    if (y < 0 || y >= 10000) fail('#NUM!'); if (y < 1900) y += 1900;
    const s = serialOf(y, m, d);
    if (s < 0 || s > LAST_DAY) fail('#NUM!');
    return s;
  });
  def('YEAR', 1, 1, 'v', c => civil(serialArg(c, 0)).y);
  def('MONTH', 1, 1, 'v', c => civil(serialArg(c, 0)).m);
  def('DAY', 1, 1, 'v', c => civil(serialArg(c, 0)).d);
  const WEEKDAY_START = {1: [0, 1], 2: [1, 1], 3: [1, 0], 11: [1, 1], 12: [2, 1], 13: [3, 1], 14: [4, 1], 15: [5, 1], 16: [6, 1], 17: [0, 1]};
  def('WEEKDAY', 1, 2, 'v', c => { const n = serialArg(c, 0), t = WEEKDAY_START[c.int(1, 1)]; if (!t) fail('#NUM!'); return (weekday(n) - t[0] + 7) % 7 + t[1]; });
  function isoWeek(n) {
    n = Math.floor(n); const wd = (weekday(n) + 6) % 7, thursday = n - wd + 3, {y} = civil(thursday);
    return Math.floor((thursday - serialOf(y, 1, 1)) / 7) + 1;
  }
  const WEEK_START = {1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0};
  def('WEEKNUM', 1, 2, 'v', c => {
    const n = strictSerial(c, 0), t = Math.trunc(c.strict(1, 1));
    if (t === 21) return isoWeek(n);
    if (!(t in WEEK_START)) fail('#NUM!');
    const jan1 = serialOf(civil(n).y, 1, 1), offset = (weekday(jan1) - WEEK_START[t] + 7) % 7;
    return Math.floor((Math.floor(n) - jan1 + offset) / 7) + 1;
  });
  def('ISOWEEKNUM', 1, 1, 'v', c => isoWeek(serialArg(c, 0)));
  // EDATE and EOMONTH keep to the calendar, where 1900/02 has 28 days: EDATE(DATE(1900,1,31),1) and EDATE(60,0) are 59.
  function addMonths(c, end) {
    const n = strictSerial(c, 0), k = Math.trunc(c.strict(1)), {y, m, d} = civil(n), target = m - 1 + k, ty = y + Math.floor(target / 12), tm = (target % 12 + 12) % 12 + 1;
    if (ty < 1900 || ty > 9999) fail('#NUM!');
    const s = serialOf(ty, tm, end ? realDaysIn(ty, tm) : Math.min(d, realDaysIn(ty, tm)));
    if (s < 0 || s > LAST_DAY) fail('#NUM!');
    return s;
  }
  def('EDATE', 2, 2, 'v', c => addMonths(c, false));
  def('EOMONTH', 2, 2, 'v', c => addMonths(c, true));
  def('DATEDIF', 3, 3, 'v', c => {
    const a = Math.floor(serialArg(c, 0)), b = Math.floor(serialArg(c, 1)), unit = c.text(2).toUpperCase();
    if (a > b) fail('#NUM!');
    const s = civil(a), e = civil(b);
    let months = (e.y - s.y) * 12 + e.m - s.m; if (e.d < s.d) months--;
    switch (unit) {
      case 'Y': return Math.floor(months / 12);
      case 'M': return months;
      case 'D': return b - a;
      case 'YM': return months % 12;
      case 'MD': { if (e.d >= s.d) return e.d - s.d; const pm = e.m === 1 ? [e.y - 1, 12] : [e.y, e.m - 1]; return daysIn(pm[0], pm[1]) - s.d + e.d; }
      case 'YD': {
        // As measured on 77 pairs of dates: when the end's month and day come before the start's, the days from the
        // start to that month and day of the next year; else the days between the two months and days in the start's
        // year, or in the end's year when the end's day of the month is smaller (day 29 of February 2025 is March 1).
        const day = (yy, mm, dd) => utc(yy, mm, dd) / DAY;
        if (e.m < s.m || e.m === s.m && e.d < s.d) return day(s.y + 1, e.m, e.d) - day(s.y, s.m, s.d);
        const yy = e.d >= s.d ? s.y : e.y;
        return day(yy, e.m, e.d) - day(yy, s.m, s.d);
      }
    }
    return fail('#NUM!');
  });
  def('DAYS', 2, 2, 'v', c => Math.floor(serialArg(c, 0)) - Math.floor(serialArg(c, 1)));
  /** Which weekdays (0 Sunday ... 6 Saturday) are days off: 1-7 (two days, 1 Saturday and Sunday) and 11-17 (one day,
      11 Sunday ... 17 Saturday) as Excel numbers them, or seven 0s and 1s from Monday. */
  function weekendOf(c, i) {
    if (c.n <= i || !c.given(i)) return [true, false, false, false, false, false, true];
    const v = c.v(i);
    if (typeof v === 'string') {
      if (!/^[01]{7}$/.test(v)) fail('#VALUE!');
      return [v[6] === '1', ...[...v.slice(0, 6)].map(ch => ch === '1')];
    }
    const k = Math.trunc(toNumber(v)), off = Array(7).fill(false);
    if (k >= 1 && k <= 7) { off[(k + 5) % 7] = true; off[(k + 6) % 7] = true; }
    else if (k >= 11 && k <= 17) off[k - 11] = true;
    else fail('#NUM!');
    return off;
  }
  function holidaysOf(c, i) {
    if (c.n <= i || !c.given(i)) return new Set();
    const out = new Set();
    for (const x of [c.raw(i)].flat()) {
      if (isErr(x)) fail(x); if (x === BLANK) continue; if (typeof x === 'boolean') fail('#VALUE!');
      const n = toNumber(x); if (n < 0) fail('#NUM!'); out.add(Math.floor(n));
    }
    return out;
  }
  function networkDays(c, weekendIndex, holidayIndex) {
    let a = Math.floor(strictSerial(c, 0)), b = Math.floor(strictSerial(c, 1)), sign = 1;
    const off = weekendOf(c, weekendIndex), holidays = holidaysOf(c, holidayIndex);
    if (a > b) { [a, b] = [b, a]; sign = -1; }
    let n = 0; for (let d = a; d <= b; d++) if (!off[weekday(d)] && !holidays.has(d)) n++;
    return sign * n;
  }
  def('NETWORKDAYS', 2, 3, null, c => networkDays(c, 99, 2));
  def('NETWORKDAYS.INTL', 2, 4, null, c => networkDays(c, 2, 3));
  // WORKDAY takes the days down to a whole number (-1.5 is -2), and with every weekday off, only 0 days.
  function workday(c, weekendIndex, holidayIndex) {
    let d = Math.floor(strictSerial(c, 0)); const k = Math.floor(c.strict(1)), off = weekendOf(c, weekendIndex), holidays = holidaysOf(c, holidayIndex), step = k < 0 ? -1 : 1;
    if (k !== 0 && off.every(Boolean)) fail('#VALUE!');
    for (let left = Math.abs(k); left > 0;) { d += step; if (d < 0 || d > LAST_DAY) fail('#NUM!'); if (!off[weekday(d)] && !holidays.has(d)) left--; }
    return d;
  }
  def('WORKDAY', 2, 3, null, c => workday(c, 99, 2));
  def('WORKDAY.INTL', 2, 4, null, c => workday(c, 2, 3));
  def('TIME', 3, 3, 'v', c => {
    const h = c.int(0), m = c.int(1), s = c.int(2);
    if (h > 32767 || m > 32767 || s > 32767) fail('#NUM!');
    const total = h * 3600 + m * 60 + s;
    if (total < 0) fail('#NUM!');
    // Whole days are taken away from the fraction: TIME(25,0,0) is 25/24-1, 0.04166666666666674, as in Excel.
    const t = total / 86400;
    return t - Math.floor(t);
  });
  const timePart = (c, part) => { const n = serialArg(c, 0), t = Math.round((n - Math.floor(n)) * 86400) % 86400; return part === 'h' ? Math.floor(t / 3600) : part === 'm' ? Math.floor(t % 3600 / 60) : t % 60; };
  def('HOUR', 1, 1, 'v', c => timePart(c, 'h'));
  def('MINUTE', 1, 1, 'v', c => timePart(c, 'm'));
  def('SECOND', 1, 1, 'v', c => timePart(c, 's'));
  const dateText = c => { const v = c.v(0); if (typeof v !== 'string') fail('#VALUE!'); const s = halfWidth(v).trim(); if (plainNumber(s) !== null) fail('#VALUE!'); const n = textNumber(s); return n === null ? fail('#VALUE!') : n; };
  def('DATEVALUE', 1, 1, 'v', c => Math.floor(dateText(c) + 1e-9));
  def('TIMEVALUE', 1, 1, 'v', c => { const n = dateText(c); return n - Math.floor(n + 1e-9) < 0 ? 0 : n - Math.floor(n + 1e-9); });
  def('TODAY', 0, 0, null, () => Math.floor(nowSerial()));
  def('NOW', 0, 0, null, () => nowSerial());

  const FUNCTIONS = Object.keys(FN);
  /** Reads (parseOnly) or computes a formula. ref(a, b) gives the value of cell a, or the values of range a:b (with the
      range), BLANK for an empty cell; at is the formula's cell ({row, col, formulaAt(r, c)}, or [row, col]). */
  function evalFormula(formula, ref, parseOnly = false, at = null) {
    const ast = parse(formula);
    if (parseOnly) return ast;
    const env = {ref: ref || (() => BLANK), row: Array.isArray(at) ? at[0] : at?.row, col: Array.isArray(at) ? at[1] : at?.col, formulaAt: at?.formulaAt || (() => null)};
    let v = run(ast, env, false);
    if (Array.isArray(v)) v = single(v, env);
    if (typeof v === 'number' && !Number.isFinite(v)) return '#NUM!';
    return v === BLANK ? 0 : v;
  }

  return {BLANK, FUNCTIONS, evalFormula, parse, numberLiteral, englishNames, isErr, fail, errorOf, MAX_TEXT, arrayOf, dims, exactDigits, roundDigits, digits15, to15, valueOfDigits, fixedParts, roundPlaces, truncPlaces,
    formatTokens, generalText, formatValue, toHalf, toFull, charOf, codeOf,
    numberText, DAY, EPOCH, LAST_DAY, serialOf, civil, weekday, daysIn, realDaysIn, validDate, clock, nowSerial, ERAS, eraOf, MONTHS, monthByName, DAYS, DAYS_JA, halfWidth,
    plainNumber, timeOf, dateOf, dateTimeOf, textNumber, toNumber, toText, toBool, sameText, compareText, compare, wildcard, matchText, criterion};
});
