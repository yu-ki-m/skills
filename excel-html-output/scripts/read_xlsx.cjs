#!/usr/bin/env node
// Reads an .xlsx file into a workbook (excel-html-workbook/v1) that the page shows as Excel shows the file, as far as
// the page can hold it. What the page cannot hold is replaced by the nearest thing it can (a number format, a border,
// a fill) or left out (charts, comments, ...), and every such case is reported.
//
//   node scripts/read_xlsx.cjs 見積.xlsx        (scripts/build.py --from-xlsx runs this)
//
// It prints {workbook, warnings, notes} as JSON: warnings are sentences for the user, notes say which cells and fields
// differ from the file on purpose (scripts/verify_xlsx.py --source uses them).
// A formula stays a formula when the page can compute it. That is asked of assets/core.js each time (it must read the
// formula and know every function in it), so a function added to core.js is kept from then on. Other formulas become
// the value Excel saved. Only Node.js built-ins are used; the file is never executed.
'use strict';
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const C = require(path.join(__dirname, '..', 'assets', 'core.js'));

const LIMIT = {rows: 500, cols: 100, sheets: 30};
const EMU_PER_PX = 9525;
// Excel's legacy palette by index; styles.xml may replace it (<indexedColors>). 64 and above are system colours.
const PALETTE = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
// The Office theme's colours (2013-2022) in Excel's theme-index order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink),
// for a file without a theme part.
const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
// A theme's system colour shows as the colour Windows has for it (not the lastClr written in the file; measured).
const SYSTEM_COLOR = {windowText: '000000', window: 'FFFFFF'};
// Column widths are counted in the widest digit of the Normal style's font. Its width at 96dpi for 11pt fonts; other
// sizes scale with the size. For Yu Gothic 11pt it is 8px (10px at 120dpi, measured in desktop Excel 16.0). Fonts whose
// width rounds differently at other display scalings (MS PGothic, Meiryo, Calibri ...) give approximate widths there.
const DIGIT_PX_11 = {'游ゴシック': 8, 'yu gothic': 8, '游ゴシック light': 8, 'yu gothic light': 8, '游明朝': 8, 'yu mincho': 8,
  'ｍｓ ｐゴシック': 8, 'ms pgothic': 8, 'ｍｓ ゴシック': 8, 'ms gothic': 8, 'ｍｓ 明朝': 8, 'ms mincho': 8, 'ｍｓ ｐ明朝': 8, 'ms pmincho': 8,
  'メイリオ': 9, 'meiryo': 9, 'meiryo ui': 9, 'calibri': 7, 'aptos narrow': 7, 'aptos': 7, 'arial': 8};
// The height (pt) of a row that fits the Normal style's font of 11pt, at 100% display scaling (other sizes in
// proportion). Excel fits rows to it unless the sheet's default height is marked as set (customHeight). These are the
// commonly known values at 100%; at 125% desktop Excel 16.0 showed 18, 14.4 and 13.2pt, within the checker's 0.8pt.
const ROW_PT_11 = {'yu gothic': 18.75, 'yu mincho': 18.75, 'calibri': 15, 'aptos': 15, 'ms pgothic': 13.5, 'ms gothic': 13.5};
// The height of an empty row in a sheet that gives no default height: Excel's own (Yu Gothic 11pt), whatever the font.
const EXCEL_ROW_PT = 18.75;
const BORDER = {hair: '0.5px solid', thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed',
  mediumDashed: '2px dashed', dotted: '1px dotted', double: '3px double'};
const NEAR_BORDER = {dashDot: '1px dashed', dashDotDot: '1px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed'};
// Share of the foreground colour in Excel's pattern fills (the rest is the background colour).
const PATTERN = {gray0625: 0.0625, gray125: 0.125, lightGray: 0.25, mediumGray: 0.5, darkGray: 0.75, lightHorizontal: 0.25, lightVertical: 0.25,
  lightDown: 0.25, lightUp: 0.25, lightGrid: 0.4, lightTrellis: 0.4, darkHorizontal: 0.5, darkVertical: 0.5, darkDown: 0.5, darkUp: 0.5, darkGrid: 0.6, darkTrellis: 0.6};
// Built-in number formats (Japanese Excel's spelling where the locale decides it).
const BUILTIN_FORMAT = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '"¥"#,##0;"¥"\\-#,##0', 6: '"¥"#,##0;[Red]"¥"\\-#,##0',
  7: '"¥"#,##0.00;"¥"\\-#,##0.00', 8: '"¥"#,##0.00;[Red]"¥"\\-#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
  27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"', 32: 'h"時"mm"分"',
  33: 'h"時"mm"分"ss"秒"', 34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: '[$-411]ggge"年"m"月"d"日"', 55: 'yyyy"年"m"月"',
  56: 'm"月"d"日"', 57: '[$-411]ge.m.d', 58: '[$-411]ggge"年"m"月"d"日"'};

// ---- ZIP and XML ----------------------------------------------------------------------------------------------------

function unzip(buffer) {
  let end = buffer.length - 22;
  while (end >= 0 && buffer.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) throw Error('xlsx（ZIP）として読めません。古い .xls やパスワード付きのファイルなら、Excelで .xlsx として保存し直してください。');
  const count = buffer.readUInt16LE(end + 10), parts = new Map();
  let p = buffer.readUInt32LE(end + 16);
  if (p === 0xffffffff) throw Error('ZIP64 の xlsx には対応していません。');
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(p) !== 0x02014b50) throw Error('xlsx（ZIP）の目次が壊れています。');
    const flags = buffer.readUInt16LE(p + 8), method = buffer.readUInt16LE(p + 10), size = buffer.readUInt32LE(p + 20);
    const nameLength = buffer.readUInt16LE(p + 28), extra = buffer.readUInt16LE(p + 30), comment = buffer.readUInt16LE(p + 32), local = buffer.readUInt32LE(p + 42);
    const name = buffer.toString('utf8', p + 46, p + 46 + nameLength);
    if (flags & 1) throw Error('暗号化された xlsx は読めません。');
    const start = local + 30 + buffer.readUInt16LE(local + 26) + buffer.readUInt16LE(local + 28), raw = buffer.subarray(start, start + size);
    if (method !== 0 && method !== 8) throw Error('未対応の圧縮方式の xlsx です: ' + method);
    parts.set(name.replace(/\\/g, '/'), method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw));
    p += 46 + nameLength + extra + comment;
  }
  return parts;
}

const ENTITY = {lt: '<', gt: '>', amp: '&', quot: '"', apos: "'"};
const decode = s => s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, e) => e[0] === '#' ? String.fromCodePoint(/^#x/i.test(e) ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENTITY[e]);

/**
 * Elements as {name, attrs, children}; text as strings. Entities other than the five XML ones are not expanded.
 * Line ends are made LF first, as an XML reader must (a CR that is meant is written as _x000D_ or &#13;).
 */
function parseXml(source) {
  const text = source.replace(/\r\n?/g, '\n');
  const root = {name: '', attrs: {}, children: []}, stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>!?]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(text))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) top.children.push(m[1]);
    else if (m[2]) { if (stack.length > 1) stack.pop(); }
    else if (m[3]) {
      const attrs = {};
      for (const a of m[4].matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[a[1]] = decode(a[2] ?? a[3]);
      const el = {name: m[3], attrs, children: []};
      top.children.push(el);
      if (!m[5]) stack.push(el);
    } else if (m[6] !== undefined) top.children.push(decode(m[6]));
  }
  return root.children.find(c => typeof c !== 'string') || root;
}
const local = name => name.replace(/^.*:/, '');
const kids = (el, name) => (el?.children || []).filter(c => typeof c !== 'string' && (!name || local(c.name) === name));
const kid = (el, name) => kids(el, name)[0];
const textOf = el => (el?.children || []).map(c => typeof c === 'string' ? c : textOf(c)).join('');
// Excel writes characters XML cannot hold as _xHHHH_ (and a literal "_x" escape with its "_" as _x005F_).
const unescapeX = s => s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));

function resolve(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const piece of target.split('/')) piece === '..' ? parts.pop() : piece && piece !== '.' && parts.push(piece);
  return parts.join('/');
}
function relationships(parts, partName) {
  const name = partName ? partName.replace(/([^/]+)$/, '_rels/$1.rels') : '_rels/.rels', out = {};
  if (!parts.has(name)) return out;
  for (const r of kids(parseXml(parts.get(name).toString('utf8')), 'Relationship')) {
    const external = r.attrs.TargetMode === 'External';
    out[r.attrs.Id] = {type: (r.attrs.Type || '').replace(/^.*\//, ''), target: external ? r.attrs.Target : resolve(partName, r.attrs.Target), external};
  }
  return out;
}

// ---- Colours ----------------------------------------------------------------------------------------------------------

// Excel's tint: the colour in HLS with HLSMAX 240 (Windows' integer conversion), lightness changed, rounded back.
// Matched every one of 132 theme colour and tint pairs that desktop Excel 16.0 shows.
function tint(hex, t) {
  if (!t) return hex;
  const HLSMAX = 240, RGBMAX = 255, div = (a, b) => Math.floor(a / b);
  const [R, G, B] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
  const cMax = Math.max(R, G, B), cMin = Math.min(R, G, B);
  let L = div((cMax + cMin) * HLSMAX + RGBMAX, 2 * RGBMAX), H, S;
  if (cMax === cMin) { S = 0; H = div(HLSMAX * 2, 3); } else {
    S = L <= HLSMAX / 2 ? div((cMax - cMin) * HLSMAX + div(cMax + cMin, 2), cMax + cMin)
      : div((cMax - cMin) * HLSMAX + div(2 * RGBMAX - cMax - cMin, 2), 2 * RGBMAX - cMax - cMin);
    const d = c => div((cMax - c) * div(HLSMAX, 6) + div(cMax - cMin, 2), cMax - cMin);
    H = R === cMax ? d(B) - d(G) : G === cMax ? div(HLSMAX, 3) + d(R) - d(B) : div(2 * HLSMAX, 3) + d(G) - d(R);
    if (H < 0) H += HLSMAX;
    if (H > HLSMAX) H -= HLSMAX;
  }
  L = t < 0 ? Math.floor(L * (1 + t)) : Math.floor(L * (1 - t) + (HLSMAX - Math.floor(HLSMAX * (1 - t))));
  const hue = (n1, n2, h) => {
    if (h < 0) h += HLSMAX;
    if (h > HLSMAX) h -= HLSMAX;
    if (h < HLSMAX / 6) return n1 + div((n2 - n1) * h + div(HLSMAX, 12), div(HLSMAX, 6));
    if (h < HLSMAX / 2) return n2;
    if (h < div(HLSMAX * 2, 3)) return n1 + div((n2 - n1) * (div(HLSMAX * 2, 3) - h) + div(HLSMAX, 12), div(HLSMAX, 6));
    return n1;
  };
  let rgb;
  if (S === 0) rgb = [L, L, L];
  else {
    const m2 = L <= HLSMAX / 2 ? div(L * (HLSMAX + S) + div(HLSMAX, 2), HLSMAX) : L + S - div(L * S + div(HLSMAX, 2), HLSMAX), m1 = 2 * L - m2;
    rgb = [hue(m1, m2, H + div(HLSMAX, 3)), hue(m1, m2, H), hue(m1, m2, H - div(HLSMAX, 3))];
  }
  return rgb.map(v => Math.max(0, Math.min(255, div(v * RGBMAX + div(HLSMAX, 2), HLSMAX))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** '#rrggbb' for an Excel colour element, or fallback when it is missing or automatic. */
function colorOf(el, book, fallback) {
  if (!el) return fallback;
  const a = el.attrs;
  let hex = null;
  if (a.rgb && /^[0-9A-Fa-f]{6,8}$/.test(a.rgb)) hex = a.rgb.slice(-6).toUpperCase();
  else if (a.theme !== undefined) hex = book.theme[+a.theme] || null;
  else if (a.indexed !== undefined) hex = +a.indexed < book.palette.length ? book.palette[+a.indexed] : +a.indexed === 65 ? 'FFFFFF' : '000000';
  if (!hex) return fallback;
  return '#' + tint(hex, a.tint ? +a.tint : 0).toLowerCase();
}
const mix = (fg, bg, share) => '#' + [1, 3, 5].map(i => Math.round(parseInt(fg.slice(i, i + 2), 16) * share + parseInt(bg.slice(i, i + 2), 16) * (1 - share))
  .toString(16).padStart(2, '0')).join('');

// ---- Number formats ---------------------------------------------------------------------------------------------------

function sections(code) {
  const out = [];
  let cur = '', quote = false, bracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\' && !quote) { cur += ch + (code[i + 1] ?? ''); i++; continue; }
    if (ch === '"') quote = !quote;
    else if (!quote && ch === '[') bracket = true;
    else if (!quote && ch === ']') bracket = false;
    if (ch === ';' && !quote && !bracket) { out.push(cur); cur = ''; } else cur += ch;
  }
  out.push(cur);
  return out;
}
// The shape of a format section without colours, locales, quotes, escapes and padding: "¥"#,##0 and [$¥-411]#,##0 → ¥#,##0.
const plain = s => s.replace(/\[\$([^\]-]*)(?:-[0-9A-Fa-f]+)?\]/g, '$1').replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1')
  .replace(/\\(.)/g, '$1').replace(/_./g, '').replace(/\*./g, '').trim();

/** The page's format for an Excel format code, and whether it shows numbers the same way. */
function mapFormat(code) {
  if (/^(General|G\/標準)$/i.test(code.trim())) return {fmt: 'General', exact: true};
  const parts = sections(code), first = plain(parts[0]);
  // A colour, a condition or a numeral system ([Red], [>=100], [DBNum1]; not a currency or locale, [$...]) shows numbers
  // otherwise than the page. Negative numbers: the page puts "-" in front, so a second section shows them the same way
  // only when it is "-" and the first section.
  const marked = s => /\[(?!\$)[^\]]*\]/.test(s);
  const same = !marked(parts[0]) && (parts.length === 1 || parts.length === 2 && plain(parts[1]) === '-' + first && !marked(parts[1]));
  const known = C.FORMATS.find(f => f === first);
  if (known) return {fmt: known, exact: same};
  const lower = first.toLowerCase(), decimals = (/\.(0+)/.exec(first) || [, ''])[1].length;
  let fmt;
  if (/[hs]/.test(lower.replace(/"[^"]*"/g, '')) && !/[yd]/.test(lower)) fmt = 'General';
  else if (/[yd]|m.*d|ge/.test(lower)) fmt = 'yyyy/mm/dd';
  else if (first.includes('%')) fmt = decimals ? '0.00%' : '0%';
  else if (/e[+-]/i.test(first) || first.includes('?')) fmt = 'General';
  else if (first.includes('¥') || first.includes('￥')) fmt = '¥#,##0';
  else if (first.includes(',')) fmt = decimals ? '#,##0.00' : '#,##0';
  else if (first.includes('0') || first.includes('#')) fmt = decimals ? '0.00' : '0';
  else if (first === '@') fmt = '@';
  else fmt = 'General';
  return {fmt, exact: false};
}

// ---- Styles -----------------------------------------------------------------------------------------------------------

// The font of a file without fonts (or without a styles part).
const NO_FONT = {name: 'Yu Gothic', size: 11, bold: false, italic: false, strike: false, underline: null, color: '#000000'};

function readStyles(xml, book) {
  const formats = {...BUILTIN_FORMAT};
  for (const f of kids(kid(xml, 'numFmts'), 'numFmt')) formats[f.attrs.numFmtId] = f.attrs.formatCode;
  const flag = (el, name) => { const e = kid(el, name); return !!e && e.attrs.val !== '0' && e.attrs.val !== 'false'; };
  const fonts = kids(kid(xml, 'fonts'), 'font').map(f => ({
    // The page names fonts in English, as its ribbon lists them: 游ゴシック is Yu Gothic.
    name: (n => C.FONT_ALIAS[n] || n)(kid(f, 'name')?.attrs.val || 'Yu Gothic'), size: +(kid(f, 'sz')?.attrs.val || 11), bold: flag(f, 'b'), italic: flag(f, 'i'), strike: flag(f, 'strike'),
    underline: (u => !u || u.attrs.val === 'none' ? null : u.attrs.val || 'single')(kid(f, 'u')), // <u val="none"/> is no underline
    vertAlign: kid(f, 'vertAlign')?.attrs.val, color: colorOf(kid(f, 'color'), book, '#000000'),
  }));
  const fills = kids(kid(xml, 'fills'), 'fill').map(f => {
    const p = kid(f, 'patternFill'), g = kid(f, 'gradientFill');
    if (g) {
      const stops = kids(g, 'stop').map(s => colorOf(kid(s, 'color'), book, '#ffffff'));
      return {fill: stops.length ? stops.reduce((acc, c, i) => i ? mix(c, acc, 1 / (i + 1)) : c) : 'none', note: 'グラデーションの塗りは、その平均の色1色にした'};
    }
    const type = p?.attrs.patternType || 'none'; // without a pattern type, Excel does not fill (whatever the colours; measured)
    if (type === 'none') return {fill: 'none'};
    const fg = colorOf(kid(p, 'fgColor'), book, '#000000');
    if (type === 'solid') return {fill: fg};
    return {fill: mix(fg, colorOf(kid(p, 'bgColor'), book, '#ffffff'), PATTERN[type] ?? 0.5), note: `網かけ（${type}）の塗りは、混ぜた色1色にした`};
  });
  const borders = kids(kid(xml, 'borders'), 'border').map(b => {
    const out = {notes: []};
    for (const [side, key] of [['top', 'borderTop'], ['right', 'borderRight'], ['bottom', 'borderBottom'], ['left', 'borderLeft']]) {
      const el = kid(b, side) || kid(b, side === 'left' ? 'start' : side === 'right' ? 'end' : side);
      const style = el?.attrs.style;
      if (!style || style === 'none') { out[key] = 'none'; continue; }
      const kind = BORDER[style] || NEAR_BORDER[style] || '1px solid';
      if (!BORDER[style]) out.notes.push({field: {borderTop: 'bTop', borderRight: 'bRight', borderBottom: 'bBottom', borderLeft: 'bLeft'}[key], reason: `罫線（${style}）は「${kind}」にした`});
      out[key] = kind + ' ' + colorOf(kid(el, 'color'), book, '#000000');
    }
    const diagonal = kid(b, 'diagonal');
    if (diagonal?.attrs.style && diagonal.attrs.style !== 'none' && (b.attrs.diagonalUp === '1' || b.attrs.diagonalDown === '1'))
      out.notes.push({field: null, reason: '斜めの罫線は移せない'});
    return out;
  });
  const xfs = kids(kid(xml, 'cellXfs'), 'xf').map(x => ({
    numFmtId: +(x.attrs.numFmtId || 0), fontId: +(x.attrs.fontId || 0), fillId: +(x.attrs.fillId || 0), borderId: +(x.attrs.borderId || 0),
    alignment: kid(x, 'alignment')?.attrs || {},
  }));
  const normalXf = kid(kid(xml, 'cellStyleXfs'), 'xf');
  // Without a Normal style (no styles part), Excel uses its own: Yu Gothic 11pt, vertically centred (Japanese Excel).
  const normal = {font: fonts[+(normalXf?.attrs.fontId || 0)] || fonts[0] || NO_FONT, vertical: kid(normalXf, 'alignment')?.attrs.vertical || (normalXf ? 'bottom' : 'center')};
  return {formats, fonts, fills, borders, xfs, normal};
}

// The Excel spelling of the page's formats that differ from it in writing (as scripts/xlsx.cjs writes them).
const EXCEL_CODE = {'¥#,##0': '"¥"#,##0'};
// For a border side: the neighbouring cell (rows, columns away) and its side that is the same line in Excel.
const NEIGHBOUR_SIDE = {bTop: [-1, 0, 'bBottom'], bBottom: [1, 0, 'bTop'], bLeft: [0, -1, 'bRight'], bRight: [0, 1, 'bLeft']};
const HORIZONTAL = {general: 'general', left: 'left', center: 'center', right: 'right', fill: 'left', justify: 'left', distributed: 'center', centerContinuous: 'center'};
const VERTICAL = {top: 'top', center: 'middle', bottom: 'bottom', justify: 'middle', distributed: 'middle'};

/** The page's style for cellXfs[index] (a full style), with what had to change: [{field, reason}] (field as verify_xlsx.py names it). */
function styleOf(styles, index) {
  const cache = styles.cache || (styles.cache = new Map());
  if (cache.has(index)) return cache.get(index);
  const xf = styles.xfs[index] || styles.xfs[0] || {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: {vertical: styles.normal.vertical}};
  const font = styles.fonts[xf.fontId] || styles.fonts[0] || NO_FONT, fill = styles.fills[xf.fillId] || {fill: 'none'}, border = styles.borders[xf.borderId] || {notes: []};
  const a = xf.alignment, notes = [];
  const st = {...C.DEFAULT};
  st.font = font.name;
  st.size = Math.min(96, Math.max(6, font.size));
  if (st.size !== font.size) notes.push({field: 'size', reason: `文字の大きさ ${font.size}pt は ${st.size}pt にした（HTMLは6〜96pt）`});
  Object.assign(st, {bold: font.bold, italic: font.italic, strike: font.strike, underline: !!font.underline, color: font.color});
  if (font.underline && font.underline !== 'single') notes.push({field: 'underline', reason: `下線（${font.underline}）は一重の下線にした`});
  if (font.vertAlign && font.vertAlign !== 'baseline') notes.push({field: null, reason: '上付き・下付きの文字は、普通の文字にした'});
  st.fill = fill.fill;
  if (fill.note) notes.push({field: 'fill', reason: fill.note});
  for (const key of ['borderTop', 'borderRight', 'borderBottom', 'borderLeft']) st[key] = border[key] || 'none';
  notes.push(...border.notes);
  const horizontal = a.horizontal || 'general';
  st.align = HORIZONTAL[horizontal] || 'general';
  if (!['general', 'left', 'center', 'right'].includes(horizontal)) notes.push({field: 'halign', reason: `横位置（${horizontal}）は「${st.align}」にした`});
  const vertical = a.vertical || 'bottom';
  st.valign = VERTICAL[vertical] || 'bottom';
  if (!['top', 'center', 'bottom'].includes(vertical)) notes.push({field: 'valign', reason: `縦位置（${vertical}）は「${st.valign}」にした`});
  st.wrap = a.wrapText === '1' || a.wrapText === 'true';
  const indent = +(a.indent || 0);
  st.indent = Math.min(15, indent);
  if (indent > 15) notes.push({field: 'indent', reason: `インデント ${indent} は 15 にした`});
  if (st.indent > 0 && st.align === 'general') { st.align = 'left'; notes.push({field: 'halign', reason: '「標準」の横位置のインデントは、左揃えのインデントにした'}); }
  if (+(a.textRotation || 0)) notes.push({field: null, reason: '文字の回転・縦書きは移せない'});
  if (a.shrinkToFit === '1' || a.shrinkToFit === 'true') notes.push({field: null, reason: '「縮小して全体を表示する」は移せない'});
  const code = styles.formats[xf.numFmtId] ?? 'General', mapped = mapFormat(code);
  st.numFmt = mapped.fmt;
  const result = {st, notes, format: {code, ...mapped}, size: font.size};
  cache.set(index, result);
  return result;
}

// ---- Cells and sheets -------------------------------------------------------------------------------------------------

function sharedStrings(parts, path, book) {
  if (!path || !parts.has(path)) return [];
  return kids(parseXml(parts.get(path).toString('utf8')), 'si').map(si => stringItem(si, book));
}
// The text of a <t> as Excel reads it: without the spaces, tabs and line ends around it, unless xml:space="preserve".
const tText = t => { const s = textOf(t); return t?.attrs['xml:space'] === 'preserve' ? s : s.replace(/^[ \t\n]+|[ \t\n]+$/g, ''); };
/** The font a run's <rPr> gives its characters; what it leaves out (name, size, colour) is the cell's. */
function runFont(rPr, book) {
  const flag = name => { const e = kid(rPr, name); return !!e && e.attrs.val !== '0' && e.attrs.val !== 'false'; }, u = kid(rPr, 'u'), name = kid(rPr, 'rFont')?.attrs.val;
  const font = {bold: flag('b'), italic: flag('i'), strike: flag('strike'), underline: !!u && u.attrs.val !== 'none'};
  if (name) font.font = C.FONT_ALIAS[name] || name;
  if (kid(rPr, 'sz')) font.size = +kid(rPr, 'sz').attrs.val;
  if (kid(rPr, 'color')) font.color = colorOf(kid(rPr, 'color'), book, '#000000');
  return font;
}
/**
 * The text of <si>/<is>: its <t>, or the <t> of its runs; phonetic readings (<rPh>) are not part of the text. Runs that
 * all have one format (one run is enough) show the text in it: font. Runs in different formats make it rich.
 */
function stringItem(si, book) {
  const runs = kids(si, 'r');
  const text = runs.length ? runs.map(r => tText(kid(r, 't'))).join('') : tText(kid(si, 't'));
  const formats = runs.map(r => kid(r, 'rPr') ? JSON.stringify(runFont(kid(r, 'rPr'), book)) : null);
  const uniform = formats.length > 0 && formats[0] !== null && formats.every(f => f === formats[0]);
  return {text: unescapeX(text), rich: !uniform && formats.some(f => f !== null), ...(uniform ? {font: JSON.parse(formats[0])} : {})};
}

/** Whether the page computes this function: core.js answers #NAME? only for a function it does not know. */
const knownFunctions = new Map();
function pageKnows(name) {
  if (!knownFunctions.has(name)) {
    let known = true;
    try { C.evalFormula('=' + name + '()', () => 0); } catch (e) { known = e.message !== '#NAME?'; }
    knownFunctions.set(name, known);
  }
  return knownFunctions.get(name);
}
/** {ok, maxRow, maxCol, functions} for a formula (with "="): ok when the page can read it and knows every function. */
function formulaSupport(f) {
  let ast;
  try { ast = C.evalFormula(f, null, true); } catch { return {ok: false, why: 'HTMLの数式として読めない（別のシートの参照・名前付きの範囲など）'}; }
  const unknown = new Set();
  let maxRow = -1, maxCol = -1;
  const walk = n => {
    if (!n || typeof n !== 'object') return;
    if (n.op === 'call') { if (!pageKnows(n.name)) unknown.add(n.name); n.args.forEach(walk); }
    else if (n.op === 'ref') for (const a of [n.a, n.b].filter(Boolean)) { const [r, c] = C.coord(a); maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
    else for (const key of ['x', 'l', 'r']) walk(n[key]);
  };
  walk(ast);
  if (unknown.size) return {ok: false, why: `HTMLで計算できない関数（${[...unknown].join('、')}）を使っている`};
  if (maxRow >= LIMIT.rows || maxCol >= LIMIT.cols) return {ok: false, why: 'HTMLの上限（500行・100列）の外を参照している'};
  return {ok: true, maxRow, maxCol};
}

function cellValue(c, book) {
  const t = c.attrs.t || 'n', v = kid(c, 'v'), raw = v ? textOf(v) : null;
  const text = s => ({t: 's', v: s.text, rich: s.rich, ...(s.font ? {font: s.font} : {})});
  if (t === 's') { const s = book.strings[+raw]; return s ? text(s) : null; }
  if (t === 'inlineStr') return text(stringItem(kid(c, 'is') || {children: []}, book));
  if (raw === null) return null;
  if (t === 'str') return {t: 's', v: unescapeX(raw)};
  if (t === 'b') return {t: 'b', v: raw === '1' || raw === 'true'};
  if (t === 'e') return {t: 'e', v: raw};
  if (t === 'd') { const d = Date.parse(raw.length <= 10 ? raw + 'T00:00:00Z' : raw.replace(/Z?$/, 'Z')); return Number.isFinite(d) ? {t: 'n', v: d / 86400000 + 25569} : null; }
  const n = Number(raw);
  return Number.isFinite(n) ? {t: 'n', v: n} : null;
}

const colLetters = c => C.addr(0, c).replace(/\d+$/, '');
const ranges = list => { const shown = list.slice(0, 8).join('、'); return list.length > 8 ? `${shown} ほか${list.length - 8}セル` : shown; };

/** Reads one worksheet into the page's sheet. report collects warnings and notes. */
function readSheet(parts, book, info, index, report) {
  const xml = parseXml(parts.get(info.path).toString('utf8')), rels = relationships(parts, info.path);
  const name = info.name, notes = {cells: {}, sheet: []}, say = (key, text) => report.add(name, key, text);
  const noteCell = (a, fields, reason) => { (notes.cells[a] ||= []).push({fields, reason}); };
  const cells = new Map(); // addr -> {r, c, value, formula, xf}
  const shared = new Map(), inCell = [];
  let nextRow = 0;
  for (const row of kids(kid(xml, 'sheetData'), 'row')) {
    const r = row.attrs.r ? +row.attrs.r - 1 : nextRow;
    nextRow = r + 1;
    let nextCol = 0;
    for (const c of kids(row, 'c')) {
      const [rr, cc] = c.attrs.r ? C.coord(c.attrs.r) : [r, nextCol];
      nextCol = cc + 1;
      const f = kid(c, 'f');
      let formula = null, formulaKind = null;
      if (f) {
        formulaKind = f.attrs.t || 'normal';
        if (formulaKind === 'shared') {
          const text = textOf(f);
          if (text) { shared.set(f.attrs.si, {text, r: rr, c: cc}); formula = text; }
          else { const master = shared.get(f.attrs.si); if (master) formula = C.shiftFormula('=' + master.text, rr - master.r, cc - master.c).slice(1); }
        } else formula = textOf(f) || null;
      }
      // A picture in a cell or a data type (a rich value, vm) is saved with an error value that Excel does not show.
      const richValue = c.attrs.vm !== undefined;
      if (richValue) inCell.push(C.addr(rr, cc));
      cells.set(C.addr(rr, cc), {r: rr, c: cc, value: richValue ? null : cellValue(c, book), formula, formulaKind, ref: f?.attrs.ref, xf: +(c.attrs.s || 0)});
    }
  }
  for (const a of inCell) noteCell(a, ['value2', 'text'], 'セルの中の画像やデータ型');
  if (inCell.length) say('vm', `セルの中の画像やデータ型（${ranges(inCell)}）は移せない`);
  const rowInfo = new Map();
  for (const row of kids(kid(xml, 'sheetData'), 'row')) if (row.attrs.r) rowInfo.set(+row.attrs.r - 1, row.attrs);
  const colDefs = kids(kid(xml, 'cols'), 'col'), colDef = c => colDefs.find(d => +d.attrs.min <= c + 1 && c + 1 <= +d.attrs.max);
  // A cell without a <c> has its row's format when the row has one, else its column's (Excel's whole-row and
  // whole-column formats), else the Normal style.
  const rowXf = r => { const a = rowInfo.get(r); return a && (a.customFormat === '1' || a.customFormat === 'true') ? +(a.s || 0) : null; };
  const colXf = c => { const d = colDef(c); return d && d.attrs.style !== undefined ? +d.attrs.style : null; };
  const xfAt = (r, c) => cells.get(C.addr(r, c))?.xf ?? rowXf(r) ?? colXf(c) ?? 0;
  const merges = kids(kid(xml, 'mergeCells'), 'mergeCell').map(m => C.range(m.attrs.ref));
  const links = new Map(), otherLinks = [];
  for (const h of kids(kid(xml, 'hyperlinks'), 'hyperlink')) {
    const q = C.range(h.attrs.ref), target = rels[h.attrs['r:id']]?.target, a = C.addr(q.r1, q.c1);
    if (target && /^https?:\/\//i.test(target)) links.set(a, target);
    else { otherLinks.push(`${a}（${target || '#' + (h.attrs.location || '')}）`); noteCell(a, ['hyperlink', 'color', 'underline'], 'http(s) 以外のリンク'); }
  }
  if (otherLinks.length) say('link', `次のリンクは、http(s) のリンクでないので移せない: ${ranges(otherLinks)}`);

  // Formulas: kept when the page can compute them, otherwise the value Excel saved.
  let refRow = -1, refCol = -1;
  const valueOnly = [], valueOnlyNote = a => noteCell(a, ['hasFormula', 'formula'], 'HTMLで計算できない数式は値にした');
  for (const [a, cell] of cells) {
    // A data table (what-if analysis) has its <f>, without text, in its first cell only; Excel computes the whole range.
    if (cell.formulaKind === 'dataTable') {
      valueOnly.push(`${a}（データテーブル ${cell.ref || a}）`);
      C.each(C.range(cell.ref || a), (r, c) => valueOnlyNote(C.addr(r, c)));
      cell.formula = null;
      continue;
    }
    if (!cell.formula) continue;
    const support = cell.formulaKind === 'array' ? {ok: false, why: '配列数式'} : formulaSupport('=' + cell.formula);
    if (support.ok) { refRow = Math.max(refRow, support.maxRow); refCol = Math.max(refCol, support.maxCol); continue; }
    valueOnly.push(`${a}（=${cell.formula}：${support.why}）`);
    valueOnlyNote(a);
    cell.formula = null;
  }
  if (valueOnly.length) say('formula', `次の数式は、HTMLで計算できないので、Excelが保存していた値にした: ${ranges(valueOnly)}`);

  // The area: what has content decides the limit; cells that only carry a format may be cut at the limit.
  let lastRow = -1, lastCol = -1, styledRow = -1, styledCol = -1;
  const pictures = readPictures(parts, book, rels, xml, name, report, notes);
  for (const [, cell] of cells) {
    const content = cell.formula || cell.value && !(cell.value.t === 's' && cell.value.v === '');
    if (content) { lastRow = Math.max(lastRow, cell.r); lastCol = Math.max(lastCol, cell.c); }
    else if (cell.xf) { styledRow = Math.max(styledRow, cell.r); styledCol = Math.max(styledCol, cell.c); }
  }
  for (const m of merges) { lastRow = Math.max(lastRow, m.r2); lastCol = Math.max(lastCol, m.c2); }
  for (const p of pictures) { lastRow = Math.max(lastRow, p.r); lastCol = Math.max(lastCol, p.c); }
  for (const a of links.keys()) { const [r, c] = C.coord(a); lastRow = Math.max(lastRow, r); lastCol = Math.max(lastCol, c); }
  if (lastRow >= LIMIT.rows || lastCol >= LIMIT.cols) return {tooLarge: {name, rows: lastRow + 1, cols: lastCol + 1}};
  const rows = Math.min(LIMIT.rows, Math.max(30, lastRow + 21, styledRow + 1, refRow + 1));
  const cols = Math.min(LIMIT.cols, Math.max(10, lastCol + 4, styledCol + 1, refCol + 1));
  if (styledRow >= rows || styledCol >= cols) say('cut', `書式だけのセルのうち、HTMLの上限（500行・100列）の外にあるものは読み込まなかった`);

  // The height of a row that fits the Normal font (for a font of unknown height, the file's default height).
  const sheetFormat = kid(xml, 'sheetFormatPr')?.attrs || {};
  const fitted = book.normalRow.pt ?? (+sheetFormat.defaultRowHeight || EXCEL_ROW_PT);

  // Cells. A merged cell takes its format from its top-left cell, and its outline from the cells on each edge.
  const out = {}, base = {...C.DEFAULT, ...book.defaultStyle};
  const mergeOf = (r, c) => merges.find(m => C.within(r, c, m));
  const hiddenBy = (r, c) => { const m = mergeOf(r, c); return m && (r !== m.r1 || c !== m.c1); };
  const edge = (r, c, key) => styleOf(book.styles, xfAt(r, c)).st[key];
  const formatNotes = new Map(), hiddenValues = [];
  // Per row, for the rows Excel fits to their text: the height of the tallest text, and whether it is approximate.
  const rowFit = new Map(), roughRows = new Set();
  const FONT_FIELDS = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color'];
  /** Puts the cell at (r, c), with the cell format xfIndex, into out; cell is its <c> as read (none for an empty cell). */
  const put = (a, r, c, xfIndex, cell) => {
    const {st: full, notes: styleNotes, format, size} = styleOf(book.styles, xfIndex);
    const st = {...full}, merge = mergeOf(r, c);
    if (merge) {
      st.borderRight = edge(merge.r1, merge.c2, 'borderRight');
      st.borderBottom = edge(merge.r2, merge.c1, 'borderBottom');
    }
    for (const n of styleNotes) {
      if (n.field) noteCell(a, [n.field], n.reason);
      // In Excel a border is also the neighbour's opposite side (the top of a cell is the bottom of the cell above).
      const side = NEIGHBOUR_SIDE[n.field];
      if (side && r + side[0] >= 0 && c + side[1] >= 0) noteCell(C.addr(r + side[0], c + side[1]), [side[2]], n.reason);
      say('style:' + n.reason, n.reason + '（' + a + ' など）');
    }
    let value = cell?.value;
    const formula = cell?.formula;
    if (book.date1904 && value?.t === 'n' && st.numFmt === 'yyyy/mm/dd') value = {...value, v: value.v + 1462};
    // A format shows numbers, and only a format with an @ shows text, otherwise than written.
    const shown = formula || value && !(value.t === 's' && value.v === '');
    if (!format.exact && shown && (value?.t === 'n' || formula || format.code.includes('@'))) {
      noteCell(a, ['numfmt', 'text'], `表示形式「${format.code}」は「${format.fmt}」にした`);
      const key = format.code + '→' + format.fmt;
      formatNotes.set(key, [...(formatNotes.get(key) || []), a]);
    } else if (shown && format.code !== (EXCEL_CODE[format.fmt] || format.fmt) && !(format.fmt === 'General' && format.exact)) {
      noteCell(a, ['numfmt'], `表示形式「${format.code}」は、同じ表示の「${format.fmt}」にした`); // shown alike, written otherwise
    }
    if (value?.rich) { noteCell(a, FONT_FIELDS, '文字ごとの書式'); say('rich', `文字ごとに違う書式（一部だけ太字など）は、セル全体の書式にした（${a} など）`); }
    // Runs in one format show the text in it; Excel then gives no single font for the cell (it reads as mixed).
    let fontSize = size;
    if (value?.font) {
      const {size: runSize, ...rest} = value.font;
      Object.assign(st, rest);
      if (runSize !== undefined) { st.size = Math.min(96, Math.max(6, runSize)); fontSize = runSize; }
      noteCell(a, FONT_FIELDS, '文字の書式（run）をセルの書式にした');
    }
    // A row without a height of its own fits its tallest text (merged cells over several rows do not count, as in Excel).
    // Known fonts of 11pt give the height Excel gives; other sizes and fonts, approximated formats and wrapped text make
    // the height approximate.
    if (shown && !(merge && merge.r2 > merge.r1)) {
      const known = ROW_PT_11[String(st.font).toLowerCase()];
      const perPt = known !== undefined ? known / 11 : st.font === book.normalRow.font ? fitted / book.normalRow.size : EXCEL_ROW_PT / 11;
      rowFit.set(r, Math.max(rowFit.get(r) ?? 0, perPt * fontSize));
      if (known === undefined || fontSize !== 11 || styleNotes.length || value?.rich || st.wrap) roughRows.add(r);
    }
    const entry = {};
    if (formula) { entry.t = 'f'; entry.f = '=' + formula; }
    else if (!value || value.t === 's' && value.v === '') { entry.t = 's'; entry.v = ''; }
    else if (value.t === 'e') { entry.t = 's'; entry.v = value.v; noteCell(a, ['value2'], 'エラーの値は、その文字にした'); say('error', `エラーの値（#N/A など）は、その文字の文字列にした（${a} など）`); }
    else { entry.t = value.t; entry.v = value.v; }
    if (links.has(a)) entry.link = links.get(a);
    const diff = Object.fromEntries(Object.entries(st).filter(([k, v]) => base[k] !== v));
    if (Object.keys(diff).length) entry.s = diff;
    if (entry.t === 's' && entry.v === '' && !entry.s && !entry.link) return;
    out[a] = entry;
  };
  for (const [a, cell] of cells) {
    if (cell.r >= rows || cell.c >= cols) continue;
    if (hiddenBy(cell.r, cell.c)) { if (cell.formula || cell.value && cell.value.v !== '') hiddenValues.push(a); continue; }
    put(a, cell.r, cell.c, cell.xf, cell);
  }
  // Cells without a <c> in a formatted row or column take that format.
  const rowXfs = Array.from({length: rows}, (_, r) => rowXf(r)), colXfs = Array.from({length: cols}, (_, c) => colXf(c));
  if (rowXfs.some(x => x !== null) || colXfs.some(x => x !== null)) {
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const xfIndex = rowXfs[r] ?? colXfs[c], a = C.addr(r, c);
      if (xfIndex !== null && !cells.has(a) && !hiddenBy(r, c)) put(a, r, c, xfIndex, null);
    }
  }
  if (hiddenValues.length) say('hidden', `結合で隠れたセル（${ranges(hiddenValues)}）の値は、Excelでも表示されないので読み込まなかった`);
  for (const [key, list] of formatNotes) { const [code, fmt] = key.split('→'); say('format:' + key, `表示形式「${code}」はHTMLにないので「${fmt}」にした: ${ranges(list)}`); }

  // Sizes: px from Excel's column width (characters of the Normal font's widest digit) and row height (pt).
  // A column without a width of its own is Excel's default: baseColWidth digits and the padding, rounded up to a
  // multiple of 8 screen pixels (72px for Yu Gothic 11pt at 96dpi; 88 pixels at 120dpi, which is 70.4px there).
  const digit = book.digitPx;
  const defaultWidth = sheetFormat.defaultColWidth ? +sheetFormat.defaultColWidth * digit
    : Math.ceil(((+(sheetFormat.baseColWidth || 8)) * digit + 2 * Math.ceil(digit / 4) + 1) / 8) * 8;
  const colWidths = [], hiddenCols = [], clampCols = [], defaultCols = [];
  for (let c = 0; c < cols; c++) {
    const def = colDef(c);
    if (!def || def.attrs.width === undefined) defaultCols.push(colLetters(c));
    let px = def && def.attrs.width !== undefined ? +def.attrs.width * digit : defaultWidth;
    if (def && (def.attrs.hidden === '1' || def.attrs.hidden === 'true' || +def.attrs.width === 0)) { hiddenCols.push(colLetters(c)); px = defaultWidth; }
    const clamped = Math.min(C.SIZE_LIMITS.colWidth[1], Math.max(C.SIZE_LIMITS.colWidth[0], Math.round(px)));
    if (clamped !== Math.round(px) && !hiddenCols.includes(colLetters(c))) clampCols.push(colLetters(c));
    colWidths.push(clamped);
  }
  // A row without a height of its own (desktop Excel 16.0, measured): the sheet's default height when it is marked as set
  // (customHeight); otherwise Excel fits the row to the Normal font, whatever height the file gives, and in a sheet
  // without a default height an empty row is Excel's own height.
  const setHeight = (sheetFormat.customHeight === '1' || sheetFormat.customHeight === 'true') && sheetFormat.defaultRowHeight ? +sheetFormat.defaultRowHeight : null;
  const filled = new Set([...cells.values()].map(cell => cell.r));
  const defaultPt = r => setHeight ?? (filled.has(r) || kid(xml, 'sheetFormatPr') ? fitted : EXCEL_ROW_PT);
  if (setHeight === null && !book.normalRow.exact) notes.sheet.push({item: 'rowHeightAll', reason: '標準のフォントに合う行の高さは近い値'});
  const rowHeights = [], hiddenRows = [], clampRows = [], fitRows = [];
  for (let r = 0; r < rows; r++) {
    const attrs = rowInfo.get(r) || {}, hidden = attrs.hidden === '1' || attrs.hidden === 'true' || attrs.ht !== undefined && +attrs.ht === 0;
    const fit = setHeight === null && attrs.ht === undefined && rowFit.has(r); // Excel fits the row to its text
    let px = (attrs.ht !== undefined ? +attrs.ht : fit ? rowFit.get(r) : defaultPt(r)) / 0.75;
    if (fit && roughRows.has(r) && !hidden) fitRows.push(r + 1);
    if (hidden) { hiddenRows.push(r + 1); px = defaultPt(r) / 0.75; }
    const height = Math.min(C.SIZE_LIMITS.rowHeight[1], Math.max(C.SIZE_LIMITS.rowHeight[0], Math.round(px)));
    if (height !== Math.round(px) && !hidden) clampRows.push(r + 1);
    rowHeights.push(height);
  }
  if (hiddenCols.length) { say('hiddenCols', `非表示の列（${hiddenCols.join('、')}）は、HTMLで隠せないので表示した`); notes.sheet.push({item: 'colWidth', cols: hiddenCols, reason: '非表示の列'}); }
  if (defaultCols.length) notes.sheet.push({item: 'colWidth', cols: defaultCols, reason: 'Excelの既定の列幅は、表示の倍率で少し変わる（96dpiで72px、120dpiで70.4px相当）'});
  if (clampCols.length) { say('clampCols', `列（${clampCols.join('、')}）の幅は、HTMLの範囲（24〜2044px）に収めた`); notes.sheet.push({item: 'colWidth', cols: clampCols, reason: '幅の範囲'}); }
  if (hiddenRows.length) { say('hiddenRows', `非表示の行（${ranges(hiddenRows.map(String))} 行目）は、HTMLで隠せないので表示した`); notes.sheet.push({item: 'rowHeight', rows: hiddenRows, reason: '非表示の行'}); }
  if (clampRows.length) { say('clampRows', `行（${ranges(clampRows.map(String))} 行目）の高さは、HTMLの範囲（16〜546px）に収めた`); notes.sheet.push({item: 'rowHeight', rows: clampRows, reason: '高さの範囲'}); }
  if (fitRows.length) notes.sheet.push({item: 'rowHeight', rows: fitRows, reason: 'Excelが中の文字に合わせて決める高さは、近い値'});

  // View: frozen panes, gridlines, zoom, selection.
  const view = kid(kid(xml, 'sheetViews'), 'sheetView'), pane = kid(view, 'pane'), out2 = {zoom: +(view?.attrs.zoomScale || 100)};
  let freeze;
  if (pane && /^frozen/.test(pane.attrs.state || '')) {
    freeze = {rows: Math.min(rows, Math.round(+(pane.attrs.ySplit || 0))), cols: Math.min(cols, Math.round(+(pane.attrs.xSplit || 0)))};
    if (view.attrs.topLeftCell && view.attrs.topLeftCell !== 'A1') say('freeze-scroll', '固定した行・列は、表の先頭からの固定にした（Excelでは途中から固定していた）');
    if (merges.some(m => m.r1 < freeze.rows && m.r2 >= freeze.rows || m.c1 < freeze.cols && m.c2 >= freeze.cols)) {
      say('freeze-merge', '固定の境目をまたぐ結合セルがあるので、枠の固定は移せなかった');
      notes.sheet.push({item: 'freeze', reason: '固定の境目をまたぐ結合'});
      freeze = undefined;
    }
  } else if (pane) { say('split', 'ウィンドウの分割は移せない'); }
  const selection = kid(view, 'selection');
  if (selection?.attrs.activeCell) {
    const [r, c] = C.coord(selection.attrs.activeCell);
    if (r < rows && c < cols) out2.active = [r, c];
    const first = (selection.attrs.sqref || '').split(' ')[0];
    if (first) { const q = C.range(first); if (q.r2 < rows && q.c2 < cols) out2.selection = q; }
  }
  for (const [tag, label] of [['conditionalFormatting', '条件付き書式'], ['dataValidations', '入力規則'], ['autoFilter', 'フィルター'], ['tableParts', 'テーブル'],
    ['sheetProtection', 'シートの保護'], ['legacyDrawing', 'コメント（メモ）'], ['rowBreaks', '改ページ'], ['colBreaks', '改ページ'], ['controls', 'フォームのコントロール']]) {
    if (kid(xml, tag)) say('drop:' + label, `${label}は移せない`);
  }
  // What Excel 2010 and later keep in the sheet's extensions.
  const EXT = {sparklineGroups: 'スパークライン', conditionalFormattings: '条件付き書式', dataValidations: '入力規則', slicerList: 'スライサー', timelineRefs: 'タイムライン'};
  for (const ext of kids(kid(xml, 'extLst'), 'ext')) for (const el of kids(ext)) { const label = EXT[local(el.name)]; if (label) say('drop:' + label, `${label}は移せない`); }
  if (Object.values(rels).some(r => /comments|threadedComment/.test(r.type))) say('drop:コメント（メモ）', 'コメント（メモ）は移せない');
  if (Object.values(rels).some(r => r.type === 'pivotTable')) say('drop:ピボットテーブル', 'ピボットテーブルは移せない（表示されていた値だけを移した）');
  if (info.hidden) say('hidden-sheet', 'Excelでは非表示のシートだが、HTMLでは表示した');

  const sheet = {id: 's' + (index + 1), name, rows, cols, colWidths, rowHeights, merges: merges.filter(m => m.r2 < rows && m.c2 < cols), cells: out};
  if (freeze && (freeze.rows || freeze.cols)) sheet.freeze = freeze;
  if (view?.attrs.showGridLines === '0' || view?.attrs.showGridLines === 'false') sheet.gridlines = false;
  if (pictures.length) sheet.images = pictures.filter(p => p.r < rows && p.c < cols);
  return {sheet, notes, view: out2, tabSelected: view?.attrs.tabSelected === '1'};
}

function readPictures(parts, book, rels, xml, sheetName, report, notes) {
  const drawing = kid(xml, 'drawing'), out = [];
  if (!drawing) return out;
  const target = rels[drawing.attrs['r:id']]?.target;
  if (!target || !parts.has(target)) return out;
  const dxml = parseXml(parts.get(target).toString('utf8')), drels = relationships(parts, target);
  const skipped = {format: 0, small: 0, large: 0};
  let other = 0;
  for (const item of kids(dxml)) {
    // A drawing written for several versions of Office is wrapped in mc:AlternateContent; the first choice is the one shown.
    const anchor = local(item.name) === 'AlternateContent' ? kids(kid(item, 'Choice') || kid(item, 'Fallback'))[0] : item;
    const kind = anchor ? local(anchor.name) : '';
    if (!/Anchor$/.test(kind)) continue;
    const pic = kid(anchor, 'pic');
    if (!pic) { other++; continue; }
    const embed = kid(kid(pic, 'blipFill'), 'blip')?.attrs['r:embed'], media = drels[embed]?.target;
    const ext = (media || '').replace(/^.*\./, '').toLowerCase(), type = {png: 'png', jpg: 'jpeg', jpeg: 'jpeg', gif: 'gif', webp: 'webp'}[ext];
    if (!media || !parts.has(media) || !type) { skipped.format++; continue; }
    const from = kid(anchor, 'from'), num = (el, n) => +textOf(kid(el, n)) || 0;
    const xfrm = kid(kid(pic, 'spPr'), 'xfrm'), size = kid(xfrm, 'ext') || kid(anchor, 'ext');
    let r = 0, c = 0, x = 0, y = 0;
    if (from) { c = num(from, 'col'); r = num(from, 'row'); x = num(from, 'colOff') / EMU_PER_PX; y = num(from, 'rowOff') / EMU_PER_PX; }
    else { const pos = kid(anchor, 'pos'); x = +(pos?.attrs.x || 0) / EMU_PER_PX; y = +(pos?.attrs.y || 0) / EMU_PER_PX; }
    const w = size ? +size.attrs.cx / EMU_PER_PX : 0, h = size ? +size.attrs.cy / EMU_PER_PX : 0;
    if (!(w >= 4 && h >= 4)) { skipped.small++; continue; }
    let asset = book.mediaAssets.get(media);
    if (!asset) {
      const data = 'data:image/' + type + ';base64,' + parts.get(media).toString('base64');
      if (data.length > 15e6) { skipped.large++; continue; }
      asset = 'img-' + (book.mediaAssets.size + 1);
      book.mediaAssets.set(media, asset);
      book.assets[asset] = data;
    }
    if (kid(kid(pic, 'blipFill'), 'srcRect') && Object.keys(kid(kid(pic, 'blipFill'), 'srcRect').attrs).length) report.add(sheetName, 'crop', '画像の切り抜きは移せないので、元の画像全体を表示した');
    if (xfrm?.attrs.rot && +xfrm.attrs.rot) report.add(sheetName, 'rotate', '画像の回転は移せない');
    out.push({id: 'pic-' + (out.length + 1), asset, r, c, x: Math.round(Math.min(20000, x)), y: Math.round(Math.min(20000, y)), w: Math.round(Math.min(10000, w)), h: Math.round(Math.min(10000, h))});
  }
  const lost = {format: n => `画像 ${n} 個は、HTMLで扱えない形式（EMF・BMPなど）なので移せなかった`, small: n => `幅か高さが4px未満の画像 ${n} 個は移せなかった`,
    large: n => `HTMLに入れられる大きさ（15MB）を超える画像 ${n} 個は移せなかった`};
  for (const [kind, n] of Object.entries(skipped)) if (n) report.add(sheetName, 'pictures-' + kind, lost[kind](n));
  if (skipped.format + skipped.small + skipped.large) notes.sheet.push({item: 'pictures', reason: '移せない画像'});
  if (other) { report.add(sheetName, 'shapes', `図形・グラフなど ${other} 個は移せない`); }
  return out;
}

/** Collects warnings once per (sheet, key); sheet null for the whole book. */
class Report {
  constructor() { this.items = new Map(); }
  add(sheet, key, text) { const k = sheet + '\u0000' + key; if (!this.items.has(k)) this.items.set(k, sheet === null ? text : `シート「${sheet}」：${text}`); }
  get list() { return [...this.items.values()]; }
}

/** {workbook, warnings, notes} for the .xlsx file in buffer. name: the file's name, for the title and the id. */
function fromXlsx(buffer, {name = 'book.xlsx'} = {}) {
  const parts = unzip(buffer);
  const rootRels = relationships(parts, ''), mainPath = Object.values(rootRels).find(r => r.type === 'officeDocument')?.target;
  if (!mainPath || !parts.has(mainPath)) throw Error('xlsx の中にブックがありません。');
  const wbXml = parseXml(parts.get(mainPath).toString('utf8')), rels = relationships(parts, mainPath);
  const byType = type => Object.values(rels).find(r => r.type === type)?.target;
  const book = {palette: [...PALETTE], theme: [...OFFICE_THEME], assets: {}, mediaAssets: new Map(), date1904: ['1', 'true'].includes(kid(wbXml, 'workbookPr')?.attrs.date1904)};
  const themePath = byType('theme');
  if (themePath && parts.has(themePath)) {
    const scheme = kid(kid(parseXml(parts.get(themePath).toString('utf8')), 'themeElements'), 'clrScheme');
    const get = n => { const el = kids(kid(scheme, n))[0]; return el ? (local(el.name) === 'sysClr' ? SYSTEM_COLOR[el.attrs.val] || el.attrs.lastClr : el.attrs.val) : null; };
    const colors = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(get);
    book.theme = colors.map((c, i) => c && /^[0-9A-Fa-f]{6}$/.test(c) ? c.toUpperCase() : OFFICE_THEME[i]);
  }
  const stylesPath = byType('styles');
  const stylesXml = stylesPath && parts.has(stylesPath) ? parseXml(parts.get(stylesPath).toString('utf8')) : {children: []};
  const indexed = kids(kid(kid(stylesXml, 'colors'), 'indexedColors'), 'rgbColor').map(c => (c.attrs.rgb || '').slice(-6).toUpperCase());
  if (indexed.length) book.palette = indexed;
  book.styles = readStyles(stylesXml, book);
  book.strings = sharedStrings(parts, byType('sharedStrings'), book);
  const normal = book.styles.normal.font;
  const digit11 = DIGIT_PX_11[String(normal.name).toLowerCase()];
  book.digitPx = Math.max(5, Math.round((digit11 ?? 8) * normal.size / 11));
  const row11 = ROW_PT_11[String(normal.name).toLowerCase()];
  book.normalRow = {pt: row11 === undefined ? null : row11 * normal.size / 11, exact: row11 !== undefined && normal.size === 11, font: normal.name, size: normal.size};
  book.defaultStyle = {font: normal.name, size: normal.size, valign: VERTICAL[book.styles.normal.vertical] || 'bottom'};

  // Only worksheets hold cells: chart sheets, dialog sheets and macro sheets are left out, and reported.
  const report = new Report();
  const allSheets = kids(kid(wbXml, 'sheets'), 'sheet').map((s, position) => ({position, name: s.attrs.name, rel: rels[s.attrs['r:id']], hidden: /hidden/i.test(s.attrs.state || '')}));
  const OTHER_SHEET = {chartsheet: 'グラフのシート', dialogsheet: 'ダイアログのシート', xlMacrosheet: 'マクロのシート', xlIntlMacrosheet: 'マクロのシート'};
  for (const s of allSheets) if (s.rel && s.rel.type !== 'worksheet') report.add(s.name, 'not-worksheet', `${OTHER_SHEET[s.rel.type] || 'ワークシートでないシート'}なので移せない`);
  const sheetInfos = allSheets.filter(s => s.rel?.type === 'worksheet' && parts.has(s.rel.target)).map(s => ({...s, path: s.rel.target}));
  if (!sheetInfos.length) throw Error('xlsx にワークシートがありません（グラフだけのシートは移せません）。');
  if (sheetInfos.length > LIMIT.sheets) throw Error(`シートが ${sheetInfos.length} 枚あり、HTMLの上限（30枚）を超えるので読み込めません。必要なシートを伝えてください。`);
  const results = sheetInfos.map((info, i) => readSheet(parts, book, info, i, report));
  const tooLarge = results.filter(r => r.tooLarge).map(r => r.tooLarge);
  if (tooLarge.length) throw Error('HTMLの上限（1シート500行・100列）を超えるので読み込めません: ' +
    tooLarge.map(t => `シート「${t.name}」は ${t.rows} 行・${t.cols} 列`).join('、') + '。読み込む範囲を指定し直してください。');
  // The sheet that was open: its position counts every sheet; one that was left out gives the next worksheet (or the last).
  const tab = +(kid(kid(wbXml, 'bookViews'), 'workbookView')?.attrs.activeTab || 0), next = sheetInfos.findIndex(s => s.position >= tab);
  const active = results[next < 0 ? results.length - 1 : next];
  const zoom = Math.round(active.view.zoom || 100), clampedZoom = Math.min(200, Math.max(50, zoom));
  if (clampedZoom !== zoom) report.add(active.sheet.name, 'zoom', `倍率 ${zoom}% は ${clampedZoom}% にした（HTMLは50〜200%）`);
  // The page has one zoom for the whole book: the zoom of the sheet that was open.
  for (const r of results) {
    const own = Math.round(r.view.zoom || 100);
    if (own === clampedZoom) continue;
    if (r === active) { r.notes.sheet.push({item: 'zoom', reason: 'HTMLの倍率は50〜200%'}); continue; } // reported above
    report.add(r.sheet.name, 'zoom-sheet', `倍率（${own}%）は、ブック全体の倍率（${clampedZoom}%）にした`);
    r.notes.sheet.push({item: 'zoom', reason: 'HTMLの倍率はブック全体で1つ'});
  }
  const coreXml = parts.has('docProps/core.xml') ? parseXml(parts.get('docProps/core.xml').toString('utf8')) : null;
  const stem = path.basename(name).replace(/\.[^.]+$/, '');
  const slug = stem.normalize('NFKC').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);
  const workbook = {
    schema: C.SCHEMA, id: ['xlsx', slug, crypto.createHash('sha256').update(buffer).digest('hex').slice(0, 8)].filter(Boolean).join('-'), revision: 'draft',
    title: textOf(kid(coreXml, 'title')).trim() || stem || 'Excel', activeSheet: active.sheet.id, defaultStyle: book.defaultStyle,
    view: {zoom: clampedZoom, ...(active.view.active ? {active: active.view.active} : {}), ...(active.view.selection ? {selection: active.view.selection} : {})},
    sheets: results.map(r => r.sheet),
  };
  if (Object.keys(book.assets).length) workbook.assets = book.assets;
  // Formulas kept as formulas: does the page compute what Excel saved?
  const cache = new Map(), differ = [];
  for (const [i, sheet] of workbook.sheets.entries()) {
    const info = sheetInfos[i], xml = parseXml(parts.get(info.path).toString('utf8'));
    for (const row of kids(kid(xml, 'sheetData'), 'row')) for (const c of kids(row, 'c')) {
      const a = c.attrs.r, cell = a && sheet.cells[a];
      if (!cell || cell.t !== 'f') continue;
      const saved = cellValue(c, book);
      if (!saved) continue;
      const [r, col] = C.coord(a), got = C.calculate(workbook, sheet, r, col, new Set(), cache);
      const same = saved.t === 'n' ? typeof got === 'number' && Math.abs(got - saved.v) <= 1e-9 * Math.max(1, Math.abs(saved.v)) : saved.t === 'e' ? got === saved.v : got === saved.v;
      if (!same) differ.push(`${sheet.name}!${a}（HTML ${JSON.stringify(got)}、Excel ${JSON.stringify(saved.v)}）`);
    }
  }
  if (differ.length) report.add(null, 'calc', `HTMLで計算した値が、Excelが保存していた値と違う数式がある: ${ranges(differ)}。HTMLの計算の仕組み（core.js）で直す候補になる`);
  C.validate(workbook);
  const notes = Object.fromEntries(results.map(r => [r.sheet.name, r.notes]));
  // Only Yu Gothic 11pt has the same digit width at every display scaling (8px at 96dpi, 10px at 120dpi).
  if (!['游ゴシック', 'yu gothic'].includes(String(normal.name).toLowerCase()) || normal.size !== 11)
    for (const n of Object.values(notes)) n.sheet.push({item: 'colWidthAll', reason: '標準のフォントが游ゴシック11ptでないので、列幅は近い値'});
  return {workbook, warnings: report.list, notes};
}

module.exports = {fromXlsx, mapFormat, tint};

if (require.main === module) {
  const file = process.argv[2];
  try {
    if (!file) throw Error('使い方: node scripts/read_xlsx.cjs <file.xlsx>');
    process.stdout.write(JSON.stringify(fromXlsx(fs.readFileSync(file), {name: path.basename(file)})));
  } catch (e) {
    process.stderr.write(e.message + '\n');
    process.exitCode = 1;
  }
}
