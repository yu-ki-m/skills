#!/usr/bin/env node
// Writes a workbook (excel-html-workbook/v1) as an .xlsx file that desktop Excel opens as the same table.
//
//   node scripts/xlsx.cjs < workbook.json > 表.xlsx        (scripts/build.py --xlsx runs this)
//
// The parts are written the way desktop Excel 16.0 (build 20326) saves the same table when it is built inside Excel
// (measured: the paste checker's native builder made the table, and Excel saved it). Formula values are the ones the
// page computes (assets/core.js); Excel recalculates every formula when it opens the file. Only Node.js built-ins.
'use strict';
const path = require('path');
const zlib = require('zlib');
const C = require(path.join(__dirname, '..', 'assets', 'core.js'));

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const TYPE = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  sheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  strings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
};
// Excel's built-in number formats, by id. The other two formats are defined in the file, spelt as Excel writes them.
const BUILTIN_FORMAT = {'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '@': 49};
const CUSTOM_FORMAT = {'yyyy/mm/dd': 'yyyy/mm/dd', '¥#,##0': '"¥"#,##0'};
const FIRST_CUSTOM_ID = 176; // Japanese Excel numbers its own custom formats from 176 on.
const BORDER = {'0.5px solid': 'hair', '1px solid': 'thin', '2px solid': 'medium', '3px solid': 'thick', '1px dashed': 'dashed',
  '2px dashed': 'mediumDashed', '1px dotted': 'dotted', '3px double': 'double'};
// Japanese Excel names these fonts in Japanese; the English names (the page's) are the same fonts.
const JAPANESE_NAME = Object.fromEntries(Object.entries(C.FONT_ALIAS).map(([japanese, english]) => [english, japanese]));
// The Normal style's font. Column widths are counted in its widest digit, which is 8px at 96dpi (10px at 120dpi), so a
// width of px / 8 gives the page's width at any display scaling. It stays Yu Gothic 11pt, as in Japanese Excel, whatever
// the book's default style is; cells carry their own fonts.
const NORMAL_FONT = {name: '游ゴシック', size: 11};
const DIGIT_PX = 8;
const DEFAULT_ROW_PX = C.height({rowHeights: []}, 0); // the page's row height when a row has none of its own (22px)
// A cell with a hyperlink gets Excel's Hyperlink cell style: an underlined font in the Office theme's hyperlink colour.
// Excel puts that font over the cell's own, so only the font name and size of the cell remain (as the native builder of
// verify_excel_paste.py makes it, and as Excel pastes a linked cell).
const LINK_COLOR = '0563C1';
const EMU_PER_PX = 9525; // a CSS pixel, 0.75pt
const PICTURE_TYPE = {png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp'};

// Text for XML. Characters XML 1.0 cannot hold (and CR, which XML readers turn into LF) are written as Excel writes them,
// _xHHHH_; a "_xHHHH_" already in the text has its "_" written as _x005F_, so that it is not read as such a character.
const text = s => C.esc(String(s).replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
  .replace(/[\x00-\x08\x0B-\x1F￾￿]/g, ch => '_x' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_'));
const attr = s => C.esc(String(s));
const num = n => String(Object.is(n, -0) ? 0 : n);
const hex = color => color.slice(1).toUpperCase();
const ref = q => q.r1 === q.r2 && q.c1 === q.c2 ? C.addr(q.r1, q.c1) : C.addr(q.r1, q.c1) + ':' + C.addr(q.r2, q.c2);

// The formula as Excel stores it: without "=", with function names and references in capitals (strings left alone).
// Returns null when the page's formula reader cannot read it, since Excel would refuse the whole file for it.
function excelFormula(f) {
  try { C.evalFormula(f, null, true); } catch { return null; }
  return f.slice(1).replace(/"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_.]*/g, t => t[0] === '"' ? t : t.toUpperCase());
}

/** Fonts, fills, borders, number formats and cell formats, each written once and referred to by number. */
class Styles {
  constructor() {
    this.fonts = [this.fontXml({...NORMAL_FONT, color: '#000000'}), this.fontXml({...NORMAL_FONT, color: '#' + LINK_COLOR, underline: true})];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.formats = new Map();
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>'];
  }
  static id(list, xml) { let i = list.indexOf(xml); if (i < 0) { i = list.length; list.push(xml); } return i; }
  fontXml({name, size, bold, italic, strike, underline, color}) {
    const japanese = C.fontCharset(name) === 128;
    return '<font>' + (bold ? '<b/>' : '') + (italic ? '<i/>' : '') + (strike ? '<strike/>' : '') + (underline ? '<u/>' : '') +
      `<sz val="${num(size)}"/><color rgb="FF${hex(color)}"/><name val="${attr(JAPANESE_NAME[name] || name)}"/>` +
      (japanese ? '<family val="3"/><charset val="128"/>' : '<family val="2"/>') + '</font>';
  }
  formatId(numFmt) {
    if (numFmt in BUILTIN_FORMAT) return BUILTIN_FORMAT[numFmt];
    if (!this.formats.has(numFmt)) this.formats.set(numFmt, FIRST_CUSTOM_ID + this.formats.size);
    return this.formats.get(numFmt);
  }
  /** The cell format number for style st (a full style), with edges (the border sides to draw), a link, a leading "'". */
  xf(st, edges, {link = false, quotePrefix = false} = {}) {
    const font = Styles.id(this.fonts, this.fontXml(link ? {name: st.font, size: st.size, underline: true, color: '#' + LINK_COLOR}
      : {name: st.font, size: st.size, bold: st.bold, italic: st.italic, strike: st.strike, underline: st.underline, color: st.color}));
    const fill = st.fill === 'none' ? 0 : Styles.id(this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${hex(st.fill)}"/><bgColor indexed="64"/></patternFill></fill>`);
    const side = (tag, v) => {
      if (v === 'none') return `<${tag}/>`;
      const [width, kind, color] = v.split(' ');
      return `<${tag} style="${BORDER[width + ' ' + kind]}"><color rgb="FF${hex(color)}"/></${tag}>`;
    };
    const border = Styles.id(this.borders, '<border>' + side('left', edges.borderLeft) + side('right', edges.borderRight) +
      side('top', edges.borderTop) + side('bottom', edges.borderBottom) + '<diagonal/></border>');
    const format = this.formatId(st.numFmt);
    // As in Excel, indenting a cell aligned "general" aligns it left.
    const horizontal = st.indent > 0 && st.align === 'general' ? 'left' : st.align;
    const align = [horizontal !== 'general' && `horizontal="${horizontal}"`, st.valign === 'middle' ? 'vertical="center"' : st.valign === 'top' && 'vertical="top"',
      st.wrap && 'wrapText="1"', st.indent > 0 && `indent="${st.indent}"`].filter(Boolean).join(' ');
    if (!format && !font && !fill && !border && !link && !quotePrefix && align === 'vertical="center"') return 0; // the Normal style itself
    const xml = `<xf numFmtId="${format}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="${link ? 1 : 0}"` + (quotePrefix ? ' quotePrefix="1"' : '') +
      (format ? ' applyNumberFormat="1"' : '') + (font ? ' applyFont="1"' : '') + (fill ? ' applyFill="1"' : '') + (border ? ' applyBorder="1"' : '') +
      (align ? ` applyAlignment="1"><alignment ${align}/></xf>` : '/>');
    return Styles.id(this.xfs, xml);
  }
  xml() {
    const list = (tag, items) => `<${tag} count="${items.length}">${items.join('')}</${tag}>`;
    const formats = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${attr(CUSTOM_FORMAT[code])}"/>`);
    return HEAD + `<styleSheet xmlns="${MAIN}">` + (formats.length ? list('numFmts', formats) : '') + list('fonts', this.fonts) + list('fills', this.fills) +
      list('borders', this.borders) +
      '<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"><alignment vertical="center"/></xf></cellStyleXfs>' +
      list('cellXfs', this.xfs) +
      '<cellStyles count="2"><cellStyle name="ハイパーリンク" xfId="1" builtinId="8"/><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
  }
}

/** Shared strings: each distinct string once. */
class Strings {
  constructor() { this.list = []; this.index = new Map(); this.count = 0; }
  id(s) { this.count++; if (!this.index.has(s)) { this.index.set(s, this.list.length); this.list.push(s); } return this.index.get(s); }
  xml() {
    return HEAD + `<sst xmlns="${MAIN}" count="${this.count}" uniqueCount="${this.list.length}">` +
      this.list.map(s => `<si><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${text(s)}</t></si>`).join('') + '</sst>';
  }
}

/**
 * The <c> element of the cell at (r, c), or '' when the cell has no value and Excel's Normal style is its format.
 * ctx holds what the whole book shares: styles, strings, warnings, and values (formula results, kept like the page's
 * renderCache so that every formula is computed once and the same way as on the page).
 */
function cellXml(w, s, r, c, ctx) {
  const {styles, strings, warnings} = ctx;
  const merge = C.mergeAt(s, r, c), owner = merge ? C.get(s, merge.r1, merge.c1) : C.get(s, r, c);
  const cell = merge && (r !== merge.r1 || c !== merge.c1) ? {t: 's', v: ''} : owner;
  const st = C.style(w, s, owner), a = C.addr(r, c);
  const edges = {borderTop: st.borderTop, borderRight: st.borderRight, borderBottom: st.borderBottom, borderLeft: st.borderLeft};
  // A merged cell keeps its outline: each cell in it draws only the sides that lie on the outline (as Excel stores it).
  if (merge) {
    if (r !== merge.r1) edges.borderTop = 'none';
    if (r !== merge.r2) edges.borderBottom = 'none';
    if (c !== merge.c1) edges.borderLeft = 'none';
    if (c !== merge.c2) edges.borderRight = 'none';
  }
  let formula = null, asText = null;
  if (cell.t === 'f') {
    formula = excelFormula(cell.f);
    if (formula === null) {
      asText = cell.f;
      warnings.push(`シート「${s.name}」の ${a}：数式「${cell.f}」は、HTMLの数式として読めないため、文字のまま文字列として書き出した`);
    }
  }
  const textValue = asText ?? (cell.t === 's' ? cell.v : null);
  const quotePrefix = textValue !== null && textValue !== '' && st.numFmt !== '@' && C.excelWouldConvert(textValue);
  const link = !!owner.link;
  const xf = styles.xf(st, edges, {link, quotePrefix}), at = `<c r="${a}"` + (xf ? ` s="${xf}"` : '');
  if (cell.t === 's' && cell.v === '' && !link && !xf) return '';
  if (textValue !== null) return textValue === '' ? `${at}/>` : `${at} t="s"><v>${strings.id(textValue)}</v></c>`;
  if (cell.t === 'n') return `${at}><v>${num(cell.v)}</v></c>`;
  if (cell.t === 'b') return `${at} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
  const f = `<f>${attr(formula)}</f>`, value = C.calculate(w, s, r, c, new Set(), ctx.values);
  if (typeof value === 'number') return `${at}>${f}<v>${num(value)}</v></c>`;
  if (typeof value === 'boolean') return `${at} t="b">${f}<v>${value ? 1 : 0}</v></c>`;
  if (value === '#CIRC!') return `${at}>${f}</c>`; // not an Excel value: Excel finds the circular reference itself
  if (C.isErr(value)) return `${at} t="e">${f}<v>${value}</v></c>`;
  return `${at} t="str">${f}<v>${text(value)}</v></c>`;
}

/** Where a picture's top-left corner lies: the cell and the offset inside it, in EMU. */
function anchorOf(s, p) {
  let {r, c, x, y} = p;
  while (c < s.cols - 1 && x >= C.width(s, c)) { x -= C.width(s, c); c++; }
  while (y >= C.height(s, r)) { y -= C.height(s, r); r++; }
  return {col: c, colOff: Math.round(x * EMU_PER_PX), row: r, rowOff: Math.round(y * EMU_PER_PX)};
}

function drawingXml(s, pictures) {
  return HEAD + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}">` +
    pictures.map(({picture: p, rel}, i) => {
      const at = anchorOf(s, p), cx = Math.round(p.w * EMU_PER_PX), cy = Math.round(p.h * EMU_PER_PX);
      let left = 0, top = 0;
      for (let c = 0; c < at.col; c++) left += C.width(s, c);
      for (let r = 0; r < at.row; r++) top += C.height(s, r);
      const x = Math.round(left * EMU_PER_PX) + at.colOff, y = Math.round(top * EMU_PER_PX) + at.rowOff;
      return `<xdr:oneCellAnchor><xdr:from><xdr:col>${at.col}</xdr:col><xdr:colOff>${at.colOff}</xdr:colOff><xdr:row>${at.row}</xdr:row><xdr:rowOff>${at.rowOff}</xdr:rowOff></xdr:from>` +
        `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="図 ${i + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip r:embed="${rel}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    }).join('') + '</xdr:wsDr>';
}

const rels = list => HEAD + `<Relationships xmlns="${PACKAGE_REL}">` + list.map(([id, type, target, external]) =>
  `<Relationship Id="${id}" Type="${type}" Target="${attr(target)}"${external ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>';

/** The worksheet part of sheet s (the index-th sheet), and the parts it refers to. */
function sheetParts(w, s, index, ctx) {
  const isActive = s.id === w.activeSheet, view = w.view || {}, zoom = Math.round(view.zoom || 100);
  const used = new Map(), mark = (r, c) => { if (!used.has(r)) used.set(r, new Set()); used.get(r).add(c); };
  for (const key of Object.keys(s.cells)) mark(...C.coord(key));
  for (const m of s.merges) C.each(m, mark);
  // The page shows a cell without a value in the default format (defaultStyle), and what is typed there takes it. When
  // that format is not Excel's Normal style, every cell of the table is written with it (cells outside stay Normal).
  if (JSON.stringify(C.style(w, s, {})) !== JSON.stringify(C.DEFAULT)) C.each({r1: 0, c1: 0, r2: s.rows - 1, c2: s.cols - 1}, mark);
  const rows = [];
  let box = null;
  for (let r = 0; r < s.rows; r++) {
    const cells = [...(used.get(r) || [])].sort((a, b) => a - b).map(c => {
      const xml = cellXml(w, s, r, c, ctx);
      if (xml) box = box ? C.rect([Math.min(box.r1, r), Math.min(box.c1, c)], [Math.max(box.r2, r), Math.max(box.c2, c)]) : C.rect([r, c]);
      return xml;
    }).join('');
    // Every written row gets its height as a set height. Excel fits a row without one to its text (18pt for Yu Gothic
    // 11pt), and ignores a height that is not marked as set (desktop Excel 16.0), so the page's 22px would be lost.
    const h = s.rowHeights[r];
    if (h == null && !cells) continue;
    rows.push(`<row r="${r + 1}" ht="${num((h ?? DEFAULT_ROW_PX) * 0.75)}" customHeight="1"` + (cells ? `>${cells}</row>` : '/>'));
  }
  // Columns: one <col> for each run of columns with the same width. A column without a width is the page's 96px.
  const cols = [];
  for (let c = 0; c < s.cols; c++) {
    const width = num(C.width(s, c) / DIGIT_PX), last = cols[cols.length - 1];
    if (last && last.width === width) last.max = c + 1; else cols.push({min: c + 1, max: c + 1, width});
  }
  const fr = s.freeze?.rows || 0, fc = s.freeze?.cols || 0, pane = fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : fc ? 'topRight' : null;
  let viewXml = (pane ? `<pane${fc ? ` xSplit="${fc}"` : ''}${fr ? ` ySplit="${fr}"` : ''} topLeftCell="${C.addr(fr, fc)}" activePane="${pane}" state="frozen"/>` : '');
  if (isActive && (view.active || view.selection)) {
    const active = view.active || [view.selection.r1, view.selection.c1];
    viewXml += `<selection${pane ? ` pane="${pane}"` : ''} activeCell="${C.addr(...active)}" sqref="${view.selection ? ref(view.selection) : C.addr(...active)}"/>`;
  } else if (pane) viewXml += `<selection pane="${pane}"/>`;
  const viewAttrs = (s.gridlines === false ? ' showGridLines="0"' : '') + (isActive ? ' tabSelected="1"' : '') + (zoom !== 100 ? ` zoomScale="${zoom}" zoomScaleNormal="${zoom}"` : '');
  const relList = [], links = [];
  for (const [key, cell] of Object.entries(s.cells)) {
    const [r, c] = C.coord(key), merge = C.mergeAt(s, r, c);
    if (!cell.link || merge && (merge.r1 !== r || merge.c1 !== c)) continue; // a cell hidden by a merge shows no link
    const id = 'rId' + (relList.length + 1);
    relList.push([id, REL + '/hyperlink', cell.link, true]);
    links.push(`<hyperlink ref="${key}" r:id="${id}"/>`);
  }
  const parts = [];
  let drawing = '';
  if (s.images?.length) {
    const drawingRels = [], pictures = s.images.map(p => {
      const target = '../media/' + ctx.media.file(p.asset);
      let rel = drawingRels.find(x => x[2] === target);
      if (!rel) drawingRels.push(rel = ['rId' + (drawingRels.length + 1), REL + '/image', target]);
      return {picture: p, rel: rel[0]};
    });
    const name = `xl/drawings/drawing${index + 1}.xml`, id = 'rId' + (relList.length + 1);
    relList.push([id, REL + '/drawing', `../drawings/drawing${index + 1}.xml`]);
    parts.push([name, drawingXml(s, pictures), TYPE.drawing], [`xl/drawings/_rels/drawing${index + 1}.xml.rels`, rels(drawingRels)]);
    drawing = `<drawing r:id="${id}"/>`;
  }
  // Printing fits every column on one page width, with as many pages down as needed: Excel's 拡大縮小印刷 →
  // 「すべての列を1ページに印刷」, written as Excel writes it. The paper is written too (A4 portrait, the default of
  // Japanese Excel), since Excel reads a page setup without a paper size as Letter (desktop Excel 16.0).
  const sheet = HEAD + `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="${box ? ref(box) : 'A1'}"/>` +
    `<sheetViews><sheetView${viewAttrs} workbookViewId="0"${viewXml ? `>${viewXml}</sheetView>` : '/>'}</sheetViews>` +
    // Rows without a height of their own are the page's 22px (16.5pt).
    `<sheetFormatPr defaultRowHeight="${num(DEFAULT_ROW_PX * 0.75)}" customHeight="1"/>` +
    `<cols>${cols.map(c => `<col min="${c.min}" max="${c.max}" width="${c.width}" customWidth="1"/>`).join('')}</cols>` +
    (rows.length ? `<sheetData>${rows.join('')}</sheetData>` : '<sheetData/>') +
    (s.merges.length ? `<mergeCells count="${s.merges.length}">${s.merges.map(m => `<mergeCell ref="${ref(m)}"/>`).join('')}</mergeCells>` : '') +
    (links.length ? `<hyperlinks>${links.join('')}</hyperlinks>` : '') +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<pageSetup paperSize="9" fitToHeight="0" orientation="portrait"/>' + drawing + '</worksheet>';
  parts.unshift([`xl/worksheets/sheet${index + 1}.xml`, sheet, TYPE.sheet]);
  if (relList.length) parts.push([`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, rels(relList)]);
  return parts;
}

/** Picture files: each asset used by a picture is stored once, as xl/media/image<n>.<type>. */
class Media {
  constructor(assets) { this.assets = assets || {}; this.files = new Map(); this.parts = []; }
  file(asset) {
    if (!this.files.has(asset)) {
      const m = /^data:image\/(png|jpeg|gif|webp);base64,(.*)$/.exec(this.assets[asset] || '');
      if (!m) throw Error('画像の中身がありません: ' + asset);
      const name = `image${this.files.size + 1}.${m[1]}`;
      this.files.set(asset, name);
      this.parts.push([`xl/media/${name}`, Buffer.from(m[2], 'base64')]);
    }
    return this.files.get(asset);
  }
}

// ZIP (stored or deflated, no zip64). The time stamp is fixed, so the same workbook always gives the same bytes.
const CRC_TABLE = Array.from({length: 256}, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
function crc32(data) { let c = 0xffffffff; for (const b of data) c = CRC_TABLE[(c ^ b) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
const DOS_DATE = (1 << 5) | 1, DOS_TIME = 0; // 1980-01-01 00:00
function zip(files) {
  const chunks = [], central = [];
  let offset = 0;
  for (const [name, content] of files) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8'), nameBytes = Buffer.from(name, 'utf8');
    const deflated = zlib.deflateRawSync(data), stored = deflated.length >= data.length, body = stored ? data : deflated, crc = crc32(data);
    const head = Buffer.alloc(30);
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt16LE(stored ? 0 : 8, 8); head.writeUInt16LE(DOS_TIME, 10); head.writeUInt16LE(DOS_DATE, 12);
    head.writeUInt32LE(crc, 14); head.writeUInt32LE(body.length, 18); head.writeUInt32LE(data.length, 22); head.writeUInt16LE(nameBytes.length, 26);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(stored ? 0 : 8, 10); entry.writeUInt16LE(DOS_TIME, 12);
    entry.writeUInt16LE(DOS_DATE, 14); entry.writeUInt32LE(crc, 16); entry.writeUInt32LE(body.length, 20); entry.writeUInt32LE(data.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28); entry.writeUInt32LE(offset, 42);
    chunks.push(head, nameBytes, body);
    central.push(entry, nameBytes);
    offset += head.length + nameBytes.length + body.length;
  }
  const directory = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10); end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, directory, end]);
}

/**
 * The .xlsx file of workbook w, as {bytes, warnings}. warnings lists what could not be written as it is in the page
 * (a formula the page cannot read is written as its text).
 */
function toXlsx(w) {
  C.validate(w);
  const styles = new Styles(), strings = new Strings(), media = new Media(w.assets), warnings = [];
  const ctx = {styles, strings, media, warnings, values: new Map()};
  const sheetFiles = w.sheets.map((s, i) => sheetParts(w, s, i, ctx));
  const n = w.sheets.length, active = Math.max(0, w.sheets.findIndex(s => s.id === w.activeSheet));
  const workbook = HEAD + `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><bookViews><workbookView${active ? ` activeTab="${active}"` : ''}/></bookViews><sheets>` +
    w.sheets.map((s, i) => `<sheet name="${attr(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
    // Excel recalculates every formula on opening, so its own results replace the page's wherever the two differ.
    '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
  const core = HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
    `xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${text(w.title)}</dc:title></cp:coreProperties>`;
  const parts = [
    ['xl/workbook.xml', workbook, TYPE.workbook],
    ['xl/_rels/workbook.xml.rels', rels([...w.sheets.map((_, i) => ['rId' + (i + 1), REL + '/worksheet', `worksheets/sheet${i + 1}.xml`]),
      ['rId' + (n + 1), REL + '/styles', 'styles.xml'], ['rId' + (n + 2), REL + '/sharedStrings', 'sharedStrings.xml']])],
    ...sheetFiles.flat(),
    ['xl/styles.xml', styles.xml(), TYPE.styles],
    ['xl/sharedStrings.xml', strings.xml(), TYPE.strings],
    ...media.parts,
    ['docProps/core.xml', core, TYPE.core],
    ['_rels/.rels', rels([['rId1', REL + '/officeDocument', 'xl/workbook.xml'], ['rId2', PACKAGE_REL + '/metadata/core-properties', 'docProps/core.xml']])],
  ];
  const kinds = [...new Set(media.parts.map(([name]) => name.replace(/^.*\./, '')))];
  const contentTypes = HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    kinds.map(k => `<Default Extension="${k}" ContentType="${PICTURE_TYPE[k]}"/>`).join('') +
    parts.filter(p => p[2]).map(([name, , type]) => `<Override PartName="/${name}" ContentType="${type}"/>`).join('') + '</Types>';
  return {bytes: zip([['[Content_Types].xml', contentTypes], ...parts.map(([name, data]) => [name, data])]), warnings};
}

module.exports = {toXlsx};

if (require.main === module) {
  let input = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', chunk => { input += chunk; });
  process.stdin.on('end', () => {
    try {
      const {bytes, warnings} = toXlsx(JSON.parse(input));
      for (const warning of warnings) process.stderr.write('注意: ' + warning + '\n');
      process.stdout.write(bytes);
    } catch (e) {
      process.stderr.write(e.message + '\n');
      process.exitCode = 1;
    }
  });
}
