// Reads an .xlsx file the way the tests look at it: the zip parts (each part's CRC is checked), and for each sheet its
// cells with resolved styles, column widths, row heights, merges, view, links and pictures.
// Only Node.js built-ins are used. It reads what scripts/xlsx.cjs writes; it is not a general xlsx reader.
// It also makes small hand-written .xlsx files (miniXlsx) for the tests of scripts/read_xlsx.cjs.
import {crc32, inflateRawSync} from 'node:zlib';

/** A zip (stored, not compressed) of parts: {name: text or Buffer}. */
export function makeZip(parts) {
  const chunks = [], central = [];
  let offset = 0;
  for (const [name, content] of Object.entries(parts)) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8'), nameBytes = Buffer.from(name, 'utf8'), crc = crc32(data);
    const head = Buffer.alloc(30), entry = Buffer.alloc(46);
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt32LE(crc, 14); head.writeUInt32LE(data.length, 18); head.writeUInt32LE(data.length, 22); head.writeUInt16LE(nameBytes.length, 26);
    entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt32LE(crc, 16); entry.writeUInt32LE(data.length, 20); entry.writeUInt32LE(data.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28); entry.writeUInt32LE(offset, 42);
    chunks.push(head, nameBytes, data);
    central.push(entry, nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(central.length / 2, 8); end.writeUInt16LE(central.length / 2, 10); end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, directory, end]);
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const escXml = s => String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
const MINI_STYLES = '<fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
  '<border><left/><right/><top/><bottom style="thin"/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"><alignment vertical="center"/></xf>' +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs>';

/** A .rels part: [[id, type (the last word of an Office relationship type, or a whole URI), target, external]]. */
export const relsXml = list => '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  list.map(([id, type, target, external]) => `<Relationship Id="${id}" Type="${type.includes('/') ? type : OFFICE_REL + type}" Target="${escXml(target)}"` +
    (external ? ' TargetMode="External"' : '') + '/>').join('') + '</Relationships>';

/**
 * The inside of <styleSheet>, in the order Excel requires. fonts, fills, borders: XML of each item (Excel's two required
 * fills come first by default); xfs: the cell formats after cellXfs[0] (which is the Normal style, normal);
 * numFmts: {id: format code}; colors: a <colors> element.
 */
export function stylesXml({numFmts = {}, fonts = ['<font><sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font>'],
  fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'],
  normal = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf>', xfs = [], colors = ''}) {
  const list = (tag, items) => `<${tag} count="${items.length}">${items.join('')}</${tag}>`, codes = Object.entries(numFmts);
  return (codes.length ? list('numFmts', codes.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escXml(code)}"/>`)) : '') +
    list('fonts', fonts) + list('fills', fills) + list('borders', borders) + list('cellStyleXfs', [normal]) +
    list('cellXfs', [normal.replace('<xf ', '<xf xfId="0" '), ...xfs]) + '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>' + colors;
}

/**
 * A hand-written .xlsx file for the tests of scripts/read_xlsx.cjs.
 * sheets: [{name, xml: the inside of <worksheet>, rels: [[id, type, target, external]], state: 'hidden', type: 'worksheet'}];
 * strings: the <si> items, styles: the inside of <styleSheet> (null: no such part); workbook: elements put into
 * workbook.xml before <sheets> (<workbookPr>, <bookViews>); theme: the inside of <a:clrScheme> (null: no theme part);
 * core: the inside of docProps/core.xml (null: no such part); parts: more parts by name (text or Buffer).
 */
export function handXlsx({sheets, strings = '', styles = MINI_STYLES, workbook = '', theme = null, core = null, parts = {}}) {
  const bookRels = sheets.map((s, i) => ['rId' + (i + 1), s.type || 'worksheet', `worksheets/sheet${i + 1}.xml`]);
  const files = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/>' +
      '<Default Extension="gif" ContentType="image/gif"/><Default Extension="emf" ContentType="image/x-emf"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      (styles === null ? '' : '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>') +
      (strings === null ? '' : '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>') +
      (theme === null ? '' : '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>') +
      Object.keys(parts).filter(name => name.startsWith('xl/drawings/drawing')).map(name => `<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('') +
      (core === null ? '' : '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>') + '</Types>',
    '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml'],
      ...(core === null ? [] : [['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml']])]),
    'xl/workbook.xml': `<workbook ${NS}>${workbook}<sheets>` +
      sheets.map((s, i) => `<sheet name="${escXml(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>',
  };
  sheets.forEach((s, i) => {
    const root = s.type === 'chartsheet' ? 'chartsheet' : 'worksheet';
    files[`xl/worksheets/sheet${i + 1}.xml`] = `<${root} ${NS}>${s.xml}</${root}>`;
    if (s.rels) files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = relsXml(s.rels);
  });
  const add = (type, target, content) => { bookRels.push(['rId' + (bookRels.length + 1), type, target]); files['xl/' + target] = content; };
  if (styles !== null) add('styles', 'styles.xml', `<styleSheet ${NS}>${styles}</styleSheet>`);
  if (strings !== null) add('sharedStrings', 'sharedStrings.xml', `<sst ${NS}>${strings}</sst>`);
  // A theme needs its fonts and formats as well for Excel to open the file.
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>', fills = fill.repeat(3), font = name => `<a:latin typeface="${name}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  if (theme !== null) add('theme', 'theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>' +
    `<a:clrScheme name="T">${theme}</a:clrScheme><a:fontScheme name="T"><a:majorFont>${font('Yu Gothic Light')}</a:majorFont><a:minorFont>${font('Yu Gothic')}</a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="T"><a:fillStyleLst>${fills}</a:fillStyleLst><a:lnStyleLst>${[6350, 12700, 19050].map(w => `<a:ln w="${w}">${fill}</a:ln>`).join('')}</a:lnStyleLst>` +
    `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${fills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`);
  files['xl/_rels/workbook.xml.rels'] = relsXml(bookRels);
  if (core !== null) files['docProps/core.xml'] = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    `xmlns:dc="http://purl.org/dc/elements/1.1/">${core}</cp:coreProperties>`;
  return makeZip({...files, ...parts});
}

/**
 * A small hand-written .xlsx with one sheet named S. sheet: the inside of <worksheet> (with <sheetData>); strings: the
 * <si> items; workbookPr: an element put into workbook.xml. Cell styles: s="1" has a thin bottom border, s="2" format 14.
 */
export const miniXlsx = ({sheet, strings = '', workbookPr = '', styles = MINI_STYLES}) => handXlsx({sheets: [{name: 'S', xml: sheet}], strings, styles, workbook: workbookPr});

/** The parts of a zip file as a Map of name -> Buffer. Throws when a part's CRC-32 does not match its data. */
export function unzip(buffer) {
  let end = buffer.length - 22;
  while (end >= 0 && buffer.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) throw new Error('ZIPの終わりの記録がありません');
  const count = buffer.readUInt16LE(end + 10), parts = new Map();
  let p = buffer.readUInt32LE(end + 16);
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(p) !== 0x02014b50) throw new Error('ZIPの目次が壊れています');
    const method = buffer.readUInt16LE(p + 10), crc = buffer.readUInt32LE(p + 16), size = buffer.readUInt32LE(p + 20);
    const nameLength = buffer.readUInt16LE(p + 28), extra = buffer.readUInt16LE(p + 30), comment = buffer.readUInt16LE(p + 32);
    const local = buffer.readUInt32LE(p + 42), name = buffer.toString('utf8', p + 46, p + 46 + nameLength);
    if (buffer.readUInt32LE(local) !== 0x04034b50) throw new Error('ZIPの部品の見出しが壊れています: ' + name);
    const start = local + 30 + buffer.readUInt16LE(local + 26) + buffer.readUInt16LE(local + 28);
    const raw = buffer.subarray(start, start + size), data = method === 8 ? inflateRawSync(raw) : Buffer.from(raw);
    if (method !== 0 && method !== 8) throw new Error('未対応の圧縮方式: ' + method);
    if (data.length !== buffer.readUInt32LE(p + 24)) throw new Error('展開後の大きさが合いません: ' + name);
    if (crc32(data) !== crc) throw new Error('CRCが合いません: ' + name);
    if (parts.has(name)) throw new Error('同じ名前の部品が2つあります: ' + name);
    parts.set(name, data);
    p += 46 + nameLength + extra + comment;
  }
  return parts;
}

const ENTITY = {lt: '<', gt: '>', amp: '&', quot: '"', apos: "'"};
const decode = text => text.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi,
  (_, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENTITY[e]);

/** A small XML parser: elements as {name, attrs, children}, text as strings. No DTDs; comments and <?…?> are skipped. */
export function parseXml(text) {
  const root = {name: '#root', attrs: {}, children: []}, stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(text))) {
    const top = stack[stack.length - 1];
    if (m[1]) {
      if (top.name !== m[1]) throw new Error(`XMLの閉じタグが合いません: ${m[1]}（開いているのは ${top.name}）`);
      stack.pop();
    } else if (m[2]) {
      const attrs = {};
      for (const a of m[3].matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[a[1]] = decode(a[2] ?? a[3]);
      const element = {name: m[2], attrs, children: []};
      top.children.push(element);
      if (!m[4]) stack.push(element);
    } else if (m[5] !== undefined && !m[0].startsWith('<')) {
      top.children.push(decode(m[5]));
    }
  }
  if (stack.length !== 1) throw new Error('XMLの要素が閉じていません: ' + stack[stack.length - 1].name);
  const elements = root.children.filter(c => typeof c !== 'string');
  if (elements.length !== 1) throw new Error('XMLの最上位の要素は1つのはずです');
  return elements[0];
}

const local = name => name.replace(/^.*:/, '');
/** Child elements of el named name (namespace prefix ignored). */
export const kids = (el, name) => (el?.children || []).filter(c => typeof c !== 'string' && (!name || local(c.name) === name));
/** The first child element named name. */
export const kid = (el, name) => kids(el, name)[0];
/** The text inside el. */
export const textOf = el => (el?.children || []).map(c => typeof c === 'string' ? c : textOf(c)).join('');

/** Excel writes characters that XML cannot hold as _xHHHH_; a literal "_x0041_" in the text is written as "_x005F_x0041_". */
const unescapeX = text => text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));

// Excel's built-in number formats used by this skill.
const BUILTIN_FORMATS = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 49: '@'};

function resolve(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const piece of target.split('/')) piece === '..' ? parts.pop() : parts.push(piece);
  return parts.join('/');
}

function relationships(parts, partName) {
  const name = partName ? partName.replace(/([^/]+)$/, '_rels/$1.rels') : '_rels/.rels', out = {};
  if (!parts.has(name)) return out;
  for (const r of kids(parseXml(parts.get(name).toString('utf8')), 'Relationship'))
    out[r.attrs.Id] = {type: r.attrs.Type.replace(/^.*\//, ''), target: r.attrs.TargetMode === 'External' ? r.attrs.Target : resolve(partName, r.attrs.Target), external: r.attrs.TargetMode === 'External'};
  return out;
}

function readStyles(xml) {
  const formats = {...BUILTIN_FORMATS};
  for (const f of kids(kid(xml, 'numFmts'), 'numFmt')) formats[f.attrs.numFmtId] = f.attrs.formatCode;
  const color = el => el?.attrs.rgb ? '#' + el.attrs.rgb.slice(2).toLowerCase() : el ? 'other' : undefined;
  const flag = (el, name) => { const e = kid(el, name); return !!e && e.attrs.val !== '0' && e.attrs.val !== 'false'; };
  const fonts = kids(kid(xml, 'fonts'), 'font').map(f => ({
    name: kid(f, 'name')?.attrs.val, size: +kid(f, 'sz')?.attrs.val, bold: flag(f, 'b'), italic: flag(f, 'i'), strike: flag(f, 'strike'),
    underline: kid(f, 'u') ? (kid(f, 'u').attrs.val || 'single') : 'none', color: color(kid(f, 'color')),
    family: kid(f, 'family') ? +kid(f, 'family').attrs.val : undefined, charset: kid(f, 'charset') ? +kid(f, 'charset').attrs.val : undefined,
  }));
  const fills = kids(kid(xml, 'fills'), 'fill').map(f => {
    const p = kid(f, 'patternFill');
    return p?.attrs.patternType === 'solid' ? color(kid(p, 'fgColor')) : p?.attrs.patternType || 'none';
  });
  const side = el => el && el.attrs.style ? {style: el.attrs.style, color: color(kid(el, 'color'))} : 'none';
  const borders = kids(kid(xml, 'borders'), 'border').map(b => ({top: side(kid(b, 'top')), bottom: side(kid(b, 'bottom')), left: side(kid(b, 'left')), right: side(kid(b, 'right'))}));
  const styleNames = {};
  for (const s of kids(kid(xml, 'cellStyles'), 'cellStyle')) styleNames[s.attrs.xfId] = {name: s.attrs.name, builtinId: +s.attrs.builtinId};
  const xf = x => {
    const a = kid(x, 'alignment')?.attrs || {};
    return {
      numFmt: formats[x.attrs.numFmtId ?? 0], font: fonts[+(x.attrs.fontId ?? 0)], fill: fills[+(x.attrs.fillId ?? 0)], borders: borders[+(x.attrs.borderId ?? 0)],
      alignment: {horizontal: a.horizontal || 'general', vertical: a.vertical || 'bottom', wrapText: a.wrapText === '1' || a.wrapText === 'true', indent: +(a.indent || 0)},
      quotePrefix: x.attrs.quotePrefix === '1' || x.attrs.quotePrefix === 'true', cellStyle: styleNames[x.attrs.xfId ?? 0]?.builtinId,
    };
  };
  return {formats, fonts, fills, borders, cellStyleXfs: kids(kid(xml, 'cellStyleXfs'), 'xf').map(xf), cellXfs: kids(kid(xml, 'cellXfs'), 'xf').map(xf), styleNames};
}

/** Column index (0-based) of an A1 address's letters. */
export const colIndex = letters => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

function readSheet(parts, path, name, styles, strings) {
  const xml = parseXml(parts.get(path).toString('utf8')), rels = relationships(parts, path);
  const cells = {}, rows = {}, cols = [];
  for (const row of kids(kid(xml, 'sheetData'), 'row')) {
    rows[+row.attrs.r] = {ht: row.attrs.ht === undefined ? undefined : +row.attrs.ht, customHeight: row.attrs.customHeight === '1'};
    for (const c of kids(row, 'c')) {
      const type = c.attrs.t || 'n', v = kid(c, 'v'), f = kid(c, 'f'), raw = v ? textOf(v) : undefined;
      let value;
      if (type === 's') value = strings[+raw];
      else if (type === 'b') value = raw === undefined ? undefined : raw === '1';
      else if (type === 'n') value = raw === undefined ? undefined : +raw;
      else if (type === 'inlineStr') value = unescapeX(textOf(kid(c, 'is')));
      else value = raw === undefined ? undefined : unescapeX(raw);
      if (cells[c.attrs.r]) throw new Error('同じセルが2回あります: ' + c.attrs.r);
      cells[c.attrs.r] = {type, value, formula: f ? textOf(f) : undefined, style: styles.cellXfs[+(c.attrs.s ?? 0)], styleIndex: +(c.attrs.s ?? 0)};
    }
  }
  for (const c of kids(kid(xml, 'cols'), 'col')) cols.push({min: +c.attrs.min, max: +c.attrs.max, width: +c.attrs.width, customWidth: c.attrs.customWidth === '1'});
  const format = kid(xml, 'sheetFormatPr')?.attrs || {};
  const viewEl = kid(kid(xml, 'sheetViews'), 'sheetView'), paneEl = kid(viewEl, 'pane');
  const view = {
    showGridLines: viewEl?.attrs.showGridLines !== '0', tabSelected: viewEl?.attrs.tabSelected === '1', zoomScale: viewEl?.attrs.zoomScale ? +viewEl.attrs.zoomScale : 100,
    pane: paneEl ? {xSplit: +(paneEl.attrs.xSplit || 0), ySplit: +(paneEl.attrs.ySplit || 0), topLeftCell: paneEl.attrs.topLeftCell, activePane: paneEl.attrs.activePane, state: paneEl.attrs.state} : null,
    selections: kids(viewEl, 'selection').map(s => ({pane: s.attrs.pane, activeCell: s.attrs.activeCell, sqref: s.attrs.sqref})),
  };
  const links = {};
  for (const h of kids(kid(xml, 'hyperlinks'), 'hyperlink')) links[h.attrs.ref] = rels[h.attrs['r:id']]?.target;
  const pictures = [];
  const drawing = kid(xml, 'drawing');
  if (drawing) {
    const drawingPath = rels[drawing.attrs['r:id']].target, drawingRels = relationships(parts, drawingPath);
    for (const anchor of kids(parseXml(parts.get(drawingPath).toString('utf8')))) {
      const from = kid(anchor, 'from'), ext = kid(anchor, 'ext'), pic = kid(anchor, 'pic');
      const num = (el, n) => +textOf(kid(el, n));
      const media = drawingRels[kid(kid(pic, 'blipFill'), 'blip').attrs['r:embed']].target;
      pictures.push({
        anchor: local(anchor.name), editAs: anchor.attrs.editAs, name: kid(kid(pic, 'nvPicPr'), 'cNvPr').attrs.name,
        from: {col: num(from, 'col'), colOff: num(from, 'colOff'), row: num(from, 'row'), rowOff: num(from, 'rowOff')},
        ext: ext ? {cx: +ext.attrs.cx, cy: +ext.attrs.cy} : null, media, bytes: parts.get(media),
      });
    }
  }
  // Page setup, with the schema's defaults for what is not written (fitToWidth/Height 1, paperSize 1 = Letter).
  const setup = kid(xml, 'pageSetup')?.attrs || {};
  const print = {
    fitToPage: kid(kid(xml, 'sheetPr'), 'pageSetUpPr')?.attrs.fitToPage === '1', fitToWidth: +(setup.fitToWidth ?? 1), fitToHeight: +(setup.fitToHeight ?? 1),
    paperSize: +(setup.paperSize ?? 1), orientation: setup.orientation || 'default',
  };
  const widthOf = c => cols.find(x => x.min <= c + 1 && c + 1 <= x.max)?.width;
  return {
    name, path, cells, rows, cols, widthOf, defaultRowHeight: format.defaultRowHeight === undefined ? undefined : +format.defaultRowHeight,
    customDefaultHeight: format.customHeight === '1', merges: kids(kid(xml, 'mergeCells'), 'mergeCell').map(m => m.attrs.ref), view, links, pictures, print,
    dimension: kid(xml, 'dimension')?.attrs.ref, root: xml,
  };
}

/** Everything the tests check in an .xlsx file (a Buffer). */
export function readXlsx(buffer) {
  const parts = unzip(buffer);
  const contentTypes = {}, defaults = {};
  const types = parseXml(parts.get('[Content_Types].xml').toString('utf8'));
  for (const d of kids(types, 'Default')) defaults[d.attrs.Extension.toLowerCase()] = d.attrs.ContentType;
  for (const o of kids(types, 'Override')) contentTypes[o.attrs.PartName.slice(1)] = o.attrs.ContentType;
  const typeOf = name => contentTypes[name] || defaults[name.replace(/^.*\./, '').toLowerCase()];
  for (const name of parts.keys()) if (!name.endsWith('/') && !typeOf(name)) throw new Error('部品の種類が[Content_Types].xmlにありません: ' + name);
  const rootRels = relationships(parts, ''), mainPath = Object.values(rootRels).find(r => r.type === 'officeDocument').target;
  const workbook = parseXml(parts.get(mainPath).toString('utf8')), rels = relationships(parts, mainPath);
  const stylesPath = Object.values(rels).find(r => r.type === 'styles').target, styles = readStyles(parseXml(parts.get(stylesPath).toString('utf8')));
  const stringsRel = Object.values(rels).find(r => r.type === 'sharedStrings');
  const strings = stringsRel ? kids(parseXml(parts.get(stringsRel.target).toString('utf8')), 'si').map(si => unescapeX(textOf(si))) : [];
  const sheets = kids(kid(workbook, 'sheets'), 'sheet').map(s => readSheet(parts, rels[s.attrs['r:id']].target, s.attrs.name, styles, strings));
  const coreRel = Object.values(rootRels).find(r => r.type === 'core-properties');
  const title = coreRel ? textOf(kid(parseXml(parts.get(coreRel.target).toString('utf8')), 'title')) : undefined;
  const calc = kid(workbook, 'calcPr')?.attrs || {};
  return {parts, typeOf, sheets, styles, strings, title, calc, activeTab: +(kid(kid(workbook, 'bookViews'), 'workbookView')?.attrs.activeTab || 0)};
}
