// tests/xlsx-import-patterns.test.mjs の書き方の種類を、デスクトップ版Excelで確かめるための見本のxlsxを書き出す。
// 見本は、テストと同じ tests/lib/xlsx.mjs の handXlsx で手書きする。Excelが開けるように、シートの表示の設定（sheetViews）も書く
// （これがないと、Excelは customHeight の既定の行の高さを、画面の倍率に応じて縮めて表示する）。
//
//   node tests/fixtures/make-xlsx-patterns.mjs <書き出すフォルダー>
//
// 書き出したファイルごとに、読み込んでからExcelで照合し、想定外の差が0件になることを確かめる（Windows・デスクトップ版Excel）。
//   python scripts/build.py --from-xlsx <フォルダー>/patterns-styles.xlsx --output <フォルダー>/patterns-styles.html
//   python scripts/verify_xlsx.py --base <フォルダー>/patterns-styles.html --xlsx <フォルダー>/patterns-styles.xlsx --source
import {mkdirSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {handXlsx, stylesXml} from '../lib/xlsx.mjs';

const out = process.argv[2];
if (!out) { console.error('使い方: node tests/fixtures/make-xlsx-patterns.mjs <書き出すフォルダー>'); process.exit(2); }
mkdirSync(out, {recursive: true});
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
const VIEW = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
const written = [];
const write = (name, bytes) => { writeFileSync(join(out, name), bytes); written.push(name); };

// ---- 書式: 1行に1つの書き方。A列は文字（表示形式は数）、B列は負の数、C列は書き方の名前。
const fonts = ['<font><sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font>'];
const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
const numFmts = {}, cases = [];
const add = (label, {font, fill, border, numFmt, alignment = 'vertical="center"', value = '文字', negative = false} = {}) => {
  const ids = {font: 0, fill: 0, border: 0, numFmt: 0};
  if (font) { fonts.push(font); ids.font = fonts.length - 1; }
  if (fill) { fills.push(fill); ids.fill = fills.length - 1; }
  if (border) { borders.push(border); ids.border = borders.length - 1; }
  if (numFmt !== undefined) { const id = 176 + Object.keys(numFmts).length; numFmts[id] = numFmt; ids.numFmt = id; }
  cases.push({label, ids, alignment, value, negative});
};
for (const color of ['<color rgb="FF00B050"/>', '<color theme="5" tint="-0.249977111117893"/>', '<color indexed="10"/>', '<color indexed="64"/>', '<color auto="1"/>'])
  add('文字の色 ' + color, {font: `<font><sz val="11"/>${color}<name val="游ゴシック"/></font>`});
for (const name of ['游ゴシック', '游明朝', 'メイリオ', 'ＭＳ ゴシック', 'ＭＳ Ｐゴシック', 'ＭＳ 明朝', 'ＭＳ Ｐ明朝', 'Yu Gothic UI', 'Arial', 'Calibri'])
  add('フォント ' + name, {font: `<font><sz val="11"/><name val="${name}"/></font>`});
for (const [label, font] of [['名前と大きさなし', '<font><b/></font>'], ['b=0 i=false strike', '<font><b val="0"/><i val="false"/><strike/><sz val="11"/><name val="游ゴシック"/></font>'],
  ...['double', 'singleAccounting', 'none'].map(u => ['下線 ' + u, `<font><u val="${u}"/><sz val="11"/><name val="游ゴシック"/></font>`]),
  ['下線', '<font><u/><sz val="11"/><name val="游ゴシック"/></font>'], ['上付き', '<font><vertAlign val="superscript"/><sz val="11"/><name val="游ゴシック"/></font>'],
  ['4pt', '<font><sz val="4"/><name val="游ゴシック"/></font>'], ['120pt', '<font><sz val="120"/><name val="游ゴシック"/></font>']]) add(label, {font});
const pattern = type => `<fill><patternFill patternType="${type}"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>`;
for (const type of ['solid', 'none', 'gray0625', 'gray125', 'lightGray', 'mediumGray', 'darkGray', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
  'lightGrid', 'lightTrellis', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis']) add('塗り ' + type, {fill: pattern(type)});
const stop = (position, rgb) => `<stop position="${position}"><color rgb="FF${rgb}"/></stop>`;
add('グラデーション2色', {fill: `<fill><gradientFill degree="90">${stop(0, 'FF0000')}${stop(1, '0000FF')}</gradientFill></fill>`});
add('グラデーション3色', {fill: `<fill><gradientFill>${stop(0, 'FF0000')}${stop(0.5, '00FF00')}${stop(1, '0000FF')}</gradientFill></fill>`});
add('種類のない塗り', {fill: '<fill><patternFill><fgColor rgb="FF00FF00"/></patternFill></fill>'});
for (const k of ['hair', 'thin', 'medium', 'thick', 'dashed', 'mediumDashed', 'dotted', 'double', 'dashDot', 'dashDotDot', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot'])
  add('罫線 ' + k, {border: `<border><left/><right/><top style="${k}"><color rgb="FF0000FF"/></top><bottom/><diagonal/></border>`});
add('斜めの罫線', {border: '<border diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color auto="1"/></diagonal></border>'});
add('罫線の色', {border: '<border><left style="thin"><color indexed="10"/></left><right style="medium"><color theme="4"/></right><top/><bottom style="thin"/><diagonal/></border>'});
add('start end', {border: '<border><start style="thin"/><end style="dotted"/><top/><bottom/><diagonal/></border>'});
for (const alignment of ['horizontal="left"', 'horizontal="center"', 'horizontal="right"', 'horizontal="fill"', 'horizontal="justify"', 'horizontal="distributed"',
  'horizontal="centerContinuous"'].map(a => a + ' vertical="center"').concat(['vertical="top"', 'vertical="bottom"', 'vertical="justify"', 'vertical="distributed"',
  'vertical="center" wrapText="1"', 'vertical="center" wrapText="true"', 'horizontal="left" vertical="center" indent="3"', 'horizontal="right" vertical="center" indent="20"',
  'vertical="center" textRotation="90"', 'vertical="center" textRotation="255"', 'vertical="center" shrinkToFit="1"', 'vertical="center" indent="2"'])) add('配置 ' + alignment, {alignment});
for (const code of ['0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', 'yyyy/mm/dd', '"¥"#,##0', '[$¥-411]#,##0', '#,##0;-#,##0', '0_ ', '#,##0;[Red]-#,##0', '#,##0_);(#,##0)',
  '"¥"#,##0;"¥"\\-#,##0', '#,##0;-#,##0;"-"', '[Red]#,##0', '[Blue]0.00', '[>=100]0', '[DBNum1]0', '0.0', '0.000', '#,##0.0', '0.0%', '￥#,##0', '#,##0"円"',
  'yyyy/m/d', 'm/d/yy', 'yyyy"年"m"月"d"日"', '[$-411]ggge"年"m"月"d"日"', 'mmm-yy', 'yyyy/m/d h:mm', 'h:mm', 'h:mm:ss AM/PM', '[h]:mm:ss', 'mm:ss', 'h"時"mm"分"',
  '0.00E+00', '# ?/?', ';;;']) add('表示形式 ' + code, {numFmt: code, value: /[ymdh]|ge/.test(code.replace(/"[^"]*"/g, '')) ? 46023.5 : 1234.567, negative: true});
add('文字の表示形式 @', {numFmt: '@'});
add('文字の表示形式 "(株)"@', {numFmt: '"(株)"@'});
const rows = cases.map((k, i) => {
  const r = i + 1, cell = (a, v) => typeof v === 'number' ? `<c r="${a}${r}" s="${r}"><v>${v}</v></c>` : `<c r="${a}${r}" s="${r}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  return `<row r="${r}">${cell('A', k.value)}${k.negative ? cell('B', -k.value) : ''}<c r="C${r}" t="inlineStr"><is><t xml:space="preserve">${esc(k.label)}</t></is></c></row>`;
}).join('');
const xfs = cases.map(({ids, alignment}) => `<xf numFmtId="${ids.numFmt}" fontId="${ids.font}" fillId="${ids.fill}" borderId="${ids.border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment ${alignment}/></xf>`);
write('patterns-styles.xlsx', handXlsx({sheets: [{name: '書式', xml: `${VIEW}<sheetFormatPr defaultRowHeight="18.75"/><cols><col min="1" max="2" width="20" customWidth="1"/><col min="3" max="3" width="40" customWidth="1"/></cols><sheetData>${rows}</sheetData>`}],
  styles: stylesXml({numFmts, fonts, fills, borders, xfs})}));

// ---- 文字: 前後の空白（xml:space）と、run の書式
const red = '<rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="游ゴシック"/><family val="3"/><charset val="128"/></rPr>';
const texts = ['<t> 前の空白</t>', '<t>後ろの空白 </t>', '<t>後ろの改行\n</t>', '<t>\t前のタブ</t>', '<t xml:space="preserve"> 残す </t>', '<t>中の  空白</t>', '<r><t>前 </t></r><r><t> 後</t></r>',
  `<r>${red}<t>太字の赤</t></r>`, `<r>${red}<t>太字</t></r><r>${red}<t>の赤</t></r>`, '<r><rPr><b/></rPr><t>大きさなし</t></r>', `<r>${red}<t>赤</t></r><r><t>普通</t></r>`];
write('patterns-text.xlsx', handXlsx({strings: '<si><t>  共有の前の空白</t></si>', styles: stylesXml({}), sheets: [{name: '文字', xml: `${VIEW}<sheetData>` +
  texts.map((t, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is>${t}</is></c></row>`).join('') + `<row r="${texts.length + 1}"><c r="A${texts.length + 1}" t="s"><v>0</v></c></row></sheetData>`}]}));

// ---- 独自のテーマとパレット
const theme = '<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
  ['ABCDEF', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF'].map((c, i) => `<a:accent${i + 1}><a:srgbClr val="${c}"/></a:accent${i + 1}>`).join('') +
  '<a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>';
const themeColors = ['<color theme="1"/>', '<color theme="4"/>', '<color theme="3"/>', '<color indexed="0"/>', '<color indexed="1"/>', '<color theme="4" tint="0.3999755851924192"/>'];
write('patterns-theme.xlsx', handXlsx({theme, sheets: [{name: 'テーマ', xml: `${VIEW}<sheetData>` + themeColors.map((_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i + 1}" t="inlineStr"><is><t>文字</t></is></c>` +
  `<c r="B${i + 1}" s="${i + 1 + themeColors.length}" t="inlineStr"><is><t>塗り</t></is></c></row>`).join('') + '</sheetData>'}],
  styles: stylesXml({fonts: [fonts[0], ...themeColors.map(color => `<font><sz val="11"/>${color}<name val="游ゴシック"/></font>`)],
    fills: [...fills, ...themeColors.map(color => `<fill><patternFill patternType="solid">${color.replace('<color', '<fgColor')}<bgColor indexed="64"/></patternFill></fill>`)],
    xfs: [...themeColors.map((_, i) => `<xf numFmtId="0" fontId="${i + 1}" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment vertical="center"/></xf>`),
      ...themeColors.map((_, i) => `<xf numFmtId="0" fontId="0" fillId="${i + 2}" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>`)],
    colors: '<colors><indexedColors><rgbColor rgb="FF123456"/><rgbColor rgb="FF654321"/></indexedColors></colors>'})}));

// ---- シート: 大きさ・表示・列と行の書式・非表示のシート
const v = (a, value = 1) => `<c r="${a}"><v>${value}</v></c>`;
const view = (inner = '', attrs = '') => `<sheetViews><sheetView${attrs} workbookViewId="0">${inner}</sheetView></sheetViews>`;
const sheetFills = [...fills, '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>',
  '<fill><patternFill patternType="solid"><fgColor rgb="FF00B0F0"/><bgColor indexed="64"/></patternFill></fill>'];
const sheets = [
  {name: '幅と高さ', xml: view() + '<sheetFormatPr baseColWidth="10" defaultRowHeight="20"/><cols><col min="2" max="2" width="1" customWidth="1"/><col min="3" max="3" width="255" customWidth="1"/></cols>' +
    `<sheetData><row r="1">${v('A1')}${v('B1')}${v('C1')}${v('D1')}</row><row r="3" hidden="1">${v('A3', 3)}</row><row r="4" ht="0" customHeight="1">${v('A4', 4)}</row>` +
    `<row r="5" ht="6" customHeight="1">${v('A5', 5)}</row><row r="6" ht="409.5" customHeight="1">${v('A6', 6)}</row></sheetData>`},
  {name: '既定の幅', xml: view() + `<sheetFormatPr defaultColWidth="12" defaultRowHeight="18.75"/><sheetData><row r="1">${v('A1')}${v('B1')}</row></sheetData>`},
  {name: '列と行の書式', xml: view() + '<sheetFormatPr defaultRowHeight="18.75"/><cols><col min="2" max="3" width="9" style="1" customWidth="1"/></cols>' +
    '<sheetData><row r="2" s="2" customFormat="1"><c r="A2" t="inlineStr"><is><t>a</t></is></c></row><row r="4"><c r="B4" t="inlineStr"><is><t>own</t></is></c></row></sheetData>'},
  {name: '列だけ固定', xml: view('<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/><selection pane="topRight"/>') + `<sheetData><row r="1">${v('A1')}${v('D1')}</row></sheetData>`},
  {name: 'frozenSplit', xml: view('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozenSplit"/><selection pane="bottomLeft"/>') + `<sheetData><row r="1">${v('A1')}</row><row r="3">${v('A3')}</row></sheetData>`},
  {name: '途中から固定', xml: view('<pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>', ' topLeftCell="A5"') + `<sheetData><row r="1">${v('A1')}</row><row r="9">${v('A9')}</row></sheetData>`},
  {name: '結合をまたぐ固定', xml: view('<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>') + `<sheetData><row r="1">${v('A1')}</row><row r="2">${v('A2')}</row></sheetData><mergeCells count="1"><mergeCell ref="A2:A3"/></mergeCells>`},
  {name: '分割', xml: view('<pane xSplit="2000" ySplit="1000" topLeftCell="C5" activePane="bottomRight"/><selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight"/>') + `<sheetData><row r="1">${v('A1')}</row></sheetData>`},
  {name: '選択と倍率', xml: view('<selection activeCell="B2" sqref="B2:C3 E5"/>', ' showGridLines="false" tabSelected="1" zoomScale="300" zoomScaleNormal="300"') + `<sheetData><row r="1">${v('A1')}</row><row r="5">${v('E5')}</row></sheetData>`},
  {name: '隠し', state: 'hidden', xml: view() + `<sheetData><row r="1">${v('A1')}</row></sheetData>`},
];
write('patterns-sheet.xlsx', handXlsx({sheets, workbook: '<bookViews><workbookView activeTab="8"/></bookViews>',
  styles: stylesXml({fills: sheetFills, xfs: [2, 3].map(fill => `<xf numFmtId="0" fontId="0" fillId="${fill}" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>`)})}));

// ---- 書式・テーマ・共有文字列の部品がないファイル
write('patterns-bare.xlsx', handXlsx({styles: null, strings: null, sheets: [{name: 'S', xml: `${VIEW}<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="B1" t="inlineStr"><is><t>文字</t></is></c></row></sheetData>`}]}));
console.log(`書き出した（${out}）: ${written.join(' ')}`);
