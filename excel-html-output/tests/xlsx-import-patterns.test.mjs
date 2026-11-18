// scripts/read_xlsx.cjs：xlsx の書き方の種類ごとに、HTMLの表のデータへの読み方を確かめる。
// ここで読むxlsxは、テストの中で手で書いた小さなファイルである（tests/lib/xlsx.mjs の handXlsx）。
// Excelの表示に合わせる期待値（色・幅・配置など）は、同じ書き方のファイルをデスクトップ版Excel 16.0 で開き、
// scripts/verify_xlsx.py --source でHTMLとの想定外の差が0件になることを確かめてある（見本は tests/fixtures/make-xlsx-patterns.mjs で作れる）。
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {test} from 'node:test';
import {SKILL_ROOT} from './lib/workbook.mjs';
import {handXlsx, makeZip, miniXlsx, relsXml, stylesXml} from './lib/xlsx.mjs';

const require = createRequire(import.meta.url);
const {fromXlsx, mapFormat} = require(join(SKILL_ROOT, 'scripts', 'read_xlsx.cjs'));
const warned = (result, pattern) => result.warnings.some(w => pattern.test(w));
const ribbonFonts = [...readFileSync(join(SKILL_ROOT, 'assets', 'shell.html'), 'utf8').match(/<select id="font"[^>]*>(.*?)<\/select>/s)[1]
  .matchAll(/<option>([^<]+)<\/option>/g)].map(m => m[1]);
const FONT = '<font><sz val="11"/><name val="游ゴシック"/></font>';

/** A cell format (an item of cellXfs) using fonts[font], fills[fill], borders[border] and the number format numFmt. */
const xf = ({font = 0, fill = 0, border = 0, numFmt = 0, alignment = 'vertical="center"'} = {}) =>
  `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0">${alignment === null ? '' : `<alignment ${alignment}/>`}</xf>`;
/** Reads a one-sheet book whose cell A{n} (the text 文字) has the cell format xfs[n - 1]; the other styles as stylesXml takes them. */
function readStyled({xfs, theme = null, ...styles}) {
  const rows = xfs.map((_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i + 1}" t="inlineStr"><is><t>文字</t></is></c></row>`).join('');
  return fromXlsx(handXlsx({sheets: [{name: 'S', xml: `<sheetData>${rows}</sheetData>`}], theme, styles: stylesXml({...styles, xfs})}));
}
const cellsOf = result => result.workbook.sheets[0].cells;
/** The field of the format of cells A1, A2, ... (undefined where it is the book's default). */
const column = (result, field, count) => Array.from({length: count}, (_, i) => cellsOf(result)['A' + (i + 1)]?.s?.[field]);

test('文字の色は、RGB・テーマ（テーマのないファイルはOfficeの既定のテーマ）・パレットの番号・システムの色・自動のどれでも、Excelが表示する色にする', () => {
  const colors = ['<color rgb="FF00B050"/>', '<color theme="5" tint="-0.249977111117893"/>', '<color indexed="10"/>', '<color indexed="64"/>', '<color auto="1"/>', '<color rgb="xyz"/>', ''];
  const fonts = [FONT, ...colors.map(color => `<font><sz val="11"/>${color}<name val="游ゴシック"/></font>`)];
  const result = readStyled({fonts, xfs: colors.map((_, i) => xf({font: i + 1}))});
  assert.deepEqual(column(result, 'color', colors.length), ['#00b050', '#c65911', '#ff0000', undefined, undefined, undefined, undefined],
    'パレットの10番は赤、64番はシステムの文字の色（黒）。自動・読めない色・色の指定なしは黒（既定なので書かない）');
});

test('ファイルにある独自のパレット（indexedColors）とテーマの色を使い、テーマのシステムの色（dk1・lt1）はWindowsの色にする', () => {
  const theme = '<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
    ['ABCDEF', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF'].map((c, i) => `<a:accent${i + 1}><a:srgbClr val="${c}"/></a:accent${i + 1}>`).join('') +
    '<a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>';
  const colors = ['<color theme="1"/>', '<color theme="4"/>', '<color theme="3"/>', '<color indexed="0"/>', '<color indexed="1"/>'];
  const result = readStyled({theme, fonts: [FONT, ...colors.map(color => `<font><sz val="11"/>${color}<name val="游ゴシック"/></font>`)],
    xfs: colors.map((_, i) => xf({font: i + 1})), colors: '<colors><indexedColors><rgbColor rgb="FF123456"/><rgbColor rgb="FF654321"/></indexedColors></colors>'});
  assert.deepEqual(column(result, 'color', colors.length), [undefined, '#abcdef', '#222222', '#123456', '#654321'],
    'テーマの番号は lt1, dk1, lt2, dk2, アクセント1… の順。dk1 のシステムの色は、ファイルにある lastClr（#111111）でなく、Windowsの文字の色（黒）');
});

test('網かけの塗り（17種類）は前景と背景の色を混ぜた1色、グラデーションは止めの色の平均の1色にして知らせ、べた塗りと塗りなしはそのまま', () => {
  const patterns = ['gray0625', 'gray125', 'lightGray', 'mediumGray', 'darkGray', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
    'lightGrid', 'lightTrellis', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis'];
  const pattern = type => `<fill><patternFill patternType="${type}"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>`;
  const stop = (position, rgb) => `<stop position="${position}"><color rgb="FF${rgb}"/></stop>`;
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>', pattern('solid'), pattern('none'), ...patterns.map(pattern),
    `<fill><gradientFill degree="90">${stop(0, 'FF0000')}${stop(1, '0000FF')}</gradientFill></fill>`,
    `<fill><gradientFill>${stop(0, 'FF0000')}${stop(0.5, '00FF00')}${stop(1, '0000FF')}</gradientFill></fill>`, '<fill><patternFill><fgColor rgb="FF00FF00"/></patternFill></fill>'];
  const result = readStyled({fills, xfs: fills.slice(2).map((_, i) => xf({fill: i + 2}))}), all = column(result, 'fill', fills.length - 2), got = all.slice(0, -1);
  assert.equal(all.at(-1), undefined, '網かけの種類（patternType）のない塗りは、前景の色があってもExcelと同じく塗らない');
  assert.deepEqual(got.slice(0, 2), ['#ff0000', undefined], 'べた塗りは前景の色、塗りなしは塗らない');
  assert.deepEqual(got.slice(2, 7), ['#ffefef', '#ffdfdf', '#ffbfbf', '#ff8080', '#ff4040'], '灰色の網かけは、前景の色を6.25%・12.5%・25%・50%・75%混ぜる');
  for (const [i, type] of patterns.entries()) {
    const [r, g, b] = [1, 3, 5].map(k => parseInt(got[i + 2].slice(k, k + 2), 16));
    assert.ok(r === 255 && g === b && g > 0 && g < 255, `${type} は赤と白の間の色: ${got[i + 2]}`);
    assert.ok(warned(result, new RegExp(`網かけ（${type}）の塗りは、混ぜた色1色にした`)), type);
  }
  assert.deepEqual(got.slice(-2), ['#800080', '#555555'], 'グラデーションは、止めの位置によらず色の平均');
  assert.ok(warned(result, /グラデーションの塗りは、その平均の色1色にした/));
});

test('罫線の13種類は、HTMLにある8種類はそのまま、ない5種類は近い線にして知らせ、斜めの罫線は外して知らせる', () => {
  const kinds = {hair: '0.5px solid', thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', mediumDashed: '2px dashed', dotted: '1px dotted',
    double: '3px double', dashDot: '1px dashed', dashDotDot: '1px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed'};
  const near = ['dashDot', 'dashDotDot', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot'], n = Object.keys(kinds).length;
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>',
    ...Object.keys(kinds).map(k => `<border><left/><right/><top style="${k}"><color rgb="FF0000FF"/></top><bottom/><diagonal/></border>`),
    '<border diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color auto="1"/></diagonal></border>',
    '<border><left style="thin"><color indexed="10"/></left><right style="medium"><color theme="4"/></right><top/><bottom style="thin"/><diagonal/></border>',
    '<border><start style="thin"/><end style="dotted"/><top/><bottom/><diagonal/></border>'];
  const result = readStyled({borders, xfs: borders.slice(1).map((_, i) => xf({border: i + 1}))}), cells = cellsOf(result);
  Object.entries(kinds).forEach(([k, kind], i) => assert.equal(cells['A' + (i + 1)].s.borderTop, kind + ' #0000ff', k));
  for (const k of Object.keys(kinds)) assert.equal(warned(result, new RegExp(`罫線（${k}）は`)), near.includes(k), k);
  for (const k of near) {
    // 上の罫線は、Excelでは上のセルの下の罫線としても読めるので、照合のために上のセルにも理由を付ける。
    const row = Object.keys(kinds).indexOf(k) + 1, reason = `罫線（${k}）は「${kinds[k]}」にした`, of = a => result.notes.S.cells[a]?.filter(n => n.reason === reason);
    assert.deepEqual([of('A' + row), of('A' + (row - 1))], [[{fields: ['bTop'], reason}], [{fields: ['bBottom'], reason}]], k);
  }
  assert.equal(cells['A' + (n + 1)].s, undefined, '斜めの罫線だけのセルは、罫線なし');
  assert.ok(warned(result, /斜めの罫線は移せない/));
  assert.deepEqual(cells['A' + (n + 2)].s, {borderLeft: '1px solid #ff0000', borderRight: '2px solid #4472c4', borderBottom: '1px solid #000000'},
    '線の色は、パレットの番号・テーマ・指定なし（黒）');
  assert.deepEqual(cells['A' + (n + 3)].s, {borderLeft: '1px solid #000000', borderRight: '1px dotted #000000'}, 'start・end は左・右');
});

test('フォント名は、日本語の名前をリボンの一覧と同じ英語の名前にする（游ゴシック → Yu Gothic など）', () => {
  const names = {'游ゴシック': 'Yu Gothic', '游明朝': 'Yu Mincho', 'メイリオ': 'Meiryo', 'ＭＳ ゴシック': 'MS Gothic', 'ＭＳ Ｐゴシック': 'MS PGothic', 'ＭＳ 明朝': 'MS Mincho',
    'ＭＳ Ｐ明朝': 'MS PMincho', 'Yu Gothic UI': 'Yu Gothic UI', 'Arial': 'Arial', 'Calibri': 'Calibri'};
  const result = readStyled({fonts: [FONT, ...Object.keys(names).map(name => `<font><sz val="11"/><name val="${name}"/></font>`)], xfs: Object.keys(names).map((_, i) => xf({font: i + 1}))});
  assert.deepEqual(column(result, 'font', Object.keys(names).length).map(f => f ?? 'Yu Gothic'), Object.values(names));
  for (const name of ['游ゴシック', 'メイリオ', 'ＭＳ Ｐゴシック']) assert.ok(ribbonFonts.includes(names[name]), `${name} のセルを選ぶと、リボンのフォントの欄にその名前が出る`);
});

test('フォントの太字・斜体・取り消し線・下線の種類・上付き、名前や大きさのないフォント、HTMLの範囲（6〜96pt）の外の大きさを読む', () => {
  const fonts = [FONT, '<font><b/></font>', '<font><b val="0"/><i val="false"/><strike/><sz val="11"/><name val="游ゴシック"/></font>',
    ...['double', 'singleAccounting', 'none'].map(u => `<font><u val="${u}"/><sz val="11"/><name val="游ゴシック"/></font>`), '<font><u/><sz val="11"/><name val="游ゴシック"/></font>',
    '<font><vertAlign val="superscript"/><sz val="11"/><name val="游ゴシック"/></font>', '<font><sz val="4"/><name val="游ゴシック"/></font>', '<font><sz val="120"/><name val="游ゴシック"/></font>'];
  const result = readStyled({fonts, xfs: fonts.slice(1).map((_, i) => xf({font: i + 1}))}), cells = cellsOf(result);
  assert.deepEqual(cells.A1.s, {bold: true}, '名前と大きさのないフォントは、游ゴシック11pt');
  assert.deepEqual(cells.A2.s, {strike: true}, 'val="0"・val="false" はその書式なし');
  assert.deepEqual([cells.A3.s, cells.A4.s, cells.A5.s, cells.A6.s], [{underline: true}, {underline: true}, undefined, {underline: true}],
    '二重下線・会計用の下線は一重の下線、val="none" は下線なし');
  assert.ok(warned(result, /下線（double）は一重の下線にした/) && warned(result, /下線（singleAccounting）は一重の下線にした/));
  assert.ok(!warned(result, /下線（none）/));
  assert.equal(cells.A7.s, undefined);
  assert.ok(warned(result, /上付き・下付きの文字は、普通の文字にした/));
  assert.deepEqual([cells.A8.s, cells.A9.s], [{size: 6}, {size: 96}]);
  assert.ok(warned(result, /文字の大きさ 4pt は 6pt にした/) && warned(result, /文字の大きさ 120pt は 96pt にした/));
});

test('横位置・縦位置・折り返し・インデントを読み、HTMLにない配置・16以上のインデント・回転・縮小して全体を表示は、近いものにするか外して知らせる', () => {
  const aligns = ['horizontal="left"', 'horizontal="center"', 'horizontal="right"', 'horizontal="fill"', 'horizontal="justify"', 'horizontal="distributed"', 'horizontal="centerContinuous"']
    .map(a => a + ' vertical="center"').concat(['vertical="top"', 'vertical="bottom"', 'vertical="justify"', 'vertical="distributed"', 'vertical="center" wrapText="1"',
      'vertical="center" wrapText="true"', 'horizontal="left" vertical="center" indent="3"', 'horizontal="right" vertical="center" indent="20"',
      'vertical="center" textRotation="90"', 'vertical="center" textRotation="255"', 'vertical="center" shrinkToFit="1"', 'vertical="center" indent="2"']);
  const result = readStyled({xfs: aligns.map(alignment => xf({alignment}))}), cells = cellsOf(result);
  assert.deepEqual(aligns.map((_, i) => cells['A' + (i + 1)].s), [{align: 'left'}, {align: 'center'}, {align: 'right'}, {align: 'left'}, {align: 'left'}, {align: 'center'},
    {align: 'center'}, {valign: 'top'}, {valign: 'bottom'}, undefined, undefined, {wrap: true}, {wrap: true}, {align: 'left', indent: 3}, {align: 'right', indent: 15},
    undefined, undefined, undefined, {align: 'left', indent: 2}]);
  assert.ok(warned(result, /「標準」の横位置のインデントは、左揃えのインデントにした/), 'HTMLのインデントは左か右に揃えたセルだけなので、左揃えにしたことを知らせる');
  for (const [kind, near] of [['fill', 'left'], ['justify', 'left'], ['distributed', 'center'], ['centerContinuous', 'center']])
    assert.ok(warned(result, new RegExp(`横位置（${kind}）は「${near}」にした`)), kind);
  for (const kind of ['justify', 'distributed']) assert.ok(warned(result, new RegExp(`縦位置（${kind}）は「middle」にした`)), kind);
  for (const pattern of [/インデント 20 は 15 にした/, /文字の回転・縦書きは移せない/, /「縮小して全体を表示する」は移せない/]) assert.ok(warned(result, pattern), String(pattern));
  assert.ok(!warned(result, /横位置（(left|center|right)）/) && !warned(result, /縦位置（(top|center|bottom)）/), 'HTMLにある配置は知らせない');
});

test('表示形式は、HTMLにある10種類ならそのまま、ないものは近い形式にし、色・条件・漢数字の指定や負の数の別の書き方があれば同じ表示ではないとする', () => {
  // [Excelの表示形式, HTMLの表示形式, 同じ表示か]
  const cases = [
    ['General', 'General', true], ['G/標準', 'General', true], ['0', '0', true], ['0.00', '0.00', true], ['#,##0', '#,##0', true], ['#,##0.00', '#,##0.00', true],
    ['0%', '0%', true], ['0.00%', '0.00%', true], ['@', '@', true], ['yyyy/mm/dd', 'yyyy/mm/dd', true], ['"¥"#,##0', '¥#,##0', true], ['[$¥-411]#,##0', '¥#,##0', true],
    ['#,##0;-#,##0', '#,##0', true], ['0_ ', '0', true],
    ['#,##0;[Red]-#,##0', '#,##0', false], ['#,##0_);(#,##0)', '#,##0', false], ['"¥"#,##0;"¥"\\-#,##0', '¥#,##0', false], ['#,##0;-#,##0;"-"', '#,##0', false],
    ['[Red]#,##0', '#,##0', false], ['[Blue]0.00', '0.00', false], ['[>=100]0', '0', false], ['[DBNum1]0', '0', false],
    ['0.0', '0.00', false], ['0.000', '0.00', false], ['#,##0.0', '#,##0.00', false], ['0.0%', '0.00%', false], ['￥#,##0', '¥#,##0', false], ['#,##0"円"', '#,##0', false],
    ['yyyy/m/d', 'yyyy/mm/dd', false], ['m/d/yy', 'yyyy/mm/dd', false], ['yyyy"年"m"月"d"日"', 'yyyy/mm/dd', false], ['[$-411]ggge"年"m"月"d"日"', 'yyyy/mm/dd', false],
    ['mmm-yy', 'yyyy/mm/dd', false], ['yyyy/m/d h:mm', 'yyyy/mm/dd', false],
    ['h:mm', 'General', false], ['h:mm:ss AM/PM', 'General', false], ['[h]:mm:ss', 'General', false], ['mm:ss', 'General', false], ['h"時"mm"分"', 'General', false],
    ['0.00E+00', 'General', false], ['# ?/?', 'General', false], [';;;', 'General', false], ['"(株)"@', 'General', false]];
  const wrong = cases.filter(([code, fmt, exact]) => { const got = mapFormat(code); return got.fmt !== fmt || got.exact !== exact; })
    .map(([code, fmt, exact]) => `${code}: ${JSON.stringify(mapFormat(code))}（期待 ${fmt}、${exact}）`);
  assert.deepEqual(wrong, []);
});

test('表示が同じでも書き方の違う表示形式（[$¥-411]#,##0 など）は、知らせずに、表示形式の書き方が違うことだけを照合に伝える', () => {
  const sheet = '<sheetData><row r="1"><c r="A1" s="1"><v>-1234</v></c><c r="B1" s="2"><v>5</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet, styles: stylesXml({numFmts: {176: '[$¥-411]#,##0', 177: '0_ '}, xfs: [xf({numFmt: 176}), xf({numFmt: 177})]})}));
  assert.deepEqual([cellsOf(result).A1.s, cellsOf(result).B1.s], [{numFmt: '¥#,##0'}, {numFmt: '0'}]);
  assert.deepEqual(result.warnings, []);
  assert.deepEqual(result.notes.S.cells.A1, [{fields: ['numfmt'], reason: '表示形式「[$¥-411]#,##0」は、同じ表示の「¥#,##0」にした'}]);
  assert.deepEqual(result.notes.S.cells.B1, [{fields: ['numfmt'], reason: '表示形式「0_ 」は、同じ表示の「0」にした'}]);
});

test('文字の前後に文字を足す表示形式（"(株)"@ など）は、文字のセルでも表示が違うことを知らせる', () => {
  const result = readStyled({numFmts: {176: '"(株)"@'}, xfs: [xf({numFmt: 176})]});
  assert.ok(warned(result, /表示形式「"\(株\)"@」はHTMLにないので「General」にした: A1/));
  assert.deepEqual(result.notes.S.cells.A1, [{fields: ['numfmt', 'text'], reason: '表示形式「"(株)"@」は「General」にした'}]);
});

test('セルの値は、真偽値・数式でない文字列（t="str"）・時刻つきの日付（t="d"）・直接の文字列の run・指数の数も読み、共有文字列の表にない番号は空にする', () => {
  const sheet = '<sheetData><row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c><c r="C1" t="b"><v>true</v></c><c r="D1" t="str"><v>文字</v></c>' +
    '<c r="E1" s="2" t="d"><v>2026-01-01T12:00:00</v></c><c r="F1" t="inlineStr"><is><r><t>前</t></r><r><rPr><b/></rPr><t>後</t></r></is></c>' +
    '<c r="G1" t="s"><v>5</v></c><c r="H1"><v>1E+3</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet, strings: '<si><t>a</t></si>'})), cells = cellsOf(result);
  assert.deepEqual([cells.A1, cells.B1, cells.C1, cells.D1], [{t: 'b', v: true}, {t: 'b', v: false}, {t: 'b', v: true}, {t: 's', v: '文字'}]);
  assert.deepEqual(cells.E1, {t: 'n', v: 46023.5, s: {numFmt: 'yyyy/mm/dd'}}, '2026-01-01 12:00 は 46023.5');
  assert.equal(cells.F1.v, '前後');
  assert.ok(warned(result, /文字ごとに違う書式.*F1/));
  assert.equal(cells.G1, undefined, '共有文字列の表にない番号は、空のセル');
  assert.deepEqual(cells.H1, {t: 'n', v: 1000});
});

test('文字の前後の空白・タブ・改行は、Excelと同じく、xml:space="preserve" のない文字（run ごと）では除き、中の空白は残す', () => {
  // デスクトップ版Excel 16.0 で、同じ書き方のセルを開いて読み取った値。
  const sheet = '<sheetData><row r="1">' + ['<t> 前の空白</t>', '<t>後ろの改行\n</t>', '<t>\t前のタブ</t>', '<t xml:space="preserve"> 残す </t>', '<t>中の  空白</t>',
    '<r><t>前 </t></r><r><t> 後</t></r>'].map((t, i) => `<c r="${'ABCDEF'[i]}1" t="inlineStr"><is>${t}</is></c>`).join('') + '<c r="G1" t="s"><v>0</v></c></row></sheetData>';
  const cells = cellsOf(fromXlsx(miniXlsx({sheet, strings: '<si><t>  共有の前の空白</t></si>'})));
  assert.deepEqual([...'ABCDEFG'].map(c => cells[c + '1'].v), ['前の空白', '後ろの改行', '前のタブ', ' 残す ', '中の  空白', '前後', '共有の前の空白']);
});

test('run がすべて同じ書式の文字（1つの run だけのものも）は、その書式で表示し、run ごとに違う書式があるときだけ知らせる', () => {
  // Excelは、セル全体の書式としては「混在」と答えるが、文字は run の書式（太字の赤）で表示する（デスクトップ版Excel 16.0）。
  const red = '<rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="メイリオ"/><family val="3"/><charset val="128"/></rPr>';
  const sheet = '<sheetData><row r="1">' + [`<r>${red}<t>太字の赤</t></r>`, `<r>${red}<t>太字</t></r><r>${red}<t>の赤</t></r>`, '<r><rPr><i/></rPr><t>斜体だけ</t></r>',
    `<r>${red}<t>赤</t></r><r><t>普通</t></r>`].map((t, i) => `<c r="${'ABCD'[i]}1" t="inlineStr"><is>${t}</is></c>`).join('') + '</row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet})), cells = cellsOf(result);
  assert.deepEqual([cells.A1.s, cells.B1.s], [{font: 'Meiryo', size: 14, bold: true, color: '#ff0000'}, {font: 'Meiryo', size: 14, bold: true, color: '#ff0000'}]);
  assert.deepEqual(cells.C1.s, {italic: true}, 'run にない大きさ・フォント名・色は、セルの書式のもの');
  assert.equal(cells.D1.s, undefined, 'run ごとに違う書式は、セルの書式');
  assert.deepEqual(result.warnings, ['シート「S」：文字ごとに違う書式（一部だけ太字など）は、セル全体の書式にした（D1 など）']);
  const fields = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color'];
  assert.deepEqual(result.notes.S.cells.A1, [{fields, reason: '文字の書式（run）をセルの書式にした'}], 'Excelがセル全体の書式を「混在」と答えることを照合に伝える');
  assert.deepEqual(result.notes.S.cells.D1, [{fields, reason: '文字ごとの書式'}]);
});

test('行の高さが決まっていない行は、Excelが中の文字に合わせて広げるので、いちばん大きい文字に合う高さにし、近い値であることを照合に伝える', () => {
  const fonts = [FONT, '<font><sz val="22"/><name val="游ゴシック"/></font>', '<font><sz val="120"/><name val="游ゴシック"/></font>', '<font><vertAlign val="superscript"/><sz val="11"/><name val="游ゴシック"/></font>'];
  const sheet = '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18.75"/><sheetData>' +
    '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>22pt</t></is></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2" s="2" t="inlineStr"><is><t>120pt</t></is></c></row>' +
    '<row r="3"><c r="A3" s="3" t="inlineStr"><is><t>上付き</t></is></c></row><row r="4"><c r="A4"><v>1</v></c></row><row r="5" ht="30" customHeight="1"><c r="A5" s="1"><v>1</v></c></row>' +
    '<row r="6"><c r="A6" s="1" t="inlineStr"><is><t>2行の結合</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A6:A7"/></mergeCells>';
  const result = fromXlsx(miniXlsx({sheet, styles: stylesXml({fonts, xfs: [1, 2, 3].map(font => xf({font}))})}));
  assert.deepEqual(result.workbook.sheets[0].rowHeights.slice(0, 7), [50, 273, 25, 25, 40, 25, 25],
    '22pt は 18.75 × 22 ÷ 11 = 37.5pt、120pt は204.5pt（文字はHTMLの上限の96pt）。Excelと同じく、複数の行にまたがる結合セルの文字には合わせない');
  assert.deepEqual(result.notes.S.sheet.filter(n => n.item === 'rowHeight'), [{item: 'rowHeight', rows: [1, 2, 3], reason: 'Excelが中の文字に合わせて決める高さは、近い値'}],
    '高さを決めた行（5行目）と、標準の文字だけの行（4行目）には付けない');
});

test('配列数式とデータテーブルは値にして知らせ、数式が参照するセルまで表の範囲を広げる', () => {
  const sheet = '<sheetData><row r="1"><c r="A1"><f t="array" ref="A1">SUM(B1:B2*C1:C2)</f><v>11</v></c><c r="B1"><v>1</v></c><c r="C1"><v>3</v></c></row>' +
    '<row r="2"><c r="B2"><v>2</v></c><c r="C2"><v>4</v></c></row><row r="5"><c r="E5"><f t="dataTable" ref="E5:E6" dt2D="0" dtr="0" r1="B1"/><v>7</v></c></row>' +
    '<row r="6"><c r="E6"><v>8</v></c><c r="F6"><f>SUM(A1:A200)+SUM(A1:AD1)</f><v>26</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet})), s = result.workbook.sheets[0];
  assert.deepEqual([s.cells.A1, s.cells.E5, s.cells.E6], [{t: 'n', v: 11}, {t: 'n', v: 7}, {t: 'n', v: 8}]);
  assert.ok(warned(result, /A1（=SUM\(B1:B2\*C1:C2\)：配列数式）/));
  assert.ok(warned(result, /E5（データテーブル E5:E6）/));
  assert.deepEqual(result.notes.S.cells.E5, [{fields: ['hasFormula', 'formula'], reason: 'HTMLで計算できない数式は値にした'}]);
  assert.equal(s.cells.F6.f, '=SUM(A1:A200)+SUM(A1:AD1)');
  assert.deepEqual([s.rows, s.cols], [200, 30], '数式が参照する A200 と AD1 が表に入る');
});

test('http(s) でないリンク（ブックの中の場所・メール・ファイル）は移さずに知らせ、http(s) のリンクは大文字で書かれていても移す', () => {
  const sheet = '<sheetData><row r="1">' + ['A1', 'B1', 'C1', 'D1'].map(a => `<c r="${a}" t="inlineStr"><is><t>${a}</t></is></c>`).join('') + '</row></sheetData>' +
    '<hyperlinks><hyperlink ref="A1" location="\'別のシート\'!A1" display="A1"/><hyperlink ref="B1" r:id="rId1"/><hyperlink ref="C1" r:id="rId2"/><hyperlink ref="D1" r:id="rId3"/></hyperlinks>';
  const rels = [['rId1', 'hyperlink', 'mailto:someone@example.com', true], ['rId2', 'hyperlink', 'file:///C:/work/a.xlsx', true], ['rId3', 'hyperlink', 'HTTPS://EXAMPLE.COM/', true]];
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: sheet, rels}]})), cells = cellsOf(result);
  assert.deepEqual(['A1', 'B1', 'C1', 'D1'].map(a => cells[a].link), [undefined, undefined, undefined, 'HTTPS://EXAMPLE.COM/']);
  assert.ok(result.warnings.includes("シート「S」：次のリンクは、http(s) のリンクでないので移せない: A1（#'別のシート'!A1）、B1（mailto:someone@example.com）、C1（file:///C:/work/a.xlsx）"),
    'リンクがいくつあっても、全部を知らせる');
  for (const a of ['A1', 'B1', 'C1']) assert.deepEqual(result.notes.S.cells[a], [{fields: ['hyperlink', 'color', 'underline'], reason: 'http(s) 以外のリンク'}], a);
});

test('結合で隠れたセルの値は、Excelでも表示されないので読み込まずに、そのセルを全部知らせる', () => {
  const sheet = '<sheetData><row r="1">' + [['A1', '表示'], ['B1', '隠れた値'], ['C1', 'これも']].map(([a, v]) => `<c r="${a}" t="inlineStr"><is><t>${v}</t></is></c>`).join('') +
    '</row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>';
  const result = fromXlsx(miniXlsx({sheet})), cells = cellsOf(result);
  assert.deepEqual([cells.A1.v, cells.B1, cells.C1], ['表示', undefined, undefined]);
  assert.ok(result.warnings.includes('シート「S」：結合で隠れたセル（B1、C1）の値は、Excelでも表示されないので読み込まなかった'));
});

test('列や行に付けた書式（列全体・行全体の書式）は、値のない空のセルにも付ける（行の書式が列の書式より優先、セル自身の書式が最優先）', () => {
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>', '<fill><patternFill patternType="solid"><fgColor rgb="FF00B0F0"/></patternFill></fill>'];
  const sheet = '<cols><col min="2" max="3" width="9" style="1" customWidth="1"/></cols><sheetData><row r="2" s="2" customFormat="1"><c r="A2" t="inlineStr"><is><t>a</t></is></c></row>' +
    '<row r="4"><c r="B4" t="inlineStr"><is><t>own</t></is></c></row></sheetData>';
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: sheet}], styles: stylesXml({fills, xfs: [xf({fill: 2}), xf({fill: 3})]})})), cells = cellsOf(result);
  const fill = a => cells[a]?.s?.fill;
  assert.deepEqual(['B1', 'C1', 'B3', 'C30'].map(fill), ['#ffff00', '#ffff00', '#ffff00', '#ffff00'], '列の書式は、その列の空のセル全部');
  assert.deepEqual(['B2', 'C2', 'D2', 'J2'].map(fill), ['#00b0f0', '#00b0f0', '#00b0f0', '#00b0f0'], '行の書式は、列の書式のある列でも行の書式');
  assert.deepEqual(['A2', 'B4', 'A1', 'D1'].map(fill), [undefined, undefined, undefined, undefined], '書式を持つセル（s なしは標準）と、書式のない列・行は塗らない');
  assert.deepEqual(cells.B4, {t: 's', v: 'own'});
});

test('列幅は、シートの既定の幅（baseColWidth・defaultColWidth）も読み、HTMLの範囲（24〜2044px）の外の幅は収めて知らせる', () => {
  const read = (format, cols = '') => fromXlsx(miniXlsx({sheet: `${format}${cols}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`}));
  assert.equal(read('<sheetFormatPr baseColWidth="10" defaultRowHeight="18.75"/>').workbook.sheets[0].colWidths[0], 88, '数字10文字分と余白を、8pxの倍数に切り上げる');
  assert.equal(read('<sheetFormatPr defaultColWidth="12" defaultRowHeight="18.75"/>').workbook.sheets[0].colWidths[0], 96);
  const clamped = read('<sheetFormatPr defaultRowHeight="18.75"/>', '<cols><col min="2" max="2" width="1" customWidth="1"/><col min="3" max="3" width="300" customWidth="1"/></cols>');
  assert.deepEqual(clamped.workbook.sheets[0].colWidths.slice(1, 3), [24, 2044]);
  assert.ok(warned(clamped, /列（B、C）の幅は、HTMLの範囲（24〜2044px）に収めた/));
  assert.ok(clamped.notes.S.sheet.some(n => n.item === 'colWidth' && n.cols.join() === 'B,C' && n.reason === '幅の範囲'));
});

test('行の高さの既定は、Excelと同じく、シートの既定の高さが決めてあればその高さ、なければ標準のフォントに合う高さにする', () => {
  // デスクトップ版Excel 16.0 で測った規則。決めていない既定の高さ（customHeight のない defaultRowHeight）は使わず、行を標準のフォントに
  // 合わせる。既定の高さを書いていないシートでは、空の行はExcelの既定（游ゴシック11ptの18.75pt）。高さは表示の倍率100%での値。
  const read = (format, normal = '游ゴシック') => fromXlsx(miniXlsx({styles: stylesXml({fonts: [`<font><sz val="11"/><name val="${normal}"/></font>`]}),
    sheet: `<sheetViews><sheetView workbookViewId="0"/></sheetViews>${format}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`}));
  const heights = (format, normal) => read(format, normal).workbook.sheets[0].rowHeights.slice(0, 2);
  assert.deepEqual(heights('<sheetFormatPr defaultRowHeight="20" customHeight="1"/>'), [27, 27], '決めた高さ（20pt）は、そのまま');
  assert.deepEqual(heights('<sheetFormatPr defaultRowHeight="20"/>'), [25, 25], '決めていない高さは使わず、游ゴシック11ptに合う18.75pt');
  assert.deepEqual(heights('<sheetFormatPr defaultRowHeight="20"/>', 'Calibri'), [20, 20], 'Calibri 11ptに合う高さは15pt');
  assert.deepEqual(heights('', 'Calibri'), [20, 25], '既定の高さのないシートでは、値のある行はフォントに合う高さ、空の行はExcelの既定');
  assert.deepEqual(heights('', 'ＭＳ Ｐゴシック'), [18, 25], 'ＭＳ Ｐゴシック11ptに合う高さは13.5pt');
  assert.ok(!read('').notes.S.sheet.some(n => n.item === 'rowHeightAll'));
  const unknown = read('<sheetFormatPr defaultRowHeight="19.5"/>', 'Foo Sans');
  assert.deepEqual(unknown.workbook.sheets[0].rowHeights.slice(0, 2), [26, 26], '高さを知らないフォントは、ファイルの既定の高さ');
  assert.ok(unknown.notes.S.sheet.some(n => n.item === 'rowHeightAll'), '行の高さが近い値であることを照合に伝える');
});

test('非表示の行（hidden・高さ0）は既定の高さで表示し、HTMLの範囲（16〜546px）より低い行は16pxにして、どちらも知らせる', () => {
  const sheet = '<sheetFormatPr defaultRowHeight="20" customHeight="1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="3" hidden="1"><c r="A3"><v>3</v></c></row>' +
    '<row r="4" ht="0" customHeight="1"/><row r="5" ht="6" customHeight="1"/><row r="6" ht="409.5" customHeight="1"/></sheetData>';
  const result = fromXlsx(miniXlsx({sheet}));
  assert.deepEqual(result.workbook.sheets[0].rowHeights.slice(0, 6), [27, 27, 27, 27, 16, 546], 'シートの既定の高さ20ptは27px');
  assert.ok(warned(result, /非表示の行（3、4 行目）は、HTMLで隠せないので表示した/));
  assert.ok(warned(result, /行（5 行目）の高さは、HTMLの範囲（16〜546px）に収めた/));
  assert.deepEqual(result.notes.S.sheet.filter(n => n.item === 'rowHeight'), [{item: 'rowHeight', rows: [3, 4], reason: '非表示の行'}, {item: 'rowHeight', rows: [5], reason: '高さの範囲'}]);
});

test('枠の固定（列だけ・frozenSplit）を読み、途中からの固定・結合をまたぐ固定・ウィンドウの分割は、移せる形にするか外して知らせる', () => {
  const read = (view, after = '') => fromXlsx(miniXlsx({sheet: `<sheetViews>${view}</sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${after}`}));
  assert.deepEqual(read('<sheetView workbookViewId="0"><pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/></sheetView>').workbook.sheets[0].freeze, {rows: 0, cols: 2});
  assert.deepEqual(read('<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozenSplit"/></sheetView>').workbook.sheets[0].freeze, {rows: 1, cols: 0});
  const scrolled = read('<sheetView workbookViewId="0" topLeftCell="A5"><pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView>');
  assert.deepEqual(scrolled.workbook.sheets[0].freeze, {rows: 2, cols: 0});
  assert.ok(warned(scrolled, /固定した行・列は、表の先頭からの固定にした（Excelでは途中から固定していた）/));
  const crossing = read('<sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView>', '<mergeCells count="1"><mergeCell ref="A2:A3"/></mergeCells>');
  assert.equal(crossing.workbook.sheets[0].freeze, undefined);
  assert.ok(warned(crossing, /固定の境目をまたぐ結合セルがあるので、枠の固定は移せなかった/));
  assert.ok(crossing.notes.S.sheet.some(n => n.item === 'freeze'));
  const split = read('<sheetView workbookViewId="0"><pane xSplit="2000" ySplit="1000" topLeftCell="C5" activePane="bottomRight"/></sheetView>');
  assert.equal(split.workbook.sheets[0].freeze, undefined);
  assert.ok(warned(split, /ウィンドウの分割は移せない/));
});

test('選択は最初の範囲を読み、表の外のアクティブなセルは読まず、倍率はHTMLの範囲（50〜200%）に収めて知らせ、枠線の false も読む', () => {
  const read = view => fromXlsx(miniXlsx({sheet: `<sheetViews>${view}</sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`}));
  const multi = read('<sheetView showGridLines="false" workbookViewId="0"><selection activeCell="B2" sqref="B2:C3 E5"/></sheetView>');
  assert.deepEqual(multi.workbook.view, {zoom: 100, active: [1, 1], selection: {r1: 1, c1: 1, r2: 2, c2: 2}});
  assert.equal(multi.workbook.sheets[0].gridlines, false);
  assert.deepEqual(read('<sheetView workbookViewId="0"><selection activeCell="Z900" sqref="Z900"/></sheetView>').workbook.view, {zoom: 100});
  for (const [zoom, want] of [[300, 200], [25, 50]]) {
    const result = read(`<sheetView zoomScale="${zoom}" workbookViewId="0"/>`);
    assert.equal(result.workbook.view.zoom, want);
    assert.deepEqual(result.warnings, [`シート「S」：倍率 ${zoom}% は ${want}% にした（HTMLは50〜200%）`], '同じことを2回知らせない');
    assert.deepEqual(result.notes.S.sheet.filter(n => n.item === 'zoom'), [{item: 'zoom', reason: 'HTMLの倍率は50〜200%'}]);
  }
});

test('条件付き書式・入力規則・フィルター・テーブル・シートの保護・コメント・改ページは移せないことを知らせ、非表示のシートは表示して知らせる', () => {
  const data = '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>';
  const all = data + '<sheetProtection sheet="1" objects="1" scenarios="1"/><autoFilter ref="A1:A1"/>' +
    '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>' +
    '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="A1"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>' +
    '<rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks><legacyDrawing r:id="rId1"/><tableParts count="1"><tablePart r:id="rId2"/></tableParts>';
  // Excel 2010 以降の機能は、シートの extLst に入る。
  const x14 = '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups/></ext>' +
    '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext>' +
    '<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations/></ext>' +
    '<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList/></ext></extLst>';
  const result = fromXlsx(handXlsx({sheets: [{name: '全部', xml: all, rels: [['rId3', 'pivotTable', '../pivotTables/pivotTable1.xml']]},
    {name: '新しい機能', xml: data.replace('</row>', '<c r="B1" t="e" vm="1"><v>#VALUE!</v></c></row>') + '<controls><control shapeId="1025" r:id="rId4"/></controls>' + x14},
    {name: '新しいコメント', xml: data, rels: [['rId1', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml']]},
    {name: '列の改ページ', xml: data + '<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>'},
    {name: '隠し', xml: data, state: 'hidden'}, {name: 'とても隠し', xml: data, state: 'veryHidden'}]}));
  for (const label of ['条件付き書式', '入力規則', 'フィルター', 'テーブル', 'シートの保護', 'コメント（メモ）', '改ページ'])
    assert.ok(result.warnings.includes(`シート「全部」：${label}は移せない`), label);
  assert.ok(result.warnings.includes('シート「全部」：ピボットテーブルは移せない（表示されていた値だけを移した）'));
  for (const label of ['スパークライン', '条件付き書式', '入力規則', 'スライサー', 'フォームのコントロール'])
    assert.ok(result.warnings.includes(`シート「新しい機能」：${label}は移せない`), label);
  assert.ok(result.warnings.includes('シート「新しい機能」：セルの中の画像やデータ型（B1）は移せない'));
  assert.equal(result.workbook.sheets[1].cells.B1, undefined, 'セルの中の画像のセルは、Excelが保存したエラーの値（#VALUE!）を表示せず、空にする');
  assert.deepEqual(result.notes['新しい機能'].cells.B1, [{fields: ['value2', 'text'], reason: 'セルの中の画像やデータ型'}]);
  assert.ok(result.warnings.includes('シート「新しいコメント」：コメント（メモ）は移せない'));
  assert.ok(result.warnings.includes('シート「列の改ページ」：改ページは移せない'));
  assert.deepEqual(result.workbook.sheets.map(s => s.name), ['全部', '新しい機能', '新しいコメント', '列の改ページ', '隠し', 'とても隠し']);
  for (const name of ['隠し', 'とても隠し']) assert.ok(result.warnings.includes(`シート「${name}」：Excelでは非表示のシートだが、HTMLでは表示した`), name);
});

test('画像は、セルからのずれで置いた画像・絶対位置の画像・JPEG・GIFを移し、同じ画像は1つにまとめ、移せない画像・図形・グラフ・切り抜き・回転は知らせる', () => {
  const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"', EMU = 9525;
  const from = (col, colOff, row, rowOff) => `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOff * EMU}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff * EMU}</xdr:rowOff></xdr:from>`;
  const to = '<xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>';
  const pic = (id, embed, w, h, {crop = '', rot = ''} = {}) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="図 ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${embed}"/>${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm${rot}><a:off x="0" y="0"/>` +
    `<a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
  const one = (col, body) => `<xdr:oneCellAnchor>${from(col, 0, 0, 0)}<xdr:ext cx="0" cy="0"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
  const drawing = `<xdr:wsDr ${XDR}>` +
    `<xdr:twoCellAnchor editAs="oneCell">${from(1, 2, 2, 1)}${to}${pic(2, 'rId1', 100, 50)}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:absoluteAnchor><xdr:pos x="${200 * EMU}" y="${100 * EMU}"/><xdr:ext cx="${50 * EMU}" cy="${50 * EMU}"/>${pic(3, 'rId1', 50, 50)}<xdr:clientData/></xdr:absoluteAnchor>` +
    one(4, pic(4, 'rId2', 30, 30, {crop: '<a:srcRect l="10000"/>', rot: ' rot="5400000"'})) + one(6, pic(5, 'rId3', 30, 30)) + one(7, pic(6, 'rId4', 2, 2)) +
    `<xdr:twoCellAnchor>${from(1, 0, 8, 0)}${to}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7" name="四角形 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${from(1, 0, 9, 0)}${to}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="8" name="グラフ 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
    one(9, pic(9, 'rId4', 20, 20)) + '</xdr:wsDr>';
  const media = {png: Buffer.from('png のかわり'), jpg: Buffer.from('jpeg のかわり'), emf: Buffer.from('emf のかわり'), gif: Buffer.from('gif のかわり')};
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: '<sheetData/><drawing r:id="rId1"/>', rels: [['rId1', 'drawing', '../drawings/drawing1.xml']]}], parts: {
    'xl/drawings/drawing1.xml': drawing,
    'xl/drawings/_rels/drawing1.xml.rels': relsXml([['rId1', 'image', '../media/image1.png'], ['rId2', 'image', '../media/image2.jpg'], ['rId3', 'image', '../media/image3.emf'], ['rId4', 'image', '../media/image4.gif']]),
    'xl/media/image1.png': media.png, 'xl/media/image2.jpg': media.jpg, 'xl/media/image3.emf': media.emf, 'xl/media/image4.gif': media.gif,
  }}));
  const {workbook: w} = result;
  assert.deepEqual(w.sheets[0].images, [{id: 'pic-1', asset: 'img-1', r: 2, c: 1, x: 2, y: 1, w: 100, h: 50}, {id: 'pic-2', asset: 'img-1', r: 0, c: 0, x: 200, y: 100, w: 50, h: 50},
    {id: 'pic-3', asset: 'img-2', r: 0, c: 4, x: 0, y: 0, w: 30, h: 30}, {id: 'pic-4', asset: 'img-3', r: 0, c: 9, x: 0, y: 0, w: 20, h: 20}], '同じ画像の中身は、1つの画像データを使う');
  assert.deepEqual(w.assets, {'img-1': 'data:image/png;base64,' + media.png.toString('base64'), 'img-2': 'data:image/jpeg;base64,' + media.jpg.toString('base64'),
    'img-3': 'data:image/gif;base64,' + media.gif.toString('base64')});
  for (const text of ['画像の切り抜きは移せないので、元の画像全体を表示した', '画像の回転は移せない', '画像 1 個は、HTMLで扱えない形式（EMF・BMPなど）なので移せなかった',
    '幅か高さが4px未満の画像 1 個は移せなかった', '図形・グラフなど 2 個は移せない']) assert.ok(result.warnings.includes('シート「S」：' + text), text);
  assert.ok(result.notes.S.sheet.some(n => n.item === 'pictures'));
});

test('互換のための書き方（mc:AlternateContent）で包んだ図形も移せないことを知らせ、包まれた画像は1つ目の選択肢を移す', () => {
  const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const anchor = (col, body) => `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="190500" cy="190500"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
  const pic = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="図 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic>';
  const shape = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="四角形 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>';
  const wrap = (choice, fallback) => `<mc:AlternateContent><mc:Choice Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: '<sheetData/><drawing r:id="rId1"/>', rels: [['rId1', 'drawing', '../drawings/drawing1.xml']]}], parts: {
    'xl/drawings/drawing1.xml': `<xdr:wsDr ${XDR}>${wrap(anchor(1, shape), anchor(1, shape))}${wrap(anchor(3, pic), anchor(5, pic))}</xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': relsXml([['rId1', 'image', '../media/image1.png']]), 'xl/media/image1.png': Buffer.from('png のかわり')}}));
  assert.deepEqual(result.workbook.sheets[0].images.map(p => [p.r, p.c, p.w]), [[0, 3, 20]], '包まれた画像は、1つ目の選択肢（mc:Choice）の位置に1つ');
  assert.ok(result.warnings.includes('シート「S」：図形・グラフなど 1 個は移せない'));
});

test('行やセルの番地（r）のない書き方、絶対パスで書いた部品の参照先、数値文字参照も読む', () => {
  const sheet = '<sheetData><row><c><v>1</v></c><c t="inlineStr"><is><t>&#x41;&#66;</t></is></c></row><row><c><v>3</v></c></row><row r="5"><c r="C5"><v>5</v></c><c><v>6</v></c></row></sheetData>';
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: sheet}], parts: {'xl/_rels/workbook.xml.rels':
    relsXml([['rId1', 'worksheet', '/xl/worksheets/sheet1.xml'], ['rId2', 'styles', '/xl/styles.xml'], ['rId3', 'sharedStrings', '/xl/sharedStrings.xml']])}}));
  const cells = cellsOf(result);
  assert.deepEqual([cells.A1, cells.B1, cells.A2, cells.C5, cells.D5], [{t: 'n', v: 1}, {t: 's', v: 'AB'}, {t: 'n', v: 3}, {t: 'n', v: 5}, {t: 'n', v: 6}],
    '番地のない行は前の行の次、番地のないセルは前のセルの右');
});

test('数式のまま移した数式は、Excelが保存したエラーの値とHTMLの計算が同じなら知らせず、違えば知らせる', () => {
  const sheet = '<sheetData><row r="1"><c r="A1"><v>0</v></c><c r="B1" t="e"><f>1/A1</f><v>#DIV/0!</v></c><c r="C1" t="e"><f>A1+1</f><v>#N/A</v></c></row></sheetData>';
  const result = fromXlsx(miniXlsx({sheet}));
  assert.deepEqual([cellsOf(result).B1.f, cellsOf(result).C1.f], ['=1/A1', '=A1+1']);
  assert.ok(!warned(result, /S!B1/), 'HTMLも #DIV/0!');
  assert.ok(warned(result, /HTMLで計算した値が、Excelが保存していた値と違う数式がある: S!C1（HTML 1、Excel "#N\/A"）/));
});

test('題名は文書のプロパティの題名、なければファイル名にし、ブックのIDはファイル名と中身から決め、開いていたシートを最初に表示する', () => {
  const sheets = [{name: '一', xml: '<sheetData/>'}, {name: '二', xml: '<sheetData/>'}];
  const titled = fromXlsx(handXlsx({sheets, core: '<dc:title>見積書 2026</dc:title>', workbook: '<bookViews><workbookView activeTab="1"/></bookViews>'}), {name: 'Quote 2026.xlsx'});
  assert.deepEqual([titled.workbook.title, titled.workbook.revision, titled.workbook.activeSheet], ['見積書 2026', 'draft', 's2']);
  assert.match(titled.workbook.id, /^xlsx-quote-2026-[0-9a-f]{8}$/);
  const untitled = fromXlsx(handXlsx({sheets, workbook: '<bookViews><workbookView activeTab="7"/></bookViews>'}), {name: '見積.xlsx'});
  assert.equal(untitled.workbook.title, '見積');
  assert.match(untitled.workbook.id, /^xlsx-[0-9a-f]{8}$/, '英数字のないファイル名は、中身から決めた部分だけ');
  assert.equal(untitled.workbook.activeSheet, 's2', 'シートの数より大きい番号は、最後のシート');
  const idOf = xml => fromXlsx(handXlsx({sheets: [{name: 'S', xml}]}), {name: 'a.xlsx'}).workbook.id;
  assert.equal(idOf('<sheetData/>'), idOf('<sheetData/>'), '同じファイルからは同じID');
  assert.notEqual(idOf('<sheetData/>'), idOf('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'), '中身が違えば違うID');
});

test('グラフのシートは移さずに知らせ、開いていたシートの番号はワークシートだけで数え直す', () => {
  const read = tab => fromXlsx(handXlsx({sheets: [{name: 'グラフ1', xml: '', type: 'chartsheet'}, {name: '表', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'},
    {name: '表2', xml: '<sheetData/>'}], workbook: `<bookViews><workbookView activeTab="${tab}"/></bookViews>`}));
  const result = read(1);
  assert.deepEqual(result.workbook.sheets.map(s => s.name), ['表', '表2']);
  assert.equal(result.workbook.activeSheet, 's1', '2番目のシート（表）を開いていた');
  assert.ok(result.warnings.includes('シート「グラフ1」：グラフのシートなので移せない'));
  assert.equal(read(0).workbook.activeSheet, 's1', 'グラフのシートを開いていたなら、その次のワークシート');
});

test('ブックやワークシートのないファイル・31枚以上のシート・ZIP64・暗号化・未対応の圧縮・壊れた目次は、理由を示して止める', () => {
  const noBook = makeZip({'[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']])});
  assert.throws(() => fromXlsx(noBook), /xlsx の中にブックがありません/);
  assert.throws(() => fromXlsx(handXlsx({sheets: [{name: 'グラフ1', xml: '', type: 'chartsheet'}]})), /xlsx にワークシートがありません/);
  assert.throws(() => fromXlsx(handXlsx({sheets: Array.from({length: 31}, (_, i) => ({name: 'S' + (i + 1), xml: '<sheetData/>'}))})), /シートが 31 枚あり、HTMLの上限（30枚）を超える/);
  const good = miniXlsx({sheet: '<sheetData/>'}), end = good.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])), central = good.readUInt32LE(end + 16);
  const broken = edit => { const b = Buffer.from(good); edit(b); return b; };
  assert.throws(() => fromXlsx(broken(b => b.writeUInt32LE(0xffffffff, end + 16))), /ZIP64 の xlsx には対応していません/);
  assert.throws(() => fromXlsx(broken(b => b.writeUInt16LE(1, central + 8))), /暗号化された xlsx は読めません/);
  assert.throws(() => fromXlsx(broken(b => b.writeUInt16LE(12, central + 10))), /未対応の圧縮方式の xlsx です: 12/);
  assert.throws(() => fromXlsx(broken(b => b.writeUInt32LE(0, central))), /目次が壊れています/);
});

test('書式・テーマ・共有文字列の部品がないファイルも読み、書式は日本語版Excelの標準のスタイル（游ゴシック11pt・上下中央）にする', () => {
  const result = fromXlsx(handXlsx({sheets: [{name: 'S', xml: '<sheetData><row r="1"><c r="A1" s="3"><v>5</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData>'}], styles: null, strings: null}));
  assert.deepEqual(cellsOf(result), {A1: {t: 'n', v: 5}}, 'セルの書式も標準のスタイル');
  assert.deepEqual(result.workbook.defaultStyle, {font: 'Yu Gothic', size: 11, valign: 'middle'});
});

test('標準のフォントが游ゴシック11ptでないブックは、そのフォントの数字の幅で列幅を求め、列幅が近い値であることを照合に伝える', () => {
  const read = (name, size, normal) => fromXlsx(miniXlsx({styles: stylesXml({fonts: [`<font><sz val="${size}"/><name val="${name}"/></font>`], ...(normal ? {normal} : {})}),
    sheet: '<cols><col min="2" max="2" width="10" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'}));
  const widths = (name, size) => read(name, size).workbook.sheets[0].colWidths.slice(0, 2);
  assert.deepEqual(widths('Calibri', 11), [64, 70], 'Calibri 11pt の数字の幅は7px（96dpi）。既定の列幅は、Excelと同じ64px');
  assert.deepEqual(widths('メイリオ', 11), [80, 90], 'メイリオ 11pt は9px');
  assert.deepEqual(widths('游ゴシック', 10), [64, 70], '大きさに比例させて四捨五入（8 × 10 ÷ 11 → 7px）');
  assert.deepEqual(widths('Foo Sans', 11), [72, 80], '数字の幅を知らないフォントは、游ゴシックと同じ8px');
  const calibri = read('Calibri', 11, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
  assert.deepEqual(calibri.workbook.defaultStyle, {font: 'Calibri', size: 11, valign: 'bottom'}, '縦位置の指定のない標準のスタイルは下揃え');
  assert.ok(calibri.notes.S.sheet.some(n => n.item === 'colWidthAll'));
  assert.ok(read('游ゴシック', 10).notes.S.sheet.some(n => n.item === 'colWidthAll'), '游ゴシックでも11ptでなければ、近い値');
  assert.ok(!read('游ゴシック', 11).notes.S.sheet.some(n => n.item === 'colWidthAll'), '游ゴシック11ptなら、どの表示の倍率でもExcelと同じ幅');
});
