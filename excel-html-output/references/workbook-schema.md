# ブックJSON（excel-html-workbook/v1）

`scripts/build.py --input` に渡すJSONの形式である。
ビルドとブラウザーは同じ検査（`assets/core.js` の `validate`）を使うので、ここに書いた制約を外れると、ビルドの時点で止まる。
ここに書いていない項目を持つブックも止まる。
例えば、シートに `hidden` を書くと、`未対応の項目 sheets[0].hidden` で止まる。
画面に出ない項目は、xlsxに書き出すと失われ、セルの項目はそのセルを画面で編集しても失われるので、黙って持たずに止める。

```json
{
  "schema": "excel-html-workbook/v1",
  "id": "estimate-2026-acme",
  "revision": "draft",
  "title": "見積明細",
  "activeSheet": "main",
  "defaultStyle": {"font": "Yu Gothic", "size": 11},
  "view": {"zoom": 100},
  "sheets": [{
    "id": "main", "name": "見積", "rows": 60, "cols": 12,
    "colWidths": [90, 220, 70, 100, 110], "rowHeights": [32, 24, 24, 44],
    "freeze": {"rows": 3, "cols": 0}, "gridlines": true,
    "merges": [{"r1": 0, "c1": 0, "r2": 0, "c2": 4}],
    "cells": {
      "A1": {"t": "s", "v": "見積明細", "s": {"bold": true, "size": 16, "align": "center"}},
      "A3": {"t": "s", "v": "品番", "s": {"bold": true, "fill": "#e2efda", "borderBottom": "2px solid #217346"}},
      "A4": {"t": "s", "v": "00123", "s": {"numFmt": "@"}},
      "B4": {"t": "s", "v": "保守作業\n（月額）", "s": {"wrap": true}},
      "C4": {"t": "n", "v": 3},
      "D4": {"t": "n", "v": 12000, "s": {"numFmt": "¥#,##0"}},
      "E4": {"t": "f", "f": "=C4*D4", "s": {"numFmt": "¥#,##0"}}
    }
  }]
}
```

この例では、B4 の改行を見せるために `rowHeights[3]` を44にしている。
セルに収まらない文字を、余白を詰めたり文字を小さくしたりしてごまかさない。

## ブックとシート

- `id` はブックの恒久的なIDで、作り直しても変えない。
  `revision` はビルドのたびに内容のSHA-256から作り直されるので、`"draft"` のままでよい。
  ユーザーから戻ってきた編集状態は、このIDとリビジョンで元のHTMLと照合する。
- `activeSheet` は最初に開くシートのID。
  シートは1〜30枚。
  シート名はExcelと同じく31文字以内で、`\ / ? * [ ] :` を含まず、重複しない。
- `rows` と `cols` は表示する行数と列数で、上限は500行・100列。
  全セルを描画するので、大きな表は表示速度を確かめる。
- `colWidths` は列幅、`rowHeights` は行の高さで、CSSピクセルで書く。
  既定は列幅96px、行の高さ22px。
  `null` や省略は既定になる。
  幅は24〜2044、高さは16〜546。
  上限は、Excelで設定できる最大の列幅と行の高さに合わせている。
- `merges` は結合で、両端を含む0始まりの矩形（`r1,c1,r2,c2`）。
  重複は不可。
  左上以外のセルに値を置かない。
  結合範囲の書式（塗り・罫線を含む）は左上のセルに書く。
  左上以外のセルの書式は画面に出ず、Excelにも渡らない。
  結合が固定行・固定列の境界をまたぐことはできない。
- `freeze` は先頭から固定する行数と列数。
  `gridlines` は true か false で、false で補助の枠線を隠す（セルの罫線は残る）。
- `view` は `zoom`（50〜200）、`active:[行,列]`、`selection:{r1,c1,r2,c2}`、`scrollTop`、`scrollLeft`。
  編集状態を書き出すと、その時点の位置が入る。

## セル

セルはA1形式の番地をキーにした辞書で、空のセルは書かなくてよい。

| t | 値 | 使いどころ |
|---|---|---|
| `s` | `v`: 文字列 | 文字。計算しない番号（管理番号、郵便番号、電話番号、コード）もここ。先頭ゼロがあれば `numFmt:"@"` |
| `n` | `v`: 有限の数値 | 数量・金額・比率・日付。表示は `numFmt` で決める |
| `b` | `v`: true / false | TRUE / FALSE |
| `f` | `f`: `=` で始まる式 | 合計・小計など。値はHTMLが計算する（キャッシュは持たない） |

- 日付は、1899-12-30 を0日目とするシリアル値（2026-01-01 は 46023）に `yyyy/mm/dd` を付ける。
  1900年のうるう年の互換やタイムゾーン付きの日時は扱わない。
- 比率は 0.25 のように小数で持ち、`0%` か `0.00%` を付ける。
- `link` に http(s) のURLを書くとリンクになる。
- セルの文字列はHTMLやJavaScriptとして実行されない。

## 書式（`s`）

書式は「既定 → `workbook.defaultStyle` → `sheet.defaultStyle` → セルの `s`」の順に重ねる。
既定は日本語版Excelの標準と同じである。

| キー | 値 | 既定 |
|---|---|---|
| `font` | フォント名（埋め込まない） | `Yu Gothic` |
| `size` | 6〜96（pt） | 11 |
| `bold`, `italic`, `underline`, `strike` | true / false | false |
| `color` | `#rrggbb` | `#000000` |
| `fill` | `#rrggbb`、または `none` | `none`（塗りなし） |
| `align` | `general` / `left` / `center` / `right` | `general`（数値は右、文字は左、真偽値は中央） |
| `valign` | `top` / `middle` / `bottom` | `middle` |
| `wrap` | true / false | false |
| `indent` | 0〜15（`align` は `left` か `right` と組み合わせる） | 0 |
| `borderTop`, `borderRight`, `borderBottom`, `borderLeft` | `none` または下の表の線＋色（例 `1px solid #808080`） | `none` |
| `numFmt` | `General` / `@` / `0` / `0.00` / `#,##0` / `#,##0.00` / `0%` / `0.00%` / `¥#,##0` / `yyyy/mm/dd` | `General` |

罫線は、Excelで表せる次の線だけを受け付ける。

| 値 | Excelの線 |
|---|---|
| `0.5px solid` | 極細 |
| `1px solid` | 細線 |
| `2px solid` | 中太線 |
| `3px solid` | 太線 |
| `1px dashed` / `2px dashed` | 破線（細）/ 破線（中） |
| `1px dotted` | 点線 |
| `3px double` | 二重線 |

Excelはセルの辺ごとに線を1本しか持てない。
隣り合うセルの境界（A1の下とA2の上など）に違う線を書くと、Excelではどちらか一方になる。
境界の線は片方のセルにだけ書くか、同じ線にそろえる。

一覧にない表示形式や線を通すと、画面とExcelの表示がずれる。
増やすときは、画面の表示（`display`・`css`）、クリップボードの書き出し（`xlCss`）、読み取り（`styleFromDecls`）を同時に実装し、`verify_excel_paste.py` で確かめる。

## 数式

HTMLで計算できるのは、同じシートのA1参照（`$` 付きを含む）と範囲、数値・文字列・`TRUE`/`FALSE`・エラー値（`#N/A` など）、演算子、括弧、次の139の関数である。
数式の読み取りと計算は `assets/formulas.js` が行い、式の評価に `eval` や `Function` は使っていない。

| 分類 | 関数 |
|---|---|
| 論理・情報 | IF, IFS, IFERROR, IFNA, AND, OR, XOR, NOT, TRUE, FALSE, NA, SWITCH, CHOOSE, ISBLANK, ISNUMBER, ISTEXT, ISNONTEXT, ISLOGICAL, ISERROR, ISERR, ISNA, ISEVEN, ISODD, N, T |
| 数学 | SUM, PRODUCT, SUMPRODUCT, ROUND, ROUNDUP, ROUNDDOWN, TRUNC, INT, ABS, MOD, QUOTIENT, SIGN, SQRT, POWER, EXP, LN, LOG, LOG10, PI, CEILING, FLOOR, CEILING.MATH, FLOOR.MATH, MROUND |
| 統計 | AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, MEDIAN, MODE, MODE.SNGL, LARGE, SMALL, RANK, RANK.EQ, STDEV, STDEV.S, STDEVP, STDEV.P, VAR, VAR.S, VARP, VAR.P, SUBTOTAL |
| 条件つきの集計 | COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS |
| 検索・参照 | VLOOKUP, HLOOKUP, XLOOKUP, LOOKUP, MATCH, XMATCH, INDEX, ROW, COLUMN, ROWS, COLUMNS |
| 文字列 | LEN, LEFT, RIGHT, MID, FIND, SEARCH, SUBSTITUTE, REPLACE, TRIM, CLEAN, UPPER, LOWER, PROPER, REPT, EXACT, CHAR, CODE, UNICHAR, UNICODE, ASC, JIS（DBCS）, VALUE, TEXT, CONCATENATE, CONCAT, TEXTJOIN |
| 日付・時刻 | TODAY, NOW, DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, ISOWEEKNUM, EDATE, EOMONTH, DATEDIF, DAYS, NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL, DATEVALUE, TIMEVALUE |

どの関数も、デスクトップ版Excelで同じ数式を計算させた値と比べて、同じ値を返すように作ってある（[Excelとの互換性](excel-compatibility.md#関数の計算)）。
引数の扱いもExcelと同じである。
例えば、数字だけの文字列（`"2"`、`"2026/1/2"`、`"12:30"`）は計算で数値として使われ、空のセルは計算では0、文字では空の文字になる。
引数は省略して空にしてもよい（`=IF(A1>0,"済",)` の偽の値は0）。
演算子の優先順位もExcelと同じで、符号の `-` はべき乗より先に計算する（`=-2^2` は4）。
同じ優先順位の演算子は左から順に計算する（`=2^3^2` は64）。
表にない関数は `#NAME?` と表示するので使わない。
TEXTBEFORE や TEXTAFTER のように、Microsoft 365 にだけある関数も表にない。
測ったExcel（Excel 2021）にはなく、貼り付けた先で `#NAME?` になるからである。

範囲は、関数の範囲の引数（SUM、COUNTIF、VLOOKUP の表など）に渡す。
1つの値を受け取る場所（`&` や IF の条件など）に範囲を書くと、Excelの数式と同じく、数式と同じ行（横に並んだ範囲なら同じ列）のセルの値だけを使う。
範囲どうしの計算は SUMPRODUCT の中でだけできる。
例えば、`=SUMPRODUCT((B2:B10="東京")*C2:C10)` は、B列が「東京」の行のC列を合計する。
列全体や行全体の参照（`B:B`、`2:2`）、配列定数（`{1,2,3}`）、別のシートの参照、名前、結果が複数のセルに広がる数式（動的配列）は使えない。

日付は、1900年1月1日を1とする通し番号（シリアル値）で計算する。
`=DATE(2026,1,2)` や `=TODAY()` の結果を日付として見せるには、セルに `numFmt:"yyyy/mm/dd"` を付ける。
Excelはこれらの数式を入力したセルに日付の表示形式を自動で付けるが、HTMLは付けないので、付けないと `46024` のような数で表示される。
TODAY と NOW は、計算したときのパソコンの時計の日付と時刻を返す。

文字をつなぐ数式（`&`・CONCATENATE・CONCAT・TEXTJOIN）は、Excelと同じ文字を返す。
数値は有効数字15桁の文字になり、セルの表示形式は使わない。
そのため、2026/01/01と表示した日付をつなぐと `46023`（シリアル値）になり、12.5%と表示した比率は `0.125` になる。
表示どおりの文字でつなぐには、TEXT関数でその文字にする（`=TEXT(A2,"yyyy年m月d日")&"納品"`、`=TEXT(B2,"#,##0")&"円"`）。
TEXT関数は、Excelの表示形式の書き方（和暦の `ggge`、曜日の `aaa` を含む）を読む。
範囲のセルをまとめてつなぐときは、CONCAT か TEXTJOIN を使う（`=TEXTJOIN("、",TRUE,B2:D2)` は空のセルを除いて「、」でつなぐ）。
細かい規則は [Excelとの互換性](excel-compatibility.md#文字をつなぐ数式) にある。

不正な引数は `#VALUE!` のようにExcelと同じエラーを表示し、循環参照は `#CIRC!` と表示する。

Excelに貼ると、参照先がすべてコピー範囲の中にある数式は、数式のまま入る。
範囲の外を参照する数式は、HTMLが計算した値で入る。

## 画像

画像は、Excelと同じくセルの上に浮かぶ画像として、シートの `images` に置く。
画像の中身は、ブックの `assets` に、IDをキーにした data URL で持つ。

```json
"assets": {"img-1": "data:image/png;base64,…"},
"sheets": [{"id": "main", "…": "…", "images": [{"id": "pic-1", "asset": "img-1", "r": 5, "c": 2, "x": 0, "y": 0, "w": 240, "h": 120}]}]
```

- `r`・`c` は、画像の左上を置くセル（0始まり）である。
  `x`・`y` は、そのセルの左上からのずれ（CSSピクセル）である。
  行や列を挿入・削除すると、画像もセルと一緒に動く。
- `w`・`h` は、表示する大きさ（CSSピクセル、4〜10000）である。
- 画像の中身は、png / jpeg / gif / webp の base64 の data URL に限る。
  SVGは、スクリプトを含められるので使えない。
- 1シートに100枚まで、画像1つの data URL は1500万文字まで。
- 画面で貼った画像は、長い辺が2000pxを超えると縮めて保存される。
  置く大きさは、貼ったときに画面で見えていた大きさ（1000×1000に収まるよう縮める）である。

## 編集状態（AIに渡すテキスト）

「AIに渡す」は、既定でTOONの形式で書き出す。
1行目の `EXCEL_HTML_STATE_V1 TOON` で形式を示し、最後の行の `EXCEL_HTML_STATE_END` で、テキストが最後まであることを示す。
その間の行は、下のJSONと同じオブジェクトを、TOON（仕様4.4、[toon-format/spec](https://github.com/toon-format/spec)）で書いたものである。

```text
EXCEL_HTML_STATE_V1 TOON
schema: excel-html-state/v1
workbookId: estimate-2026-acme
baseRevision: "sha256:…"
exportedAt: "2026-11-17T01:23:45.678Z"
workbook:
  schema: excel-html-workbook/v1
  id: estimate-2026-acme
  revision: "sha256:…"
  title: 見積明細
  activeSheet: main
  sheets[1]:
    - id: main
      name: 見積
      rows: 60
      cols: 12
      colWidths[5]: 90,220,70,100,110
      rowHeights[4]: 32,24,24,44
      merges[1]{r1,c1,r2,c2}:
        0,0,0,4
      cells:
        A1:
          t: s
          v: 見積明細
          s:
            bold: true
            size: 16
            align: center
        A4:
          t: s
          v: "00123"
          s:
            numFmt: @
        B4:
          t: s
          v: "保守作業\n（月額）"
          s:
            wrap: true
        E4:
          t: f
          f: =C4*D4
          s:
            numFmt: "¥#,##0"
EXCEL_HTML_STATE_END
```

TOONの読み方は次のとおりである。

- `名前: 値` の1行が1つの項目で、行頭の字下げ（空白2つで1段）が入れ子を表す。
- `colWidths[5]: 90,220,…` は要素が5つの配列である。
  `merges[1]{r1,c1,r2,c2}:` は、同じ項目名を持つオブジェクトを表にしたもので、次の行から1行が1要素になる。
  `[ ]` の数と実際の要素の数が合わないテキストは、途中で切れたか書き換えられたものとして拒む。
- `- ` で始まる行は、配列の要素である。
  要素がオブジェクトなら、最初の項目を `- ` の行に書き、残りの項目をその下にそろえて書く。
- 数・`true`・`false`・`null` と区別がつかない文字列や、`:` `,` `"` `#` などを含む文字列は、`"00123"` のように引用符で囲む。
  改行は `\n`、引用符は `\"` と書く。

ダイアログの「形式」でJSONを選んだときは、次のように書き出す。

```text
EXCEL_HTML_STATE_V1
{"schema":"excel-html-state/v1","workbookId":"…","baseRevision":"sha256:…","exportedAt":"…","workbook":{…ブック全体…}}
```

`build.py --state` は、どちらの形式も読む。
差分ではなく、ブック全体の写しである。
値、型、数式、書式、空セル、結合、寸法、シートの順、表示位置を含む。
画像を貼った場合は、使われている画像の中身（`assets`）も入るので、数MBになることがある。
途中で切れたテキストは拒否し、推測で補わない。
TOONでは、最後の行がないテキストと、`[ ]` の数が合わないテキストを、途中で切れたものとして扱う。
更新しないセルや、知らない業務用の項目を削除しない。
