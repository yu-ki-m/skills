# ブックJSON（excel-html-workbook/v1）

`scripts/build.py --input` に渡すJSONの形式である。
ビルドとブラウザーは同じ検査（`assets/core.js` の `validate`）を使うので、ここに書いた制約を外れると、ビルドの時点で止まる。
ここに書いていない項目を持つブックも止まる。
例えば、シートに `hidden` を書くと、`未対応の項目 sheets[0].hidden` で止まる。
画面に出ない項目は、xlsxに書き出すと失われ、セルの項目はそのセルを画面で編集しても失われるので、黙って持たずに止める。
`--input` のビルドは、その前に [範囲でまとめて書いた所](#範囲でまとめて書く) をセルごとの形に広げ、後で [罫線の食い違い](#書式s) も調べる。

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
    "id": "main", "name": "見積", "rows": 30, "cols": 8,
    "colWidths": [90, 220, 70, 100, 110], "rowHeights": [32, 24, 24, 44],
    "freeze": {"rows": 3, "cols": 0},
    "merges": ["A1:E1"],
    "cells": {
      "A1": {"t": "s", "v": "見積明細", "s": {"bold": true, "size": 16, "align": "center"}}
    },
    "values": [{"at": "A3", "rows": [
      ["品番", "品名", "数量", "単価", "金額"],
      ["00123", "保守作業\n（月額）", 3, 12000],
      ["00124", "部品交換", 1, 4500],
      [null, null, null, "合計"]
    ]}],
    "formulas": [
      {"range": "E4:E5", "f": "=C4*D4"},
      {"range": "E6", "f": "=SUM(E4:E5)"}
    ],
    "styles": [
      {"range": "A3:E6", "s": {"borderTop": "1px solid #808080", "borderRight": "1px solid #808080", "borderBottom": "1px solid #808080", "borderLeft": "1px solid #808080"}},
      {"range": "A3:E3", "s": {"bold": true, "fill": "#e2efda", "borderBottom": "2px solid #217346"}},
      {"range": "A4:A5", "s": {"numFmt": "@"}},
      {"range": "B4", "s": {"wrap": true}},
      {"range": "D4:E6", "s": {"numFmt": "¥#,##0"}},
      {"range": "D6:E6", "s": {"bold": true}}
    ]
  }]
}
```

この例では、B4 の改行を見せるために `rowHeights[3]` を44にしている。
セルに収まらない文字を、余白を詰めたり文字を小さくしたりしてごまかさない。
金額の列は、`values` では書かずに空け（行の終わりのセルは省ける）、`formulas` で数式を入れている。
結合は `"A1:E1"` のように範囲で書いている（`--input` で書ける省略形で、ビルドが `{"r1": 0, "c1": 0, "r2": 0, "c2": 4}` の形にする）。
見出しの下の中太線は、格子の細線の上に重ねて書けば、4行目の上の線も同じ中太線になる。
広げると、例えば E5 は `{"t": "f", "f": "=C5*D5", "s": {…4辺の細線…, "numFmt": "¥#,##0"}}` のセルになる。

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
- `merges` は結合で、両端を含む0始まりの矩形（`{"r1": 0, "c1": 0, "r2": 0, "c2": 4}` は A1:E1）。
  `--input` のブックでは、`"A1:E1"` のように範囲で書いてもよい。
  重複は不可。
  左上以外のセルに値を置かない。
  結合範囲の書式（塗り・罫線を含む）は、`cells` では左上のセルに書き、[styles](#範囲でまとめて書く) では結合セル全体を含む範囲に書く（左上だけの範囲では止まる）。
  左上以外のセルの書式は画面に出ず、Excelにも渡らない。
  結合が固定行・固定列の境界をまたぐことはできない（またぐと、どの結合がどの境目をまたぐかを示して止まる）。
- `freeze` は先頭から固定する行数と列数で、`{"rows": 3, "cols": 0}` のように `rows` と `cols` の両方を書く（固定しない方は 0）。
  固定しないなら `freeze` を書かない。
  `gridlines` は true か false で、false で補助の枠線を隠す（セルの罫線は残る）。
- `view` は `zoom`（50〜200）、`active:[行,列]`、`selection:{r1,c1,r2,c2}`、`scrollTop`、`scrollLeft`。
  編集状態を書き出すと、その時点の位置が入る。

## セル

セルはA1形式の番地をキーにした辞書で、空のセルは書かなくてよい。
例えば `"A4": {"t": "s", "v": "00123", "s": {"numFmt": "@"}}` は、先頭ゼロの品番の文字のセルである。
冒頭の例の題名のように、1つだけのセルは `cells` に書き、表のように続くセルは [範囲でまとめて書く](#範囲でまとめて書く)。
「AIに渡す」の状態テキストと `--extract-json` で取り出したJSONは、すべてのセルをこの形で持つ。

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
| `indent` | 0〜15（`align` は `left` か `right` と組み合わせる。`general` のままなら、Excelと同じく左詰めになり、数値も左に寄る） | 0 |
| `borderTop`, `borderRight`, `borderBottom`, `borderLeft` | `none` または下の表の線＋色（例 `1px solid #808080`） | `none` |
| `numFmt` | `General` / `@` / `0` / `0.00` / `#,##0` / `#,##0.00` / `0%` / `0.00%` / `¥#,##0` / `yyyy/mm/dd` | `General` |

罫線は、Excelで表せる次の線だけを受け付ける。
値は、線のあとに空白と `#rrggbb` の色を続けて書く（例 `3px solid #000000`）。
色のない `3px solid` や、`black`・`#000` のような色は、受け付ける線と色の書き方を示して止まる。

| 線 | 値の例 | Excelの線 |
|---|---|---|
| `0.5px solid` | `0.5px solid #808080` | 極細 |
| `1px solid` | `1px solid #808080` | 細線 |
| `2px solid` | `2px solid #217346` | 中太線 |
| `3px solid` | `3px solid #000000` | 太線 |
| `1px dashed` / `2px dashed` | `1px dashed #000000` | 破線（細）/ 破線（中） |
| `1px dotted` | `1px dotted #000000` | 点線 |
| `3px double` | `3px double #000000` | 二重線 |

Excelはセルの辺ごとに線を1本しか持てない。
隣り合うセルの境界（A1の下とA2の上など）に違う線を書くと、Excelではどちらか一方になる。
境界の線は片方のセルにだけ書くか、同じ線にそろえる。
そろっていないと、`--input` のビルドは「罫線の食い違い」として、場所と2本の線を示して止まる。
[範囲でまとめて書く](#範囲でまとめて書く) `styles` で引いた線は、隣のセルの側の線も同じ線にそろう。
ただし、1つの項目で範囲の向かい合う辺（左と右、上と下）に違う線を書くと、範囲の中の境界が食い違う（外枠だけを太くするなら、辺ごとに1行・1列の範囲で書く）。
画面で編集した後の表（`--state`）とxlsxから読み込んだ表の食い違いは、ユーザーが付けた線なので、止めずに「注意:」で知らせる。
それらを `--extract-json` で取り出して直したJSONから作るときは、`--keep-user-content` を付けると、同じく注意を出して続ける。

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
ビルドは、数式を次のように調べる。

| 数式 | 例 | ビルドの扱い |
|---|---|---|
| HTMLの数式として読めない | 名前（`=A1*税率`）、別のシートの参照（`=Sheet2!A1`）、閉じていないかっこ（`=SUM(A1`） | `--input` は式を示して止まる。編集状態の反映では「注意:」（画面では `#NAME?` か `#VALUE!`、xlsxでは文字列） |
| HTMLでは計算できず、Excelでは計算されることがある | 表にない関数（`#NAME?`）、シートの `rows`・`cols` の外の参照（`#REF!`）、循環参照（`#CIRC!`） | 関数名や番地と直し方を「注意:」で示す |
| エラーのセルを参照しているだけ | 上のセルを参照する `=A1*2` | 「注意:」で示す（参照先を直すと直る） |
| Excelでも同じエラーになる | `=INDEX(B1:B3,5)`、`=#REF!` | 「注意:」で示す |
| 表示形式が「@」のセルにある | `numFmt:"@"` の合計 | 「注意:」で示す（画面やExcelで編集すると文字になる） |

Excelのセルに入らない長さ（文字は32,767文字、数式は8,192文字、数式の中の文字は255文字まで）は、`--input` のビルドが場所を示して止め、編集状態の反映では「注意:」で知らせる。
`--input` のビルドが止まる理由（読めない数式、長さ、罫線の食い違い）は、1回のエラーにまとめて示される。
画面で編集した後の表から取り出したJSONを直して作るときは、`--keep-user-content` を付けると、ユーザーの内容を「注意:」にして作る。

Excelに貼ると、参照先がすべてコピー範囲の中にある数式は、数式のまま入る。
範囲の外を参照する数式は、HTMLが計算した値で入る。

## 範囲でまとめて書く

`build.py --input` に渡すブックJSONでは、シートに `values`・`formulas`・`styles` を書ける（冒頭の例）。
ビルドは、検査の前にこれらを `cells` のセルごとの形に広げる。
埋め込まれるブック、「AIに渡す」の状態テキスト、xlsxは、セルごとに書いたときと同じになる。
同じ書式や同じ形の数式が続く表は、セルごとに書かずにこの形で書く。
明細50行の見積（全セルに罫線、数式153個）は、セルごとに書くと約7.5万字、この形なら約2,400字である。

| 項目 | 書き方 | 広げ方 |
|---|---|---|
| `values` | `{"at": "A3", "rows": [[…], …]}` | `at` のセルを左上に、`rows` を1行ずつ書く。文字は文字のセル、`=` で始まる文字は数式、数値は数値、`true`・`false` は真偽値になる。`null` と `""` は何も書かない（数式を入れる列や、結合で隠れるセルを空けるのに使う）。行の終わりの空けるセルは書かなくてよく、行ごとに長さが違ってよい。セルのオブジェクト（`=` で始まる文字の `{"t": "s", "v": "=で始まる文字"}`、`link` を持つものなど）も置け、`t` は必ず書く |
| `formulas` | `{"range": "E4:E53", "f": "=C4*D4"}` | `f` を範囲の左上のセルの式とし、ほかのセルには、Excelで式を複写したときと同じく、相対参照をずらした式を入れる（`$` を付けた行・列は動かない）。合計の行のように横へも複写できる（`{"range": "E54:G54", "f": "=SUM(E4:E53)"}`） |
| `styles` | `{"range": "A3:G54", "s": {…}}` | 範囲のセルの書式に、書いた順に重ねる。後に書いたものほど優先し、`cells` や `values` のセル自身の `s` にも重なる |

- 数量や金額は、引用符を付けずに数値で書く。
  `"30,000"` や `"2"` のように引用符を付けると文字のセルになり、SUM が数えない。
  `values` に数値に見える文字列があると「数値に見える文字列」、日付や時刻に見える文字列があると「日付や時刻に見える文字列」として、ビルドが「注意:」で知らせる。
  日付として計算するなら、シリアル値に `numFmt:"yyyy/mm/dd"` を付けて書く。
  `-なし` や `@担当` のような、記号で始まるだけの普通の文字は知らせない（Excelにも文字のまま渡る）。
  先頭ゼロの番号のように文字のままにするものは、`styles` で `numFmt:"@"` を付ける（注意も出ない）。
  文字の先頭の `'` は、文字の一部として残る（Excelで入力するときのような、文字にする印にはならない）。
- 罫線は、Excelと同じく、セルの辺ごとに1本として扱う。
  `styles` で辺に線を引くと、隣のセルの同じ辺に別の線があれば、それも同じ線にする（隣のセルに線がなければ何も足さない）。
  隣のセルの線には、ブックやシートの `defaultStyle` の線も含む。
  そのため、表全体に細線の格子を引いてから、見出しの行に `borderBottom` の中太線を重ねても、食い違いにならない。
  線を `none` にすると、隣のセルの側の線も消える。
  1つの項目で、範囲の向かい合う辺（`borderLeft` と `borderRight` など）に違う線を書くと、範囲の中の境界が食い違う。
  外枠だけを太くするなら、上・下・左・右の辺ごとに、1行か1列の範囲で書く。
- 結合セルには、左上のセルに書式が付き、結合セルの辺は1つのセルの辺として扱う。
  範囲が結合セルの一部だけを含むと、どこに付けるかが決まらないので、ビルドが止まる（結合セル全体を含む範囲にするか、範囲から外す）。
  範囲の外の結合セルの辺の一部にだけ接して線を引き、その結合セルの辺に別の線があるときも、辺の残りに同じ線を引かなければ止まる。
  結合セルの辺は1本の線なので、そろえると範囲の外まで線が変わるからである（例: 結合 A5:D6 の上で C4 の下だけに線を引く）。
  辺全体に接する範囲（A4:D4）か、辺全体を覆う複数の範囲（A4:B4 と C4:D4）に同じ線を引くか、結合セル自身（A5:D6 の `borderTop`）に書けば、結合セルの辺もそろう。
- 同じセルに値や数式を2回書くと（`cells` と `values`、`values` どうし、`formulas` とほか）、ビルドが止まる。
  `cells` の書式だけのセル（`{"t": "s", "v": "", "s": {…}}`）には値を重ねられ、その書式は残る。
- 範囲はシートの `rows`・`cols` の内側に限り、列全体（`B:B`）は書けない。
  シート全体の書式は、`styles` で広い範囲に書かずに、シートの `defaultStyle` に書く。
  広い範囲に書くと、値のないセルまで書式を持つセルになり、HTMLと状態テキストが大きくなるからである。
- `cells`・`merges`・`rowHeights`・`colWidths` は省いてよい（空として扱う）。
  `merges` は `"A1:E1"` のように範囲で書いてもよい。
- 書式の値を誤ると（`valign` を `"center"` と書く、色を `"#ddd"` と書くなど）、どの `styles` の項目の、どの値かと、受け付ける値を示して止まる。
- 広げた結果は `python scripts/build.py --input C:/作業/見積.json --output C:/作業/広げた.json --extract-json` で見られる。
- 状態テキスト（`--state`）と `--from-xlsx` で読み込んだブックは、セルごとの形だけを持つ（`values` などを書くと、未対応の項目として止まる）。
  `--extract-json` で取り出したJSONに `styles` を足すと、範囲の書式をまとめて変えられる（例: 見出しの塗りを変える）。

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
  defaultStyle:
    font: Yu Gothic
    size: 11
  view:
    zoom: 100
  sheets[1]:
    - id: main
      name: 見積
      rows: 30
      cols: 8
      colWidths[5]: 90,220,70,100,110
      rowHeights[4]: 32,24,24,44
      freeze:
        rows: 3
        cols: 0
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
            borderTop: 2px solid #217346
            borderRight: 1px solid #808080
            borderBottom: 1px solid #808080
            borderLeft: 1px solid #808080
            numFmt: @
        B4:
          t: s
          v: "保守作業\n（月額）"
          s:
            borderTop: 2px solid #217346
            borderRight: 1px solid #808080
            borderBottom: 1px solid #808080
            borderLeft: 1px solid #808080
            wrap: true
        E4:
          t: f
          f: =C4*D4
          s:
            borderTop: 2px solid #217346
            borderRight: 1px solid #808080
            borderBottom: 1px solid #808080
            borderLeft: 1px solid #808080
            numFmt: "¥#,##0"
EXCEL_HTML_STATE_END
```

これは冒頭の例から作ったHTMLの編集状態で、`cells` はA1・A4・B4・E4だけを示している（実際のテキストには、すべてのセルが入る）。
`view` も `zoom` だけを示している（実際には、書き出した時点の `selection`・`active`・`scrollTop`・`scrollLeft` も入る）。
`values` などの省略形は、セルごとの形に広がっている。

TOONの読み方は次のとおりである。

- `名前: 値` の1行が1つの項目で、行頭の字下げ（空白2つで1段）が入れ子を表す。
- `colWidths[5]: 90,220,…` は要素が5つの配列である。
  `merges[1]{r1,c1,r2,c2}:` は、同じ項目名を持つオブジェクトを表にしたもので、次の行から1行が1要素になる。
  `[ ]` の数と実際の要素の数が合わないテキストは、途中で切れたか書き換えられたものとして拒む。
- `- ` で始まる行は、配列の要素である。
  要素がオブジェクトなら、最初の項目を `- ` の行に書き、残りの項目をその下にそろえて書く。
- 数・`true`・`false`・`null` と区別がつかない文字列、`:` `,` `"` `[` `]` `{` `}` を含む文字列、`-` か `#` で始まる文字列、前後に空白がある文字列は、`"00123"` のように引用符で囲む。
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
