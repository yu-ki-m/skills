# Desktop Excel side of verify_excel_paste.py and verify_xlsx.py (Windows, PowerShell 7, Excel installed).
#   native : build the sheet natively in Excel from the workbook JSON, dump it, let Excel copy it and paste
#            Excel's own HTML back (what an Excel-to-Excel HTML paste gives)
#   paste  : put the given HTML on the clipboard exactly as Chrome/Edge do and paste it at the range origin
#   xlsx   : open the .xlsx file (read-only) and dump the range of the sheet named SheetName, with the sheet's column
#            widths, row heights, view and pictures; then build the sheet natively and dump it. No clipboard.
# Dumps every cell of the range (value, type, formula, format, font, fill, alignment, borders, merge, link).
param(
  [Parameter(Mandatory)][ValidateSet('native', 'paste', 'xlsx')][string]$Mode,
  [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][int]$R1, [Parameter(Mandatory)][int]$C1,
  [Parameter(Mandatory)][int]$R2, [Parameter(Mandatory)][int]$C2,
  [string]$Workbook, [string]$Sheet, [string]$Html, [string]$Text, [string]$CfHtml, [switch]$Pdf,
  [string]$Xlsx, [string]$SheetName, [int]$Rows, [int]$Cols)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\clip.ps1"
$enc = New-Object System.Text.UTF8Encoding($false)

# A colour Excel does not give as one number (a gradient fill, a text in several colours) is $null.
function Num($v) { if ($null -eq $v -or $v -is [DBNull]) { $null } else { [int]$v } }

function Dump-Range($xl, $ws) {
  $cells = @()
  for ($r = $R1; $r -le $R2; $r++) { for ($c = $C1; $c -le $C2; $c++) {
    $x = $ws.Cells.Item($r, $c); $b = $x.Borders
    $bd = { param($i) $e = $b.Item($i); if ($e.LineStyle -eq -4142) { 'none' } else { "$($e.LineStyle)/$($e.Weight)/$(Num $e.Color)" } }
    $cells += [ordered]@{
      addr = $x.Address(0,0); text = $x.Text; value2 = $x.Value2; hasFormula = $x.HasFormula; formula = $x.Formula
      prefix = $x.PrefixCharacter; numfmt = $x.NumberFormat
      font = $x.Font.Name; size = $x.Font.Size; bold = $x.Font.Bold; italic = $x.Font.Italic
      underline = $x.Font.Underline; strike = $x.Font.Strikethrough; color = Num $x.Font.Color
      fill = if ($x.Interior.Pattern -eq -4142) { 'none' } else { Num $x.Interior.Color }
      halign = $x.HorizontalAlignment; valign = $x.VerticalAlignment; wrap = $x.WrapText; indent = $x.IndentLevel
      mergeArea = if ($x.MergeCells) { $x.MergeArea.Address(0,0) } else { '' }
      hyperlink = if ($x.Hyperlinks.Count) { $x.Hyperlinks.Item(1).Address } else { '' }
      bTop = & $bd 8; bBottom = & $bd 9; bLeft = & $bd 7; bRight = & $bd 10
    }
  } }
  $cols = @(); for ($c = $C1; $c -le $C2; $c++) { $cols += [ordered]@{ chars = $ws.Columns.Item($c).ColumnWidth; pt = $ws.Columns.Item($c).Width } }
  $rowsInfo = @(); for ($r = $R1; $r -le $R2; $r++) { $rowsInfo += $ws.Rows.Item($r).RowHeight }
  return [ordered]@{ excel = $xl.Version; cells = $cells; columns = $cols; rowHeights = $rowsInfo }
}
function Save-Dump($xl, $ws, $path) {
  [IO.File]::WriteAllText($path, ((Dump-Range $xl $ws) | ConvertTo-Json -Depth 6), $enc)
  if ($Pdf) {
    $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
    $ws.PageSetup.PrintGridlines = $true; $ws.PageSetup.PrintHeadings = $true
    $ws.PageSetup.PrintArea = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($R2, $C2)).Address()
    $ws.ExportAsFixedFormat(0, [IO.Path]::ChangeExtension($path, '.pdf'))
  }
}
function Paste-At($xl, $path) {
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Activate(); $ws.Cells.Item($R1, $C1).Select() | Out-Null
  $ws.Paste()
  Save-Dump $xl $ws $path
}

# NormalFromBook (used for xlsx files): the reference book's Normal style takes the book's default font, and a cell gets a
# font of its own only where it differs. Empty cells then have the book's default font, as in the file, and Japanese
# text in a Latin default font (Calibri) keeps one font name: set cell by cell, Excel reads such a cell's font back as no
# single name (desktop Excel 16.0).
function Build-Native($xl, [switch]$NormalFromBook) {
  $w = Get-Content -Raw -Encoding utf8 $Workbook | ConvertFrom-Json -AsHashtable
  $s = $w.sheets | Where-Object { $_.id -eq $Sheet } | Select-Object -First 1
  if (-not $s) { throw "sheet not found: $Sheet" }
  $DEFAULT = [ordered]@{font='Yu Gothic';size=11;bold=$false;italic=$false;underline=$false;strike=$false;color='#000000';fill='none';align='general';valign='middle';wrap=$false;numFmt='General';borderTop='none';borderRight='none';borderBottom='none';borderLeft='none';indent=0}
  $ALIGN = @{general=1; left=-4131; center=-4108; right=-4152}; $VALIGN = @{top=-4160; middle=-4108; bottom=-4107}
  $LINE = @{solid=1; dashed=-4115; dotted=-4118; double=-4119}; $WEIGHT = @{'0.5'=1; '1'=2; '2'=-4138; '3'=4}
  $styleOf = { param($cell) $st = [ordered]@{}; foreach ($k in $DEFAULT.Keys) { $st[$k] = $DEFAULT[$k] }
    foreach ($src in @($w.defaultStyle, $s.defaultStyle, $cell.s)) { if ($src) { foreach ($k in $src.Keys) { $st[$k] = $src[$k] } } }; $st }
  $bgr = { param($hex) $h = $hex.TrimStart('#'); [int]("0x" + $h.Substring(4,2) + $h.Substring(2,2) + $h.Substring(0,2)) }
  $coord = { param($a) $m = [regex]::Match($a, '^([A-Z]+)(\d+)$'); $c = 0; foreach ($ch in $m.Groups[1].Value.ToCharArray()) { $c = $c * 26 + [int]$ch - 64 }; @([int]$m.Groups[2].Value, $c) }

  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $bookFont = & $styleOf @{}
  if ($NormalFromBook) { foreach ($style in $wb.Styles) { if ($style.BuiltIn -and $style.Name -eq 'Normal') { $style.Font.Name = $bookFont.font; $style.Font.Size = [double]$bookFont.size } } }
  $scratch = $ws.Cells.Item(1048000, 16000)
  $probe = $ws.Columns.Item(16001); $probe.ColumnWidth = 10; $p10 = $probe.Width; $probe.ColumnWidth = 20; $p20 = $probe.Width
  $slope = ($p20 - $p10) / 10; $icpt = $p10 - 10 * $slope
  for ($c = 0; $c -lt $s.colWidths.Count; $c++) { $px = $s.colWidths[$c]; if ($null -ne $px) { $ws.Columns.Item($c + 1).ColumnWidth = [math]::Max(0, ($px * 0.75 - $icpt) / $slope) } }
  for ($r = 0; $r -lt $s.rowHeights.Count; $r++) { $px = $s.rowHeights[$r]; if ($null -ne $px) { $ws.Rows.Item($r + 1).RowHeight = $px * 0.75 } }
  $mergeOf = @{}; $hidden = @{}
  foreach ($m in $s.merges) {
    $rg = $ws.Range($ws.Cells.Item([int]$m.r1 + 1, [int]$m.c1 + 1), $ws.Cells.Item([int]$m.r2 + 1, [int]$m.c2 + 1)); $rg.Merge()
    $mergeOf[$rg.Cells.Item(1, 1).Address(0, 0)] = $rg
    # The page draws a merged range with its top-left cell's format only, so the other cells' formats are not built.
    foreach ($hc in $rg.Cells) { $ha = $hc.Address(0, 0); if ($ha -ne $rg.Cells.Item(1, 1).Address(0, 0)) { $hidden[$ha] = $true } }
  }
  foreach ($key in $s.cells.Keys) {
    if ($hidden.ContainsKey($key)) { continue }
    $cell = $s.cells[$key]; $rc = & $coord $key; $x = $ws.Cells.Item($rc[0], $rc[1]); $st = & $styleOf $cell
    $fx = $x; if ($mergeOf.ContainsKey($key)) { $fx = $mergeOf[$key] }
    if ($st.numFmt -ne 'General') { $fx.NumberFormat = $st.numFmt }
    switch ($cell.t) {
      's' {
        $v = [string]$cell.v; $needPrefix = $false
        if ($v -ne '' -and $st.numFmt -ne '@') { $scratch.Value2 = $v; $needPrefix = $scratch.HasFormula -or -not ($scratch.Value2 -is [string]) -or [string]$scratch.Value2 -ne $v; $scratch.Clear() | Out-Null }
        $x.Value2 = $(if ($needPrefix) { "'" + $v } else { $v })
      }
      # ConvertFrom-Json reads an integer past the range of Int64 (-99999999999999900000, as JavaScript writes such a number)
      # as a BigInteger, which [double] cuts toward zero; parsed from its digits it is the nearest double, as in the book.
      # The value is set as [double]$n: set from an expression without that cast ($(...) or the result of Parse), the COM
      # call failed with "Unable to cast object of type 'System.Double' to type 'System.String'" (PowerShell 7.6).
      'n' { $n = $cell.v; if ($n -is [System.Numerics.BigInteger]) { $n = [double]::Parse($n.ToString(), [Globalization.CultureInfo]::InvariantCulture) }; $x.Value2 = [double]$n }
      'b' { $x.Value2 = [bool]$cell.v }
      # .Formula takes English names: JIS, as the Japanese Excel shows DBCS (and the page takes it), is written as DBCS.
      # Excel gives a cell left General a format of its own for some formulas (=DATE(...) a date); the page's cell is
      # General, so that is set back (the Japanese Excel takes General only by its own name, G/標準).
      'f' {
        $x.Formula = [regex]::Replace([string]$cell.f, '"(?:[^"]|"")*"|(?<![A-Za-z0-9_.])JIS(?=\s*\()', { param($m) if ($m.Value[0] -eq '"') { $m.Value } else { 'DBCS' } }, 'IgnoreCase')
        if ($st.numFmt -eq 'General' -and $fx.NumberFormat -ne 'General') { try { $fx.NumberFormat = 'General' } catch { $fx.NumberFormatLocal = 'G/標準' } }
      }
    }
    $f = $fx.Font
    if (-not $NormalFromBook -or $st.font -ne $bookFont.font) { $f.Name = $st.font }
    if (-not $NormalFromBook -or $st.size -ne $bookFont.size) { $f.Size = [double]$st.size }
    $f.Bold = $st.bold; $f.Italic = $st.italic
    $f.Underline = $(if ($st.underline) { 2 } else { -4142 }); $f.Strikethrough = $st.strike; $f.Color = & $bgr $st.color
    if ($st.fill -ne 'none') { $fx.Interior.Color = & $bgr $st.fill }
    $fx.HorizontalAlignment = $ALIGN[$st.align]; $fx.VerticalAlignment = $VALIGN[$st.valign]; $fx.WrapText = $st.wrap
    if ($st.indent -gt 0) { $fx.IndentLevel = [int]$st.indent }
    foreach ($pair in @(@('borderTop', 8), @('borderBottom', 9), @('borderLeft', 7), @('borderRight', 10))) {
      $v = $st[$pair[0]]
      if ($v -ne 'none') { $m = [regex]::Match($v, '^([\d.]+)px (\w+) (#[0-9a-fA-F]{6})$'); $e = $fx.Borders.Item($pair[1])
        $e.LineStyle = $LINE[$m.Groups[2].Value]; if ($m.Groups[2].Value -ne 'double') { $e.Weight = $WEIGHT[$m.Groups[1].Value] }; $e.Color = & $bgr $m.Groups[3].Value }
    }
    if ($cell.link) { [void]$ws.Hyperlinks.Add($x, $cell.link, '', '', [string]$cell.v); $x.Font.Name = $st.font; $x.Font.Size = [double]$st.size }
  }
  return $ws
}

# The sheet's sizes, view and pictures as Excel shows them. Excel writes a repair log (error*.xml in TEMP) when it has to
# repair a file to open it; such logs that appear while opening are reported. The file is opened before any other book:
# with another book already open, hidden Excel reports frozen panes of 2 rows and 1 column as SplitRow 1 and SplitColumn 0
# (desktop Excel 16.0), although the panes themselves are right.
function Open-Xlsx($xl) {
  $before = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName)
  $wb = $xl.Workbooks.Open($Xlsx, 0, $true)
  $repairs = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.FullName } |
    ForEach-Object { Get-Content -Raw -Encoding utf8 $_.FullName })
  $active = $wb.ActiveSheet.Name
  $names = @(foreach ($s in $wb.Worksheets) { $s.Name })
  # A hidden sheet cannot be activated, and its view would be read from the sheet that is: it is shown first (in this
  # read-only copy, which is never saved).
  $ws = $wb.Worksheets.Item($SheetName); if ($ws.Visible -ne -1) { $ws.Visible = -1 }; $ws.Activate(); $win = $wb.Windows.Item(1)
  # Pictures only (msoPicture 13): a comment, a chart or a drawn shape is a shape too.
  $pictures = @(foreach ($sh in $ws.Shapes) { if ($sh.Type -eq 13) { [ordered]@{ name = $sh.Name; type = $sh.Type; topLeft = $sh.TopLeftCell.Address(0, 0)
    left = $sh.Left; top = $sh.Top; width = $sh.Width; height = $sh.Height; placement = $sh.Placement } } })
  # Reading the page setup asks the printer driver, which fails where no printer is installed; then it is left out.
  $page = try { $p = $ws.PageSetup; [ordered]@{ paperSize = $p.PaperSize; orientation = $p.Orientation; zoom = $p.Zoom; fitWide = $p.FitToPagesWide; fitTall = $p.FitToPagesTall } } catch { $null }
  $info = [ordered]@{ repairs = $repairs; activeSheet = $active; sheetNames = $names; page = $page; freeze = $win.FreezePanes; splitRow = $win.SplitRow
    splitColumn = $win.SplitColumn; gridlines = $win.DisplayGridlines; zoom = $win.Zoom
    colWidths = @(for ($c = 1; $c -le $Cols; $c++) { $ws.Columns.Item($c).Width }); rowHeights = @(for ($r = 1; $r -le $Rows; $r++) { $ws.Rows.Item($r).RowHeight })
    pictures = $pictures; activeCell = $xl.ActiveCell.Address(0, 0); selection = $xl.Selection.Address(0, 0) }
  [IO.File]::WriteAllText("$Out.sheet.json", ($info | ConvertTo-Json -Depth 6), $enc)
  Save-Dump $xl $ws "$Out.xlsx.json"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  if ($Mode -eq 'xlsx') {
    Open-Xlsx $xl
    $ws = Build-Native $xl -NormalFromBook
    Save-Dump $xl $ws "$Out.native.json"
  } elseif ($Mode -eq 'native') {
    $ws = Build-Native $xl
    Save-Dump $xl $ws "$Out.native.json"
    $ws.Range($ws.Cells.Item($R1, $C1), $ws.Cells.Item($R2, $C2)).Copy() | Out-Null
    $h = [RawClip]::GetHtml(); $t = [RawClip]::GetText()
    [IO.File]::WriteAllText("$Out.excelcopy.html", $h, $enc)
    [RawClip]::SetHtml($h.Substring($h.IndexOf('<html')), $t)
    Paste-At $xl "$Out.excelcopy.json"
  } else {
    $t = [IO.File]::ReadAllText($Text, $enc)
    if ($CfHtml) { [RawClip]::SetRaw([IO.File]::ReadAllBytes($CfHtml), $t) }  # bytes captured from a real browser copy
    else {
      $h = [IO.File]::ReadAllText($Html, $enc)
      # Chrome and Edge wrap text/html this way when they write CF_HTML.
      [RawClip]::SetHtml("<html>`r`n<body>`r`n<!--StartFragment-->" + $h + "<!--EndFragment-->`r`n</body>`r`n</html>", $t)
    }
    Paste-At $xl "$Out.json"
  }
  'OK'
} finally {
  foreach ($wbk in @($xl.Workbooks)) { $wbk.Close($false) }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
