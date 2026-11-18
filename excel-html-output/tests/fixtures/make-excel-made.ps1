# Builds tests/fixtures/excel-made.xlsx inside hidden Excel (no clipboard): a quote sheet with what the importer keeps
# and what it has to replace or leave out, and a price-list sheet that a VLOOKUP refers to.
param([Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Png)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = '見積'; $pl = $wb.Worksheets.Item(2); $pl.Name = '単価表'
  $pl.Range('A1').Value2 = '品番'; $pl.Range('B1').Value2 = '単価'
  $pl.Range('A2').NumberFormat = '@'; $pl.Range('A2').Value2 = '00123'; $pl.Range("B2").Value2 = [double]1200
  $pl.Range('A3').NumberFormat = '@'; $pl.Range('A3').Value2 = '00124'; $pl.Range("B3").Value2 = [double]800
  $title = $ws.Range('A1:F1'); $title.Merge(); $title.Value2 = '御見積書'
  $title.Font.Bold = $true; $title.Font.Size = 16; $title.HorizontalAlignment = -4108
  $title.Interior.ThemeColor = 5; $title.Interior.TintAndShade = 0.7999816888943144
  $title.BorderAround(1, -4138, 0, 0x466217) | Out-Null
  $heads = @('品番', '品名', '数量', '単価', '金額', '単価表の単価', '日付')
  for ($i = 0; $i -lt $heads.Count; $i++) { $ws.Cells.Item(3, $i + 1).Value2 = $heads[$i] }
  $h = $ws.Range('A3:G3'); $h.Font.Bold = $true; $h.Font.Color = 0xFFFFFF; $h.Interior.Color = 0x467321; $h.HorizontalAlignment = -4108
  $h.Borders.Item(9).LineStyle = 1; $h.Borders.Item(9).Weight = 2
  $rows = @(@('00123', 'Webサイト保守', 12, 30000), @("2行の`n品名", '初期設定', 1, 120000), @('00124', '交通費', 1, 8000))
  for ($r = 0; $r -lt 3; $r++) {
    $a = $ws.Cells.Item($r + 4, 1); $a.NumberFormat = '@'; $a.Value2 = [string]$rows[$r][0]
    $ws.Cells.Item($r + 4, 2).Value2 = [string]$rows[$r][1]; $ws.Cells.Item($r + 4, 3).Value2 = [double]$rows[$r][2]; $ws.Cells.Item($r + 4, 4).Value2 = [double]$rows[$r][3]
  }
  $ws.Range('A5').WrapText = $true
  $ws.Range('C4:C6').NumberFormat = '#,##0'; $ws.Range('D4:D5').NumberFormat = '"¥"#,##0'; $ws.Range('D6').NumberFormat = '"¥"#,##0;[赤]"¥"-#,##0'
  $ws.Range('E4:E6').Formula = '=C4*D4'; $ws.Range('E4:E9').NumberFormat = '"¥"#,##0'
  $ws.Range('D7').Value2 = '小計'; $ws.Range('E7').Formula = '=SUM(E4:E6)'
  $ws.Range('D8').Value2 = '消費税'; $ws.Range('E8').Formula = '=ROUNDDOWN(E7*0.1,0)'
  $ws.Range('D9').Value2 = '合計'; $ws.Range('E9').Formula = '=E7+E8'; $ws.Range('D9:E9').Font.Bold = $true
  $ws.Range('E9').Borders.Item(8).LineStyle = 1
  $ws.Range('F4').Formula = '=VLOOKUP(A4,単価表!A:B,2,FALSE)'; $ws.Range('F4').NumberFormat = '"¥"#,##0'
  $ws.Range('F5').Formula = '=単価表!B3'
  $ws.Range('G4').Value2 = [double]46346; $ws.Range('G4').NumberFormat = 'yyyy/m/d'
  $ws.Range('G5').Value2 = [double]46346; $ws.Range('G5').NumberFormat = 'yyyy/mm/dd'
  $ws.Range('G6').Value2 = [double]0.5208333333333334; $ws.Range('G6').NumberFormat = 'h:mm'
  $rich = $ws.Range('B10'); $rich.Value2 = '太字と普通'; $rich.Characters(1, 2).Font.Bold = $true
  [void]$ws.Hyperlinks.Add($ws.Range('A12'), 'https://example.com/', '', '', 'サイト')
  $dd = $ws.Range('C12'); $dd.Value2 = '一点鎖線'; $dd.Borders.Item(9).LineStyle = 4; $dd.Borders.Item(9).Weight = 2
  $pat = $ws.Range('D12'); $pat.Value2 = '網かけ'; $pat.Interior.Pattern = -4124; $pat.Interior.PatternColor = 0x0000FF
  $rot = $ws.Range('E12'); $rot.Value2 = '回転'; $rot.Orientation = 45
  $ws.Range('H2').Value2 = '隠した列'; $ws.Columns.Item(8).Hidden = $true
  [void]$ws.Range('B4').AddComment('コメント')
  [void]$ws.Range('C4:C6').Validation.Add(1, 1, 1, '1', '100')
  $fc = $ws.Range('E4:E6').FormatConditions.Add(1, 5, '=100000'); $fc.Interior.Color = 0x9999FF
  $widths = @(8.5, 18, 7, 10, 12, 12, 11); for ($c = 0; $c -lt $widths.Count; $c++) { $ws.Columns.Item($c + 1).ColumnWidth = [double]$widths[$c] }
  $ws.Rows.Item(1).RowHeight = [double]30; $ws.Rows.Item(5).RowHeight = [double]36
  [void]$ws.Shapes.AddPicture($Png, 0, -1, $ws.Range('B14').Left + 6, $ws.Range('B14').Top + 3, 90, 45)
  $ws.Activate(); $win = $wb.Windows.Item(1); $win.SplitRow = 3; $win.SplitColumn = 0; $win.FreezePanes = $true; $win.Zoom = 120
  $ws.Range('B5').Select() | Out-Null
  $pl.Activate(); $wb.Windows.Item(1).DisplayGridlines = $false; $ws.Activate(); $ws.Range('B5').Select() | Out-Null
  $wb.SaveAs($Out, 51)
  'saved'
} finally {
  foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
