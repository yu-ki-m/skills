# Builds tests/fixtures/excel-text.xlsx inside hidden Excel (no clipboard): formulas that join text (&, CONCATENATE, CONCAT,
# TEXTJOIN) on one sheet, with the values Excel computed for them. tests/xlsx-import.test.mjs reads the file and checks
# that every formula stays a formula and that the page computes what Excel saved.
#   A1:D14  the cells the formulas refer to (A6 is #DIV/0!, A14 is #VALUE!; D1 is the text TRUE)
#   E1:I25  formulas whose row or column matters (a range used as one value: implicit intersection)
#   H30:H   other formulas;  M1:M  numbers, and N1:N each of them joined as text (=M1&"")
param([Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = '結合'
  $ws.Range('A1').Value2 = 'a'; $ws.Range('A3').Value2 = [double]1; $ws.Range('A4').Value2 = $true
  $ws.Range('A5').Formula = '=""'; $ws.Range('A6').Formula = '=1/0'
  $ws.Range('A7').Value2 = [double]0.125; $ws.Range('A7').NumberFormat = '0%'
  $ws.Range('A8').Value2 = [double]46023; $ws.Range('A8').NumberFormat = 'yyyy/mm/dd'
  $ws.Range('A9').Value2 = ' sp '; $ws.Range('A10').Value2 = 'x'; $ws.Range('A14').Formula = '="a"+1'
  $ws.Range('B1').Value2 = 'b'; $ws.Range('B2').Value2 = 'c'; $ws.Range('B4').Value2 = [double]2.5
  $ws.Range('C1').Value2 = '-'; $ws.Range('C2').Value2 = '+'
  $ws.Range('D1').NumberFormat = '@'; $ws.Range('D1').Value2 = 'TRUE'
  $ws.Range('D2').Value2 = 'abc'; $ws.Range('D3').Value2 = [double]0; $ws.Range('D4').Value2 = [double]2
  $placed = [ordered]@{
    F1 = '=A1:A3&"x"'; F2 = '=A1:A3&"x"'; F3 = '=A1:A3&"x"'; F20 = '=A1:A3&"x"'
    G1 = '=CONCATENATE(A1:A3)'; G2 = '=CONCATENATE(A1:A3,"!")'; G3 = '=CONCATENATE(A1:A3)'; G21 = '=CONCATENATE(A1:A3)'
    A22 = '=CONCATENATE(A1:B1)'; B23 = '=CONCATENATE(A1:B1)'; B24 = '="<"&A1:B1'; E2 = '=A1:B2&"x"'; E1 = '=CONCATENATE(A1:B2)'
    A26 = '=A1:A1&"x"'; B26 = '=CONCATENATE(B2:B2)'
    I1 = '=TEXTJOIN(",",D1:D4,"a","","b")'; I3 = '=TEXTJOIN(",",D1:D4,"a","","b")'; I25 = '=TEXTJOIN(",",D1:D4,"a","","b")'
  }
  foreach ($a in $placed.Keys) { $ws.Range($a).Formula = $placed[$a] }
  $formulas = @(
    '="a"&"b"', '=A1&B1', '=A2&"x"', '=A3&"x"', '=A4&"x"', '=TRUE&""', '=FALSE&1', '=(1=1)&""', '=1&2', '="1"&2+1', '="a"&-1', '=50%&""', '=0.1+0.2&""',
    '=A5&"x"', '=A6&"x"', '=A7&""', '=A8&""', '=A9&"|"', '=A6&A14', '=A14&A6', '=A14+A6', '=A6*A14',
    '=CONCATENATE("a","b",1,TRUE)', '=CONCATENATE(A1)', '=CONCATENATE(A2)', '=CONCATENATE(1/3)', '=CONCATENATE(A7,A8)', '=CONCATENATE(A6,"x")',
    '=CONCATENATE("x",A14,A6)', '=CONCATENATE("a",,"b")', '=CONCATENATE(,)', '=CONCATENATE("a",)',
    '=CONCAT("a","b")', '=CONCAT(A1:B2)', '=CONCAT(A1:A5)', '=CONCAT(A7,A8)', '=CONCAT(1/3)', '=CONCAT(A1:B4)', '=CONCAT(B1:B4,A1:A4)', '=CONCAT(A2)',
    '=CONCAT(A5)', '=CONCAT(A2:B3)', '=CONCAT(A1:A10)', '=CONCAT(A14,A6)', '=CONCAT(A6:A6,A14)', '=CONCAT("a",,"b")', '=CONCAT("a",)',
    '=IFERROR(CONCAT(A6),"err")', '=IF(TRUE,CONCAT("a",1),0)', '=CONCAT("a")&TEXTJOIN("-",TRUE,"b","c")', '=SUM(1,2)&"件"', '=CONCAT(SUM(A3:A4))',
    '=TEXTJOIN(",",TRUE,A1:A5)', '=TEXTJOIN(",",FALSE,A1:A5)', '=TEXTJOIN(C1:C2,TRUE,"a","b","c","d")', '=TEXTJOIN(C1:C3,FALSE,"a","b","c","d","e")',
    '=TEXTJOIN(B1:C2,FALSE,"1","2","3","4","5","6")', '=TEXTJOIN(C1:C2,TRUE,"a","","b","c")', '=TEXTJOIN(C1:C2,TRUE,"","a","b","c")',
    '=TEXTJOIN(C1:C2,FALSE,"a","","b","c")', '=TEXTJOIN(1,TRUE,"a","b")', '=TEXTJOIN(TRUE,TRUE,"a","b")', '=TEXTJOIN(0.5,TRUE,"a","b")', '=TEXTJOIN("",TRUE,"a","b")',
    '=TEXTJOIN(A2,TRUE,"a","b")', '=TEXTJOIN(A2:A3,FALSE,"a","b","c")', '=TEXTJOIN(A1,TRUE,"x","y")', '=TEXTJOIN(",",TRUE,"")', '=TEXTJOIN(",",TRUE,A2:A2)',
    '=TEXTJOIN(",",TRUE,A1:B2)', '=TEXTJOIN(",",FALSE,A1:B2)', '=TEXTJOIN(", ",TRUE,A9,A10)', '=TEXTJOIN(",",TRUE,A1:A2,"z",A3)', '=TEXTJOIN(",",TRUE,1/3,2/3)',
    '=TEXTJOIN(",",FALSE,A2,A3)', '=TEXTJOIN(",",TRUE,A4,FALSE)', '=TEXTJOIN(",",TRUE,A1:B4)', '=TEXTJOIN(",",FALSE,A1:B4)', '=TEXTJOIN(",",TRUE,"a")',
    '=TEXTJOIN(",",1,"a","","b")', '=TEXTJOIN(",",0,"a","","b")', '=TEXTJOIN(",",2.5,"a","","b")', '=TEXTJOIN(",","TRUE","a","","b")', '=TEXTJOIN(",","true","a","","b")',
    '=TEXTJOIN(",","FALSE","a","","b")', '=TEXTJOIN(",","x","a","b")', '=TEXTJOIN(",","","a","b")', '=TEXTJOIN(",","1","a","","b")', '=TEXTJOIN(",",A2,"a","","b")',
    '=TEXTJOIN(",",A1,"a","","b")', '=TEXTJOIN(",",A3,"a","","b")', '=TEXTJOIN(",",A4,"a","","b")', '=TEXTJOIN(",",D1,"a","","b")', '=TEXTJOIN(",",D2,"a","","b")',
    '=TEXTJOIN(",",A6,"a")', '=TEXTJOIN(A6,TRUE,"a")', '=TEXTJOIN(A5:A6,TRUE,"a","b","c")', '=TEXTJOIN(",",TRUE,A6:A6)', '=TEXTJOIN(",",TRUE,A6,A14)',
    '=TEXTJOIN(A14,TRUE,A6)', '=TEXTJOIN(",",A14,A6)', '=TEXTJOIN(A6,A14,"a")', '=TEXTJOIN(",","x",A6)', '=TEXTJOIN(A6,"x","a")',
    '=CONCAT(A6,"a"+1)', '=CONCAT("a"+1,A6)', '=CONCATENATE(A6,"a"+1)', '=A6&("a"+1)', '=TEXTJOIN(A6,TRUE,"a"+1)', '=TEXTJOIN(",",TRUE,A6,"a"+1)',
    '=TEXTJOIN(A6,TRUE,A14)', '=TEXTJOIN(A14,A6,"a")', '=TEXTJOIN(A6,"a"+1,"b")',
    '=TEXTJOIN(,TRUE,"a","b")', '=TEXTJOIN(",",,"a","","b")', '=TEXTJOIN(",",TRUE,"a",,"b")', '=TEXTJOIN(",",FALSE,"a",,"b")', '=TEXTJOIN(",",FALSE,"a",)',
    '=TEXTJOIN(",",TRUE,)', '=TEXTJOIN(,,A1:A3)'
  )
  for ($i = 0; $i -lt $formulas.Count; $i++) { $ws.Cells.Item(30 + $i, 8).Formula = $formulas[$i] }
  $numbers = @((1 / 3), (2 / 3), 0.30000000000000004, 46023, 0.125, -1.5, 0, 1234567890123456, 123456789012345.6, 9.999999999999989e19, -9.99999999999999e19,
    1e20, 1.5e20, -1e20, 1.23456789012345e21, 0.000123456789012345, -0.000123456789012345, 0.0000123456789012345, 0.000012345678901234, 1e-5, 1e-12, 1.5e-12,
    1.23456789012345e-12, 1e-18, -1e-18, 1.5e-18, 1e-19, 1e-307, 2.2250738585072014e-308, 1.7976931348623157e308, 100000000000000.5, 100000000000001.5,
    -100000000000001.5, 1000000000000.375, 12345678901234.75, 123456789012.1875, 1.430511474609375e-6, 1234567890.000005, 12345678901.00005, 1.234567890123456,
    7.777777777777778, 0.9999999999999999, 9.999999999999998, (1 / 7))
  for ($i = 0; $i -lt $numbers.Count; $i++) { $ws.Cells.Item($i + 1, 13).Value2 = [double]$numbers[$i]; $ws.Cells.Item($i + 1, 14).Formula = '=M' + ($i + 1) + '&""' }
  $wb.SaveAs($Out, 51)
  'saved'
} finally {
  foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
