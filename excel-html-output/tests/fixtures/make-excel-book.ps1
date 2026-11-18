# Builds a workbook inside hidden Excel (no clipboard) from a JSON description, lets Excel compute it, and saves it as xlsx.
# The description is {"sheets": [{"name": ..., "cells": {"A1": value, ...}}]}, where a value is one of
#   a number, a string, true / false          the value (a string Excel would read as something else is kept as text)
#   {"f": "=..."}                              a formula, set with Formula (as a formula in an xlsx file is read)
#   {"v": value, "fmt": "yyyy/mm/dd"}          a value with a number format
# With -Results, the value Excel computed for each formula is also written there as JSON ({"sheet!A1": value}; an error
# value is its text, such as "#N/A"), for a quick look without opening the file.
param([Parameter(Mandatory)][string]$Spec, [Parameter(Mandatory)][string]$Out, [string]$Results)
$ErrorActionPreference = 'Stop'
$enc = New-Object System.Text.UTF8Encoding($false)
$Out = [IO.Path]::GetFullPath($Out)  # Excel saves only to a full path with backslashes
$book = Get-Content -Raw -Encoding utf8 $Spec | ConvertFrom-Json -AsHashtable
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt $book.sheets.Count) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $formulas = New-Object System.Collections.Generic.List[object]
  $refused = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $book.sheets.Count; $i++) {
    $s = $book.sheets[$i]; $ws = $wb.Worksheets.Item($i + 1); $ws.Name = $s.name
    $scratch = $ws.Cells.Item(1048000, 16000)
    foreach ($a in $s.cells.Keys) {
      $v = $s.cells[$a]; $x = $ws.Range($a)
      if ($v -is [System.Collections.IDictionary]) {
        if ($v.ContainsKey('fmt')) { $x.NumberFormat = $v.fmt }
        if ($v.ContainsKey('f')) {
          try { $x.Formula = $v.f; $formulas.Add(@($s.name, $a, $x)) } catch { $refused.Add("$($s.name)!$a $($v.f)") }
          continue
        }
        $v = $v.v
      }
      if ($v -is [string]) {
        # A string Excel would read as a number, date, boolean or formula is entered with a leading ', as text.
        $scratch.Value2 = $v; $asText = $scratch.HasFormula -or -not ($scratch.Value2 -is [string]) -or [string]$scratch.Value2 -ne $v; $scratch.Clear() | Out-Null
        $x.Value2 = $(if ($asText) { "'" + $v } else { $v })
      } elseif ($v -is [bool]) { $x.Value2 = [bool]$v }
      elseif ($v -is [System.Numerics.BigInteger]) { $x.Value2 = [double]::Parse($v.ToString(), [Globalization.CultureInfo]::InvariantCulture) }
      else { $x.Value2 = [double]$v }
    }
  }
  # A formula Excel does not take (a wrong number of arguments, say) stops the build, so that the file has every formula.
  if ($refused.Count) { throw ("Excel did not take these formulas:`n" + ($refused -join "`n")) }
  $xl.CalculateFull()
  if ($Results) {
    # Error values come through COM as numbers; their names are those Excel shows.
    $errors = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!' }
    $computed = [ordered]@{}
    foreach ($item in $formulas) { $x = $item[2]; $v = $x.Value2; $computed["$($item[0])!$($item[1])"] = $(if ($v -is [int]) { if ($errors.ContainsKey($v)) { $errors[$v] } else { 'error ' + $v + ' ' + $x.Text } } else { $v }) }
    [IO.File]::WriteAllText($Results, ($computed | ConvertTo-Json -Depth 3), $enc)
  }
  $wb.Worksheets.Item(1).Activate()
  $wb.SaveAs($Out, 51)
  'saved ' + $formulas.Count + ' formulas'
} finally {
  foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
