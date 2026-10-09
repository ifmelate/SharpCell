<#
.SYNOPSIS
    Builds SharpCell's own reference workbooks: Excel calculates the cases and saves them as .xlsx.

.DESCRIPTION
    Each tools/excel-oracle/cases/*.csv becomes tests/corpus/sharpcell/<name>.xlsx. Excel is driven
    through COM: the script writes the inputs and formulas, runs a full rebuild calculation and saves
    the workbook, so the cached values in the file are Excel's. A file named <name>.1904.csv uses
    the 1904 date system. oracle-meta.json records the Excel version and settings used.

    CSV columns (header required):
      sheet    sheet name; sheets are created in order of first use
      cell     A1 address (empty for a name)
      content  a formula starting with '=', or a constant: a number (invariant, '.' decimals),
               TRUE/FALSE, or text; text that would read as something else starts with an apostrophe
      kind     empty: dynamic array formula (Range.Formula2)
               legacy: formula entered the pre-dynamic-array way (Range.Formula, implicit intersection)
               array:<A1:B2>: array formula (Ctrl+Shift+Enter) over that area, content is the formula
               name:<Name>: defined name with content as its formula; cell is ignored

    Re-running replaces every output file. Requires Windows with desktop Excel (2021 or Microsoft 365
    for dynamic arrays and LAMBDA).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\excel-oracle\Build-Oracle.ps1
#>
[CmdletBinding()]
param(
    [string]$Cases = (Join-Path $PSScriptRoot 'cases'),
    [string]$Output = (Join-Path $PSScriptRoot '..\..\tests\corpus\sharpcell')
)

$ErrorActionPreference = 'Stop'
$xlOpenXMLWorkbook = 51
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

function Set-Content-Value($range, [string]$content) {
    if ($content.StartsWith("'")) { $range.Value2 = $content; return }
    if ($content -eq 'TRUE') { $range.Value2 = $true; return }
    if ($content -eq 'FALSE') { $range.Value2 = $false; return }
    $number = 0.0
    if ([double]::TryParse($content, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$number)) {
        $range.Value2 = $number
        return
    }
    $range.Value2 = $content
}

New-Item -ItemType Directory -Force -Path $Output | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$files = @()
try {
    foreach ($csv in Get-ChildItem -Path $Cases -Filter *.csv | Sort-Object Name) {
        $rows = @(Import-Csv -Path $csv.FullName -Encoding UTF8)
        $is1904 = $csv.BaseName.EndsWith('.1904')
        $name = $csv.BaseName
        Write-Host "$name ($($rows.Count) rows)"

        $workbook = $excel.Workbooks.Add()
        $workbook.Date1904 = $is1904
        while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete() }
        $sheets = @{}

        foreach ($row in $rows) {
            $kind = if ($row.kind) { $row.kind } else { '' }
            if ($kind.StartsWith('name:')) {
                $workbook.Names.Add($kind.Substring(5), $row.content) | Out-Null
                continue
            }

            if (-not $sheets.ContainsKey($row.sheet)) {
                if ($sheets.Count -eq 0) {
                    $sheet = $workbook.Worksheets.Item(1)
                } else {
                    $sheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
                }
                $sheet.Name = $row.sheet
                $sheets[$row.sheet] = $sheet
            }

            $sheet = $sheets[$row.sheet]
            if ($kind.StartsWith('array:')) {
                $sheet.Range($kind.Substring(6)).FormulaArray = $row.content
            } elseif ($kind -eq 'legacy') {
                $sheet.Range($row.cell).Formula = $row.content
            } elseif ($row.content.StartsWith('=')) {
                $sheet.Range($row.cell).Formula2 = $row.content
            } else {
                Set-Content-Value $sheet.Range($row.cell) $row.content
            }
        }

        $excel.CalculateFullRebuild()
        $target = Join-Path (Resolve-Path $Output) "$name.xlsx"
        if (Test-Path $target) { Remove-Item $target }
        $workbook.SaveAs($target, $xlOpenXMLWorkbook)
        $workbook.Close($false)
        $files += [ordered]@{ file = "$name.xlsx"; cases = $rows.Count; dateSystem = $(if ($is1904) { '1904' } else { '1900' }) }
    }

    $meta = [ordered]@{
        excelVersion = $excel.Version
        excelBuild = $excel.Build
        operatingSystem = $excel.OperatingSystem
        decimalSeparator = $excel.International(3)
        listSeparator = $excel.International(5)
        generated = (Get-Date).ToString('yyyy-MM-dd')
        files = $files
    }
    $meta | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Output 'oracle-meta.json') -Encoding UTF8
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
