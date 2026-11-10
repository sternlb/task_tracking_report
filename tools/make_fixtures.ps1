# Dev only (T-2.6): create test fixtures that were really saved by Excel.
# Needs an installed Excel; the tests themselves do not.
#
#   powershell -ExecutionPolicy Bypass -File tools\make_fixtures.ps1
#
# Creates in tests\fixtures:
#   excel_filled.xlsx       --init template filled and saved by Excel (shared strings, Excel styles,
#                           dates as date cells and as text, percent cell, error rows)
#   excel_with_import.xlsx  the same plus a Polarion-like export pasted into the Import sheet
#   polarion_export.xlsx    synthetic Polarion-style export (title rows above the header, durations)
#   polarion_export.xls     the same as old binary .xls (must be rejected with a hint)
#   excel_1904.xlsx         workbook in the 1904 date system
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$fx = Join-Path $root 'tests\fixtures'
New-Item -ItemType Directory -Force $fx | Out-Null
$template = Join-Path $fx '_template.xlsx'
Push-Location $root
python -m statusreport --init $template --force | Out-Null
Pop-Location

function Set-Date($cell, [int]$y, [int]$m, [int]$d, [bool]$date1904 = $false) {
    # Excel date serial; the 1904 system counts 1462 days less
    [double]$serial = ([datetime]::new($y, $m, $d)).ToOADate()
    if ($date1904) { $serial -= 1462 }
    $cell.Value2 = $serial
}
function Set-Text($cell, [string]$text) {
    $cell.NumberFormat = '@'
    $cell.Value2 = $text
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    # ---- Polarion-like export ------------------------------------------------------------
    $pe = $xl.Workbooks.Add()
    $ws = $pe.Worksheets.Item(1)
    $ws.Name = 'Work Items'
    $ws.Cells.Item(1, 1).Value2 = 'Polarion export - Project Falcon (synthetic test data)'
    $headers = 'ID', 'Title', 'Type', 'Status', 'Assignee', 'Planned In', 'Due Date', 'Initial Estimate', 'Time Spent', 'Remaining Estimate'
    for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(3, $c + 1).Value2 = $headers[$c] }
    $data = @(
        @('PRJ-101', 'Login dialog with SSO', 'Task', 'In Progress', 'Anna', '1.2.0', @(2026, 11, 27), '3d', '1d 2h', '1d 6h'),
        @('PRJ-201', 'Flash driver', 'Task', 'Done', 'Gus', '1.2.0', @(2026, 10, 30), '1w 2d', '1w 3d', '0h'),
        @('PRJ-202', 'Watchdog handling', 'Task', 'Open', 'Hana', '1.10.0', @(2026, 12, 11), '4h 30m', '', '4h 30m'),
        @('PRJ-203', 'Startup self test', 'Task', 'In Progress', 'Ivo', '1.10.0', @(2026, 12, 4), '2d', '6h', '10h')
    )
    $r = 4
    foreach ($row in $data) {
        for ($c = 0; $c -lt $row.Count; $c++) {
            $cell = $ws.Cells.Item($r, $c + 1)
            $v = $row[$c]
            if ($v -is [array]) { $cell.NumberFormat = 'dd.mm.yyyy'; Set-Date $cell $v[0] $v[1] $v[2] }
            elseif ($v -ne '') { Set-Text $cell $v }
        }
        $r++
    }
    $pe.SaveAs((Join-Path $fx 'polarion_export.xlsx'), 51)
    $pe.SaveAs((Join-Path $fx 'polarion_export.xls'), 56)
    $exportRange = $ws.Range('A3:J7')

    # ---- filled template -----------------------------------------------------------------
    $wb = $xl.Workbooks.Open($template)
    $t = $wb.Worksheets.Item('Tasks')
    # row 2 is the #EXAMPLE row from the template and stays
    $t.Cells.Item(3, 1).Value2 = 'PRJ-101'; $t.Cells.Item(3, 2).Value2 = 'Login dialog with SSO'
    $t.Cells.Item(3, 3).Value2 = '1.2.0'; $t.Cells.Item(3, 4).Value2 = 'Anna'
    Set-Date $t.Cells.Item(3, 5) 2026 11 2; Set-Date $t.Cells.Item(3, 6) 2026 11 27
    $t.Cells.Item(3, 7).Value2 = '3d'; $t.Cells.Item(3, 8).Value2 = 10; $t.Cells.Item(3, 10).Value2 = 'in progress'
    $t.Cells.Item(3, 12).Value2 = 'waiting for API'

    $t.Cells.Item(4, 1).Value2 = 'PRJ-102'; $t.Cells.Item(4, 2).Value2 = 'CAN bus driver'
    $t.Cells.Item(4, 3).Value2 = '1.2.0'
    Set-Text $t.Cells.Item(4, 5) '01.10.2026'; Set-Text $t.Cells.Item(4, 6) '13/11/2026'
    $t.Cells.Item(4, 7).Value2 = 40; $t.Cells.Item(4, 8).Value2 = 32

    $t.Cells.Item(5, 1).Value2 = 'PRJ-103'; $t.Cells.Item(5, 2).Value2 = 'Diagnostic trouble codes'
    $t.Cells.Item(5, 3).Value2 = '1.10.0'
    $t.Cells.Item(5, 9).NumberFormat = '0%'; $t.Cells.Item(5, 9).Value2 = 0.9

    $t.Cells.Item(6, 1).Value2 = 'PRJ-104'; $t.Cells.Item(6, 2).Value2 = 'Bootloader update'
    $t.Cells.Item(6, 3).Value2 = '1.2.0'; $t.Cells.Item(6, 11).Value2 = 'Red'; $t.Cells.Item(6, 12).Value2 = 'escalated'

    $t.Cells.Item(7, 1).Value2 = 'PRJ-105'; $t.Cells.Item(7, 2).Value2 = 'Bad date row'
    $t.Cells.Item(7, 3).Value2 = '1.2.0'; Set-Text $t.Cells.Item(7, 6) '31.02.2026'

    $t.Cells.Item(8, 1).Value2 = 'PRJ-101'; $t.Cells.Item(8, 2).Value2 = 'Duplicate of row 3'; $t.Cells.Item(8, 3).Value2 = '1.2.0'
    $t.Cells.Item(9, 2).Value2 = 'Row without ID'; $t.Cells.Item(9, 3).Value2 = '1.2.0'

    $t.Cells.Item(10, 1).Value2 = 'PRJ-106'; $t.Cells.Item(10, 2).Value2 = 'Unit tests CRC'
    $t.Cells.Item(10, 3).Value2 = '1.2.0'; $t.Cells.Item(10, 10).Value2 = 'Done'

    $s = $wb.Worksheets.Item('Settings')
    for ($row = 2; $row -le 20; $row++) {
        switch ($s.Cells.Item($row, 1).Value2) {
            'Project' { $s.Cells.Item($row, 2).Value2 = 'Falcon' }
            'Customer' { $s.Cells.Item($row, 2).Value2 = 'ACME Automotive' }
            'Report date' { Set-Date $s.Cells.Item($row, 2) 2026 11 10 }
            'Version order' { $s.Cells.Item($row, 2).Value2 = '1.2.0, 1.10.0' }
        }
    }
    $wb.SaveAs((Join-Path $fx 'excel_filled.xlsx'), 51)

    # ---- with the export pasted into Import --------------------------------------------------
    $exportRange.Copy($wb.Worksheets.Item('Import').Range('A1')) | Out-Null
    # keep only overlay rows in Tasks: PRJ-101 comment and a Tasks-only row
    $t.Range('A4:L10').ClearContents() | Out-Null
    $t.Cells.Item(3, 2).Value2 = ''         # title comes from the import
    $t.Cells.Item(3, 8).Value2 = ''         # spent comes from the import
    $t.Cells.Item(4, 1).Value2 = 'PRJ-301'; $t.Cells.Item(4, 2).Value2 = 'Only in Tasks'; $t.Cells.Item(4, 3).Value2 = '1.10.0'
    $wb.SaveAs((Join-Path $fx 'excel_with_import.xlsx'), 51)
    $wb.Close($false)
    $pe.Close($false)

    # ---- 1904 date system --------------------------------------------------------------------
    $d = $xl.Workbooks.Add()
    $d.Date1904 = $true
    $ws = $d.Worksheets.Item(1)
    $ws.Name = 'Dates'
    $ws.Cells.Item(1, 1).Value2 = 'Date'
    $ws.Cells.Item(2, 1).NumberFormat = 'yyyy-mm-dd'
    Set-Date $ws.Cells.Item(2, 1) 2026 11 10 $true
    $d.SaveAs((Join-Path $fx 'excel_1904.xlsx'), 51)
    $d.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $template -ErrorAction SilentlyContinue
}
Get-ChildItem $fx | Select-Object Name, Length
