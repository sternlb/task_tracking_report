# Dev only: open a .pptx in the installed PowerPoint, list slides and shapes,
# optionally export each slide as PNG. Not part of the tool (constraint C-1).
#
#   powershell -File tools\ppt_check.ps1 out\spike.pptx -PngDir out\png
#
# PowerPoint is only closed again if it was not running before.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$PngDir,
    [int]$Width = 1600
)
$Height = [int]($Width * 9 / 16)
$ErrorActionPreference = 'Stop'
$full = (Resolve-Path $Path).Path
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($full, -1, 0, 0)
    "Opened: $full"
    "Slide size: $($pres.PageSetup.SlideWidth) x $($pres.PageSetup.SlideHeight) pt, slides: $($pres.Slides.Count)"
    foreach ($s in $pres.Slides) {
        "Slide $($s.SlideIndex): layout '$($s.CustomLayout.Name)', $($s.Shapes.Count) shapes"
        foreach ($sh in $s.Shapes) {
            $info = "  - $($sh.Name)"
            if ($sh.HasTable) { $info += " table $($sh.Table.Rows.Count)x$($sh.Table.Columns.Count)" }
            if ($sh.HasChart) { $info += " chart, $($sh.Chart.SeriesCollection().Count) series" }
            if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $info += " text '" + $sh.TextFrame.TextRange.Text.Substring(0, [Math]::Min(50, $sh.TextFrame.TextRange.Text.Length)) + "'" }
            $info
        }
        if ($PngDir) {
            $dir = (New-Item -ItemType Directory -Force $PngDir).FullName
            $s.Export((Join-Path $dir "slide$($s.SlideIndex).png"), "PNG", $Width, $Height)
        }
    }
    $pres.Close()
}
finally {
    if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
