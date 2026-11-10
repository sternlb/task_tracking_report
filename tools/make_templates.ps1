# Dev only (T-6.3): foreign PowerPoint templates saved by real PowerPoint, as stand-ins
# for the customer template. Needs an installed PowerPoint; the tests do not.
#
#   powershell -ExecutionPolicy Bypass -File tools\make_templates.ps1
#
# Creates in tests\fixtures\templates:
#   office_16x9.pptx   default design of this Office (layout names in the Office language),
#                      with existing slides: picture, notes, two sections, a custom show,
#                      slide numbers and a footer
#   office_16x9.potx   the same saved as template (.potx)
#   facet_4x3.pptx     built-in design "Facet", 4:3, footer "ACME confidential", a table slide
#   renamed_cust.pptx  design "Dividend", only two layouts left, renamed to "Cover" and "Content"
#                      (no "Title Only" layout at all)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = Join-Path $root 'tests\fixtures\templates'
New-Item -ItemType Directory -Force $out | Out-Null
$themes = 'C:\Program Files\Microsoft Office\root\Document Themes 16'

# a small picture for the "existing content" slides
Add-Type -AssemblyName System.Drawing
$png = Join-Path $env:TEMP 'sw-status-report-picture.png'
$bmp = New-Object System.Drawing.Bitmap 120, 80
$g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::SteelBlue); $g.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function Layout-By-Type($pres, [int]$ppLayout) {
    foreach ($l in $pres.SlideMaster.CustomLayouts) { if ($l.Layout -eq $ppLayout) { return $l } }
    return $pres.SlideMaster.CustomLayouts.Item(1)
}
function Add-Content-Slides($pres) {
    $title = Layout-By-Type $pres 1       # ppLayoutTitle
    $text = Layout-By-Type $pres 16       # ppLayoutObject ("Title and Content")
    $s1 = $pres.Slides.AddSlide(1, $title)
    $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'Old customer deck'
    $s1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'speaker notes that must disappear'
    $s2 = $pres.Slides.AddSlide(2, $text)
    $s2.Shapes.Item(1).TextFrame.TextRange.Text = 'Picture'
    # (AddChart2 fails without a window; removing charts is covered by tests/test_pptx.py)
    [void]$s2.Shapes.AddPicture($png, 0, -1, 500, 150, 120, 80)
    $s3 = $pres.Slides.AddSlide(3, $text)
    $s3.Shapes.Item(1).TextFrame.TextRange.Text = 'Agenda'
    [void]$pres.SectionProperties.AddBeforeSlide(1, 'Intro')
    [void]$pres.SectionProperties.AddBeforeSlide(3, 'Details')
    [void]$pres.SlideShowSettings.NamedSlideShows.Add('Short', [int[]]@($s1.SlideID, $s3.SlideID))
    foreach ($s in $pres.Slides) { $s.HeadersFooters.SlideNumber.Visible = -1 }
}

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
try {
    # 1. default Office design, 16:9, with content to strip
    $p = $app.Presentations.Add(0)
    $p.SlideMaster.HeadersFooters.Footer.Visible = -1
    $p.SlideMaster.HeadersFooters.Footer.Text = 'Customer template footer'
    $p.SlideMaster.HeadersFooters.SlideNumber.Visible = -1
    Add-Content-Slides $p
    $p.SaveAs((Join-Path $out 'office_16x9.pptx'), 24)   # ppSaveAsOpenXMLPresentation
    $p.SaveAs((Join-Path $out 'office_16x9.potx'), 26)   # ppSaveAsOpenXMLTemplate
    $p.Close()

    # 2. Facet design, 4:3, footer text, one table slide
    $p = $app.Presentations.Add(0)
    $p.PageSetup.SlideSize = 1                            # ppSlideSizeOnScreen (4:3)
    $p.ApplyTheme((Join-Path $themes 'Facet.thmx'))
    $p.SlideMaster.HeadersFooters.Footer.Visible = -1
    $p.SlideMaster.HeadersFooters.Footer.Text = 'ACME confidential'
    $p.SlideMaster.HeadersFooters.SlideNumber.Visible = -1
    $s = $p.Slides.AddSlide(1, (Layout-By-Type $p 11))  # ppLayoutTitleOnly
    $s.Shapes.Item(1).TextFrame.TextRange.Text = 'A table that must disappear'
    [void]$s.Shapes.AddTable(3, 3, 40, 120, 400, 120)
    $p.SaveAs((Join-Path $out 'facet_4x3.pptx'), 24)
    $p.Close()

    # 3. Dividend design, only two layouts with custom names
    $p = $app.Presentations.Add(0)
    $p.ApplyTheme((Join-Path $themes 'Dividend-new.thmx'))
    $keepTitle = $p.SlideMaster.CustomLayouts.Item(1)   # title slide in every Office design
    $keepText = $p.SlideMaster.CustomLayouts.Item(2)    # "Title and Content"
    $keepTitle.Name = 'Cover'
    $keepText.Name = 'Content'
    for ($i = $p.SlideMaster.CustomLayouts.Count; $i -ge 1; $i--) {
        $l = $p.SlideMaster.CustomLayouts.Item($i)
        if ($l.Name -ne 'Cover' -and $l.Name -ne 'Content') { $l.Delete() }
    }
    $s = $p.Slides.AddSlide(1, $keepText)
    $s.Shapes.Item(1).TextFrame.TextRange.Text = 'Leftover slide'
    $p.SaveAs((Join-Path $out 'renamed_cust.pptx'), 24)
    $p.Close()
}
finally {
    if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Remove-Item $png -ErrorAction SilentlyContinue
}
Get-ChildItem $out | Select-Object Name, Length
