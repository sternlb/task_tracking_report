# Dev only: build the short tutorial video (about one minute, no sound, German captions)
# with the installed PowerPoint: scenes as animated slides, exported as MP4.
#
#   powershell -ExecutionPolicy Bypass -File tools\make_tutorial_video.ps1 [-PngDir out\video_png] [-NoVideo]
#
# Output: docs/tutorial/video/tutorial.mp4 and docs/tutorial/video/poster.png
# Uses the slide images in docs/tutorial/img. If ffmpeg is on the PATH, the video is
# re-encoded to a much smaller file for embedding. Saved as UTF-8 *with* BOM, because
# Windows PowerShell 5.1 would otherwise misread the German umlauts.
param(
    [string]$PngDir,
    [switch]$NoVideo
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$img = Join-Path $root 'docs\tutorial\img'
$outDir = Join-Path $root 'docs\tutorial\video'
New-Item -ItemType Directory -Force $outDir | Out-Null

# ---- colours (BGR integers for COM) and fonts ----------------------------------------
function RGB([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }
$NAVY = RGB 31 78 121; $INK = RGB 22 32 44; $PAPER = RGB 243 245 248; $MUTED = RGB 88 101 119
$WHITE = RGB 255 255 255; $LINE = RGB 214 221 230; $TERM = RGB 18 26 36; $TERMINK = RGB 215 224 234
$GREEN = RGB 46 158 68; $ORANGE = RGB 240 162 2; $RED = RGB 214 40 40; $SOFT = RGB 220 231 242
$FONT = 'Segoe UI'; $FONTB = 'Segoe UI Semibold'; $MONO = 'Consolas'

# ---- COM constants ---------------------------------------------------------------------
$ppLayoutBlank = 12; $msoShapeRect = 1; $msoShapeRounded = 5; $msoShapeOval = 9; $msoShapeRightArrow = 33
$fxAppear = 1; $fxFade = 10; $fxCustom = 0; $fxGrowShrink = 59
$afterPrev = 3; $withPrev = 2; $motion = 1
$fadeSmooth = 3849

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

# ---- helpers ---------------------------------------------------------------------------
function New-Scene([int]$seconds, [int]$bg = $PAPER) {
    $s = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
    $s.FollowMasterBackground = 0
    $s.Background.Fill.Solid()
    $s.Background.Fill.ForeColor.RGB = $bg
    $t = $s.SlideShowTransition
    $t.EntryEffect = $fadeSmooth; $t.Duration = 0.6
    $t.AdvanceOnClick = 0; $t.AdvanceOnTime = -1; $t.AdvanceTime = $seconds
    return $s
}
function Add-Text($s, [string]$text, [double]$l, [double]$t, [double]$w, [double]$h, [double]$size = 18,
                  [int]$color = $INK, [string]$font = $FONT, [int]$align = 1, [switch]$Bold) {
    $shp = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
    $tf = $shp.TextFrame
    $tf.WordWrap = -1; $tf.AutoSize = 0
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $r = $tf.TextRange
    $r.Text = $text
    $r.Font.Name = $font; $r.Font.Size = $size; $r.Font.Color.RGB = $color
    if ($Bold) { $r.Font.Bold = -1 }
    $r.ParagraphFormat.Alignment = $align
    return $shp
}
function Add-Box($s, [double]$l, [double]$t, [double]$w, [double]$h, [int]$fill, [int]$shape = $msoShapeRounded, [int]$line = -1) {
    $shp = $s.Shapes.AddShape($shape, $l, $t, $w, $h)
    $shp.Fill.Solid(); $shp.Fill.ForeColor.RGB = $fill
    if ($line -lt 0) { $shp.Line.Visible = 0 } else { $shp.Line.ForeColor.RGB = $line; $shp.Line.Weight = 1 }
    if ($shape -eq $msoShapeRounded) { $shp.Adjustments.Item(1) = 0.12 }
    $shp.Shadow.Visible = 0
    return $shp
}
function Set-BoxText($shp, [string]$text, [double]$size, [int]$color, [string]$font = $FONT, [int]$align = 2) {
    $r = $shp.TextFrame.TextRange
    $r.Text = $text; $r.Font.Name = $font; $r.Font.Size = $size; $r.Font.Color.RGB = $color
    $r.ParagraphFormat.Alignment = $align
    $shp.TextFrame.WordWrap = -1
}
function Add-Outline($s, [double]$l, [double]$t, [double]$w, [double]$h, [int]$color = $ORANGE) {
    $shp = $s.Shapes.AddShape($msoShapeRounded, $l, $t, $w, $h)
    $shp.Fill.Visible = 0; $shp.Line.ForeColor.RGB = $color; $shp.Line.Weight = 3; $shp.Shadow.Visible = 0
    return $shp
}
function Reveal($s, $shp, [double]$delay = 0.3, [double]$dur = 0.5, [int]$fx = $fxFade) {
    $e = $s.TimeLine.MainSequence.AddEffect($shp, $fx, 0, $afterPrev)
    $e.Timing.TriggerDelayTime = $delay
    if ($fx -ne $fxAppear) { $e.Timing.Duration = $dur }
    return $e
}
function Caption($s, [string]$text) {
    $bar = Add-Box $s 0 474 960 66 $NAVY $msoShapeRect
    $tb = Add-Text $s $text 48 486 864 44 19 $WHITE $FONT 1
    $tb.TextFrame.VerticalAnchor = 3
}
function Chip($s, [string]$text) { [void](Add-Text $s $text 48 32 600 22 13 $MUTED $FONTB) }
function Heading($s, [string]$text) { [void](Add-Text $s $text 48 56 864 44 30 $INK $FONTB) }

# ---- scene 1: intro ------------------------------------------------------------------------
$s = New-Scene 6 $NAVY
[void](Add-Text $s 'SW Status Report' 48 150 864 60 46 $WHITE $FONTB)
[void](Add-Text $s 'Vom Excel-Plan zum Kunden-Deck – in einer Minute erklärt' 48 214 864 34 22 $SOFT)
$labels = @('Tasks.xlsx', 'make_report.cmd', 'Status-Deck .pptx')
$x = 48
for ($i = 0; $i -lt 3; $i++) {
    $b = Add-Box $s $x 320 230 64 $WHITE
    Set-BoxText $b $labels[$i] 18 $NAVY $FONTB
    [void](Reveal $s $b 0.25 0.4)
    if ($i -lt 2) {
        $a = Add-Box $s ($x + 246) 340 48 24 $SOFT $msoShapeRightArrow
        [void](Reveal $s $a 0.1 0.3)
    }
    $x += 310
}

# ---- scene 2: setup ---------------------------------------------------------------------------
$s = New-Scene 7
Chip $s 'SCHRITT 1 · EINMALIG'
Heading $s 'ZIP entpacken – fertig'
$win = Add-Box $s 200 120 560 330 $WHITE $msoShapeRounded $LINE
$bar = Add-Box $s 200 120 560 36 $SOFT $msoShapeRect
[void](Add-Text $s 'C:\Tools\sw-status-report' 216 128 520 22 13 $INK $MONO)
$files = @(
    @('statusreport.pyz', 'das Tool – eine einzige Datei'),
    @('make_report.cmd', 'der Startknopf'),
    @('Tasks.xlsx', 'leere Eingabe-Excel'),
    @('Tutorial (Deutsch).html', 'diese Anleitung'),
    @('README.md', 'Kurzreferenz (Englisch)'))
$y = 172
foreach ($f in $files) {
    $n = Add-Text $s $f[0] 224 $y 250 24 16 $INK $MONO
    $d = Add-Text $s $f[1] 480 ($y + 2) 270 24 14 $MUTED
    [void](Reveal $s $n 0.2 0.35); [void]$s.TimeLine.MainSequence.AddEffect($d, $fxFade, 0, $withPrev)
    $y += 52
}
Caption $s 'Keine Installation. Nur Python 3.9 oder neuer muss da sein (py --version).'

# ---- scene 3: the Tasks sheet ------------------------------------------------------------------
$s = New-Scene 10
Chip $s 'SCHRITT 2 · JEDE WOCHE'
Heading $s 'Aufgaben im Blatt „Tasks“ pflegen'
$cols = @('ID', 'Title', 'Version', 'Start', 'Due', 'Estimate (h)', 'Spent (h)', 'Light override')
$widths = @(70, 230, 70, 86, 86, 92, 76, 110)
$rows = @(
    @('PRJ-201', 'Login dialog with single sign-on', '1.2.0', '2026-10-19', '2026-11-27', '80', '44', ''),
    @('PRJ-202', 'CAN bus driver for the new ECU', '1.2.0', '2026-10-01', '2026-11-13', '120', '92', ''),
    @('PRJ-206', 'Unit tests for the CRC module', '1.2.0', '2026-10-05', '2026-11-20', '40', '40', ''),
    @('PRJ-210', 'Error memory and freeze frames', '1.2.0', '2026-10-26', '2026-12-18', '60', '8', 'Orange'))
$tblShape = $s.Shapes.AddTable($rows.Count + 1, $cols.Count, 40, 130, 820, 170)
$tbl = $tblShape.Table
$tbl.ApplyStyle('{5940675A-B579-460E-94D1-54222C63F5DA}')   # "No Style, Table Grid"
for ($c = 1; $c -le $cols.Count; $c++) { $tbl.Columns.Item($c).Width = $widths[$c - 1] }
for ($r = 1; $r -le $rows.Count + 1; $r++) {
    for ($c = 1; $c -le $cols.Count; $c++) {
        $cell = $tbl.Cell($r, $c).Shape
        $txt = if ($r -eq 1) { $cols[$c - 1] } else { $rows[$r - 2][$c - 1] }
        $tr = $cell.TextFrame.TextRange
        $tr.Text = $txt; $tr.Font.Name = $FONT; $tr.Font.Size = 12
        $cell.TextFrame.MarginLeft = 5; $cell.TextFrame.MarginRight = 4
        if ($r -eq 1) { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = $NAVY; $tr.Font.Color.RGB = $WHITE; $tr.Font.Bold = -1 }
        else { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = $WHITE; $tr.Font.Color.RGB = $INK }
    }
}
$req = Add-Outline $s 34 126 372 ($tblShape.Height + 8)
$reqLabel = Add-Text $s 'Pflicht: ID, Title, Version' 40 ($tblShape.Top + $tblShape.Height + 14) 380 24 16 $ORANGE $FONTB
[void](Reveal $s $req 0.8 0.4); [void]$s.TimeLine.MainSequence.AddEffect($reqLabel, $fxFade, 0, $withPrev)
$plan = Add-Outline $s (40 + 70 + 230 + 70 + 3) 126 (86 + 86 + 92 + 76 - 6) ($tblShape.Height + 8) $NAVY
$planLabel = Add-Text $s 'Plan und Stunden ergeben die Ampel' 420 ($tblShape.Top + $tblShape.Height + 14) 340 24 16 $NAVY $FONTB
[void](Reveal $s $plan 1.2 0.4); [void]$s.TimeLine.MainSequence.AddEffect($planLabel, $fxFade, 0, $withPrev)
$ov = Add-Outline $s (40 + 820 - 107) 126 113 ($tblShape.Height + 8) $GREEN
$ovLabel = Add-Text $s 'Ampel von Hand' 740 ($tblShape.Top + $tblShape.Height + 14) 140 24 16 $GREEN $FONTB
[void](Reveal $s $ov 1.2 0.4); [void]$s.TimeLine.MainSequence.AddEffect($ovLabel, $fxFade, 0, $withPrev)
Caption $s 'Eine Zeile pro Arbeitspaket. Oder einfach den Polarion-Export ins Blatt „Import“ kopieren.'

# ---- scene 4: drag and drop ----------------------------------------------------------------------
$s = New-Scene 8
Chip $s 'SCHRITT 3'
Heading $s 'Speichern, schließen, Excel auf make_report.cmd ziehen'
$target = Add-Box $s 560 170 180 180 $WHITE $msoShapeRounded $LINE
$tIcon = Add-Box $s 610 190 80 80 $TERM
Set-BoxText $tIcon '>_' 30 $TERMINK $MONO
[void](Add-Text $s 'make_report.cmd' 560 290 180 24 15 $INK $MONO 2)
$src = Add-Box $s 220 170 180 180 $WHITE $msoShapeRounded $LINE
$sIcon = Add-Box $s 270 190 80 80 $GREEN
Set-BoxText $sIcon 'X' 36 $WHITE $FONTB
[void](Add-Text $s 'Tasks.xlsx' 220 290 180 24 15 $INK $MONO 2)
# the file that is dragged: a copy of the Excel icon moving onto the launcher
$drag = Add-Box $s 270 190 80 80 $GREEN
Set-BoxText $drag 'X' 36 $WHITE $FONTB
$drag.Shadow.Visible = -1
$e = $s.TimeLine.MainSequence.AddEffect($drag, $fxCustom, 0, $afterPrev)
$e.Timing.TriggerDelayTime = 1.0; $e.Timing.Duration = 1.6
$m = $e.Behaviors.Add($motion)
$m.MotionEffect.FromX = 0; $m.MotionEffect.FromY = 0
$m.MotionEffect.ToX = 340 / 960 * 100; $m.MotionEffect.ToY = 0
$pulse = $s.TimeLine.MainSequence.AddEffect($target, $fxGrowShrink, 0, $afterPrev)
$pulse.Timing.Duration = 0.5; $pulse.Timing.AutoReverse = -1
try { $pulse.EffectParameters.Size = 108 } catch {}
$hint = Add-Text $s 'Ein schwarzes Fenster öffnet sich …' 560 370 300 24 16 $MUTED
[void](Reveal $s $hint 0.2 0.4)
Caption $s 'Ungespeicherte Änderungen sieht das Tool nicht – also erst speichern und Excel schließen.'

# ---- scene 5: console -----------------------------------------------------------------------------
$s = New-Scene 11
$win = Add-Box $s 70 40 820 412 $TERM $msoShapeRect
$tb = Add-Box $s 70 40 820 30 (RGB 36 50 67) $msoShapeRect
[void](Add-Text $s 'make_report.cmd' 86 46 400 20 12 $TERMINK $MONO)
$lines = @(
    @('Project:      Project Falcon / ACME Automotive', $TERMINK),
    @('Report date:  2026-11-10 (CW 46/2026)', $TERMINK),
    @('Tasks:        22 in 3 version(s)', $TERMINK),
    @('History:      previous week 2026-W45; trend CW 40 .. CW 46', $TERMINK),
    @('', $TERMINK),
    @('  Version 1.2.0     53 % (46 %)  GREEN    15 task(s)', (RGB 111 208 138)),
    @('  Version 2.0.0     23 % (20 %)  ORANGE    4 task(s)', (RGB 255 194 77)),
    @('  Version 1.1.0    100 % (88 %)  GREEN     3 task(s)', (RGB 111 208 138)),
    @('', $TERMINK),
    @('Deck written:   Project Falcon Status 2026-W46.pptx (6 slides)', $TERMINK),
    @('Snapshot saved: history\2026-W46.json', $TERMINK),
    @('', $TERMINK),
    @('Done.', (RGB 111 208 138)))
$y = 84
foreach ($ln in $lines) {
    if ($ln[0] -ne '') {
        $t = Add-Text $s $ln[0] 92 $y 780 20 14 $ln[1] $MONO
        [void](Reveal $s $t 0.28 0.15 $fxAppear)
    }
    $y += 26
}
Caption $s 'Das Tool liest die Excel, rechnet die Ampeln, schreibt das Deck und merkt sich die Woche.'

# ---- scene 6: the result ------------------------------------------------------------------------
$s = New-Scene 8
Chip $s 'SCHRITT 4 · ERGEBNIS'
Heading $s 'Das fertige Deck liegt neben deiner Excel'
$p1 = $s.Shapes.AddPicture((Join-Path $img 'title.png'), 0, -1, 48, 128, 420, 236.25)
$p2 = $s.Shapes.AddPicture((Join-Path $img 'trend.png'), 0, -1, 492, 128, 420, 236.25)
foreach ($p in @($p1, $p2)) { $p.Line.Visible = -1; $p.Line.ForeColor.RGB = $LINE }
[void](Reveal $s $p1 0.3 0.5); [void](Reveal $s $p2 0.4 0.5)
[void](Add-Text $s 'Titelfolie' 48 374 420 24 15 $MUTED)
[void](Add-Text $s 'Fortschritt über die Wochen, eine Linie pro Version' 492 374 420 24 15 $MUTED)
Caption $s 'Danach kommt pro Softwareversion eine Tabelle.'

# ---- scene 7: reading the table ----------------------------------------------------------------
$s = New-Scene 12
$L = 30; $T = 34; $W = 700; $H = $W * 9 / 16
$pic = $s.Shapes.AddPicture((Join-Path $img 'table.png'), 0, -1, $L, $T, $W, $H)
$pic.Line.Visible = -1; $pic.Line.ForeColor.RGB = $LINE
# positions inside the rendered table slide (relative to the picture)
$hl1 = Add-Outline $s ($L + 0.525 * $W) ($T + 0.287 * $H) (0.12 * $W) (0.05 * $H)
$hl2 = Add-Outline $s ($L + 0.682 * $W) ($T + 0.384 * $H) (0.064 * $W) (0.05 * $H)
$b1 = Add-Box $s ($L + 0.525 * $W - 26) ($T + 0.287 * $H - 2) 22 22 $ORANGE $msoShapeOval
Set-BoxText $b1 '1' 12 $WHITE $FONTB
$b2 = Add-Box $s ($L + 0.682 * $W - 26) ($T + 0.384 * $H - 2) 22 22 $ORANGE $msoShapeOval
Set-BoxText $b2 '2' 12 $WHITE $FONTB
$n1 = Add-Text $s "1  Fortschritt jetzt, fett.`rIn Klammern: Vorwoche." 748 80 190 80 15 $INK
$n2 = Add-Text $s "2  Ampel jetzt.`rKleine Ampel in Klammern: so war es letzte Woche." 748 190 190 110 15 $INK
$n3 = Add-Text $s "Unten: Legende mit deinen Schwellwerten." 748 330 190 60 15 $MUTED
[void](Reveal $s $hl1 0.8 0.4); [void]$s.TimeLine.MainSequence.AddEffect($b1, $fxFade, 0, $withPrev); [void]$s.TimeLine.MainSequence.AddEffect($n1, $fxFade, 0, $withPrev)
[void](Reveal $s $hl2 1.6 0.4); [void]$s.TimeLine.MainSequence.AddEffect($b2, $fxFade, 0, $withPrev); [void]$s.TimeLine.MainSequence.AddEffect($n2, $fxFade, 0, $withPrev)
[void](Reveal $s $n3 1.4 0.4)
Caption $s 'Grün = im Plan · Orange = mehr als 10 pp zurück · Rot = mehr als 25 pp oder überfällig'

# ---- scene 8: every week ---------------------------------------------------------------------------
$s = New-Scene 8
Chip $s 'JEDE WOCHE'
Heading $s 'Drei Handgriffe'
$steps = @('Liste aktualisieren', 'Excel auf make_report.cmd ziehen', 'Deck präsentieren')
$x = 48
for ($i = 0; $i -lt 3; $i++) {
    $b = Add-Box $s $x 150 250 110 $WHITE $msoShapeRounded $LINE
    Set-BoxText $b ($([string]($i + 1)) + "`r" + $steps[$i]) 18 $INK $FONTB
    $b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $NAVY
    $b.TextFrame.TextRange.Paragraphs(1).Font.Size = 28
    [void](Reveal $s $b 0.3 0.4)
    if ($i -lt 2) { $a = Add-Box $s ($x + 262) 193 40 24 $SOFT $msoShapeRightArrow; [void](Reveal $s $a 0.1 0.3) }
    $x += 314
}
$warn = Add-Box $s 48 300 864 70 $WHITE $msoShapeRounded $ORANGE
Set-BoxText $warn 'Den Ordner „history“ neben der Excel nicht löschen: daraus kommen die Vorwochenwerte und der Trend.' 16 $INK $FONT 1
$warn.TextFrame.MarginLeft = 18
[void](Reveal $s $warn 0.6 0.5)
Caption $s 'Zum Prüfen ohne zu schreiben: --check. Alles Weitere steht unten im Tutorial.'

# ---- scene 9: outro ---------------------------------------------------------------------------------
$s = New-Scene 4 $NAVY
[void](Add-Text $s 'Fertig für den Kundentermin.' 48 190 864 60 40 $WHITE $FONTB 2)
[void](Add-Text $s 'Details, Ampel-Rechner und Hilfe: weiter unten in diesem Tutorial.' 48 262 864 34 20 $SOFT $FONT 2)

# ---- output ------------------------------------------------------------------------------------------
try {
    if ($PngDir) {
        $dir = (New-Item -ItemType Directory -Force $PngDir).FullName
        foreach ($sl in $pres.Slides) { $sl.Export((Join-Path $dir "scene$($sl.SlideIndex).png"), 'PNG', 1280, 720) }
        "frames: $dir"
    }
    $pres.Slides.Item(1).Export((Join-Path $outDir 'poster.png'), 'PNG', 1280, 720)
    if (-not $NoVideo) {
        $mp4 = Join-Path $outDir 'tutorial.mp4'
        if (Test-Path $mp4) { Remove-Item $mp4 }
        # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality)
        $pres.CreateVideo($mp4, -1, 5, 720, 25, 80)
        $waited = 0
        while ($pres.CreateVideoStatus -eq 1 -or $pres.CreateVideoStatus -eq 2) {   # in progress / queued
            Start-Sleep -Seconds 2; $waited += 2
            if ($waited -gt 600) { throw 'video export timed out' }
        }
        "video status: $($pres.CreateVideoStatus) (3 = done) after $waited s"
        if (Test-Path $mp4) { "video from PowerPoint: $([int]((Get-Item $mp4).Length / 1024)) KB" }
        # optional: re-encode for embedding (mostly still frames compress very well)
        $ffmpeg = (Get-Command ffmpeg -ErrorAction SilentlyContinue).Source
        if ($ffmpeg -and (Test-Path $mp4)) {
            $small = Join-Path $outDir 'tutorial.small.mp4'
            & $ffmpeg -v error -y -i $mp4 -an -c:v libx264 -preset slow -tune stillimage -crf 30 -r 20 `
                -pix_fmt yuv420p -movflags +faststart $small
            if ($LASTEXITCODE -eq 0 -and (Test-Path $small)) { Move-Item -Force $small $mp4 }
        }
        if (Test-Path $mp4) { "video: $mp4 ($([int]((Get-Item $mp4).Length / 1024)) KB)" }
    }
}
finally {
    $pres.Close()
    if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
