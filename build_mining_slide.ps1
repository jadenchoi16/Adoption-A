# =============================================================================
# build_mining_slide.ps1
# Builds a single-slide PPTX titled "High mining does NOT mean high adoption."
# Output: mining_vs_adoption.pptx in the repo root.
# Run:   powershell -ExecutionPolicy Bypass -File .\build_mining_slide.ps1
# =============================================================================

$ErrorActionPreference = "Stop"

function RGB([string]$hex) {
    $h = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($h.Substring(0,2), 16)
    $g = [Convert]::ToInt32($h.Substring(2,2), 16)
    $b = [Convert]::ToInt32($h.Substring(4,2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

# Palette
$BG_LIGHT     = RGB '#fafaf7'
$BG_CARD      = RGB '#ffffff'
$INK          = RGB '#1a2a3a'
$INK_BODY     = RGB '#2c3e50'
$INK_MUTED    = RGB '#5a6b7c'
$INK_DIM      = RGB '#7a8a9a'
$BLUE_DEEP    = RGB '#0a4c6a'
$TEAL         = RGB '#5ba8b5'
$WARM         = RGB '#9b5a30'
$CORAL        = RGB '#c06030'
$GREEN        = RGB '#2e7d32'
$BORDER_LIGHT = RGB '#ebebeb'

$FONT_HEAD = "Cambria"
$FONT_BODY = "Calibri"
$FONT_MONO = "Consolas"

$SLIDE_W = 960
$SLIDE_H = 540

Write-Host "Launching PowerPoint..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1

$pres = $ppt.Presentations.Add($true)
$pres.PageSetup.SlideSize = 15  # ppSlideSize16x9

while ($pres.Slides.Count -gt 0) { $pres.Slides.Item(1).Delete() }

# helpers ---------------------------------------------------------------------

function Add-Text {
    param(
        [Parameter(Mandatory=$true)] $slide,
        [Parameter(Mandatory=$true)] [string]$text,
        [double]$left, [double]$top, [double]$width, [double]$height,
        [string]$font = $FONT_BODY,
        [double]$size = 16,
        [int]$color = $INK_BODY,
        [bool]$bold = $false,
        [bool]$italic = $false,
        [int]$align = 1,
        [bool]$upper = $false
    )
    $tb = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $tb.TextFrame.MarginLeft = 0
    $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop = 0
    $tb.TextFrame.MarginBottom = 0
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0
    $r = $tb.TextFrame.TextRange
    if ($upper) { $text = $text.ToUpper() }
    $r.Text = $text
    $r.Font.Name = $font
    $r.Font.Size = $size
    $r.Font.Color.RGB = $color
    $r.Font.Bold = if ($bold) { -1 } else { 0 }
    $r.Font.Italic = if ($italic) { -1 } else { 0 }
    $r.ParagraphFormat.Alignment = $align
    return $tb
}

function Add-Rect {
    param(
        [Parameter(Mandatory=$true)] $slide,
        [double]$left, [double]$top, [double]$width, [double]$height,
        [int]$fill = $BG_CARD,
        [int]$borderColor = -1,
        [double]$borderWidth = 0,
        [double]$cornerRadius = 0.04
    )
    if ($cornerRadius -gt 0) {
        $shape = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
        try { $shape.Adjustments.Item(1) = $cornerRadius } catch {}
    } else {
        $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    }
    $shape.Fill.ForeColor.RGB = $fill
    if ($borderColor -lt 0) {
        $shape.Line.Visible = 0
    } else {
        $shape.Line.ForeColor.RGB = $borderColor
        $shape.Line.Weight = $borderWidth
        $shape.Line.Visible = -1
    }
    return $shape
}

# Build the one slide ---------------------------------------------------------

$slide = $pres.Slides.Add(1, 12)  # ppLayoutBlank
$slide.FollowMasterBackground = $false
$slide.Background.Fill.Visible = -1
$slide.Background.Fill.ForeColor.RGB = $BG_LIGHT

# Eyebrow
[void](Add-Text -slide $slide -text "ADDITIONAL FINDING  -  MINING vs. ADOPTION" `
                -left 60 -top 60 -width 700 -height 18 `
                -font $FONT_MONO -size 11 -color $BLUE_DEEP -bold $true -upper $true)

# Title
[void](Add-Text -slide $slide -text "High mining does NOT mean high adoption." `
                -left 60 -top 92 -width 840 -height 60 `
                -font $FONT_HEAD -size 32 -color $INK -bold $true)

# Lede
[void](Add-Text -slide $slide -text "Mining and consumer adoption move on completely different inputs. Five of our top six adopters mine essentially zero Bitcoin." `
                -left 60 -top 154 -width 840 -height 44 `
                -font $FONT_BODY -size 14 -color $INK_BODY -italic $true)

# Two-column cards
$colTop = 215
$colHeight = 230
$leftX = 60
$rightX = 500
$colW = 400

# LEFT CARD: Top miners and their adoption rank ------------------------------
[void](Add-Rect -slide $slide -left $leftX -top $colTop -width $colW -height $colHeight -fill $BG_CARD -borderColor $BORDER_LIGHT -borderWidth 0.5)
[void](Add-Rect -slide $slide -left $leftX -top $colTop -width $colW -height 4 -fill $WARM -cornerRadius 0)

[void](Add-Text -slide $slide -text "TOP MINERS  -  WHERE DO THEY RANK ON ADOPTION?" `
                -left ($leftX + 18) -top ($colTop + 14) -width ($colW - 36) -height 14 `
                -font $FONT_MONO -size 9 -color $WARM -bold $true -upper $true)

# Header row
[void](Add-Text -slide $slide -text "Country" -left ($leftX + 18) -top ($colTop + 38) -width 140 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true)
[void](Add-Text -slide $slide -text "Hashrate" -left ($leftX + 158) -top ($colTop + 38) -width 100 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -align 2)
[void](Add-Text -slide $slide -text "Adoption rank" -left ($leftX + 258) -top ($colTop + 38) -width 124 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -align 3)

$leftRows = @(
    @("United States", "37.5%",  "#4",          $TEAL),
    @("Russia",        "16.4%",  "#12",         $INK_BODY),
    @("Ethiopia",      "2.6%",   "#20 (LAST)",  $CORAL),
    @("Indonesia",     "2.1%",   "#10",         $INK_BODY)
)
$rowY = $colTop + 60
foreach ($r in $leftRows) {
    [void](Add-Text -slide $slide -text $r[0] -left ($leftX + 18) -top $rowY -width 140 -height 22 -font $FONT_HEAD -size 14 -color $INK -bold $true)
    [void](Add-Text -slide $slide -text $r[1] -left ($leftX + 158) -top ($rowY + 2) -width 100 -height 20 -font $FONT_MONO -size 14 -color $INK_BODY -align 2)
    [void](Add-Text -slide $slide -text $r[2] -left ($leftX + 258) -top ($rowY + 2) -width 124 -height 20 -font $FONT_MONO -size 14 -color $r[3] -bold $true -align 3)
    $rowY += 38
}

[void](Add-Text -slide $slide -text "Mixed. The US is the only top miner that's also a top adopter. Ethiopia mines a lot but is the lowest-adoption country in the index." `
                -left ($leftX + 18) -top ($colTop + $colHeight - 56) -width ($colW - 36) -height 48 `
                -font $FONT_BODY -size 10 -color $INK_MUTED -italic $true)

# RIGHT CARD: Top adopters and their mining score ----------------------------
[void](Add-Rect -slide $slide -left $rightX -top $colTop -width $colW -height $colHeight -fill $BG_CARD -borderColor $BORDER_LIGHT -borderWidth 0.5)
[void](Add-Rect -slide $slide -left $rightX -top $colTop -width $colW -height 4 -fill $TEAL -cornerRadius 0)

[void](Add-Text -slide $slide -text "TOP ADOPTERS  -  HOW MUCH DO THEY MINE?" `
                -left ($rightX + 18) -top ($colTop + 14) -width ($colW - 36) -height 14 `
                -font $FONT_MONO -size 9 -color $TEAL -bold $true -upper $true)

[void](Add-Text -slide $slide -text "Country" -left ($rightX + 18) -top ($colTop + 38) -width 140 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true)
[void](Add-Text -slide $slide -text "Adoption rank" -left ($rightX + 158) -top ($colTop + 38) -width 100 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -align 2)
[void](Add-Text -slide $slide -text "Mining score" -left ($rightX + 258) -top ($colTop + 38) -width 124 -height 14 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -align 3)

$rightRows = @(
    @("Vietnam",      "#1", "0.010",  $CORAL),
    @("Ukraine",      "#2", "0.013",  $CORAL),
    @("Philippines",  "#3", "0.010",  $CORAL),
    @("United States","#4", "0.983",  $TEAL),
    @("Pakistan",     "#5", "0.010",  $CORAL)
)
$rowY = $colTop + 60
foreach ($r in $rightRows) {
    [void](Add-Text -slide $slide -text $r[0] -left ($rightX + 18) -top $rowY -width 140 -height 18 -font $FONT_HEAD -size 12.5 -color $INK -bold $true)
    [void](Add-Text -slide $slide -text $r[1] -left ($rightX + 158) -top ($rowY + 1) -width 100 -height 18 -font $FONT_MONO -size 13 -color $INK_BODY -align 2)
    [void](Add-Text -slide $slide -text $r[2] -left ($rightX + 258) -top ($rowY + 1) -width 124 -height 18 -font $FONT_MONO -size 13 -color $r[3] -bold $true -align 3)
    $rowY += 30
}

[void](Add-Text -slide $slide -text "Five of six adopters mine ~0%. The pattern is one-way: the US is an outlier, not a rule." `
                -left ($rightX + 18) -top ($colTop + $colHeight - 36) -width ($colW - 36) -height 30 `
                -font $FONT_BODY -size 10 -color $INK_MUTED -italic $true)

# Bottom takeaway band -------------------------------------------------------
$tkTop = 460
[void](Add-Rect -slide $slide -left 60 -top $tkTop -width 840 -height 60 -fill (RGB '#0a1620') -cornerRadius 0.04)
[void](Add-Text -slide $slide -text "THE TAKEAWAY" `
                -left 76 -top ($tkTop + 10) -width 200 -height 14 `
                -font $FONT_MONO -size 9 -color $TEAL -bold $true -upper $true)
[void](Add-Text -slide $slide -text "Mining and adoption are different businesses with different drivers. Mining follows cheap energy; adoption follows financial necessity. They are largely independent, which is exactly why this project measures the technical and economic pillars separately." `
                -left 76 -top ($tkTop + 26) -width 808 -height 30 `
                -font $FONT_BODY -size 11 -color (RGB '#d0dae3'))

# Save -----------------------------------------------------------------------
$out = (Join-Path (Get-Location) "mining_vs_adoption.pptx")
if (Test-Path $out) { Remove-Item $out -Force }
$pres.SaveAs($out, 24)  # ppSaveAsOpenXMLPresentation
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Host ""
Write-Host "Wrote: $out"
