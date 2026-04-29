# =============================================================================
# build_pptx.ps1
# Builds presentation.pptx from scratch using PowerPoint COM.
# Run from the repo root:  powershell -ExecutionPolicy Bypass -File .\build_pptx.ps1
# =============================================================================

$ErrorActionPreference = "Stop"

# ---------- helpers ----------

function RGB([string]$hex) {
    $h = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($h.Substring(0,2), 16)
    $g = [Convert]::ToInt32($h.Substring(2,2), 16)
    $b = [Convert]::ToInt32($h.Substring(4,2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

# Color palette (matching the website / HTML deck)
$BG_LIGHT       = RGB '#fafaf7'
$BG_DARK        = RGB '#0a1620'
$BG_TINT        = RGB '#f4f0e8'
$BG_CARD        = RGB '#ffffff'
$BG_CARD_DARK   = RGB '#122433'
$BG_FORMULA     = RGB '#0a1620'
$INK            = RGB '#1a2a3a'
$INK_BODY       = RGB '#2c3e50'
$INK_MUTED      = RGB '#5a6b7c'
$INK_DIM        = RGB '#7a8a9a'
$INK_DARK_BG    = RGB '#f0f4f8'
$INK_DARK_BODY  = RGB '#d0dae3'
$INK_DARK_MUTED = RGB '#9bdce8'
$BLUE_DEEP      = RGB '#0a4c6a'
$BLUE_MID       = RGB '#1a6b7a'
$TEAL_LIGHT     = RGB '#5ba8b5'
$WARM           = RGB '#9b5a30'
$WARM_MID       = RGB '#d4915c'
$CORAL          = RGB '#c06030'
$GREEN          = RGB '#1a7a4e'
$TECH_BLUE      = RGB '#2e5fa1'
$RED            = RGB '#8b1a1a'
$BORDER_LIGHT   = RGB '#ebebeb'
$BORDER_TINT    = RGB '#ebe6db'

# Slide dimensions in points (16:9 widescreen)
$SLIDE_W = 960
$SLIDE_H = 540

# Default fonts (most universal across Windows + Google Slides)
$FONT_HEAD = "Cambria"      # serif headline (closest to Source Serif 4)
$FONT_BODY = "Calibri"
$FONT_MONO = "Consolas"

# ---------- start PowerPoint ----------
Write-Host "Launching PowerPoint..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1   # COM can't add slides reliably with hidden window; PPT will close at the end

$pres = $ppt.Presentations.Add($true)
$pres.PageSetup.SlideSize = 15  # ppSlideSize16x9

# Track current slide index for incrementing layout
$slideIndex = 0

# Remove the default starter slide if any
while ($pres.Slides.Count -gt 0) { $pres.Slides.Item(1).Delete() }

# ---------- low-level helpers ----------

function New-BlankSlide {
    param([int]$bgColor = $BG_LIGHT)
    $script:slideIndex++
    $s = $script:pres.Slides.Add($script:slideIndex, 12)  # 12 = ppLayoutBlank
    $s.FollowMasterBackground = $false
    $s.Background.Fill.Visible = -1
    $s.Background.Fill.ForeColor.RGB = $bgColor
    return $s
}

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
        [int]$align = 1,   # 1=left, 2=center, 3=right
        [double]$tracking = 0,
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
    # PowerPoint COM doesn't expose letter-spacing on Font (tracking arg accepted but ignored).
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
        $shape = $slide.Shapes.AddShape(5, $left, $top, $width, $height)  # 5 = msoShapeRoundedRectangle
        try { $shape.Adjustments.Item(1) = $cornerRadius } catch {}
    } else {
        $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)  # 1 = msoShapeRectangle
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

function Add-Line {
    param(
        [Parameter(Mandatory=$true)] $slide,
        [double]$x1, [double]$y1, [double]$x2, [double]$y2,
        [int]$color, [double]$weight = 1
    )
    $ln = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $ln.Line.ForeColor.RGB = $color
    $ln.Line.Weight = $weight
    return $ln
}

function Add-Eyebrow {
    param($slide, [string]$text, [int]$color = $BLUE_DEEP, [double]$top = 60)
    Add-Text -slide $slide -text $text -left 60 -top $top -width 700 -height 18 `
             -font $FONT_MONO -size 11 -color $color -bold $true -tracking 2 -upper $true | Out-Null
}

function Add-PresenterTag {
    param($slide, [string]$who, [int]$color = $BLUE_DEEP, [bool]$onDark = $false)
    $base = if ($onDark) { RGB '#9bdce8' } else { $INK_DIM }
    $tb = $slide.Shapes.AddTextbox(1, $SLIDE_W - 360, 28, 340, 16)
    $tb.TextFrame.MarginLeft = 0; $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop = 0;  $tb.TextFrame.MarginBottom = 0
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0
    $r = $tb.TextFrame.TextRange
    $r.Text = "PRESENTED BY: $who"
    $r.Font.Name = $FONT_MONO
    $r.Font.Size = 9
    $r.Font.Color.RGB = $base
    $r.ParagraphFormat.Alignment = 3  # right
    # Highlight the name with brand color
    $idx = "PRESENTED BY: ".Length + 1
    $name = $r.Characters($idx, $r.Text.Length - $idx + 1)
    $name.Font.Color.RGB = $color
    $name.Font.Bold = -1
    return $tb
}

function Add-Heading {
    param($slide, [string]$text, [double]$top = 100, [double]$size = 30, [int]$color = $INK)
    Add-Text -slide $slide -text $text -left 60 -top $top -width 840 -height 100 `
             -font $FONT_HEAD -size $size -color $color -bold $true | Out-Null
}

function Add-Lede {
    param($slide, [string]$text, [double]$top, [int]$color = $INK_BODY, [double]$width = 840)
    Add-Text -slide $slide -text $text -left 60 -top $top -width $width -height 60 `
             -font $FONT_BODY -size 14 -color $color -italic $false | Out-Null
}

function Add-Body {
    param($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height,
          [int]$color = $INK_BODY, [double]$size = 12, [bool]$bold = $false)
    Add-Text -slide $slide -text $text -left $left -top $top -width $width -height $height `
             -font $FONT_BODY -size $size -color $color -bold $bold | Out-Null
}

# Pillar card: top accent stripe + card body + title + weight + body
function Add-PillarCard {
    param(
        $slide,
        [double]$left, [double]$top, [double]$width, [double]$height,
        [string]$title, [string]$weightLabel, [string]$desc,
        [int]$accent
    )
    $card = Add-Rect -slide $slide -left $left -top $top -width $width -height $height `
                    -fill $BG_CARD -borderColor $BORDER_LIGHT -borderWidth 0.5
    # accent stripe (top)
    Add-Rect -slide $slide -left $left -top $top -width $width -height 4 `
             -fill $accent -cornerRadius 0 | Out-Null
    Add-Text -slide $slide -text $title -left ($left + 14) -top ($top + 14) `
             -width ($width - 28) -height 22 -font $FONT_HEAD -size 14 -color $INK -bold $true | Out-Null
    Add-Text -slide $slide -text $weightLabel -left ($left + 14) -top ($top + 36) `
             -width ($width - 28) -height 12 -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -tracking 1.4 | Out-Null
    Add-Text -slide $slide -text $desc -left ($left + 14) -top ($top + 54) `
             -width ($width - 28) -height ($height - 60) -font $FONT_BODY -size 10 -color $INK_BODY | Out-Null
}

# Indicator row: weight | name + desc
function Add-IndicatorRow {
    param($slide, [double]$top, [string]$weight, [string]$name, [string]$desc, [int]$accent)
    Add-Text -slide $slide -text $weight -left 60 -top $top -width 90 -height 30 `
             -font $FONT_MONO -size 22 -color $accent -bold $true | Out-Null
    Add-Text -slide $slide -text $name -left 160 -top $top -width 740 -height 18 `
             -font $FONT_HEAD -size 13 -color $INK -bold $true | Out-Null
    Add-Text -slide $slide -text $desc -left 160 -top ($top + 20) -width 740 -height 18 `
             -font $FONT_BODY -size 10 -color $INK_MUTED | Out-Null
    # subtle separator below
    Add-Line -slide $slide -x1 60 -y1 ($top + 42) -x2 900 -y2 ($top + 42) -color $BORDER_TINT -weight 0.5 | Out-Null
}

# Big stat
function Add-BigStat {
    param($slide, [double]$left, [double]$top, [double]$width, [double]$height,
          [string]$num, [string]$label, [int]$accent = $BLUE_DEEP, [int]$bg = $BG_CARD,
          [int]$labelColor = $INK_BODY)
    $card = Add-Rect -slide $slide -left $left -top $top -width $width -height $height `
                    -fill $bg -borderColor -1
    # left accent stripe
    Add-Rect -slide $slide -left $left -top $top -width 5 -height $height -fill $accent -cornerRadius 0 | Out-Null
    Add-Text -slide $slide -text $num -left ($left + 16) -top ($top + 14) `
             -width ($width - 32) -height 56 -font $FONT_MONO -size 42 -color $accent -bold $true | Out-Null
    Add-Text -slide $slide -text $label -left ($left + 16) -top ($top + 72) `
             -width ($width - 32) -height ($height - 78) -font $FONT_BODY -size 10.5 -color $labelColor | Out-Null
}

# Leader row (rank, country, driver, score)
function Add-LeaderRow {
    param($slide, [double]$top, [string]$rank, [string]$country, [string]$driver, [string]$score, [int]$accent = $BLUE_MID)
    Add-Rect -slide $slide -left 60 -top $top -width 840 -height 42 -fill $BG_CARD -cornerRadius 0.05 | Out-Null
    Add-Rect -slide $slide -left 60 -top $top -width 4 -height 42 -fill $accent -cornerRadius 0 | Out-Null
    Add-Text -slide $slide -text $rank -left 76 -top ($top + 12) -width 50 -height 22 `
             -font $FONT_MONO -size 16 -color $BLUE_DEEP -bold $true | Out-Null
    Add-Text -slide $slide -text $country -left 130 -top ($top + 12) -width 180 -height 22 `
             -font $FONT_HEAD -size 13 -color $INK -bold $true | Out-Null
    Add-Text -slide $slide -text $driver -left 320 -top ($top + 14) -width 480 -height 22 `
             -font $FONT_BODY -size 10 -color $INK_MUTED | Out-Null
    Add-Text -slide $slide -text $score -left 800 -top ($top + 12) -width 90 -height 22 `
             -font $FONT_MONO -size 16 -color $BLUE_DEEP -bold $true -align 3 | Out-Null
}

# Tier card (Comprehensive / Developing / Ambiguous / Hostile)
function Add-TierCard {
    param($slide, [double]$left, [double]$top, [double]$width, [double]$height,
          [string]$range, [string]$tierName, [string]$desc, [int]$accent)
    Add-Rect -slide $slide -left $left -top $top -width $width -height $height -fill $BG_CARD -cornerRadius 0.06 | Out-Null
    Add-Rect -slide $slide -left $left -top $top -width 4 -height $height -fill $accent -cornerRadius 0 | Out-Null
    Add-Text -slide $slide -text $range -left ($left + 14) -top ($top + 12) -width ($width - 28) -height 18 `
             -font $FONT_MONO -size 13 -color $accent -bold $true | Out-Null
    Add-Text -slide $slide -text $tierName -left ($left + 14) -top ($top + 30) -width ($width - 28) -height 22 `
             -font $FONT_HEAD -size 13 -color $INK -bold $true | Out-Null
    Add-Text -slide $slide -text $desc -left ($left + 14) -top ($top + 54) -width ($width - 28) -height ($height - 60) `
             -font $FONT_BODY -size 10 -color $INK_MUTED | Out-Null
}

# vs/then card (when -> then -> body)
function Add-VsCard {
    param($slide, [double]$left, [double]$top, [double]$width, [double]$height,
          [string]$when, [string]$then, [string]$body, [int]$accent = $BLUE_MID,
          [int]$bg = $BG_CARD, [int]$inkColor = $INK, [int]$bodyColor = $INK_MUTED, [int]$whenColor = -1)
    Add-Rect -slide $slide -left $left -top $top -width $width -height $height -fill $bg -cornerRadius 0.04 | Out-Null
    Add-Rect -slide $slide -left $left -top $top -width $width -height 4 -fill $accent -cornerRadius 0 | Out-Null
    if ($whenColor -lt 0) { $whenColor = $INK_DIM }
    Add-Text -slide $slide -text $when -left ($left + 16) -top ($top + 14) -width ($width - 32) -height 14 `
             -font $FONT_MONO -size 9 -color $whenColor -upper $true -tracking 1.6 | Out-Null
    Add-Text -slide $slide -text $then -left ($left + 16) -top ($top + 30) -width ($width - 32) -height 38 `
             -font $FONT_HEAD -size 14 -color $inkColor -bold $true | Out-Null
    Add-Text -slide $slide -text $body -left ($left + 16) -top ($top + 72) -width ($width - 32) -height ($height - 80) `
             -font $FONT_BODY -size 10 -color $bodyColor | Out-Null
}

# Bulleted list (custom bullets, simple horizontal dash)
function Add-BulletList {
    param($slide, [double]$left, [double]$top, [double]$width, [string[]]$items,
          [int]$color = $INK_BODY, [double]$size = 11, [int]$accent = $BLUE_DEEP, [double]$rowHeight = 36)
    for ($i = 0; $i -lt $items.Length; $i++) {
        $y = $top + ($i * $rowHeight)
        # dash bullet
        Add-Line -slide $slide -x1 $left -y1 ($y + 8) -x2 ($left + 14) -y2 ($y + 8) -color $accent -weight 1.5 | Out-Null
        Add-Text -slide $slide -text $items[$i] -left ($left + 22) -top $y -width ($width - 22) -height ($rowHeight - 4) `
                 -font $FONT_BODY -size $size -color $color | Out-Null
    }
}

# =============================================================================
# SLIDES
# =============================================================================

# ---------- 01 Title ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-Eyebrow -slide $s -text "TAC 456  -  Capstone Seminar  -  April 2026" -color $TEAL_LIGHT -top 70
Add-Text -slide $s -text "Blockchain" -left 60 -top 120 -width 840 -height 60 `
         -font $FONT_HEAD -size 56 -color $INK_DARK_BG -bold $true | Out-Null
Add-Text -slide $s -text "Adoption Atlas" -left 60 -top 184 -width 840 -height 70 `
         -font $FONT_HEAD -size 64 -color $TEAL_LIGHT -bold $true | Out-Null
Add-Text -slide $s -text "Measuring blockchain readiness across economic, regulatory, and technical dimensions in 20 countries." `
         -left 60 -top 280 -width 760 -height 60 -font $FONT_BODY -size 18 -color (RGB '#c0d2dc') | Out-Null
Add-Text -slide $s -text "FEDERICO BECERRA  -  NICO TAYLOR  -  JADEN CHOI  -  ALEXIS VO  -  ALON MUTTER  -  ARI CHOE" `
         -left 60 -top 460 -width 840 -height 16 -font $FONT_MONO -size 11 -color (RGB '#7e8c98') -tracking 1.6 | Out-Null

# ---------- 02 Problem ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS"
Add-Eyebrow -slide $s -text "01  -  The Problem"
Add-Heading -slide $s -text "Existing rankings answer the wrong question." -top 90 -size 30
Add-Lede -slide $s -text "Current blockchain adoption metrics are fragmented and conflate speculation with genuine utility." -top 162
Add-Text -slide $s -text "What others measure" -left 60 -top 220 -width 400 -height 22 -font $FONT_HEAD -size 16 -color $INK -bold $true | Out-Null
Add-BulletList -slide $s -left 60 -top 254 -width 400 -items @(
    "Chainalysis: composite of on-chain activity, weighted toward CEX volume and DeFi flows.",
    "Triple-A: a single ownership-percentage statistic per country.",
    "Country reports (PwC, Library of Congress): legal status, no behavioral data."
) -size 11 -rowHeight 50
Add-Text -slide $s -text "What is missing" -left 500 -top 220 -width 400 -height 22 -font $FONT_HEAD -size 16 -color $INK -bold $true | Out-Null
Add-BulletList -slide $s -left 500 -top 254 -width 400 -items @(
    "None separate why a country ranks high. India, Vietnam, and the US sit near the top of Chainalysis for completely different reasons.",
    "None combine usage, legal clarity, and infrastructure into one comparable picture.",
    "None reveal the tradeoff between regulation and adoption."
) -size 11 -rowHeight 56

# ---------- 03 Question ----------
$s = New-BlankSlide -bgColor $BG_TINT
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS"
Add-Eyebrow -slide $s -text "02  -  Our Question" -color $WARM
Add-Text -slide $s -text "How should global blockchain adoption be measured and visualized across economic, technical, and regulatory dimensions?" `
         -left 60 -top 130 -width 840 -height 200 -font $FONT_HEAD -size 32 -color $INK -bold $true | Out-Null
Add-Lede -slide $s -text "A defensible answer requires keeping the three dimensions separable, not collapsing them into a single number that hides which one is doing the work." -top 380

# ---------- 04 Thesis ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS" -color $TEAL_LIGHT -onDark $true
Add-Eyebrow -slide $s -text "03  -  Our Thesis" -color $TEAL_LIGHT
Add-Heading -slide $s -text "What we expected vs. what we found." -top 90 -size 28 -color $INK_DARK_BG
Add-VsCard -slide $s -left 60 -top 170 -width 400 -height 240 `
           -when "Original thesis" -then "High trading volume would not correlate with strong technical infrastructure or regulatory clarity." `
           -body "i.e., countries with the most activity are not the ones with the best frameworks for it." `
           -accent (RGB '#7e8c98') -bg $BG_CARD_DARK -inkColor $INK_DARK_BG -bodyColor $INK_DARK_BODY -whenColor (RGB '#9aa5b0')
Add-VsCard -slide $s -left 500 -top 170 -width 400 -height 240 `
           -when "Revised thesis" -then "High crypto activity does not mean a country is truly blockchain-ready." `
           -body "Some countries are active out of necessity, others have built real infrastructure, and you cannot tell the difference without looking at all three dimensions separately." `
           -accent $CORAL -bg $BG_CARD_DARK -inkColor $INK_DARK_BG -bodyColor $INK_DARK_BODY -whenColor (RGB '#9aa5b0')
Add-Text -slide $s -text "The original thesis was partially confirmed. The revised thesis is what the data actually showed." `
         -left 60 -top 440 -width 840 -height 30 -font $FONT_BODY -size 12 -color $INK_DARK_MUTED -italic $true | Out-Null

# ---------- 05 Method ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS"
Add-Eyebrow -slide $s -text "04  -  Method"
Add-Heading -slide $s -text "Three pillars, weighted equally." -top 90 -size 30
Add-Lede -slide $s -text "Each country gets one score per pillar, normalized to 0 to 1. The overall score is their average. The point is to keep the pillars separable, so the underlying driver of any country's position stays visible." -top 162
Add-PillarCard -slide $s -left 60 -top 250 -width 270 -height 150 `
               -title "Economic Adoption" -weightLabel "Weight 1/3" `
               -desc "Real-world crypto use: ownership, exchange volume per GDP, remittances, Chainalysis ranks." `
               -accent $GREEN
Add-PillarCard -slide $s -left 345 -top 250 -width 270 -height 150 `
               -title "Regulatory Clarity" -weightLabel "Weight 1/3" `
               -desc "Six legal sub-dimensions: stablecoins, licensing, tax, banking, enforcement, DeFi." `
               -accent $WARM
Add-PillarCard -slide $s -left 630 -top 250 -width 270 -height 150 `
               -title "Technical Infrastructure" -weightLabel "Weight 1/3" `
               -desc "BTC and ETH nodes, mining hashrate, Lightning density, electricity, internet quality." `
               -accent $TECH_BLUE
# formula box
Add-Rect -slide $s -left 60 -top 420 -width 840 -height 50 -fill $BG_FORMULA -cornerRadius 0.04 | Out-Null
Add-Text -slide $s -text "Overall = (Economic x 0.333) + (Regulatory x 0.333) + (Technical x 0.333)" `
         -left 76 -top 432 -width 800 -height 28 -font $FONT_MONO -size 16 -color (RGB '#d0dae3') | Out-Null

# ---------- 06 Pillar I Economic ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "NICO TAYLOR" -color $GREEN
Add-Eyebrow -slide $s -text "05  -  Pillar I  -  Economic Dimension" -color $GREEN
Add-Heading -slide $s -text "How much economic value is flowing through blockchain, and is it speculation or productive use?" -top 90 -size 22
Add-IndicatorRow -slide $s -top 195 -weight "30%" -name "Crypto Ownership" -desc "Population penetration rate from Triple-A. The strongest signal of real consumer adoption." -accent $BLUE_DEEP
Add-IndicatorRow -slide $s -top 245 -weight "20%" -name "CEX Activity" -desc "Average of Chainalysis Retail and Institutional CEX sub-rankings, PPP-adjusted. Captures broad market participation." -accent $BLUE_DEEP
Add-IndicatorRow -slide $s -top 295 -weight "20%" -name "CEX Volume / GDP" -desc "CoinGecko 24h volume divided by national GDP, log-adjusted. Catches Lira, Won, and Naira-style speculation." -accent $BLUE_DEEP
Add-IndicatorRow -slide $s -top 345 -weight "20%" -name "Remittances per Capita" -desc "World Bank inbound flows divided by population. Indicates real-world cross-border utility (Ukraine, Philippines, Bangladesh)." -accent $BLUE_DEEP
Add-IndicatorRow -slide $s -top 395 -weight "10%" -name "DeFi Activity" -desc "Chainalysis DeFi value-received sub-rank. Down-weighted because it is volatile and concentrated in a small group." -accent $BLUE_DEEP

# ---------- 07 Pillar II Technical ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "FEDERICO BECERRA" -color $TECH_BLUE
Add-Eyebrow -slide $s -text "06  -  Pillar II  -  Technical Infrastructure" -color $TECH_BLUE
Add-Heading -slide $s -text "How is blockchain infrastructure being built and where?" -top 90 -size 28
Add-IndicatorRow -slide $s -top 175 -weight "25%" -name "Bitcoin Nodes" -desc "Reachable BTC nodes per 100k internet users (Bitnodes)." -accent $TECH_BLUE
Add-IndicatorRow -slide $s -top 225 -weight "25%" -name "Ethereum Nodes" -desc "Reachable ETH nodes per 100k internet users (Ethernodes)." -accent $TECH_BLUE
Add-IndicatorRow -slide $s -top 275 -weight "20%" -name "Lightning Network" -desc "LN nodes per 100k internet users (mempool.emzy.de). A proxy for second-layer adoption depth." -accent $TECH_BLUE
Add-IndicatorRow -slide $s -top 325 -weight "15%" -name "Mining Hashrate" -desc "Country share of global BTC hashrate (Hashrate Index). Used directly because mining is geographically concentrated." -accent $TECH_BLUE
Add-IndicatorRow -slide $s -top 375 -weight "15%" -name "Enabling Environment" -desc "Composite of internet penetration (40%), electricity access (30%), and average internet speed (30%)." -accent $TECH_BLUE

# ---------- 08 Technical Reasoning ----------
$s = New-BlankSlide -bgColor $BG_TINT
Add-PresenterTag -slide $s -who "FEDERICO BECERRA" -color $TECH_BLUE
Add-Eyebrow -slide $s -text "07  -  Why this technical weighting" -color $TECH_BLUE
Add-Heading -slide $s -text "The reasoning behind the five technical indicators." -top 90 -size 26
Add-BulletList -slide $s -left 60 -top 180 -width 840 -accent $TECH_BLUE -size 13 -rowHeight 64 -items @(
    "Balances decentralization and economic commitment. Nodes (BTC / ETH / LN) measure decentralization and participation, while hashrate captures capital-intensive commitment to the network.",
    "Per-100k-internet-users normalization improves cross-country comparability. Makes small but highly engaged countries comparable to large economies.",
    "Incorporates real-world infrastructure constraints. Crypto infrastructure cannot scale without underlying connectivity and energy systems.",
    "Weights reflect ecosystem importance. Prioritizes core protocol infrastructure (BTC and ETH at 25% each), while including scalability and security layers (LN 20%, Mining 15%)."
)

# ---------- 09 Pillar III Regulatory ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS" -color $WARM
Add-Eyebrow -slide $s -text "08  -  Pillar III  -  Regulatory Dimension" -color $WARM
Add-Heading -slide $s -text "How clear, comprehensive, and enforceable is each country's framework?" -top 90 -size 22
Add-Lede -slide $s -text "Each country is rated on six sub-dimensions, scored 1 (absent / hostile), 2 (partial), or 3 (clear and codified). Sums divided by 18 normalize to 0 to 1. Equal weight on each: no single dimension is sufficient on its own." -top 165
# 3x2 grid of sub-dimension cards
$rWidths = 270; $rHeight = 95
$cols = @(60, 345, 630)
$rows = @(245, 350)
$labels = @(
    @("Exchange Legality",         "Whether centralized exchanges can operate legally at all in the jurisdiction."),
    @("Licensing Framework (VASP)","Defined licensing regime for centralized exchanges and other Virtual Asset Service Providers."),
    @("Stablecoin Regulation",     "Specific legislation governing stablecoin issuance, reserves, and consumer protection."),
    @("Tax Clarity",               "Codified rules for taxing crypto gains, income, and corporate holdings."),
    @("Banking System Access",     "Whether crypto businesses can access traditional banking rails legally."),
    @("Enforcement & Stance",      "Active, observable enforcement of the rules that exist on paper, plus government posture.")
)
for ($i = 0; $i -lt 6; $i++) {
    $col = $i % 3; $row = [math]::Floor($i / 3)
    Add-PillarCard -slide $s -left $cols[$col] -top $rows[$row] -width $rWidths -height $rHeight `
                   -title $labels[$i][0] -weightLabel "Weight 1/6" -desc $labels[$i][1] -accent $WARM
}

# ---------- 10 Score Interpretation ----------
$s = New-BlankSlide -bgColor $BG_TINT
Add-PresenterTag -slide $s -who "ALON / ARI / ALEXIS" -color $WARM
Add-Eyebrow -slide $s -text "09  -  Score Interpretation" -color $WARM
Add-Heading -slide $s -text "Higher scores reflect clarity and usability, not strictness." -top 90 -size 24
# 2x2 grid
$tw = 410; $th = 130
Add-TierCard -slide $s -left 60 -top 180 -width $tw -height $th -range "0.8 - 1.0" -tierName "Comprehensive" -desc "Clear legal frameworks across all categories. Businesses and users can operate with high certainty." -accent (RGB '#2e7d32')
Add-TierCard -slide $s -left 490 -top 180 -width $tw -height $th -range "0.6 - 0.8" -tierName "Developing" -desc "Partial regulation with some gaps. Key areas defined, but inconsistencies remain." -accent $WARM_MID
Add-TierCard -slide $s -left 60 -top 330 -width $tw -height $th -range "0.4 - 0.6" -tierName "Ambiguous" -desc "Unclear or fragmented rules. High uncertainty for users and companies." -accent $CORAL
Add-TierCard -slide $s -left 490 -top 330 -width $tw -height $th -range "0.2 - 0.4" -tierName "Hostile" -desc "Restrictive or prohibitive policies. Limited or no legal pathway for participation." -accent $RED
Add-Text -slide $s -text "Five of 20 reach Comprehensive: Japan (0.944), Ukraine (0.889), Brazil (0.889), Thailand (0.833), United Kingdom (0.833)." `
         -left 60 -top 478 -width 840 -height 24 -font $FONT_BODY -size 11 -color $INK_MUTED -italic $true | Out-Null

# ---------- 11 Demo ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-PresenterTag -slide $s -who "JADEN CHOI" -color $TEAL_LIGHT -onDark $true
Add-Eyebrow -slide $s -text "10  -  Live Demo" -color $TEAL_LIGHT
Add-Heading -slide $s -text "The deliverable is an interactive web atlas." -top 90 -size 28 -color $INK_DARK_BG
Add-Lede -slide $s -text "Five linked surfaces: home, map, country profiles, findings, and three methodology pages." -top 158 -color $INK_DARK_BODY
# demo card
Add-Rect -slide $s -left 60 -top 220 -width 840 -height 280 -fill $BG_CARD_DARK -borderColor (RGB '#1f3445') -borderWidth 0.5 -cornerRadius 0.04 | Out-Null
Add-Text -slide $s -text "Demo walkthrough" -left 78 -top 240 -width 800 -height 22 -font $FONT_HEAD -size 16 -color $INK_DARK_BG -bold $true | Out-Null
$demoSteps = @(
    @("01", "Map view.", "Show the geographic heat map. Switch dimension tabs (Overall / Economic / Regulatory / Technical) to show how the same country shifts color across pillars. Hover for tier labels."),
    @("02", "Map features.", "Search countries from the rankings table; compare ranks across the three dimensions; click into any country."),
    @("03", "Country profile.", "Open Vietnam (#1 economic) and Japan (#1 regulatory) to compare radar shapes and the 'What this score means' section."),
    @("04", "Findings page.", "Skim the six findings to set up the next slides.")
)
for ($i = 0; $i -lt 4; $i++) {
    $y = 280 + ($i * 50)
    Add-Rect -slide $s -left 78 -top ($y - 2) -width 38 -height 24 -fill (RGB '#1f3445') -cornerRadius 0.18 | Out-Null
    Add-Text -slide $s -text $demoSteps[$i][0] -left 78 -top ($y + 2) -width 38 -height 20 -font $FONT_MONO -size 12 -color $TEAL_LIGHT -align 2 -bold $true | Out-Null
    Add-Text -slide $s -text $demoSteps[$i][1] -left 130 -top ($y + 2) -width 200 -height 20 -font $FONT_BODY -size 11 -color $TEAL_LIGHT -bold $true | Out-Null
    Add-Text -slide $s -text $demoSteps[$i][2] -left 330 -top ($y + 2) -width 550 -height 40 -font $FONT_BODY -size 10 -color $INK_DARK_BODY | Out-Null
}

# ---------- 12 Data Breakdowns ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "FEDERICO & NICO"
Add-Eyebrow -slide $s -text "11  -  Data Breakdowns"
Add-Heading -slide $s -text "Three full data pages, one per pillar." -top 90 -size 28
Add-Lede -slide $s -text "Every claim on the site is auditable. The methodology pages expose the raw indicators, normalization, and source citations behind each score." -top 162
Add-PillarCard -slide $s -left 60 -top 245 -width 270 -height 170 -title "econ-methodology.html" -weightLabel "Pillar I  -  Economic" -desc "Raw ownership, volume, remittance, and Chainalysis sub-rank tables for all 20 countries. Three views: scored, raw, and normalized." -accent $GREEN
Add-PillarCard -slide $s -left 345 -top 245 -width 270 -height 170 -title "regulatory-methodology.html" -weightLabel "Pillar III  -  Regulatory" -desc "Per-country 1 to 3 ratings on each of the six sub-dimensions, with source citations from Library of Congress and PwC." -accent $WARM
Add-PillarCard -slide $s -left 630 -top 245 -width 270 -height 170 -title "tech-methodology.html" -weightLabel "Pillar II  -  Technical" -desc "Node counts, hashrate share, Lightning density, and the enabling-environment composite. All values per 100k internet users where applicable." -accent $TECH_BLUE
Add-Text -slide $s -text "The full data is also accessible via the country profile pages, which surface a per-country rollup with links to the underlying methodology." `
         -left 60 -top 432 -width 840 -height 30 -font $FONT_BODY -size 11 -color $INK_MUTED | Out-Null

# ---------- 13 Section Header: Findings ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-Text -slide $s -text "PART TWO" -left 60 -top 130 -width 300 -height 22 `
         -font $FONT_MONO -size 14 -color $TEAL_LIGHT -tracking 4 -bold $true | Out-Null
Add-Text -slide $s -text "Key Findings" -left 60 -top 180 -width 840 -height 110 `
         -font $FONT_HEAD -size 80 -color $INK_DARK_BG -bold $true | Out-Null
Add-Text -slide $s -text "Six insights from analyzing 20 countries across the three dimensions." `
         -left 60 -top 320 -width 800 -height 30 -font $FONT_BODY -size 16 -color $INK_DARK_MUTED -italic $true | Out-Null

# ---------- 14 Finding I (Alon) ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "ALON MUTTER"
Add-Eyebrow -slide $s -text "12  -  Finding I  -  Adoption vs. Regulation"
Add-Heading -slide $s -text "An inverse relationship, with two notable exceptions." -top 90 -size 24
Add-Lede -slide $s -text "Heavy regulation slows consumer-led adoption. Light regulation enables rapid but potentially risky growth." -top 158
Add-VsCard -slide $s -left 60 -top 220 -width 410 -height 200 `
           -when "The exceptions" -then "Ukraine & Brazil" `
           -body "The only countries achieving both high adoption (#2 and #11 economic) and Comprehensive regulation (0.889 each). Both reached this through proactive frameworks responding to existing demand." `
           -accent $BLUE_MID
Add-VsCard -slide $s -left 490 -top 220 -width 410 -height 200 `
           -when "The pattern" -then "High adoption, low clarity" `
           -body "Vietnam (0.761 econ / 0.556 reg), Pakistan (0.494 / 0.556), Argentina (0.262 / 0.611). Activity outruns rules wherever economic necessity is the driver." `
           -accent $CORAL
Add-Text -slide $s -text "The takeaway. The ideal balance, high adoption with clear rules, remains rare globally. Most countries pick one or the other." `
         -left 60 -top 440 -width 840 -height 30 -font $FONT_BODY -size 12 -color $INK_BODY -bold $true | Out-Null

# ---------- 15 Finding II (Ari) ----------
$s = New-BlankSlide -bgColor $BG_TINT
Add-PresenterTag -slide $s -who "ARI CHOE" -color $WARM
Add-Eyebrow -slide $s -text "13  -  Finding II  -  Developing Nations Lead" -color $WARM
Add-Heading -slide $s -text "Every top-five adopter has a distress driver." -top 90 -size 26
Add-Lede -slide $s -text "The leaders look nothing like a list of 'most regulated' countries. The US is the only stable, developed economy in the top five." -top 158
Add-LeaderRow -slide $s -top 215 -rank "#1" -country "Vietnam" -driver "Highest crypto ownership in the world (21.2%) + remittance corridors" -score "0.761"
Add-LeaderRow -slide $s -top 263 -rank "#2" -country "Ukraine" -driver "Active conflict, currency stress, \$503/capita remittances (highest in index)" -score "0.733"
Add-LeaderRow -slide $s -top 311 -rank "#3" -country "Philippines" -driver "Major remittance corridor (\$341/capita), 13.4% ownership" -score "0.632"
Add-LeaderRow -slide $s -top 359 -rank "#4" -country "United States" -driver "Outlier: 15.6% ownership, \$76.6B daily CEX volume, deep institutional flows" -score "0.576"
Add-LeaderRow -slide $s -top 407 -rank "#5" -country "Pakistan" -driver "Currency volatility, financial exclusion, top-3 retail and institutional CEX ranks" -score "0.494"
Add-Text -slide $s -text "Three structural drivers recur: currency instability (Venezuela 10.3%, Argentina 9.7% ownership), remittance corridors, and financial exclusion -- all stronger predictors than favorable policy." `
         -left 60 -top 460 -width 840 -height 26 -font $FONT_BODY -size 11 -color $INK_BODY | Out-Null

# ---------- 16 Finding III (Alexis) ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "ALEXIS VO"
Add-Eyebrow -slide $s -text "14  -  Finding III  -  The Regulatory Landscape"
Add-Heading -slide $s -text "Only 5 of 20 countries have comprehensive regulation." -top 90 -size 28
Add-BigStat -slide $s -left 60 -top 200 -width 270 -height 220 -num "5/20" `
            -label "Comprehensive only. Japan (0.944), Ukraine (0.889), Brazil (0.889), Thailand (0.833), UK (0.833). The remaining 15 are Developing, Ambiguous, or Hostile." -accent $BLUE_DEEP
Add-BigStat -slide $s -left 345 -top 200 -width 270 -height 220 -num "0/20" `
            -label "DeFi recognition. No country in the index scores a 3 on DeFi. Traditional legal frameworks cannot classify decentralized protocols." -accent $CORAL
Add-BigStat -slide $s -left 630 -top 200 -width 270 -height 220 -num "3" `
            -label "Hostile-tier countries. Bangladesh, Yemen, and Ethiopia all score 0.333 across all six sub-dimensions. Operating there carries maximum legal risk." -accent $BLUE_MID
Add-Text -slide $s -text "The takeaway. Even the most regulated jurisdictions still have a global blind spot: stablecoin and DeFi rules. That is the next regulatory frontier everywhere." `
         -left 60 -top 446 -width 840 -height 30 -font $FONT_BODY -size 12 -color $INK_BODY -bold $true | Out-Null

# ---------- 17 Finding IV (Jaden) ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-PresenterTag -slide $s -who "JADEN CHOI" -color $TEAL_LIGHT -onDark $true
Add-Eyebrow -slide $s -text "15  -  Finding IV  -  Infrastructure Concentration" -color $TEAL_LIGHT
Add-Heading -slide $s -text "Crypto's 'decentralized' layer is geographically concentrated." -top 90 -size 26 -color $INK_DARK_BG
Add-BigStat -slide $s -left 60 -top 180 -width 320 -height 250 -num "58.6%" `
            -label "of global Bitcoin hashrate is produced by just four countries: US (37.5%), Russia (16.4%), Ethiopia (2.6%), Indonesia (2.1%)." `
            -accent $TEAL_LIGHT -bg $BG_CARD_DARK -labelColor $INK_DARK_BODY
Add-Text -slide $s -text "Exception: The Ethiopia Paradox" -left 410 -top 190 -width 480 -height 28 -font $FONT_HEAD -size 18 -color $INK_DARK_BG -bold $true | Out-Null
Add-Text -slide $s -text "Hostile regulation (0.333). Near-zero consumer adoption (0.173). Yet Ethiopia produces 2.6% of global Bitcoin hashrate, powered by cheap hydroelectric energy from the Grand Ethiopian Renaissance Dam." `
         -left 410 -top 224 -width 480 -height 70 -font $FONT_BODY -size 12 -color $INK_DARK_BODY | Out-Null
Add-Text -slide $s -text "Indonesia mirrors this pattern: 2.1% global hashrate driven by geothermal energy, while also being the #1 country for DeFi adoption per Chainalysis." `
         -left 410 -top 304 -width 480 -height 60 -font $FONT_BODY -size 12 -color $INK_DARK_BODY | Out-Null
Add-Text -slide $s -text "Mining concentration is now an energy-policy and geopolitical question, not just a technical one." `
         -left 410 -top 374 -width 480 -height 50 -font $FONT_BODY -size 12 -color $INK_DARK_MUTED -italic $true | Out-Null

# ---------- 18 Finding V (Nico) ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "NICO TAYLOR" -color $GREEN
Add-Eyebrow -slide $s -text "16  -  Finding V  -  Crisis Drives Adoption" -color $GREEN
Add-Heading -slide $s -text "Sanctions and economic crisis are stronger catalysts than policy." -top 90 -size 24
Add-Lede -slide $s -text "Restrictive rules alone cannot prevent adoption when citizens face economic necessity." -top 158
Add-VsCard -slide $s -left 60 -top 220 -width 410 -height 200 `
           -when "Russia" -then "Sanctions, regulation 'Developing' (0.611)" `
           -body "Yet 16.4% of global Bitcoin hashrate (#2 worldwide) and an Advanced-tier technical infrastructure score (0.584). International isolation has paradoxically deepened the country's crypto presence." `
           -accent $BLUE_MID
Add-VsCard -slide $s -left 490 -top 220 -width 410 -height 200 `
           -when "Venezuela" -then "Ambiguous regulation (0.556), hyperinflation" `
           -body "10.3% crypto ownership, higher than Japan, South Korea, or the UK. USDT is functionally a parallel dollar; restrictive policy has not slowed it." `
           -accent $CORAL
Add-Text -slide $s -text "For policymakers. Channel adoption through clear frameworks rather than attempting to block it. Bans push usage into P2P and self-custody, where regulators have less visibility, not more control." `
         -left 60 -top 440 -width 840 -height 30 -font $FONT_BODY -size 12 -color $INK_BODY -bold $true | Out-Null

# ---------- 19 Finding VI (Federico) ----------
$s = New-BlankSlide -bgColor $BG_TINT
Add-PresenterTag -slide $s -who "FEDERICO BECERRA" -color $TECH_BLUE
Add-Eyebrow -slide $s -text "17  -  Finding VI  -  Market Entry Recommendations" -color $TECH_BLUE
Add-Heading -slide $s -text "Five distinct risk-reward profiles." -top 90 -size 26
Add-Lede -slide $s -text "The split-pillar score lets us group countries by the kind of opportunity (and risk) they actually offer." -top 158
# 5 cards across
$cw = 168; $ch = 240; $startX = 60; $gap = 12
$markets = @(
    @("Low Risk", "Proven Markets", "US, UK, Japan, Brazil", "Legal certainty + strong infrastructure. Best for regulated exchanges and institutional / enterprise products.", (RGB '#2e7d32')),
    @("High Growth", "Developing Regulation", "Ukraine, India, Philippines, Indonesia, Thailand", "Large user bases with regulation moving in the right direction. Ideal for wallets, remittance, DeFi, P2P.", $WARM_MID),
    @("High Risk", "Necessity-Driven", "Vietnam, Nigeria, Pakistan, Argentina, Venezuela", "Massive necessity-driven adoption. Regulatory standing can shift rapidly. High reward, high downside.", $CORAL),
    @("Infrastructure", "Mining / Node Plays", "Russia, Ethiopia, Indonesia", "Mining and node infrastructure opportunities driven by cheap energy, regardless of regulatory environment.", $TECH_BLUE),
    @("Avoid", "No Viable Path", "Bangladesh, Yemen", "Hostile regulation plus minimal infrastructure. No viable entry point currently.", $RED)
)
for ($i = 0; $i -lt 5; $i++) {
    $x = $startX + $i * ($cw + $gap)
    $m = $markets[$i]
    Add-Rect -slide $s -left $x -top 215 -width $cw -height $ch -fill $BG_CARD -cornerRadius 0.05 | Out-Null
    Add-Rect -slide $s -left $x -top 215 -width $cw -height 3 -fill $m[4] -cornerRadius 0 | Out-Null
    Add-Text -slide $s -text $m[0] -left ($x + 12) -top 226 -width ($cw - 24) -height 14 `
             -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -tracking 1.2 | Out-Null
    Add-Text -slide $s -text $m[1] -left ($x + 12) -top 244 -width ($cw - 24) -height 32 `
             -font $FONT_HEAD -size 12 -color $INK -bold $true | Out-Null
    Add-Text -slide $s -text $m[2] -left ($x + 12) -top 280 -width ($cw - 24) -height 50 `
             -font $FONT_BODY -size 9.5 -color $INK_BODY -bold $true | Out-Null
    Add-Text -slide $s -text $m[3] -left ($x + 12) -top 332 -width ($cw - 24) -height 110 `
             -font $FONT_BODY -size 9 -color $INK_MUTED | Out-Null
}

# ---------- 20 Final Conclusion (Federico) ----------
$s = New-BlankSlide -bgColor $BG_LIGHT
Add-PresenterTag -slide $s -who "FEDERICO BECERRA" -color $TECH_BLUE
Add-Eyebrow -slide $s -text "18  -  Final Conclusion  -  Who Should Act"
Add-Heading -slide $s -text "Different audiences, different reads." -top 90 -size 28
Add-Lede -slide $s -text "The split-pillar score is most useful when it changes a decision. Five audiences, five concrete reads." -top 158
$audiences = @(
    @("Business Executives", "Identify entry markets and allocate resources with legal certainty.", "Use the risk-reward profiles to match the product to the country's underlying problem.", $CORAL),
    @("Lawyers & Compliance", "Pinpoint legal risk by jurisdiction.", "Sub-dimension breakdown (stablecoin, licensing, tax, DeFi) makes risk legible at a granular level.", $BLUE_MID),
    @("Policymakers", "Pursue balance, not restriction.", "Ukraine and Brazil prove comprehensive regulation and high adoption can coexist. Bans push usage out of view.", $WARM),
    @("Investors & VCs", "Read the trajectory, not just the level.", "Adoption + improving regulation signals the highest-potential markets. Infrastructure scores reveal where the network is being built.", $TECH_BLUE),
    @("Blockchain Developers", "Where to build infrastructure next.", "Node and validator concentration data guides network expansion, validator location selection, and underserved-region targeting.", $GREEN)
)
for ($i = 0; $i -lt 5; $i++) {
    $x = 60 + $i * 172
    $a = $audiences[$i]
    Add-Rect -slide $s -left $x -top 215 -width 160 -height 230 -fill $BG_CARD -cornerRadius 0.05 | Out-Null
    Add-Rect -slide $s -left $x -top 215 -width 160 -height 4 -fill $a[3] -cornerRadius 0 | Out-Null
    Add-Text -slide $s -text $a[0] -left ($x + 12) -top 230 -width 136 -height 14 `
             -font $FONT_MONO -size 9 -color $INK_DIM -upper $true -tracking 1.2 | Out-Null
    Add-Text -slide $s -text $a[1] -left ($x + 12) -top 248 -width 136 -height 60 `
             -font $FONT_HEAD -size 12 -color $INK -bold $true | Out-Null
    Add-Text -slide $s -text $a[2] -left ($x + 12) -top 314 -width 136 -height 130 `
             -font $FONT_BODY -size 9.5 -color $INK_MUTED | Out-Null
}
Add-Text -slide $s -text "One takeaway. No single number captures blockchain readiness. The split-pillar approach makes the underlying driver of every country's position legible to whoever needs to act on it." `
         -left 60 -top 462 -width 840 -height 32 -font $FONT_BODY -size 11.5 -color $INK_BODY -bold $true | Out-Null

# ---------- 21 Q&A ----------
$s = New-BlankSlide -bgColor $BG_DARK
Add-Text -slide $s -text "FIN" -left 60 -top 130 -width 200 -height 22 `
         -font $FONT_MONO -size 14 -color $TEAL_LIGHT -tracking 4 -bold $true | Out-Null
Add-Text -slide $s -text "Questions?" -left 60 -top 180 -width 840 -height 130 `
         -font $FONT_HEAD -size 96 -color $INK_DARK_BG -bold $true | Out-Null
Add-Text -slide $s -text "Blockchain Adoption Atlas  -  TAC 456 Capstone  -  April 2026" -left 60 -top 360 -width 840 -height 22 `
         -font $FONT_BODY -size 14 -color $INK_DARK_BODY | Out-Null
Add-Text -slide $s -text "Federico Becerra  -  Nico Taylor  -  Jaden Choi  -  Alexis Vo  -  Alon Mutter  -  Ari Choe" -left 60 -top 392 -width 840 -height 22 `
         -font $FONT_MONO -size 11 -color (RGB '#7e8c98') -tracking 1.6 | Out-Null

# ---------- save ----------
$out = (Join-Path (Get-Location) "presentation.pptx")
if (Test-Path $out) { Remove-Item $out -Force }
$pres.SaveAs($out, 24)  # 24 = ppSaveAsOpenXMLPresentation
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Host ""
Write-Host "Wrote: $out"
