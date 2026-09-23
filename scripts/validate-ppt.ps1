param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputPath,
    [string]$PreviewDirectory
)

$ErrorActionPreference = 'Stop'
$msoTrue = -1

function Safe {
    param([scriptblock]$Expression)
    try { & $Expression } catch { $null }
}

function Convert-OfficeRgbToHex {
    param($Rgb)
    if ($null -eq $Rgb) { return $null }
    $value = [long]$Rgb
    if ($value -lt 0) { return 'MIXED' }
    return ('#{0:X2}{1:X2}{2:X2}' -f ($value -band 0xFF), (($value -shr 8) -band 0xFF), (($value -shr 16) -band 0xFF))
}

function Get-RedFillRegions {
    param($Shapes, [string]$ParentPath = '')
    $redColors = @('#E21413', '#B82020', '#EE7271', '#F98A8A', '#F3A1A1', '#F9D0D0', '#FADBDF', '#F0B6BA')
    $regions = [Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        $path = if ($ParentPath) { "$ParentPath/$($shape.Name)" } else { [string]$shape.Name }
        $fillVisible = Safe { [int]$shape.Fill.Visible }
        $fillColor = Convert-OfficeRgbToHex (Safe { $shape.Fill.ForeColor.RGB })
        $fillTransparency = Safe { [double]$shape.Fill.Transparency }
        if ($fillVisible -eq -1 -and $fillColor -in $redColors -and ($null -eq $fillTransparency -or $fillTransparency -lt 1)) {
            $regions.Add([pscustomobject]@{
                left = [double]$shape.Left
                top = [double]$shape.Top
                right = [double]$shape.Left + [double]$shape.Width
                bottom = [double]$shape.Top + [double]$shape.Height
                shape = $path
            })
        }
        if ((Safe { [int]$shape.Type }) -eq 6) {
            foreach ($region in (Get-RedFillRegions -Shapes $shape.GroupItems -ParentPath $path)) { $regions.Add($region) }
        }
    }
    return $regions.ToArray()
}

function Visit-Shapes {
    param($Shapes, [int]$SlideIndex, [Collections.Generic.List[object]]$Warnings, [string]$LayoutId, [object[]]$RedRegions, [string]$ParentPath = '')
    $yahei = -join @([char]0x5FAE, [char]0x8F6F, [char]0x96C5, [char]0x9ED1)
    $allowedColors = @('#E21413', '#000000', '#0D0D0D', '#FFFFFF', '#F9D0D0', '#F3A1A1', '#EE7271', '#F98A8A', '#FADBDF', '#F0B6BA', '#B82020', '#F2F2F2', '#BFBFBF', '#D9D9D9', 'MIXED')

    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        $path = if ($ParentPath) { "$ParentPath/$($shape.Name)" } else { [string]$shape.Name }
        $left = [double]$shape.Left; $top = [double]$shape.Top; $width = [double]$shape.Width; $height = [double]$shape.Height
        if ($left -lt -2 -or $top -lt -2 -or ($left + $width) -gt 962 -or ($top + $height) -gt 542) {
            $Warnings.Add([ordered]@{ code = 'shape_outside_canvas'; slide = $SlideIndex; shape = $path; message = 'Shape extends outside the 960 × 540 pt canvas.' })
        }

        if ((Safe { $shape.HasTextFrame }) -eq $msoTrue -and (Safe { $shape.TextFrame.HasText }) -eq $msoTrue) {
            $range = $shape.TextFrame.TextRange
            $text = (($range.Text -replace "`r", ' ' -replace "`n", ' ').Trim())
            $fontName = Safe { [string]$range.Font.Name }
            $fontSize = Safe { [double]$range.Font.Size }
            $fontColor = Convert-OfficeRgbToHex (Safe { $range.Font.Color.RGB })
            $fontAllowed = $fontName -eq $yahei -or $fontName -match 'YaHei|Arial|^\+'
            if ($fontName -and -not $fontAllowed) {
                $Warnings.Add([ordered]@{ code = 'unexpected_font'; slide = $SlideIndex; shape = $path; message = "Unexpected font '$fontName'." })
            }
            if ($fontSize -gt 0) {
                $minimum = if ($top -lt 50 -and $fontSize -le 10) { 9 } elseif ($text -match 'Copyright') { 8.35 } else { 12 }
                if ($fontSize -lt $minimum) {
                    $Warnings.Add([ordered]@{ code = 'font_too_small'; slide = $SlideIndex; shape = $path; message = "Font size $fontSize pt is below the expected minimum $minimum pt." })
                }
            }
            if ($fontColor -and $fontColor -notin $allowedColors) {
                $Warnings.Add([ordered]@{ code = 'unexpected_text_color'; slide = $SlideIndex; shape = $path; message = "Unexpected text color '$fontColor'." })
            }
            $textArea = [math]::Max(1, $width * $height)
            foreach ($region in $RedRegions) {
                $overlapWidth = [math]::Max(0, [math]::Min(($left + $width), $region.right) - [math]::Max($left, $region.left))
                $overlapHeight = [math]::Max(0, [math]::Min(($top + $height), $region.bottom) - [math]::Max($top, $region.top))
                if (($overlapWidth * $overlapHeight / $textArea) -ge 0.2 -and $fontColor -ne '#FFFFFF') {
                    $Warnings.Add([ordered]@{ code = 'nonwhite_text_on_red'; slide = $SlideIndex; shape = $path; message = "Text color '$fontColor' overlaps red-filled shape '$($region.shape)'; all text on red backgrounds must be white (#FFFFFF)." })
                    break
                }
            }
            $boundHeight = Safe { [double]$shape.TextFrame2.TextRange.BoundHeight }
            $boundWidth = Safe { [double]$shape.TextFrame2.TextRange.BoundWidth }
            if ($boundHeight -and $boundHeight -gt [math]::Max(($height * 1.3), ($height + 8))) {
                $Warnings.Add([ordered]@{ code = 'text_overflow_height'; slide = $SlideIndex; shape = $path; message = "Text bound height $([math]::Round($boundHeight,1)) exceeds shape height $([math]::Round($height,1))." })
            }
            if ((Safe { [int]$shape.TextFrame.WordWrap }) -eq 0 -and $boundWidth -and $boundWidth -gt ($width + 2)) {
                $Warnings.Add([ordered]@{ code = 'text_overflow_width'; slide = $SlideIndex; shape = $path; message = "Text bound width $([math]::Round($boundWidth,1)) exceeds shape width $([math]::Round($width,1))." })
            }
            if (($top + $height) -gt 480 -and $LayoutId -notin @('cover', 'closing') -and $text -notmatch 'Copyright') {
                $Warnings.Add([ordered]@{ code = 'bottom_safe_zone'; slide = $SlideIndex; shape = $path; message = 'Text enters the bottom brand safe zone.' })
            }
        }

        if ((Safe { [int]$shape.Type }) -eq 6) {
            Visit-Shapes -Shapes $shape.GroupItems -SlideIndex $SlideIndex -Warnings $Warnings -LayoutId $LayoutId -RedRegions $RedRegions -ParentPath $path
        }
    }
}

$resolvedInput = [IO.Path]::GetFullPath($InputPath)
if (-not (Test-Path -LiteralPath $resolvedInput)) { throw "Input '$resolvedInput' does not exist." }

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
$redThemeLayouts = @('cover', 'section-red', 'closing', 'swot', 'keyword-three', 'keyword-radial', 'stair-red', 'icon-library-red')
try {
    $presentation = $app.Presentations.Open($resolvedInput, $msoTrue, 0, 0)
    $warnings = [Collections.Generic.List[object]]::new()

    if ([math]::Abs([double]$presentation.PageSetup.SlideWidth - 960) -gt 0.5 -or [math]::Abs([double]$presentation.PageSetup.SlideHeight - 540) -gt 0.5) {
        $warnings.Add([ordered]@{ code = 'wrong_canvas'; slide = $null; shape = $null; message = "Canvas is $($presentation.PageSetup.SlideWidth) × $($presentation.PageSetup.SlideHeight) pt; expected 960 × 540 pt." })
    }
    if ($presentation.Slides.Count -eq 0) {
        $warnings.Add([ordered]@{ code = 'empty_deck'; slide = $null; shape = $null; message = 'Presentation has no slides.' })
    }

    foreach ($slide in $presentation.Slides) {
        $layoutTag = Safe { [string]$slide.Tags.Item('dongpeng_layout') }
        $redRegions = @(Get-RedFillRegions -Shapes $slide.Shapes)
        if ($layoutTag -in $redThemeLayouts) {
            $redRegions += [pscustomobject]@{ left = 0; top = 0; right = 960; bottom = 540; shape = "red-theme layout '$layoutTag'" }
        }
        foreach ($backgroundFill in @((Safe { $slide.Background.Fill }), (Safe { $slide.CustomLayout.Background.Fill }))) {
            $backgroundColor = Convert-OfficeRgbToHex (Safe { $backgroundFill.ForeColor.RGB })
            if ($backgroundColor -in @('#E21413', '#B82020', '#EE7271', '#F98A8A', '#F3A1A1', '#F9D0D0', '#FADBDF', '#F0B6BA')) {
                $redRegions += [pscustomobject]@{ left = 0; top = 0; right = 960; bottom = 540; shape = 'slide background' }
            }
        }
        Visit-Shapes -Shapes $slide.Shapes -SlideIndex ([int]$slide.SlideIndex) -Warnings $warnings -LayoutId $layoutTag -RedRegions $redRegions
        if (-not $layoutTag) {
            $warnings.Add([ordered]@{ code = 'missing_layout_tag'; slide = [int]$slide.SlideIndex; shape = $null; message = 'Slide has no dongpeng_layout tag; use explicit inspection for legacy slides.' })
        }
    }

    if ($PreviewDirectory) {
        $preview = [IO.Path]::GetFullPath($PreviewDirectory)
        New-Item -ItemType Directory -Force -Path $preview | Out-Null
        $presentation.Export($preview, 'PNG', 1600, 900)
    }

    $result = [ordered]@{
        input = $resolvedInput
        valid = $warnings.Count -eq 0
        slide_count = [int]$presentation.Slides.Count
        width_pt = [double]$presentation.PageSetup.SlideWidth
        height_pt = [double]$presentation.PageSetup.SlideHeight
        warning_count = $warnings.Count
        warnings = $warnings
    }
    $json = $result | ConvertTo-Json -Depth 10
    if ($OutputPath) {
        $resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedOutput) | Out-Null
        Set-Content -Encoding UTF8 -LiteralPath $resolvedOutput -Value $json
    }
    $json
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    try { $app.Quit() } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
