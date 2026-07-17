param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputPath
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

function Get-ShapeRecords {
    param($Shapes, [string]$ParentPath = '')
    $records = [Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        $path = if ($ParentPath) { "$ParentPath/$($shape.Name)" } else { [string]$shape.Name }
        $text = ''
        $fontName = $null
        $fontSize = $null
        $fontColor = $null
        if ((Safe { $shape.HasTextFrame }) -eq $msoTrue -and (Safe { $shape.TextFrame.HasText }) -eq $msoTrue) {
            $range = $shape.TextFrame.TextRange
            $text = (($range.Text -replace "`r", ' | ' -replace "`n", ' | ').Trim())
            $fontName = Safe { [string]$range.Font.Name }
            $fontSize = Safe { [double]$range.Font.Size }
            $fontColor = Convert-OfficeRgbToHex (Safe { $range.Font.Color.RGB })
        }
        $records.Add([ordered]@{
            path = $path
            name = [string]$shape.Name
            type = Safe { [int]$shape.Type }
            left = [math]::Round([double]$shape.Left, 2)
            top = [math]::Round([double]$shape.Top, 2)
            width = [math]::Round([double]$shape.Width, 2)
            height = [math]::Round([double]$shape.Height, 2)
            rotation = Safe { [math]::Round([double]$shape.Rotation, 2) }
            text = $text
            font_name = $fontName
            font_size = $fontSize
            font_color = $fontColor
            has_table = (Safe { $shape.HasTable }) -eq $msoTrue
            has_chart = (Safe { $shape.HasChart }) -eq $msoTrue
            z_order = Safe { [int]$shape.ZOrderPosition }
        })
        if ((Safe { [int]$shape.Type }) -eq 6) {
            foreach ($child in (Get-ShapeRecords -Shapes $shape.GroupItems -ParentPath $path)) { $records.Add($child) }
        }
    }
    return $records
}

$resolvedInput = [IO.Path]::GetFullPath($InputPath)
if (-not (Test-Path -LiteralPath $resolvedInput)) { throw "Input '$resolvedInput' does not exist." }

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $app.Presentations.Open($resolvedInput, $msoTrue, 0, 0)
    $slides = [Collections.Generic.List[object]]::new()
    foreach ($slide in $presentation.Slides) {
        $titleCandidates = @()
        $shapes = @(Get-ShapeRecords -Shapes $slide.Shapes)
        foreach ($shape in $shapes) {
            if ($shape.text -and $shape.font_size -ge 24) { $titleCandidates += $shape.text }
        }
        $slides.Add([ordered]@{
            index = [int]$slide.SlideIndex
            slide_id = [int]$slide.SlideID
            layout_name = Safe { [string]$slide.CustomLayout.Name }
            dongpeng_layout = Safe { [string]$slide.Tags.Item('dongpeng_layout') }
            dongpeng_theme = Safe { [string]$slide.Tags.Item('dongpeng_theme') }
            title_candidates = $titleCandidates
            shapes = $shapes
        })
    }
    $result = [ordered]@{
        input = $resolvedInput
        slide_count = [int]$presentation.Slides.Count
        width_pt = [double]$presentation.PageSetup.SlideWidth
        height_pt = [double]$presentation.PageSetup.SlideHeight
        slides = $slides
    }
    $json = $result | ConvertTo-Json -Depth 12
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
