param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24
$ppSaveAsPDF = 32
$msoTextOrientationHorizontal = 1

function Has-Property {
    param($Object, [string]$Name)
    return $null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]
}

function Resolve-SpecPath {
    param([string]$Path, [string]$BaseDirectory)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $BaseDirectory $Path))
}

function Clear-ReadOnlyFile {
    param([string]$Path)
    $item = Get-Item -LiteralPath $Path
    if ($item.IsReadOnly) { $item.IsReadOnly = $false }
}

function Convert-HexToOfficeRgb {
    param([string]$Hex)
    $value = $Hex.Trim().TrimStart('#')
    if ($value.Length -ne 6) { throw "Invalid color '$Hex'. Use #RRGGBB." }
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $r + ($g -shl 8) + ($b -shl 16)
}

function Get-ShapeEntries {
    param($Shapes, [string]$ParentPath = '')
    $entries = [Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        $path = if ($ParentPath) { "$ParentPath/$($shape.Name)" } else { [string]$shape.Name }
        $entries.Add([pscustomobject]@{ Shape = $shape; Path = $path })
        if ([int]$shape.Type -eq 6) {
            $children = Get-ShapeEntries -Shapes $shape.GroupItems -ParentPath $path
            foreach ($child in $children) { $entries.Add($child) }
        }
    }
    return $entries
}

function Get-ShapeText {
    param($Shape)
    try {
        if ($Shape.HasTextFrame -eq $msoTrue -and $Shape.TextFrame.HasText -eq $msoTrue) {
            return [string]$Shape.TextFrame.TextRange.Text
        }
    } catch {}
    return ''
}

function Get-SlideLayoutTag {
    param($Slide)
    try { return [string]$Slide.Tags.Item('dongpeng_layout') } catch { return '' }
}

function Find-SlideByOperation {
    param($Presentation, $Operation)

    if (Has-Property $Operation 'slide_id') {
        foreach ($slide in $Presentation.Slides) {
            if ([int]$slide.SlideID -eq [int]$Operation.slide_id) { return $slide }
        }
        throw "Slide ID $($Operation.slide_id) was not found."
    }

    if (Has-Property $Operation 'slide_layout') {
        $occurrence = if (Has-Property $Operation 'occurrence') { [int]$Operation.occurrence } else { 1 }
        $found = 0
        foreach ($slide in $Presentation.Slides) {
            if ((Get-SlideLayoutTag $slide) -eq [string]$Operation.slide_layout) {
                $found++
                if ($found -eq $occurrence) { return $slide }
            }
        }
        throw "Slide layout '$($Operation.slide_layout)' occurrence $occurrence was not found."
    }

    if (-not (Has-Property $Operation 'slide')) { throw 'The operation requires slide, slide_id, or slide_layout.' }
    $index = [int]$Operation.slide
    if ($index -lt 1 -or $index -gt $Presentation.Slides.Count) { throw "Slide index $index is out of range." }
    return $Presentation.Slides.Item($index)
}

function Get-LayoutDefinition {
    param($LayoutMap, [string]$LayoutId)
    $property = $LayoutMap.layouts.PSObject.Properties[$LayoutId]
    if ($null -eq $property) { throw "Unknown layout_id '$LayoutId'." }
    return $property.Value
}

function Get-ShapeSelector {
    param($Slide, $Operation, $LayoutMap)

    if (Has-Property $Operation 'selector') { return $Operation.selector }
    if (-not (Has-Property $Operation 'role')) { throw 'The operation requires selector or role.' }

    $layoutId = if (Has-Property $Operation 'layout_id') { [string]$Operation.layout_id } else { Get-SlideLayoutTag $Slide }
    if ([string]::IsNullOrWhiteSpace($layoutId)) {
        throw "Slide $($Slide.SlideIndex) has no dongpeng_layout tag. Provide layout_id or an explicit selector."
    }
    $layout = Get-LayoutDefinition -LayoutMap $LayoutMap -LayoutId $layoutId
    $roleProperty = $layout.roles.PSObject.Properties[[string]$Operation.role]
    if ($null -eq $roleProperty) { throw "Role '$($Operation.role)' is not registered for layout '$layoutId'." }
    return $roleProperty.Value
}

function Find-Shape {
    param($Slide, $Selector, [switch]$RequireText)

    $tolerance = if (Has-Property $Selector 'tolerance') { [double]$Selector.tolerance } else { 3.0 }
    $matches = [Collections.Generic.List[object]]::new()
    foreach ($entry in (Get-ShapeEntries -Shapes $Slide.Shapes)) {
        $shape = $entry.Shape
        if ($RequireText -and $shape.HasTextFrame -ne $msoTrue) { continue }
        if (Has-Property $Selector 'name') {
            if (-not ([string]$shape.Name).Equals([string]$Selector.name, [StringComparison]::OrdinalIgnoreCase)) { continue }
        }
        if (Has-Property $Selector 'path') {
            if (-not ([string]$entry.Path).Equals([string]$Selector.path, [StringComparison]::OrdinalIgnoreCase)) { continue }
        }
        if (Has-Property $Selector 'type') {
            if ([int]$shape.Type -ne [int]$Selector.type) { continue }
        }
        if (Has-Property $Selector 'left') {
            if ([math]::Abs([double]$shape.Left - [double]$Selector.left) -gt $tolerance) { continue }
        }
        if (Has-Property $Selector 'top') {
            if ([math]::Abs([double]$shape.Top - [double]$Selector.top) -gt $tolerance) { continue }
        }
        if (Has-Property $Selector 'text_contains') {
            if ((Get-ShapeText $shape).IndexOf([string]$Selector.text_contains, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }
        $matches.Add($entry)
    }

    if ($matches.Count -eq 0) {
        throw "No shape matched selector: $($Selector | ConvertTo-Json -Compress -Depth 5)"
    }
    $matchIndex = if (Has-Property $Selector 'match_index') { [int]$Selector.match_index } else { 1 }
    if ($matchIndex -lt 1 -or $matchIndex -gt $matches.Count) {
        throw "Selector matched $($matches.Count) shapes; match_index $matchIndex is invalid."
    }
    return $matches[$matchIndex - 1].Shape
}

function Apply-Geometry {
    param($Shape, $Object)
    foreach ($name in @('left', 'top', 'width', 'height', 'rotation')) {
        if (Has-Property $Object $name) {
            $propertyName = (Get-Culture).TextInfo.ToTitleCase($name)
            $Shape.$propertyName = [double]$Object.$name
        }
    }
}

function Apply-Font {
    param($TextRange, $Font)
    if ($null -eq $Font) { return }
    if (Has-Property $Font 'name') { $TextRange.Font.Name = [string]$Font.name }
    if (Has-Property $Font 'size') { $TextRange.Font.Size = [double]$Font.size }
    if (Has-Property $Font 'bold') { $TextRange.Font.Bold = if ([bool]$Font.bold) { $msoTrue } else { $msoFalse } }
    if (Has-Property $Font 'italic') { $TextRange.Font.Italic = if ([bool]$Font.italic) { $msoTrue } else { $msoFalse } }
    if (Has-Property $Font 'color') { $TextRange.Font.Color.RGB = Convert-HexToOfficeRgb ([string]$Font.color) }
}

function Apply-Paragraph {
    param($TextRange, $Paragraph)
    if ($null -eq $Paragraph) { return }
    if (Has-Property $Paragraph 'alignment') {
        $map = @{ left = 1; center = 2; right = 3; justify = 4; distribute = 5 }
        $key = ([string]$Paragraph.alignment).ToLowerInvariant()
        if (-not $map.ContainsKey($key)) { throw "Unknown paragraph alignment '$key'." }
        $TextRange.ParagraphFormat.Alignment = $map[$key]
    }
    if (Has-Property $Paragraph 'line_spacing') { $TextRange.ParagraphFormat.SpaceWithin = [double]$Paragraph.line_spacing }
    if (Has-Property $Paragraph 'space_before') { $TextRange.ParagraphFormat.SpaceBefore = [double]$Paragraph.space_before }
    if (Has-Property $Paragraph 'space_after') { $TextRange.ParagraphFormat.SpaceAfter = [double]$Paragraph.space_after }
}

function Move-ShapeToZPosition {
    param($Shape, [int]$TargetPosition)
    try {
        while ([int]$Shape.ZOrderPosition -gt $TargetPosition) { $Shape.ZOrder(3) }
        while ([int]$Shape.ZOrderPosition -lt $TargetPosition) { $Shape.ZOrder(2) }
    } catch {}
}

function Set-PictureFit {
    param($Shape, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Fit = 'cover', [double]$FocusX = 0.5, [double]$FocusY = 0.5)

    $fitMode = $Fit.ToLowerInvariant()
    $originalWidth = [double]$Shape.Width
    $originalHeight = [double]$Shape.Height
    if ($originalWidth -le 0 -or $originalHeight -le 0) { throw 'The picture has invalid dimensions.' }

    if ($fitMode -eq 'stretch') {
        $Shape.LockAspectRatio = $msoFalse
        $Shape.Left = $Left; $Shape.Top = $Top; $Shape.Width = $Width; $Shape.Height = $Height
        return
    }

    if ($fitMode -eq 'contain') {
        $scale = [math]::Min($Width / $originalWidth, $Height / $originalHeight)
        $newWidth = $originalWidth * $scale
        $newHeight = $originalHeight * $scale
        $Shape.LockAspectRatio = $msoFalse
        $Shape.Width = $newWidth; $Shape.Height = $newHeight
        $Shape.Left = $Left + (($Width - $newWidth) * $FocusX)
        $Shape.Top = $Top + (($Height - $newHeight) * $FocusY)
        $Shape.LockAspectRatio = $msoTrue
        return
    }

    if ($fitMode -ne 'cover') { throw "Unknown picture fit '$Fit'." }

    $scale = [math]::Max($Width / $originalWidth, $Height / $originalHeight)
    $pictureWidth = $originalWidth * $scale
    $pictureHeight = $originalHeight * $scale
    try {
        $crop = $Shape.PictureFormat.Crop
        $crop.ShapeLeft = $Left
        $crop.ShapeTop = $Top
        $crop.ShapeWidth = $Width
        $crop.ShapeHeight = $Height
        $crop.PictureWidth = $pictureWidth
        $crop.PictureHeight = $pictureHeight
        $crop.PictureOffsetX = ($Width - $pictureWidth) * ($FocusX - 0.5)
        $crop.PictureOffsetY = ($Height - $pictureHeight) * ($FocusY - 0.5)
    } catch {
        throw "PowerPoint could not apply a non-distorting cover crop: $($_.Exception.Message)"
    }
}

function Add-PictureToBox {
    param($Slide, [string]$Path, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Fit, [double]$FocusX, [double]$FocusY)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Image '$Path' does not exist." }
    $shape = $Slide.Shapes.AddPicture($Path, $msoFalse, $msoTrue, 0, 0, -1, -1)
    Set-PictureFit -Shape $shape -Left $Left -Top $Top -Width $Width -Height $Height -Fit $Fit -FocusX $FocusX -FocusY $FocusY
    return $shape
}

function Set-ShapeStyle {
    param($Shape, $Operation)
    Apply-Geometry -Shape $Shape -Object $Operation
    if (Has-Property $Operation 'fill') {
        $fill = $Operation.fill
        if (Has-Property $fill 'visible') { $Shape.Fill.Visible = if ([bool]$fill.visible) { $msoTrue } else { $msoFalse } }
        if (Has-Property $fill 'color') { $Shape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$fill.color) }
        if (Has-Property $fill 'transparency') { $Shape.Fill.Transparency = [double]$fill.transparency }
    }
    if (Has-Property $Operation 'line') {
        $line = $Operation.line
        if (Has-Property $line 'visible') { $Shape.Line.Visible = if ([bool]$line.visible) { $msoTrue } else { $msoFalse } }
        if (Has-Property $line 'color') { $Shape.Line.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$line.color) }
        if (Has-Property $line 'weight') { $Shape.Line.Weight = [double]$line.weight }
        if (Has-Property $line 'transparency') { $Shape.Line.Transparency = [double]$line.transparency }
    }
    if (Has-Property $Operation 'z_order') {
        $zMap = @{ front = 0; back = 1; forward = 2; backward = 3 }
        $key = ([string]$Operation.z_order).ToLowerInvariant()
        if (-not $zMap.ContainsKey($key)) { throw "Unknown z_order '$key'." }
        $Shape.ZOrder($zMap[$key])
    }
}

$resolvedSpecPath = [IO.Path]::GetFullPath($SpecPath)
$specDirectory = Split-Path -Parent $resolvedSpecPath
$skillRoot = Split-Path -Parent $PSScriptRoot
$layoutMap = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $skillRoot 'assets\layout-map.json') | ConvertFrom-Json
$spec = Get-Content -Raw -Encoding UTF8 -LiteralPath $resolvedSpecPath | ConvertFrom-Json

if (-not (Has-Property $spec 'output')) { throw 'The specification requires output.' }
$outputPath = Resolve-SpecPath -Path ([string]$spec.output) -BaseDirectory $specDirectory
$outputDirectory = Split-Path -Parent $outputPath
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$overwrite = (Has-Property $spec 'overwrite') -and [bool]$spec.overwrite

if ((Test-Path -LiteralPath $outputPath) -and -not $overwrite) {
    throw "Output '$outputPath' already exists. Set overwrite to true only with explicit user authorization."
}

$templatePath = if (Has-Property $spec 'template') {
    Resolve-SpecPath -Path ([string]$spec.template) -BaseDirectory $specDirectory
} else {
    Join-Path (Join-Path $skillRoot 'assets') ([string]$layoutMap.template_file)
}

$inputPath = if (Has-Property $spec 'input') { Resolve-SpecPath -Path ([string]$spec.input) -BaseDirectory $specDirectory } else { $null }
$newDeck = (Has-Property $spec 'new_deck') -and [bool]$spec.new_deck
if (-not $newDeck -and [string]::IsNullOrWhiteSpace($inputPath)) { throw 'Provide input or set new_deck to true.' }
if ($inputPath -and -not (Test-Path -LiteralPath $inputPath)) { throw "Input '$inputPath' does not exist." }
if (-not (Test-Path -LiteralPath $templatePath)) { throw "Template '$templatePath' does not exist." }

$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
$templatePresentation = $null
$result = [ordered]@{ output = $outputPath; operations = @(); pdf = $null; preview_directory = $null }

try {
    if ($newDeck) {
        Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force
        Clear-ReadOnlyFile -Path $outputPath
        $presentation = $app.Presentations.Open($outputPath, $msoFalse, $msoFalse, $msoFalse)
        while ($presentation.Slides.Count -gt 0) { $presentation.Slides.Item($presentation.Slides.Count).Delete() }
    } else {
        if ([IO.Path]::GetFullPath($inputPath).Equals($outputPath, [StringComparison]::OrdinalIgnoreCase)) {
            if (-not $overwrite) { throw 'Editing in place requires overwrite=true and explicit user authorization.' }
            $presentation = $app.Presentations.Open($inputPath, $msoFalse, $msoFalse, $msoFalse)
        } else {
            Copy-Item -LiteralPath $inputPath -Destination $outputPath -Force
            Clear-ReadOnlyFile -Path $outputPath
            $presentation = $app.Presentations.Open($outputPath, $msoFalse, $msoFalse, $msoFalse)
        }
    }

    $templatePresentation = $app.Presentations.Open($templatePath, $msoTrue, $msoFalse, $msoFalse)

    $operationIndex = 0
    foreach ($operation in $spec.operations) {
        $operationIndex++
        $opName = ([string]$operation.op).ToLowerInvariant()
        try {
            switch ($opName) {
                'insert_template_slide' {
                    $layoutId = [string]$operation.layout_id
                    $layout = Get-LayoutDefinition -LayoutMap $layoutMap -LayoutId $layoutId
                    $sourceSlide = $templatePresentation.Slides.Item([int]$layout.template_slide)
                    $sourceSlide.Copy()
                    $pastedRange = $presentation.Slides.Paste()
                    $newSlide = $pastedRange.Item(1)
                    $position = if (Has-Property $operation 'position') { [int]$operation.position } else { $presentation.Slides.Count }
                    if ($position -lt 1) { $position = 1 }
                    if ($position -gt $presentation.Slides.Count) { $position = $presentation.Slides.Count }
                    $newSlide.MoveTo($position)
                    $newSlide.Tags.Add('dongpeng_layout', $layoutId)
                    $newSlide.Tags.Add('dongpeng_theme', [string]$layout.theme)
                }
                'duplicate_slide' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $duplicate = $slide.Duplicate().Item(1)
                    $position = if (Has-Property $operation 'position') { [int]$operation.position } else { $slide.SlideIndex + 1 }
                    $duplicate.MoveTo($position)
                }
                'delete_slide' {
                    (Find-SlideByOperation -Presentation $presentation -Operation $operation).Delete()
                }
                'move_slide' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $slide.MoveTo([int]$operation.to)
                }
                'set_slide_tag' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $slide.Tags.Add([string]$operation.name, [string]$operation.value)
                }
                'set_text' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector -RequireText
                    if ($shape.HasTextFrame -ne $msoTrue) { throw "Shape '$($shape.Name)' does not support text." }
                    if (Has-Property $operation 'text') { $shape.TextFrame.TextRange.Text = [string]$operation.text }
                    $range = $shape.TextFrame.TextRange
                    if (Has-Property $operation 'font') { Apply-Font -TextRange $range -Font $operation.font }
                    if (Has-Property $operation 'paragraph') { Apply-Paragraph -TextRange $range -Paragraph $operation.paragraph }
                    if (Has-Property $operation 'margins') {
                        $m = $operation.margins
                        if (Has-Property $m 'left') { $shape.TextFrame.MarginLeft = [double]$m.left }
                        if (Has-Property $m 'right') { $shape.TextFrame.MarginRight = [double]$m.right }
                        if (Has-Property $m 'top') { $shape.TextFrame.MarginTop = [double]$m.top }
                        if (Has-Property $m 'bottom') { $shape.TextFrame.MarginBottom = [double]$m.bottom }
                    }
                    Apply-Geometry -Shape $shape -Object $operation
                }
                'add_textbox' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
                    $shape.TextFrame.TextRange.Text = [string]$operation.text
                    if (Has-Property $operation 'name') { $shape.Name = [string]$operation.name }
                    if (Has-Property $operation 'font') { Apply-Font -TextRange $shape.TextFrame.TextRange -Font $operation.font }
                    if (Has-Property $operation 'paragraph') { Apply-Paragraph -TextRange $shape.TextFrame.TextRange -Paragraph $operation.paragraph }
                }
                'insert_image' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $imagePath = Resolve-SpecPath -Path ([string]$operation.path) -BaseDirectory $specDirectory
                    $fit = if (Has-Property $operation 'fit') { [string]$operation.fit } else { 'cover' }
                    $focusX = if (Has-Property $operation 'focus_x') { [double]$operation.focus_x } else { 0.5 }
                    $focusY = if (Has-Property $operation 'focus_y') { [double]$operation.focus_y } else { 0.5 }
                    $shape = Add-PictureToBox -Slide $slide -Path $imagePath -Left ([double]$operation.left) -Top ([double]$operation.top) -Width ([double]$operation.width) -Height ([double]$operation.height) -Fit $fit -FocusX $focusX -FocusY $focusY
                    if (Has-Property $operation 'name') { $shape.Name = [string]$operation.name }
                    if (Has-Property $operation 'z_order') { Set-ShapeStyle -Shape $shape -Operation $operation }
                }
                'replace_image' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $oldShape = Find-Shape -Slide $slide -Selector $selector
                    $imagePath = Resolve-SpecPath -Path ([string]$operation.path) -BaseDirectory $specDirectory
                    $left = [double]$oldShape.Left; $top = [double]$oldShape.Top; $width = [double]$oldShape.Width; $height = [double]$oldShape.Height
                    $name = [string]$oldShape.Name; $z = [int]$oldShape.ZOrderPosition
                    $oldShape.Delete()
                    $fit = if (Has-Property $operation 'fit') { [string]$operation.fit } else { 'cover' }
                    $focusX = if (Has-Property $operation 'focus_x') { [double]$operation.focus_x } else { 0.5 }
                    $focusY = if (Has-Property $operation 'focus_y') { [double]$operation.focus_y } else { 0.5 }
                    $newShape = Add-PictureToBox -Slide $slide -Path $imagePath -Left $left -Top $top -Width $width -Height $height -Fit $fit -FocusX $focusX -FocusY $focusY
                    try { $newShape.Name = $name } catch {}
                    Move-ShapeToZPosition -Shape $newShape -TargetPosition $z
                }
                'copy_shape_from_template' {
                    $targetSlide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $sourceSlide = $templatePresentation.Slides.Item([int]$operation.template_slide)
                    $sourceShape = Find-Shape -Slide $sourceSlide -Selector $operation.selector
                    $sourceShape.Copy()
                    $newShape = $targetSlide.Shapes.Paste().Item(1)
                    Apply-Geometry -Shape $newShape -Object $operation
                }
                'set_shape' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    Set-ShapeStyle -Shape (Find-Shape -Slide $slide -Selector $selector) -Operation $operation
                }
                'delete_shape' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    (Find-Shape -Slide $slide -Selector $selector).Delete()
                }
                'duplicate_shape' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = (Find-Shape -Slide $slide -Selector $selector).Duplicate().Item(1)
                    Apply-Geometry -Shape $shape -Object $operation
                    if (Has-Property $operation 'name') { $shape.Name = [string]$operation.name }
                }
                'add_shape' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $typeMap = @{ rectangle = 1; rounded_rectangle = 5; oval = 9; triangle = 7; hexagon = 10 }
                    $shapeType = if (Has-Property $operation 'shape_type' -and $operation.shape_type -is [string]) { $typeMap[([string]$operation.shape_type).ToLowerInvariant()] } else { [int]$operation.shape_type }
                    if ($null -eq $shapeType) { throw "Unknown shape_type '$($operation.shape_type)'." }
                    $shape = $slide.Shapes.AddShape($shapeType, [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
                    if (Has-Property $operation 'name') { $shape.Name = [string]$operation.name }
                    Set-ShapeStyle -Shape $shape -Operation $operation
                    if (Has-Property $operation 'text') {
                        $shape.TextFrame.TextRange.Text = [string]$operation.text
                        if (Has-Property $operation 'font') { Apply-Font -TextRange $shape.TextFrame.TextRange -Font $operation.font }
                        if (Has-Property $operation 'paragraph') { Apply-Paragraph -TextRange $shape.TextFrame.TextRange -Paragraph $operation.paragraph }
                    }
                }
                'set_table_cell' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    if ($shape.HasTable -ne $msoTrue) { throw "Shape '$($shape.Name)' is not a table." }
                    $cellShape = $shape.Table.Cell([int]$operation.row, [int]$operation.column).Shape
                    if (Has-Property $operation 'text') { $cellShape.TextFrame.TextRange.Text = [string]$operation.text }
                    if (Has-Property $operation 'font') { Apply-Font -TextRange $cellShape.TextFrame.TextRange -Font $operation.font }
                    if (Has-Property $operation 'fill') {
                        if (Has-Property $operation.fill 'color') { $cellShape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb ([string]$operation.fill.color) }
                    }
                }
                'add_table_row' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    $null = $shape.Table.Rows.Add()
                }
                'delete_table_row' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    $shape.Table.Rows.Item([int]$operation.row).Delete()
                }
                'add_table_column' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    $null = $shape.Table.Columns.Add()
                }
                'delete_table_column' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    $shape.Table.Columns.Item([int]$operation.column).Delete()
                }
                'set_chart_data' {
                    $slide = Find-SlideByOperation -Presentation $presentation -Operation $operation
                    $selector = Get-ShapeSelector -Slide $slide -Operation $operation -LayoutMap $layoutMap
                    $shape = Find-Shape -Slide $slide -Selector $selector
                    if ($shape.HasChart -ne $msoTrue) { throw "Shape '$($shape.Name)' is not a chart." }
                    $categories = @($operation.categories)
                    $series = @($operation.series)
                    if ($categories.Count -eq 0 -or $series.Count -eq 0) { throw 'set_chart_data requires categories and series.' }
                    foreach ($item in $series) {
                        if (@($item.values).Count -ne $categories.Count) { throw "Series '$($item.name)' value count must match categories." }
                    }
                    $chart = $shape.Chart
                    $chart.ChartData.Activate()
                    $workbook = $chart.ChartData.Workbook
                    $worksheet = $workbook.Worksheets.Item(1)
                    $null = $worksheet.Cells.Clear()
                    $worksheet.Cells.Item(1, 1).Value2 = ''
                    for ($column = 0; $column -lt $series.Count; $column++) {
                        $worksheet.Cells.Item(1, $column + 2).Value2 = [string]$series[$column].name
                    }
                    for ($row = 0; $row -lt $categories.Count; $row++) {
                        $worksheet.Cells.Item($row + 2, 1).Value2 = [string]$categories[$row]
                        for ($column = 0; $column -lt $series.Count; $column++) {
                            $worksheet.Cells.Item($row + 2, $column + 2).Value2 = [double]@($series[$column].values)[$row]
                        }
                    }
                    $range = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($categories.Count + 1, $series.Count + 1))
                    $sourceAddress = $range.Address($msoTrue, $msoTrue, 1, $msoTrue)
                    $null = $chart.SetSourceData($sourceAddress)
                    if (Has-Property $operation 'data_labels') {
                        $labelMode = ([string]$operation.data_labels).ToLowerInvariant()
                        for ($seriesIndex = 0; $seriesIndex -lt $series.Count; $seriesIndex++) {
                            $chartSeries = $chart.SeriesCollection($seriesIndex + 1)
                            if ($labelMode -eq 'none') {
                                try { $chartSeries.DataLabels().Delete() } catch {}
                                continue
                            }
                            try { $labels = $chartSeries.DataLabels() } catch {
                                $chartSeries.ApplyDataLabels()
                                $labels = $chartSeries.DataLabels()
                            }
                            $values = @($series[$seriesIndex].values)
                            $total = 0.0
                            foreach ($value in $values) { $total += [double]$value }
                            for ($pointIndex = 0; $pointIndex -lt $values.Count; $pointIndex++) {
                                if ($labelMode -eq 'percentage') {
                                    $labelText = if ($total -eq 0) { '0%' } else { ('{0:0.#}%' -f (([double]$values[$pointIndex] / $total) * 100)) }
                                } elseif ($labelMode -eq 'value') {
                                    $labelText = [string]$values[$pointIndex]
                                } else {
                                    throw "Unknown data_labels mode '$labelMode'."
                                }
                                $labels.Item($pointIndex + 1).Text = $labelText
                            }
                        }
                    }
                    if (Has-Property $operation 'title') {
                        $chart.HasTitle = $msoTrue
                        $chart.ChartTitle.Text = [string]$operation.title
                    }
                    $null = $chart.Refresh()
                    $null = $workbook.Close($msoTrue)
                }
                'export_pdf' {
                    $result.pdf = Resolve-SpecPath -Path ([string]$operation.path) -BaseDirectory $specDirectory
                }
                'export_png' {
                    $result.preview_directory = Resolve-SpecPath -Path ([string]$operation.directory) -BaseDirectory $specDirectory
                    $result.preview_width = if (Has-Property $operation 'width') { [int]$operation.width } else { 1600 }
                    $result.preview_height = if (Has-Property $operation 'height') { [int]$operation.height } else { 900 }
                }
                'save' { }
                default { throw "Unsupported operation '$opName'." }
            }
            $result.operations += [ordered]@{ index = $operationIndex; op = $opName; status = 'ok' }
        } catch {
            throw "Operation $operationIndex '$opName' failed: $($_.Exception.Message)"
        }
    }

    if ($presentation.Slides.Count -eq 0) { throw 'The output presentation has no slides.' }
    $presentation.SaveAs($outputPath, $ppSaveAsOpenXMLPresentation)

    if ($result.pdf) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $result.pdf) | Out-Null
        $presentation.SaveAs($result.pdf, $ppSaveAsPDF)
    }
    if ($result.preview_directory) {
        New-Item -ItemType Directory -Force -Path $result.preview_directory | Out-Null
        $presentation.Export($result.preview_directory, 'PNG', [int]$result.preview_width, [int]$result.preview_height)
    }

    $result.slide_count = [int]$presentation.Slides.Count
    $result | ConvertTo-Json -Depth 8
}
finally {
    if ($null -ne $templatePresentation) {
        try { $templatePresentation.Close() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($templatePresentation) | Out-Null
    }
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
