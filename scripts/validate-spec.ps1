[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Add-Issue {
    param(
        [ValidateSet('error', 'warning', 'notice')]
        [string]$Level,
        [string]$Code,
        [string]$Message,
        [int]$OperationIndex = -1
    )

    $item = [ordered]@{ code = $Code; message = $Message }
    if ($OperationIndex -ge 0) { $item.operation_index = $OperationIndex }
    switch ($Level) {
        'error'   { $script:Errors += [pscustomobject]$item }
        'warning' { $script:Warnings += [pscustomobject]$item }
        'notice'  { $script:Notices += [pscustomobject]$item }
    }
}

function Has-Property {
    param([object]$Object, [string]$Name)
    return $null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]
}

function Normalize-Color {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $text = [string]$Value
    if ($text -match '^#[0-9A-Fa-f]{6}$') { return $text.ToUpperInvariant() }
    return $text
}

function Test-ColorObject {
    param([object]$Object, [string]$Context, [int]$OperationIndex)
    if ($null -eq $Object -or -not (Has-Property $Object 'color')) { return }
    $color = Normalize-Color $Object.color
    if ($color -notmatch '^#[0-9A-F]{6}$') {
        Add-Issue error 'invalid_color' "$Context color must use #RRGGBB: $color" $OperationIndex
        return
    }
    if ($script:AllowedColors -notcontains $color) {
        Add-Issue error 'unregistered_color' "$Context uses an unregistered color: $color" $OperationIndex
    }
}

function Resolve-SpecRelativePath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $null }
    if ([IO.Path]::IsPathRooted($PathValue)) { return [IO.Path]::GetFullPath($PathValue) }
    return [IO.Path]::GetFullPath((Join-Path $script:SpecDirectory $PathValue))
}

$script:Errors = @()
$script:Warnings = @()
$script:Notices = @()
$script:AllowedColors = @(
    '#E21413', '#000000', '#0D0D0D', '#FFFFFF',
    '#F9D0D0', '#F3A1A1', '#EE7271', '#F98A8A',
    '#FADBDF', '#F0B6BA', '#B82020', '#F2F2F2',
    '#BFBFBF', '#D9D9D9'
)
$yahei = -join ([char[]](0x5FAE, 0x8F6F, 0x96C5, 0x9ED1))
$allowedFonts = @($yahei, 'Microsoft YaHei', 'Microsoft YaHei UI', 'Arial', '+mj-lt', '+mn-lt', '+mj-ea', '+mn-ea')
$allowedFits = @('cover', 'contain', 'stretch')
$geometryFields = @('left', 'top', 'width', 'height')

$resolvedSpecPath = [IO.Path]::GetFullPath($SpecPath)
if (-not (Test-Path -LiteralPath $resolvedSpecPath -PathType Leaf)) {
    throw "Spec file not found: $resolvedSpecPath"
}
$script:SpecDirectory = Split-Path -Parent $resolvedSpecPath

try {
    $spec = Get-Content -LiteralPath $resolvedSpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
catch {
    throw "Spec is not valid JSON: $($_.Exception.Message)"
}

$layoutMapPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'assets\layout-map.json'
$layoutMap = Get-Content -LiteralPath $layoutMapPath -Raw -Encoding UTF8 | ConvertFrom-Json
$registeredLayouts = @($layoutMap.layouts.PSObject.Properties.Name)

if (-not (Has-Property $spec 'output') -or [string]::IsNullOrWhiteSpace([string]$spec.output)) {
    Add-Issue error 'missing_output' 'Top-level output is required.'
}
if (-not (Has-Property $spec 'operations') -or $null -eq $spec.operations) {
    Add-Issue error 'missing_operations' 'Top-level operations array is required.'
}

$isNewDeck = (Has-Property $spec 'new_deck') -and [bool]$spec.new_deck
$hasInput = (Has-Property $spec 'input') -and -not [string]::IsNullOrWhiteSpace([string]$spec.input)
if ($isNewDeck -and $hasInput) {
    Add-Issue error 'conflicting_source' 'new_deck:true and input cannot be used together.'
}
if (-not $isNewDeck -and -not $hasInput) {
    Add-Issue error 'missing_source' 'Edit mode requires input; create mode requires new_deck:true.'
}

if (Has-Property $spec 'output') {
    $resolvedOutput = Resolve-SpecRelativePath ([string]$spec.output)
    if ($hasInput) {
        $resolvedInput = Resolve-SpecRelativePath ([string]$spec.input)
        if ($resolvedOutput -and $resolvedInput -and $resolvedOutput -eq $resolvedInput) {
            Add-Issue error 'output_overwrites_input' 'output must not be the same file as input.'
        }
    }
    $templateValue = if ((Has-Property $spec 'template') -and $spec.template) { [string]$spec.template } else { [string]$layoutMap.template_file }
    $resolvedTemplate = Resolve-SpecRelativePath $templateValue
    if ($resolvedOutput -and $resolvedTemplate -and $resolvedOutput -eq $resolvedTemplate) {
        Add-Issue error 'output_overwrites_template' 'output must not overwrite the master template.'
    }
}

$operationNames = @()
if (Has-Property $spec 'operations' -and $null -ne $spec.operations) {
    $allOperations = @($spec.operations)
    for ($index = 0; $index -lt $allOperations.Count; $index++) {
        $op = $allOperations[$index]
        $operationName = if (Has-Property $op 'op') { [string]$op.op } else { '' }
        if ([string]::IsNullOrWhiteSpace($operationName)) {
            Add-Issue error 'missing_operation_name' 'Operation is missing op.' $index
            continue
        }
        $operationNames += $operationName

        if ($operationName -eq 'insert_template_slide') {
            if (-not (Has-Property $op 'layout_id') -or [string]::IsNullOrWhiteSpace([string]$op.layout_id)) {
                Add-Issue error 'missing_layout_id' 'insert_template_slide requires layout_id.' $index
            }
            elseif ($registeredLayouts -notcontains [string]$op.layout_id) {
                Add-Issue error 'unregistered_layout' "Unregistered layout: $($op.layout_id)" $index
            }
        }

        if (Has-Property $op 'font' -and $null -ne $op.font) {
            if ((Has-Property $op.font 'name') -and $op.font.name -and $allowedFonts -notcontains [string]$op.font.name) {
                Add-Issue error 'unregistered_font' "Unregistered font: $($op.font.name)" $index
            }
            Test-ColorObject $op.font 'Font' $index
            if ((Has-Property $op.font 'size') -and $null -ne $op.font.size) {
                $size = [double]$op.font.size
                $roleOrName = (([string]$op.role) + ' ' + ([string]$op.name)).ToLowerInvariant()
                $smallTextException = $roleOrName -match 'section|label|copyright|page|footer'
                if ($size -lt 12 -and -not $smallTextException) {
                    Add-Issue warning 'small_body_text' "Body or annotation text is below 12 pt: $size pt." $index
                }
            }
        }

        if (Has-Property $op 'fill') { Test-ColorObject $op.fill 'Fill' $index }
        if (Has-Property $op 'line') { Test-ColorObject $op.line 'Line' $index }

        foreach ($field in $geometryFields) {
            if (-not (Has-Property $op $field) -or $null -eq $op.$field) { continue }
            $value = [double]$op.$field
            if ($value -lt 0) {
                Add-Issue error 'negative_geometry' "$field must not be negative: $value" $index
            }
        }
        if ((Has-Property $op 'left') -and (Has-Property $op 'width') -and $null -ne $op.left -and $null -ne $op.width) {
            if (([double]$op.left + [double]$op.width) -gt 960.01) {
                Add-Issue warning 'outside_canvas_x' 'Object extends past the 960 pt right edge.' $index
            }
        }
        if ((Has-Property $op 'top') -and (Has-Property $op 'height') -and $null -ne $op.top -and $null -ne $op.height) {
            if (([double]$op.top + [double]$op.height) -gt 540.01) {
                Add-Issue warning 'outside_canvas_y' 'Object extends past the 540 pt bottom edge.' $index
            }
        }

        if (Has-Property $op 'fit' -and $op.fit) {
            if ($allowedFits -notcontains [string]$op.fit) {
                Add-Issue error 'invalid_image_fit' "Image fit must be cover, contain, or stretch: $($op.fit)" $index
            }
            elseif ([string]$op.fit -eq 'stretch') {
                Add-Issue warning 'stretched_image' 'Image fit is stretch and may distort the image.' $index
            }
        }

        if ($operationName -in @('replace_image', 'insert_image') -and (Has-Property $op 'path')) {
            $imagePath = Resolve-SpecRelativePath ([string]$op.path)
            if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) {
                Add-Issue error 'missing_image' "Image file not found: $imagePath" $index
            }
        }

        if ($operationName -eq 'set_text' -and (Has-Property $op 'text')) {
            $text = [string]$op.text
            if ([string]::IsNullOrWhiteSpace($text)) {
                Add-Issue warning 'empty_set_text' 'set_text writes empty text; use delete_shape for unused objects.' $index
            }
            foreach ($line in ($text -split "`r?`n")) {
                if ($line.Trim() -match '^[\u3400-\u9FFF]$') {
                    Add-Issue warning 'orphan_character' "Text has a one-character line: $($line.Trim())" $index
                    break
                }
            }
        }
    }
}

foreach ($required in @('save', 'export_pdf', 'export_png')) {
    if ($operationNames -notcontains $required) {
        Add-Issue warning "missing_$required" "Candidate spec should include $required."
    }
}

$insertCount = @($operationNames | Where-Object { $_ -eq 'insert_template_slide' }).Count
$notesCount = @($operationNames | Where-Object { $_ -eq 'set_speaker_notes' }).Count
if ($insertCount -gt 0 -and $notesCount -lt $insertCount) {
    Add-Issue warning 'missing_speaker_notes' "Inserted $insertCount slides but wrote notes for only $notesCount slides."
}

if ((Has-Property $spec 'overwrite') -and [bool]$spec.overwrite) {
    Add-Issue notice 'overwrite_enabled' 'overwrite:true is for a user-approved final path only.'
}
foreach ($field in @('input', 'template', 'output')) {
    if ((Has-Property $spec $field) -and $spec.$field -and [IO.Path]::IsPathRooted([string]$spec.$field)) {
        Add-Issue notice 'absolute_path' "$field uses an absolute path; use a relative path for cross-machine sharing."
    }
}

$result = [ordered]@{
    input = $resolvedSpecPath
    valid = ($script:Errors.Count -eq 0)
    error_count = $script:Errors.Count
    warning_count = $script:Warnings.Count
    notice_count = $script:Notices.Count
    errors = @($script:Errors)
    warnings = @($script:Warnings)
    notices = @($script:Notices)
}

$json = $result | ConvertTo-Json -Depth 8
if ($OutputPath) {
    $resolvedReportPath = if ([IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path $script:SpecDirectory $OutputPath }
    $parent = Split-Path -Parent $resolvedReportPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($resolvedReportPath), $json, [Text.UTF8Encoding]::new($false))
}

$json
if ($script:Errors.Count -gt 0) { exit 1 }
