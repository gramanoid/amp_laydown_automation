<#
.SYNOPSIS
    Sanitize primary columns in PowerPoint tables (COM-based).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR BULK OPERATIONS ⚠️

    This script uses PowerPoint COM automation and is DEPRECATED due to
    catastrophic performance issues (60x slower than Python).

    **Replacement:**
    Use Python-based post-processing instead:
      .\tools\PostProcessNormalize.ps1 -PresentationPath deck.pptx

    Or Python CLI directly:
      py -m amp_automation.presentation.postprocess.cli `
        --presentation-path deck.pptx `
        --operations normalize

    **Status:** DEPRECATED as of 27 Oct 2025
    **Migration:** Use PostProcessNormalize.ps1 or Python CLI
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [int[]]$Columns = @(1, 2, 3),
    [int]$HeaderRows = 1,
    [double]$RowHeightPt = 8.4,
    [int[]]$SlideIndexFilter = @(),
    [string]$TableNameRegex = '.*'
)

$msoTrue = -1
$msoFalse = 0
$msoAutoSizeNone = 0
$msoAnchorTop = 1

function Stop-PowerPointInstances {
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not $_.HasExited) {
            try { $_.CloseMainWindow() | Out-Null } catch {}
        }
    }
    Start-Sleep -Milliseconds 200
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Open-Presentation {
    param([string]$Path)

    $app = New-Object -ComObject PowerPoint.Application
    try { $app.Visible = $msoFalse } catch { Write-Verbose "PowerPoint refused to hide: $_" }
    return @{ App = $app; Pres = $app.Presentations.Open($Path, $msoFalse, $msoFalse, $msoFalse) }
}

function Close-Presentation {
    param($Context)

    if ($Context.Pres) {
        try { $Context.Pres.Save() } catch {}
        try { $Context.Pres.Close() } catch {}
    }
    if ($Context.App) {
        try { $Context.App.Quit() } catch {}
    }
}

function Get-MainTable {
    param($Slide, [string]$NameRegex)

    $tables = @()
    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.HasTable -eq $msoTrue -and $shape.Name -match $NameRegex) {
            $area = $shape.Table.Rows.Count * $shape.Table.Columns.Count
            $tables += [pscustomobject]@{ Shape = $shape; Area = $area }
        }
    }
    $selected = $tables | Sort-Object Area -Descending | Select-Object -First 1
    if (-not $selected) { return $null }
    if (-not ($selected | Get-Member -Name Shape -ErrorAction SilentlyContinue)) { return $null }
    return $selected.Shape.Table
}

function Get-CellText {
    param($Cell)

    if (-not $Cell) { return $null }
    if (-not ($Cell | Get-Member -Name Shape -ErrorAction SilentlyContinue)) { return $null }

    try {
        $text = $Cell.Shape.TextFrame2.TextRange.Text
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return ($text -replace '\s+$', '')
    } catch {
        return $null
    }
}

function Capture-ColumnValues {
    param($Table, [int]$Column, [int]$HeaderRows)

    $rowCount = $Table.Rows.Count
    $values = New-Object string[] $rowCount
    $last = $null

    for ($row = 1; $row -le $rowCount; $row++) {
        $text = $null
        try {
            $cell = $Table.Cell($row, $Column)
            $text = Get-CellText -Cell $cell
        } catch {
            $text = $null
        }

        if ($row -le $HeaderRows) {
            $values[$row - 1] = $text
            if ($text) { $last = $text }
            continue
        }

        if (-not $text) { $text = $last } else { $last = $text }
        $values[$row - 1] = $text
    }

    return $values
}

function Set-CellText {
    param($Cell, [string]$Text)

    if (-not $Cell) { return }
    if (-not ($Cell | Get-Member -Name Shape -ErrorAction SilentlyContinue)) {
        Write-Verbose ('Skipping cell lacking Shape property (type: {0})' -f $Cell.GetType().FullName)
        return
    }

    $tf2 = $null
    try { $tf2 = $Cell.Shape.TextFrame2 } catch {}
    if ($tf2) {
        try { $tf2.AutoSize = $msoAutoSizeNone } catch {}
        try { $tf2.VerticalAnchor = $msoAnchorTop } catch {}
        try { $tf2.MarginTop = 1 } catch {}
        try { $tf2.MarginBottom = 1 } catch {}
        try { $tf2.MarginLeft = 1.5 } catch {}
        try { $tf2.MarginRight = 1.5 } catch {}
    }

    $tf = $null
    try { $tf = $Cell.Shape.TextFrame } catch {}
    if ($tf) {
        try { $tf.AutoSize = 0 } catch {}
        try { $tf.VerticalAnchor = 1 } catch {}
        try { $tf.MarginTop = 0 } catch {}
        try { $tf.MarginBottom = 0 } catch {}
        try { $tf.MarginLeft = 0 } catch {}
        try { $tf.MarginRight = 0 } catch {}
    }

    $value = if ($Text) { $Text } else { '' }

    if ($tf2) {
        try { $tf2.TextRange.Text = $value; return } catch {}
    }
    if ($tf) {
        try { $tf.TextRange.Text = $value; return } catch {}
    }
}

function Sanitize-Table {
    param($Table, [int[]]$Columns, [int]$HeaderRows, [double]$RowHeightPt)

    $rowCount = $Table.Rows.Count
    if ($rowCount -lt 2) { return }

    foreach ($col in ($Columns | Sort-Object -Descending)) {
        if ($col -gt $Table.Columns.Count) { continue }

        Write-Verbose ("Rebuilding column {0}" -f $col)

        $width = $Table.Columns.Item($col).Width
        $values = Capture-ColumnValues -Table $Table -Column $col -HeaderRows $HeaderRows

        $Table.Columns.Item($col).Delete()
        [void]$Table.Columns.Add($col)
        $newCol = $Table.Columns.Item($col)
        $newCol.Width = $width

        for ($row = 1; $row -le $rowCount; $row++) {
        $value = $values[$row - 1]
        try {
            $dest = $Table.Cell($row, $col)
        } catch {
            Write-Verbose ("Unable to access cell row {0} col {1}: {2}" -f $row, $col, $_)
            continue
        }

        try {
            Set-CellText -Cell $dest -Text $value
        } catch {
            Write-Verbose ("Failed to set text for row {0} col {1}: {2}" -f $row, $col, $_)
        }
    }
    }

    for ($row = [Math]::Max($HeaderRows + 1, 1); $row -le $rowCount; $row++) {
        try { $Table.Rows.Item($row).Height = $RowHeightPt } catch {
            Write-Verbose ("Unable to set row {0} height: {1}" -f $row, $_)
        }
    }
}

if (-not (Test-Path -LiteralPath $PresentationPath)) {
    throw "Presentation not found: $PresentationPath"
}

$fullPath = [System.IO.Path]::GetFullPath($PresentationPath)

Stop-PowerPointInstances

$ctx = Open-Presentation -Path $fullPath
try {
$slides = $ctx.Pres.Slides
foreach ($slide in $slides) {
    $index = $slide.SlideIndex
    if ($SlideIndexFilter.Count -gt 0 -and -not ($SlideIndexFilter -contains $index)) { continue }
    $table = Get-MainTable -Slide $slide -NameRegex $TableNameRegex
    if (-not $table) { continue }
    Write-Verbose ("Slide {0}: sanitizing columns {1}" -f $index, ($Columns -join ','))
    Sanitize-Table -Table $table -Columns $Columns -HeaderRows $HeaderRows -RowHeightPt $RowHeightPt
    }
    Write-Verbose ("Sanitization complete: {0}" -f $fullPath)
}
finally {
    Close-Presentation -Context $ctx
    Stop-PowerPointInstances
}
