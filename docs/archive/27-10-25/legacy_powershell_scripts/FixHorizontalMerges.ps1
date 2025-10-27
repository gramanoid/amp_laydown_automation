<#
.SYNOPSIS
    Fix horizontal merges in PowerPoint tables using JSON instructions (COM-based).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR DIAGNOSTIC TOOLS ⚠️

    This diagnostic script uses PowerPoint COM automation and is DEPRECATED.
    COM-based tools are slow and error-prone compared to Python alternatives.

    **Replacement:**
    Python-based post-processing handles merge operations automatically:
      .\tools\PostProcessNormalize.ps1 -PresentationPath deck.pptx

    **Status:** DEPRECATED as of 27 Oct 2025
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use for emergency repairs on legacy decks with specific merge issues!
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [string]$InstructionsPath = "docs/22-10-25/merged_cells_analysis/merged_cells_fix_instructions.json"
)

if (-not (Test-Path -LiteralPath $PresentationPath)) {
    throw "Presentation not found: $PresentationPath"
}

if (-not (Test-Path -LiteralPath $InstructionsPath)) {
    throw "Instructions file not found: $InstructionsPath"
}

$instructions = Get-Content -LiteralPath $InstructionsPath -Raw | ConvertFrom-Json
$entries = $instructions.merged_cells

function Normalize-Label([string]$Text) {
    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
}

function Should-KeepMerge($entry) {
    if ($entry.empty) { return $false }
    $label = Normalize-Label($entry.content).ToUpper()
    if (-not $label) { return $false }
    if ($label.StartsWith("MONTHLY TOTAL")) { return $true }
    return $label -in @("GRAND TOTAL", "CARRIED FORWARD")
}

function Get-Table([object]$slide) {
    foreach ($shape in $slide.Shapes) {
        if ($shape.Name -eq "MainDataTable" -and $shape.HasTable) { return $shape.Table }
    }
    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTable) { return $shape.Table }
    }
    return $null
}

function Get-HorizontalSpan([object]$table, [int]$rowIdx, [int]$startCol, [int]$maxCols) {
    try {
        $cell = $table.Cell($rowIdx, $startCol)
    } catch {
        return 1
    }

    if ($null -eq $cell) { return 1 }
    $limit = [Math]::Min($maxCols, $table.Columns.Count)
    $span = 1
    for ($col = $startCol + 1; $col -le $limit; $col++) {
        try {
            $next = $table.Cell($rowIdx, $col)
        } catch {
            break
        }
        if (-not [object]::ReferenceEquals($cell.Shape, $next.Shape)) { break }
        $span += 1
    }
    return $span
}

function Unmerge-Horizontal([object]$table, [int]$rowIdx, [int]$startCol, [int]$maxCols) {
    $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -startCol $startCol -maxCols $maxCols
    if ($span -gt 1) {
        try {
            $table.Cell($rowIdx, $startCol).Split(1, $span) | Out-Null
        } catch {
            Write-Verbose ("Unable to split row {0} columns {1}-{2}: {3}" -f $rowIdx, $startCol, ($startCol + $span - 1), $_)
        }
    }
    $limit = [Math]::Min($maxCols, $table.Columns.Count)
    for ($col = $startCol + 1; $col -le $limit; $col++) {
        try {
            $table.Cell($rowIdx, $col).Shape.TextFrame.TextRange.Text = ""
        } catch { }
    }
}

function Ensure-Merge([object]$table, [int]$rowIdx, [string]$text, [double]$fontSize, [bool]$bold) {
    $maxCols = [Math]::Min(3, $table.Columns.Count)
    Unmerge-Horizontal -table $table -rowIdx $rowIdx -startCol 1 -maxCols $maxCols

    try {
        $cell = $table.Cell($rowIdx, 1)
    } catch {
        return
    }

    for ($col = 2; $col -le $maxCols; $col++) {
        try {
            $target = $table.Cell($rowIdx, $col)
        } catch {
            continue
        }
        if (-not [object]::ReferenceEquals($cell.Shape, $target.Shape)) {
            try {
                $null = $cell.Merge($target)
                $cell = $table.Cell($rowIdx, 1)
            } catch {
                Write-Verbose ("Unable to merge row {0} column {1}: {2}" -f $rowIdx, $col, $_)
                break
            }
        }
    }

    try {
        $range = $cell.Shape.TextFrame.TextRange
        if ($text) { $range.Text = $text }
        $range.ParagraphFormat.Alignment = 2
        $range.Parent.VerticalAnchor = 3
        $range.Font.Size = $fontSize
        $range.Font.Bold = if ($bold) { -1 } else { 0 }
    } catch {
        Write-Verbose ("Unable to format merged cell on row {0}: {1}" -f $rowIdx, $_)
    }
}

$summaryFontSize = 7.0
$monthlyFontSize = 6.5

$pp = $null
$presentation = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = -1 } catch { }
    $presentation = $pp.Presentations.Open($PresentationPath, $false, $false, $false)

    $processed = 0
    foreach ($entry in $entries) {
        $slideIndex = [int]$entry.slide
        $rowIndex = [int]$entry.row
        $content = Normalize-Label $entry.content
        $keepMerge = Should-KeepMerge $entry

        if ($slideIndex -lt 1 -or $slideIndex -gt $presentation.Slides.Count) {
            Write-Verbose ("Skipping slide {0}: out of range" -f $slideIndex)
            continue
        }

        $slide = $presentation.Slides.Item($slideIndex)
        $table = Get-Table $slide
        if ($null -eq $table) {
            Write-Verbose ("Skipping slide {0}: no table found" -f $slideIndex)
            continue
        }

        if ($rowIndex -lt 1 -or $rowIndex -gt $table.Rows.Count) {
            Write-Verbose ("Skipping slide {0} row {1}: out of range" -f $slideIndex, $rowIndex)
            continue
        }

        if ($keepMerge) {
            $isMonthly = $content.ToUpper().StartsWith("MONTHLY TOTAL")
            $fontSize = if ($isMonthly) { $monthlyFontSize } else { $summaryFontSize }
            $bold = $true
            Ensure-Merge -table $table -rowIdx $rowIndex -text $content -fontSize $fontSize -bold $bold
        }
        else {
            Unmerge-Horizontal -table $table -rowIdx $rowIndex -startCol 1 -maxCols 3
            if ($content) {
                try {
                    $range = $table.Cell($rowIndex, 1).Shape.TextFrame.TextRange
                    $range.Text = $content
                    $range.ParagraphFormat.Alignment = 1
                    $range.Font.Bold = 0
                } catch {
                    Write-Verbose ("Unable to restyle unmerged cell on slide {0} row {1}: {2}" -f $slideIndex, $rowIndex, $_)
                }
            }
        }
        $processed++
    }

    Write-Verbose ("Processed {0} rows based on instructions" -f $processed)
    $presentation.Save()
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($pp -ne $null) {
        $pp.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
