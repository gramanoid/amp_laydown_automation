<#
.SYNOPSIS
    Verify allowed horizontal merges in PowerPoint tables (COM-based).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR DIAGNOSTIC TOOLS ⚠️

    This verification script uses PowerPoint COM automation and is DEPRECATED.
    COM-based tools are 60x slower than Python alternatives.

    **Replacement:**
    Python post-processing ensures correct merges automatically:
      .\tools\PostProcessNormalize.ps1 -PresentationPath deck.pptx

    **Status:** DEPRECATED as of 27 Oct 2025
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use for legacy deck verification!
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [switch]$FailOnViolation
)

$normalizedPath = [System.IO.Path]::GetFullPath($PresentationPath).Replace('/', '\')

if (-not (Test-Path -LiteralPath $normalizedPath)) {
    throw "Presentation not found: $normalizedPath"
}

function Normalize-Label([string]$Text) {
    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
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

function Get-HorizontalSpan([object]$table, [int]$rowIdx, [int]$startCol) {
    try {
        $cell = $table.Cell($rowIdx, $startCol)
    } catch {
        return 0
    }

    if ($null -eq $cell) { return 0 }
    $span = 1
    for ($col = $startCol + 1; $col -le $table.Columns.Count; $col++) {
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

$allowedLabels = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$null = $allowedLabels.Add("MONTHLY TOTAL (� 000)")
$null = $allowedLabels.Add("MONTHLY TOTAL")
$null = $allowedLabels.Add("GRAND TOTAL")
$null = $allowedLabels.Add("CARRIED FORWARD")

$pp = $null
$presentation = $null

$mergedRows = New-Object System.Collections.Generic.List[object]
$violations = New-Object System.Collections.Generic.List[object]

try {
    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = 0 } catch { }

    $presentation = $pp.Presentations.Open($normalizedPath, 0, 0, 0)

    for ($slideIdx = 1; $slideIdx -le $presentation.Slides.Count; $slideIdx++) {
        $slide = $presentation.Slides.Item($slideIdx)
        $table = Get-Table $slide
        if (-not $table) { continue }

        for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
            $label = ""
            try {
                $label = Normalize-Label($table.Cell($rowIdx, 1).Shape.TextFrame.TextRange.Text)
            } catch { }

            $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -startCol 1
            if ($span -gt 1) {
                $entry = [pscustomobject]@{
                    Slide = $slideIdx
                    Row   = $rowIdx
                    Span  = $span
                    Label = $label
                }
                $mergedRows.Add($entry)

                if (-not $allowedLabels.Contains($label)) {
                    $violations.Add($entry) | Out-Null
                }
            }

            for ($colIdx = 2; $colIdx -le [Math]::Min(3, $table.Columns.Count); $colIdx++) {
                $colSpan = Get-HorizontalSpan -table $table -rowIdx $rowIdx -startCol $colIdx
                if ($colSpan -gt 1) {
                    $entry = [pscustomobject]@{
                        Slide = $slideIdx
                        Row   = $rowIdx
                        Span  = $colSpan
                        Label = "<merge starts at column $colIdx>"
                    }
                    $violations.Add($entry) | Out-Null
                }
            }
        }
    }
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

Write-Output "Merged rows across columns 1-3: $($mergedRows.Count)"
if ($mergedRows.Count -gt 0) {
    $mergedRows | Sort-Object Slide, Row | Format-Table -AutoSize | Out-String | Write-Output
}

if ($violations.Count -gt 0) {
    Write-Warning ("Violations detected: {0}" -f $violations.Count)
    $violations | Sort-Object Slide, Row | Format-Table -AutoSize | Out-String | Write-Warning
    if ($FailOnViolation) {
        throw "Horizontal merge violations detected."
    }
} else {
    Write-Output "No horizontal merge violations detected."
}

