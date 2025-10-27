<#
.SYNOPSIS
    Inspect column spans in PowerPoint tables (COM-based diagnostic).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR DIAGNOSTIC TOOLS ⚠️

    This diagnostic script uses PowerPoint COM automation and is DEPRECATED.
    COM-based inspection tools are slow and limited compared to Python alternatives.

    **Replacement:**
    Python-pptx provides better table inspection capabilities:
      py -m amp_automation.presentation.postprocess.cli `
        --presentation-path deck.pptx `
        --operations postprocess-all --verbose

    **Status:** DEPRECATED as of 27 Oct 2025
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use for quick legacy deck inspection!
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PresentationPath,
    [int]$SlideIndex = 1,
    [int]$Column = 1,
    [int]$Rows = 10
)

if (-not (Test-Path -LiteralPath $PresentationPath)) {
    throw "Presentation not found: $PresentationPath"
}

$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = 0 } catch { Write-Verbose "PowerPoint refused to hide: $_" }
    $pres = $pp.Presentations.Open($PresentationPath, 0, 0, 0)

    $slide = $pres.Slides.Item($SlideIndex)
    $table = $null
    foreach ($shape in @($slide.Shapes)) {
        if ($shape.Name -eq 'MainDataTable' -and $shape.HasTable) {
            $table = $shape.Table
            break
        }
    }

    if (-not $table) {
        Write-Output "No MainDataTable found on slide $SlideIndex"
        return
    }

    $maxRow = [Math]::Min($Rows, $table.Rows.Count)
    for ($row = 1; $row -le $maxRow; $row++) {
        try {
            $cell = $table.Cell($row, $Column)
            $span = 1
            for ($cursor = $row + 1; $cursor -le $table.Rows.Count; $cursor++) {
                $candidate = $table.Cell($cursor, $Column)
                if ([object]::ReferenceEquals($candidate.Shape, $cell.Shape)) {
                    $span++
                } else {
                    break
                }
            }
            Write-Output "Row $row -> span $span"
        } catch {
            Write-Output "Row $row -> error $($_.Exception.Message)"
        }
    }
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
}
