<#
.SYNOPSIS
    Rebuild campaign cell merges in PowerPoint tables (COM-based).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR BULK OPERATIONS ⚠️

    This script uses PowerPoint COM automation for bulk table operations
    and is DEPRECATED due to catastrophic performance issues.

    **Performance:**
    - COM automation: 10+ hours for large decks
    - Python replacement: <1 second (1,800x faster!)

    **Replacement:**
    Use Python-based post-processing instead:
      .\tools\PostProcessNormalize.ps1 -PresentationPath deck.pptx

    Or Python CLI directly:
      py -m amp_automation.presentation.postprocess.cli `
        --presentation-path deck.pptx `
        --operations merge-campaign

    **Why Deprecated:**
    COM automation requires PowerPoint instances, is slow, error-prone,
    and cannot run in parallel. Python-pptx is faster, more reliable,
    and doesn't require PowerPoint to be installed.

    **Status:** DEPRECATED as of 27 Oct 2025
    **Migration:** Use PostProcessNormalize.ps1 or Python CLI
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use this script for emergency repairs on legacy decks!
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [int[]]$SlideFilter,

    [double]$DataRowHeight = 8.4,
    [int]$HeaderRows = 1,
    [int]$PrimaryColumns = 3
)

if (-not (Test-Path -LiteralPath $PresentationPath)) {
    throw "Presentation not found: $PresentationPath"
}

$msoTrue = -1
$msoFalse = 0

$campaignFontSize = 6.0
$monthlyTotalFontSize = 6.5
$summaryFontSize = 7.0

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

    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = $msoFalse } catch { Write-Verbose "PowerPoint refused to hide: $_" }
    $pres = $pp.Presentations.Open($Path, $msoFalse, $msoFalse, $msoFalse)
    return @{ App = $pp; Pres = $pres }
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
    param($Slide)

    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.Name -eq 'MainDataTable' -and $shape.HasTable) {
            return $shape.Table
        }
    }
    return $null
}

function Normalize-Label {
    param([string]$Text)

    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
}

function Set-CellFixedLayout {
    param($cell)

    if (-not $cell) { return }

    $shape = $null
    try { $shape = $cell.Shape } catch { return }
    if (-not $shape) { return }

    $tf = $null
    try { $tf = $shape.TextFrame } catch {}
    if ($tf) {
        try { $tf.AutoSize = 0 } catch {}
        try { $tf.WordWrap = -1 } catch {}
        try { $tf.MarginLeft = 0 } catch {}
        try { $tf.MarginRight = 0 } catch {}
        try { $tf.MarginTop = 0 } catch {}
        try { $tf.MarginBottom = 0 } catch {}
        try { $tf.VerticalAnchor = 3 } catch {}
    }

    $tf2 = $null
    try { $tf2 = $shape.TextFrame2 } catch {}
    if ($tf2) {
        try { $tf2.AutoSize = 0 } catch {}
        try { $tf2.MarginLeft = 0 } catch {}
        try { $tf2.MarginRight = 0 } catch {}
        try { $tf2.MarginTop = 0 } catch {}
        try { $tf2.MarginBottom = 0 } catch {}
        try { $tf2.VerticalAnchor = 3 } catch {}
    }
}

function Set-RowHeight {
    param($table, [int]$rowIdx, [double]$height)

    try {
        $row = $table.Rows.Item($rowIdx)
        $row.Height = $height
    } catch {
        Write-Verbose ("Unable to set row {0} height: {1}" -f $rowIdx, $_)
    }
}

function Clear-PrimaryCells {
    param($table, [int]$rowIdx, [int]$primaryColumns)

    $limit = [Math]::Min($primaryColumns, $table.Columns.Count)
    for ($col = 1; $col -le $limit; $col++) {
        try {
            $cell = $table.Cell($rowIdx, $col)
            $tf2 = $null
            try { $tf2 = $cell.Shape.TextFrame2 } catch {}
            if ($tf2) {
                $tf2.TextRange.Text = ''
            } else {
                $cell.Shape.TextFrame.TextRange.Text = ''
            }
            Set-CellFixedLayout -cell $cell
        } catch {}
    }
}

function Merge-CampaignBlock {
    param(
        $table,
        [int]$startRow,
        [int]$endRow,
        [string]$label,
        [double]$rowHeight,
        [int]$primaryColumns
    )

    if ($endRow -lt $startRow) { return }

    for ($row = $startRow + 1; $row -le $endRow; $row++) {
        Clear-PrimaryCells -table $table -rowIdx $row -primaryColumns $primaryColumns
    }

    try {
        $topCell = $table.Cell($startRow, 1)
        $bottomCell = $table.Cell($endRow, 1)
        if ($endRow -gt $startRow) {
            $null = $topCell.Merge($bottomCell)
        }
        Set-CellFixedLayout -cell $topCell
        $range = $topCell.Shape.TextFrame.TextRange
        $range.Text = $label
        $range.Font.Size = $campaignFontSize
        $range.Font.Bold = -1
        $range.ParagraphFormat.Alignment = 2
        $range.Parent.VerticalAnchor = 3
    } catch {
        Write-Warning ("Campaign merge failed rows {0}-{1}: {2}" -f $startRow, $endRow, $_)
    }

    for ($row = $startRow; $row -le $endRow; $row++) {
        Set-RowHeight -table $table -rowIdx $row -height $rowHeight
    }
}

function Merge-MonthlyTotal {
    param(
        $table,
        [int]$rowIdx,
        [string]$label,
        [int]$primaryColumns,
        [double]$rowHeight
    )

    Clear-PrimaryCells -table $table -rowIdx $rowIdx -primaryColumns $primaryColumns

    try {
        $anchor = $table.Cell($rowIdx, 1)
        for ($col = 2; $col -le [Math]::Min($primaryColumns, $table.Columns.Count); $col++) {
            $target = $table.Cell($rowIdx, $col)
            if (-not [object]::ReferenceEquals($anchor.Shape, $target.Shape)) {
                $null = $anchor.Merge($target)
            }
        }
        Set-CellFixedLayout -cell $anchor
        $range = $anchor.Shape.TextFrame.TextRange
        $range.Text = $label
        $range.Font.Size = $monthlyTotalFontSize
        $range.Font.Bold = -1
        $range.ParagraphFormat.Alignment = 2
        $range.Parent.VerticalAnchor = 3
    } catch {
        Write-Warning ("Monthly total merge failed row {0}: {1}" -f $rowIdx, $_)
    }

    Set-RowHeight -table $table -rowIdx $rowIdx -height $rowHeight
}

function Merge-SummaryRow {
    param(
        $table,
        [int]$rowIdx,
        [string]$label,
        [int]$primaryColumns,
        [double]$rowHeight
    )

    Clear-PrimaryCells -table $table -rowIdx $rowIdx -primaryColumns $primaryColumns

    try {
        $anchor = $table.Cell($rowIdx, 1)
        for ($col = 2; $col -le [Math]::Min($primaryColumns, $table.Columns.Count); $col++) {
            $target = $table.Cell($rowIdx, $col)
            if (-not [object]::ReferenceEquals($anchor.Shape, $target.Shape)) {
                $null = $anchor.Merge($target)
            }
        }
        Set-CellFixedLayout -cell $anchor
        $range = $anchor.Shape.TextFrame.TextRange
        $range.Text = $label
        $range.Font.Size = $summaryFontSize
        $range.Font.Bold = -1
        $range.ParagraphFormat.Alignment = 2
        $range.Parent.VerticalAnchor = 3
    } catch {
        Write-Warning ("Summary merge failed row {0}: {1}" -f $rowIdx, $_)
    }

    Set-RowHeight -table $table -rowIdx $rowIdx -height $rowHeight
}

$fullPath = [System.IO.Path]::GetFullPath($PresentationPath)

Stop-PowerPointInstances

$ctx = Open-Presentation -Path $fullPath
try {
    $slides = @($ctx.Pres.Slides)
    foreach ($slide in $slides) {
        $index = $slide.SlideIndex
        if ($SlideFilter -and -not ($SlideFilter -contains $index)) { continue }

        $table = Get-MainTable -Slide $slide
        if (-not $table) { continue }

        Write-Verbose ("Slide {0}: rebuilding campaign merges" -f $index)

        $rowCount = $table.Rows.Count
        if ($rowCount -le $HeaderRows) { continue }

        $campaignStart = $null
        $campaignLabel = $null

        for ($row = $HeaderRows + 1; $row -le $rowCount; $row++) {
            $cellText = $null
            try { $cellText = $table.Cell($row, 1).Shape.TextFrame.TextRange.Text } catch {}
            $normalized = Normalize-Label $cellText
            $upper = $normalized.ToUpper()

            $isMonthlyTotal = $upper.StartsWith('MONTHLY TOTAL')
            $isGrandTotal = $upper -eq 'GRAND TOTAL'
            $isCarriedForward = $upper -eq 'CARRIED FORWARD'
            $isEmpty = [string]::IsNullOrWhiteSpace($normalized)

            if (-not $isMonthlyTotal -and -not $isGrandTotal -and -not $isCarriedForward -and -not $isEmpty) {
                if (-not $campaignStart) {
                    $campaignStart = $row
                    $campaignLabel = $normalized
                }
            }

            if ($isMonthlyTotal) {
                if ($campaignStart) {
                    Merge-CampaignBlock -table $table -startRow $campaignStart -endRow ($row - 1) -label $campaignLabel -rowHeight $DataRowHeight -primaryColumns $PrimaryColumns
                    $campaignStart = $null
                    $campaignLabel = $null
                }

                Merge-MonthlyTotal -table $table -rowIdx $row -label $normalized -primaryColumns $PrimaryColumns -rowHeight $DataRowHeight
                continue
            }

            if ($isGrandTotal -or $isCarriedForward) {
                if ($campaignStart) {
                    Merge-CampaignBlock -table $table -startRow $campaignStart -endRow ($row - 1) -label $campaignLabel -rowHeight $DataRowHeight -primaryColumns $PrimaryColumns
                    $campaignStart = $null
                    $campaignLabel = $null
                }

                Merge-SummaryRow -table $table -rowIdx $row -label $normalized -primaryColumns $PrimaryColumns -rowHeight $DataRowHeight
                continue
            }

            if ($isEmpty -and $campaignStart) {
                Merge-CampaignBlock -table $table -startRow $campaignStart -endRow ($row - 1) -label $campaignLabel -rowHeight $DataRowHeight -primaryColumns $PrimaryColumns
                $campaignStart = $null
                $campaignLabel = $null
            }
        }

        if ($campaignStart) {
            Merge-CampaignBlock -table $table -startRow $campaignStart -endRow $rowCount -label $campaignLabel -rowHeight $DataRowHeight -primaryColumns $PrimaryColumns
        }
    }

    Write-Verbose ("Rebuild complete: {0}" -f $fullPath)
}
finally {
    Close-Presentation -Context $ctx
    Stop-PowerPointInstances
}
