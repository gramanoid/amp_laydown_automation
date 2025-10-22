param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath
)

$normalizedPath = [System.IO.Path]::GetFullPath($PresentationPath).Replace('/', '\')

if (-not (Test-Path -LiteralPath $normalizedPath)) {
    throw "Presentation not found: $normalizedPath"
}

Write-Verbose ("Using presentation path: {0}" -f $normalizedPath)

$campaignFontSize = 6.0
$monthlyTotalFontSize = 6.5
$summaryFontSize = 7.0

function Normalize-Label {
    param([string]$Text)
    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
}

function Reset-VerticalSpan {
    param($table, [int]$rowIdx, [int]$colIdx)

    try {
        $cell = $table.Cell($rowIdx, $colIdx)
    } catch {
        return
    }

    if ($null -eq $cell) { return }

    $top = $rowIdx
    while ($top -gt 1) {
        try {
            $above = $table.Cell($top - 1, $colIdx)
        } catch {
            break
        }

        if (-not [object]::ReferenceEquals($cell.Shape, $above.Shape)) { break }
        $top -= 1
    }

    $bottom = $rowIdx
    while ($bottom -lt $table.Rows.Count) {
        try {
            $below = $table.Cell($bottom + 1, $colIdx)
        } catch {
            break
        }

        if (-not [object]::ReferenceEquals($cell.Shape, $below.Shape)) { break }
        $bottom += 1
    }

    $span = $bottom - $top + 1
    if ($span -gt 1) {
        try {
            $table.Cell($top, $colIdx).Split($span, 1) | Out-Null
        } catch {
            Write-Verbose ("Unable to split vertical span on column {0} rows {1}-{2}: {3}" -f $colIdx, $top, $bottom, $_)
        }
    }
}

function Reset-HorizontalSpan {
    param($table, [int]$rowIdx, [int]$startCol, [int]$endCol)

    try {
        $row = $table.Rows.Item($rowIdx)
    } catch {
        return
    }

    $limit = [Math]::Min($endCol, $table.Columns.Count)
    $col = $startCol
    while ($col -le $limit) {
        try {
            $cell = $table.Cell($rowIdx, $col)
        } catch {
            break
        }

        if ($null -eq $cell) {
            $col += 1
            continue
        }

        $span = 1
        for ($next = $col + 1; $next -le $limit; $next++) {
            try {
                $nextCell = $table.Cell($rowIdx, $next)
            } catch {
                break
            }

            if (-not [object]::ReferenceEquals($cell.Shape, $nextCell.Shape)) { break }
            $span += 1
        }

        if ($span -gt 1) {
            try {
                $cell.Split(1, $span) | Out-Null
                for ($offset = 1; $offset -lt $span; $offset++) {
                    try {
                        $table.Cell($rowIdx, $col + $offset).Shape.TextFrame.TextRange.Text = ""
                    } catch { }
                }
            } catch {
                Write-Verbose ("Unable to split horizontal span on row {0} columns {1}-{2}: {3}" -f $rowIdx, $col, ($col + $span - 1), $_)
            }
        }

        $col += $span
    }
}

function Reset-MonthlyTotalRow {
    param($table, [int]$rowIdx)

    try {
        $cell = $table.Cell($rowIdx, 1)
    } catch {
        return
    }

    if ($null -eq $cell) { return }

    $colSpan = 1
    for ($col = 2; $col -le $table.Columns.Count; $col++) {
        try {
            $nextCell = $table.Cell($rowIdx, $col)
        } catch {
            break
        }

        if (-not [object]::ReferenceEquals($cell.Shape, $nextCell.Shape)) { break }
        $colSpan += 1
    }

    if ($colSpan -gt 1) {
        try {
            $cell.Split(1, $colSpan) | Out-Null
        } catch {
            Write-Verbose ("Unable to split monthly total span on row {0}: {1}" -f $rowIdx, $_)
        }
    }
}

$pp = $null
$presentation = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = -1 } catch { }
    try {
        $presentation = $pp.Presentations.Open($normalizedPath, $false, $false, $true)
    } catch {
        Write-Verbose ("Initial open failed for {0}: {1}" -f $normalizedPath, $_)
        Start-Sleep -Milliseconds 500
        $presentation = $pp.Presentations.Open($normalizedPath, $false, $false, $true)
    }

    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -ne "MainDataTable" -or -not $shape.HasTable) { continue }

            $table = $shape.Table
            $rowCount = $table.Rows.Count
            if ($rowCount -le 1) { continue }

            $campaignStart = $null
            for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
                try {
                    $labelCell = $table.Cell($rowIdx, 1)
                } catch {
                    Write-Verbose ("Unable to read label cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_)
                    continue
                }

                $normalized = Normalize-Label -Text $labelCell.Shape.TextFrame.TextRange.Text
                $normalizedUpper = $normalized.ToUpper()

                $isMonthlyTotal = $normalizedUpper.StartsWith("MONTHLY TOTAL")
                $isGrandTotal = $normalizedUpper -eq "GRAND TOTAL"
                $isCarriedForward = $normalizedUpper -eq "CARRIED FORWARD"

                if ($normalized -and -not $isMonthlyTotal -and -not $isGrandTotal -and -not $isCarriedForward) {
                    if ($campaignStart -eq $null) { $campaignStart = $rowIdx }
                }

                if ($isMonthlyTotal -and $campaignStart -ne $null) {
                    $campaignEnd = $rowIdx - 1
                    if ($campaignEnd -gt $campaignStart) {
                        $maxCampaignCol = [Math]::Min(3, $table.Columns.Count)
                        for ($campaignRow = $campaignStart; $campaignRow -le $campaignEnd; $campaignRow++) {
                            for ($colIdx = 1; $colIdx -le $maxCampaignCol; $colIdx++) {
                                Reset-VerticalSpan -table $table -rowIdx $campaignRow -colIdx $colIdx
                            }
                            Reset-HorizontalSpan -table $table -rowIdx $campaignRow -startCol 1 -endCol $maxCampaignCol
                        }

                        $topCell = $table.Cell($campaignStart, 1)
                        $bottomCell = $table.Cell($campaignEnd, 1)
                        if (-not [object]::ReferenceEquals($topCell.Shape, $bottomCell.Shape)) {
                            try {
                                $null = $topCell.Merge($bottomCell)
                            } catch {
                                Write-Verbose ("Campaign merge failed on slide {0} rows {1}-{2}: {3}" -f $slide.SlideIndex, $campaignStart, $campaignEnd, $_)
                            }
                        }

                        try {
                            $campaignCell = $table.Cell($campaignStart, 1)
                            $campaignRange = $campaignCell.Shape.TextFrame.TextRange
                            if (-not (Normalize-Label -Text $campaignRange.Text)) { $campaignRange.Text = $normalized }
                            $campaignRange.ParagraphFormat.Alignment = 2
                            $campaignRange.Parent.VerticalAnchor = 3
                            $campaignRange.Font.Size = $campaignFontSize
                            $campaignRange.Font.Bold = -1
                        } catch {
                            Write-Verbose ("Unable to style campaign cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $campaignStart, $_)
                        }
                    }

                    $campaignStart = $null
                }

                if ($isMonthlyTotal) {
                    $maxCampaignCol = [Math]::Min(3, $table.Columns.Count)
                    for ($colIdx = 1; $colIdx -le $maxCampaignCol; $colIdx++) {
                        Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                    }
                    Reset-HorizontalSpan -table $table -rowIdx $rowIdx -startCol 1 -endCol $maxCampaignCol

                    try {
                        Reset-MonthlyTotalRow -table $table -rowIdx $rowIdx

                        $mergedCell = $table.Cell($rowIdx, 1)
                        if ($table.Columns.Count -ge 2 -and -not [object]::ReferenceEquals($mergedCell.Shape, $table.Cell($rowIdx, 2).Shape)) {
                            $null = $mergedCell.Merge($table.Cell($rowIdx, 2))
                            $mergedCell = $table.Cell($rowIdx, 1)
                        }
                        if ($table.Columns.Count -ge 3 -and -not [object]::ReferenceEquals($mergedCell.Shape, $table.Cell($rowIdx, 3).Shape)) {
                            $null = $mergedCell.Merge($table.Cell($rowIdx, 3))
                            $mergedCell = $table.Cell($rowIdx, 1)
                        }

                        $monthlyRange = $mergedCell.Shape.TextFrame.TextRange
                        $monthlyRange.Text = $normalized
                        $monthlyRange.ParagraphFormat.Alignment = 2
                        $monthlyRange.Parent.VerticalAnchor = 3
                        $monthlyRange.Font.Size = $monthlyTotalFontSize
                        $monthlyRange.Font.Bold = -1
                    } catch {
                        Write-Verbose ("Monthly total merge failed on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_)
                    }
                }

                if ($isGrandTotal -or $isCarriedForward) {
                    try {
                        $maxCampaignCol = [Math]::Min(3, $table.Columns.Count)
                        for ($colIdx = 1; $colIdx -le $maxCampaignCol; $colIdx++) {
                            Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                        }
                        Reset-HorizontalSpan -table $table -rowIdx $rowIdx -startCol 1 -endCol $maxCampaignCol

                        $summaryCell = $table.Cell($rowIdx, 1)
                        if ($table.Columns.Count -ge 2 -and -not [object]::ReferenceEquals($summaryCell.Shape, $table.Cell($rowIdx, 2).Shape)) {
                            $null = $summaryCell.Merge($table.Cell($rowIdx, 2))
                            $summaryCell = $table.Cell($rowIdx, 1)
                        }
                        if ($table.Columns.Count -ge 3 -and -not [object]::ReferenceEquals($summaryCell.Shape, $table.Cell($rowIdx, 3).Shape)) {
                            $null = $summaryCell.Merge($table.Cell($rowIdx, 3))
                            $summaryCell = $table.Cell($rowIdx, 1)
                        }

                        $summaryRange = $summaryCell.Shape.TextFrame.TextRange
                        $summaryRange.Text = $normalized
                        $summaryRange.Font.Size = $summaryFontSize
                        $summaryRange.Font.Bold = -1
                        $summaryRange.ParagraphFormat.Alignment = 2
                        $summaryRange.Parent.VerticalAnchor = 3
                    } catch {
                        Write-Verbose ("Unable to style summary cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_)
                    }
                }
            }

            $targetHeight = 8.4
            for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
                try {
                    $row = $table.Rows.Item($rowIdx)
                    try { $row.HeightRule = 1 } catch { }
                    $row.Height = $targetHeight
                } catch {
                    Write-Verbose ("Unable to enforce row height on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_)
                }
            }
        }
    }

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
