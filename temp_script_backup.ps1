param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath
)

if (-not (Test-Path -LiteralPath $PresentationPath)) {
    throw "Presentation not found: $PresentationPath"
}

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
    $presentation = $pp.Presentations.Open($PresentationPath, $false, $false, $false)

    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -ne "MainDataTable" -or -not $shape.HasTable) { continue }

            $table = $shape.Table
            $rowCount = $table.Rows.Count
            if ($rowCount -le 1) { continue }

            # Pre-pass: clear any vertical spans across columns 1-3 so later merges see clean cells.
            for ($colIdx = 1; $colIdx -le [Math]::Min(3, $table.Columns.Count); $colIdx++) {
                for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
                    Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                }
            }

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
                        Reset-VerticalSpan -table $table -rowIdx $campaignStart -colIdx 1

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
                    Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx 1
                    Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx 2
                    Reset-VerticalSpan -table $table -rowIdx $rowIdx -colIdx 3
                    Reset-MonthlyTotalRow -table $table -rowIdx $rowIdx

                    try {
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
                        $summaryRange = $table.Cell($rowIdx, 1).Shape.TextFrame.TextRange
                        $summaryRange.Font.Size = $summaryFontSize
                        $summaryRange.Font.Bold = -1
                    } catch {
                        Write-Verbose ("Unable to style summary cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_)
                    }
                }
            }

            $targetHeight = 8.4
            for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
                try {
                    $row = $table.Rows.Item($rowIdx)
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
