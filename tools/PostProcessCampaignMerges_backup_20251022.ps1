param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [int[]]$SlideFilter
)

$normalizedPath = [System.IO.Path]::GetFullPath($PresentationPath).Replace('/', '\')

if (-not (Test-Path -LiteralPath $normalizedPath)) {
    throw "Presentation not found: $normalizedPath"
}

Write-Verbose ("Using presentation path: {0}" -f $normalizedPath)

$unblockSucceeded = $false
try {
    Unblock-File -LiteralPath $normalizedPath -ErrorAction Stop
    $unblockSucceeded = $true
} catch {
    Write-Verbose ("Unblock-File skipped for {0}: {1}" -f $normalizedPath, $_)
}
if ($unblockSucceeded) {
    Write-Verbose ("Removed zone identifier for: {0}" -f $normalizedPath)
}

$campaignFontSize = 6.0
$monthlyTotalFontSize = 6.5
$summaryFontSize = 7.0

function Stop-PowerPointInstances {
    param([int]$WaitMilliseconds = 500)

    $existing = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if (-not $existing) { return }

    Write-Verbose ("Stopping {0} lingering POWERPNT processes" -f $existing.Count)
    foreach ($proc in $existing) {
        try {
            if (-not $proc.HasExited) {
                $null = $proc.CloseMainWindow()
            }
        } catch { }
    }

    if ($WaitMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $WaitMilliseconds
    }

    $remaining = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($remaining) {
        foreach ($proc in $remaining) {
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            } catch { }
        }
        Start-Sleep -Milliseconds 250
    }
}

function Close-BlankPresentations {
    param(
        [Parameter(Mandatory = $true)]
        $Application
    )

    try {
        $presentations = $Application.Presentations
        if (-not $presentations) { return }

        for ($index = $presentations.Count; $index -ge 1; $index--) {
            $candidate = $null
            try { $candidate = $presentations.Item($index) } catch { continue }
            if (-not $candidate) { continue }

            $fullName = $null
            try { $fullName = $candidate.FullName } catch { }

            $isBlank = -not $fullName
            if ($isBlank) {
                try { $candidate.Saved = $true } catch { }
                try {
                    $candidate.Close()
                    Write-Verbose ("Closed blank presentation (index {0})." -f $index)
                } catch {
                    Write-Verbose ("Unable to close blank presentation (index {0}): {1}" -f $index, $_)
                }
            }
        }
    } catch {
        Write-Verbose ("Error while closing blank presentations: {0}" -f $_)
    }
}

function Open-PresentationWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        $Application,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$MaxAttempts = 3,
        [int]$DelayMilliseconds = 750
    )

    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return $Application.Presentations.Open($Path, 0, 0, 0)
        } catch {
            $lastError = $_
            Write-Verbose ("Attempt {0}/{1} to open {2} failed: {3}" -f $attempt, $MaxAttempts, $Path, $_)

            try {
                $pvCollection = $Application.ProtectedViewWindows
                if ($pvCollection -and $pvCollection.Count -gt 0) {
                    for ($pvIndex = 1; $pvIndex -le $pvCollection.Count; $pvIndex++) {
                        $pvWindow = $null
                        try { $pvWindow = $pvCollection.Item($pvIndex) } catch { continue }
                        if (-not $pvWindow) { continue }

                        $presentationPath = $null
                        try { $presentationPath = $pvWindow.PresentationPath } catch { }
                        $caption = $null
                        try { $caption = $pvWindow.Caption } catch { }

                        $matchesPath = $false
                        if ($presentationPath -and $presentationPath -eq $Path) {
                            $matchesPath = $true
                        } elseif ($caption -and $caption -like "*$([System.IO.Path]::GetFileName($Path))*") {
                            $matchesPath = $true
                        }

                        if ($matchesPath) {
                            $label = "<unknown protected view>"
                            if ($presentationPath) {
                                $label = $presentationPath
                            } elseif ($caption) {
                                $label = $caption
                            }
                            Write-Verbose ("Protected View detected for {0}; attempting to edit." -f $label)
                            try {
                                $null = $pvWindow.Edit()
                            } catch {
                                Write-Verbose ("Protected View edit failed: {0}" -f $_)
                            }
                        }
                    }
                }
            } catch { }

            if ($attempt -lt $MaxAttempts) {
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
        }
    }

    $message = "Failed to open $Path after $MaxAttempts attempts."
    if ($lastError) {
        $message += " Last error: $($lastError.Exception.Message)"
    }
    throw [System.InvalidOperationException]::new($message)
}

function Normalize-Label {
    param([string]$Text)
    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
}

function Reset-VerticalSpan {
    param($table, [int]$rowIdx, [int]$colIdx)

    $span = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
    $originalSpan = $span
    if ($span -le 1) { return }

    $iterations = 0
    while ($span -gt 1 -and $iterations -lt 64) {
        $iterations++
        $splitRows = if ($span -gt 2) { 2 } else { $span }
        try {
            $table.Cell($rowIdx, $colIdx).Split($splitRows, 1) | Out-Null
        } catch {
            Write-Verbose ("Unable to split vertical span on column {0} starting row {1}: {2}" -f $colIdx, $rowIdx, $_)
            break
        }
        $span = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
    }

    $clearRows = [Math]::Min($table.Rows.Count, $rowIdx + $originalSpan - 1)
    for ($targetRow = $rowIdx + 1; $targetRow -le $clearRows; $targetRow++) {
        try {
            $table.Cell($targetRow, $colIdx).Shape.TextFrame.TextRange.Text = ""
        } catch { }
    }
}

function Reset-HorizontalSpan {
    param($table, [int]$rowIdx, [int]$startCol, [int]$endCol)

    # Replaced by Ensure-FirstColumns logic; no-op placeholder to satisfy existing calls
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

function Reset-ColumnGroup {
    param($table, [int]$maxCols)

    if ($maxCols -lt 1) { return }

    for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
        Ensure-FirstColumns -table $table -rowIdx $rowIdx -maxCols $maxCols
    }
}

function Get-VerticalSpan {
    param($table, [int]$rowIdx, [int]$colIdx)

    try {
        $cell = $table.Cell($rowIdx, $colIdx)
    } catch {
        return 1
    }

    if ($null -eq $cell) { return 1 }

    $span = 1
    for ($cursor = $rowIdx + 1; $cursor -le $table.Rows.Count; $cursor++) {
        try {
            $next = $table.Cell($cursor, $colIdx)
        } catch {
            break
        }
        if (-not [object]::ReferenceEquals($cell.Shape, $next.Shape)) { break }
        $span += 1
    }
    return $span
}

function Get-HorizontalSpan {
    param($table, [int]$rowIdx, [int]$colIdx)

    try {
        $cell = $table.Cell($rowIdx, $colIdx)
    } catch {
        return 1
    }

    if ($null -eq $cell) { return 1 }

    $span = 1
    for ($cursor = $colIdx + 1; $cursor -le $table.Columns.Count; $cursor++) {
        try {
            $next = $table.Cell($rowIdx, $cursor)
        } catch {
            break
        }
        if (-not [object]::ReferenceEquals($cell.Shape, $next.Shape)) { break }
        $span += 1
    }
    return $span
}

function Ensure-FirstColumns {
    param(
        [object]$table,
        [int]$rowIdx,
        [int]$maxCols
    )

    $target = [Math]::Min($maxCols, $table.Columns.Count)
    if ($target -lt 1) { return }

    # Break vertical spans in the target columns
    for ($colIdx = 1; $colIdx -le $target; $colIdx++) {
        $vSpan = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
        if ($vSpan -gt 1) {
            try {
                $table.Cell($rowIdx, $colIdx).Split($vSpan, 1) | Out-Null
            } catch {
                Write-Verbose ("Unable to split vertical span on slide row {0} column {1}: {2}" -f $rowIdx, $colIdx, $_)
            }
        }
    }

    # Ensure column 1 is split into target columns (up to 3)
    $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx 1
    if ($span -gt $target) { $span = $target }
    if ($span -gt 1) {
        try {
            $table.Cell($rowIdx, 1).Split(1, $span) | Out-Null
        } catch {
            Write-Verbose ("Unable to split primary span on row {0}: {1}" -f $rowIdx, $_)
        }
    }

    # Ensure columns 2..target do not remain merged with column 1 or beyond target
    for ($colIdx = 2; $colIdx -le $target; $colIdx++) {
        $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
        if ($span -gt 1) {
            $split = [Math]::Min($span, $target - $colIdx + 1)
            if ($split -gt 1) {
                try {
                    $table.Cell($rowIdx, $colIdx).Split(1, $split) | Out-Null
                } catch {
                    Write-Verbose ("Unable to split secondary span on row {0} column {1}: {2}" -f $rowIdx, $colIdx, $_)
                }
            }
        }
        try {
            $table.Cell($rowIdx, $colIdx).Shape.TextFrame.TextRange.Text = ""
        } catch { }
    }
}

$pp = $null
$presentation = $null
try {
    Stop-PowerPointInstances

    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = 0 } catch { }
    try { $pp.DisplayAlerts = 1 } catch { }
    try { $pp.AutomationSecurity = 3 } catch { }

    Close-BlankPresentations -Application $pp

    $presentation = Open-PresentationWithRetry -Application $pp -Path $normalizedPath

    foreach ($slide in $presentation.Slides) {
        if ($SlideFilter -and -not ($SlideFilter -contains $slide.SlideIndex)) { continue }

        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -ne "MainDataTable" -or -not $shape.HasTable) { continue }

            $table = $shape.Table
            $rowCount = $table.Rows.Count
            if ($rowCount -le 1) { continue }

            $maxCampaignCol = [Math]::Min(3, $table.Columns.Count)
            Reset-ColumnGroup -table $table -maxCols $maxCampaignCol

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
