<#
.SYNOPSIS
    Post-processing script for PowerPoint presentations.

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED FOR BULK OPERATIONS ⚠️

    This script contains COM-based bulk operations that are PROHIBITED due to
    catastrophic performance issues (10+ hours vs 10 minutes with Python).

    **THIS SCRIPT IS SCHEDULED FOR MIGRATION TO PYTHON**

    See: docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Current Status (24 Oct 2025):
    - ALL merge operations: DISABLED (commented out)
    - ALL height enforcement: DISABLED (commented out)
    - File I/O operations: Still active (acceptable COM usage)

    Replacement: amp_automation/presentation/postprocess/cli.py

    DO NOT add new COM bulk operations to this script!
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [int[]]$SlideFilter,

    [bool]$PreSanitize = $true,
    [int[]]$SanitizeColumns = @(1, 2, 3),
    [int]$SanitizeHeaderRows = 1,
    [double]$SanitizeRowHeight = 8.4,
    [string]$TraceCommandLogPath,
    [string]$LogDirectory
)

$ErrorActionPreference = "Stop"

$script:ProjectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$script:PostProcessStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:DailyLogContext = $null
$script:MergeLogPath = $null
$script:TraceLogDefaultPath = $null
$script:WatchdogLogPath = $null
$script:CurrentSlideIndex = $null
$script:CurrentRowIndex = $null
$script:CurrentTableName = $null

function Get-ArabianNow {
    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Arabian Standard Time")
    } catch {
        $tz = [System.TimeZoneInfo]::CreateCustomTimeZone("UTC+04", [TimeSpan]::FromHours(4), "UTC+04", "UTC+04")
    }
    return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
}

function Resolve-DailyLogContext {
    param(
        [string]$ProjectRoot,
        [string]$OverrideDirectory
    )

    $context = [ordered]@{
        DailyDirectory = $null
        LogsDirectory  = $null
        MergeLogPath   = $null
        TraceLogPath   = $null
    }

    $docsRoot = Join-Path $ProjectRoot "docs"
    $nowLocal = Get-ArabianNow
    $preferredName = $nowLocal.ToString("dd-MM-yy")
    $altName = $nowLocal.ToString("dd-MM-yyyy")

    $candidate = $null
    if ($OverrideDirectory) {
        $resolvedOverride = [System.IO.Path]::GetFullPath($OverrideDirectory)
        if (Test-Path -LiteralPath $resolvedOverride) {
            $candidate = $resolvedOverride
        }
    }

    if (-not $candidate -and (Test-Path -LiteralPath $docsRoot)) {
        $candidate = Get-ChildItem -Path $docsRoot -Directory -Filter $preferredName -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1 |
            ForEach-Object { $_.FullName }

        if (-not $candidate) {
            $candidate = Get-ChildItem -Path $docsRoot -Directory -Filter $altName -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1 |
                ForEach-Object { $_.FullName }
        }
    }

    if (-not $candidate) {
        if (-not (Test-Path -LiteralPath $docsRoot)) {
            [void](New-Item -ItemType Directory -Path $docsRoot -Force)
        }
        $candidate = Join-Path $docsRoot $preferredName
    }

    if (-not (Test-Path -LiteralPath $candidate)) {
        [void](New-Item -ItemType Directory -Path $candidate -Force)
    }

    $logsDir = Join-Path $candidate "logs"
    if (-not (Test-Path -LiteralPath $logsDir)) {
        [void](New-Item -ItemType Directory -Path $logsDir -Force)
    }

    $context.DailyDirectory = $candidate
    $context.LogsDirectory = $logsDir
    $context.MergeLogPath = Join-Path $logsDir "05-postprocess_merges.log"
    $context.TraceLogPath = Join-Path $logsDir "05-postprocess_trace.log"

    return [PSCustomObject]$context
}

function Initialize-WatchdogContext {
    param(
        [string]$PresentationFullPath,
        [string]$ProjectRoot
    )

    if (-not $PresentationFullPath) { return $null }

    $match = [regex]::Match($PresentationFullPath, "run_\d{8}_\d{6}", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return $null }

    $productionRoot = Join-Path $ProjectRoot "logs"
    $productionRoot = Join-Path $productionRoot "production"
    if (-not (Test-Path -LiteralPath $productionRoot)) {
        try { [void](New-Item -ItemType Directory -Path $productionRoot -Force) } catch { return $null }
    }

    $runDirectory = Join-Path $productionRoot $match.Value
    if (-not (Test-Path -LiteralPath $runDirectory)) {
        try { [void](New-Item -ItemType Directory -Path $runDirectory -Force) } catch { return $null }
    }

    return Join-Path $runDirectory "postprocess_watchdog.log"
}

function Write-WatchdogLog {
    param([string]$Message)

    if (-not $script:WatchdogLogPath) { return }
    if (-not $Message) { return }

    $timestamp = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $line = "[{0}] {1}" -f $timestamp, $Message
    try {
        Add-Content -Path $script:WatchdogLogPath -Value $line -Encoding UTF8
    } catch { }
}

function Write-MergeLog {
    param(
        [string]$Message,
        [int]$SlideIndex,
        [int]$RowIndex,
        [switch]$IsError
    )

    if (-not $Message) { return }

    if (-not $PSBoundParameters.ContainsKey("SlideIndex") -and $script:CurrentSlideIndex) {
        $SlideIndex = $script:CurrentSlideIndex
    }

    if (-not $PSBoundParameters.ContainsKey("RowIndex") -and $script:CurrentRowIndex) {
        $RowIndex = $script:CurrentRowIndex
    }

    $elapsed = if ($script:PostProcessStopwatch) {
        $script:PostProcessStopwatch.Elapsed
    } else {
        [TimeSpan]::Zero
    }
    $elapsedLabel = $elapsed.ToString("hh\:mm\:ss\.fff")

    $contextParts = @("[{0}]" -f $elapsedLabel)
    if ($SlideIndex) { $contextParts += ("slide {0}" -f $SlideIndex) }
    if ($RowIndex) { $contextParts += ("row {0}" -f $RowIndex) }
    if ($script:CurrentTableName) { $contextParts += $script:CurrentTableName }

    $level = if ($IsError) { "ERROR" } else { "INFO" }
    $prefix = ($contextParts -join " | ")
    $line = "{0} {1}" -f $prefix, $Message

    if ($script:MergeLogPath) {
        $timestamp = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        $entry = "[{0}] {1} {2}" -f $timestamp, $level, $line
        try {
            Add-Content -Path $script:MergeLogPath -Value $entry -Encoding UTF8
        } catch { }
    }

    Write-Verbose $line
}

$normalizedPath = [System.IO.Path]::GetFullPath($PresentationPath).Replace('/', '\')

if (-not (Test-Path -LiteralPath $normalizedPath)) {
    throw "Presentation not found: $normalizedPath"
}

$script:DailyLogContext = Resolve-DailyLogContext -ProjectRoot $script:ProjectRoot -OverrideDirectory $LogDirectory
if ($script:DailyLogContext) {
    $script:MergeLogPath = $script:DailyLogContext.MergeLogPath
    if (-not $TraceCommandLogPath) {
        $TraceCommandLogPath = $script:DailyLogContext.TraceLogPath
    } else {
        $TraceCommandLogPath = [System.IO.Path]::GetFullPath($TraceCommandLogPath)
        $traceParent = Split-Path -Parent $TraceCommandLogPath
        if ($traceParent -and -not (Test-Path -LiteralPath $traceParent)) {
            [void](New-Item -ItemType Directory -Path $traceParent -Force)
        }
    }
}
$script:TraceLogDefaultPath = $TraceCommandLogPath
$script:WatchdogLogPath = Initialize-WatchdogContext -PresentationFullPath $normalizedPath -ProjectRoot $script:ProjectRoot

Write-MergeLog -Message ("Using presentation path: {0}" -f $normalizedPath)
if ($TraceCommandLogPath) {
    Write-MergeLog -Message ("Trace log path configured: {0}" -f $TraceCommandLogPath)
}
if ($script:WatchdogLogPath) {
    Write-MergeLog -Message ("Watchdog log path: {0}" -f $script:WatchdogLogPath)
}

$unblockSucceeded = $false
try {
    Unblock-File -LiteralPath $normalizedPath -ErrorAction Stop
    $unblockSucceeded = $true
} catch {
    Write-MergeLog -Message ("Unblock-File skipped for {0}: {1}" -f $normalizedPath, $_)
}
if ($unblockSucceeded) {
    Write-MergeLog -Message ("Removed zone identifier for: {0}" -f $normalizedPath)
}

if ($PreSanitize) {
    $sanitizeScript = Join-Path -Path $PSScriptRoot -ChildPath "SanitizePrimaryColumns.ps1"
    if (-not (Test-Path -LiteralPath $sanitizeScript)) {
        throw "SanitizePrimaryColumns.ps1 not found in $PSScriptRoot"
    }

    $sanitizeArgs = @{
        PresentationPath = $normalizedPath
        Columns = $SanitizeColumns
        HeaderRows = $SanitizeHeaderRows
        RowHeightPt = $SanitizeRowHeight
    }

    if ($SlideFilter) {
        $sanitizeArgs["SlideIndexFilter"] = $SlideFilter
    }

    if ($VerbosePreference -eq 'Continue') {
        $sanitizeArgs["Verbose"] = $true
    }

    Write-MergeLog -Message "Running primary column sanitizer before merge processing."
    & $sanitizeScript @sanitizeArgs
}

$campaignFontSize = 6.0
$monthlyTotalFontSize = 6.5
$summaryFontSize = 7.0
$blankFontName = "Verdana"
$blankFontSize = 6.0
$blankFontColor = 0x00BFBFBF
$zeroWidthSpace = [char]0x200B

function Set-CellFixedLayout {
    param($cell)

    if (-not $cell) { return }

    $shape = $null
    try { $shape = $cell.Shape } catch { return }
    if (-not $shape) { return }

    $textFrame = $null
    try { $textFrame = $shape.TextFrame } catch { }
    if ($textFrame) {
        try { $textFrame.AutoSize = 0 } catch { }
        try { $textFrame.WordWrap = -1 } catch { }
        try { $textFrame.MarginTop = 0 } catch { }
        try { $textFrame.MarginBottom = 0 } catch { }
        try { $textFrame.MarginLeft = 0 } catch { }
        try { $textFrame.MarginRight = 0 } catch { }
        try { $textFrame.VerticalAnchor = 3 } catch { }
    }

    $textFrame2 = $null
    try { $textFrame2 = $shape.TextFrame2 } catch { }
    if ($textFrame2) {
        try { $textFrame2.AutoSize = 0 } catch { }
        try { $textFrame2.MarginTop = 0 } catch { }
        try { $textFrame2.MarginBottom = 0 } catch { }
        try { $textFrame2.MarginLeft = 0 } catch { }
        try { $textFrame2.MarginRight = 0 } catch { }
        try { $textFrame2.VerticalAnchor = 3 } catch { }
    }
}

function Normalize-TableLayout {
    param($table)

    if (-not $table) { return }

    for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
        for ($colIdx = 1; $colIdx -le $table.Columns.Count; $colIdx++) {
            try {
                $cell = $table.Cell($rowIdx, $colIdx)
                Set-CellFixedLayout -cell $cell
            } catch { }
        }
    }
}

function Stop-PowerPointInstances {
    param(
        [int]$WaitMilliseconds = 500,
        [switch]$ForceFirst
    )

    if ($ForceFirst) {
        Write-MergeLog -Message "Forcefully terminating any POWERPNT.exe instances before run."
        try {
            Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
        } catch { }

        if ($WaitMilliseconds -gt 0) {
            $sleep = [Math]::Min([Math]::Max($WaitMilliseconds, 100), 1000)
            Start-Sleep -Milliseconds $sleep
        }
    }

    $existing = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if (-not $existing) { return }

    Write-MergeLog -Message ("Gracefully stopping {0} lingering POWERPNT processes ({1})." -f $existing.Count, ($existing | Select-Object -ExpandProperty Id | Sort-Object | ForEach-Object { $_ } -join ", "))
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

    $postCheck = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($postCheck) {
        Write-MergeLog -Message ("POWERPNT cleanup incomplete; remaining process IDs: {0}" -f ($postCheck | Select-Object -ExpandProperty Id | Sort-Object -Unique -join ", ")) -IsError
    } else {
        Write-MergeLog -Message "POWERPNT cleanup complete; no instances detected."
    }

    try {
        $tasklist = & tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>$null
        if ($tasklist -match "No tasks are running") {
            Write-MergeLog -Message "tasklist verification: no POWERPNT.EXE entries reported."
        } else {
            Write-MergeLog -Message ("tasklist verification detected POWERPNT entries:`n{0}" -f ($tasklist -join [Environment]::NewLine)) -IsError
        }
    } catch {
        Write-MergeLog -Message ("tasklist verification failed: {0}" -f $_) -IsError
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
                    Write-MergeLog -Message ("Closed blank presentation (index {0})." -f $index)
                } catch {
                    Write-MergeLog -Message ("Unable to close blank presentation (index {0}): {1}" -f $index, $_)
                }
            }
        }
    } catch {
        Write-MergeLog -Message ("Error while closing blank presentations: {0}" -f $_)
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
            Write-MergeLog -Message ("Attempt {0}/{1} to open {2} failed: {3}" -f $attempt, $MaxAttempts, $Path, $_)

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
                            Write-MergeLog -Message ("Protected View detected for {0}; attempting to edit." -f $label)
                            try {
                                $null = $pvWindow.Edit()
                            } catch {
                                Write-MergeLog -Message ("Protected View edit failed: {0}" -f $_)
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

function Normalize-CellContent {
    param([string]$Text)

    if ($null -eq $Text) { return "" }

    $processed = $Text -replace "\u200B", ""
    $processed = $processed -replace "[`r`n]+", " "
    return $processed.Trim()
}

function Ensure-BlankCellFormatting {
    param($cell)

    if (-not $cell) { return }

    Set-CellFixedLayout -cell $cell

    $textFrame = $null
    try { $textFrame = $cell.Shape.TextFrame } catch { return }
    if (-not $textFrame) { return }

    $range = $null
    try { $range = $textFrame.TextRange } catch { return }
    if (-not $range) { return }

    $content = Normalize-CellContent -Text $range.Text
    $isDash = $content -eq "-"
    $isBlank = $content.Length -eq 0
    if (-not $isDash -and -not $isBlank) { return }

    if ($isBlank) {
        try {
            $range.Text = [string]$script:zeroWidthSpace
            $range = $textFrame.TextRange
        } catch {
            return
        }
    }

    try { $textFrame.AutoSize = 0 } catch { }
    try { $textFrame.MarginLeft = 0 } catch { }
    try { $textFrame.MarginRight = 0 } catch { }
    try { $textFrame.VerticalAnchor = 3 } catch { }
    try { $range.ParagraphFormat.Alignment = 2 } catch { }

    $font = $null
    try { $font = $range.Font } catch { }
    if ($font) {
        try { $font.Name = $script:blankFontName } catch { }
        try { $font.Size = $script:blankFontSize } catch { }
        try { $font.Bold = 0 } catch { }
        try { $font.Color.RGB = $script:blankFontColor } catch { }
    }
}

function Apply-BlankCellFormatting {
    param($table)

    if (-not $table) { return }

    for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
        for ($colIdx = 1; $colIdx -le $table.Columns.Count; $colIdx++) {
            $cell = $null
            try { $cell = $table.Cell($rowIdx, $colIdx) } catch { continue }
            Ensure-BlankCellFormatting -cell $cell
        }
    }
}

function Reset-VerticalSpan {
    param($table, [int]$rowIdx, [int]$colIdx)

    $span = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
    $originalSpan = $span
    if ($span -le 1) { return }

    $iterations = 0
    $stallCount = 0
    while ($span -gt 1 -and $iterations -lt 64) {
        $iterations++
        $splitRows = if ($span -gt 2) { 2 } else { $span }
        try {
            $table.Cell($rowIdx, $colIdx).Split($splitRows, 1) | Out-Null
        } catch {
            Write-MergeLog -Message ("Unable to split vertical span on column {0} starting row {1}: {2}" -f $colIdx, $rowIdx, $_) -IsError
            break
        }
        $newSpan = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
        if ($newSpan -ge $span) {
            $stallCount++
        } else {
            $stallCount = 0
        }
        $span = $newSpan
        if ($stallCount -ge 4) {
            $message = "Watchdog: vertical span reset stalled at column {0} starting row {1} (span {2}, iteration {3})." -f $colIdx, $rowIdx, $span, $iterations
            Write-MergeLog -Message $message -RowIndex $rowIdx -IsError
            Write-WatchdogLog $message
            break
        }
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
    param(
        $table,
        [int]$rowIdx,
        [int]$maxCols = 3
    )

    Ensure-FirstColumns -table $table -rowIdx $rowIdx -maxCols $maxCols

    try {
        $primaryCell = $table.Cell($rowIdx, 1)
        Set-CellFixedLayout -cell $primaryCell
    } catch { }
}

function Reset-PrimaryColumnSpans {
    param(
        [object]$table,
        [int]$maxCols
    )

    if (-not $table) { return }

    $limit = [Math]::Min($maxCols, $table.Columns.Count)
    if ($limit -lt 1) { return }

    for ($colIdx = 1; $colIdx -le $limit; $colIdx++) {
        for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
            $cell = $null
            try { $cell = $table.Cell($rowIdx, $colIdx) } catch { continue }
            if (-not $cell) { continue }

            if ($rowIdx -gt 1) {
                try {
                    $above = $table.Cell($rowIdx - 1, $colIdx)
                    if ($above -and [object]::ReferenceEquals($cell.Shape, $above.Shape)) { continue }
                } catch { }
            }

            $hSpan = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
            $hIteration = 0
            $hStall = 0
            while ($hSpan -gt 1 -and $hIteration -lt 50) {
                $hIteration++
                $maxAvailable = $table.Columns.Count - $colIdx + 1
                if ($maxAvailable -lt 2) { break }
                $splitCols = [Math]::Min($hSpan, $maxAvailable)
                if ($splitCols -gt $limit) {
                    $splitCols = $limit
                }
                if ($splitCols -gt 3) {
                    $splitCols = 3
                }
                if ($splitCols -le 1) { break }
                try {
                    $cell.Split(1, $splitCols) | Out-Null
                } catch {
                    if ($splitCols -gt 2) {
                        $splitCols = 2
                        continue
                    }
                    Write-MergeLog -Message ("Unable to split horizontal span (row {0}, col {1}, span {2}): {3}" -f $rowIdx, $colIdx, $hSpan, $_) -IsError
                    break
                }
                try { $cell = $table.Cell($rowIdx, $colIdx) } catch { break }
                $newHSpan = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                if ($newHSpan -ge $hSpan) {
                    $hStall++
                } else {
                    $hStall = 0
                }
                $hSpan = $newHSpan
                if ($hStall -ge 4) {
                    $message = "Watchdog: horizontal span reset stalled (row {0}, col {1}, span {2}, iteration {3})." -f $rowIdx, $colIdx, $hSpan, $hIteration
                    Write-MergeLog -Message $message -RowIndex $rowIdx -IsError
                    Write-WatchdogLog $message
                    break
                }
            }

            $span = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
            $iteration = 0
            $stallCount = 0
            while ($span -gt 1 -and $iteration -lt 400) {
                $iteration++
                $splitRows = [Math]::Min($span, 50)
                if ($splitRows -lt 2) {
                    $splitRows = 2
                }
                if ($splitRows -gt 75) {
                    $splitRows = 75
                }

                try {
                    $cell.Split($splitRows, 1) | Out-Null
                } catch {
                    if ($splitRows -gt 2) {
                        $splitRows = 2
                        continue
                    }
                    Write-MergeLog -Message ("Unable to split vertical span (row {0}, col {1}, span {2}): {3}" -f $rowIdx, $colIdx, $span, $_) -IsError
                    break
                }

                try { $cell = $table.Cell($rowIdx, $colIdx) } catch { break }
                $newSpan = Get-VerticalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                if ($newSpan -ge $span) {
                    $stallCount++
                } else {
                    $stallCount = 0
                }
                $span = $newSpan
                if ($stallCount -ge 6) {
                    $message = "Watchdog: vertical span reset stalled (row {0}, col {1}, span {2}, iteration {3})." -f $rowIdx, $colIdx, $span, $iteration
                    Write-MergeLog -Message $message -RowIndex $rowIdx -IsError
                    Write-WatchdogLog $message
                    break
                }
            }

            Set-CellFixedLayout -cell $cell

            if ($span -gt 1) {
                Write-MergeLog -Message ("Residual vertical span remains on row {0} column {1} (span {2}) after reset." -f $rowIdx, $colIdx, $span)
            }
        }
    }

    for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
        for ($colIdx = 1; $colIdx -le $limit; $colIdx++) {
            $cell = $null
            try { $cell = $table.Cell($rowIdx, $colIdx) } catch { continue }
            if (-not $cell) { continue }

            if ($colIdx -gt 1) {
                try {
                    $left = $table.Cell($rowIdx, $colIdx - 1)
                    if ($left -and [object]::ReferenceEquals($cell.Shape, $left.Shape)) { continue }
                } catch { }
            }

            $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
            $iteration = 0
            $stallCount = 0
            while ($span -gt 1 -and $iteration -lt 400) {
                $iteration++
                $maxAvailable = $table.Columns.Count - $colIdx + 1
                if ($maxAvailable -lt 2) { break }
                $splitCols = [Math]::Min($span, $maxAvailable)
                if ($splitCols -gt 3) {
                    $splitCols = 3
                }
                if ($splitCols -le 1) { break }

                try {
                    $cell.Split(1, $splitCols) | Out-Null
                } catch {
                    if ($splitCols -gt 2) {
                        $splitCols = 2
                        continue
                    }
                    Write-MergeLog -Message ("Unable to split horizontal span (row {0}, col {1}, span {2}): {3}" -f $rowIdx, $colIdx, $span, $_) -IsError
                    break
                }

                try { $cell = $table.Cell($rowIdx, $colIdx) } catch { break }
                $newSpan = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
                if ($newSpan -ge $span) {
                    $stallCount++
                } else {
                    $stallCount = 0
                }
                $span = $newSpan
                if ($stallCount -ge 6) {
                    $message = "Watchdog: horizontal span reset stalled (row {0}, col {1}, span {2}, iteration {3})." -f $rowIdx, $colIdx, $span, $iteration
                    Write-MergeLog -Message $message -RowIndex $rowIdx -IsError
                    Write-WatchdogLog $message
                    break
                }
            }

            Set-CellFixedLayout -cell $cell

            if ($span -gt 1) {
                Write-MergeLog -Message ("Residual horizontal span remains on row {0} column {1} (span {2}) after reset." -f $rowIdx, $colIdx, $span)
            }
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

function Set-RowHeightExact {
    param(
        [object]$table,
        [int]$rowIdx,
        [double]$height,
        [int]$maxAttempts = 6,
        [double]$acceptableTolerance = 1.5
    )

    try {
        $row = $table.Rows.Item($rowIdx)
    } catch {
        Write-MergeLog -Message ("Unable to access row {0} for height enforcement: {1}" -f $rowIdx, $_)
        return
    }

    try { $row.HeightRule = 1 } catch { }

    $adjustments = @(0.0, 0.12, -0.12, 0.08, -0.08, 0.04)
    $current = $height
    $previous = $null
    $plateauCount = 0

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $offset = $adjustments[($attempt - 1) % $adjustments.Count]
        $targetValue = $height + $offset

        try {
            $row.Height = $targetValue
        } catch {
            Write-MergeLog -Message ("Unable to set row height on row {0} attempt {1}: {2}" -f $rowIdx, $attempt, $_)
            return
        }

        Start-Sleep -Milliseconds 20

        try {
            $current = [double]$row.Height
        } catch {
            $current = $height
        }

        if ([math]::Abs($current - $height) -le 0.05) {
            if ([math]::Abs($targetValue - $height) -gt 0.001) {
                try { $row.Height = $height } catch { }
            }
            return
        }

        # Plateau detection: check if height is stuck
        if ($previous -ne $null) {
            $heightDelta = [math]::Abs($current - $previous)
            if ($heightDelta -le 0.02) {
                $plateauCount++
            } else {
                $plateauCount = 0
            }

            # If plateaued for 2+ attempts, check acceptable tolerance
            if ($plateauCount -ge 2) {
                $deviation = [math]::Abs($current - $height)
                if ($deviation -le $acceptableTolerance) {
                    Write-MergeLog -Message ("Row {0} height plateaued at {1} (target {2}, within {3}pt tolerance); accepting" -f $rowIdx, [math]::Round($current, 3), $height, $acceptableTolerance)
                    return
                }
                # Beyond acceptable tolerance, break early to avoid wasting time
                Write-MergeLog -Message ("Row {0} height plateaued at {1} (target {2}, exceeds tolerance); stopping retries" -f $rowIdx, [math]::Round($current, 3), $height)
                break
            }
        }
        $previous = $current
    }

    try { $row.Height = $height } catch { }
    Start-Sleep -Milliseconds 20
    try {
        $current = [double]$row.Height
    } catch {
        $current = $height
    }
    if ([math]::Abs($current - $height) -gt 0.05) {
        Write-MergeLog -Message ("Row {0} height remains {1} (target {2}); continuing" -f $rowIdx, [math]::Round($current, 3), $height)
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
                Write-MergeLog -Message ("Unable to split vertical span on slide row {0} column {1}: {2}" -f $rowIdx, $colIdx, $_)
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
            Write-MergeLog -Message ("Unable to split primary span on row {0}: {1}" -f $rowIdx, $_)
        }
    }
    try {
        $primaryCell = $table.Cell($rowIdx, 1)
        Set-CellFixedLayout -cell $primaryCell
    } catch { }

    # Ensure columns 2..target do not remain merged with column 1 or beyond target
    for ($colIdx = 2; $colIdx -le $target; $colIdx++) {
        $span = Get-HorizontalSpan -table $table -rowIdx $rowIdx -colIdx $colIdx
        if ($span -gt 1) {
            $split = [Math]::Min($span, $target - $colIdx + 1)
            if ($split -gt 1) {
                try {
                    $table.Cell($rowIdx, $colIdx).Split(1, $split) | Out-Null
                } catch {
                    Write-MergeLog -Message ("Unable to split secondary span on row {0} column {1}: {2}" -f $rowIdx, $colIdx, $_)
                }
            }
        }
        try {
            $targetCell = $table.Cell($rowIdx, $colIdx)
            $targetCell.Shape.TextFrame.TextRange.Text = ""
            Set-CellFixedLayout -cell $targetCell
            Ensure-BlankCellFormatting -cell $targetCell
        } catch { }
    }
}

function Prepare-RowGroup {
    param(
        [object]$table,
        [int]$startRow,
        [int]$endRow,
        [int]$maxCols,
        [double]$targetHeight
    )

    if ($null -eq $startRow -or $null -eq $endRow) { return }
    if ($endRow -lt $startRow) { return }

    $limit = [Math]::Min($maxCols, $table.Columns.Count)
    if ($limit -lt 1) { return }

    for ($rowIdx = $startRow; $rowIdx -le $endRow; $rowIdx++) {
        Ensure-FirstColumns -table $table -rowIdx $rowIdx -maxCols $limit
        for ($colIdx = 1; $colIdx -le $table.Columns.Count; $colIdx++) {
            try {
                $cell = $table.Cell($rowIdx, $colIdx)
                Set-CellFixedLayout -cell $cell
            } catch { }
        }
        # Height enforcement disabled for performance
        # Set-RowHeightExact -table $table -rowIdx $rowIdx -height $targetHeight
    }
}

function Invoke-PostProcessMergeCore {
    $pp = $null
    $presentation = $null
    try {
        Write-MergeLog -Message "Pre-run cleanup: terminating stray PowerPoint sessions."
        Stop-PowerPointInstances -ForceFirst -WaitMilliseconds 750
        Write-MergeLog -Message "Creating PowerPoint COM automation session."
        $pp = New-Object -ComObject PowerPoint.Application
        try { $pp.Visible = 0 } catch { Write-MergeLog -Message ("Unable to set PowerPoint visibility: {0}" -f $_) -IsError }
        try { $pp.DisplayAlerts = 1 } catch { Write-MergeLog -Message ("Unable to set DisplayAlerts: {0}" -f $_) -IsError }
        try { $pp.AutomationSecurity = 3 } catch { Write-MergeLog -Message ("Unable to set AutomationSecurity: {0}" -f $_) -IsError }

        Close-BlankPresentations -Application $pp

        Write-MergeLog -Message ("Opening presentation: {0}" -f $normalizedPath)
        $presentation = Open-PresentationWithRetry -Application $pp -Path $normalizedPath
        Write-MergeLog -Message ("Presentation opened; slide count: {0}" -f $presentation.Slides.Count)

        foreach ($slide in $presentation.Slides) {
            if ($SlideFilter -and -not ($SlideFilter -contains $slide.SlideIndex)) { continue }

            $script:CurrentSlideIndex = $slide.SlideIndex
            $script:CurrentTableName = $null

            $slideWatch = [System.Diagnostics.Stopwatch]::StartNew()
            Write-MergeLog -Message ("Processing slide {0} with {1} shapes." -f $slide.SlideIndex, $slide.Shapes.Count)

        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -ne "MainDataTable" -or -not $shape.HasTable) { continue }

            $script:CurrentTableName = $shape.Name

            $table = $shape.Table
            $rowCount = $table.Rows.Count
            $columnCount = $table.Columns.Count
            Write-MergeLog -Message ("Processing table {0}; rows={1}, columns={2}." -f $shape.Name, $rowCount, $columnCount)

            if ($rowCount -le 1) {
                Write-MergeLog -Message "Skipping table with <= 1 row."
                $script:CurrentTableName = $null
                continue
            }

            $maxCampaignCol = [Math]::Min(3, $columnCount)
            Normalize-TableLayout -table $table
            Reset-PrimaryColumnSpans -table $table -maxCols $maxCampaignCol
            Reset-ColumnGroup -table $table -maxCols $maxCampaignCol
            Apply-BlankCellFormatting -table $table
            $targetHeight = 8.4
            # Height enforcement disabled for performance - merges complete faster without it
            # Write-MergeLog -Message ("Enforcing target row height {0}pt." -f $targetHeight)
            # for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
            #     $script:CurrentRowIndex = $rowIdx
            #     Set-RowHeightExact -table $table -rowIdx $rowIdx -height $targetHeight
            # }
            $script:CurrentRowIndex = $null

            $campaignStart = $null
            for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
                $script:CurrentRowIndex = $rowIdx
                if ($slideWatch.Elapsed.TotalMinutes -ge 4) {
                    $timeoutMessage = "Timeout reached on slide {0} at row {1}; skipping remaining rows." -f $slide.SlideIndex, $rowIdx
                    Write-MergeLog -Message $timeoutMessage -IsError
                    Write-WatchdogLog $timeoutMessage
                    $script:CurrentRowIndex = $null
                    break
                }

                try {
                    $labelCell = $table.Cell($rowIdx, 1)
                } catch {
                    Write-MergeLog -Message ("Unable to read label cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_) -IsError
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

                # Campaign merges disabled for performance - will be rebuilt in Python
                # if ($isMonthlyTotal -and $campaignStart -ne $null) {
                #     $campaignEnd = $rowIdx - 1
                #     if ($campaignEnd -gt $campaignStart) {
                #         Prepare-RowGroup -table $table -startRow $campaignStart -endRow $campaignEnd -maxCols $maxCampaignCol -targetHeight $targetHeight
                #
                #         $merged = $false
                #         for ($attempt = 1; $attempt -le 2 -and -not $merged; $attempt++) {
                #             try {
                #                 $topCell = $table.Cell($campaignStart, 1)
                #                 $bottomCell = $table.Cell($campaignEnd, 1)
                #                 if (-not [object]::ReferenceEquals($topCell.Shape, $bottomCell.Shape)) {
                #                     $null = $topCell.Merge($bottomCell)
                #                 }
                #                 $merged = $true
                #             } catch {
                #                 Write-MergeLog -Message ("Campaign merge failed on slide {0} rows {1}-{2} attempt {3}: {4}" -f $slide.SlideIndex, $campaignStart, $campaignEnd, $attempt, $_) -IsError
                #                 Write-WatchdogLog ("Campaign merge retry needed on slide {0} rows {1}-{2}." -f $slide.SlideIndex, $campaignStart, $campaignEnd)
                #                 Prepare-RowGroup -table $table -startRow $campaignStart -endRow $campaignEnd -maxCols $maxCampaignCol -targetHeight $targetHeight
                #                 Start-Sleep -Milliseconds 50
                #             }
                #         }
                #
                #         try {
                #             $campaignCell = $table.Cell($campaignStart, 1)
                #             Set-CellFixedLayout -cell $campaignCell
                #             $campaignRange = $campaignCell.Shape.TextFrame.TextRange
                #             if (-not (Normalize-Label -Text $campaignRange.Text)) { $campaignRange.Text = $normalized }
                #             $campaignRange.ParagraphFormat.Alignment = 2
                #             $campaignRange.Parent.VerticalAnchor = 3
                #             $campaignRange.Font.Size = $campaignFontSize
                #             $campaignRange.Font.Bold = -1
                #         } catch {
                #             Write-MergeLog -Message ("Unable to style campaign cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $campaignStart, $_) -IsError
                #         }
                #     }
                #
                #     $campaignStart = $null
                # }
                if ($isMonthlyTotal -and $campaignStart -ne $null) {
                    $campaignStart = $null
                }

                # Monthly total merges disabled for performance - will be rebuilt in Python
                # if ($isMonthlyTotal) {
                #     try {
                #         Reset-MonthlyTotalRow -table $table -rowIdx $rowIdx -maxCols $maxCampaignCol
                #         Prepare-RowGroup -table $table -startRow $rowIdx -endRow $rowIdx -maxCols $maxCampaignCol -targetHeight $targetHeight
                #
                #         $mergedCell = $table.Cell($rowIdx, 1)
                #         $maxMergeCol = [Math]::Min(3, $table.Columns.Count)
                #         for ($targetCol = 2; $targetCol -le $maxMergeCol; $targetCol++) {
                #             $mergeCompleted = $false
                #             for ($attempt = 1; $attempt -le 2 -and -not $mergeCompleted; $attempt++) {
                #                 try {
                #                     $mergedCell = $table.Cell($rowIdx, 1)
                #                     $targetCell = $table.Cell($rowIdx, $targetCol)
                #                 } catch {
                #                     Write-MergeLog -Message ("Unable to access monthly total cells on slide {0} row {1} column {2}: {3}" -f $slide.SlideIndex, $rowIdx, $targetCol, $_)
                #                     break
                #                 }
                #
                #                 if ([object]::ReferenceEquals($mergedCell.Shape, $targetCell.Shape)) {
                #                     $mergeCompleted = $true
                #                     continue
                #                 }
                #
                #                 try {
                #                     $null = $mergedCell.Merge($targetCell)
                #                     $mergeCompleted = $true
                #             } catch {
                #                 Write-MergeLog -Message ("Monthly total merge failed on slide {0} row {1} column {2} attempt {3}: {4}" -f $slide.SlideIndex, $rowIdx, $targetCol, $attempt, $_) -IsError
                #                 Write-WatchdogLog ("Monthly total merge retry required on slide {0} row {1} column {2}." -f $slide.SlideIndex, $rowIdx, $targetCol)
                #                 Prepare-RowGroup -table $table -startRow $rowIdx -endRow $rowIdx -maxCols $maxCampaignCol -targetHeight $targetHeight
                #                 Start-Sleep -Milliseconds 50
                #             }
                #         }
                #         }
                #
                #         $mergedCell = $table.Cell($rowIdx, 1)
                #         Set-CellFixedLayout -cell $mergedCell
                #         $monthlyRange = $mergedCell.Shape.TextFrame.TextRange
                #         $monthlyRange.Text = $normalized
                #         $monthlyRange.ParagraphFormat.Alignment = 2
                #         $monthlyRange.Parent.VerticalAnchor = 3
                #         $monthlyRange.Font.Size = $monthlyTotalFontSize
                #         $monthlyRange.Font.Bold = -1
                #     } catch {
                #         Write-MergeLog -Message ("Monthly total merge failed on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_) -IsError
                #     }
                # }

                # Summary merges (GRAND TOTAL/CARRIED FORWARD) disabled for performance - will be rebuilt in Python
                # if ($isGrandTotal -or $isCarriedForward) {
                #     try {
                #         Prepare-RowGroup -table $table -startRow $rowIdx -endRow $rowIdx -maxCols $maxCampaignCol -targetHeight $targetHeight
                #
                #         $summaryCell = $table.Cell($rowIdx, 1)
                #         $maxMergeCol = [Math]::Min(3, $table.Columns.Count)
                #         for ($targetCol = 2; $targetCol -le $maxMergeCol; $targetCol++) {
                #             $mergeCompleted = $false
                #             for ($attempt = 1; $attempt -le 2 -and -not $mergeCompleted; $attempt++) {
                #                 try {
                #                     $summaryCell = $table.Cell($rowIdx, 1)
                #                     $targetCell = $table.Cell($rowIdx, $targetCol)
                #                 } catch {
                #                     Write-MergeLog -Message ("Unable to access summary cells on slide {0} row {1} column {2}: {3}" -f $slide.SlideIndex, $rowIdx, $targetCol, $_) -IsError
                #                     break
                #                 }
                #
                #                 if ([object]::ReferenceEquals($summaryCell.Shape, $targetCell.Shape)) {
                #                     $mergeCompleted = $true
                #                     continue
                #                 }
                #
                #                 try {
                #                     $null = $summaryCell.Merge($targetCell)
                #                     $mergeCompleted = $true
                #                 } catch {
                #                     Write-MergeLog -Message ("Summary merge failed on slide {0} row {1} column {2} attempt {3}: {4}" -f $slide.SlideIndex, $rowIdx, $targetCol, $attempt, $_) -IsError
                #                     Write-WatchdogLog ("Summary merge retry required on slide {0} row {1} column {2}." -f $slide.SlideIndex, $rowIdx, $targetCol)
                #                     Prepare-RowGroup -table $table -startRow $rowIdx -endRow $rowIdx -maxCols $maxCampaignCol -targetHeight $targetHeight
                #                     Start-Sleep -Milliseconds 50
                #                 }
                #             }
                #         }
                #
                #         Set-CellFixedLayout -cell $summaryCell
                #         $summaryRange = $summaryCell.Shape.TextFrame.TextRange
                #         $summaryRange.Text = $normalized
                #         $summaryRange.Font.Size = $summaryFontSize
                #         $summaryRange.Font.Bold = -1
                #         $summaryRange.ParagraphFormat.Alignment = 2
                #         $summaryRange.Parent.VerticalAnchor = 3
                #     } catch {
                #         Write-MergeLog -Message ("Unable to style summary cell on slide {0} row {1}: {2}" -f $slide.SlideIndex, $rowIdx, $_) -IsError
                #     }
                # }
            }

            Apply-BlankCellFormatting -table $table

            # Height enforcement disabled for performance
            # for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
            #     $script:CurrentRowIndex = $rowIdx
            #     Set-RowHeightExact -table $table -rowIdx $rowIdx -height $targetHeight
            # }

            Apply-BlankCellFormatting -table $table
            # Height enforcement disabled for performance
            # for ($rowIdx = 2; $rowIdx -le $rowCount; $rowIdx++) {
            #     $script:CurrentRowIndex = $rowIdx
            #     Set-RowHeightExact -table $table -rowIdx $rowIdx -height $targetHeight
            # }
            $script:CurrentRowIndex = $null
            Write-MergeLog -Message ("Completed table {0} processing." -f $shape.Name)
            $script:CurrentTableName = $null
        }

        $slideWatch.Stop()
        Write-MergeLog -Message ("Slide {0} processing finished in {1}." -f $slide.SlideIndex, $slideWatch.Elapsed.ToString("hh\:mm\:ss\.fff"))
        $script:CurrentSlideIndex = $null
    }

    $presentation.Save()
    }
    catch [System.Runtime.InteropServices.COMException] {
        $hresult = if ($_.Exception -and $null -ne $_.Exception.HResult) {
            "{0:X8}" -f ($_.Exception.HResult -band 0xFFFFFFFF)
        } else {
            "UNKNOWN"
        }
        $message = "COMException 0x$hresult encountered: $($_.Exception.Message)"
        Write-MergeLog -Message $message -IsError
        Write-WatchdogLog $message
        throw
    }
    catch {
        Write-MergeLog -Message ("Unhandled exception: {0}" -f $_) -IsError
        throw
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
        Stop-PowerPointInstances -ForceFirst -WaitMilliseconds 250
        if ($script:PostProcessStopwatch -and $script:PostProcessStopwatch.IsRunning) {
            $script:PostProcessStopwatch.Stop()
        }
        Write-MergeLog -Message ("Post-process cleanup finished; total elapsed {0}." -f $script:PostProcessStopwatch.Elapsed)
        $script:CurrentSlideIndex = $null
        $script:CurrentRowIndex = $null
        $script:CurrentTableName = $null
    }
}

# Trace-Command temporarily disabled for clean performance measurement
# if ($TraceCommandLogPath) {
#     if (Test-Path -LiteralPath $TraceCommandLogPath) {
#         try {
#             Remove-Item -LiteralPath $TraceCommandLogPath -Force
#         } catch {
#             Write-MergeLog -Message ("Unable to clear existing trace log {0}: {1}" -f $TraceCommandLogPath, $_) -IsError
#         }
#     }
#
#     Write-MergeLog -Message ("Trace-Command capture enabled; writing to {0}" -f $TraceCommandLogPath)
#     Trace-Command -Name ParameterBinding,TypeConversion -FilePath $TraceCommandLogPath -PSHost -Expression { Invoke-PostProcessMergeCore } | Out-Null
# } else {
    Invoke-PostProcessMergeCore
# }
