<#
.SYNOPSIS
    Audit campaign merges in PowerPoint tables (COM-based diagnostic).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR DIAGNOSTIC TOOLS ⚠️

    This diagnostic script uses PowerPoint COM automation and is DEPRECATED.
    COM-based tools are 60x slower than Python alternatives.

    **Replacement:**
    Python CLI provides better diagnostic capabilities:
      py -m amp_automation.presentation.postprocess.cli `
        --presentation-path deck.pptx `
        --operations postprocess-all --verbose

    **Status:** DEPRECATED as of 27 Oct 2025
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use for legacy deck analysis!
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [string]$OutputPath
)

$normalizedPath = [System.IO.Path]::GetFullPath($PresentationPath).Replace('/', '\')

if (-not (Test-Path -LiteralPath $normalizedPath)) {
    throw "Presentation not found: $normalizedPath"
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $defaultName = "merge_audit_$timestamp.csv"
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath $defaultName
}

$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath).Replace('/', '\')
$outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutput)
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

function Stop-PowerPointInstances {
    param([int]$WaitMilliseconds = 250)

    $existing = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if (-not $existing) { return }

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
        Start-Sleep -Milliseconds 100
    }
}

function Normalize-Label([string]$Text) {
    if (-not $Text) { return "" }
    $clean = [regex]::Replace($Text, "\u00A0", " ")
    $clean = [regex]::Replace($clean, "\s+", " ")
    return $clean.Trim()
}

function Get-DataTable([object]$slide) {
    foreach ($shape in $slide.Shapes) {
        if ($shape.Name -eq "MainDataTable" -and $shape.HasTable) { return $shape.Table }
    }
    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTable) { return $shape.Table }
    }
    return $null
}

function Get-Cell($table, [int]$rowIdx, [int]$colIdx) {
    try {
        return $table.Cell($rowIdx, $colIdx)
    } catch {
        return $null
    }
}

function Get-HorizontalSpan($table, [int]$rowIdx, [int]$colIdx) {
    $cell = Get-Cell $table $rowIdx $colIdx
    if (-not $cell) { return 0 }

    $span = 1
    for ($cursor = $colIdx + 1; $cursor -le $table.Columns.Count; $cursor++) {
        $next = Get-Cell $table $rowIdx $cursor
        if (-not $next) { break }
        if (-not [object]::ReferenceEquals($cell.Shape, $next.Shape)) { break }
        $span += 1
    }
    return $span
}

function Get-VerticalSpan($table, [int]$rowIdx, [int]$colIdx) {
    $cell = Get-Cell $table $rowIdx $colIdx
    if (-not $cell) { return 0 }

    $span = 1
    for ($cursor = $rowIdx + 1; $cursor -le $table.Rows.Count; $cursor++) {
        $next = Get-Cell $table $cursor $colIdx
        if (-not $next) { break }
        if (-not [object]::ReferenceEquals($cell.Shape, $next.Shape)) { break }
        $span += 1
    }
    return $span
}

$entries = New-Object System.Collections.Generic.List[object]

$pp = $null
$presentation = $null
try {
    Stop-PowerPointInstances

    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = 0 } catch { }
    try { $pp.DisplayAlerts = 1 } catch { }
    try { $pp.AutomationSecurity = 3 } catch { }

    $presentation = $pp.Presentations.Open($normalizedPath, 0, 0, 0)

    for ($slideIdx = 1; $slideIdx -le $presentation.Slides.Count; $slideIdx++) {
        $slide = $presentation.Slides.Item($slideIdx)
        $table = Get-DataTable $slide
        if (-not $table) { continue }

        for ($rowIdx = 1; $rowIdx -le $table.Rows.Count; $rowIdx++) {
            $label = ""
            $cell = Get-Cell $table $rowIdx 1
            if ($cell) {
                try {
                    $label = Normalize-Label($cell.Shape.TextFrame.TextRange.Text)
                } catch { }
            }

            # Horizontal spans starting at column 1-3
            for ($colIdx = 1; $colIdx -le [Math]::Min(3, $table.Columns.Count); $colIdx++) {
                $span = Get-HorizontalSpan $table $rowIdx $colIdx
                if ($span -gt 1) {
                    # Ensure leftmost cell capture
                    $leftNeighbor = Get-Cell $table $rowIdx ($colIdx - 1)
                    $isLeftEdge = ($colIdx -eq 1) -or (-not $leftNeighbor) -or (-not $cell) -or (-not [object]::ReferenceEquals($leftNeighbor.Shape, (Get-Cell $table $rowIdx $colIdx).Shape))
                    if ($isLeftEdge) {
                        $entries.Add([pscustomobject]@{
                            Slide       = $slideIdx
                            Row         = $rowIdx
                            Column      = $colIdx
                            SpanRows    = 1
                            SpanColumns = $span
                            Type        = "Horizontal"
                            Label       = $label
                        })
                    }
                }
            }

            # Vertical span check for column 1
            $vSpan = Get-VerticalSpan $table $rowIdx 1
            if ($vSpan -gt 1) {
                $above = Get-Cell $table ($rowIdx - 1) 1
                $isTop = ($rowIdx -eq 1) -or (-not $above) -or (-not $cell) -or (-not [object]::ReferenceEquals($above.Shape, $cell.Shape))
                if ($isTop) {
                    $entries.Add([pscustomobject]@{
                        Slide       = $slideIdx
                        Row         = $rowIdx
                        Column      = 1
                        SpanRows    = $vSpan
                        SpanColumns = 1
                        Type        = "Vertical"
                        Label       = $label
                    })
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
    Stop-PowerPointInstances
}

$entries |
    Sort-Object Slide, Row, Column |
    Export-Csv -LiteralPath $resolvedOutput -NoTypeInformation -Encoding UTF8

Write-Output ("Merge audit written to {0}" -f $resolvedOutput)

