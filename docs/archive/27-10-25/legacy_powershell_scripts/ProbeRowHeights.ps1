<#
.SYNOPSIS
    Probe row heights in PowerPoint tables (COM-based diagnostic).

.DESCRIPTION
    ⚠️ WARNING - DEPRECATED: COM AUTOMATION FOR DIAGNOSTIC TOOLS ⚠️

    This diagnostic script uses PowerPoint COM automation and is DEPRECATED.
    COM-based probing tools are slow and limited compared to Python alternatives.

    **Replacement:**
    Python-pptx provides better table inspection capabilities:
      py -c "from pptx import Presentation; prs = Presentation('deck.pptx'); ..."

    **Status:** DEPRECATED as of 27 Oct 2025
    **Documentation:** docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md

    Only use for quick legacy deck inspection!
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
    $defaultName = "row_height_probe_$timestamp.csv"
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

function Find-DataTable {
    param($slide)

    foreach ($shape in $slide.Shapes) {
        if ($shape.Name -eq "MainDataTable" -and $shape.HasTable) {
            return $shape.Table
        }
    }

    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTable) {
            return $shape.Table
        }
    }

    return $null
}

$EMU_PER_POINT = 12700.0
$rows = New-Object System.Collections.Generic.List[object]

$pp = $null
$presentation = $null
try {
    Stop-PowerPointInstances

    $pp = New-Object -ComObject PowerPoint.Application
    try { $pp.Visible = 0 } catch { }
    try { $pp.DisplayAlerts = 1 } catch { }
    try { $pp.AutomationSecurity = 3 } catch { }

    try {
        $presentation = $pp.Presentations.Open($normalizedPath, 0, 0, 0)
    } catch {
        Stop-PowerPointInstances
        Start-Sleep -Milliseconds 250
        $presentation = $pp.Presentations.Open($normalizedPath, 0, 0, 0)
    }

    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $table = Find-DataTable -slide $slide
        if (-not $table) {
            continue
        }

        for ($rowIndex = 1; $rowIndex -le $table.Rows.Count; $rowIndex++) {
            try {
                $row = $table.Rows.Item($rowIndex)
                $heightPt = [double]$row.Height
                $heightEmu = [math]::Round($heightPt * $EMU_PER_POINT, 3)
                $rows.Add([pscustomobject]@{
                    SlideIndex = $slideIndex
                    RowIndex   = $rowIndex
                    HeightEmu  = $heightEmu
                    HeightPt   = [math]::Round($heightPt, 3)
                })
            } catch {
                $rows.Add([pscustomobject]@{
                    SlideIndex = $slideIndex
                    RowIndex   = $rowIndex
                    HeightEmu  = ""
                    HeightPt   = ""
                    Error      = $_.Exception.Message
                })
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

$rows |
    Sort-Object SlideIndex, RowIndex |
    Export-Csv -LiteralPath $resolvedOutput -NoTypeInformation -Encoding UTF8

Write-Output ("Row height probe written to {0}" -f $resolvedOutput)
