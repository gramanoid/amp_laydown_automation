<#
.SYNOPSIS
    Python-based post-processing for PowerPoint presentations.

.DESCRIPTION
    Modern replacement for PostProcessCampaignMerges.ps1 using Python instead of COM.

    This script delegates all bulk operations to the Python CLI, providing:
    - Fast table normalization (<1 minute for 88 slides vs 10+ hours COM)
    - Cell formatting and layout consistency
    - Optional edge-case merge repairs

    Architecture:
    - Cell merges: Created during generation (assembly.py), not post-processing
    - Post-processing: Normalization, formatting, edge case fixes only

    See: docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md
         openspec/changes/clarify-postprocessing-architecture/

.PARAMETER PresentationPath
    Path to the PowerPoint presentation file (.pptx).

.PARAMETER Operations
    Comma-separated list of operations to perform.
    Options: normalize, reset-spans, merge-campaign, merge-monthly, merge-summary
    Default: normalize

.PARAMETER SlideFilter
    Array of slide numbers to process (1-based). If omitted, processes all slides.

.PARAMETER Verbose
    Enable verbose logging output.

.EXAMPLE
    .\PostProcessNormalize.ps1 -PresentationPath "output\presentations\deck.pptx"

    Runs normalization on all slides.

.EXAMPLE
    .\PostProcessNormalize.ps1 -PresentationPath "deck.pptx" -Operations "normalize,merge-campaign" -SlideFilter 2,3,4 -VerboseOutput

    Runs normalization and campaign merge repair on slides 2, 3, and 4 with verbose output.

.NOTES
    Requires: Python 3.13+ with python-pptx package
    Replaces: PostProcessCampaignMerges.ps1 (deprecated COM-based script)
    Created: 2025-10-24
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [string]$Operations = "normalize",

    [int[]]$SlideFilter,

    [switch]$VerboseOutput
)

$ErrorActionPreference = "Stop"

# Resolve paths
$script:ProjectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$script:PresentationFullPath = [System.IO.Path]::GetFullPath($PresentationPath)

# Verify presentation exists
if (-not (Test-Path $script:PresentationFullPath)) {
    Write-Error "Presentation not found: $script:PresentationFullPath"
    exit 1
}

# Get Arabian time for logging
function Get-ArabianNow {
    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Arabian Standard Time")
    } catch {
        $tz = [System.TimeZoneInfo]::CreateCustomTimeZone("UTC+04", [TimeSpan]::FromHours(4), "UTC+04", "UTC+04")
    }
    return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
}

# Setup logging
$nowLocal = Get-ArabianNow
$timestamp = $nowLocal.ToString("yyyyMMdd_HHmmss")
$docsRoot = Join-Path $script:ProjectRoot "docs"
$dailyDirName = $nowLocal.ToString("dd-MM-yy")
$dailyDir = Get-ChildItem -Path $docsRoot -Directory -Filter "*$dailyDirName*" -ErrorAction SilentlyContinue | Select-Object -First 1

if ($dailyDir) {
    $logsDir = Join-Path $dailyDir.FullName "logs"
    if (-not (Test-Path $logsDir)) {
        $logsDir = $dailyDir.FullName
    }
} else {
    $logsDir = $docsRoot
}

$logFile = Join-Path $logsDir "postprocess_normalize_${timestamp}.log"
Write-Host "Log file: $logFile"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = (Get-ArabianNow).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"
    Write-Host $logLine
    Add-Content -Path $logFile -Value $logLine
}

# Start processing
Write-Log "=== Python Post-Processing Started ===" "INFO"
Write-Log "Presentation: $script:PresentationFullPath" "INFO"
Write-Log "Operations: $Operations" "INFO"
if ($SlideFilter) {
    Write-Log "Slide filter: $($SlideFilter -join ', ')" "INFO"
}

# Build Python CLI command
$pythonArgs = @(
    "-m", "amp_automation.presentation.postprocess.cli",
    "--presentation-path", $script:PresentationFullPath,
    "--operations", $Operations
)

if ($SlideFilter) {
    $slideFilterStr = $SlideFilter -join ','
    $pythonArgs += "--slide-filter"
    $pythonArgs += $slideFilterStr
}

if ($VerboseOutput) {
    $pythonArgs += "--verbose"
}

# Execute Python CLI
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Executing: python $($pythonArgs -join ' ')" "INFO"

try {
    $pythonOutput = & python $pythonArgs 2>&1
    $exitCode = $LASTEXITCODE

    # Log Python output
    $pythonOutput | ForEach-Object {
        Add-Content -Path $logFile -Value $_
        if ($VerboseOutput) {
            Write-Host $_
        }
    }

    $stopwatch.Stop()
    $elapsed = $stopwatch.Elapsed.ToString("mm\:ss")

    if ($exitCode -eq 0) {
        Write-Log "Post-processing completed successfully in $elapsed" "SUCCESS"
        exit 0
    } else {
        Write-Log "Post-processing failed with exit code $exitCode after $elapsed" "ERROR"
        exit $exitCode
    }
} catch {
    $stopwatch.Stop()
    $elapsed = $stopwatch.Elapsed.ToString("mm\:ss")
    Write-Log "Post-processing failed after $elapsed : $_" "ERROR"
    Write-Log "Exception: $($_.Exception.Message)" "ERROR"
    exit 1
}
