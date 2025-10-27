param(
    [string]$PresentationPath,
    [switch]$Trace,
    [switch]$KillPowerPointFirst = $true,
    [bool]$PreSanitize = $true
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot "..\\.."))
$mergeScript = Join-Path $projectRoot "tools\PostProcessCampaignMerges.ps1"

function Get-ArabianNow {
    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Arabian Standard Time")
    } catch {
        $tz = [System.TimeZoneInfo]::CreateCustomTimeZone("UTC+04", [TimeSpan]::FromHours(4), "UTC+04", "UTC+04")
    }
    return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
}

function Resolve-LatestSanitizedDeck {
    param([string]$Root)

    $presentations = Join-Path $Root "output\presentations"
    if (-not (Test-Path -LiteralPath $presentations)) { return $null }

    return Get-ChildItem -Path $presentations -Filter "*sanitized*.pptx" -Recurse -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 |
        ForEach-Object { $_.FullName }
}

function Resolve-DailyLogDirectory {
    param([string]$Root)

    $now = Get-ArabianNow
    $docsRoot = Join-Path $Root "docs"
    $preferred = $now.ToString("dd-MM-yy")
    $alternate = $now.ToString("dd-MM-yyyy")

    $candidate = $null
    if (Test-Path -LiteralPath $docsRoot) {
        $candidate = Get-ChildItem -Path $docsRoot -Directory -Filter $preferred -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1 |
            ForEach-Object { $_.FullName }
        if (-not $candidate) {
            $candidate = Get-ChildItem -Path $docsRoot -Directory -Filter $alternate -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1 |
                ForEach-Object { $_.FullName }
        }
    }

    if (-not $candidate) {
        if (-not (Test-Path -LiteralPath $docsRoot)) {
            [void](New-Item -ItemType Directory -Path $docsRoot -Force)
        }
        $candidate = Join-Path $docsRoot $preferred
    }

    if (-not (Test-Path -LiteralPath $candidate)) {
        [void](New-Item -ItemType Directory -Path $candidate -Force)
    }

    $logs = Join-Path $candidate "logs"
    if (-not (Test-Path -LiteralPath $logs)) {
        [void](New-Item -ItemType Directory -Path $logs -Force)
    }

    return @{
        DailyDirectory = $candidate
        LogsDirectory  = $logs
        Timestamp      = $now
    }
}

if (-not (Test-Path -LiteralPath $mergeScript)) {
    throw "PostProcessCampaignMerges.ps1 not found at $mergeScript"
}

$deckPath = $PresentationPath
if (-not $deckPath) {
    $deckPath = Resolve-LatestSanitizedDeck -Root $projectRoot
}
if (-not $deckPath) {
    throw "No sanitized presentation found. Provide -PresentationPath explicitly."
}
$deckPath = [System.IO.Path]::GetFullPath($deckPath)
if (-not (Test-Path -LiteralPath $deckPath)) {
    throw "Presentation not found: $deckPath"
}

$logContext = Resolve-DailyLogDirectory -Root $projectRoot
$logsDir = $logContext.LogsDirectory
$dailyDir = $logContext.DailyDirectory
$nowLocal = $logContext.Timestamp

$transcriptName = "05-postprocess_repro_{0}.log" -f $nowLocal.ToString("yyyyMMdd_HHmmss")
$transcriptPath = Join-Path $logsDir $transcriptName

$tracePath = if ($Trace) {
    Join-Path $logsDir ("05-postprocess_trace_{0}.log" -f $nowLocal.ToString("yyyyMMdd_HHmmss"))
} else {
    $null
}

Write-Host ("Deck: {0}" -f $deckPath)
Write-Host ("Logs: {0}" -f $logsDir)

if ($KillPowerPointFirst.IsPresent) {
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch { }
    }
    Start-Sleep -Milliseconds 500
}

Start-Transcript -Path $transcriptPath -Force | Out-Null
try {
    $arguments = @{
        PresentationPath = $deckPath
        LogDirectory     = $dailyDir
        PreSanitize      = $PreSanitize
    }

    if ($tracePath) {
        $arguments["TraceCommandLogPath"] = $tracePath
    }

    & $mergeScript @arguments -Verbose
}
finally {
    Stop-Transcript | Out-Null
}
