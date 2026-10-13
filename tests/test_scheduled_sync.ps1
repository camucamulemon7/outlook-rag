$ErrorActionPreference = "Stop"
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $root | Out-Null
$config = Join-Path $root "config.json"
Set-Content -LiteralPath $config -Value '{}'
$global:calls = [System.Collections.Generic.List[string]]::new()
$global:fixtureState = "running"
function Invoke-FixtureRag {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$arguments)
    if ($arguments[0] -ne "--config" -or $arguments[1] -ne $config) { throw "Configuration arguments missing" }
    $command = $arguments[2..($arguments.Count - 1)] -join " "
    $global:calls.Add($command)
    $global:LASTEXITCODE = 0
    switch ($command) {
        "job status" { @{ state = $global:fixtureState; job_id = "fixture"; options = @{ reconcile = $true } } | ConvertTo-Json }
        "status" { @{ data_dir = $root; chunks = 1000; vector_coverage = @{ unindexed_rows = 300 } } | ConvertTo-Json }
        default { '{}' }
    }
}
try {
    # Run in a child scope: an active worker returns without upkeep or another start.
    & "$PSScriptRoot/../scripts/sync-scheduled.ps1" -Config $config -Executable Invoke-FixtureRag
    if ($global:calls.Count -ne 1) { throw "Active job was not reused" }
    $global:calls.Clear()
    $global:fixtureState = "completed"
    & "$PSScriptRoot/../scripts/sync-scheduled.ps1" -Config $config -Executable Invoke-FixtureRag
    if (-not $global:calls.Contains("optimize")) { throw "ANN tail was not optimized" }
    if (-not $global:calls.Contains("maintain")) { throw "Weekly upkeep was not run" }
    if (-not $global:calls.Contains("job start --folders auto --since-days 0")) { throw "Incremental all-store sync did not start" }
    if (-not (Test-Path -LiteralPath (Join-Path $root "schedule-state.json"))) { throw "Schedule state missing" }
    Write-Output "Scheduled sync scenarios passed"
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force
}
