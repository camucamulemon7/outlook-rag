param(
    [Parameter(Mandatory = $true)][string]$Config,
    [string]$Executable = "outlook-rag",
    [int]$ReconcileEveryHours = 24,
    [int]$MaintainEveryDays = 7
)
# Intended for an opt-in Task Scheduler task under the logged-in Outlook user.
$ErrorActionPreference = "Stop"
$Config = (Resolve-Path -LiteralPath $Config).Path
if ($ReconcileEveryHours -lt 1 -or $MaintainEveryDays -lt 1) { throw "Intervals must be positive" }
function Invoke-Rag([string[]]$Arguments) {
    $output = & $Executable --config $Config @Arguments
    if ($LASTEXITCODE -ne 0) { throw "outlook-rag command failed" }
    return ($output -join "`n" | ConvertFrom-Json)
}
$job = Invoke-Rag @("job", "status")
if ($job.state -in @("starting", "running", "waiting")) { return }
$index = Invoke-Rag @("status")
$statePath = Join-Path $index.data_dir "schedule-state.json"
$state = @{ reconcileAt = 0; maintainAt = 0; completedJob = "" }
if (Test-Path -LiteralPath $statePath) {
    $saved = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    $state.reconcileAt = $saved.reconcileAt
    $state.maintainAt = $saved.maintainAt
    $state.completedJob = $saved.completedJob
}
$now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
if ($job.state -eq "completed" -and $job.job_id -ne $state.completedJob) {
    if ($job.options.reconcile) { $state.reconcileAt = $now }
    $state.completedJob = $job.job_id
    # optimize() also absorbs appended rows. Small indexes continue exact search.
    if ($index.chunks -ge 256 -and $index.vector_coverage.unindexed_rows -ge [Math]::Max(256, $index.chunks * 0.1)) {
        $null = Invoke-Rag @("optimize")
    }
    if ($now - $state.maintainAt -ge $MaintainEveryDays * 86400) {
        $null = Invoke-Rag @("maintain")
        $state.maintainAt = $now
    }
}
$arguments = @("job", "start", "--folders", "auto", "--since-days", "0")
if ($now - $state.reconcileAt -ge $ReconcileEveryHours * 3600) { $arguments += "--reconcile" }
$null = Invoke-Rag $arguments
# Update only after successful commands. State holds timestamps/IDs, no credentials.
$temporary = "$statePath.$([guid]::NewGuid().ToString('N')).tmp"
$state | ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding UTF8
Move-Item -LiteralPath $temporary -Destination $statePath -Force
