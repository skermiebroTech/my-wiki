#Requires -RunAsAdministrator
# =============================================================================
# Test-DriverInstaller.ps1  v1.0.0
# Pulls Install-Drivers-auto-7z.ps1 from the repo and runs it headlessly
# for Dell, HP, and Lenovo with SkipInstall + SkipCleanup so you can
# verify extraction without touching the driver store.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File Test-DriverInstaller.ps1
#   powershell -ExecutionPolicy Bypass -File Test-DriverInstaller.ps1 -OEM Dell
#   powershell -ExecutionPolicy Bypass -File Test-DriverInstaller.ps1 -OEM HP,Lenovo
#   powershell -ExecutionPolicy Bypass -File Test-DriverInstaller.ps1 -RunInstall
#
# Parameters:
#   -OEM              Comma-separated list of OEMs to test (default: Dell,HP,Lenovo)
#   -RunInstall       Also run pnputil driver install (default: extract-only)
#   -Branch           GitHub branch to pull from (default: main)
#   -DellModel        Override Dell model string (default: Latitude 7320)
#   -HPModel          Override HP model string (default: EliteBook x360 1030 G8)
#   -LenovoMachineType Override Lenovo machine type, e.g. 20XX (default: read from WMI)
# =============================================================================

param(
    [string[]]$OEM                = @("Dell", "HP", "Lenovo"),
    [switch]$RunInstall,
    [string]$Branch               = "main",
    [string]$DellModel            = "Latitude 7320",
    [string]$HPModel              = "HP EliteBook x360 1030 G8 Notebook PC",
    [string]$LenovoMachineType    = ""
)

$ScriptVersion = "1.0.3"
$RepoUrl       = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/$Branch/Install-Drivers-auto-7z.ps1"
$LogFile       = Join-Path ([Environment]::GetFolderPath("UserProfile")) `
                     ("Downloads\Test-DriverInstaller_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
$ScriptCache   = "$env:TEMP\Install-Drivers-auto-test.ps1"

# OEM test definitions
# MachineType: Lenovo-specific — passed as -MachineType to the main script.
#   Use the first 4 chars of the device serial, e.g. serial "20XXS1J203" -> "20XX"
#   Leave blank to read machine type from WMI (requires a real Lenovo machine).
$OEMConfig = @{
    Dell   = @{ Manufacturer = "Dell Inc."; Model = $DellModel;  MachineType = "" }
    HP     = @{ Manufacturer = "HP";        Model = $HPModel;    MachineType = "" }
    Lenovo = @{ Manufacturer = "LENOVO";    Model = "";          MachineType = $LenovoMachineType }
}

# -----------------------------------------------------------------------------
function Write-Header { param([string]$m)
    Write-Host ""
    Write-Host "  ==========================================" -ForegroundColor DarkGray
    Write-Host "  $m" -ForegroundColor White
    Write-Host "  ==========================================" -ForegroundColor DarkGray
}
function Write-Step  { param([string]$m); Write-Host "`n  --> $m" -ForegroundColor Cyan }
function Write-OK    { param([string]$m); Write-Host "  [PASS] $m" -ForegroundColor Green }
function Write-Fail  { param([string]$m); Write-Host "  [FAIL] $m" -ForegroundColor Red }
function Write-Info  { param([string]$m); Write-Host "  $m" -ForegroundColor Gray }
function Write-Warn  { param([string]$m); Write-Host "  [WARN] $m" -ForegroundColor Yellow }

function Add-Log { param([string]$m)
    $ts = Get-Date -Format 'HH:mm:ss'
    Add-Content -Path $LogFile -Value "[$ts] $m" -Encoding UTF8
}

# -----------------------------------------------------------------------------
Write-Header "Driver Installer Test Runner  v$ScriptVersion"
Write-Info "Branch:      $Branch"
Write-Info "OEMs:        $($OEM -join ', ')"
Write-Info "Run install: $RunInstall"
Write-Info "Log:         $LogFile"
New-Item -ItemType File -Path $LogFile -Force | Out-Null

# -----------------------------------------------------------------------------
Write-Step "Downloading Install-Drivers-auto-7z.ps1 from repo..."
Add-Log "Downloading script from: $RepoUrl"
try {
    $ProgressPreference = 'SilentlyContinue'
    # Use curl.exe to preserve UTF-8 encoding — Invoke-WebRequest mangles special chars
    $curlExit = (Start-Process curl.exe -ArgumentList "--silent --location --output `"$ScriptCache`" `"$RepoUrl`"" -Wait -PassThru).ExitCode
    if ($curlExit -ne 0) { throw "curl.exe failed with exit code $curlExit" }
    $sizeMB = [math]::Round((Get-Item $ScriptCache).Length / 1KB, 1)
    Write-OK "Downloaded ($sizeMB KB)"
    Add-Log "Download OK: $sizeMB KB"

    # Extract version from downloaded script
    $versionLine = Get-Content $ScriptCache | Select-String '^\$ScriptVersion\s*=' | Select-Object -First 1
    $version = if ($versionLine) { $versionLine -replace '.*"(.*)".*','$1' } else { "unknown" }
    Write-Info "Script version: $version"
    Add-Log "Script version: $version"
} catch {
    Write-Fail "Failed to download script: $_"
    Add-Log "FAIL: Download failed: $_"
    exit 1
}

# -----------------------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()

# Run all OEMs in parallel using PowerShell jobs
$jobs = [System.Collections.Generic.List[object]]::new()

foreach ($oemName in $OEM) {
    if (-not $OEMConfig.ContainsKey($oemName)) {
        Write-Warn "Unknown OEM '$oemName' - skipping. Valid: $($OEMConfig.Keys -join ', ')"
        continue
    }

    $cfg        = $OEMConfig[$oemName]
    $driverRoot = "C:\DRIVERS\$oemName"   # separate folder per OEM so parallel runs don't collide
    $extractDir = "$driverRoot\$($oemName)_Extracted"

    # Build arg string — pass DriverRoot so each OEM uses its own folder
    $argList  = "-ExecutionPolicy Bypass"
    $argList += " -File `"$ScriptCache`""
    $argList += " -Manufacturer `"$($cfg.Manufacturer)`""
    if ($cfg.Model)       { $argList += " -Model `"$($cfg.Model)`"" }
    if ($cfg.MachineType) { $argList += " -MachineType `"$($cfg.MachineType)`"" }
    $argList += " -DriverRoot `"$driverRoot`""
    $argList += " -SkipCleanup"
    if (-not $RunInstall) { $argList += " -SkipInstall" }

    Write-Header "Starting: $oemName"
    Write-Info "Manufacturer: $($cfg.Manufacturer)"
    if ($cfg.Model)       { Write-Info "Model:        $($cfg.Model)" }
    if ($cfg.MachineType) { Write-Info "MachineType:  $($cfg.MachineType)" }
    Write-Info "SkipInstall:  $(-not $RunInstall)"
    Add-Log "=== START $oemName ==="
    Add-Log "Running: powershell $argList"

    # Capture vars needed inside the job scriptblock
    $jobArgList    = $argList
    $jobOemName    = $oemName
    $jobDriverRoot = $driverRoot
    $jobExtractDir = $extractDir
    $jobLogFile    = $LogFile
    $jobRunInstall = $RunInstall

    $job = Start-Job -ScriptBlock {
        param($argList, $oemName, $driverRoot, $extractDir, $logFile, $runInstall)

        function JobLog { param([string]$m)
            $ts = Get-Date -Format 'HH:mm:ss'
            $line = "[$ts][$oemName] $m"
            Add-Content -Path $logFile -Value $line -Encoding UTF8
            $line  # also output so we can capture it
        }

        $startTime = Get-Date

        # Run the installer script
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = "powershell.exe"
        $psi.Arguments              = $argList
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $proc.Start() | Out-Null

        $outputLines = [System.Collections.Generic.List[string]]::new()
        while (-not $proc.HasExited) {
            $line = $proc.StandardOutput.ReadLine()
            if ($line -ne $null) { $outputLines.Add($line); JobLog "OUT: $line" | Out-Null }
        }
        $remaining = $proc.StandardOutput.ReadToEnd()
        foreach ($line in ($remaining -split "`n" | Where-Object { $_.Trim() })) {
            $outputLines.Add($line); JobLog "OUT: $line" | Out-Null
        }
        $stderr = $proc.StandardError.ReadToEnd().Trim()
        if ($stderr) { JobLog "ERR: $stderr" | Out-Null }

        $exitCode    = $proc.ExitCode
        $durationSec = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)

        # Evaluate results
        $testPassed = $true
        $testNotes  = [System.Collections.Generic.List[string]]::new()
        $fileCount  = 0
        $infCount   = 0

        if ($exitCode -ne 0) {
            $testNotes.Add("Exit code: $exitCode")
            $testPassed = $false
        }

        if (Test-Path $extractDir) {
            $infFiles  = @(Get-ChildItem $extractDir -Recurse -Filter "*.inf" -EA SilentlyContinue)
            $allFiles  = @(Get-ChildItem $extractDir -Recurse -File -EA SilentlyContinue)
            $fileCount = $allFiles.Count
            $infCount  = $infFiles.Count
            if ($fileCount -eq 0) { $testNotes.Add("Extraction produced 0 files"); $testPassed = $false }
            if ($infCount  -eq 0) { $testNotes.Add("No INF files found");          $testPassed = $false }
        } else {
            $testNotes.Add("Extract directory missing"); $testPassed = $false
        }

        if ($oemName -in @("Dell","HP")) {
            if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
                $testNotes.Add("7-Zip not removed"); $testPassed = $false
            }
        }

        $successLine = $outputLines | Where-Object { $_ -match "SUCCESS:|Driver installation complete" } | Select-Object -First 1
        $failLine    = $outputLines | Where-Object { $_ -match "FAILED:|did not complete" } | Select-Object -First 1
        if (-not $successLine -and $failLine) {
            $testNotes.Add("Script reported failure"); $testPassed = $false
        }
        if (-not $successLine -and -not $failLine) {
            $testNotes.Add("Outcome unclear from output")
        }

        if ($runInstall) {
            $pnpLines = $outputLines | Where-Object { $_ -match "pnputil|Driver package added|INFs installed" }
            if (-not $pnpLines) { $testNotes.Add("No pnputil output found") }
        }

        # Cleanup extracted files
        if (Test-Path $driverRoot) {
            Remove-Item $driverRoot -Recurse -Force -EA SilentlyContinue
        }

        JobLog "=== END $oemName - $( if ($testPassed) {'PASS'} else {'FAIL'} ) ===" | Out-Null

        # Return result object
        [ordered]@{
            OEM      = $oemName
            Status   = if ($testPassed) { "PASS" } else { "FAIL" }
            Duration = $durationSec
            Files    = $fileCount
            INFs     = $infCount
            Notes    = $testNotes -join "; "
        }
    } -ArgumentList $jobArgList, $jobOemName, $jobDriverRoot, $jobExtractDir, $LogFile, $jobRunInstall

    $jobs.Add([ordered]@{ OEM = $oemName; Job = $job; StartTime = (Get-Date) })
    Write-OK "$oemName job started (PID tracking via job ID $($job.Id))"
}

# Wait for all jobs and stream status updates
Write-Header "Waiting for all OEM jobs to complete..."
$completed = @{}

while ($completed.Count -lt $jobs.Count) {
    foreach ($entry in $jobs) {
        $oemName = $entry.OEM
        $job     = $entry.Job
        if ($completed.ContainsKey($oemName)) { continue }

        $elapsed = [math]::Round(((Get-Date) - $entry.StartTime).TotalSeconds, 0)

        if ($job.State -in @("Completed","Failed","Stopped")) {
            $completed[$oemName] = $true
            $colour = if ($job.State -eq "Completed") { "Cyan" } else { "Red" }
            Write-Host "  [$oemName] Job finished ($elapsed sec) - state: $($job.State)" -ForegroundColor $colour
        } else {
            Write-Host "`r  Waiting... $oemName`: ${elapsed}s  " -NoNewline -ForegroundColor DarkGray
        }
    }
    if ($completed.Count -lt $jobs.Count) { Start-Sleep -Milliseconds 2000 }
}
Write-Host ""

# Collect results
foreach ($entry in $jobs) {
    $result = Receive-Job $entry.Job -EA SilentlyContinue
    if ($result) { $results.Add($result) }
    else {
        # Job failed to return a result
        $results.Add([ordered]@{
            OEM      = $entry.OEM
            Status   = "FAIL"
            Duration = [math]::Round(((Get-Date) - $entry.StartTime).TotalSeconds, 1)
            Files    = 0
            INFs     = 0
            Notes    = "Job failed to return result (state: $($entry.Job.State))"
        })
    }
    Remove-Job $entry.Job -Force -EA SilentlyContinue
}
# Remove cached script
Remove-Item $ScriptCache -Force -EA SilentlyContinue

# -----------------------------------------------------------------------------
Write-Header "Test Summary"
$passCount = ($results | Where-Object { $_.Status -eq "PASS" }).Count
$failCount = ($results | Where-Object { $_.Status -eq "FAIL" }).Count

foreach ($r in $results) {
    $colour = if ($r.Status -eq "PASS") { "Green" } else { "Red" }
    Write-Host ("  {0,-8} {1,-6}  {2,6}s  {3,5} files  {4,4} INFs  {5}" -f `
        $r.OEM, $r.Status, $r.Duration, $r.Files, $r.INFs, $r.Notes) -ForegroundColor $colour
}

Write-Host ""
Write-Host ("  {0} passed  {1} failed  out of {2} tests" -f $passCount, $failCount, $results.Count) `
    -ForegroundColor (if ($failCount -eq 0) { "Green" } else { "Red" })
Write-Host "  Log: $LogFile" -ForegroundColor Gray
Write-Host ""

Add-Log "=== SUMMARY: $passCount passed, $failCount failed ==="
foreach ($r in $results) { Add-Log "$($r.OEM): $($r.Status)  $($r.Notes)" }

exit $failCount