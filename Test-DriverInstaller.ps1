#Requires -RunAsAdministrator
# =============================================================================
# Test-DriverInstaller.ps1  v1.0.0
# Pulls Install-Drivers-auto.ps1 from the repo and runs it headlessly
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

$ScriptVersion = "1.0.1"
$RepoUrl       = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/$Branch/Install-Drivers-auto.ps1"
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
Write-Step "Downloading Install-Drivers-auto.ps1 from repo..."
Add-Log "Downloading script from: $RepoUrl"
try {
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $RepoUrl -OutFile $ScriptCache -UseBasicParsing -ErrorAction Stop
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

foreach ($oemName in $OEM) {
    if (-not $OEMConfig.ContainsKey($oemName)) {
        Write-Warn "Unknown OEM '$oemName' — skipping. Valid: $($OEMConfig.Keys -join ', ')"
        continue
    }

    $cfg = $OEMConfig[$oemName]
    Write-Header "Testing: $oemName"
    Add-Log "=== START $oemName ==="

    $testStart  = Get-Date
    $driverRoot = "C:\DRIVERS"
    $extractDir = switch ($oemName) {
        "Dell"   { "$driverRoot\Dell_Extracted" }
        "HP"     { "$driverRoot\HP_Extracted" }
        "Lenovo" { "$driverRoot\Lenovo_Extracted" }
    }

    # Build argument list
    $args = @(
        "-ExecutionPolicy", "Bypass",
        "-File", $ScriptCache,
        "-Manufacturer", $cfg.Manufacturer,
        "-SkipCleanup"   # always keep files so we can inspect
    )
    if ($cfg.Model)        { $args += @("-Model",       $cfg.Model) }
    if ($cfg.MachineType)  { $args += @("-MachineType", $cfg.MachineType) }
    if (-not $RunInstall)  { $args += "-SkipInstall" }

    Write-Step "Running headless install..."
    Write-Info "Manufacturer: $($cfg.Manufacturer)"
    if ($cfg.Model)       { Write-Info "Model:        $($cfg.Model)" }
    if ($cfg.MachineType) { Write-Info "MachineType:  $($cfg.MachineType)" }
    Write-Info "SkipInstall:  $(-not $RunInstall)"
    Add-Log "Running: powershell $($args -join ' ')"

    # Run the script and capture output
    $outputLines = [System.Collections.Generic.List[string]]::new()
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "powershell.exe"
    $psi.Arguments              = $args -join " "
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    # Stream output live while also capturing it
    $lastDot = Get-Date
    while (-not $proc.HasExited) {
        $line = $proc.StandardOutput.ReadLine()
        if ($line -ne $null) {
            $outputLines.Add($line)
            Add-Log "  OUT: $line"
            # Show progress dots so we know it's alive
            if (((Get-Date) - $lastDot).TotalSeconds -ge 10) {
                Write-Host "  ." -NoNewline -ForegroundColor DarkGray
                $lastDot = Get-Date
            }
        }
    }
    Write-Host ""
    # Drain remaining output
    $remaining = $proc.StandardOutput.ReadToEnd()
    foreach ($line in ($remaining -split "`n" | Where-Object { $_.Trim() })) {
        $outputLines.Add($line)
        Add-Log "  OUT: $line"
    }
    $stderr = $proc.StandardError.ReadToEnd().Trim()
    if ($stderr) { Add-Log "  ERR: $stderr" }

    $exitCode   = $proc.ExitCode
    $durationSec = [math]::Round(((Get-Date) - $testStart).TotalSeconds, 1)
    Add-Log "$oemName exit code: $exitCode  duration: ${durationSec}s"

    # ----- Evaluate results -----
    $testPassed   = $true
    $testNotes    = [System.Collections.Generic.List[string]]::new()

    # Check 1: process exit code
    if ($exitCode -ne 0) {
        Write-Fail "Process exited with code $exitCode"
        $testNotes.Add("Exit code: $exitCode")
        $testPassed = $false
    } else {
        Write-OK "Process exited cleanly"
    }

    # Check 2: extraction directory exists and has files
    if (Test-Path $extractDir) {
        $infFiles  = @(Get-ChildItem $extractDir -Recurse -Filter "*.inf" -EA SilentlyContinue)
        $allFiles  = @(Get-ChildItem $extractDir -Recurse -File -EA SilentlyContinue)
        $fileCount = $allFiles.Count
        $infCount  = $infFiles.Count

        if ($fileCount -gt 0) {
            Write-OK "Extraction: $fileCount files found"
        } else {
            Write-Fail "Extraction: 0 files found in $extractDir"
            $testNotes.Add("Extraction produced 0 files")
            $testPassed = $false
        }

        if ($infCount -gt 0) {
            Write-OK "INF files:  $infCount found"
        } else {
            Write-Fail "INF files:  0 .inf files found — pnputil would have nothing to install"
            $testNotes.Add("No INF files found")
            $testPassed = $false
        }
        $testNotes.Add("Files: $fileCount  INFs: $infCount")
    } else {
        Write-Fail "Extract dir not found: $extractDir"
        $testNotes.Add("Extract directory missing")
        $testPassed = $false
        $fileCount = 0
        $infCount  = 0
    }

    # Check 3: 7-Zip was removed (Dell/HP only)
    if ($oemName -in @("Dell", "HP")) {
        $7zExe = "C:\Program Files\7-Zip\7z.exe"
        if (-not (Test-Path $7zExe)) {
            Write-OK "7-Zip removed cleanly"
        } else {
            Write-Fail "7-Zip still present after run — not safe to sysprep"
            $testNotes.Add("7-Zip not removed")
            $testPassed = $false
        }
    }

    # Check 4: look for SUCCESS/FAILED in output
    $successLine = $outputLines | Where-Object { $_ -match "SUCCESS:|Driver installation complete" } | Select-Object -First 1
    $failLine    = $outputLines | Where-Object { $_ -match "FAILED:|did not complete" } | Select-Object -First 1
    if ($successLine) {
        Write-OK "Script reported success"
    } elseif ($failLine) {
        Write-Fail "Script reported failure: $failLine"
        $testNotes.Add("Script reported failure")
        $testPassed = $false
    } else {
        Write-Warn "Could not determine script outcome from output"
        $testNotes.Add("Outcome unclear from output")
    }

    # Check 5: if RunInstall, look for pnputil output
    if ($RunInstall) {
        $pnpLines = $outputLines | Where-Object { $_ -match "pnputil|Driver package added|INFs installed" }
        if ($pnpLines) {
            Write-OK "pnputil ran ($($pnpLines.Count) relevant log lines)"
        } else {
            Write-Warn "No pnputil output detected — check log"
            $testNotes.Add("No pnputil output found")
        }
    }

    # Summary for this OEM
    $status = if ($testPassed) { "PASS" } else { "FAIL" }
    $colour = if ($testPassed) { "Green" } else { "Red" }
    Write-Host ""
    Write-Host "  $oemName result: $status  ($durationSec sec)" -ForegroundColor $colour
    Add-Log "$oemName result: $status  duration: ${durationSec}s"

    $results.Add([ordered]@{
        OEM      = $oemName
        Status   = $status
        Duration = $durationSec
        Files    = if ($null -ne $fileCount) { $fileCount } else { 0 }
        INFs     = if ($null -ne $infCount)  { $infCount  } else { 0 }
        Notes    = $testNotes -join "; "
    })

    # Clean up extracted files unless -SkipCleanup was passed to THIS test runner
    # (we always pass -SkipCleanup to the main script so we can inspect here,
    #  then we clean up ourselves)
    if (Test-Path $driverRoot) {
        Write-Info "Cleaning up $driverRoot..."
        try { Remove-Item $driverRoot -Recurse -Force -EA Stop; Write-Info "  Removed." }
        catch { Write-Warn "Could not remove $driverRoot`: $_" }
    }

    Add-Log "=== END $oemName ==="
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
