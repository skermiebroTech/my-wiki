<#
.SYNOPSIS
    Windows Fingerprint Reader Diagnostic.

.DESCRIPTION
    Reports on biometric hardware, the Windows Biometric Service,
    Windows Hello availability, current-user enrollment, and runs
    a live authentication prompt.

    Designed to be launched from Win+R with no dependencies:

        powershell -ep bypass -nop -c "irm https://raw.githubusercontent.com/<user>/<repo>/main/fp-test.ps1 | iex"

.NOTES
    Uses the WinRT UserConsentVerifier API. It authenticates with whichever
    Windows Hello method the user has enrolled (fingerprint, face, or PIN).
    For a pure fingerprint test, enroll only a fingerprint, or watch which
    modality the Hello prompt requests.
#>

$Host.UI.RawUI.WindowTitle = 'Fingerprint Reader Test'
$ErrorActionPreference = 'Continue'

function Write-Pass { param($t) Write-Host '  [ OK ] ' -ForegroundColor Green  -NoNewline; Write-Host $t }
function Write-Warn { param($t) Write-Host '  [WARN] ' -ForegroundColor Yellow -NoNewline; Write-Host $t }
function Write-Fail { param($t) Write-Host '  [FAIL] ' -ForegroundColor Red    -NoNewline; Write-Host $t }
function Write-Info { param($t) Write-Host '  [INFO] ' -ForegroundColor Gray   -NoNewline; Write-Host $t }
function Write-Section { param($t) Write-Host ''; Write-Host $t -ForegroundColor Cyan }

Clear-Host
Write-Host ''
Write-Host '  ===============================================' -ForegroundColor Cyan
Write-Host '     Windows Fingerprint Reader Diagnostic'        -ForegroundColor Cyan
Write-Host '  ===============================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Computer : $env:COMPUTERNAME" -ForegroundColor Gray
Write-Host "  User     : $env:USERNAME"     -ForegroundColor Gray
Write-Host "  Date     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if ($os) { Write-Host "  OS       : $($os.Caption) ($($os.Version))" -ForegroundColor Gray }

# ---------- 1. Biometric hardware ----------
Write-Section '[1] Biometric hardware (PnP class: Biometric)'
try {
    $devices = Get-PnpDevice -Class Biometric -ErrorAction Stop
    if (-not $devices) {
        Write-Fail 'No biometric devices detected. This PC may not have a fingerprint reader, or its driver is missing.'
    }
    foreach ($d in $devices) {
        switch ($d.Status) {
            'OK'       { Write-Pass "$($d.FriendlyName)" ; Write-Info "InstanceId: $($d.InstanceId)" }
            'Error'    { Write-Fail "$($d.FriendlyName) -- driver reports an error" }
            'Unknown'  { Write-Warn "$($d.FriendlyName) -- status unknown" }
            default    { Write-Warn "$($d.FriendlyName) -- $($d.Status)" }
        }
    }
} catch {
    Write-Fail "Could not enumerate biometric devices: $($_.Exception.Message)"
}

# ---------- 2. Windows Biometric Service ----------
Write-Section '[2] Windows Biometric Service (WbioSrvc)'
$svc = Get-Service -Name WbioSrvc -ErrorAction SilentlyContinue
if (-not $svc) {
    Write-Fail 'Service WbioSrvc not found.'
} else {
    if ($svc.Status -eq 'Running') { Write-Pass "Status: Running" }
    else                            { Write-Warn "Status: $($svc.Status) (expected Running)" }
    Write-Info "Start type: $($svc.StartType)"
}

# ---------- 3. Windows Hello availability ----------
Write-Section '[3] Windows Hello availability'
$availability = $null
$awaitReady   = $false
try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
    [Windows.Security.Credentials.UI.UserConsentVerifier, Windows.Security.Credentials.UI, ContentType = WindowsRuntime] | Out-Null

    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]

    function Await {
        param($winRtTask, $resultType)
        $m = $asTaskGeneric.MakeGenericMethod($resultType)
        $t = $m.Invoke($null, @($winRtTask))
        $t.Wait(-1) | Out-Null
        $t.Result
    }
    $awaitReady = $true

    $availability = Await `
        ([Windows.Security.Credentials.UI.UserConsentVerifier]::CheckAvailabilityAsync()) `
        ([Windows.Security.Credentials.UI.UserConsentVerifierAvailability])

    switch ("$availability") {
        'Available'            { Write-Pass 'Windows Hello is available and configured' }
        'DeviceNotPresent'     { Write-Fail 'No biometric/PIN device is present' }
        'NotConfiguredForUser' { Write-Warn 'Hello is not configured for this user (no PIN/biometric enrolled)' }
        'DisabledByPolicy'     { Write-Fail 'Disabled by group policy' }
        'DeviceBusy'           { Write-Warn 'Device is busy' }
        default                { Write-Warn "Availability: $availability" }
    }
} catch {
    Write-Fail "Could not query Windows Hello: $($_.Exception.Message)"
}

# ---------- 4. Per-user enrollment ----------
Write-Section '[4] Fingerprint enrollment for current user'
try {
    $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WinBio\AccountInfo\$sid"
    if (Test-Path $key) {
        Write-Pass 'Biometric enrollment record found for this user.'
        Write-Info "Registry: $key"
    } else {
        Write-Warn 'No biometric enrollment record for the current SID.'
        Write-Info 'Enroll via Settings > Accounts > Sign-in options > Fingerprint recognition.'
    }
} catch {
    Write-Warn "Could not check enrollment registry: $($_.Exception.Message)"
}

# ---------- 5. Live authentication test ----------
Write-Section '[5] Live authentication test'
if (-not $awaitReady -or "$availability" -ne 'Available') {
    Write-Warn 'Skipping live test (Windows Hello is not available on this system).'
} else {
    Write-Host ''
    Write-Host '  A Windows Hello prompt will appear in a moment.'        -ForegroundColor Yellow
    Write-Host '  Touch the fingerprint sensor when prompted.'            -ForegroundColor Yellow
    Write-Host '  Press any key to launch the prompt...'                  -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

    try {
        $result = Await `
            ([Windows.Security.Credentials.UI.UserConsentVerifier]::RequestVerificationAsync('Fingerprint reader diagnostic test')) `
            ([Windows.Security.Credentials.UI.UserConsentVerificationResult])

        switch ("$result") {
            'Verified'             { Write-Pass 'Authentication SUCCEEDED -- the reader works.' }
            'DeviceNotPresent'     { Write-Fail 'No biometric device present.' }
            'NotConfiguredForUser' { Write-Warn 'User has no enrolled credential.' }
            'DisabledByPolicy'     { Write-Fail 'Disabled by policy.' }
            'DeviceBusy'           { Write-Warn 'Device busy.' }
            'RetriesExhausted'     { Write-Fail 'Too many failed attempts -- sensor or fingerprint may be faulty.' }
            'Canceled'             { Write-Warn 'Prompt was cancelled by the user.' }
            default                { Write-Warn "Result: $result" }
        }
    } catch {
        Write-Fail "Verification call failed: $($_.Exception.Message)"
    }
}

Write-Host ''
Write-Host '  ===============================================' -ForegroundColor Cyan
Write-Host '     Diagnostic complete'                          -ForegroundColor Cyan
Write-Host '  ===============================================' -ForegroundColor Cyan
Write-Host ''
Write-Host '  Press any key to close...' -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')