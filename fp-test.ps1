<#
.SYNOPSIS
    Windows Fingerprint Reader Diagnostic - works in OOBE, audit mode, and normal Windows.

.DESCRIPTION
    Calls the Windows Biometric Framework (winbio.dll) directly so the live
    sensor test does not require a PIN, an enrolled fingerprint, or any user
    account configuration. Suitable for running during OOBE (Shift+F10) and
    audit mode (Win+R).

    Steps performed:
      1. Show host info and detected setup phase (OOBE / Audit / Normal).
      2. Enumerate biometric hardware via PnP.
      3. Verify the Windows Biometric Service (WbioSrvc) is running.
      4. Enumerate fingerprint units via WinBioEnumBiometricUnits.
      5. Open a session and locate the sensor.
      6. Live capture: WinBioCaptureSample waits for a finger touch.

.NOTES
    Launch one-liner (Win+R in audit mode, or Shift+F10 cmd in OOBE):

        powershell -ep bypass -nop -c "irm https://raw.githubusercontent.com/<user>/<repo>/main/fp-test.ps1 | iex"

    Offline (script on USB stick):

        powershell -ep bypass -f X:\fp-test.ps1
#>

$Host.UI.RawUI.WindowTitle = 'Fingerprint Reader Test'
$ErrorActionPreference = 'Continue'

# ---------- helpers ----------
function Pass    { param($t) Write-Host '  [ OK ] ' -ForegroundColor Green  -NoNewline; Write-Host $t }
function Warn    { param($t) Write-Host '  [WARN] ' -ForegroundColor Yellow -NoNewline; Write-Host $t }
function Fail    { param($t) Write-Host '  [FAIL] ' -ForegroundColor Red    -NoNewline; Write-Host $t }
function Info    { param($t) Write-Host '  [INFO] ' -ForegroundColor Gray   -NoNewline; Write-Host $t }
function Section { param($t) Write-Host ''; Write-Host $t -ForegroundColor Cyan }

function Get-WinBioError {
    param([int]$Hr)
    $hex = '0x{0:X8}' -f $Hr
    switch ($Hr) {
        0           { 'S_OK' ; break }
        -2147024891 { "E_ACCESSDENIED ($hex) -- run as admin/SYSTEM" ; break }              # 0x80070005
        -2143289339 { "WINBIO_E_DEVICE_BUSY ($hex)" ; break }                                # 0x80098005
        -2143289337 { "WINBIO_E_INVALID_DEVICE_STATE ($hex)" ; break }                       # 0x80098007
        -2143289334 { "WINBIO_E_BAD_CAPTURE ($hex) -- sensor works, sample rejected" ; break } # 0x8009800A
        -2143289320 { "WINBIO_E_NO_MATCH ($hex)" ; break }                                   # 0x80098018
        -2143289299 { "WINBIO_E_CANCELED ($hex)" ; break }                                   # 0x8009802D
        -2143289324 { "WINBIO_E_CAPTURE_ABORTED ($hex)" ; break }                            # 0x80098014
        -2143289316 { "WINBIO_E_UNSUPPORTED_FACTOR ($hex)" ; break }                         # 0x8009801C
        default     { "HRESULT $hex" }
    }
}

# ---------- WinBio P/Invoke ----------
if (-not ('WinBio' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class WinBio
{
    public const uint TYPE_FINGERPRINT = 0x00000008;
    public const uint POOL_SYSTEM  = 1;
    public const uint POOL_PRIVATE = 2;
    public const uint FLAG_DEFAULT = 0x00000000;
    public const byte PURPOSE_VERIFY = 0x01;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WINBIO_UNIT_SCHEMA
    {
        public uint UnitId;
        public uint PoolType;
        public uint BiometricFactor;
        public uint SensorSubType;
        public uint Capabilities;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string DeviceInstanceId;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Description;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Manufacturer;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Model;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string SerialNumber;
        public uint FirmwareMajor;
        public uint FirmwareMinor;
    }

    [DllImport("winbio.dll")] public static extern int WinBioEnumBiometricUnits(
        uint Factor, out IntPtr UnitSchemaArray, out IntPtr UnitCount);

    [DllImport("winbio.dll")] public static extern int WinBioOpenSession(
        uint Factor, uint PoolType, uint Flags,
        IntPtr UnitArray, IntPtr UnitCount, IntPtr DatabaseId,
        out IntPtr SessionHandle);

    [DllImport("winbio.dll")] public static extern int WinBioCloseSession(IntPtr SessionHandle);

    [DllImport("winbio.dll")] public static extern int WinBioLocateSensor(
        IntPtr SessionHandle, out uint UnitId);

    [DllImport("winbio.dll")] public static extern int WinBioCaptureSample(
        IntPtr SessionHandle, byte Purpose, byte Flags,
        out uint UnitId, out IntPtr Sample, out IntPtr SampleSize,
        out int RejectDetail);

    [DllImport("winbio.dll")] public static extern int WinBioCancel(IntPtr SessionHandle);

    [DllImport("winbio.dll")] public static extern int WinBioFree(IntPtr Address);
}
'@
}

# ---------- header ----------
Clear-Host
Write-Host ''
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host '     Windows Fingerprint Reader Diagnostic'             -ForegroundColor Cyan
Write-Host '     (OOBE / Audit / Normal compatible)'                -ForegroundColor Cyan
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Computer : $env:COMPUTERNAME"                                          -ForegroundColor Gray
Write-Host "  Identity : $([Security.Principal.WindowsIdentity]::GetCurrent().Name)" -ForegroundColor Gray
Write-Host "  Date     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"                  -ForegroundColor Gray
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if ($os) { Write-Host "  OS       : $($os.Caption) ($($os.Version))" -ForegroundColor Gray }

# Setup phase
$imgState = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Setup\State' -ErrorAction SilentlyContinue).ImageState
$mode = switch -Wildcard ($imgState) {
    'IMAGE_STATE_COMPLETE' { 'Normal Windows' ; break }
    '*RESEAL_TO_OOBE*'     { 'OOBE'           ; break }
    '*RESEAL_TO_AUDIT*'    { 'Audit'          ; break }
    default                { if ($imgState) { $imgState } else { 'Unknown' } }
}
$modeColor = if ($mode -eq 'Normal Windows') { 'Gray' } else { 'Magenta' }
Write-Host "  Mode     : $mode" -ForegroundColor $modeColor

# Admin check
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ''
    Warn 'Not running as Administrator. WinBio session open may fail with access denied.'
}

# ---------- 1. PnP ----------
Section '[1] Biometric hardware (PnP class: Biometric)'
try {
    $devices = @(Get-PnpDevice -Class Biometric -ErrorAction Stop)
    if (-not $devices) {
        Fail 'No biometric devices found. The PC has no fingerprint reader, or the driver is not installed.'
    }
    foreach ($d in $devices) {
        switch ($d.Status) {
            'OK'    { Pass $d.FriendlyName ; Info "InstanceId: $($d.InstanceId)" }
            'Error' { Fail "$($d.FriendlyName) -- driver error" }
            default { Warn "$($d.FriendlyName) -- status $($d.Status)" }
        }
    }
} catch {
    Fail "PnP query failed: $($_.Exception.Message)"
}

# ---------- 2. WbioSrvc ----------
Section '[2] Windows Biometric Service (WbioSrvc)'
$svc = Get-Service -Name WbioSrvc -ErrorAction SilentlyContinue
if (-not $svc) {
    Fail 'Service WbioSrvc not found.'
} elseif ($svc.Status -eq 'Running') {
    Pass "Running (start type: $($svc.StartType))"
} else {
    Warn "Status: $($svc.Status). Attempting to start..."
    try { Start-Service WbioSrvc -ErrorAction Stop; Pass 'Service started.' }
    catch { Fail "Could not start service: $($_.Exception.Message)" }
}

# ---------- 3. WinBio enumeration ----------
Section '[3] Biometric units (WinBioEnumBiometricUnits)'
$unitArr = [IntPtr]::Zero
$unitCnt = [IntPtr]::Zero
try {
    $hr = [WinBio]::WinBioEnumBiometricUnits([WinBio]::TYPE_FINGERPRINT, [ref] $unitArr, [ref] $unitCnt)
    if ($hr -ne 0) {
        Fail "WinBioEnumBiometricUnits: $(Get-WinBioError $hr)"
    } else {
        $count = $unitCnt.ToInt64()
        if ($count -eq 0) {
            Fail 'WBF reports zero fingerprint units.'
        } else {
            $size = [System.Runtime.InteropServices.Marshal]::SizeOf([type][WinBio+WINBIO_UNIT_SCHEMA])
            for ($i = 0; $i -lt $count; $i++) {
                $p = [IntPtr]::Add($unitArr, $i * $size)
                $u = [System.Runtime.InteropServices.Marshal]::PtrToStructure($p, [type][WinBio+WINBIO_UNIT_SCHEMA])
                Pass "Unit $($u.UnitId): $($u.Manufacturer) $($u.Model)"
                if ($u.Description)  { Info "Description : $($u.Description)" }
                if ($u.SerialNumber) { Info "Serial      : $($u.SerialNumber)" }
                Info "Firmware    : $($u.FirmwareMajor).$($u.FirmwareMinor)"
            }
        }
    }
} catch {
    Fail "Enumeration failed: $($_.Exception.Message)"
} finally {
    if ($unitArr -ne [IntPtr]::Zero) { [WinBio]::WinBioFree($unitArr) | Out-Null }
}

# ---------- 4. Open session ----------
Section '[4] Open biometric session'
$session = [IntPtr]::Zero
$hr = [WinBio]::WinBioOpenSession(
    [WinBio]::TYPE_FINGERPRINT, [WinBio]::POOL_SYSTEM, [WinBio]::FLAG_DEFAULT,
    [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [ref] $session)

if ($hr -ne 0) {
    Warn "SYSTEM pool: $(Get-WinBioError $hr)"
    Info 'Retrying with PRIVATE pool...'
    $hr = [WinBio]::WinBioOpenSession(
        [WinBio]::TYPE_FINGERPRINT, [WinBio]::POOL_PRIVATE, [WinBio]::FLAG_DEFAULT,
        [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [ref] $session)
}

if ($hr -ne 0) {
    Fail "WinBioOpenSession: $(Get-WinBioError $hr)"
    Write-Host ''
    Write-Host '  Press Enter to close...' -ForegroundColor Gray
    Read-Host | Out-Null
    return
}
Pass 'Session opened.'

$locUnit = 0
$hr = [WinBio]::WinBioLocateSensor($session, [ref] $locUnit)
if ($hr -eq 0) { Pass "Sensor located on unit $locUnit" }
else           { Warn "WinBioLocateSensor: $(Get-WinBioError $hr)" }

# ---------- 5. Live capture ----------
Section '[5] Live capture test'
Write-Host ''
Write-Host '  Place your finger on the sensor when ready.' -ForegroundColor Yellow
Write-Host '  Press Esc to cancel.'                         -ForegroundColor Yellow
Write-Host ''

# Run capture in a background runspace so we can listen for Esc.
# The WinBio type lives in the shared AppDomain, so it is visible inside the runspace.
$captureScript = {
    param($sess)
    $unit = 0; $sample = [IntPtr]::Zero; $size = [IntPtr]::Zero; $reject = 0
    $hr = [WinBio]::WinBioCaptureSample(
        $sess, [WinBio]::PURPOSE_VERIFY, [WinBio]::FLAG_DEFAULT,
        [ref] $unit, [ref] $sample, [ref] $size, [ref] $reject)
    if ($sample -ne [IntPtr]::Zero) { [WinBio]::WinBioFree($sample) | Out-Null }
    [pscustomobject]@{ HResult = $hr; UnitId = $unit; SampleSize = $size.ToInt64(); Reject = $reject }
}

$ps = [PowerShell]::Create()
$null = $ps.AddScript($captureScript).AddArgument($session)
$async = $ps.BeginInvoke()

$cancelled   = $false
$tick        = 0
$canReadKeys = $true
try { $null = [Console]::KeyAvailable } catch { $canReadKeys = $false }

Write-Host '  Waiting' -NoNewline -ForegroundColor DarkGray
while (-not $async.IsCompleted) {
    if ($canReadKeys) {
        try {
            if ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true)
                if ($k.Key -eq 'Escape') {
                    [WinBio]::WinBioCancel($session) | Out-Null
                    $cancelled = $true
                    break
                }
            }
        } catch {}
    }
    Start-Sleep -Milliseconds 250
    $tick++
    if ($tick % 4 -eq 0) { Write-Host '.' -NoNewline -ForegroundColor DarkGray }
}

$captureResult = $ps.EndInvoke($async) | Select-Object -First 1
$ps.Dispose()
Write-Host ''
Write-Host ''

if ($cancelled) {
    Warn 'Cancelled by user.'
} elseif ($captureResult.HResult -eq 0) {
    Pass "Capture succeeded on unit $($captureResult.UnitId). Sample size: $($captureResult.SampleSize) bytes."
    Pass 'The fingerprint reader is WORKING.'
} else {
    Fail "Capture: $(Get-WinBioError $captureResult.HResult)"
    if ($captureResult.Reject -ne 0) { Info "Reject detail code: $($captureResult.Reject)" }
}

[WinBio]::WinBioCloseSession($session) | Out-Null

# ---------- done ----------
Write-Host ''
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host '     Diagnostic complete'                                -ForegroundColor Cyan
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host ''
try {
    Write-Host '  Press any key to close...' -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} catch {
    Write-Host '  Press Enter to close...' -ForegroundColor Gray
    Read-Host | Out-Null
}