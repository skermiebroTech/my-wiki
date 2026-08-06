@echo off
setlocal

set "DUCKY=%~d0"

REM === Switch the default audio output to a device with "speaker" in its name ===
REM The PowerShell that does this lives at the bottom of this file, after the
REM marker line. The command below reads this file, grabs everything after the
REM last marker, and runs it. cmd never executes those lines itself because of
REM the "exit /b" further down.
REM Fire-and-forget: launched in a hidden background process so systembuddy
REM starts instantly and the audio switch happens on its own.
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Invoke-Expression (((Get-Content -LiteralPath '%~f0' -Raw) -split ':PSBODY:')[-1])"

REM Check every drive except the Ducky first
for %%D in (D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if /I not "%%D:"=="%DUCKY%" (
        if exist "%%D:\systembuddy.exe" (
            cd /d "%%D:\"
            start "" "%%D:\systembuddy.exe"
            exit /b
        )
    )
)

REM Fall back to the Ducky copy
if exist "%~dp0systembuddy.exe" (
    cd /d "%~dp0"
    start "" "%~dp0systembuddy.exe"
    exit /b
)

echo systembuddy.exe was not found.
pause
exit /b

:PSBODY:
# ---------------------------------------------------------------------------
# Set the default audio render (output) device whose name contains "speaker".
# Uses the built-in Windows IPolicyConfig COM interface - no extra tools and no
# administrator rights required.
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[Guid("f8679f50-850a-41cf-9c72-430f290290c8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPolicyConfig
{
    [PreserveSig] int GetMixFormat();
    [PreserveSig] int GetDeviceFormat();
    [PreserveSig] int ResetDeviceFormat();
    [PreserveSig] int SetDeviceFormat();
    [PreserveSig] int GetProcessingPeriod();
    [PreserveSig] int SetProcessingPeriod();
    [PreserveSig] int GetShareMode();
    [PreserveSig] int SetShareMode();
    [PreserveSig] int GetPropertyValue();
    [PreserveSig] int SetPropertyValue();
    [PreserveSig] int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int role);
    [PreserveSig] int SetEndpointVisibility();
}

[ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
public class CPolicyConfigClient { }
'@

$base  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
$kName = '{a45c254e-df1c-4efd-8020-67d146a850e0},14'  # full friendly name, e.g. "Speakers (Realtek Audio)"
$kDesc = '{a45c254e-df1c-4efd-8020-67d146a850e0},2'   # short description,   e.g. "Speakers"
$target = $null

foreach ($dev in Get-ChildItem $base -ErrorAction SilentlyContinue) {
    $state = (Get-ItemProperty -Path $dev.PSPath -Name DeviceState -ErrorAction SilentlyContinue).DeviceState
    if ($state -ne 1) { continue }   # 1 = active/enabled endpoint
    $props = Get-ItemProperty -Path (Join-Path $dev.PSPath 'Properties') -ErrorAction SilentlyContinue
    $name = $props.$kName
    $desc = $props.$kDesc
    if (($name -match 'speaker') -or ($desc -match 'speaker')) {
        $target = [pscustomobject]@{
            Id   = '{0.0.0.00000000}.' + $dev.PSChildName
            Name = if ($name) { $name } else { $desc }
        }
        break
    }
}

if ($target) {
    $cfg = [IPolicyConfig](New-Object CPolicyConfigClient)
    foreach ($role in 0, 1, 2) { [void]$cfg.SetDefaultEndpoint($target.Id, $role) }  # Console, Multimedia, Communications
    Write-Host "Default audio output set to: $($target.Name)"
} else {
    Write-Host "No active output device with 'speaker' in the name was found."
}
