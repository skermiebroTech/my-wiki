<#
.SYNOPSIS
    Windows 11 OOBE bypass. Creates a local administrator, marks setup complete, and reboots.

.DESCRIPTION
    Run this from the Shift+F10 command prompt on any Windows 11 OOBE screen.
    The script does four things:

      1. Creates a local account (default name "user") with a blank password
         and adds it to the local Administrators group.
      2. Writes the registry values that tell Windows setup is finished, so
         OOBE does not run again. No Microsoft account, no network, no
         privacy screens, no region or keyboard screens.
      3. Deletes the leftover defaultuser0 placeholder account.
      4. Reboots. The machine comes up at the normal sign-in screen.

    The account has NO PASSWORD. Anyone at the keyboard gets administrator
    rights. Use this on bench machines only. Set a password before the
    machine leaves the workshop.

.PARAMETER UserName
    Name of the local administrator to create. Default: user

.PARAMETER AutoLogon
    Also enable automatic sign-in, so the machine boots straight to the
    desktop with no click. Off by default.

.PARAMETER NoReboot
    Apply the changes but do not reboot.

.PARAMETER Force
    Allow the script to run when Windows is NOT in OOBE. Without this switch
    the script refuses to run on a machine that is already set up.

.NOTES
    Launch one-liner (Shift+F10 at any OOBE screen, network available):

        powershell -ep bypass -nop -c "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Bypass-OOBE.ps1 | iex"

    With parameters (iex cannot take parameters, so use a script block):

        powershell -ep bypass -nop -c "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Bypass-OOBE.ps1))) -AutoLogon"

    Offline (no network at OOBE). Put Bypass-OOBE.cmd on a USB stick and run:

        X:\Bypass-OOBE.cmd
#>

[CmdletBinding()]
param(
    [string]$UserName = 'user',
    [switch]$AutoLogon,
    [switch]$NoReboot,
    [switch]$Force
)

$SCRIPT_VERSION = '1.2.0'
$ErrorActionPreference = 'Continue'
try { $Host.UI.RawUI.WindowTitle = 'OOBE Bypass' } catch {}

# ---------- helpers ----------
function Pass    { param($t) Write-Host '  [ OK ] ' -ForegroundColor Green  -NoNewline; Write-Host $t }
function Warn    { param($t) Write-Host '  [WARN] ' -ForegroundColor Yellow -NoNewline; Write-Host $t }
function Fail    { param($t) Write-Host '  [FAIL] ' -ForegroundColor Red    -NoNewline; Write-Host $t }
function Info    { param($t) Write-Host '  [INFO] ' -ForegroundColor Gray   -NoNewline; Write-Host $t }
function Section { param($t) Write-Host ''; Write-Host $t -ForegroundColor Cyan }

function Set-Reg {
    param($Path, $Name, $Value, $Type)
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction Stop | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop | Out-Null
}

Write-Host ''
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host "     Windows 11 OOBE Bypass  v$SCRIPT_VERSION"          -ForegroundColor Cyan
Write-Host '  ===================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Identity : $([Security.Principal.WindowsIdentity]::GetCurrent().Name)" -ForegroundColor Gray
Write-Host "  Date     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"                  -ForegroundColor Gray
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if ($os) { Write-Host "  OS       : $($os.Caption) ($($os.Version))" -ForegroundColor Gray }

# ---------- 0. checks ----------
Section '  [0] Checks'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Fail 'Not running as Administrator. Use the Shift+F10 prompt at OOBE, or an elevated console.'
    return
}
Pass 'Running with administrator rights.'

# OOBE detection.
# Windows 11 24H2 and LTSC 26100 clear SystemSetupInProgress and set ImageState
# to IMAGE_STATE_COMPLETE BEFORE the OOBE pages appear. Those two values alone
# report "already set up" in the middle of OOBE. Test several signals instead
# and accept any one of them.
$setup    = Get-ItemProperty 'HKLM:\SYSTEM\Setup' -ErrorAction SilentlyContinue
$imgState = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State' -ErrorAction SilentlyContinue).ImageState
$who      = [Security.Principal.WindowsIdentity]::GetCurrent().Name

# Windows deletes defaultuser0 when OOBE finishes, so it is a strong signal.
& net.exe user defaultuser0 2>$null | Out-Null
$hasDefaultUser0 = ($LASTEXITCODE -eq 0)

$signals = [ordered]@{
    'identity is defaultuser0'      = ($who -like '*\defaultuser0')
    'defaultuser0 account exists'   = $hasDefaultUser0
    'SystemSetupInProgress = 1'     = ($setup.SystemSetupInProgress -eq 1)
    'OOBEInProgress = 1'            = ($setup.OOBEInProgress -eq 1)
    'Setup CmdLine starts OOBE'     = ($setup.CmdLine -match 'windeploy|oobe')
    'ImageState is not complete'    = ([bool]$imgState -and $imgState -ne 'IMAGE_STATE_COMPLETE')
}

foreach ($k in $signals.Keys) {
    if ($signals[$k]) { Pass "OOBE signal: $k" }
}
Info "ImageState='$imgState' SystemSetupInProgress='$($setup.SystemSetupInProgress)' CmdLine='$($setup.CmdLine)'"

$inOobe = @($signals.Values | Where-Object { $_ }).Count -gt 0

if ($inOobe) {
    Pass 'Windows is in OOBE.'
} elseif ($Force) {
    Warn 'Windows is already set up. -Force given, so the script continues.'
} else {
    Fail 'Windows is already set up. OOBE is not running.'
    Info 'This script creates a passwordless administrator. Add -Force only if you want that here.'
    return
}

# ---------- 1. local administrator ----------
Section '  [1] Local administrator'

# Resolve the Administrators group by SID, so a non-English Windows still works.
try {
    $adminGroup = ([Security.Principal.SecurityIdentifier]'S-1-5-32-544').Translate(
                    [Security.Principal.NTAccount]).Value -replace '^.*\\'
} catch {
    $adminGroup = 'Administrators'
}

& net.exe user $UserName 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    Warn "Account '$UserName' already exists. The script clears its password."
    try {
        Set-LocalUser -Name $UserName -Password (New-Object System.Security.SecureString) -ErrorAction Stop
    } catch {
        # PowerShell drops an empty '' argument to a native command, so quote it.
        & net.exe user $UserName '""' 2>$null | Out-Null
    }
} else {
    & net.exe user $UserName /add /passwordreq:no 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Fail "Could not create '$UserName'. A password policy may block a blank password."
        return
    }
    Pass "Created account '$UserName' with a blank password."
}

& net.exe localgroup $adminGroup $UserName /add 2>$null | Out-Null
$addExit = $LASTEXITCODE

# Confirm the membership instead of trusting the exit code.
$inAdmins = $null
try {
    $inAdmins = [bool](Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop |
                       Where-Object { $_.Name -like "*\$UserName" })
} catch {
    # Get-LocalGroupMember throws on orphaned SIDs. Fall back to the exit code.
    # net.exe returns 2 when the account is already a member.
    $inAdmins = ($addExit -eq 0 -or $addExit -eq 2)
}
if ($inAdmins) {
    Pass "'$UserName' is a member of $adminGroup."
} else {
    Fail "Could not add '$UserName' to $adminGroup. net.exe exit code $addExit."
    return
}

try {
    Set-LocalUser -Name $UserName -PasswordNeverExpires $true -ErrorAction Stop
    Pass 'Password set to never expire.'
} catch {
    Warn 'Could not set the password expiry flag. Not a problem for a blank password.'
}

# ---------- 2. mark setup complete ----------
Section '  [2] Mark OOBE complete'

$values = @(
    # Tell Windows setup finished, so windeploy.exe does not start OOBE again.
    @{ P = 'HKLM:\SYSTEM\Setup'; N = 'CmdLine';               V = '';  T = 'String' }
    @{ P = 'HKLM:\SYSTEM\Setup'; N = 'SetupType';             V = 0;   T = 'DWord'  }
    @{ P = 'HKLM:\SYSTEM\Setup'; N = 'OOBEInProgress';        V = 0;   T = 'DWord'  }
    @{ P = 'HKLM:\SYSTEM\Setup'; N = 'SystemSetupInProgress'; V = 0;   T = 'DWord'  }
    @{ P = 'HKLM:\SYSTEM\Setup\Status\ChildCompletion'; N = 'setup.exe'; V = 3; T = 'DWord' }
    # Sysprep state, so a sysprepped image does not reseal back into OOBE.
    @{ P = 'HKLM:\SYSTEM\Setup\Status\SysprepStatus'; N = 'GeneralizationState'; V = 7; T = 'DWord' }
    @{ P = 'HKLM:\SYSTEM\Setup\Status\SysprepStatus'; N = 'CleanupState';        V = 2; T = 'DWord' }
    # Extra safety, in case any OOBE page still appears.
    @{ P = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE'; N = 'UnattendCreatedUser'; V = 1; T = 'DWord' }
    @{ P = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE';       N = 'BypassNRO';           V = 1; T = 'DWord' }
    @{ P = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE';             N = 'DisablePrivacyExperience'; V = 1; T = 'DWord' }
    # Skip the "Hi, we are setting things up" first sign-in animation.
    @{ P = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; N = 'EnableFirstLogonAnimation'; V = 0; T = 'DWord' }
)

$written = 0
foreach ($v in $values) {
    try { Set-Reg -Path $v.P -Name $v.N -Value $v.V -Type $v.T; $written++ }
    catch { Warn "Could not write $($v.P)\$($v.N): $($_.Exception.Message)" }
}
Pass "Wrote $written of $($values.Count) registry values."

# Also mark the current image complete, so nothing re-enters OOBE.
try {
    Set-Reg -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State' `
            -Name 'ImageState' -Value 'IMAGE_STATE_COMPLETE' -Type 'String'
    Pass 'ImageState set to IMAGE_STATE_COMPLETE.'
} catch {
    Warn 'Could not set ImageState.'
}

# ---------- 3. automatic sign-in ----------
Section '  [3] Automatic sign-in'

$winlogon = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
if ($AutoLogon) {
    Set-Reg -Path $winlogon -Name 'AutoAdminLogon'    -Value '1'             -Type 'String'
    Set-Reg -Path $winlogon -Name 'DefaultUserName'   -Value $UserName       -Type 'String'
    Set-Reg -Path $winlogon -Name 'DefaultPassword'   -Value ''              -Type 'String'
    Set-Reg -Path $winlogon -Name 'DefaultDomainName' -Value $env:COMPUTERNAME -Type 'String'
    Remove-ItemProperty -Path $winlogon -Name 'AutoLogonCount' -ErrorAction SilentlyContinue
    Pass "Automatic sign-in enabled for '$UserName'."
} else {
    Info 'Not enabled. Add -AutoLogon to boot straight to the desktop.'
}

# ---------- 4. clean up ----------
Section '  [4] Clean up'

& net.exe user defaultuser0 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    & net.exe user defaultuser0 /delete 2>$null | Out-Null
    Pass 'Removed the defaultuser0 placeholder account.'
} else {
    Info 'No defaultuser0 account to remove.'
}

# ---------- 5. reboot ----------
Section '  [5] Reboot'

Write-Host ''
Write-Host "  Sign in as '$UserName' with an empty password." -ForegroundColor Green
Write-Host '  Set a password before this machine leaves the workshop.' -ForegroundColor Yellow
Write-Host ''

if ($NoReboot) {
    Info 'Reboot skipped. Run "shutdown /r /t 0" when you are ready.'
    return
}

for ($i = 5; $i -ge 1; $i--) {
    Write-Host "`r  Reboot in $i second(s). Press Ctrl+C to stop. " -ForegroundColor Cyan -NoNewline
    Start-Sleep -Seconds 1
}
Write-Host ''
& shutdown.exe /r /t 0 /f
