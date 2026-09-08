@echo off
REM ===================================================================
REM  Bypass-OOBE.cmd  v1.2.0
REM  Windows 11 OOBE bypass. No network needed.
REM
REM  Put this file on a USB stick. At any Windows 11 OOBE screen press
REM  Shift+F10 and run it:
REM
REM      X:\Bypass-OOBE.cmd
REM
REM  Optional first argument sets the account name (default: user).
REM  Add /autologon to boot straight to the desktop.
REM  Add /force to run on a machine that is already set up.
REM
REM      X:\Bypass-OOBE.cmd tech /autologon
REM
REM  WARNING: the account has NO PASSWORD and is a local administrator.
REM  Set a password before the machine leaves the workshop.
REM ===================================================================

setlocal EnableExtensions

set "ACCT=user"
set "AUTOLOGON=0"
set "FORCE=0"
set "SIGNALS=0"

REM ---- read arguments ----
REM Read the argument before the shift. A %~1 inside a parenthesised block
REM freezes at parse time, so the shift must not sit inside the block.
:parse
if "%~1"=="" goto parsed
set "ARG=%~1"
shift
if /i "%ARG%"=="/autologon" (set "AUTOLOGON=1") else if /i "%ARG%"=="/force" (set "FORCE=1") else (set "ACCT=%ARG%")
goto parse
:parsed

echo.
echo   ===================================================
echo      Windows 11 OOBE Bypass  v1.2.0
echo   ===================================================
echo.
echo   Account: %ACCT%  (no password, local administrator)
echo.

REM ---- 0. is Windows in OOBE? ----
REM Windows 11 24H2 and LTSC 26100 clear SystemSetupInProgress and set
REM ImageState to IMAGE_STATE_COMPLETE BEFORE the OOBE pages appear. Those two
REM values alone report "already set up" in the middle of OOBE. Count several
REM signals instead and accept any one of them.
echo   [0] Checks

if /i "%USERNAME%"=="defaultuser0" call :hit "identity is defaultuser0"

REM Windows deletes defaultuser0 when OOBE finishes, so it is a strong signal.
net user defaultuser0 >nul 2>&1
if not errorlevel 1 call :hit "defaultuser0 account exists"

reg query "HKLM\SYSTEM\Setup" /v SystemSetupInProgress 2>nul | find "0x1" >nul
if not errorlevel 1 call :hit "SystemSetupInProgress = 1"

reg query "HKLM\SYSTEM\Setup" /v OOBEInProgress 2>nul | find "0x1" >nul
if not errorlevel 1 call :hit "OOBEInProgress = 1"

reg query "HKLM\SYSTEM\Setup" /v CmdLine 2>nul | findstr /i "windeploy oobe" >nul
if not errorlevel 1 call :hit "Setup CmdLine starts OOBE"

REM ImageState counts only when the value exists and is not COMPLETE.
set "IMGDONE="
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState 2>nul | find /i "IMAGE_STATE_COMPLETE" >nul
if not errorlevel 1 set "IMGDONE=1"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState 2>nul | find /i "IMAGE_STATE" >nul
if errorlevel 1 set "IMGDONE=1"
if not defined IMGDONE call :hit "ImageState is not complete"

if %SIGNALS% GTR 0 goto oobe_ok
if "%FORCE%"=="1" (
    echo   [WARN] Windows is already set up. /force given, so the script continues.
    goto oobe_ok
)
echo   [FAIL] Windows is already set up. OOBE is not running.
echo          This script creates a passwordless administrator.
echo          Add /force only if you want that here.
echo.
pause
exit /b 1
:oobe_ok
echo.

REM ---- 1. create the local administrator ----
REM "Administrators" is the English group name. Change it on a localised image.
net user "%ACCT%" /add /passwordreq:no >nul 2>&1
if errorlevel 1 net user "%ACCT%" "" >nul 2>&1
net localgroup Administrators "%ACCT%" /add >nul 2>&1

REM Stop here if the account does not exist, so we never reboot into nothing.
net user "%ACCT%" >nul 2>&1
if errorlevel 1 (
    echo   [FAIL] Could not create the account "%ACCT%".
    echo          A password policy may block a blank password.
    pause
    exit /b 1
)
echo   [ OK ] Local administrator ready.

REM ---- 2. mark setup complete ----
REM An omitted /d writes an empty REG_SZ, which clears oobe\windeploy.exe.
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /f >nul
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f >nul
reg add "HKLM\SYSTEM\Setup" /v OOBEInProgress /t REG_DWORD /d 0 /f >nul
reg add "HKLM\SYSTEM\Setup" /v SystemSetupInProgress /t REG_DWORD /d 0 /f >nul
reg add "HKLM\SYSTEM\Setup\Status\ChildCompletion" /v setup.exe /t REG_DWORD /d 3 /f >nul
reg add "HKLM\SYSTEM\Setup\Status\SysprepStatus" /v GeneralizationState /t REG_DWORD /d 7 /f >nul
reg add "HKLM\SYSTEM\Setup\Status\SysprepStatus" /v CleanupState /t REG_DWORD /d 2 /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState /t REG_SZ /d IMAGE_STATE_COMPLETE /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE" /v UnattendCreatedUser /t REG_DWORD /d 1 /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f >nul
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\OOBE" /v DisablePrivacyExperience /t REG_DWORD /d 1 /f >nul
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableFirstLogonAnimation /t REG_DWORD /d 0 /f >nul
echo   [ OK ] OOBE marked complete.

REM ---- 3. automatic sign-in ----
if "%AUTOLOGON%"=="1" (
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f >nul
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d "%ACCT%" /f >nul
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /f >nul
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultDomainName /t REG_SZ /d "%COMPUTERNAME%" /f >nul
    reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoLogonCount /f >nul 2>&1
    echo   [ OK ] Automatic sign-in enabled.
)

REM ---- 4. clean up ----
net user defaultuser0 >nul 2>&1
if not errorlevel 1 (
    net user defaultuser0 /delete >nul 2>&1
    echo   [ OK ] Removed the defaultuser0 placeholder.
)

echo.
echo   Sign in as "%ACCT%" with an empty password.
echo   Set a password before this machine leaves the workshop.
echo.
echo   Reboot in 5 seconds. Press Ctrl+C to stop.
shutdown /r /t 5 /f
endlocal
exit /b 0

REM ---- subroutine: record one OOBE signal ----
:hit
set /a SIGNALS+=1
echo   [ OK ] OOBE signal: %~1
goto :eof
