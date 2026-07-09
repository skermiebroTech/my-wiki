# =====================================================================
# Setup-TestMachine.ps1
#
# Bootstraps a Windows 11 machine for testing the driver installer
# (Install-Drivers-auto.ps1). Installs the dev/test toolchain:
#
#   1. Git for Windows      (winget Git.Git, direct-download fallback)
#   2. GitHub Desktop       (winget GitHub.GitHubDesktop, direct fallback)
#   3. Claude Code          (official installer: irm https://claude.ai/install.ps1)
#
# Then configures the git identity and clones the my-wiki repo so the
# driver installer scripts are on the machine ready to run.
#
# Usage (elevated PowerShell, fresh Win11 box):
#   Set-ExecutionPolicy Bypass -Scope Process -Force
#   .\Setup-TestMachine.ps1
#
# Parameters:
#   -SkipClone      Don't clone the my-wiki repo.
#   -GitUserName    git config --global user.name  (default: Skermiebro Tech Tips)
#   -GitUserEmail   git config --global user.email (default: GitHub noreply address)
#
# Notes:
#   - Self-elevates via UAC if not already admin. GitHub Desktop and
#     Claude Code install PER-USER, so run this logged in as the account
#     that will do the testing (fine on shop test machines that log in
#     as Administrator). Elevating from a DIFFERENT admin account would
#     put the per-user tools on the wrong profile.
#   - winget is used with --source winget to dodge msstore-source issues
#     on fresh images / built-in Administrator accounts. If winget is
#     absent or broken, Git and GitHub Desktop fall back to direct
#     downloads from their vendors.
#
# v1.0.0 - Initial version.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$SkipClone,
    [string]$GitUserName  = "Skermiebro Tech Tips",
    [string]$GitUserEmail = "67085932+skermiebroTech@users.noreply.github.com",
    [string]$RepoUrl      = "https://github.com/skermiebroTech/my-wiki.git",
    [string]$ClonePath    = (Join-Path $env:USERPROFILE "my-wiki")
)

$SCRIPT_VERSION = "1.0.0"
$ErrorActionPreference = "Stop"
# Fresh Win11 + Windows PowerShell 5.1 may still default to TLS 1.0 for .NET downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Log {
    param([string]$Msg)
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Msg)
}

# --- Self-elevate ----------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Log "Not elevated - relaunching with UAC prompt..."
    $argList = @("-NoProfile","-ExecutionPolicy","Bypass","-File","`"$PSCommandPath`"")
    if ($SkipClone) { $argList += "-SkipClone" }
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
    exit 0
}

Log "=== Test machine setup v$SCRIPT_VERSION ==="
$failures = @()

# --- PATH refresh helper (installers append PATH in the registry, not
#     in this session) --------------------------------------------------
function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user    = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machine;$user"
}

# --- winget availability ---------------------------------------------
function Test-Winget {
    try { winget --version 2>$null | Out-Null; return $true } catch { return $false }
}
$wingetOk = Test-Winget
if (-not $wingetOk) {
    # Fresh images sometimes ship App Installer unregistered for this user
    Log "winget not found - attempting to re-register App Installer..."
    try {
        Get-AppxPackage Microsoft.DesktopAppInstaller |
            ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }
        $wingetOk = Test-Winget
    } catch {}
    if ($wingetOk) { Log "  winget registered OK." }
    else { Log "  winget unavailable - will use direct-download fallbacks." }
}

function Install-ViaWinget {
    param([string]$Id, [string]$Display)
    if (-not $wingetOk) { return $false }
    Log "Installing $Display via winget ($Id)..."
    & winget install --id $Id --exact --source winget --silent `
        --accept-source-agreements --accept-package-agreements
    if ($LASTEXITCODE -eq 0) { return $true }
    # -1978335189 = APPINSTALLER_CLI_ERROR_UPDATE_NOT_APPLICABLE (already installed / no update)
    if ($LASTEXITCODE -eq -1978335189) { Log "  Already installed."; return $true }
    Log "  winget install failed (exit $LASTEXITCODE) - trying direct download."
    return $false
}

# --- 1. Git for Windows ----------------------------------------------
if (Get-Command git -EA SilentlyContinue) {
    Log "Git already installed: $(git --version)"
} else {
    $ok = Install-ViaWinget -Id "Git.Git" -Display "Git for Windows"
    if (-not $ok) {
        Log "Downloading Git for Windows from GitHub releases..."
        try {
            $rel = Invoke-RestMethod "https://api.github.com/repos/git-for-windows/git/releases/latest"
            $asset = $rel.assets | Where-Object { $_.name -match '^Git-.*-64-bit\.exe$' } | Select-Object -First 1
            $exe = Join-Path $env:TEMP $asset.name
            Invoke-WebRequest $asset.browser_download_url -OutFile $exe
            Log "  Running $($asset.name) silently..."
            Start-Process $exe -ArgumentList "/VERYSILENT","/NORESTART" -Wait
            $ok = $true
        } catch { Log "  Direct download failed: $($_.Exception.Message)" }
    }
    Update-SessionPath
    if ($ok -and (Get-Command git -EA SilentlyContinue)) {
        Log "  Git installed: $(git --version)"
    } else {
        $failures += "Git"
    }
}

# --- 2. GitHub Desktop (per-user Squirrel installer) ------------------
$ghDesktopExe = Join-Path $env:LOCALAPPDATA "GitHubDesktop\GitHubDesktop.exe"
if (Test-Path $ghDesktopExe) {
    Log "GitHub Desktop already installed."
} else {
    $ok = Install-ViaWinget -Id "GitHub.GitHubDesktop" -Display "GitHub Desktop"
    if (-not $ok) {
        Log "Downloading GitHub Desktop directly..."
        try {
            $exe = Join-Path $env:TEMP "GitHubDesktopSetup-x64.exe"
            Invoke-WebRequest "https://central.github.com/deployments/desktop/desktop/latest/win32?format=exe&arch=x64" -OutFile $exe
            Log "  Running installer silently..."
            Start-Process $exe -ArgumentList "-s" -Wait
            $ok = $true
        } catch { Log "  Direct download failed: $($_.Exception.Message)" }
    }
    # Squirrel finishes the first install asynchronously - give it a moment
    $deadline = (Get-Date).AddSeconds(90)
    while (-not (Test-Path $ghDesktopExe) -and (Get-Date) -lt $deadline) { Start-Sleep 3 }
    if (Test-Path $ghDesktopExe) { Log "  GitHub Desktop installed." }
    else {
        Log "  GitHub Desktop exe not found yet - it may still be finishing; check the Start menu."
        if (-not $ok) { $failures += "GitHub Desktop" }
    }
}

# --- 3. Claude Code (official native installer, per-user) -------------
Update-SessionPath
if (Get-Command claude -EA SilentlyContinue) {
    Log "Claude Code already installed: $(claude --version)"
} else {
    Log "Installing Claude Code (irm https://claude.ai/install.ps1)..."
    try {
        Invoke-RestMethod "https://claude.ai/install.ps1" | Invoke-Expression
    } catch { Log "  Claude Code installer failed: $($_.Exception.Message)" }
    Update-SessionPath
    # Installer puts the binary in %USERPROFILE%\.local\bin
    $claudeBin = Join-Path $env:USERPROFILE ".local\bin"
    if (-not (Get-Command claude -EA SilentlyContinue) -and (Test-Path (Join-Path $claudeBin "claude.exe"))) {
        $env:Path = "$env:Path;$claudeBin"
    }
    if (Get-Command claude -EA SilentlyContinue) {
        Log "  Claude Code installed: $(claude --version)"
        Log "  NOTE: run 'claude' once to log in to your Anthropic account."
    } else {
        $failures += "Claude Code"
    }
}

# --- 4. Git identity ---------------------------------------------------
if (Get-Command git -EA SilentlyContinue) {
    if (-not (git config --global user.name 2>$null)) {
        git config --global user.name  $GitUserName
        git config --global user.email $GitUserEmail
        Log "Configured git identity: $GitUserName <$GitUserEmail>"
    } else {
        Log "Git identity already configured: $(git config --global user.name)"
    }
}

# --- 5. Clone the repo -------------------------------------------------
if (-not $SkipClone -and (Get-Command git -EA SilentlyContinue)) {
    if (Test-Path (Join-Path $ClonePath ".git")) {
        Log "Repo already cloned at $ClonePath - pulling latest..."
        git -C $ClonePath pull --ff-only
    } else {
        Log "Cloning $RepoUrl -> $ClonePath ..."
        git clone $RepoUrl $ClonePath
    }
    Log "Driver installer scripts are in: $ClonePath"
    Log "  (pushing changes requires signing in via GitHub Desktop first)"
}

# --- Summary -----------------------------------------------------------
Log "=== Setup complete ==="
if ($failures.Count -gt 0) {
    Log "FAILED: $($failures -join ', ') - install manually."
    exit 1
}
Log "All tools installed OK. Sign in to GitHub Desktop and run 'claude' to authenticate."
exit 0
