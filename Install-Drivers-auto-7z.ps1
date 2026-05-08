# =============================================================
# Install-Drivers-auto.ps1
# Version: 1.5.6
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode (GUI):
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
#
# Run headlessly with arguments:
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer Dell
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer HP -Model "EliteBook x360 1030 G8 Notebook PC"
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer Lenovo -MachineType 20XX -SkipInstall -SkipCleanup
#
# Parameters:
#   -Manufacturer  Override WMI manufacturer detection (Dell, HP, Lenovo, Microsoft)
#   -Model         Override WMI model detection
#   -Headless      Skip GUI, write to console only (auto-set when any param is passed)
#   -SkipInstall   Download and extract only, skip pnputil driver installation
#   -SkipCleanup   Keep C:\DRIVERS after run for inspection
#
# Supports: Dell, HP, Lenovo, Microsoft (Surface)
#
# v1.5.6 - Added -MachineType param for Lenovo machine type override
# v1.5.5 - Added headless/parameter mode for testing and automation
# v1.5.4 - 7-Zip integration for Dell and HP extraction
#   Dell: 7-Zip pass-1 replaces /s /e= (verified identical output, 1.1x faster)
#   HP:   7-Zip pass-1 replaces /s /e /f (verified identical output, 4.7x faster)
#   Lenovo: unchanged - Inno Setup proprietary format, 7-Zip cannot extract
#   7-Zip is installed silently at start and removed before cleanup
# =============================================================

param(
    [string]$Manufacturer = "",
    [string]$Model        = "",
    [string]$MachineType  = "",   # Lenovo only: override 4-char machine type prefix (e.g. 20XX)
    [switch]$Headless,
    [switch]$SkipInstall,
    [switch]$SkipCleanup
)

# Auto-enable headless when any override param is passed
if ($Manufacturer -or $Model -or $MachineType -or $SkipInstall -or $SkipCleanup) { $Headless = $true }

$ScriptVersion   = "1.5.6"
$SpinnerFrames   = @('⠋','⠙','⠹','⠸','⠼','⠴','⠦','⠧','⠇','⠏')
$SpinnerIndex    = 0
$CancelRequested = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# Set timezone to Brisbane (UTC+10, no DST) and sync clock
# Runs before log file creation so the filename timestamp is correct
tzutil /s "E. Australia Standard Time"
Start-Service w32tm -ErrorAction SilentlyContinue
w32tm /resync /force | Out-Null

# =========================
# LOG FILE SETUP
# =========================
$LogFile = Join-Path ([Environment]::GetFolderPath("UserProfile")) `
    ("Downloads\DriverInstaller_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
New-Item -ItemType File -Path $LogFile -Force | Out-Null

# =========================
# FONT CONSTANTS
# =========================
$FontMono      = New-Object System.Drawing.Font("Courier New", 9,   [System.Drawing.FontStyle]::Regular)
$FontMonoSm    = New-Object System.Drawing.Font("Courier New", 8,   [System.Drawing.FontStyle]::Regular)
$FontUI        = New-Object System.Drawing.Font("Segoe UI",    9,   [System.Drawing.FontStyle]::Regular)
$FontUIBold    = New-Object System.Drawing.Font("Segoe UI",    9,   [System.Drawing.FontStyle]::Bold)
$FontUIBoldSm  = New-Object System.Drawing.Font("Segoe UI",    8,   [System.Drawing.FontStyle]::Bold)
$FontUISmall   = New-Object System.Drawing.Font("Segoe UI",    7.5, [System.Drawing.FontStyle]::Regular)
$FontTitleBold = New-Object System.Drawing.Font("Segoe UI",    13,  [System.Drawing.FontStyle]::Bold)

# =========================
# FORM SETUP
# =========================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Driver Installer Tool  v$ScriptVersion"
$form.Size            = New-Object System.Drawing.Size(580, 560)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 245)

$title           = New-Object System.Windows.Forms.Label
$title.AutoSize  = $true
$title.Font      = $FontTitleBold
$title.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$title.Location  = New-Object System.Drawing.Point(20, 15)
$title.Text      = "Driver Installer"
$title.UseCompatibleTextRendering = $false
$form.Controls.Add($title)

$versionLabel           = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize  = $true
$versionLabel.Font      = $FontUISmall
$versionLabel.ForeColor = [System.Drawing.Color]::Gray
$versionLabel.Text      = "v$ScriptVersion"
$versionLabel.Location  = New-Object System.Drawing.Point(510, 20)
$versionLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($versionLabel)

$statusBox             = New-Object System.Windows.Forms.RichTextBox
$statusBox.Multiline   = $true
$statusBox.ScrollBars  = "Vertical"
$statusBox.Size        = New-Object System.Drawing.Size(536, 200)
$statusBox.Location    = New-Object System.Drawing.Point(20, 55)
$statusBox.ReadOnly    = $true
$statusBox.BackColor   = [System.Drawing.Color]::FromArgb(22, 22, 22)
$statusBox.ForeColor   = [System.Drawing.Color]::FromArgb(190, 255, 190)
$statusBox.Font        = $FontMono
$statusBox.BorderStyle = "FixedSingle"
$form.Controls.Add($statusBox)

# ---- DOWNLOAD group ----
$dlGroupBox           = New-Object System.Windows.Forms.GroupBox
$dlGroupBox.Text      = "Download"
$dlGroupBox.Font      = $FontUIBoldSm
$dlGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
$dlGroupBox.Size      = New-Object System.Drawing.Size(536, 68)
$dlGroupBox.Location  = New-Object System.Drawing.Point(20, 265)
$form.Controls.Add($dlGroupBox)

$dlSpinnerLabel           = New-Object System.Windows.Forms.Label
$dlSpinnerLabel.AutoSize  = $true
$dlSpinnerLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$dlSpinnerLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
$dlSpinnerLabel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$dlSpinnerLabel.Location  = New-Object System.Drawing.Point(72, 0)
$dlSpinnerLabel.Text      = ""
$dlSpinnerLabel.UseCompatibleTextRendering = $false
$dlGroupBox.Controls.Add($dlSpinnerLabel)

$dlBar                       = New-Object System.Windows.Forms.ProgressBar
$dlBar.Size                  = New-Object System.Drawing.Size(508, 18)
$dlBar.Location              = New-Object System.Drawing.Point(12, 20)
$dlBar.Style                 = "Marquee"
$dlBar.MarqueeAnimationSpeed = 25
$dlBar.Minimum               = 0
$dlBar.Maximum               = 100
$dlGroupBox.Controls.Add($dlBar)

$dlLabel           = New-Object System.Windows.Forms.Label
$dlLabel.AutoSize  = $false
$dlLabel.Size      = New-Object System.Drawing.Size(508, 17)
$dlLabel.Location  = New-Object System.Drawing.Point(12, 42)
$dlLabel.Font      = $FontMonoSm
$dlLabel.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$dlLabel.Text      = "Waiting..."
$dlLabel.UseCompatibleTextRendering = $false
$dlGroupBox.Controls.Add($dlLabel)

# ---- EXTRACT group ----
$exGroupBox           = New-Object System.Windows.Forms.GroupBox
$exGroupBox.Text      = "Extract"
$exGroupBox.Font      = $FontUIBoldSm
$exGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 140, 80)
$exGroupBox.Size      = New-Object System.Drawing.Size(536, 68)
$exGroupBox.Location  = New-Object System.Drawing.Point(20, 340)
$form.Controls.Add($exGroupBox)

$exSpinnerLabel           = New-Object System.Windows.Forms.Label
$exSpinnerLabel.AutoSize  = $true
$exSpinnerLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$exSpinnerLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 140, 80)
$exSpinnerLabel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$exSpinnerLabel.Location  = New-Object System.Drawing.Point(58, 0)
$exSpinnerLabel.Text      = ""
$exSpinnerLabel.UseCompatibleTextRendering = $false
$exGroupBox.Controls.Add($exSpinnerLabel)

$exBar                       = New-Object System.Windows.Forms.ProgressBar
$exBar.Size                  = New-Object System.Drawing.Size(508, 18)
$exBar.Location              = New-Object System.Drawing.Point(12, 20)
$exBar.Style                 = "Marquee"
$exBar.MarqueeAnimationSpeed = 30
$exBar.Minimum               = 0
$exBar.Maximum               = 100
$exGroupBox.Controls.Add($exBar)

$exLabel           = New-Object System.Windows.Forms.Label
$exLabel.AutoSize  = $false
$exLabel.Size      = New-Object System.Drawing.Size(508, 17)
$exLabel.Location  = New-Object System.Drawing.Point(12, 42)
$exLabel.Font      = $FontMonoSm
$exLabel.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$exLabel.Text      = "Waiting..."
$exLabel.UseCompatibleTextRendering = $false
$exGroupBox.Controls.Add($exLabel)

# ---- OVERALL group ----
$overallGroupBox           = New-Object System.Windows.Forms.GroupBox
$overallGroupBox.Text      = "Overall"
$overallGroupBox.Font      = $FontUIBoldSm
$overallGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$overallGroupBox.Size      = New-Object System.Drawing.Size(536, 48)
$overallGroupBox.Location  = New-Object System.Drawing.Point(20, 415)
$form.Controls.Add($overallGroupBox)

$overallSpinnerLabel           = New-Object System.Windows.Forms.Label
$overallSpinnerLabel.AutoSize  = $true
$overallSpinnerLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$overallSpinnerLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$overallSpinnerLabel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$overallSpinnerLabel.Location  = New-Object System.Drawing.Point(62, 0)
$overallSpinnerLabel.Text      = ""
$overallSpinnerLabel.UseCompatibleTextRendering = $false
$overallGroupBox.Controls.Add($overallSpinnerLabel)

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Size     = New-Object System.Drawing.Size(508, 18)
$progress.Location = New-Object System.Drawing.Point(12, 20)
$progress.Style    = "Continuous"
$progress.Minimum  = 0
$progress.Maximum  = 100
$overallGroupBox.Controls.Add($progress)

$logLabel           = New-Object System.Windows.Forms.Label
$logLabel.AutoSize  = $false
$logLabel.Size      = New-Object System.Drawing.Size(536, 16)
$logLabel.Location  = New-Object System.Drawing.Point(20, 468)
$logLabel.ForeColor = [System.Drawing.Color]::Gray
$logLabel.Font      = $FontUISmall
$logLabel.Text      = "Log: $LogFile"
$logLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($logLabel)

# =========================
# SOUND TOGGLE CHECKBOX
# =========================
$soundCheckbox                   = New-Object System.Windows.Forms.CheckBox
$soundCheckbox.Text              = "Sound FX"
$soundCheckbox.Checked           = $true
$soundCheckbox.Font              = $FontUIBold
$soundCheckbox.ForeColor         = [System.Drawing.Color]::FromArgb(60, 60, 60)
$soundCheckbox.AutoSize          = $true
$soundCheckbox.Location          = New-Object System.Drawing.Point(20, 490)
$soundCheckbox.UseCompatibleTextRendering = $false
$form.Controls.Add($soundCheckbox)

$button            = New-Object System.Windows.Forms.Button
$button.Text       = "Install Drivers"
$button.Size       = New-Object System.Drawing.Size(155, 36)
$button.Location   = New-Object System.Drawing.Point(155, 483)
$button.Font       = $FontUIBold
$button.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
$button.ForeColor  = [System.Drawing.Color]::White
$button.FlatStyle  = "Flat"
$button.FlatAppearance.BorderSize = 0
$form.Controls.Add($button)

$cancelButton            = New-Object System.Windows.Forms.Button
$cancelButton.Text       = "Cancel"
$cancelButton.Size       = New-Object System.Drawing.Size(100, 36)
$cancelButton.Location   = New-Object System.Drawing.Point(320, 483)
$cancelButton.Font       = $FontUIBold
$cancelButton.BackColor  = [System.Drawing.Color]::FromArgb(160, 160, 160)
$cancelButton.ForeColor  = [System.Drawing.Color]::White
$cancelButton.FlatStyle  = "Flat"
$cancelButton.FlatAppearance.BorderSize = 0
$cancelButton.Enabled    = $false
$form.Controls.Add($cancelButton)

# =========================
# SOUND HELPER
# =========================
function Play-Sound {
    param(
        [ValidateSet("Start","DownloadComplete","ExtractComplete","DriverAdded","Success","Failure","Cancel")]
        [string]$Event
    )
    if (-not $soundCheckbox.Checked) { return }
    $mediaDir = "$env:SystemRoot\Media"
    $wavCandidates = switch ($Event) {
        "Start"            { @("Windows Notify.wav", "Windows Notify System Generic.wav", "chimes.wav") }
        "DownloadComplete" { @("Windows Print complete.wav", "Windows Notify.wav", "chimes.wav") }
        "ExtractComplete"  { @("Windows Print complete.wav", "Windows Notify.wav", "chimes.wav") }
        "DriverAdded"      { @("Windows Navigation Start.wav", "Windows Notify Calendar.wav", "Windows Notify.wav") }
        "Success"          { @("Windows Logon.wav", "Windows Notify.wav", "tada.wav") }
        "Failure"          { @("Windows Critical Stop.wav", "Windows Foreground.wav", "chord.wav") }
        "Cancel"           { @("Windows Critical Stop.wav", "Windows Foreground.wav", "chord.wav") }
    }
    $wavFile = $null
    foreach ($candidate in $wavCandidates) {
        $path = Join-Path $mediaDir $candidate
        if (Test-Path $path) { $wavFile = $path; break }
    }
    if (-not $wavFile) { return }
    try { $player = New-Object System.Media.SoundPlayer $wavFile; $player.Play() } catch {}
}

# =========================
# BUTTON STATE HELPERS
# =========================
function Set-ButtonRunning {
    if ($script:Headless) { return }
    $button.Enabled         = $false
    $button.BackColor       = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $cancelButton.Enabled   = $true
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60)
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ButtonIdle {
    if ($script:Headless) { return }
    $button.Enabled         = $true
    $button.BackColor       = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $cancelButton.Enabled   = $false
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
    [System.Windows.Forms.Application]::DoEvents()
}

# =========================
# HELPERS
# =========================
function Log($msg) {
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $msg"
    if ($script:Headless) {
        Write-Host $line
    } else {
        $statusBox.AppendText("$line`r`n")
        $statusBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function SetProgress($val) {
    if ($script:Headless) { return }
    $progress.Value = [math]::Min([math]::Max([int]$val, 0), 100)
    [System.Windows.Forms.Application]::DoEvents()
}

function SetDownload {
    param([int]$Pct, [string]$Label)
    if ($script:Headless) { return }
    if ($Pct -ge 100) {
        $dlBar.Style = "Continuous"
        $dlBar.Value = 100
    } elseif ($dlBar.Style -ne "Marquee") {
        $dlBar.Style                 = "Marquee"
        $dlBar.MarqueeAnimationSpeed = 25
    }
    $dlLabel.Text = $Label
    [System.Windows.Forms.Application]::DoEvents()
}

function SetExtract {
    param([int]$Pct, [string]$Label)
    if ($script:Headless) { return }
    if ($Pct -ge 100) {
        $exBar.Style = "Continuous"
        $exBar.Value = 100
    } elseif ($Pct -lt 0) {
        if ($exBar.Style -ne "Marquee") {
            $exBar.Style                 = "Marquee"
            $exBar.MarqueeAnimationSpeed = 30
        }
    } else {
        if ($exBar.Style -ne "Continuous") { $exBar.Style = "Continuous" }
        $exBar.Value = [math]::Min($Pct, 99)
    }
    $exLabel.Text = $Label
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-DlSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $dlSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-DlSpinner {
    if ($script:Headless) { return }
    param([bool]$Success = $true)
    $dlSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $dlSpinnerLabel.ForeColor = if ($Success) { [System.Drawing.Color]::FromArgb(0, 100, 180) } else { [System.Drawing.Color]::FromArgb(200, 40, 40) }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-ExSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex  = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $exSpinnerLabel.Text  = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-ExSpinner {
    if ($script:Headless) { return }
    param([bool]$Success = $true)
    $exSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $exSpinnerLabel.ForeColor = if ($Success) { [System.Drawing.Color]::FromArgb(0, 140, 80) } else { [System.Drawing.Color]::FromArgb(200, 40, 40) }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-OverallSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex      = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $overallSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-OverallSpinner {
    if ($script:Headless) { return }
    param([bool]$Success = $true)
    $overallSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $overallSpinnerLabel.ForeColor = if ($Success) { [System.Drawing.Color]::FromArgb(80, 80, 80) } else { [System.Drawing.Color]::FromArgb(200, 40, 40) }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-AllSpinners {
    if ($script:Headless) { return }
    $script:SpinnerIndex      = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $f                        = $SpinnerFrames[$script:SpinnerIndex]
    $dlSpinnerLabel.Text      = " " + $f
    $exSpinnerLabel.Text      = " " + $f
    $overallSpinnerLabel.Text = " " + $f
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-Cancelled {
    if ($script:CancelRequested) {
        Log "Operation cancelled."
        SetDownload -Pct 0 -Label "Cancelled"
        SetExtract  -Pct 0 -Label "Cancelled"
        Stop-DlSpinner      -Success $false
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Play-Sound -Event "Cancel"
        Set-ButtonIdle
        return $true
    }
    return $false
}

function Assert-Curl {
    if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
        Log "ERROR: curl.exe not found. Windows 10 1803+ required."
        return $false
    }
    return $true
}

function Get-MissingDriverCount {
    try {
        $count = @(Get-CimInstance Win32_PnPEntity |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 }).Count
        return $count
    } catch {
        Log "  WARNING: Could not query PnP device status: $($_.Exception.Message)"
        return -1
    }
}

# =========================
# 7-ZIP HELPERS
# Installed at start of Start-Install for Dell/HP extraction.
# Removed before final cleanup so nothing persists to the customer.
# Lenovo uses Inno Setup - 7-Zip cannot extract its proprietary format.
# =========================
$script:7zExe         = "C:\Program Files\7-Zip\7z.exe"
$script:7zInstaller   = "$env:TEMP\7z-installer.exe"
$script:7zInstalled   = $false

function Install-7Zip {
    if (Test-Path $script:7zExe) {
        Log "7-Zip already present - skipping install."
        $script:7zInstalled = $true
        return $true
    }
    Log "Installing 7-Zip (temporary - will be removed after extraction)..."
    SetDownload -Pct 0 -Label "Downloading 7-Zip..."
    try {
        $psi                        = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = "curl.exe"
        $psi.Arguments              = "--silent --location --max-time 60 --connect-timeout 15 " +
                                      "--output `"$($script:7zInstaller)`" " +
                                      "`"https://www.7-zip.org/a/7z2409-x64.exe`""
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $proc                       = New-Object System.Diagnostics.Process
        $proc.StartInfo             = $psi
        $proc.Start() | Out-Null
        while (-not $proc.HasExited) { Start-Sleep -Milliseconds 400; Step-AllSpinners }
        if ($proc.ExitCode -ne 0 -or -not (Test-Path $script:7zInstaller)) {
            Log "  7-Zip download failed (curl exit $($proc.ExitCode))."
            return $false
        }
        Start-Process $script:7zInstaller -ArgumentList "/S" -Wait
        if (-not (Test-Path $script:7zExe)) {
            Log "  7-Zip installer ran but 7z.exe not found."
            return $false
        }
        $script:7zInstalled = $true
        Log "  7-Zip installed OK."
        return $true
    } catch {
        Log "  7-Zip install error: $($_.Exception.Message)"
        return $false
    }
}

function Remove-7Zip {
    Log "Removing 7-Zip..."
    try {
        $uninstaller = "C:\Program Files\7-Zip\Uninstall.exe"
        if (Test-Path $uninstaller) {
            Start-Process $uninstaller -ArgumentList "/S" -Wait
            Start-Sleep -Seconds 2
        }
        # Belt-and-braces: remove folder if uninstaller left anything
        $folder = "C:\Program Files\7-Zip"
        if (Test-Path $folder) { Remove-Item $folder -Recurse -Force -EA SilentlyContinue }
    } catch {
        Log "  WARNING: 7-Zip uninstall error: $($_.Exception.Message)"
    }
    # Always clean up the installer temp file
    Remove-Item $script:7zInstaller -Force -EA SilentlyContinue
    $script:7zInstalled = $false

    # Verify
    if (Test-Path $script:7zExe) {
        Log "  WARNING: 7z.exe still present after uninstall - check manually before sysprep."
    } else {
        Log "  7-Zip removed OK."
    }
}

# =========================
# GOOGLE SHEETS ANALYTICS
# =========================
$SHEETS_WEBHOOK = "https://script.google.com/macros/s/AKfycbygEF0i6j_6rSstmfQ2sQPLn0KjkqxZwUwIRjyCsd911IP9kALucv2cImMFumGoUUs/exec"

$script:AnalyticsManufacturer    = ""
$script:AnalyticsModel           = ""
$script:AnalyticsSerial          = ""
$script:AnalyticsOsVersion       = ""
$script:AnalyticsOsBuild         = 0
$script:AnalyticsInfCount        = 0
$script:AnalyticsDownloadMB      = 0.0
$script:AnalyticsStartTime       = $null
$script:AnalyticsMissingBefore   = -1
$script:AnalyticsMissingAfter    = -1

function Send-AnalyticsEvent {
    param(
        [ValidateSet("success","failure","cancelled")]
        [string]$Result
    )
    $durationSec = 0
    if ($script:AnalyticsStartTime) {
        $durationSec = [int]((Get-Date) - $script:AnalyticsStartTime).TotalSeconds
    }
    $payload = @"
{
  "result":         "$Result",
  "manufacturer":   "$($script:AnalyticsManufacturer -replace '"','\"')",
  "model":          "$($script:AnalyticsModel -replace '"','\"')",
  "serial":         "$($script:AnalyticsSerial -replace '"','\"')",
  "os_version":     "$($script:AnalyticsOsVersion -replace '"','\"')",
  "os_build":       $($script:AnalyticsOsBuild),
  "inf_count":      $($script:AnalyticsInfCount),
  "download_mb":    $($script:AnalyticsDownloadMB),
  "missing_before": $($script:AnalyticsMissingBefore),
  "missing_after":  $($script:AnalyticsMissingAfter),
  "duration_sec":   $durationSec,
  "script_version": "$ScriptVersion"
}
"@
    Log "Sending analytics (result=$Result, model=$($script:AnalyticsModel), infs=$($script:AnalyticsInfCount), dl=$($script:AnalyticsDownloadMB)MB, missing=$($script:AnalyticsMissingBefore)->$($script:AnalyticsMissingAfter), duration=${durationSec}s)..."
    try {
        $payloadFile = Join-Path $env:TEMP "analytics_payload_$(Get-Date -Format 'HHmmss').json"
        $utf8NoBom   = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($payloadFile, $payload, $utf8NoBom)
        $curlArgs = "--silent --max-time 15 --connect-timeout 10 " +
                    "--location " +
                    "-H `"Content-Type: application/json`" " +
                    "--data `@`"$payloadFile`" " +
                    "`"$SHEETS_WEBHOOK`""
        $psi                        = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = "curl.exe"
        $psi.Arguments              = $curlArgs
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $proc                       = New-Object System.Diagnostics.Process
        $proc.StartInfo             = $psi
        $proc.Start() | Out-Null
        $stdout = $proc.StandardOutput.ReadToEnd().Trim()
        $stderr = $proc.StandardError.ReadToEnd().Trim()
        $proc.WaitForExit()
        Remove-Item $payloadFile -EA SilentlyContinue
        if ($proc.ExitCode -ne 0) {
            Log "  Analytics warning: curl exit $($proc.ExitCode) - $stderr"
        } elseif ($stdout -eq "OK") {
            Log "  Analytics sent OK - row written to Google Sheet."
        } else {
            Log "  Analytics unexpected response: $stdout"
        }
    } catch {
        Log "  Analytics error: $($_.Exception.Message)"
    }
}

# =========================
# CURL DOWNLOAD
# =========================
function Invoke-CurlDownload {
    param(
        [string]$Url,
        [string]$OutFile
    )
    $fileName = [System.IO.Path]::GetFileName($OutFile)
    Log "Downloading: $fileName"
    Log "  URL: $Url"
    SetDownload -Pct 0 -Label "Connecting..."

    $totalBytes = 0
    try {
        $headResult = & curl.exe --silent --head --max-time 15 --connect-timeout 10 `
            --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" `
            --write-out "%{content_length_download}" --output NUL "$Url" 2>$null
        if ($headResult -match '^\d+$' -and [long]$headResult -gt 0) {
            $totalBytes = [long]$headResult
        }
    } catch {}

    $totalMB = if ($totalBytes -gt 0) { [math]::Round($totalBytes / 1MB, 1) } else { 0 }
    if ($totalMB -gt 0) { Log "  Expected size: $totalMB MB" }

    $psi                 = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "curl.exe"
    $psi.Arguments       = "--location --fail --connect-timeout 30 " +
                           "--retry 10 --retry-delay 5 --retry-all-errors " +
                           "--continue-at - " +
                           "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64)`" " +
                           "--output `"$OutFile`" `"$Url`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc                = New-Object System.Diagnostics.Process
    $proc.StartInfo      = $psi
    $proc.Start() | Out-Null

    $lastSize = 0; $stall = 0; $prevSize = 0
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 700
        $sz = if (Test-Path $OutFile) { (Get-Item $OutFile -EA SilentlyContinue).Length } else { 0 }
        if ($sz -gt $lastSize) { $stall = 0; $lastSize = $sz } else { $stall++ }
        $mbDone    = [math]::Round($sz / 1MB, 1)
        $speedMbps = [math]::Round(($sz - $prevSize) * 8 / 1MB / 0.7, 1)
        $prevSize  = $sz
        $speedStr  = if ($speedMbps -gt 0) { "  $speedMbps Mbps" } else { "" }
        if ($totalMB -gt 0) {
            $pct = [math]::Min([int](($sz / $totalBytes) * 100), 99)
            SetDownload -Pct $pct -Label "$mbDone MB / $totalMB MB  ($pct%)$speedStr"
            Log "  $mbDone MB / $totalMB MB ($pct%)$speedStr"
        } else {
            SetDownload -Pct 0 -Label "$mbDone MB received...$speedStr"
            Log "  $mbDone MB received$speedStr"
        }
        Step-AllSpinners
        if ($stall -gt 300) {
            Log "  WARNING: download stalled 3.5 min - aborting."
            $proc.Kill()
            SetDownload -Pct 0 -Label "Stalled - aborted."
            Stop-DlSpinner -Success $false
            return $false
        }
    }

    if ($proc.ExitCode -ne 0) {
        Log "  curl failed (exit $($proc.ExitCode))"
        SetDownload -Pct 0 -Label "Failed (curl exit $($proc.ExitCode))"
        Stop-DlSpinner -Success $false
        return $false
    }
    if (-not (Test-Path $OutFile) -or (Get-Item $OutFile).Length -eq 0) {
        Log "  File missing or empty after download."
        SetDownload -Pct 0 -Label "Failed - file empty."
        Stop-DlSpinner -Success $false
        return $false
    }

    $finalMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 1)
    if ($finalMB -gt $script:AnalyticsDownloadMB) { $script:AnalyticsDownloadMB = $finalMB }
    Log "  Download complete: $finalMB MB"
    SetDownload -Pct 100 -Label "Complete - $finalMB MB"
    Stop-DlSpinner -Success $true
    Play-Sound -Event "DownloadComplete"
    return $true
}

# =========================
# EXTRACTION WATCHER
# =========================
function Watch-Extraction {
    param(
        [System.Diagnostics.Process]$ExtractProc,
        [string]$DestPath,
        [int]$TotalFiles    = 0,
        [int]$StallLimitSec = 300
    )
    $stall     = 0
    $lastCount = 0
    $script:SpinnerIndex      = 0
    $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
    $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]
    SetExtract -Pct -1 -Label "Extracting..."

    while (-not $ExtractProc.HasExited) {
        Start-Sleep -Milliseconds 700
        $count = if (Test-Path $DestPath) {
            (Get-ChildItem $DestPath -Recurse -ErrorAction SilentlyContinue).Count
        } else { 0 }
        if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }
        if ($TotalFiles -gt 0) {
            $remaining = [math]::Max($TotalFiles - $count, 0)
            SetExtract -Pct -1 -Label "$count / $TotalFiles files  ($remaining remaining)"
        } else {
            SetExtract -Pct -1 -Label "$count files extracted..."
        }
        Step-ExSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "  Extraction cancelled by user."
            try { $ExtractProc.Kill() } catch {}
            break
        }
        if ($stall -gt [int]($StallLimitSec * 1.25)) {
            Log "  WARNING: extraction stalled - killing process."
            try { $ExtractProc.Kill() } catch {}
            break
        }
    }
    Start-Sleep -Seconds 2
    $finalCount = if (Test-Path $DestPath) {
        (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
    } else { 0 }
    SetExtract -Pct 100 -Label "Done - $finalCount files extracted"
    Stop-ExSpinner      -Success $true
    Stop-OverallSpinner -Success $true
    Play-Sound -Event "ExtractComplete"
    Log "  Extraction finished: $finalCount files in $DestPath"
}

# =========================
# INF INSTALLER
# =========================
function Install-DriversFromPath {
    param([string]$BasePath)
    $infs = Get-ChildItem $BasePath -Recurse -Filter *.inf -ErrorAction SilentlyContinue
    if (-not $infs -or $infs.Count -eq 0) {
        Log "No INF files found under: $BasePath"
        return $false
    }
    $total = $infs.Count; $i = 0
    Log "Found $total INF file(s)."
    if ($SkipInstall) {
        Log "  SkipInstall flag set - skipping pnputil. Extraction verified OK."
        $script:AnalyticsInfCount = 0
        SetProgress 100
        SetExtract -Pct 100 -Label "Extract complete ($total INFs found, install skipped)"
        Stop-ExSpinner      -Success $true
        Stop-OverallSpinner -Success $true
        return $true
    }
    Log "Installing via pnputil..."
    $exGroupBox.Text     = "Install INFs"
    $exSpinnerLabel.Text = " " + $SpinnerFrames[0]
    $script:SpinnerIndex = 0
    $exBar.Style         = "Continuous"
    $exBar.Value         = 0

    foreach ($inf in $infs) {
        $i++
        SetProgress (60 + [int](($i / $total) * 38))
        $infPct    = [int](($i / $total) * 100)
        $remaining = $total - $i
        SetExtract -Pct $infPct -Label "$i / $total INFs  ($remaining remaining) - $($inf.Name)"
        Log "[$i/$total] $($inf.Name)"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        foreach ($l in $out) { Log "  $l" }
        Play-Sound -Event "DriverAdded"
        Step-ExSpinner
        Step-OverallSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "INF installation cancelled at $i / $total"
            $script:AnalyticsInfCount = $i
            break
        }
    }
    $script:AnalyticsInfCount = $i
    SetProgress 100
    SetExtract -Pct 100 -Label "All $total INFs installed."
    Stop-ExSpinner      -Success $true
    Stop-OverallSpinner -Success $true
    $exGroupBox.Text = "Extract / Install"
    Log "All INFs processed."
    return $true
}

# =========================
# SHARED EXTRACT RUNNER
# =========================
function Start-PackExtraction {
    param(
        [string]$PackFile,
        [string]$DestPath,
        [int]$StallLimitSec = 300,
        [ValidateSet("Dell","HP","Lenovo","")]
        [string]$Vendor = ""
    )
    if (-not (Test-Path $DestPath)) {
        New-Item -Path $DestPath -ItemType Directory -Force | Out-Null
    }
    $ext = [System.IO.Path]::GetExtension($PackFile).ToLower()

    switch ($ext) {
        ".zip" {
            Log "Extracting ZIP..."
            SetExtract -Pct -1 -Label "Starting ZIP extraction..."
            $zipJob = Start-Job {
                param($src, $dst)
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::ExtractToDirectory($src, $dst)
            } -ArgumentList $PackFile, $DestPath

            $stall = 0; $lastCount = 0
            $script:SpinnerIndex      = 0
            $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
            $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]
            while ($zipJob.State -eq "Running") {
                Start-Sleep -Milliseconds 700
                $count = if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
                if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }
                SetExtract -Pct -1 -Label "$count files extracted..."
                Step-ExSpinner
                Step-OverallSpinner
                [System.Windows.Forms.Application]::DoEvents()
                if ($stall -gt 375) { Log "  ZIP stalled - stopping."; Stop-Job $zipJob; break }
            }
            Receive-Job $zipJob -EA SilentlyContinue | Out-Null
            Remove-Job  $zipJob
            $finalCount = if (Test-Path $DestPath) {
                (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
            } else { 0 }
            SetExtract -Pct 100 -Label "Done - $finalCount files extracted"
            Stop-ExSpinner      -Success $true
            Stop-OverallSpinner -Success $true
            Play-Sound -Event "ExtractComplete"
            Log "  ZIP extraction complete. $finalCount files."
        }

        ".cab" {
            Log "Extracting CAB..."
            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = "expand.exe"
            $psi.Arguments       = "`"$PackFile`" -F:* `"$DestPath`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            $exProc              = New-Object System.Diagnostics.Process
            $exProc.StartInfo    = $psi
            $exProc.Start() | Out-Null
            Watch-Extraction -ExtractProc $exProc -DestPath $DestPath -StallLimitSec $StallLimitSec
        }

        default {
            # ------------------------------------------------------------------
            # Dell and HP: use 7-Zip pass-1 when available.
            #   Verified against real packs - byte-for-byte identical output,
            #   faster than vendor extractors (Dell 1.1x, HP 4.7x).
            #   Falls back to vendor extractor if 7-Zip unavailable or yields 0 files.
            #
            # Lenovo: always uses Inno Setup /VERYSILENT /EXTRACT=YES.
            #   7-Zip cannot extract Lenovo's proprietary Inno payload format.
            # ------------------------------------------------------------------

            $CountFiles = {
                if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
            }

            # Try 7-Zip pass-1 for Dell and HP
            if (($Vendor -eq "Dell" -or $Vendor -eq "HP") -and (Test-Path $script:7zExe)) {
                Log "  Extracting with 7-Zip (pass 1)..."
                SetExtract -Pct -1 -Label "Extracting with 7-Zip..."
                $script:SpinnerIndex      = 0
                $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
                $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]

                $psi                 = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName        = $script:7zExe
                $psi.Arguments       = "x `"$PackFile`" -o`"$DestPath`" -y"
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow  = $true
                $sevenProc           = New-Object System.Diagnostics.Process
                $sevenProc.StartInfo = $psi
                $sevenProc.Start() | Out-Null

                while (-not $sevenProc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    Step-OverallSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:CancelRequested) { try { $sevenProc.Kill() } catch {}; break }
                }
                Start-Sleep -Seconds 1
                $n7z = & $CountFiles

                if ($n7z -gt 0) {
                    Log "  7-Zip extraction complete: $n7z files."
                    SetExtract -Pct 100 -Label "Done - $n7z files extracted"
                    Stop-ExSpinner      -Success $true
                    Stop-OverallSpinner -Success $true
                    Play-Sound -Event "ExtractComplete"
                    return  # success - skip fallback chain
                }
                Log "  7-Zip yielded 0 files - falling back to vendor extractor."
            }

            # Vendor extractor fallback (always used for Lenovo, fallback for Dell/HP)
            Log "Extracting EXE pack (Vendor=$Vendor)..."

            function TrySync {
                param([string]$ExeArgs)
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = $ExeArgs
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                while (-not $proc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    Step-OverallSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:CancelRequested) { try { $proc.Kill() } catch {}; break }
                }
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            function TryAsyncHP {
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = "/s /e /f `"$DestPath`""
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                $start = Get-Date
                while (-not $proc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if (((Get-Date) - $start).TotalSeconds -gt 30 -and $n -eq 0) {
                        Log "  HP format timed out with no output - killing."
                        try { $proc.Kill() } catch {}
                        break
                    }
                }
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            $attempts = [System.Collections.Generic.List[object]]::new()
            switch ($Vendor) {
                "Dell"   {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
                "Lenovo" {
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
                "HP"     {
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                }
                default  {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
            }

            $extracted = $false
            foreach ($attempt in $attempts) {
                Log "  Trying: $($attempt.Label)..."
                $n = & $attempt.Action
                Log "  Result: $n files in $DestPath"
                if ($n -gt 0) {
                    SetExtract -Pct 100 -Label "Done - $n files extracted"
                    Stop-ExSpinner -Success $true; Stop-OverallSpinner -Success $true
                    Play-Sound -Event "ExtractComplete"
                    Log "  Extraction finished: $n files in $DestPath"
                    $extracted = $true
                    break
                }
            }

            if (-not $extracted) {
                Log "  All primary formats failed - trying Lenovo legacy (-s -fdest)..."
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = "-s -f`"$DestPath`""
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                Watch-Extraction -ExtractProc $proc -DestPath $DestPath -StallLimitSec $StallLimitSec
            }
        }
    }
}

# =========================
# DELL
# =========================
function Start-DellDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== DELL: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $serviceTag = $null
    try {
        $serviceTag = (Get-CimInstance Win32_BIOS).SerialNumber.Trim()
        Log "Service Tag: $serviceTag"
        $script:AnalyticsSerial = $serviceTag
    } catch {
        Log "Could not read Service Tag: $($_.Exception.Message)"
        return $false
    }
    if (-not $serviceTag -or $serviceTag.Length -lt 4) { Log "Invalid Service Tag."; return $false }

    $isWin11 = $false
    try {
        $isWin11 = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber -ge 22000
        Log "OS: $(if ($isWin11) {'Windows 11'} else {'Windows 10'})"
    } catch {}

    Log "Downloading Dell DriverPackCatalog.cab..."
    $catalogCab = Join-Path $env:TEMP "DellDriverPackCatalog.cab"
    $catalogXml = Join-Path $env:TEMP "DriverPackCatalog.xml"
    Remove-Item $catalogCab -EA SilentlyContinue
    Remove-Item $catalogXml -EA SilentlyContinue

    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/DriverPackCatalog.cab" -OutFile $catalogCab)) {
        Log "Failed to download Dell DriverPackCatalog."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 15

    Log "Extracting Dell DriverPackCatalog CAB..."
    SetExtract -Pct 10 -Label "Extracting catalog..."
    $expandOut = & expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1
    Log "  expand.exe: $expandOut"
    if (-not (Test-Path $catalogXml) -or (Get-Item $catalogXml).Length -eq 0) {
        Log "DriverPackCatalog CAB extraction failed."
        return $false
    }
    SetExtract -Pct 30 -Label "Catalog extracted OK"
    SetProgress 20

    Log "Parsing DriverPackCatalog..."
    try {
        $rawXml   = [System.IO.File]::ReadAllText($catalogXml).TrimStart([char]0xFEFF)
        [xml]$cat = $rawXml
    } catch {
        Log "Failed to parse DriverPackCatalog XML: $($_.Exception.Message)"
        return $false
    }

    $searchNames = @()
    if ($ModelName) {
        $searchNames += $ModelName
        $searchNames += ($ModelName -replace '\s+Notebook.*$','')
        $searchNames += ($ModelName -replace '^Dell\s+','')
        $searchNames = $searchNames | Select-Object -Unique | Where-Object { $_.Length -gt 3 }
    }
    Log "Searching catalog - model tokens: $($searchNames -join ' | ')"

    $candidates = @()
    foreach ($pkg in $cat.SelectNodes("//*[local-name()='DriverPackage']")) {
        $modelMatched = $false
        foreach ($modelNode in $pkg.SelectNodes(".//*[local-name()='Model']")) {
            $nameAttr = $modelNode.GetAttribute("name")
            foreach ($tok in $searchNames) {
                if ($nameAttr -ieq $tok) { $modelMatched = $true; break }
            }
            if ($modelMatched) { break }
        }
        if (-not $modelMatched) { continue }

        $supportsWin11 = $false; $supportsWin10 = $false
        foreach ($osNode in $pkg.SelectNodes(".//*[local-name()='OperatingSystem']")) {
            $osDisp = ""
            try { $osDisp = $osNode.SelectSingleNode("*[local-name()='Display']").InnerText } catch {}
            if ($osDisp -match "(?i)windows 11") { $supportsWin11 = $true }
            if ($osDisp -match "(?i)windows 10") { $supportsWin10 = $true }
        }
        $pkgName = ""
        try { $pkgName = $pkg.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch {}
        $pkgPath = $pkg.GetAttribute("path")
        $candidates += [PSCustomObject]@{ Path = $pkgPath; DisplayName = $pkgName; Win11 = $supportsWin11; Win10 = $supportsWin10 }
        Log "  Candidate: $pkgName  [W11=$supportsWin11 W10=$supportsWin10]"
    }

    if ($candidates.Count -eq 0) {
        Log "No driver pack found in DriverPackCatalog for '$ModelName'."
        Log "Opening Dell support page..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    $chosen = $null
    if ($isWin11) {
        $chosen = $candidates | Where-Object { $_.Win11 } | Select-Object -First 1
        if (-not $chosen) { $chosen = $candidates[0] }
    } else {
        $chosen = $candidates | Where-Object { $_.Win10 } | Select-Object -First 1
        if (-not $chosen) { $chosen = $candidates[0] }
    }

    Log "Selected: $($chosen.DisplayName)"
    $packPath = $chosen.Path
    if (-not $packPath) { Log "Driver pack entry has no path - unexpected catalog format."; return $false }

    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Pack file: $([System.IO.Path]::GetFileName($packPath))"
    Log "Pack URL:  $packUrl"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) { Log "Dell driver pack download failed."; return $false }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Dell_Extracted"
    Log "Extracting Dell pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "Dell"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# HP
# =========================
function Start-HpDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== HP: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    $osBuild = $null; $isWin11 = $false
    try {
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
        $isWin11 = $osBuild -ge 22000
        Log "OS Build: $osBuild  ($(if ($isWin11) {'Win11'} else {'Win10'}))"
    } catch { Log "Could not read OS build: $($_.Exception.Message)" }

    $matrixUrl  = "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html"
    $matrixFile = Join-Path $env:TEMP "HP_DPMatrix.html"
    Remove-Item $matrixFile -EA SilentlyContinue

    Log "Downloading HP Driver Pack Matrix..."
    SetExtract -Pct 5 -Label "Downloading matrix page..."
    if (-not (Invoke-CurlDownload -Url $matrixUrl -OutFile $matrixFile)) {
        Log "Failed to download HP Driver Pack Matrix."
        Start-Process $matrixUrl
        return $false
    }
    if (Test-Cancelled) { return $false }

    $matrixHtml = [System.IO.File]::ReadAllText($matrixFile)
    Log "  Matrix HTML: $([math]::Round($matrixHtml.Length/1KB)) KB"
    SetExtract -Pct 20 -Label "Parsing matrix..."
    SetProgress 20

    $searchTokens = @()
    if ($ModelName) {
        $stripped = $ModelName -replace '(?i)^HP\s+', ''
        $searchTokens += $stripped
        $searchTokens += ($stripped -replace '\s+Notebook.*$', '')
        $searchTokens += ($stripped -replace '\s+PC.*$',       '')
        $searchTokens = $searchTokens | Select-Object -Unique | Where-Object { $_.Length -gt 4 }
    }
    Log "Search tokens: $($searchTokens -join ' | ')"

    $packUrl = $null; $packSpNum = $null
    $flat    = $matrixHtml -replace "`r`n|`r|`n", " " -replace "\s{2,}", " "
    $rows    = [regex]::Matches($flat, '(?i)<tr[^>]*>(.*?)</tr>')

    foreach ($row in $rows) {
        $rowHtml = $row.Groups[1].Value
        $cells   = [regex]::Matches($rowHtml, '(?i)<t[dh][^>]*>(.*?)</t[dh]>')
        if ($cells.Count -lt 2) { continue }
        $modelCell = [regex]::Replace($cells[0].Groups[1].Value, '<[^>]+>', ' ')
        $modelCell = [System.Net.WebUtility]::HtmlDecode($modelCell) -replace '\s+', ' '
        $matched = $false
        foreach ($tok in $searchTokens) {
            if ($modelCell -match [regex]::Escape($tok)) { $matched = $true; break }
        }
        if (-not $matched) { continue }
        Log "  Matched matrix row: $($modelCell.Trim() -replace '\s+',' ')"
        $allLinks = [regex]::Matches($rowHtml, '(?i)href="([^"]*sp\d+\.exe)"')
        if ($allLinks.Count -eq 0) { Log "  Row matched but contains no .exe links - skipping."; continue }
        $bestUrl = $allLinks[0].Groups[1].Value
        if (-not $bestUrl.StartsWith("http")) { $bestUrl = "https://ftp.hp.com$bestUrl" }
        if (-not $isWin11) {
            foreach ($lm in $allLinks) {
                $href = $lm.Groups[1].Value
                if (-not $href.StartsWith("http")) { $href = "https://ftp.hp.com$href" }
                $aTag  = [regex]::Match($rowHtml, "(?i)<a[^>]+href=""[^""]*$([regex]::Escape([System.IO.Path]::GetFileName($href)))[^""]*""[^>]*>")
                $title = if ($aTag.Success) { $aTag.Value } else { "" }
                if ($title -eq "" -or $title -match "(?i)windows 10") { $bestUrl = $href; break }
            }
        }
        $packUrl   = $bestUrl
        $packSpNum = [regex]::Match($packUrl, '(?i)(sp\d+)\.exe').Groups[1].Value
        Log "  Selected SoftPaq: $packSpNum"
        Log "  URL: $packUrl"
        break
    }

    if (Test-Cancelled) { return $false }
    if (-not $packUrl) {
        Log "Model '$ModelName' not found in HP Driver Pack Matrix."
        Log "Opening HP Driver Pack Matrix for manual selection..."
        Start-Process $matrixUrl
        return $false
    }

    SetExtract  -Pct 40 -Label "Matrix OK - $packSpNum"
    SetProgress 25
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot "$packSpNum.exe"
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) { Log "HP driver pack download failed."; return $false }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "HP_Extracted"
    Log "Extracting HP SoftPaq..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "HP"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# LENOVO
# =========================
function Start-LenovoDriverInstall {
    param([string]$DriverRoot)

    Log "=== LENOVO: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $machineType = $null
    if ($MachineType) {
        # Use override from -MachineType param - take first 4 chars uppercased
        $machineType = $MachineType.Substring(0, [math]::Min(4, $MachineType.Length)).ToUpper()
        Log "Machine type: $MachineType  ->  prefix: $machineType  [overridden via param]"
    } else {
        try {
            $sku = (Get-CimInstance Win32_ComputerSystemProduct).Name.Trim()
            if ($sku.Length -ge 4) {
                $machineType = $sku.Substring(0, 4).ToUpper()
                Log "Machine type: $sku  ->  prefix: $machineType"
            }
        } catch { Log "Could not read machine type: $($_.Exception.Message)" }
    }
    if (-not $machineType) { Log "Cannot determine Lenovo machine type."; return $false }

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    $winVer     = (Get-CimInstance Win32_OperatingSystem).Version
    $osAttr     = if ($winVer -match "^10\.0\.2") { "win11" } else { "win10" }
    $osFallback = if ($osAttr -eq "win11") { "win10" } else { "win11" }
    Log "Detected OS tag: $osAttr"

    Log "Fetching Lenovo catalogv2.xml..."
    $catalogFile = Join-Path $env:TEMP "lenovo_catalogv2.xml"
    if (-not (Invoke-CurlDownload -Url "https://download.lenovo.com/cdrt/td/catalogv2.xml" -OutFile $catalogFile)) {
        Log "Failed to download Lenovo catalog."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 20

    Log "Parsing Lenovo catalog..."
    SetExtract -Pct 10 -Label "Parsing catalog..."
    try {
        $bytes    = [System.IO.File]::ReadAllBytes($catalogFile)
        $rawText  = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        [xml]$cat = $rawText
        Log "Catalog parsed OK."
    } catch { Log "Failed to parse Lenovo catalog: $($_.Exception.Message)"; return $false }
    SetExtract -Pct 30 -Label "Catalog parsed"
    SetProgress 25

    $packUrl = $null
    foreach ($model in $cat.ModelList.Model) {
        $types = @($model.Types.Type)
        if (-not ($types | Where-Object { $_ -like "$machineType*" })) { continue }
        Log "Matched model: $($model.name)"
        foreach ($os in @($osAttr, $osFallback)) {
            $nodes = @($model.SCCM | Where-Object { $_.os -eq $os })
            if ($nodes.Count -gt 0) {
                $url = ($nodes | Select-Object -Last 1)."#text"
                if ($url -match "^https?://") { $packUrl = $url; Log "Driver pack URL [$os]: $packUrl"; break }
            }
        }
        break
    }

    if (-not $packUrl) {
        Log "No SCCM driver pack found for machine type '$machineType'."
        Start-Process "https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html"
        return $false
    }
    SetProgress 28
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) { Log "Lenovo driver pack download failed."; return $false }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Lenovo_Extracted"
    Log "Extracting Lenovo pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "Lenovo"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# MICROSOFT (SURFACE)
#
# No machine-readable catalog exists - uses a hardcoded lookup table
# mapping WMI model names to Microsoft Download Center page IDs.
#
# Flow:
#   1. Match WMI model to a Download Center page ID.
#   2. Fetch the details page and scrape the direct .msi href from HTML.
#   3. When multiple MSIs exist for different OS builds, pick the one
#      whose embedded build number is <= the device OS build (per
#      Microsoft's own documented guidance).
#   4. Download the MSI, then run: msiexec /i <file> /quiet /norestart
#      No extraction step - the MSI installs drivers directly.
#   5. Analytics InfCount stays 0 (MSI-based install, no pnputil).
#
# To add new Surface models: add an entry to $SurfaceDownloadIds below.
# Source: https://support.microsoft.com/en-us/surface/drivers-firmware/
# =========================

$SurfaceDownloadIds = [ordered]@{
    # Surface Pro
    "Surface Pro 12"                          = "108199"
    "Surface Pro for Business (11th Edition)" = "108013"
    "Surface Pro (11th Edition)"              = "106119"
    "Surface Pro 10 with 5G"                  = "106292"
    "Surface Pro 10"                          = "105947"
    "Surface Pro 9 with 5G"                   = "105941"
    "Surface Pro 9"                           = "104680"
    "Surface Pro 8"                           = "103503"
    "Surface Pro 7+"                          = "102633"
    "Surface Pro 7"                           = "100419"
    "Surface Pro 6"                           = "57514"
    "Surface Pro with LTE"                    = "56278"
    "Surface Pro (5th Gen)"                   = "55484"
    "Surface Pro 5"                           = "55484"
    "Surface Pro 4"                           = "49498"
    "Surface Pro 3"                           = "38826"
    "Surface Pro 2"                           = "49042"
    # Surface Laptop
    "Surface Laptop 7"                        = "106123"
    "Surface Laptop 6"                        = "105950"
    "Surface Laptop 5"                        = "104220"
    "Surface Laptop 4"                        = "102924"
    "Surface Laptop 3"                        = "100429"
    "Surface Laptop 2"                        = "57515"
    "Surface Laptop Studio 2"                 = "105386"
    "Surface Laptop Studio"                   = "103505"
    "Surface Laptop Go 3"                     = "105941"
    "Surface Laptop Go 2"                     = "103739"
    "Surface Laptop Go"                       = "101304"
    # Surface Book
    "Surface Book 3"                          = "101315"
    "Surface Book 2"                          = "56261"
    "Surface Book"                            = "49497"
    # Surface Go
    "Surface Go 4"                            = "105386"
    "Surface Go 3"                            = "103504"
    "Surface Go 2"                            = "101304"
    "Surface Go"                              = "100145"
    # Surface Studio
    "Surface Studio 2+"                       = "104679"
    "Surface Studio 2"                        = "57593"
    "Surface Studio"                          = "54311"
}

function Start-MicrosoftSurfaceDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== MICROSOFT SURFACE: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    $osBuild = 0
    try {
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
        Log "OS Build: $osBuild"
    } catch { Log "Could not read OS build: $($_.Exception.Message)" }

    $pageId = $null
    foreach ($entry in $SurfaceDownloadIds.GetEnumerator()) {
        if ($ModelName -match [regex]::Escape($entry.Key)) {
            $pageId = $entry.Value
            Log "Matched model '$($entry.Key)' -> Download Center ID: $pageId"
            break
        }
    }
    if (-not $pageId) {
        foreach ($entry in $SurfaceDownloadIds.GetEnumerator()) {
            if ($ModelName -ilike "*$($entry.Key)*") {
                $pageId = $entry.Value
                Log "Fuzzy-matched '$($entry.Key)' -> Download Center ID: $pageId"
                break
            }
        }
    }

    if (-not $pageId) {
        Log "No Download Center entry found for: '$ModelName'"
        Log "Note: Surface Pro X uses Windows Update only (ARM - no MSI available)."
        Log "Opening Surface driver downloads page for manual selection..."
        Start-Process "https://support.microsoft.com/en-us/surface/drivers-firmware/download-drivers-and-firmware-for-surface-pro"
        return $false
    }

    $detailsUrl  = "https://www.microsoft.com/en-us/download/details.aspx?id=$pageId"
    $detailsFile = Join-Path $env:TEMP "surface_dl_page_$pageId.html"
    Remove-Item $detailsFile -EA SilentlyContinue

    Log "Fetching download details page (ID=$pageId)..."
    SetExtract -Pct 5 -Label "Fetching download page..."
    SetProgress 15

    $psi                        = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "curl.exe"
    $psi.Arguments              = "--silent --location --max-time 30 --connect-timeout 15 " +
                                  "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36`" " +
                                  "--output `"$detailsFile`" `"$detailsUrl`""
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardError  = $true
    $fetchProc                  = New-Object System.Diagnostics.Process
    $fetchProc.StartInfo        = $psi
    $fetchProc.Start() | Out-Null

    $start = Get-Date
    while (-not $fetchProc.HasExited) {
        Start-Sleep -Milliseconds 400
        Step-AllSpinners
        if (((Get-Date) - $start).TotalSeconds -gt 35) {
            Log "  Timeout fetching download page."
            try { $fetchProc.Kill() } catch {}
            break
        }
    }
    $fetchProc.WaitForExit()

    if (-not (Test-Path $detailsFile) -or (Get-Item $detailsFile).Length -lt 1000) {
        Log "Failed to fetch download details page for ID=$pageId"
        Log "Opening page manually: $detailsUrl"
        Start-Process $detailsUrl
        return $false
    }
    if (Test-Cancelled) { return $false }

    Log "Parsing download page for MSI links..."
    SetExtract -Pct 20 -Label "Parsing download page..."
    $pageHtml = [System.IO.File]::ReadAllText($detailsFile)

    $msiMatches = [regex]::Matches($pageHtml, 'href="(https://download\.microsoft\.com/[^"]+\.msi)"')
    if ($msiMatches.Count -eq 0) {
        $msiMatches = [regex]::Matches($pageHtml, '(https://download\.microsoft\.com/[^\s"<>]+\.msi)')
    }

    if ($msiMatches.Count -eq 0) {
        Log "No MSI links found on page ID=$pageId - page may require JavaScript."
        Log "Opening page for manual download: $detailsUrl"
        Start-Process $detailsUrl
        return $false
    }
    Log "  Found $($msiMatches.Count) MSI link(s)."

    $msiCandidates = [System.Collections.Generic.List[object]]::new()
    $seen          = @{}
    foreach ($m in $msiMatches) {
        $url = $m.Groups[1].Value
        if ($seen.ContainsKey($url)) { continue }
        $seen[$url] = $true
        $fileName   = [System.IO.Path]::GetFileName(([System.Uri]$url).LocalPath)
        $buildMatch = [regex]::Match($fileName, '_Win\d+_(\d{5})_')
        $msiOsBuild = if ($buildMatch.Success) { [int]$buildMatch.Groups[1].Value } else { 0 }
        Log "  MSI: $fileName  (target build: $(if ($msiOsBuild -gt 0) { $msiOsBuild } else { 'unknown' }))"
        $msiCandidates.Add([PSCustomObject]@{ Url = $url; FileName = $fileName; OsBuild = $msiOsBuild })
    }

    $chosen = $null
    if ($osBuild -gt 0) {
        $eligible = $msiCandidates |
            Where-Object { $_.OsBuild -gt 0 -and $_.OsBuild -le $osBuild } |
            Sort-Object OsBuild -Descending
        if ($eligible) {
            $chosen = $eligible[0]
            Log "Selected MSI (best build match): $($chosen.FileName)  [target=$($chosen.OsBuild) <= device=$osBuild]"
        } else {
            $chosen = $msiCandidates | Where-Object { $_.OsBuild -gt 0 } | Sort-Object OsBuild | Select-Object -First 1
            if (-not $chosen) { $chosen = $msiCandidates[0] }
            Log "No MSI at/below build $osBuild - using lowest available: $($chosen.FileName)"
        }
    } else {
        $chosen = $msiCandidates[0]
        Log "OS build unknown - using first MSI on page: $($chosen.FileName)"
    }

    Log "MSI URL: $($chosen.Url)"
    SetExtract -Pct 40 -Label "MSI selected: $($chosen.FileName)"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $msiFile = Join-Path $DriverRoot $chosen.FileName
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $chosen.Url -OutFile $msiFile)) {
        Log "Surface MSI download failed."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    Log "Installing Surface MSI: $($chosen.FileName)"
    Log "  msiexec /i `"$msiFile`" /quiet /norestart"
    $exGroupBox.Text             = "Install MSI"
    $exSpinnerLabel.Text         = " " + $SpinnerFrames[0]
    $script:SpinnerIndex         = 0
    $exBar.Style                 = "Marquee"
    $exBar.MarqueeAnimationSpeed = 30
    SetExtract -Pct -1 -Label "Installing MSI - this may take several minutes..."

    $psi                 = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "msiexec.exe"
    $psi.Arguments       = "/i `"$msiFile`" /quiet /norestart /l*v `"$DriverRoot\msiexec_install.log`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $msiProc             = New-Object System.Diagnostics.Process
    $msiProc.StartInfo   = $psi
    $msiProc.Start() | Out-Null

    $elapsed = 0
    while (-not $msiProc.HasExited) {
        Start-Sleep -Milliseconds 700
        $elapsed += 0.7
        $mins = [int]($elapsed / 60)
        $secs = [int]($elapsed % 60)
        SetExtract -Pct -1 -Label "Installing MSI... ($mins`m $secs`s elapsed)"
        Step-ExSpinner
        Step-OverallSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "  MSI install cancelled by user."
            try { $msiProc.Kill() } catch {}
            break
        }
        if ($elapsed -gt 1800) {
            Log "  WARNING: MSI install exceeded 30 minutes - aborting."
            try { $msiProc.Kill() } catch {}
            break
        }
    }

    $exitCode   = $msiProc.ExitCode
    $msiSuccess = ($exitCode -eq 0 -or $exitCode -eq 3010)
    Log "  msiexec exit code: $exitCode"

    if ($msiSuccess) {
        $rebootNeeded = ($exitCode -eq 3010)
        SetExtract -Pct 100 -Label "MSI installed $(if ($rebootNeeded) {'- reboot required'} else {'successfully'})"
        Stop-ExSpinner      -Success $true
        Stop-OverallSpinner -Success $true
        Play-Sound -Event "ExtractComplete"
        Log "  Surface MSI installation complete.$(if ($rebootNeeded) {' Reboot required.'})"
        $script:AnalyticsInfCount = 0
        SetProgress 100
        $exGroupBox.Text = "Install MSI"
        return $true
    } else {
        SetExtract -Pct 0 -Label "MSI install failed (exit $exitCode)"
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Play-Sound -Event "Failure"
        Log "  MSI installation failed. See: $DriverRoot\msiexec_install.log"
        return $false
    }
}

# =========================
# DEVICE INFO DUMP
# =========================
function Write-DeviceInfo {
    Log "============================================"
    Log "  DEVICE INFORMATION DUMP"
    Log "============================================"

    Log "-- System --"
    try {
        $cs = Get-CimInstance Win32_ComputerSystem
        Log "  Manufacturer       : $($cs.Manufacturer)"
        Log "  Model              : $($cs.Model)"
        Log "  SystemSKUNumber    : $($cs.SystemSKUNumber)"
        Log "  SystemFamily       : $($cs.SystemFamily)"
        Log "  PCSystemType       : $($cs.PCSystemType)"
        Log "  TotalPhysRAM (GB)  : $([math]::Round($cs.TotalPhysicalMemory/1GB,2))"
        Log "  Domain             : $($cs.Domain)"
        Log "  UserName           : $($cs.UserName)"
    } catch { Log "  [Win32_ComputerSystem ERROR] $($_.Exception.Message)" }

    try {
        $csp = Get-CimInstance Win32_ComputerSystemProduct
        Log "  CSProduct.Name     : $($csp.Name)"
        Log "  CSProduct.Version  : $($csp.Version)"
        Log "  CSProduct.UUID     : $($csp.UUID)"
        Log "  CSProduct.Vendor   : $($csp.Vendor)"
    } catch { Log "  [Win32_ComputerSystemProduct ERROR] $($_.Exception.Message)" }

    Log "-- BIOS --"
    try {
        $bios = Get-CimInstance Win32_BIOS
        Log "  SerialNumber       : $($bios.SerialNumber)"
        Log "  SMBIOSBIOSVersion  : $($bios.SMBIOSBIOSVersion)"
        Log "  ReleaseDate        : $($bios.ReleaseDate)"
        Log "  Manufacturer       : $($bios.Manufacturer)"
        Log "  Name               : $($bios.Name)"
        Log "  Version            : $($bios.Version)"
    } catch { Log "  [Win32_BIOS ERROR] $($_.Exception.Message)" }

    Log "-- Baseboard --"
    try {
        $bb = Get-CimInstance Win32_BaseBoard
        Log "  Product            : $($bb.Product)"
        Log "  Manufacturer       : $($bb.Manufacturer)"
        Log "  SerialNumber       : $($bb.SerialNumber)"
        Log "  Version            : $($bb.Version)"
    } catch { Log "  [Win32_BaseBoard ERROR] $($_.Exception.Message)" }

    Log "-- Enclosure --"
    try {
        $enc = Get-CimInstance Win32_SystemEnclosure
        Log "  ChassisTypes       : $($enc.ChassisTypes -join ',')"
        Log "  SMBIOSAssetTag     : $($enc.SMBIOSAssetTag)"
        Log "  SerialNumber       : $($enc.SerialNumber)"
        Log "  Manufacturer       : $($enc.Manufacturer)"
    } catch { Log "  [Win32_SystemEnclosure ERROR] $($_.Exception.Message)" }

    Log "-- Operating System --"
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        Log "  Caption            : $($os.Caption)"
        Log "  Version            : $($os.Version)"
        Log "  BuildNumber        : $($os.BuildNumber)"
        Log "  OSArchitecture     : $($os.OSArchitecture)"
        Log "  SystemDrive        : $($os.SystemDrive)"
        Log "  WindowsDirectory   : $($os.WindowsDirectory)"
        Log "  InstallDate        : $($os.InstallDate)"
        Log "  LastBootUpTime     : $($os.LastBootUpTime)"
    } catch { Log "  [Win32_OperatingSystem ERROR] $($_.Exception.Message)" }

    Log "-- Processor --"
    try {
        $cpus = Get-CimInstance Win32_Processor
        foreach ($cpu in $cpus) {
            Log "  Name               : $($cpu.Name)"
            Log "  DeviceID           : $($cpu.DeviceID)"
            Log "  Manufacturer       : $($cpu.Manufacturer)"
            Log "  MaxClockSpeed      : $($cpu.MaxClockSpeed) MHz"
            Log "  NumberOfCores      : $($cpu.NumberOfCores)"
            Log "  NumberOfLogical    : $($cpu.NumberOfLogicalProcessors)"
            Log "  ProcessorId        : $($cpu.ProcessorId)"
        }
    } catch { Log "  [Win32_Processor ERROR] $($_.Exception.Message)" }

    Log "-- Video Controller --"
    try {
        $gpus = Get-CimInstance Win32_VideoController
        foreach ($gpu in $gpus) {
            Log "  Name               : $($gpu.Name)"
            Log "  DeviceID           : $($gpu.DeviceID)"
            Log "  AdapterRAM         : $([math]::Round($gpu.AdapterRAM/1MB))MB"
            Log "  DriverVersion      : $($gpu.DriverVersion)"
            Log "  DriverDate         : $($gpu.DriverDate)"
            Log "  VideoModeDesc      : $($gpu.VideoModeDescription)"
            Log "  ---"
        }
    } catch { Log "  [Win32_VideoController ERROR] $($_.Exception.Message)" }

    Log "-- Network Adapters --"
    try {
        $nics = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
        foreach ($nic in $nics) {
            Log "  Name               : $($nic.Name)"
            Log "  MACAddress         : $($nic.MACAddress)"
            Log "  AdapterType        : $($nic.AdapterType)"
            Log "  ---"
        }
    } catch { Log "  [Win32_NetworkAdapter ERROR] $($_.Exception.Message)" }

    Log "-- Disk Drives --"
    try {
        $disks = Get-CimInstance Win32_DiskDrive
        foreach ($d in $disks) {
            $sizeGB = if ($d.Size) { [math]::Round($d.Size/1GB,1) } else { "?" }
            Log "  Model              : $($d.Model)"
            Log "  SerialNumber       : $($d.SerialNumber)"
            Log "  InterfaceType      : $($d.InterfaceType)"
            Log "  Size               : $sizeGB GB"
            Log "  MediaType          : $($d.MediaType)"
            Log "  ---"
        }
    } catch { Log "  [Win32_DiskDrive ERROR] $($_.Exception.Message)" }

    Log "-- Physical Memory --"
    try {
        $dimms = Get-CimInstance Win32_PhysicalMemory
        foreach ($d in $dimms) {
            $sz = if ($d.Capacity) { [math]::Round($d.Capacity/1GB,1) } else { "?" }
            Log "  BankLabel          : $($d.BankLabel)"
            Log "  DeviceLocator      : $($d.DeviceLocator)"
            Log "  Capacity           : $sz GB"
            Log "  Speed              : $($d.Speed) MHz"
            Log "  Manufacturer       : $($d.Manufacturer)"
            Log "  PartNumber         : $($d.PartNumber)"
            Log "  ---"
        }
    } catch { Log "  [Win32_PhysicalMemory ERROR] $($_.Exception.Message)" }

    Log "-- Battery --"
    try {
        $batts = Get-CimInstance Win32_Battery
        if ($batts) {
            foreach ($b in $batts) {
                Log "  Name               : $($b.Name)"
                Log "  EstimatedRuntime   : $($b.EstimatedRunTime) min"
                Log "  BatteryStatus      : $($b.BatteryStatus)"
                Log "  DesignCapacity     : $($b.DesignCapacity) mWh"
                Log "  FullChargeCapacity : $($b.FullChargeCapacity) mWh"
            }
        } else { Log "  (no battery detected - desktop?)" }
    } catch { Log "  [Win32_Battery ERROR] $($_.Exception.Message)" }

    Log "-- PnP Devices (problem state / no driver) --"
    try {
        $problem = Get-CimInstance Win32_PnPEntity |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
            Select-Object -Property Name, DeviceID, ConfigManagerErrorCode
        if ($problem) {
            foreach ($p in $problem) {
                Log "  [ERR $($p.ConfigManagerErrorCode)] $($p.Name)"
                Log "    DeviceID: $($p.DeviceID)"
            }
        } else { Log "  (none - all PnP devices have drivers)" }
    } catch { Log "  [Win32_PnPEntity ERROR] $($_.Exception.Message)" }

    Log "-- pnputil driver store (first 20 OEM INFs) --"
    try {
        $pnpOut = & pnputil /enum-drivers 2>&1 | Select-Object -First 60
        foreach ($line in $pnpOut) { Log "  $line" }
    } catch { Log "  [pnputil ERROR] $($_.Exception.Message)" }

    Log "-- Environment --"
    Log "  TEMP               : $env:TEMP"
    Log "  COMPUTERNAME       : $env:COMPUTERNAME"
    Log "  USERNAME           : $env:USERNAME"
    Log "  PROCESSOR_ARCH     : $env:PROCESSOR_ARCHITECTURE"
    Log "  PS Version         : $($PSVersionTable.PSVersion)"
    Log "  curl.exe path      : $($(Get-Command curl.exe -EA SilentlyContinue).Source)"

    Log "-- Disk Space --"
    try {
        $c = Get-PSDrive C -EA Stop
        Log "  C: Used (GB)       : $([math]::Round($c.Used/1GB,2))"
        Log "  C: Free (GB)       : $([math]::Round($c.Free/1GB,2))"
    } catch { Log "  [PSDrive C: ERROR] $($_.Exception.Message)" }

    Log "============================================"
    Log "  END DEVICE INFORMATION DUMP"
    Log "============================================"
}

# =========================
# MAIN
# =========================
function Start-Install {

    $script:CancelRequested          = $false
    $script:AnalyticsInfCount        = 0
    $script:AnalyticsDownloadMB      = 0.0
    $script:AnalyticsSerial          = ""
    $script:AnalyticsManufacturer    = ""
    $script:AnalyticsModel           = ""
    $script:AnalyticsOsVersion       = ""
    $script:AnalyticsOsBuild         = 0
    $script:AnalyticsMissingBefore   = -1
    $script:AnalyticsMissingAfter    = -1
    $script:AnalyticsStartTime       = Get-Date
    $script:7zInstalled              = $false
    $script:Headless                 = [bool]$Headless

    Set-ButtonRunning
    SetProgress 0
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."
    $exGroupBox.Text          = "Extract"
    $dlSpinnerLabel.Text      = ""
    $exSpinnerLabel.Text      = ""
    $overallSpinnerLabel.Text = ""
    $script:SpinnerIndex      = 0

    $id  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pri = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $pri.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if ($script:Headless) {
            Write-Error "ERROR: Script must be run as Administrator."
            exit 1
        }
        Log "Not running as admin - re-launching elevated..."
        Start-Process powershell `
            "-WindowStyle Hidden -ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex`"" `
            -Verb RunAs
        $form.Close()
        exit
    }

    Log "Driver Installer v$ScriptVersion"
    Log "Log: $LogFile"
    Log "--------------------------------------------"

    Play-Sound -Event "Start"

    $cs = Get-CimInstance Win32_ComputerSystem

    # Use param overrides if provided, otherwise read from WMI
    $manufacturer = if ($Manufacturer) { $Manufacturer } else { $cs.Manufacturer.Trim() }
    $model        = if ($Model)        { $Model }        else { $cs.Model.Trim() }

    $script:AnalyticsManufacturer = $manufacturer
    $script:AnalyticsModel        = $model
    try {
        $osObj = Get-CimInstance Win32_OperatingSystem
        $script:AnalyticsOsVersion = $osObj.Version
        $script:AnalyticsOsBuild   = [int]$osObj.BuildNumber
    } catch {}

    if (-not $script:Headless) {
        try { [System.Windows.Forms.Clipboard]::SetText($model) } catch {}
        $title.Text = "Driver Installer - $model"
    }

    $overrideNote = if ($Manufacturer -or $Model) { "  [OVERRIDDEN via param]" } else { "  (from WMI)" }
    Log "Manufacturer : $manufacturer$overrideNote"
    Log "Model        : $model$overrideNote"

    Write-DeviceInfo
    SetProgress 5

    # Install 7-Zip for Dell and HP extraction (not needed for Lenovo or Surface)
    if ($manufacturer -match "Dell|HP|Hewlett") {
        Log "Installing 7-Zip for fast extraction..."
        if (-not (Install-7Zip)) {
            Log "WARNING: 7-Zip unavailable - will fall back to vendor extractor."
        }
    }

    # Snapshot missing drivers before install
    Log "Checking for devices with missing drivers..."
    $script:AnalyticsMissingBefore = Get-MissingDriverCount
    Log "Missing drivers BEFORE install: $($script:AnalyticsMissingBefore)"

    $driverRoot = "C:\DRIVERS"
    $success    = $false

    if ($manufacturer -match "Dell") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-DellDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "HP|Hewlett") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-HpDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "Lenovo") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-LenovoDriverInstall -DriverRoot $driverRoot
    } elseif ($manufacturer -match "Microsoft") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-MicrosoftSurfaceDriverInstall -DriverRoot $driverRoot -ModelName $model
    } else {
        Log "Unsupported manufacturer: $manufacturer"
        Log "Supported OEMs: Dell, HP, Lenovo, Microsoft (Surface)"
        Send-AnalyticsEvent -Result "failure"
        if ($script:Headless) {
            Write-Host "UNSUPPORTED: Manufacturer '$manufacturer' is not supported."
            Write-Host "Supported: Dell, HP, Lenovo, Microsoft (Surface)"
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Manufacturer '$manufacturer' is not supported.`nSupported: Dell, HP, Lenovo, Microsoft (Surface)",
                "Unsupported Manufacturer", "OK", "Warning"
            )
        }
        # Remove 7-Zip before returning on unsupported manufacturer
        if ($script:7zInstalled) { Remove-7Zip }
        Set-ButtonIdle
        return
    }

    # Remove 7-Zip before cleanup - must happen before C:\DRIVERS is deleted
    if ($script:7zInstalled) { Remove-7Zip }

    # Snapshot missing drivers after install
    $script:AnalyticsMissingAfter = Get-MissingDriverCount
    $missingDelta = if ($script:AnalyticsMissingBefore -ge 0 -and $script:AnalyticsMissingAfter -ge 0) {
        $script:AnalyticsMissingBefore - $script:AnalyticsMissingAfter
    } else { 0 }
    Log "Missing drivers AFTER  install: $($script:AnalyticsMissingAfter)"
    Log "Devices resolved by this install: $missingDelta"

    if ($script:CancelRequested) { Send-AnalyticsEvent -Result "cancelled" }

    Log "--------------------------------------------"
    if ($success) {
        SetProgress 100
        SetDownload -Pct 100 -Label "Complete"
        SetExtract  -Pct 100 -Label "Complete"
        Log "Driver installation complete!"
        Log "Log saved to: $LogFile"
        Send-AnalyticsEvent -Result "success"
        Play-Sound -Event "Success"

        if ($SkipCleanup) {
            Log "SkipCleanup flag set - keeping $driverRoot for inspection."
        } elseif (Test-Path $driverRoot) {
            Log "Cleaning up $driverRoot..."
            try {
                Remove-Item $driverRoot -Recurse -Force -ErrorAction Stop
                Log "  $driverRoot removed."
            } catch { Log "  WARNING: Could not remove $driverRoot - $($_.Exception.Message)" }
        }

        $missingLine = if ($script:AnalyticsMissingBefore -ge 0 -and $script:AnalyticsMissingAfter -ge 0) {
            "`n`nMissing drivers:  $($script:AnalyticsMissingBefore) -> $($script:AnalyticsMissingAfter)  ($missingDelta resolved)"
        } else { "" }

        if ($script:Headless) {
            Write-Host "SUCCESS: Drivers installed for $model.$missingLine"
            Write-Host "Run complete. Reboot when ready."
        } else {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Drivers installed successfully for:`n$model$missingLine`n`nReboot now to complete installation?",
                "Installation Complete", "YesNo", "Information"
            )
            if ($result -eq "Yes") { Restart-Computer -Force }
            else { Set-ButtonIdle }
        }
    } else {
        if (-not $script:CancelRequested) { Send-AnalyticsEvent -Result "failure" }
        SetDownload -Pct 0 -Label "Failed - see log"
        SetExtract  -Pct 0 -Label "Failed - see log"
        Stop-DlSpinner      -Success $false
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Play-Sound -Event "Failure"
        Log "Driver installation did not complete. Check log: $LogFile"
        if ($script:Headless) {
            Write-Host "FAILED: Driver installation did not complete. Check log: $LogFile"
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Driver installation failed or no pack was found.`nCheck the log:`n`n$LogFile",
                "Installation Failed", "OK", "Error"
            )
            Set-ButtonIdle
        }
    }
}

# =========================
# WIRE UP + LAUNCH
# =========================
if ($Headless) {
    # Headless mode - run directly, no GUI
    Start-Install
} else {
    # GUI mode - wire up form and show
    $button.Add_Click({ Start-Install })

    $cancelButton.Add_Click({
        if ($cancelButton.Enabled) {
            $script:CancelRequested = $true
            Log "--- Cancel requested by user ---"
            Play-Sound -Event "Cancel"
            $cancelButton.Enabled   = $false
            $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
            [System.Windows.Forms.Application]::DoEvents()
        }
    })

    $form.Add_Shown({
        $form.Activate()
        Start-Sleep -Milliseconds 300
        Log "Running startup checks..."
        Start-Install
    })

    [void]$form.ShowDialog()
}