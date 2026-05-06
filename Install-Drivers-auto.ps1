# =============================================================
# Install-Drivers-auto.ps1
# Version: 1.4.1
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode:
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
#
# Supports: Dell, HP, Lenovo
# =============================================================

$ScriptVersion   = "1.4.1"
$SpinnerFrames   = @('⠋','⠙','⠹','⠸','⠼','⠴','⠦','⠧','⠇','⠏')
$SpinnerIndex    = 0
$CancelRequested = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

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
$form.Size            = New-Object System.Drawing.Size(580, 530)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 245)

# Title label
$title           = New-Object System.Windows.Forms.Label
$title.AutoSize  = $true
$title.Font      = $FontTitleBold
$title.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$title.Location  = New-Object System.Drawing.Point(20, 15)
$title.Text      = "Driver Installer"
$title.UseCompatibleTextRendering = $false
$form.Controls.Add($title)

# Version label
$versionLabel           = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize  = $true
$versionLabel.Font      = $FontUISmall
$versionLabel.ForeColor = [System.Drawing.Color]::Gray
$versionLabel.Text      = "v$ScriptVersion"
$versionLabel.Location  = New-Object System.Drawing.Point(510, 20)
$versionLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($versionLabel)

# Status box
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

# Log path label
$logLabel           = New-Object System.Windows.Forms.Label
$logLabel.AutoSize  = $false
$logLabel.Size      = New-Object System.Drawing.Size(536, 16)
$logLabel.Location  = New-Object System.Drawing.Point(20, 468)
$logLabel.ForeColor = [System.Drawing.Color]::Gray
$logLabel.Font      = $FontUISmall
$logLabel.Text      = "Log: $LogFile"
$logLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($logLabel)

# Install button
$button            = New-Object System.Windows.Forms.Button
$button.Text       = "Install Drivers"
$button.Size       = New-Object System.Drawing.Size(155, 36)
$button.Location   = New-Object System.Drawing.Point(155, 488)
$button.Font       = $FontUIBold
$button.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
$button.ForeColor  = [System.Drawing.Color]::White
$button.FlatStyle  = "Flat"
$button.FlatAppearance.BorderSize = 0
$form.Controls.Add($button)

# Cancel button
$cancelButton            = New-Object System.Windows.Forms.Button
$cancelButton.Text       = "Cancel"
$cancelButton.Size       = New-Object System.Drawing.Size(100, 36)
$cancelButton.Location   = New-Object System.Drawing.Point(320, 488)
$cancelButton.Font       = $FontUIBold
$cancelButton.BackColor  = [System.Drawing.Color]::FromArgb(160, 160, 160)
$cancelButton.ForeColor  = [System.Drawing.Color]::White
$cancelButton.FlatStyle  = "Flat"
$cancelButton.FlatAppearance.BorderSize = 0
$cancelButton.Enabled    = $false
$form.Controls.Add($cancelButton)

# =========================
# BUTTON STATE HELPERS
# =========================
function Set-ButtonRunning {
    $button.Enabled         = $false
    $button.BackColor       = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $cancelButton.Enabled   = $true
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60)
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ButtonIdle {
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
    $statusBox.AppendText("$line`r`n")
    $statusBox.ScrollToCaret()
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    [System.Windows.Forms.Application]::DoEvents()
}

function SetProgress($val) {
    $progress.Value = [math]::Min([math]::Max([int]$val, 0), 100)
    [System.Windows.Forms.Application]::DoEvents()
}

function SetDownload {
    param([int]$Pct, [string]$Label)
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
    $script:SpinnerIndex = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $dlSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-DlSpinner {
    param([bool]$Success = $true)
    $dlSpinnerLabel.Text      = if ($Success) { " ✓" } else { " ✗" }
    $dlSpinnerLabel.ForeColor = if ($Success) {
        [System.Drawing.Color]::FromArgb(0, 100, 180)
    } else {
        [System.Drawing.Color]::FromArgb(200, 40, 40)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-ExSpinner {
    $script:SpinnerIndex  = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $exSpinnerLabel.Text  = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-ExSpinner {
    param([bool]$Success = $true)
    $exSpinnerLabel.Text      = if ($Success) { " ✓" } else { " ✗" }
    $exSpinnerLabel.ForeColor = if ($Success) {
        [System.Drawing.Color]::FromArgb(0, 140, 80)
    } else {
        [System.Drawing.Color]::FromArgb(200, 40, 40)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-OverallSpinner {
    $script:SpinnerIndex      = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $overallSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-OverallSpinner {
    param([bool]$Success = $true)
    $overallSpinnerLabel.Text      = if ($Success) { " ✓" } else { " ✗" }
    $overallSpinnerLabel.ForeColor = if ($Success) {
        [System.Drawing.Color]::FromArgb(80, 80, 80)
    } else {
        [System.Drawing.Color]::FromArgb(200, 40, 40)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-AllSpinners {
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

    # HEAD request to get file size
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
            Log "  WARNING: download stalled 3.5 min — aborting."
            $proc.Kill()
            SetDownload -Pct 0 -Label "Stalled — aborted."
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
        SetDownload -Pct 0 -Label "Failed — file empty."
        Stop-DlSpinner -Success $false
        return $false
    }

    $finalMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 1)
    Log "  Download complete: $finalMB MB"
    SetDownload -Pct 100 -Label "Complete — $finalMB MB"
    Stop-DlSpinner -Success $true
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
            Log "  WARNING: extraction stalled — killing process."
            try { $ExtractProc.Kill() } catch {}
            break
        }
    }

    Start-Sleep -Seconds 2
    $finalCount = if (Test-Path $DestPath) {
        (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
    } else { 0 }

    SetExtract -Pct 100 -Label "Done — $finalCount files extracted"
    Stop-ExSpinner      -Success $true
    Stop-OverallSpinner -Success $true
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
    Log "Found $total INF file(s) — installing via pnputil..."
    $exGroupBox.Text     = "Install INFs"
    $exSpinnerLabel.Text = " " + $SpinnerFrames[0]
    $script:SpinnerIndex = 0
    $exBar.Style         = "Continuous"
    $exBar.Value         = 0

    foreach ($inf in $infs) {
        $i++
        $overallPct = 60 + [int](($i / $total) * 38)
        SetProgress $overallPct
        $infPct    = [int](($i / $total) * 100)
        $remaining = $total - $i
        SetExtract -Pct $infPct -Label "$i / $total INFs  ($remaining remaining)  —  $($inf.Name)"
        Log "[$i/$total] $($inf.Name)"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        foreach ($l in $out) { Log "  $l" }
        Step-ExSpinner
        Step-OverallSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "INF installation cancelled at $i / $total"
            break
        }
    }

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
# HP SoftPaq EXEs use:  /s /e /f "<dest>"
# Dell/Lenovo use Inno: /VERYSILENT /DIR="<dest>" /EXTRACT=YES
# We try HP flags first; if no files appear after 5 s we retry with Inno flags.
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
                if ($stall -gt 375) { Log "  ZIP stalled — stopping."; Stop-Job $zipJob; break }
            }
            Receive-Job $zipJob -EA SilentlyContinue | Out-Null
            Remove-Job  $zipJob
            $finalCount = if (Test-Path $DestPath) {
                (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
            } else { 0 }
            SetExtract -Pct 100 -Label "Done — $finalCount files extracted"
            Stop-ExSpinner      -Success $true
            Stop-OverallSpinner -Success $true
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
            # EXE extractor format varies by OEM:
            #   Dell:    /s /e="<dest>"                        (synchronous)
            #   Lenovo:  -s -f"<dest>"                         (synchronous)
            #   HP:      /s /e /f "<dest>"                     (async)
            #   Inno:    /VERYSILENT /DIR="<dest>" /EXTRACT=YES (async, last resort)
            #
            # We try the vendor-specific format first and stop immediately if it
            # produces files. Only fall through to the other formats if it fails.

            Log "Extracting EXE pack (Vendor=$Vendor)..."

            $CountFiles = {
                if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
            }

            function TrySync {
                param([string]$ExeArgs)
                $p = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = $ExeArgs
                # UseShellExecute = $true + WindowStyle Hidden causes Windows to
                # set SW_HIDE in STARTUPINFO, which child GUI processes inherit.
                # CreateNoWindow alone only suppresses the direct console window.
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                $proc.WaitForExit()
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            function TryAsyncHP {
                $p = New-Object System.Diagnostics.ProcessStartInfo
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
                        Log "  HP format timed out with no output — killing."
                        try { $proc.Kill() } catch {}
                        break
                    }
                }
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            # Build the attempt list — vendor-specific format goes first
            # Each entry: [label, scriptblock that returns file count]
            $attempts = [System.Collections.Generic.List[object]]::new()

            switch ($Vendor) {
                "Dell" {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";  Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (-s -fdest)"; Action = { TrySync "-s -f`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)"; Action = { TryAsyncHP } })
                }
                "Lenovo" {
                    $attempts.Add(@{ Label = "Lenovo (-s -fdest)"; Action = { TrySync "-s -f`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";  Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)"; Action = { TryAsyncHP } })
                }
                "HP" {
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)"; Action = { TryAsyncHP } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";  Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (-s -fdest)"; Action = { TrySync "-s -f`"$DestPath`"" } })
                }
                default {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";  Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (-s -fdest)"; Action = { TrySync "-s -f`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)"; Action = { TryAsyncHP } })
                }
            }

            $extracted = $false
            foreach ($attempt in $attempts) {
                Log "  Trying: $($attempt.Label)..."
                $n = & $attempt.Action
                Log "  Result: $n files in $DestPath"
                if ($n -gt 0) {
                    SetExtract -Pct 100 -Label "Done — $n files extracted"
                    Stop-ExSpinner -Success $true; Stop-OverallSpinner -Success $true
                    Log "  Extraction finished: $n files in $DestPath"
                    $extracted = $true
                    break
                }
            }

            # Last resort: Inno Setup (only if all vendor formats failed)
            if (-not $extracted) {
                Log "  All vendor formats failed — trying Inno Setup (/VERYSILENT /DIR /EXTRACT=YES)..."
                $p = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES"
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
    } catch {
        Log "Could not read Service Tag: $($_.Exception.Message)"
        return $false
    }
    if (-not $serviceTag -or $serviceTag.Length -lt 4) {
        Log "Invalid Service Tag."
        return $false
    }

    # Detect OS build for Win10/Win11 pack preference
    $isWin11 = $false
    try {
        $isWin11 = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber -ge 22000
        Log "OS: $(if ($isWin11) {'Windows 11'} else {'Windows 10'})"
    } catch {}

    # ---------------------------------------------------------------
    # Dell publishes TWO catalogs:
    #   CatalogPC.cab       - all individual drivers/firmware (NOT driver packs)
    #   DriverPackCatalog.cab - driver packs only (what we need)
    #
    # DriverPackCatalog.xml structure (confirmed by Dell docs):
    #   <DriverPackManifest>
    #     <DriverPackage path="FOLDER.../Model-Platform_Win11_x.x_Axx.exe" ...>
    #       <Name><Display lang="en">Latitude 7430 Windows 11 Driver Pack</Display></Name>
    #       <SupportedSystems>
    #         <Brand key="4" prefix="LAT">
    #           <Model systemID="0B0B" name="Latitude 7430">
    #             <Display>Latitude-7430</Display>
    #           </Model>
    #         </Brand>
    #       </SupportedSystems>
    #       <SupportedOperatingSystems>
    #         <OperatingSystem osCode="W21P4" ...>
    #           <Display>Windows 11</Display>
    #         </OperatingSystem>
    #       </SupportedOperatingSystems>
    #     </DriverPackage>
    #   </DriverPackManifest>
    #
    # Match strategy (per Dell docs): use the 'name' attribute on <Model> nodes
    # since systemID is not readily accessible via WMI. The 'name' value matches
    # the WMI Win32_ComputerSystem.Model string (e.g. "Latitude 7430").
    # ---------------------------------------------------------------

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

    # Build model name search tokens from WMI model string
    # WMI: "Latitude 7430" — the catalog 'name' attribute uses the same value
    # Also try stripping the family prefix in case of slight differences
    $searchNames = @()
    if ($ModelName) {
        $searchNames += $ModelName                                      # "Latitude 7430"
        $searchNames += ($ModelName -replace '\s+Notebook.*$','')       # strip "Notebook PC" suffix
        $searchNames += ($ModelName -replace '^Dell\s+','')             # strip leading "Dell "
        $searchNames = $searchNames | Select-Object -Unique | Where-Object { $_.Length -gt 3 }
    }
    Log "Searching catalog — model tokens: $($searchNames -join ' | ')"

    # Collect all matching driver pack entries
    # Root element is <DriverPackManifest>, packages are <DriverPackage> nodes
    $candidates = @()
    foreach ($pkg in $cat.SelectNodes("//*[local-name()='DriverPackage']")) {

        # Match on 'name' attribute of <Model> nodes (Dell-recommended method).
        # Use exact case-insensitive equality only — substring matching causes
        # "Latitude 7330 Rugged Extreme" to match a search for "Latitude 7330".
        $modelMatched = $false
        foreach ($modelNode in $pkg.SelectNodes(".//*[local-name()='Model']")) {
            $nameAttr = $modelNode.GetAttribute("name")
            foreach ($tok in $searchNames) {
                if ($nameAttr -ieq $tok) {
                    $modelMatched = $true; break
                }
            }
            if ($modelMatched) { break }
        }
        if (-not $modelMatched) { continue }

        # Determine OS support from <SupportedOperatingSystems>
        $supportsWin11 = $false
        $supportsWin10 = $false
        foreach ($osNode in $pkg.SelectNodes(".//*[local-name()='OperatingSystem']")) {
            $osDisp = ""
            try { $osDisp = $osNode.SelectSingleNode("*[local-name()='Display']").InnerText } catch {}
            if ($osDisp -match "(?i)windows 11") { $supportsWin11 = $true }
            if ($osDisp -match "(?i)windows 10") { $supportsWin10 = $true }
        }

        $pkgName = ""
        try { $pkgName = $pkg.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch {}
        $pkgPath = $pkg.GetAttribute("path")

        $candidates += [PSCustomObject]@{
            Path        = $pkgPath
            DisplayName = $pkgName
            Win11       = $supportsWin11
            Win10       = $supportsWin10
        }
        Log "  Candidate: $pkgName  [W11=$supportsWin11 W10=$supportsWin10]"
    }

    if ($candidates.Count -eq 0) {
        Log "No driver pack found in DriverPackCatalog for '$ModelName'."
        Log "Opening Dell support page..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    # Pick best match for detected OS
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
    if (-not $packPath) {
        Log "Driver pack entry has no path — unexpected catalog format."
        return $false
    }

    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Pack file: $([System.IO.Path]::GetFileName($packPath))"
    Log "Pack URL:  $packUrl"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) {
        New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null
    }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) {
        Log "Dell driver pack download failed."
        return $false
    }
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
# Strategy: download the HP Driver Pack Matrix HTML page (the authoritative
# source HP actually maintains), locate the row matching this machine's model
# name, and pick the best SoftPaq EXE URL for the detected OS version.
#
# The matrix table structure (from ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html):
#   - Each <tr> starts with a cell listing one or more model names
#   - Subsequent cells contain either "-" or an <a href="...spNNNNN.exe"> link
#   - Column order (left to right) is newest OS to oldest OS
#
# We download the raw HTML, find the table row(s) whose model cell contains
# our model string, then walk across the columns preferring the best OS match.
# =========================
function Start-HpDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== HP: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    # Detect OS build to prefer the right column
    $osBuild   = $null
    $isWin11   = $false
    try {
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
        $isWin11 = $osBuild -ge 22000
        Log "OS Build: $osBuild  ($(if ($isWin11) {'Win11'} else {'Win10'}))"
    } catch {
        Log "Could not read OS build: $($_.Exception.Message)"
    }

    # Download the Driver Pack Matrix HTML — this is the page HP maintains with
    # every supported model and direct .exe links, so it's always up to date.
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

    # Build search tokens from the model name.
    # e.g. "HP ProBook 430 G8 Notebook PC" -> try progressively shorter substrings
    # so we can still match if HP abbreviates the name in the table.
    $searchTokens = @()
    if ($ModelName) {
        # Remove "HP " prefix since it appears on every row and would match everything
        $stripped = $ModelName -replace '(?i)^HP\s+', ''
        $searchTokens += $stripped                                    # "ProBook 430 G8 Notebook PC"
        $searchTokens += ($stripped -replace '\s+Notebook.*$', '')   # "ProBook 430 G8"
        $searchTokens += ($stripped -replace '\s+PC.*$',       '')   # may trim further
        $searchTokens = $searchTokens | Select-Object -Unique | Where-Object { $_.Length -gt 4 }
    }
    Log "Search tokens: $($searchTokens -join ' | ')"

    # Parse table rows with a regex — we need the raw HTML cells, not markdown.
    # Pattern: find <tr>...</tr> blocks, split into <td>/<th> cells, check first
    # cell for a model name match, then scan remaining cells for .exe hrefs.
    $packUrl   = $null
    $packSpNum = $null

    # Normalise line endings and collapse whitespace inside tags for easier regex
    $flat = $matrixHtml -replace "`r`n|`r|`n", " " -replace "\s{2,}", " "

    # Extract all <tr>...</tr> blocks
    $rows = [regex]::Matches($flat, '(?i)<tr[^>]*>(.*?)</tr>')

    foreach ($row in $rows) {
        $rowHtml = $row.Groups[1].Value

        # Split into cells
        $cells = [regex]::Matches($rowHtml, '(?i)<t[dh][^>]*>(.*?)</t[dh]>')
        if ($cells.Count -lt 2) { continue }

        # First cell = model name(s) — strip all tags to get plain text
        $modelCell = [regex]::Replace($cells[0].Groups[1].Value, '<[^>]+>', ' ')
        $modelCell = [System.Net.WebUtility]::HtmlDecode($modelCell) -replace '\s+', ' '

        # Check if any of our search tokens appear in this cell
        $matched = $false
        foreach ($tok in $searchTokens) {
            if ($modelCell -match [regex]::Escape($tok)) { $matched = $true; break }
        }
        if (-not $matched) { continue }

        Log "  Matched matrix row: $($modelCell.Trim() -replace '\s+',' ')"

        # Collect all spNNNNN.exe hrefs from this row in column order (left = newest OS)
        $allLinks = [regex]::Matches($rowHtml, '(?i)href="([^"]*sp\d+\.exe)"')
        if ($allLinks.Count -eq 0) {
            Log "  Row matched but contains no .exe links — skipping."
            continue
        }

        # All links are valid candidates. We want the newest (leftmost = index 0).
        # But we do a light preference: if Win11 pick first link, if Win10 skip
        # pure Win11-only links by checking the tooltip text for "Windows 10".
        # In practice for most models the same pack covers both, so just take first.
        $bestUrl = $allLinks[0].Groups[1].Value
        if (-not $bestUrl.StartsWith("http")) {
            $bestUrl = "https://ftp.hp.com$bestUrl"
        }

        # If the machine is Win10 and the first link's tooltip only mentions Win11,
        # walk forward to find a Win10-labelled link.
        if (-not $isWin11) {
            foreach ($lm in $allLinks) {
                $href = $lm.Groups[1].Value
                if (-not $href.StartsWith("http")) { $href = "https://ftp.hp.com$href" }
                # The title attribute of the surrounding <a> tag has the version tooltip
                $aTag = [regex]::Match($rowHtml, "(?i)<a[^>]+href=""[^""]*$([regex]::Escape([System.IO.Path]::GetFileName($href)))[^""]*""[^>]*>")
                $title = if ($aTag.Success) { $aTag.Value } else { "" }
                # Accept if title is blank (we can't tell) or explicitly mentions Win10
                if ($title -eq "" -or $title -match "(?i)windows 10") {
                    $bestUrl = $href
                    break
                }
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

    SetExtract  -Pct 40 -Label "Matrix OK — $packSpNum"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) {
        New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null
    }

    $packFile = Join-Path $DriverRoot "$packSpNum.exe"
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) {
        Log "HP driver pack download failed."
        return $false
    }
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
    try {
        $sku = (Get-CimInstance Win32_ComputerSystemProduct).Name.Trim()
        if ($sku.Length -ge 4) {
            $machineType = $sku.Substring(0, 4).ToUpper()
            Log "Machine type: $sku  ->  prefix: $machineType"
        }
    } catch {
        Log "Could not read machine type: $($_.Exception.Message)"
    }
    if (-not $machineType) {
        Log "Cannot determine Lenovo machine type."
        return $false
    }

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
    } catch {
        Log "Failed to parse Lenovo catalog: $($_.Exception.Message)"
        return $false
    }
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
                if ($url -match "^https?://") {
                    $packUrl = $url
                    Log "Driver pack URL [$os]: $packUrl"
                    break
                }
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

    if (-not (Test-Path $DriverRoot)) {
        New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null
    }
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) {
        Log "Lenovo driver pack download failed."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Lenovo_Extracted"
    Log "Extracting Lenovo pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "Lenovo"
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# DEVICE INFO DUMP
# Collects every WMI/CIM field that could be useful for debugging
# driver pack lookup failures. All errors are caught individually so
# one bad query never stops the rest from printing.
# =========================
function Write-DeviceInfo {
    Log "============================================"
    Log "  DEVICE INFORMATION DUMP"
    Log "============================================"

    # ---- System / Chassis ----
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

    # ---- BIOS ----
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

    # ---- Baseboard (HP Platform ID lives here) ----
    Log "-- Baseboard --"
    try {
        $bb = Get-CimInstance Win32_BaseBoard
        Log "  Product            : $($bb.Product)"
        Log "  Manufacturer       : $($bb.Manufacturer)"
        Log "  SerialNumber       : $($bb.SerialNumber)"
        Log "  Version            : $($bb.Version)"
    } catch { Log "  [Win32_BaseBoard ERROR] $($_.Exception.Message)" }

    # ---- Enclosure / Chassis type ----
    Log "-- Enclosure --"
    try {
        $enc = Get-CimInstance Win32_SystemEnclosure
        Log "  ChassisTypes       : $($enc.ChassisTypes -join ',')"
        Log "  SMBIOSAssetTag     : $($enc.SMBIOSAssetTag)"
        Log "  SerialNumber       : $($enc.SerialNumber)"
        Log "  Manufacturer       : $($enc.Manufacturer)"
    } catch { Log "  [Win32_SystemEnclosure ERROR] $($_.Exception.Message)" }

    # ---- Operating System ----
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

    # ---- CPU ----
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

    # ---- GPU ----
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

    # ---- Network adapters ----
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

    # ---- Storage ----
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

    # ---- RAM sticks ----
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

    # ---- Battery (laptops) ----
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
        } else {
            Log "  (no battery detected — desktop?)"
        }
    } catch { Log "  [Win32_Battery ERROR] $($_.Exception.Message)" }

    # ---- PnP devices with missing drivers ----
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
        } else {
            Log "  (none — all PnP devices have drivers)"
        }
    } catch { Log "  [Win32_PnPEntity ERROR] $($_.Exception.Message)" }

    # ---- Installed drivers via pnputil ----
    Log "-- pnputil driver store (first 20 OEM INFs) --"
    try {
        $pnpOut = & pnputil /enum-drivers 2>&1 | Select-Object -First 60
        foreach ($line in $pnpOut) { Log "  $line" }
    } catch { Log "  [pnputil ERROR] $($_.Exception.Message)" }

    # ---- Environment snapshot ----
    Log "-- Environment --"
    Log "  TEMP               : $env:TEMP"
    Log "  COMPUTERNAME       : $env:COMPUTERNAME"
    Log "  USERNAME           : $env:USERNAME"
    Log "  PROCESSOR_ARCH     : $env:PROCESSOR_ARCHITECTURE"
    Log "  PS Version         : $($PSVersionTable.PSVersion)"
    Log "  curl.exe path      : $($(Get-Command curl.exe -EA SilentlyContinue).Source)"

    # ---- Free disk space on C: ----
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

    $script:CancelRequested = $false
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
        Log "Not running as admin — re-launching elevated..."
        Start-Process powershell `
            "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex`"" `
            -Verb RunAs
        exit
    }

    Log "Driver Installer v$ScriptVersion"
    Log "Log: $LogFile"
    Log "--------------------------------------------"

    $cs           = Get-CimInstance Win32_ComputerSystem
    $manufacturer = $cs.Manufacturer.Trim()
    $model        = $cs.Model.Trim()

    try { [System.Windows.Forms.Clipboard]::SetText($model) } catch {}

    Log "Manufacturer : $manufacturer"
    Log "Model        : $model  (copied to clipboard)"
    $title.Text = "Driver Installer - $model"

    # Full device dump — logged before any OEM work so the log is useful even on early failure
    Write-DeviceInfo

    SetProgress 5

    $driverRoot = "C:\DRIVERS"
    $success    = $false

    if ($manufacturer -match "Dell") {
        if (-not (Assert-Curl)) { Set-ButtonIdle; return }
        $success = Start-DellDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "HP|Hewlett") {
        if (-not (Assert-Curl)) { Set-ButtonIdle; return }
        $success = Start-HpDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "Lenovo") {
        if (-not (Assert-Curl)) { Set-ButtonIdle; return }
        $success = Start-LenovoDriverInstall -DriverRoot $driverRoot
    } else {
        Log "Unsupported manufacturer: $manufacturer"
        Log "Supported OEMs: Dell, HP, Lenovo"
        [System.Windows.Forms.MessageBox]::Show(
            "Manufacturer '$manufacturer' is not supported.`nSupported: Dell, HP, Lenovo",
            "Unsupported Manufacturer", "OK", "Warning"
        )
        Set-ButtonIdle
        return
    }

    Log "--------------------------------------------"
    if ($success) {
        SetProgress 100
        SetDownload -Pct 100 -Label "Complete"
        SetExtract  -Pct 100 -Label "Complete"
        Log "Driver installation complete!"
        Log "Log saved to: $LogFile"

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Drivers installed successfully for:`n$model`n`nReboot now to complete installation?",
            "Installation Complete", "YesNo", "Information"
        )
        if ($result -eq "Yes") { Restart-Computer -Force }
        else { Set-ButtonIdle }
    } else {
        SetDownload -Pct 0 -Label "Failed — see log"
        SetExtract  -Pct 0 -Label "Failed — see log"
        Stop-DlSpinner      -Success $false
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Log "Driver installation did not complete. Check log: $LogFile"
        [System.Windows.Forms.MessageBox]::Show(
            "Driver installation failed or no pack was found.`nCheck the log:`n`n$LogFile",
            "Installation Failed", "OK", "Error"
        )
        Set-ButtonIdle
    }
}

# =========================
# WIRE UP + LAUNCH
# =========================
$button.Add_Click({ Start-Install })

$cancelButton.Add_Click({
    if ($cancelButton.Enabled) {
        $script:CancelRequested = $true
        Log "--- Cancel requested by user ---"
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