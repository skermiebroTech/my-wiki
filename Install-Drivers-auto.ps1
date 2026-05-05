# =============================================================
# Install-Drivers-auto.ps1
# Version: 1.3.6
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode:
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
#
# Supports: Dell, HP, Lenovo
# =============================================================

$ScriptVersion   = "1.3.6"
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
        [int]$StallLimitSec = 300
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
            # Try HP SoftPaq silent-extract flags first
            Log "Extracting EXE (trying HP SoftPaq flags)..."
            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = $PackFile
            $psi.Arguments       = "/s /e /f `"$DestPath`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            $exProc              = New-Object System.Diagnostics.Process
            $exProc.StartInfo    = $psi
            $exProc.Start() | Out-Null

            # Wait 5 s then check whether any files appeared
            Start-Sleep -Seconds 5
            $earlyCount = if (Test-Path $DestPath) {
                (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
            } else { 0 }

            if ($earlyCount -eq 0 -and -not $exProc.HasExited) {
                Log "  HP flags produced no output — retrying with Inno Setup flags..."
                try { $exProc.Kill() } catch {}
                Start-Sleep -Seconds 1
                $psi2                 = New-Object System.Diagnostics.ProcessStartInfo
                $psi2.FileName        = $PackFile
                $psi2.Arguments       = "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES"
                $psi2.UseShellExecute = $false
                $psi2.CreateNoWindow  = $true
                $exProc               = New-Object System.Diagnostics.Process
                $exProc.StartInfo     = $psi2
                $exProc.Start() | Out-Null
            }

            Watch-Extraction -ExtractProc $exProc -DestPath $DestPath -StallLimitSec $StallLimitSec
        }
    }
}

# =========================
# HP SoftPaq URL builder
# HP stores EXEs at: https://ftp.hp.com/pub/softpaq/sp{low}-{high}/{spNum}.exe
# Bucket size = 500, e.g. sp171622 -> sp171501-172000
# =========================
function Get-HpSoftpaqUrl {
    param([string]$SpNum)
    # Strip any leading "sp" prefix and extract digits
    $digits = $SpNum -replace '[^0-9]', ''
    $low    = ([math]::Floor([int]$digits / 500) * 500) + 1
    $high   = $low + 499
    return "https://ftp.hp.com/pub/softpaq/sp${low}-${high}/sp${digits}.exe"
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
    $sysId      = $null
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
    try {
        $sysId = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber.Trim()
        Log "System SKU: $sysId"
    } catch {
        Log "Could not read SystemSKUNumber: $($_.Exception.Message)"
    }

    # Download Dell catalog CAB
    Log "Downloading Dell CatalogPC.cab..."
    $catalogCab = Join-Path $env:TEMP "DellCatalogPC.cab"
    $catalogXml = Join-Path $env:TEMP "CatalogPC.xml"
    Remove-Item $catalogCab -EA SilentlyContinue
    Remove-Item $catalogXml -EA SilentlyContinue

    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/CatalogPC.cab" -OutFile $catalogCab)) {
        Log "Failed to download Dell catalog."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 15

    # Extract catalog — check explicitly
    Log "Extracting Dell catalog CAB..."
    SetExtract -Pct 10 -Label "Extracting Dell catalog..."
    $expandOut = & expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1
    Log "  expand.exe: $expandOut"
    if (-not (Test-Path $catalogXml) -or (Get-Item $catalogXml).Length -eq 0) {
        Log "Dell catalog CAB extraction failed."
        return $false
    }
    SetExtract -Pct 30 -Label "Catalog extracted OK"
    SetProgress 20

    # Parse catalog
    Log "Parsing Dell catalog..."
    try {
        $rawXml   = [System.IO.File]::ReadAllText($catalogXml).TrimStart([char]0xFEFF)
        [xml]$cat = $rawXml
    } catch {
        Log "Failed to parse Dell catalog XML: $($_.Exception.Message)"
        return $false
    }

    # Build SKU candidate list
    $skuCandidates = @()
    if ($sysId) {
        $skuCandidates += $sysId
        $skuCandidates += $sysId.TrimStart('0')
        $skuCandidates += $sysId.ToUpper()
        $skuCandidates += ('0x' + $sysId)
        $skuCandidates = $skuCandidates | Select-Object -Unique
    }

    # Search catalog using local-name() to avoid namespace issues
    $packNode = $null
    foreach ($comp in $cat.SelectNodes("//*[local-name()='SoftwareComponent']")) {
        $dispName = ""
        try {
            $dispName = $comp.SelectSingleNode(
                "*[local-name()='Name']/*[local-name()='Display']").InnerText
        } catch {}
        if ($dispName -notmatch "(?i)driver\s*pack") { continue }

        if ($skuCandidates.Count -gt 0) {
            $matched = $false
            foreach ($sNode in $comp.SelectNodes(".//*[local-name()='SystemID']")) {
                $val = $sNode.InnerText.Trim()
                foreach ($candidate in $skuCandidates) {
                    if ($val -ieq $candidate) { $matched = $true; break }
                }
                if ($matched) { break }
            }
            if (-not $matched) { continue }
        }
        $packNode = $comp
        break
    }

    # Fallback: model name substring match
    if (-not $packNode -and $ModelName) {
        Log "SKU match failed — trying model name fallback..."
        $modelShort = ($ModelName -replace '[^a-zA-Z0-9]', '').ToLower()
        foreach ($comp in $cat.SelectNodes("//*[local-name()='SoftwareComponent']")) {
            $dispName = ""
            try {
                $dispName = $comp.SelectSingleNode(
                    "*[local-name()='Name']/*[local-name()='Display']").InnerText
            } catch {}
            if ($dispName -notmatch "(?i)driver\s*pack") { continue }
            $dispClean = ($dispName -replace '[^a-zA-Z0-9]', '').ToLower()
            if ($dispClean -like "*$modelShort*" -or $modelShort -like "*$dispClean*") {
                $packNode = $comp
                break
            }
        }
    }

    if (-not $packNode) {
        Log "No driver pack found for SKU '$sysId'."
        Log "Opening Dell support page..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    # 'path' attribute holds the relative URL path on downloads.dell.com
    $packPath = $packNode.GetAttribute("path")
    if (-not $packPath) {
        Log "Driver pack node missing 'path' attribute — unexpected catalog format."
        return $false
    }
    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Driver pack: $([System.IO.Path]::GetFileName($packPath))"
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
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# HP
# Strategy:
#   1. Try imagepal/ref/{id}/{id}.cab  (newer models)
#   2. Try imagepal/{id}/{id}.cab      (older models)
#   3. Parse XML for a SoftPaq entry whose Category contains "Driver Pack"
#   4. Extract SoftPaq number and build EXE URL via Get-HpSoftpaqUrl
# =========================
function Start-HpDriverInstall {
    param([string]$DriverRoot)

    Log "=== HP: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $platformId = $null
    try {
        $platformId = (Get-CimInstance Win32_BaseBoard).Product.Trim()
        Log "HP Platform ID: $platformId"
    } catch {
        Log "Could not read HP platform ID: $($_.Exception.Message)"
        return $false
    }
    if (-not $platformId) {
        Log "Empty HP platform ID."
        return $false
    }

    $catalogUrls = @(
        "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$platformId/$platformId.cab",
        "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/$platformId/$platformId.cab"
    )

    $packUrl   = $null
    $packSpNum = $null

    foreach ($catUrl in $catalogUrls) {
        Log "Trying HP catalog: $catUrl"
        $hpCab = Join-Path $env:TEMP "HP_${platformId}.cab"
        $hpXml = Join-Path $env:TEMP "HP_${platformId}.xml"
        Remove-Item $hpCab -EA SilentlyContinue
        Remove-Item $hpXml -EA SilentlyContinue

        if (-not (Invoke-CurlDownload -Url $catUrl -OutFile $hpCab)) {
            Log "  Catalog download failed — trying next."
            continue
        }

        Log "  Extracting HP catalog CAB..."
        $expandOut = & expand.exe "`"$hpCab`"" "`"$hpXml`"" 2>&1
        Log "  expand.exe: $expandOut"
        if (-not (Test-Path $hpXml) -or (Get-Item $hpXml).Length -eq 0) {
            Log "  CAB extraction failed — trying next."
            continue
        }

        try {
            $rawXml     = [System.IO.File]::ReadAllText($hpXml).TrimStart([char]0xFEFF)
            [xml]$hpCat = $rawXml
        } catch {
            Log "  XML parse failed: $($_.Exception.Message)"
            continue
        }

        # HP platform XML nodes can be <SoftPaq> or <UpdateInfo> depending on schema version.
        # Category text examples: "Driver Pack", "Manageability - Driver Pack"
        $foundNode = $null
        foreach ($node in $hpCat.SelectNodes(
            "//*[local-name()='SoftPaq' or local-name()='UpdateInfo']")) {
            $catText = ""
            try {
                $catText = $node.SelectSingleNode("*[local-name()='Category']").InnerText
            } catch {}
            if ($catText -notmatch "(?i)driver\s*pack") { continue }
            $foundNode = $node
            break
        }

        if (-not $foundNode) {
            Log "  No driver pack entry in this catalog — trying next."
            continue
        }

        # Extract the SoftPaq number from child elements or attributes
        $spNum = $null
        foreach ($field in @("SoftPaqNum","Id","Number","SPNumber")) {
            try {
                $val = $foundNode.SelectSingleNode("*[local-name()='$field']").InnerText.Trim()
                if ($val -match '(?i)^sp(\d+)$') { $spNum = "sp$($Matches[1])"; break }
                if ($val -match '^\d+$')          { $spNum = "sp$val";            break }
            } catch {}
        }
        if (-not $spNum) {
            # Try the node's own 'id' or 'number' attribute
            foreach ($attr in @("id","number","softpaqnum")) {
                $val = $foundNode.GetAttribute($attr)
                if ($val -match '(?i)^sp(\d+)$') { $spNum = "sp$($Matches[1])"; break }
                if ($val -match '^\d+$')          { $spNum = "sp$val";            break }
            }
        }

        if (-not $spNum) {
            Log "  Could not determine SoftPaq number."
            continue
        }

        $packSpNum = $spNum
        $packUrl   = Get-HpSoftpaqUrl -SpNum $packSpNum
        Log "  Found driver pack: $packSpNum"
        Log "  Download URL: $packUrl"
        break
    }

    if (Test-Cancelled) { return $false }

    if (-not $packUrl) {
        Log "All HP catalog sources exhausted for platform '$platformId'."
        Log "Opening HP Driver Pack Matrix..."
        Start-Process "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html"
        return $false
    }

    SetExtract  -Pct 40 -Label "Catalog OK — $packSpNum"
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
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300
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
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
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
    SetProgress 5

    $driverRoot = "C:\DRIVERS"
    $success    = $false

    if ($manufacturer -match "Dell") {
        if (-not (Assert-Curl)) { Set-ButtonIdle; return }
        $success = Start-DellDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "HP|Hewlett") {
        if (-not (Assert-Curl)) { Set-ButtonIdle; return }
        $success = Start-HpDriverInstall -DriverRoot $driverRoot
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