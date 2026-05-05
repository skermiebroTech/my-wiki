# =============================================================
# Install-Drivers-auto.ps1
# Version: 1.2.5
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode:
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
#
# Supports: Dell, HP, Lenovo
# Detects manufacturer, downloads correct driver pack silently,
# extracts to C:\DRIVERS, installs all INFs via pnputil.
# =============================================================

$ScriptVersion = "1.2.5"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Prevent sleep / display timeout during install
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
$form.Size            = New-Object System.Drawing.Size(580, 560)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 245)

# ---- Title label ----
$title           = New-Object System.Windows.Forms.Label
$title.AutoSize  = $true
$title.Font      = $FontTitleBold
$title.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$title.Location  = New-Object System.Drawing.Point(20, 15)
$title.Text      = "Driver Installer"
$title.UseCompatibleTextRendering = $false
$form.Controls.Add($title)

# ---- Version label (top-right) ----
$versionLabel           = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize  = $true
$versionLabel.Font      = $FontUISmall
$versionLabel.ForeColor = [System.Drawing.Color]::Gray
$versionLabel.Text      = "v$ScriptVersion"
$versionLabel.Location  = New-Object System.Drawing.Point(510, 20)
$versionLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($versionLabel)

# ---- Status box (dark terminal — Courier New for crispness) ----
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

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Size     = New-Object System.Drawing.Size(508, 18)
$progress.Location = New-Object System.Drawing.Point(12, 20)
$progress.Style    = "Continuous"
$progress.Minimum  = 0
$progress.Maximum  = 100
$overallGroupBox.Controls.Add($progress)

# ---- Log path label ----
$logLabel           = New-Object System.Windows.Forms.Label
$logLabel.AutoSize  = $false
$logLabel.Size      = New-Object System.Drawing.Size(536, 16)
$logLabel.Location  = New-Object System.Drawing.Point(20, 468)
$logLabel.ForeColor = [System.Drawing.Color]::Gray
$logLabel.Font      = $FontUISmall
$logLabel.Text      = "Log: $LogFile"
$logLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($logLabel)

# ---- Button ----
$button            = New-Object System.Windows.Forms.Button
$button.Text       = "Install Drivers"
$button.Size       = New-Object System.Drawing.Size(160, 36)
$button.Location   = New-Object System.Drawing.Point(200, 490)
$button.Font       = $FontUIBold
$button.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
$button.ForeColor  = [System.Drawing.Color]::White
$button.FlatStyle  = "Flat"
$button.FlatAppearance.BorderSize = 0
$form.Controls.Add($button)

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
    # Bar is Marquee — scrolls automatically. Snap to full Continuous on 100%.
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
    # Pct -ge 100  -> snap to filled Continuous bar (done)
    # Pct -eq -1   -> Marquee mode (active extraction, count in label)
    # Pct 0..99    -> Continuous proportional (INF install loop)
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

function Assert-Curl {
    if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
        Log "ERROR: curl.exe not found. Windows 10 1803+ required."
        return $false
    }
    return $true
}

# =========================
# PRE-SCAN: exact file count for ZIP via ZipFile API.
# EXE/CAB return 0 — extraction uses Marquee + live count label instead.
# =========================
function Get-PackFileCount {
    param([string]$PackFile)

    $ext = [System.IO.Path]::GetExtension($PackFile).ToLower()
    if ($ext -eq ".zip") {
        try {
            $zip   = [System.IO.Compression.ZipFile]::OpenRead($PackFile)
            $count = $zip.Entries.Count
            $zip.Dispose()
            Log "  Pre-scan: $count files in ZIP"
            return $count
        } catch {
            Log "  Pre-scan ZIP failed: $($_.Exception.Message)"
            return 0
        }
    } else {
        # EXE/CAB: no reliable pre-scan without extracting.
        # Return 0 so Marquee + live file count label is used.
        Log "  Pre-scan: EXE/CAB — live count mode"
        return 0
    }
}

# =========================
# CURL DOWNLOAD
# HEAD request first to get total size, then live MB/total/% bar
# =========================
function Invoke-CurlDownload {
    param(
        [string]$Url,
        [string]$OutFile,
        [int]$TimeoutSec = 600
    )

    $fileName = [System.IO.Path]::GetFileName($OutFile)
    Log "Downloading: $fileName"
    Log "  URL: $Url"
    SetDownload -Pct 0 -Label "Connecting..."

    # Silent HEAD to get Content-Length
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

    # Start download
    $psi                 = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "curl.exe"
    $psi.Arguments       = "--location --fail --max-time $TimeoutSec --connect-timeout 30 " +
                           "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64)`" " +
                           "--output `"$OutFile`" `"$Url`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc                = New-Object System.Diagnostics.Process
    $proc.StartInfo      = $psi
    $proc.Start() | Out-Null

    $lastSize = 0; $stall = 0
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 700
        $sz = if (Test-Path $OutFile) { (Get-Item $OutFile -EA SilentlyContinue).Length } else { 0 }
        if ($sz -gt $lastSize) { $stall = 0; $lastSize = $sz } else { $stall++ }
        $mbDone = [math]::Round($sz / 1MB, 1)

        if ($totalMB -gt 0) {
            $pct = [math]::Min([int](($sz / $totalBytes) * 100), 99)
            SetDownload -Pct $pct -Label "$mbDone MB / $totalMB MB  ($pct%)"
            Log "  $mbDone MB / $totalMB MB ($pct%)"
        } else {
            SetDownload -Pct 0 -Label "$mbDone MB received..."
            Log "  $mbDone MB received"
        }

        if ($stall -gt 90) {
            Log "  WARNING: download stalled 63s — aborting."
            $proc.Kill()
            SetDownload -Pct 0 -Label "Stalled — aborted."
            return $false
        }
    }

    if ($proc.ExitCode -ne 0) {
        Log "  curl failed (exit $($proc.ExitCode))"
        SetDownload -Pct 0 -Label "Failed (curl exit $($proc.ExitCode))"
        return $false
    }
    if (-not (Test-Path $OutFile) -or (Get-Item $OutFile).Length -eq 0) {
        Log "  File missing or empty after download."
        SetDownload -Pct 0 -Label "Failed — file empty."
        return $false
    }

    $finalMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 1)
    Log "  Download complete: $finalMB MB"
    SetDownload -Pct 100 -Label "Complete — $finalMB MB"
    return $true
}

# =========================
# EXTRACTION WATCHER
# Polls destination folder file count against a known/estimated total.
# Shows "X / Y files  (Z%)" so the bar is truly proportional.
# =========================
function Watch-Extraction {
    param(
        [System.Diagnostics.Process]$ExtractProc,
        [string]$DestPath,
        [int]$TotalFiles    = 0,   # exact for ZIP, 0 for EXE/CAB
        [int]$StallLimitSec = 300
    )

    $stall     = 0
    $lastCount = 0

    # -1 = Marquee mode (scrolling bar, count shown in label)
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

        [System.Windows.Forms.Application]::DoEvents()

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
    Log "  Extraction finished: $finalCount files in $DestPath"
}

# =========================
# INF INSTALLER
# Reuses Extract bar (relabelled) to show INF install progress
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
    $exGroupBox.Text = "Install INFs"
    # Switch from Marquee to Continuous so percentage values are visible
    $exBar.Style = "Continuous"
    $exBar.Value = 0

    foreach ($inf in $infs) {
        $i++
        $overallPct = 60 + [int](($i / $total) * 38)   # 60–98% of overall bar
        SetProgress $overallPct
        $infPct = [int](($i / $total) * 100)
        $remaining = $total - $i
        SetExtract -Pct $infPct -Label "$i / $total INFs  ($remaining remaining)  —  $($inf.Name)"
        Log "[$i/$total] $($inf.Name)"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        foreach ($l in $out) { Log "  $l" }
        [System.Windows.Forms.Application]::DoEvents()
    }

    SetProgress 100
    SetExtract -Pct 100 -Label "All $total INFs installed."
    $exGroupBox.Text = "Extract / Install"
    Log "All INFs processed."
    return $true
}

# =========================
# SHARED EXTRACT RUNNER
# Handles ZIP (background job), CAB (expand.exe), EXE (Inno Setup)
# Pre-scans for file count before starting so bar is proportional
# =========================
function Start-PackExtraction {
    param(
        [string]$PackFile,
        [string]$DestPath,
        [int]$StallLimitSec = 300
    )

    if (-not (Test-Path $DestPath)) { New-Item -Path $DestPath -ItemType Directory -Force | Out-Null }

    Log "Pre-scanning pack for file count..."
    SetExtract -Pct 0 -Label "Pre-scanning pack..."
    $totalFiles = Get-PackFileCount -PackFile $PackFile

    $ext = [System.IO.Path]::GetExtension($PackFile).ToLower()

    switch ($ext) {
        ".zip" {
            Log "Extracting ZIP in background..."
            SetExtract -Pct 0 -Label "Starting ZIP extraction..."
            $zipJob = Start-Job {
                param($src, $dst)
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::ExtractToDirectory($src, $dst)
            } -ArgumentList $PackFile, $DestPath

            $stall = 0; $lastCount = 0
            SetExtract -Pct -1 -Label "Starting ZIP extraction..."
            while ($zipJob.State -eq "Running") {
                Start-Sleep -Milliseconds 700
                $count = if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
                if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }

                if ($totalFiles -gt 0) {
                    $remaining = [math]::Max($totalFiles - $count, 0)
                    SetExtract -Pct -1 -Label "$count / $totalFiles files  ($remaining remaining)"
                } else {
                    SetExtract -Pct -1 -Label "$count files extracted..."
                }
                [System.Windows.Forms.Application]::DoEvents()
                if ($stall -gt 375) { Log "  ZIP stalled — stopping job."; Stop-Job $zipJob; break }
            }
            Receive-Job $zipJob -ErrorAction SilentlyContinue | Out-Null
            Remove-Job  $zipJob
            $finalCount = if (Test-Path $DestPath) {
                (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
            } else { 0 }
            SetExtract -Pct 100 -Label "Done — $finalCount files extracted"
            Log "  ZIP extraction complete. $finalCount files."
        }

        ".cab" {
            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = "expand.exe"
            $psi.Arguments       = "`"$PackFile`" -F:* `"$DestPath`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            $exProc              = New-Object System.Diagnostics.Process
            $exProc.StartInfo    = $psi
            $exProc.Start() | Out-Null
            Watch-Extraction -ExtractProc $exProc -DestPath $DestPath `
                             -TotalFiles $totalFiles -StallLimitSec $StallLimitSec
        }

        default {
            # Inno Setup EXE (Dell, HP, Lenovo SCCM packs)
            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = $PackFile
            $psi.Arguments       = "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES"
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            $exProc              = New-Object System.Diagnostics.Process
            $exProc.StartInfo    = $psi
            $exProc.Start() | Out-Null
            Watch-Extraction -ExtractProc $exProc -DestPath $DestPath `
                             -TotalFiles $totalFiles -StallLimitSec $StallLimitSec
        }
    }
}

# =========================
# DELL
# =========================
function Start-DellDriverInstall {
    param([string]$DriverRoot)

    Log "=== DELL: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $serviceTag = $null
    try {
        $serviceTag = (Get-CimInstance Win32_BIOS).SerialNumber.Trim()
        Log "Service Tag: $serviceTag"
    } catch {
        Log "Could not read Service Tag: $($_.Exception.Message)"; return $false
    }
    if (-not $serviceTag -or $serviceTag.Length -lt 4) { Log "Invalid Service Tag."; return $false }

    $sysId = $null
    try {
        $sysId = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber.Trim()
        Log "System SKU: $sysId"
    } catch { Log "Could not read SystemSKUNumber: $($_.Exception.Message)" }

    # Download catalog
    Log "Downloading Dell CatalogPC.cab..."
    $catalogCab = Join-Path $env:TEMP "DellCatalogPC.cab"
    $catalogXml = Join-Path $env:TEMP "CatalogPC.xml"
    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/CatalogPC.cab" -OutFile $catalogCab -TimeoutSec 180)) {
        Log "Failed to download Dell catalog."; return $false
    }
    SetProgress 15

    # Extract catalog
    SetExtract -Pct 10 -Label "Extracting Dell catalog..."
    Log "Extracting Dell catalog..."
    expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1 | Out-Null
    if (-not (Test-Path $catalogXml)) { Log "Dell catalog extraction failed."; return $false }
    SetExtract -Pct 40 -Label "Catalog extracted OK"
    SetProgress 20

    # Parse catalog
    Log "Parsing Dell catalog..."
    try {
        $rawXml   = [System.IO.File]::ReadAllText($catalogXml).TrimStart([char]0xFEFF)
        [xml]$cat = $rawXml
    } catch {
        Log "Failed to parse Dell catalog XML: $($_.Exception.Message)"; return $false
    }

    # Find driver pack matching SKU
    $packNode = $null
    foreach ($comp in $cat.SelectNodes("//SoftwareComponent")) {
        $dispName = ""
        try { $dispName = $comp.Name.Display.InnerText } catch {}
        if ($dispName -notmatch "(?i)driver\s*pack") { continue }
        if ($sysId) {
            $matched = $false
            foreach ($s in $comp.SelectNodes(".//SystemID")) {
                if ($s.InnerText.Trim() -ieq $sysId) { $matched = $true; break }
            }
            if (-not $matched) { continue }
        }
        $packNode = $comp; break
    }

    if (-not $packNode) {
        Log "No driver pack found for SKU '$sysId'. Opening Dell support page..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    $packPath = $packNode.GetAttribute("path")
    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Driver pack: $([System.IO.Path]::GetFileName($packPath))"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "Dell driver pack download failed."; return $false
    }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Dell_Extracted"
    Log "Extracting Dell pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 180
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# HP
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
        Log "Could not read HP platform ID: $($_.Exception.Message)"; return $false
    }
    if (-not $platformId) { Log "Empty HP platform ID."; return $false }

    # Download platform catalog
    Log "Fetching HP platform catalog for: $platformId"
    $hpCatalogUrl = "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$platformId/$platformId.cab"
    $hpCatalogCab = Join-Path $env:TEMP "HP_$platformId.cab"
    $hpCatalogXml = Join-Path $env:TEMP "HP_${platformId}.xml"

    if (-not (Invoke-CurlDownload -Url $hpCatalogUrl -OutFile $hpCatalogCab -TimeoutSec 120)) {
        Log "HP platform catalog not found. Opening HP matrix..."
        Start-Process "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html"
        return $false
    }
    SetProgress 15

    SetExtract -Pct 10 -Label "Extracting HP catalog..."
    expand.exe "`"$hpCatalogCab`"" "`"$hpCatalogXml`"" 2>&1 | Out-Null
    if (-not (Test-Path $hpCatalogXml)) { Log "HP catalog extraction failed."; return $false }
    SetExtract -Pct 40 -Label "Catalog extracted OK"
    SetProgress 20

    Log "Parsing HP catalog..."
    try {
        $rawXml     = [System.IO.File]::ReadAllText($hpCatalogXml).TrimStart([char]0xFEFF)
        [xml]$hpCat = $rawXml
    } catch {
        Log "Failed to parse HP catalog: $($_.Exception.Message)"; return $false
    }

    $packNode = $null
    foreach ($sp in $hpCat.SelectNodes("//SoftPaq")) {
        $catNode = $sp.SelectSingleNode("Category")
        if ($catNode -and $catNode.InnerText -match "(?i)driver\s*pack") { $packNode = $sp; break }
    }
    if (-not $packNode) {
        $packNode = $hpCat.SelectSingleNode(
            "//SoftPaq[contains(translate(Category,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'driver')]"
        )
    }
    if (-not $packNode) { Log "No HP driver pack found for '$platformId'."; return $false }

    $packUrl  = $packNode.SelectSingleNode("Url").InnerText.Trim()
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    Log "HP driver pack: $([System.IO.Path]::GetFileName($packFile))"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "HP driver pack download failed."; return $false
    }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "HP_Extracted"
    Log "Extracting HP SoftPaq..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 180
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
    } catch { Log "Could not read machine type: $($_.Exception.Message)" }
    if (-not $machineType) { Log "Cannot determine Lenovo machine type."; return $false }

    $winVer     = (Get-CimInstance Win32_OperatingSystem).Version
    $osAttr     = if ($winVer -match "^10\.0\.2") { "win11" } else { "win10" }
    $osFallback = if ($osAttr -eq "win11") { "win10" } else { "win11" }
    Log "Detected OS tag: $osAttr"

    Log "Fetching Lenovo catalogv2.xml..."
    $catalogFile = Join-Path $env:TEMP "lenovo_catalogv2.xml"
    if (-not (Invoke-CurlDownload -Url "https://download.lenovo.com/cdrt/td/catalogv2.xml" -OutFile $catalogFile -TimeoutSec 120)) {
        Log "Failed to download Lenovo catalog."; return $false
    }
    SetProgress 20

    Log "Parsing Lenovo catalog..."
    SetExtract -Pct 10 -Label "Parsing catalog..."
    try {
        $bytes   = [System.IO.File]::ReadAllBytes($catalogFile)
        $rawText = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        [xml]$cat = $rawText
        Log "Catalog parsed OK."
    } catch {
        Log "Failed to parse Lenovo catalog: $($_.Exception.Message)"; return $false
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

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "Lenovo driver pack download failed."; return $false
    }
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

    $button.Enabled = $false
    SetProgress 0
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."
    $exGroupBox.Text = "Extract"

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
        if (-not (Assert-Curl)) { $button.Enabled = $true; return }
        $success = Start-DellDriverInstall -DriverRoot $driverRoot
    } elseif ($manufacturer -match "HP|Hewlett") {
        if (-not (Assert-Curl)) { $button.Enabled = $true; return }
        $success = Start-HpDriverInstall -DriverRoot $driverRoot
    } elseif ($manufacturer -match "Lenovo") {
        if (-not (Assert-Curl)) { $button.Enabled = $true; return }
        $success = Start-LenovoDriverInstall -DriverRoot $driverRoot
    } else {
        Log "Unsupported manufacturer: $manufacturer"
        Log "Supported OEMs: Dell, HP, Lenovo"
        [System.Windows.Forms.MessageBox]::Show(
            "Manufacturer '$manufacturer' is not supported.`nSupported: Dell, HP, Lenovo",
            "Unsupported Manufacturer", "OK", "Warning"
        )
        $button.Enabled = $true
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
        else { $button.Enabled = $true }
    } else {
        SetDownload -Pct 0 -Label "Failed — see log"
        SetExtract  -Pct 0 -Label "Failed — see log"
        Log "Driver installation did not complete. Check log: $LogFile"
        [System.Windows.Forms.MessageBox]::Show(
            "Driver installation failed or no pack was found.`nCheck the log:`n`n$LogFile",
            "Installation Failed", "OK", "Error"
        )
        $button.Enabled = $true
    }
}

# =========================
# WIRE UP + LAUNCH
# =========================
$button.Add_Click({ Start-Install })

$form.Add_Shown({
    $form.Activate()
    Start-Sleep -Milliseconds 300
    Log "Running startup checks..."
    Start-Install
})

[void]$form.ShowDialog()