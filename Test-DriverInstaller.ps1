#Requires -RunAsAdministrator
# =============================================================================
# Test-DriverInstaller.ps1  v2.4.0
# =============================================================================

param(
    [string[]]$OEM             = @("Dell", "HP", "Lenovo"),
    [switch]$RunInstall,
    [string]$Branch            = "main",
    [string]$DellModel         = "Latitude 7320",
    [string]$HPModel           = "HP EliteBook x360 1030 G8 Notebook PC",
    [string]$LenovoMachineType = "20XX"
)

$ScriptVersion = "2.4.0"
$RepoUrl       = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/$Branch/Install-Drivers-auto-7z.ps1"
$LogFile       = Join-Path ([Environment]::GetFolderPath("UserProfile")) `
                     ("Downloads\Test-DriverInstaller_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
$ScriptCache   = "$env:TEMP\Install-Drivers-auto-test.ps1"

$OEMConfig = @{
    Dell   = @{ Manufacturer = "Dell Inc."; Model = $DellModel;  MachineType = "" }
    HP     = @{ Manufacturer = "HP";        Model = $HPModel;    MachineType = "" }
    Lenovo = @{ Manufacturer = "LENOVO";    Model = "";          MachineType = $LenovoMachineType }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================================================================
# PALETTE + FONTS
# =============================================================================
$C = @{
    Bg        = [System.Drawing.Color]::FromArgb(10,  12,  18)
    Panel     = [System.Drawing.Color]::FromArgb(18,  22,  32)
    PanelBord = [System.Drawing.Color]::FromArgb(35,  45,  65)
    TextDim   = [System.Drawing.Color]::FromArgb(80,  95,  120)
    TextMid   = [System.Drawing.Color]::FromArgb(140, 160, 190)
    TextBrt   = [System.Drawing.Color]::FromArgb(215, 225, 240)
    White     = [System.Drawing.Color]::FromArgb(240, 245, 255)
    Accent    = [System.Drawing.Color]::FromArgb(50,  140, 255)
    AccentDim = [System.Drawing.Color]::FromArgb(25,  70,  140)
    Green     = [System.Drawing.Color]::FromArgb(50,  210, 120)
    GreenDim  = [System.Drawing.Color]::FromArgb(15,  45,  28)
    Red       = [System.Drawing.Color]::FromArgb(255, 80,  80)
    RedDim    = [System.Drawing.Color]::FromArgb(55,  15,  15)
    RedBtn    = [System.Drawing.Color]::FromArgb(160, 30,  30)
    RedBtnHot = [System.Drawing.Color]::FromArgb(200, 50,  50)
    Yellow    = [System.Drawing.Color]::FromArgb(255, 200, 60)
    YellowDim = [System.Drawing.Color]::FromArgb(55,  45,  10)
    LogBg     = [System.Drawing.Color]::FromArgb(8,   10,  15)
    DellLog   = [System.Drawing.Color]::FromArgb(80,  180, 255)
    HPLog     = [System.Drawing.Color]::FromArgb(100, 220, 140)
    LenovoLog = [System.Drawing.Color]::FromArgb(255, 190, 60)
}

$F = @{
    Title   = New-Object System.Drawing.Font("Consolas", 13, [System.Drawing.FontStyle]::Bold)
    Sub     = New-Object System.Drawing.Font("Consolas", 8,  [System.Drawing.FontStyle]::Regular)
    SubBold = New-Object System.Drawing.Font("Consolas", 8,  [System.Drawing.FontStyle]::Bold)
    OEM     = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    Stat    = New-Object System.Drawing.Font("Consolas", 18, [System.Drawing.FontStyle]::Bold)
    StatSm  = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    Log     = New-Object System.Drawing.Font("Consolas", 7,  [System.Drawing.FontStyle]::Regular)
    Badge   = New-Object System.Drawing.Font("Consolas", 9,  [System.Drawing.FontStyle]::Bold)
    Kill    = New-Object System.Drawing.Font("Consolas", 9,  [System.Drawing.FontStyle]::Bold)
}

# =============================================================================
# LAYOUT CONSTANTS
# =============================================================================
$W          = 920
$H          = 720
$GAP        = 12
$HEADER_H   = 62
$PANEL_H    = 170
$PANEL_W    = [int](($W - ($GAP * 4)) / 3)   # 3 equal columns
$PANEL_Y    = $HEADER_H + $GAP
$SUMMARY_H  = 34
$SUMMARY_Y  = $PANEL_Y + $PANEL_H + $GAP
$LOG_CAP_H  = 18
$LOG_Y      = $SUMMARY_Y + $SUMMARY_H + $GAP
$LOG_H      = $H - $LOG_Y - $GAP
$OEMOrder   = @("Dell", "HP", "Lenovo")
$LogColors  = @{ Dell = $C.DellLog; HP = $C.HPLog; Lenovo = $C.LenovoLog }

# =============================================================================
# FORM
# =============================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Driver Installer Test Suite  v$ScriptVersion"
$form.ClientSize      = New-Object System.Drawing.Size($W, $H)
$form.MinimumSize     = New-Object System.Drawing.Size($W, 600)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $C.Bg
$form.ForeColor       = $C.TextBrt
$form.FormBorderStyle = "Sizable"

# =============================================================================
# HEADER
# =============================================================================
$header           = New-Object System.Windows.Forms.Panel
$header.Location  = New-Object System.Drawing.Point(0, 0)
$header.Size      = New-Object System.Drawing.Size($W, $HEADER_H)
$header.BackColor = $C.Panel
$form.Controls.Add($header)

$accentBar           = New-Object System.Windows.Forms.Panel
$accentBar.Location  = New-Object System.Drawing.Point(0, $HEADER_H)
$accentBar.Size      = New-Object System.Drawing.Size($W, 2)
$accentBar.BackColor = $C.Accent
$form.Controls.Add($accentBar)

$titleLbl           = New-Object System.Windows.Forms.Label
$titleLbl.Text      = "DRIVER INSTALLER TEST SUITE"
$titleLbl.Font      = $F.Title
$titleLbl.ForeColor = $C.White
$titleLbl.AutoSize  = $true
$titleLbl.Location  = New-Object System.Drawing.Point(16, 10)
$titleLbl.BackColor = [System.Drawing.Color]::Transparent
$header.Controls.Add($titleLbl)

$metaLbl           = New-Object System.Windows.Forms.Label
$metaLbl.Text      = "v$ScriptVersion  |  branch: $Branch  |  $(if ($RunInstall) {'FULL INSTALL'} else {'EXTRACT ONLY'})"
$metaLbl.Font      = $F.Sub
$metaLbl.ForeColor = $C.TextDim
$metaLbl.AutoSize  = $true
$metaLbl.Location  = New-Object System.Drawing.Point(18, 38)
$metaLbl.BackColor = [System.Drawing.Color]::Transparent
$header.Controls.Add($metaLbl)

# Kill button — top right of header, always visible
$killBtn                           = New-Object System.Windows.Forms.Button
$killBtn.Text                      = "KILL ALL JOBS"
$killBtn.Font                      = $F.Kill
$killBtn.ForeColor                 = $C.White
$killBtn.BackColor                 = $C.RedBtn
$killBtn.FlatStyle                 = "Flat"
$killBtn.FlatAppearance.BorderSize = 0
$killBtn.Size                      = New-Object System.Drawing.Size(130, 30)
$killBtn.Location                  = New-Object System.Drawing.Point(($W - 146), 16)
$killBtn.Cursor                    = [System.Windows.Forms.Cursors]::Hand
$header.Controls.Add($killBtn)

$elapsedLbl           = New-Object System.Windows.Forms.Label
$elapsedLbl.Text      = "00:00"
$elapsedLbl.Font      = $F.StatSm
$elapsedLbl.ForeColor = $C.TextDim
$elapsedLbl.AutoSize  = $true
$elapsedLbl.Location  = New-Object System.Drawing.Point(($W - 146), 50)
$elapsedLbl.BackColor = [System.Drawing.Color]::Transparent
$header.Controls.Add($elapsedLbl)

# =============================================================================
# OEM STATUS PANELS
# =============================================================================
$oemPanels = @{}

for ($pi = 0; $pi -lt 3; $pi++) {
    $oemName = $OEMOrder[$pi]
    $px      = $GAP + $pi * ($PANEL_W + $GAP)
    $logCol  = $LogColors[$oemName]

    $p           = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($px, $PANEL_Y)
    $p.Size      = New-Object System.Drawing.Size($PANEL_W, $PANEL_H)
    $p.BackColor = $C.Panel
    $form.Controls.Add($p)

    # Top colour strip — matches the log colour
    $strip           = New-Object System.Windows.Forms.Panel
    $strip.Location  = New-Object System.Drawing.Point(0, 0)
    $strip.Size      = New-Object System.Drawing.Size($PANEL_W, 3)
    $strip.BackColor = $C.AccentDim
    $p.Controls.Add($strip)

    # Status badge
    $badge           = New-Object System.Windows.Forms.Label
    $badge.Text      = "IDLE"
    $badge.Font      = $F.Badge
    $badge.ForeColor = $C.TextDim
    $badge.BackColor = [System.Drawing.Color]::Transparent
    $badge.AutoSize  = $true
    $badge.Location  = New-Object System.Drawing.Point(12, 10)
    $p.Controls.Add($badge)

    # OEM name
    $nameL           = New-Object System.Windows.Forms.Label
    $nameL.Text      = $oemName.ToUpper()
    $nameL.Font      = $F.OEM
    $nameL.ForeColor = $C.TextBrt
    $nameL.BackColor = [System.Drawing.Color]::Transparent
    $nameL.AutoSize  = $true
    $nameL.Location  = New-Object System.Drawing.Point(12, 28)
    $p.Controls.Add($nameL)

    # Model sub-label
    $cfg = $OEMConfig[$oemName]
    $mt  = if ($cfg.Model) { $cfg.Model } elseif ($cfg.MachineType) { "Type: $($cfg.MachineType)" } else { "WMI detection" }
    if ($mt.Length -gt 34) { $mt = $mt.Substring(0, 31) + "..." }
    $modelL           = New-Object System.Windows.Forms.Label
    $modelL.Text      = $mt
    $modelL.Font      = $F.Sub
    $modelL.ForeColor = $C.TextDim
    $modelL.BackColor = [System.Drawing.Color]::Transparent
    $modelL.AutoSize  = $false
    $modelL.Size      = New-Object System.Drawing.Size(($PANEL_W - 24), 14)
    $modelL.Location  = New-Object System.Drawing.Point(12, 50)
    $p.Controls.Add($modelL)

    # Divider
    $div           = New-Object System.Windows.Forms.Panel
    $div.Location  = New-Object System.Drawing.Point(12, 68)
    $div.Size      = New-Object System.Drawing.Size(($PANEL_W - 24), 1)
    $div.BackColor = $C.PanelBord
    $p.Controls.Add($div)

    # Stats row
    $statDefs = @(
        @{ Key = "Files"; X = 10  }
        @{ Key = "INFs";  X = 100 }
        @{ Key = "Sec";   X = 190 }
    )
    $statLabels = @{}
    foreach ($sd in $statDefs) {
        $nL           = New-Object System.Windows.Forms.Label
        $nL.Text      = "--"
        $nL.Font      = $F.Stat
        $nL.ForeColor = $C.TextDim
        $nL.BackColor = [System.Drawing.Color]::Transparent
        $nL.AutoSize  = $true
        $nL.Location  = New-Object System.Drawing.Point($sd.X, 74)
        $p.Controls.Add($nL)

        $cL           = New-Object System.Windows.Forms.Label
        $cL.Text      = $sd.Key
        $cL.Font      = $F.Sub
        $cL.ForeColor = $C.TextDim
        $cL.BackColor = [System.Drawing.Color]::Transparent
        $cL.AutoSize  = $true
        $cL.Location  = New-Object System.Drawing.Point(($sd.X + 2), 106)
        $p.Controls.Add($cL)
        $statLabels[$sd.Key] = $nL
    }

    # Notes label
    $notesL           = New-Object System.Windows.Forms.Label
    $notesL.Text      = ""
    $notesL.Font      = $F.Sub
    $notesL.ForeColor = $C.TextDim
    $notesL.BackColor = [System.Drawing.Color]::Transparent
    $notesL.AutoSize  = $false
    $notesL.Size      = New-Object System.Drawing.Size(($PANEL_W - 24), 14)
    $notesL.Location  = New-Object System.Drawing.Point(12, 122)
    $p.Controls.Add($notesL)

    # Progress bar
    $barY = $PANEL_H - 5
    $bar                       = New-Object System.Windows.Forms.ProgressBar
    $bar.Style                 = "Marquee"
    $bar.MarqueeAnimationSpeed = 0
    $bar.Location              = New-Object System.Drawing.Point(0, $barY)
    $bar.Size                  = New-Object System.Drawing.Size($PANEL_W, 5)
    $bar.Minimum               = 0
    $bar.Maximum               = 100
    $p.Controls.Add($bar)

    $oemPanels[$oemName] = @{
        Panel    = $p
        Strip    = $strip
        Badge    = $badge
        Stats    = $statLabels
        Bar      = $bar
        NotesLbl = $notesL
        LogColor = $logCol
    }
}

# =============================================================================
# SUMMARY BAR
# =============================================================================
$summaryPanel           = New-Object System.Windows.Forms.Panel
$summaryPanel.Location  = New-Object System.Drawing.Point($GAP, $SUMMARY_Y)
$summaryPanel.Size      = New-Object System.Drawing.Size(($W - $GAP * 2), $SUMMARY_H)
$summaryPanel.BackColor = $C.Panel
$form.Controls.Add($summaryPanel)

$summaryLbl           = New-Object System.Windows.Forms.Label
$summaryLbl.Text      = "Initialising..."
$summaryLbl.Font      = $F.SubBold
$summaryLbl.ForeColor = $C.TextMid
$summaryLbl.BackColor = [System.Drawing.Color]::Transparent
$summaryLbl.AutoSize  = $false
$summaryLbl.Size      = New-Object System.Drawing.Size(($W - $GAP * 2 - 24), $SUMMARY_H)
$summaryLbl.Location  = New-Object System.Drawing.Point(12, 0)
$summaryLbl.TextAlign = "MiddleLeft"
$summaryPanel.Controls.Add($summaryLbl)

# =============================================================================
# THREE LOG BOXES — one per OEM
# =============================================================================
$oemLogs = @{}

for ($pi = 0; $pi -lt 3; $pi++) {
    $oemName = $OEMOrder[$pi]
    $px      = $GAP + $pi * ($PANEL_W + $GAP)
    $logCol  = $LogColors[$oemName]

    # Caption label
    $cap           = New-Object System.Windows.Forms.Label
    $cap.Text      = $oemName.ToUpper() + " LOG"
    $cap.Font      = $F.SubBold
    $cap.ForeColor = $logCol
    $cap.BackColor = [System.Drawing.Color]::Transparent
    $cap.AutoSize  = $true
    $cap.Location  = New-Object System.Drawing.Point($px, $LOG_Y)
    $form.Controls.Add($cap)

    # Log box
    $logBoxH = $LOG_H - $LOG_CAP_H - 2
    $lb             = New-Object System.Windows.Forms.RichTextBox
    $lb.Multiline   = $true
    $lb.ScrollBars  = "Vertical"
    $lb.ReadOnly    = $true
    $lb.BackColor   = $C.LogBg
    $lb.ForeColor   = $logCol
    $lb.Font        = $F.Log
    $lb.BorderStyle = "None"
    $lb.Location    = New-Object System.Drawing.Point($px, ($LOG_Y + $LOG_CAP_H + 2))
    $lb.Size        = New-Object System.Drawing.Size($PANEL_W, $logBoxH)
    $form.Controls.Add($lb)

    $oemLogs[$oemName] = @{ Box = $lb; Caption = $cap }
}

# =============================================================================
# RESIZE HANDLER
# =============================================================================
$form.Add_Resize({
    $cw = $form.ClientSize.Width
    $ch = $form.ClientSize.Height
    $newPW = [int](($cw - ($GAP * 4)) / 3)
    $header.Width  = $cw
    $accentBar.Width = $cw
    $killBtn.Location = New-Object System.Drawing.Point(($cw - 146), 16)
    $elapsedLbl.Location = New-Object System.Drawing.Point(($cw - 146), 50)
    $summaryPanel.Width = $cw - $GAP * 2
    $summaryLbl.Width   = $cw - $GAP * 2 - 24
    $newLogH = $ch - $LOG_Y - $GAP
    $newLogBoxH = $newLogH - $LOG_CAP_H - 2
    for ($i = 0; $i -lt 3; $i++) {
        $n   = $OEMOrder[$i]
        $npx = $GAP + $i * ($newPW + $GAP)
        $oemPanels[$n].Panel.Location = New-Object System.Drawing.Point($npx, $PANEL_Y)
        $oemPanels[$n].Panel.Width    = $newPW
        $oemPanels[$n].Bar.Width      = $newPW
        $oemPanels[$n].Strip.Width    = $newPW
        $oemPanels[$n].NotesLbl.Width = $newPW - 24
        $oemLogs[$n].Caption.Location = New-Object System.Drawing.Point($npx, $LOG_Y)
        $oemLogs[$n].Box.Location     = New-Object System.Drawing.Point($npx, ($LOG_Y + $LOG_CAP_H + 2))
        $oemLogs[$n].Box.Width        = $newPW
        $oemLogs[$n].Box.Height       = $newLogBoxH
    }
    [System.Windows.Forms.Application]::DoEvents()
})

# =============================================================================
# HELPERS
# =============================================================================
function UILog {
    param([string]$msg, [string]$OEM = "")
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $msg"
    # Write to per-OEM log box if specified, else to all
    if ($OEM -and $oemLogs.ContainsKey($OEM)) {
        $oemLogs[$OEM].Box.AppendText("$line`r`n")
        $oemLogs[$OEM].Box.ScrollToCaret()
    } else {
        foreach ($n in $OEMOrder) {
            $oemLogs[$n].Box.AppendText("$line`r`n")
            $oemLogs[$n].Box.ScrollToCaret()
        }
    }
    Add-Content -Path $LogFile -Value $(if ($OEM) {"[$OEM] $line"} else {$line}) -Encoding UTF8
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-OEMStatus {
    param([string]$OEM, [string]$Status)
    $pnl = $oemPanels[$OEM]
    switch ($Status) {
        "RUNNING" {
            $pnl.Badge.Text              = "RUNNING"
            $pnl.Badge.ForeColor         = $C.Yellow
            $pnl.Strip.BackColor         = $C.Yellow
            $pnl.Panel.BackColor         = $C.YellowDim
            $pnl.Bar.MarqueeAnimationSpeed = 25
            $pnl.Bar.Style               = "Marquee"
        }
        "PASS" {
            $pnl.Badge.Text      = "PASS"
            $pnl.Badge.ForeColor = $C.Green
            $pnl.Strip.BackColor = $C.Green
            $pnl.Panel.BackColor = $C.GreenDim
            $pnl.Bar.Style       = "Continuous"
            $pnl.Bar.Value       = 100
        }
        "FAIL" {
            $pnl.Badge.Text      = "FAIL"
            $pnl.Badge.ForeColor = $C.Red
            $pnl.Strip.BackColor = $C.Red
            $pnl.Panel.BackColor = $C.RedDim
            $pnl.Bar.Style       = "Continuous"
            $pnl.Bar.Value       = 100
        }
        "SKIP" {
            $pnl.Badge.Text      = "SKIP"
            $pnl.Badge.ForeColor = $C.TextDim
            $pnl.Strip.BackColor = $C.PanelBord
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-OEMStats {
    param([string]$OEM, [int]$Files, [int]$INFs, [double]$Sec, [string]$Notes = "")
    $pnl = $oemPanels[$OEM]
    $pnl.Stats["Files"].Text      = "$Files"
    $pnl.Stats["INFs"].Text       = "$INFs"
    $pnl.Stats["Sec"].Text        = "$([int]$Sec)"
    $pnl.Stats["Files"].ForeColor = if ($Files -gt 0) { $C.White } else { $C.Red }
    $pnl.Stats["INFs"].ForeColor  = if ($INFs  -gt 0) { $C.White } else { $C.Red }
    $pnl.Stats["Sec"].ForeColor   = $C.TextMid
    if ($Notes) {
        $short = if ($Notes.Length -gt 36) { $Notes.Substring(0, 33) + "..." } else { $Notes }
        $pnl.NotesLbl.Text      = $short
        $pnl.NotesLbl.ForeColor = $C.Red
    }
    [System.Windows.Forms.Application]::DoEvents()
}

foreach ($n in $OEMOrder) { if ($OEM -notcontains $n) { Set-OEMStatus -OEM $n -Status "SKIP" } }

# =============================================================================
# TIMER
# =============================================================================
$globalStart          = Get-Date
$ticker               = New-Object System.Windows.Forms.Timer
$ticker.Interval      = 500
$ticker.Add_Tick({
    $e = [int]((Get-Date) - $globalStart).TotalSeconds
    $elapsedLbl.Text = "{0:D2}:{1:D2}" -f ([int]($e / 60)), ($e % 60)
    [System.Windows.Forms.Application]::DoEvents()
})
$ticker.Start()

# =============================================================================
# KILL BUTTON
# =============================================================================
$script:KillRequested = $false
$script:ActiveJobs    = $null

$killBtn.Add_Click({
    if ($script:KillRequested) { return }
    $script:KillRequested      = $true
    $killBtn.Text              = "KILLING..."
    $killBtn.BackColor         = [System.Drawing.Color]::FromArgb(80, 10, 10)
    $killBtn.Enabled           = $false
    $summaryLbl.Text           = "Kill requested - stopping all jobs..."
    $summaryLbl.ForeColor      = $C.Red
    UILog "=== KILL REQUESTED BY USER ==="

    # Hard kill — read PID files written by each job and terminate immediately
    foreach ($n in $OEMOrder) {
        $pidFile = "$env:TEMP\driver_test_pid_$n.txt"
        if (Test-Path $pidFile) {
            try {
                $childPid = [int](Get-Content $pidFile -Raw)
                $proc = Get-Process -Id $childPid -EA SilentlyContinue
                if ($proc) {
                    $proc.Kill()
                    UILog "[$n] Hard killed PID $childPid" -OEM $n
                }
                Remove-Item $pidFile -Force -EA SilentlyContinue
            } catch { UILog "[$n] Kill error: $_" -OEM $n }
        }
    }

    # Also stop the PowerShell job wrappers
    if ($script:ActiveJobs) {
        foreach ($e in $script:ActiveJobs) {
            try {
                Stop-Job   $e.Job -EA SilentlyContinue
                Remove-Job $e.Job -Force -EA SilentlyContinue
                UILog "[$($e.OEM)] Job wrapper stopped" -OEM $e.OEM
                Set-OEMStatus -OEM $e.OEM -Status "FAIL"
            } catch {}
        }
    }

    foreach ($n in $OEMOrder) {
        $dr = "C:\DRIVERS\$n"
        if (Test-Path $dr) { Remove-Item $dr -Recurse -Force -EA SilentlyContinue }
    }
    Remove-Item $ScriptCache -Force -EA SilentlyContinue
    $ticker.Stop()
    $summaryLbl.Text = "All jobs killed. Cleaned up."
    UILog "=== KILL COMPLETE ==="
    $killBtn.Text      = "KILLED"
    $killBtn.BackColor = [System.Drawing.Color]::FromArgb(40, 10, 10)
    [System.Windows.Forms.Application]::DoEvents()
})

# =============================================================================
# MAIN LOGIC
# =============================================================================
$form.Add_Shown({
    $form.Activate()
    [System.Windows.Forms.Application]::DoEvents()
    New-Item -ItemType File -Path $LogFile -Force | Out-Null

    # -- Download script ------------------------------------------------------
    $summaryLbl.Text      = "Downloading Install-Drivers-auto-7z.ps1..."
    $summaryLbl.ForeColor = $C.Accent
    UILog "Repo: $RepoUrl"

    $cx = (Start-Process curl.exe -ArgumentList "--silent --location --output `"$ScriptCache`" `"$RepoUrl`"" -Wait -PassThru).ExitCode
    if ($cx -ne 0 -or -not (Test-Path $ScriptCache)) {
        $summaryLbl.Text      = "ERROR: Could not download script (curl exit $cx)"
        $summaryLbl.ForeColor = $C.Red
        UILog "FAIL: Download failed"
        return
    }
    $kb  = [math]::Round((Get-Item $ScriptCache).Length / 1KB, 1)
    $vl  = Get-Content $ScriptCache | Select-String '^\$ScriptVersion\s*=' | Select-Object -First 1
    $ver = if ($vl) { ($vl.ToString() -replace '.*"(.*)".*', '$1') } else { "?" }
    UILog "Downloaded $kb KB  |  Script v$ver"
    $summaryLbl.Text      = "Script v$ver ready. Launching parallel jobs..."
    $summaryLbl.ForeColor = $C.TextBrt

    # -- Launch jobs ----------------------------------------------------------
    $results = [System.Collections.Generic.List[object]]::new()
    $jobs    = [System.Collections.Generic.List[object]]::new()

    foreach ($oemName in $OEM) {
        if (-not $OEMConfig.ContainsKey($oemName)) {
            UILog "Unknown OEM '$oemName' - skipping"
            continue
        }

        $cfg        = $OEMConfig[$oemName]
        $driverRoot = "C:\DRIVERS\$oemName"
        $extractDir = "$driverRoot\$($oemName)_Extracted"

        $al  = "-ExecutionPolicy Bypass -File `"$ScriptCache`""
        $al += " -Manufacturer `"$($cfg.Manufacturer)`""
        if ($cfg.Model)       { $al += " -Model `"$($cfg.Model)`"" }
        if ($cfg.MachineType) { $al += " -MachineType `"$($cfg.MachineType)`"" }
        $al += " -DriverRoot `"$driverRoot`" -SkipCleanup"
        if (-not $RunInstall) { $al += " -SkipInstall" }

        Set-OEMStatus -OEM $oemName -Status "RUNNING"
        UILog "[$oemName] Launching job..." -OEM $oemName

        $jAl = $al; $jN = $oemName; $jDr = $driverRoot
        $jEd = $extractDir; $jLf = $LogFile; $jRi = $RunInstall

        $job = Start-Job -ScriptBlock {
            param($al, $n, $dr, $ed, $lf, $ri)
            function JL { param([string]$m)
                Add-Content $lf "[$([datetime]::Now.ToString('HH:mm:ss'))][$n] $m" -Encoding UTF8
                $m
            }
            $st  = Get-Date
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName               = "powershell.exe"
            $psi.Arguments              = $al
            $psi.UseShellExecute        = $false
            $psi.CreateNoWindow         = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $psi
            $proc.Start() | Out-Null
            # Write PID to temp file so kill button can terminate immediately
            $pidFile = "$env:TEMP\driver_test_pid_$n.txt"
            $proc.Id | Set-Content $pidFile -Encoding UTF8
            $out = [System.Collections.Generic.List[string]]::new()

            # Use a concurrent queue + background runspace so ReadLine() never
            # blocks the main loop. The reader runspace pushes lines into the queue;
            # the main loop drains it every 700ms and writes to the pipeline so
            # Receive-Job -Keep in the UI picks them up live.
            $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
            $readerDone = [System.Threading.ManualResetEventSlim]::new($false)

            $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $rs.Open()
            $rs.SessionStateProxy.SetVariable("proc",      $proc)
            $rs.SessionStateProxy.SetVariable("queue",     $queue)
            $rs.SessionStateProxy.SetVariable("readerDone",$readerDone)

            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $rs
            $ps.AddScript({
                try {
                    while (-not $proc.HasExited) {
                        if ($proc.StandardOutput.Peek() -ge 0) {
                            $l = $proc.StandardOutput.ReadLine()
                            if ($null -ne $l) { $queue.Enqueue($l) }
                        } else { Start-Sleep -Milliseconds 30 }
                    }
                    # Drain remainder
                    $tail = $proc.StandardOutput.ReadToEnd()
                    if ($tail) {
                        foreach ($tl in ($tail -split "`n")) {
                            $tl = $tl.TrimEnd("`r")
                            if ($tl) { $queue.Enqueue($tl) }
                        }
                    }
                } finally { $readerDone.Set() }
            }) | Out-Null
            $ps.BeginInvoke() | Out-Null

            # Main drain loop — every 700ms flush queue to pipeline
            $lastActivity = Get-Date
            $lastFileSeen  = ""
            while (-not $readerDone.IsSet -or -not $queue.IsEmpty) {
                Start-Sleep -Milliseconds 700
                $flushed = 0
                $l = $null
                while ($queue.TryDequeue([ref]$l)) {
                    $out.Add($l)
                    JL $l | Out-Null
                    Write-Output $l     # hits pipeline immediately
                    $flushed++
                    $lastActivity = Get-Date
                }
                # If nothing printed in last 700ms emit a heartbeat so UI knows we are alive
                if ($flushed -eq 0 -and -not $readerDone.IsSet) {
                    $elapsed = [int]((Get-Date) - $st).TotalSeconds
                    $fc = if (Test-Path $ed) { @(Get-ChildItem $ed -Recurse -File -EA SilentlyContinue).Count } else { 0 }
                    $pulse = "  ... [$n] running ${elapsed}s$(if ($fc -gt 0){" | $fc files extracted"})"
                    Write-Output $pulse
                }
            }

            $ps.Dispose(); $rs.Close(); $rs.Dispose()

            $err = $proc.StandardError.ReadToEnd().Trim()
            if ($err) {
                $errLine = "ERR: $err"
                $out.Add($errLine)
                JL $errLine | Out-Null
                Write-Output $errLine
            }
            Remove-Item $pidFile -Force -EA SilentlyContinue

            $exit = $proc.ExitCode
            $dur  = [math]::Round(((Get-Date) - $st).TotalSeconds, 1)
            $ok   = $true
            $notes = [System.Collections.Generic.List[string]]::new()
            $fc = 0; $ic = 0

            if ($exit -ne 0) { $notes.Add("Exit:$exit"); $ok = $false }
            if (Test-Path $ed) {
                $fc = @(Get-ChildItem $ed -Recurse -File -EA SilentlyContinue).Count
                $ic = @(Get-ChildItem $ed -Recurse -Filter "*.inf" -EA SilentlyContinue).Count
                if ($fc -eq 0) { $notes.Add("0 files");    $ok = $false }
                if ($ic -eq 0) { $notes.Add("0 INFs");     $ok = $false }
            } else { $notes.Add("No extract dir"); $ok = $false }

            if ($n -in @("Dell","HP") -and (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
                $notes.Add("7z not removed"); $ok = $false
            }
            $sl = $out | Where-Object { $_ -match "SUCCESS:|complete" } | Select-Object -First 1
            $fl = $out | Where-Object { $_ -match "FAILED:|did not"  } | Select-Object -First 1
            if (-not $sl -and $fl)         { $notes.Add("Script fail");     $ok = $false }
            if (-not $sl -and -not $fl)    { $notes.Add("Unclear outcome") }

            if (Test-Path $dr) { Remove-Item $dr -Recurse -Force -EA SilentlyContinue }
            JL "END $n $(if ($ok) {'PASS'} else {'FAIL'})" | Out-Null

            [ordered]@{
                OEM      = $n
                Status   = if ($ok) { "PASS" } else { "FAIL" }
                Duration = $dur
                Files    = $fc
                INFs     = $ic
                Notes    = ($notes -join "; ")
                Output   = ($out -join "`n")
            }
        } -ArgumentList $jAl, $jN, $jDr, $jEd, $jLf, $jRi

        $jobs.Add([ordered]@{ OEM = $oemName; Job = $job; Start = (Get-Date) })
        $script:ActiveJobs = $jobs
        UILog "[$oemName] Job $($job.Id) started" -OEM $oemName
    }

    # -- Wait loop with live log streaming ------------------------------------
    $done     = @{}
    $lastLine = @{}
    foreach ($n in $OEMOrder) { $lastLine[$n] = 0 }

    while ($done.Count -lt $jobs.Count) {
        foreach ($e in $jobs) {
            $n = $e.OEM; $j = $e.Job
            $sec = [int]((Get-Date) - $e.Start).TotalSeconds
            $oemPanels[$n].Stats["Sec"].Text = "$sec"

            # Stream all new output lines from the job unfiltered
            $newLines = Receive-Job $j -Keep -EA SilentlyContinue
            if ($newLines) {
                $allLines = @($newLines | Where-Object { $_ -is [string] })
                $start    = $lastLine[$n]
                for ($li = $start; $li -lt $allLines.Count; $li++) {
                    $line = $allLines[$li].ToString()
                    $oemLogs[$n].Box.AppendText("$line`r`n")
                    $oemLogs[$n].Box.ScrollToCaret()
                }
                $lastLine[$n] = $allLines.Count
            }

            if (-not $done.ContainsKey($n) -and $j.State -in @("Completed","Failed","Stopped")) {
                $done[$n] = $true
                UILog "[$n] Job finished in ${sec}s" -OEM $n
            }
        }

        $still = ($jobs | Where-Object { -not $done.ContainsKey($_.OEM) } | ForEach-Object { $_.OEM }) -join ", "
        if ($still) {
            $summaryLbl.Text      = "Running: $still"
            $summaryLbl.ForeColor = $C.Yellow
        }
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:KillRequested) { break }
        if ($done.Count -lt $jobs.Count) { Start-Sleep -Milliseconds 600 }
    }

    # -- Collect final results ------------------------------------------------
    foreach ($e in $jobs) {
        # Drain all remaining pipeline output into log boxes first
        # Drain any final lines not yet shown
        $remaining = @(Receive-Job $e.Job -Keep -EA SilentlyContinue)
        $startIdx  = $lastLine[$e.OEM]
        if ($remaining.Count -gt $startIdx) {
            for ($ri = $startIdx; $ri -lt $remaining.Count; $ri++) {
                $item = $remaining[$ri]
                if ($item -is [string] -and $item.Trim()) {
                    $oemLogs[$e.OEM].Box.AppendText("$item`r`n")
                    $oemLogs[$e.OEM].Box.ScrollToCaret()
                }
            }
        }
        # Now get the actual result - the last hashtable emitted by the job
        $allOutput = @(Receive-Job $e.Job -EA SilentlyContinue)
        $resultObj = $allOutput | Where-Object {
            $_ -is [System.Collections.Specialized.OrderedDictionary] -or $_ -is [hashtable]
        } | Select-Object -Last 1
        if (-not $resultObj) {
            $resultObj = [ordered]@{ OEM=$e.OEM; Status="FAIL"; Duration=0; Files=0; INFs=0; Notes="No result" }
        }
        $results.Add($resultObj)
        Set-OEMStatus -OEM $resultObj.OEM -Status $resultObj.Status
        Set-OEMStats  -OEM $resultObj.OEM -Files $resultObj.Files -INFs $resultObj.INFs `
                      -Sec $resultObj.Duration -Notes $resultObj.Notes
        $noteStr = if ($resultObj.Notes) { " | $($resultObj.Notes)" } else { "" }
        UILog "$($resultObj.OEM): $($resultObj.Status)  Files:$($resultObj.Files) INFs:$($resultObj.INFs) Time:$($resultObj.Duration)s$noteStr" -OEM $resultObj.OEM
        Remove-Job $e.Job -Force -EA SilentlyContinue
        Remove-Item "$env:TEMP\driver_test_pid_$($e.OEM).txt" -Force -EA SilentlyContinue
    }
    Remove-Item $ScriptCache -Force -EA SilentlyContinue

    # -- Summary --------------------------------------------------------------
    $pass  = ($results | Where-Object { $_.Status -eq "PASS" }).Count
    $fail  = ($results | Where-Object { $_.Status -eq "FAIL" }).Count
    $total = $results.Count
    $dur   = [int]((Get-Date) - $globalStart).TotalSeconds
    $ticker.Stop()

    Add-Content $LogFile "=== SUMMARY: $pass passed, $fail failed ===" -Encoding UTF8
    foreach ($r in $results) {
        Add-Content $LogFile "$($r.OEM): $($r.Status)  $($r.Notes)" -Encoding UTF8
    }

    if ($fail -eq 0) {
        $summaryLbl.Text      = "All $pass/$total tests PASSED in ${dur}s"
        $summaryLbl.ForeColor = $C.Green
        UILog "=== ALL TESTS PASSED ($pass/$total) in ${dur}s ==="
    } else {
        $summaryLbl.Text      = "$fail FAILED  |  $pass passed  |  $total total  |  ${dur}s"
        $summaryLbl.ForeColor = $C.Red
        UILog "=== $fail FAILED, $pass passed of $total in ${dur}s ==="
    }
})

$form.Add_FormClosed({ $ticker.Stop() })
[void]$form.ShowDialog()