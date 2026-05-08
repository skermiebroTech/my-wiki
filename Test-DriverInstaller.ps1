#Requires -RunAsAdministrator
# =============================================================================
# Test-DriverInstaller.ps1  v2.0.0
# Pulls Install-Drivers-auto-7z.ps1 from the repo and runs headless tests
# for Dell, HP, and Lenovo in parallel with a live Windows Forms dashboard.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Test-DriverInstaller.ps1 | iex"
#
# Parameters:
#   -OEM              OEMs to test (default: Dell,HP,Lenovo)
#   -RunInstall       Also run pnputil driver install (default: extract-only)
#   -Branch           GitHub branch to pull from (default: main)
#   -DellModel        Override Dell model string
#   -HPModel          Override HP model string
#   -LenovoMachineType  Lenovo machine type prefix e.g. 20XX
# =============================================================================

param(
    [string[]]$OEM             = @("Dell", "HP", "Lenovo"),
    [switch]$RunInstall,
    [string]$Branch            = "main",
    [string]$DellModel         = "Latitude 7320",
    [string]$HPModel           = "HP EliteBook x360 1030 G8 Notebook PC",
    [string]$LenovoMachineType = "20XX"
)

$ScriptVersion = "2.0.0"
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
# COLOUR PALETTE
# =============================================================================
$C = @{
    Bg         = [System.Drawing.Color]::FromArgb(10,  12,  18)
    Panel      = [System.Drawing.Color]::FromArgb(18,  22,  32)
    PanelBord  = [System.Drawing.Color]::FromArgb(35,  45,  65)
    TextDim    = [System.Drawing.Color]::FromArgb(80,  95, 120)
    TextMid    = [System.Drawing.Color]::FromArgb(140, 160, 190)
    TextBright = [System.Drawing.Color]::FromArgb(215, 225, 240)
    White      = [System.Drawing.Color]::FromArgb(240, 245, 255)
    Accent     = [System.Drawing.Color]::FromArgb(50,  140, 255)
    AccentDim  = [System.Drawing.Color]::FromArgb(25,  70,  140)
    Green      = [System.Drawing.Color]::FromArgb(50,  210, 120)
    GreenDim   = [System.Drawing.Color]::FromArgb(15,  45,  28)
    Red        = [System.Drawing.Color]::FromArgb(255, 80,  80)
    RedDim     = [System.Drawing.Color]::FromArgb(55,  15,  15)
    Yellow     = [System.Drawing.Color]::FromArgb(255, 200, 60)
    YellowDim  = [System.Drawing.Color]::FromArgb(55,  45,  10)
    LogBg      = [System.Drawing.Color]::FromArgb(8,   10,  15)
    LogText    = [System.Drawing.Color]::FromArgb(100, 200, 130)
}

$F = @{
    Title   = New-Object System.Drawing.Font("Consolas", 14, [System.Drawing.FontStyle]::Bold)
    Sub     = New-Object System.Drawing.Font("Consolas", 8,  [System.Drawing.FontStyle]::Regular)
    SubBold = New-Object System.Drawing.Font("Consolas", 8,  [System.Drawing.FontStyle]::Bold)
    OEM     = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    Stat    = New-Object System.Drawing.Font("Consolas", 18, [System.Drawing.FontStyle]::Bold)
    StatSm  = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    Log     = New-Object System.Drawing.Font("Consolas", 8,  [System.Drawing.FontStyle]::Regular)
    Badge   = New-Object System.Drawing.Font("Consolas", 9,  [System.Drawing.FontStyle]::Bold)
}

# =============================================================================
# FORM
# =============================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Driver Installer Test Suite  v$ScriptVersion"
$form.Size            = New-Object System.Drawing.Size(900, 680)
$form.MinimumSize     = New-Object System.Drawing.Size(900, 580)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $C.Bg
$form.ForeColor       = $C.TextBright
$form.FormBorderStyle = "Sizable"

# ---- Header -----------------------------------------------------------------
$header           = New-Object System.Windows.Forms.Panel
$header.Dock      = "Top"
$header.Height    = 58
$header.BackColor = $C.Panel
$form.Controls.Add($header)

$accentLine           = New-Object System.Windows.Forms.Panel
$accentLine.Dock      = "Top"
$accentLine.Height    = 2
$accentLine.BackColor = $C.Accent
$form.Controls.Add($accentLine)

$titleLabel           = New-Object System.Windows.Forms.Label
$titleLabel.Text      = "DRIVER INSTALLER TEST SUITE"
$titleLabel.Font      = $F.Title
$titleLabel.ForeColor = $C.White
$titleLabel.AutoSize  = $true
$titleLabel.Location  = New-Object System.Drawing.Point(20, 10)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$header.Controls.Add($titleLabel)

$vLabel           = New-Object System.Windows.Forms.Label
$vLabel.Text      = "v$ScriptVersion  |  branch: $Branch  |  mode: $(if ($RunInstall) {'FULL INSTALL'} else {'EXTRACT ONLY'})"
$vLabel.Font      = $F.Sub
$vLabel.ForeColor = $C.TextDim
$vLabel.AutoSize  = $true
$vLabel.Location  = New-Object System.Drawing.Point(22, 36)
$vLabel.BackColor = [System.Drawing.Color]::Transparent
$header.Controls.Add($vLabel)

# ---- OEM Panels (3 columns) -------------------------------------------------
$OEMOrder  = @("Dell", "HP", "Lenovo")
$panelW    = 272
$panelH    = 170
$panelGap  = 14
$panelTopY = 72
$oemPanels = @{}

for ($pi = 0; $pi -lt 3; $pi++) {
    $oemName = $OEMOrder[$pi]
    $px      = $panelGap + $pi * ($panelW + $panelGap)

    $p            = New-Object System.Windows.Forms.Panel
    $p.Size       = New-Object System.Drawing.Size($panelW, $panelH)
    $p.Location   = New-Object System.Drawing.Point($px, $panelTopY)
    $p.BackColor  = $C.Panel
    $form.Controls.Add($p)

    # Top colour strip
    $strip           = New-Object System.Windows.Forms.Panel
    $strip.Size      = New-Object System.Drawing.Size($panelW, 3)
    $strip.Location  = New-Object System.Drawing.Point(0, 0)
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
    $nameL.ForeColor = $C.TextBright
    $nameL.BackColor = [System.Drawing.Color]::Transparent
    $nameL.AutoSize  = $true
    $nameL.Location  = New-Object System.Drawing.Point(12, 30)
    $p.Controls.Add($nameL)

    # Model/type sub-label
    $cfg = $OEMConfig[$oemName]
    $mt  = if ($cfg.Model) { $cfg.Model } elseif ($cfg.MachineType) { "Type: $($cfg.MachineType)" } else { "WMI" }
    if ($mt.Length -gt 36) { $mt = $mt.Substring(0, 33) + "..." }
    $modelL           = New-Object System.Windows.Forms.Label
    $modelL.Text      = $mt
    $modelL.Font      = $F.Sub
    $modelL.ForeColor = $C.TextDim
    $modelL.BackColor = [System.Drawing.Color]::Transparent
    $modelL.AutoSize  = $false
    $modelL.Size      = New-Object System.Drawing.Size(248, 14)
    $modelL.Location  = New-Object System.Drawing.Point(12, 52)
    $p.Controls.Add($modelL)

    # Divider
    $div           = New-Object System.Windows.Forms.Panel
    $div.Size      = New-Object System.Drawing.Size(248, 1)
    $div.Location  = New-Object System.Drawing.Point(12, 70)
    $div.BackColor = $C.PanelBord
    $p.Controls.Add($div)

    # Stats: Files / INFs / Sec
    $statsData = @(
        @{ Key="Files"; X=10  }
        @{ Key="INFs";  X=100 }
        @{ Key="Sec";   X=190 }
    )
    $statLabels = @{}
    foreach ($s in $statsData) {
        $numL           = New-Object System.Windows.Forms.Label
        $numL.Text      = "--"
        $numL.Font      = $F.Stat
        $numL.ForeColor = $C.TextDim
        $numL.BackColor = [System.Drawing.Color]::Transparent
        $numL.AutoSize  = $true
        $numL.Location  = New-Object System.Drawing.Point($s.X, 76)
        $p.Controls.Add($numL)

        $capL           = New-Object System.Windows.Forms.Label
        $capL.Text      = $s.Key
        $capL.Font      = $F.Sub
        $capL.ForeColor = $C.TextDim
        $capL.BackColor = [System.Drawing.Color]::Transparent
        $capL.AutoSize  = $true
        $capL.Location  = New-Object System.Drawing.Point(($s.X + 2), 108)
        $p.Controls.Add($capL)

        $statLabels[$s.Key] = $numL
    }

    # Notes label
    $notesL           = New-Object System.Windows.Forms.Label
    $notesL.Text      = ""
    $notesL.Font      = $F.Sub
    $notesL.ForeColor = $C.TextDim
    $notesL.BackColor = [System.Drawing.Color]::Transparent
    $notesL.AutoSize  = $false
    $notesL.Size      = New-Object System.Drawing.Size(248, 14)
    $notesL.Location  = New-Object System.Drawing.Point(12, 124)
    $p.Controls.Add($notesL)

    # Progress bar
    $barY = [int]$panelH - 5
    $bar                       = New-Object System.Windows.Forms.ProgressBar
    $bar.Style                 = "Marquee"
    $bar.MarqueeAnimationSpeed = 0
    $bar.Size                  = New-Object System.Drawing.Size($panelW, 5)
    $bar.Location              = New-Object System.Drawing.Point(0, $barY)
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
    }
}

# ---- Summary bar ------------------------------------------------------------
$summaryY             = [int]$panelTopY + [int]$panelH + 10
$summaryPanel         = New-Object System.Windows.Forms.Panel
$summaryPanel.Size    = New-Object System.Drawing.Size(858, 36)
$summaryPanel.Location= New-Object System.Drawing.Point($panelGap, $summaryY)
$summaryPanel.BackColor = $C.Panel
$form.Controls.Add($summaryPanel)

$summaryLbl           = New-Object System.Windows.Forms.Label
$summaryLbl.Text      = "Initialising..."
$summaryLbl.Font      = $F.SubBold
$summaryLbl.ForeColor = $C.TextMid
$summaryLbl.BackColor = [System.Drawing.Color]::Transparent
$summaryLbl.AutoSize  = $false
$summaryLbl.Size      = New-Object System.Drawing.Size(720, 36)
$summaryLbl.Location  = New-Object System.Drawing.Point(12, 0)
$summaryLbl.TextAlign = "MiddleLeft"
$summaryPanel.Controls.Add($summaryLbl)

$elapsedLbl           = New-Object System.Windows.Forms.Label
$elapsedLbl.Text      = "00:00"
$elapsedLbl.Font      = $F.StatSm
$elapsedLbl.ForeColor = $C.TextDim
$elapsedLbl.BackColor = [System.Drawing.Color]::Transparent
$elapsedLbl.AutoSize  = $true
$elapsedLbl.Location  = New-Object System.Drawing.Point(810, 8)
$summaryPanel.Controls.Add($elapsedLbl)

# ---- Log box ----------------------------------------------------------------
$logY        = [int]$summaryY + 36 + 8
$logCaption  = New-Object System.Windows.Forms.Label
$logCaption.Text      = "LIVE LOG"
$logCaption.Font      = $F.SubBold
$logCaption.ForeColor = $C.TextDim
$logCaption.BackColor = [System.Drawing.Color]::Transparent
$logCaption.AutoSize  = $true
$logCaption.Location  = New-Object System.Drawing.Point([int]$panelGap, [int]$logY)
$form.Controls.Add($logCaption)

$logBox             = New-Object System.Windows.Forms.RichTextBox
$logBox.Multiline   = $true
$logBox.ScrollBars  = "Vertical"
$logBox.ReadOnly    = $true
$logBox.BackColor   = $C.LogBg
$logBox.ForeColor   = $C.LogText
$logBox.Font        = $F.Log
$logBox.BorderStyle = "None"
$logBoxTop = [int]$logY + 18
$logBox.Location    = New-Object System.Drawing.Point([int]$panelGap, $logBoxTop)
$logBoxH = 680 - [int]$logY - 55
$logBox.Size        = New-Object System.Drawing.Size(858, $logBoxH)
$form.Controls.Add($logBox)
$logBox.BringToFront()
$logCaption.BringToFront()

# ---- Footer -----------------------------------------------------------------
$footer           = New-Object System.Windows.Forms.Panel
$footer.Dock      = "Bottom"
$footer.Height    = 28
$footer.BackColor = $C.Panel
$form.Controls.Add($footer)

$logPathLbl           = New-Object System.Windows.Forms.Label
$logPathLbl.Text      = "Log: $LogFile"
$logPathLbl.Font      = $F.Sub
$logPathLbl.ForeColor = $C.TextDim
$logPathLbl.BackColor = [System.Drawing.Color]::Transparent
$logPathLbl.AutoSize  = $true
$logPathLbl.Location  = New-Object System.Drawing.Point(12, 7)
$footer.Controls.Add($logPathLbl)

$killBtn                            = New-Object System.Windows.Forms.Button
$killBtn.Text                       = "KILL ALL JOBS"
$killBtn.Font                       = $F.SubBold
$killBtn.ForeColor                  = $C.White
$killBtn.BackColor                  = [System.Drawing.Color]::FromArgb(130, 20, 20)
$killBtn.FlatStyle                  = "Flat"
$killBtn.FlatAppearance.BorderSize  = 0
$killBtn.Size                       = New-Object System.Drawing.Size(120, 22)
$killBtn.Location                   = New-Object System.Drawing.Point(762, 3)
$footer.Controls.Add($killBtn)

# Resize handler
$form.Add_Resize({
    $w = $form.ClientSize.Width; $h = $form.ClientSize.Height
    $header.Width      = $w
    $accentLine.Width  = $w
    $summaryPanel.Width = $w - ($panelGap * 2)
    $logBox.Width      = $w - ($panelGap * 2)
    $logBox.Height     = $h - [int]$logY - 55
    [System.Windows.Forms.Application]::DoEvents()
})

# =============================================================================
# HELPERS
# =============================================================================
function UILog {
    param([string]$msg)
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $msg"
    $logBox.AppendText("$line`r`n")
    $logBox.ScrollToCaret()
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-OEMStatus {
    param([string]$OEM, [string]$Status)
    $pnl = $oemPanels[$OEM]
    switch ($Status) {
        "RUNNING" {
            $pnl.Badge.Text      = "RUNNING"
            $pnl.Badge.ForeColor = $C.Yellow
            $pnl.Strip.BackColor = $C.Yellow
            $pnl.Panel.BackColor = $C.YellowDim
            $pnl.Bar.MarqueeAnimationSpeed = 25
            $pnl.Bar.Style = "Marquee"
        }
        "PASS" {
            $pnl.Badge.Text      = "PASS"
            $pnl.Badge.ForeColor = $C.Green
            $pnl.Strip.BackColor = $C.Green
            $pnl.Panel.BackColor = $C.GreenDim
            $pnl.Bar.Style = "Continuous"; $pnl.Bar.Value = 100
        }
        "FAIL" {
            $pnl.Badge.Text      = "FAIL"
            $pnl.Badge.ForeColor = $C.Red
            $pnl.Strip.BackColor = $C.Red
            $pnl.Panel.BackColor = $C.RedDim
            $pnl.Bar.Style = "Continuous"; $pnl.Bar.Value = 100
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
    $pnl.Stats["Files"].Text = "$Files"
    $pnl.Stats["INFs"].Text  = "$INFs"
    $pnl.Stats["Sec"].Text   = "$([int]$Sec)"
    $pnl.Stats["Files"].ForeColor = if ($Files -gt 0) { $C.White } else { $C.Red }
    $pnl.Stats["INFs"].ForeColor  = if ($INFs  -gt 0) { $C.White } else { $C.Red }
    $pnl.Stats["Sec"].ForeColor   = $C.TextMid
    if ($Notes) {
        $short = if ($Notes.Length -gt 38) { $Notes.Substring(0,35) + "..." } else { $Notes }
        $pnl.NotesLbl.Text      = $short
        $pnl.NotesLbl.ForeColor = $C.Red
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Mark skipped OEMs
foreach ($n in $OEMOrder) { if ($OEM -notcontains $n) { Set-OEMStatus -OEM $n -Status "SKIP" } }

# =============================================================================
# TIMER — elapsed clock
# =============================================================================
$globalStart = Get-Date
$ticker      = New-Object System.Windows.Forms.Timer
$ticker.Interval = 500
$ticker.Add_Tick({
    $e = [int]((Get-Date) - $globalStart).TotalSeconds
    $elapsedLbl.Text = "{0:D2}:{1:D2}" -f ([int]($e/60)), ($e%60)
    [System.Windows.Forms.Application]::DoEvents()
})
$ticker.Start()

# =============================================================================
# MAIN LOGIC
# =============================================================================
$script:KillRequested = $false
$script:ActiveJobs    = $null   # will be set once jobs are launched

$killBtn.Add_Click({
    if ($script:KillRequested) { return }
    $script:KillRequested = $true
    $killBtn.Text      = "KILLING..."
    $killBtn.BackColor = [System.Drawing.Color]::FromArgb(80, 10, 10)
    $killBtn.Enabled   = $false
    $summaryLbl.Text      = "Kill requested - stopping all jobs..."
    $summaryLbl.ForeColor = $C.Red
    UILog "=== KILL REQUESTED BY USER ==="

    # Stop all tracked jobs
    if ($script:ActiveJobs) {
        foreach ($e in $script:ActiveJobs) {
            try {
                Stop-Job  $e.Job -EA SilentlyContinue
                Remove-Job $e.Job -Force -EA SilentlyContinue
                UILog "[$($e.OEM)] Job stopped"
                Set-OEMStatus -OEM $e.OEM -Status "FAIL"
            } catch {}
        }
    }

    # Kill any powershell child processes running the installer
    try {
        Get-Process powershell -EA SilentlyContinue |
            Where-Object { $_.MainWindowTitle -eq "" -and $_.Id -ne $PID } |
            ForEach-Object { $_.Kill(); UILog "Killed PID $($_.Id)" }
    } catch {}

    # Clean up driver folders
    foreach ($n in $OEMOrder) {
        $dr = "C:\DRIVERS\$n"
        if (Test-Path $dr) {
            Remove-Item $dr -Recurse -Force -EA SilentlyContinue
            UILog "[$n] Cleaned up $dr"
        }
    }
    Remove-Item $ScriptCache -Force -EA SilentlyContinue
    $ticker.Stop()
    $summaryLbl.Text = "All jobs killed. Cleaned up."
    UILog "=== KILL COMPLETE ==="
    $killBtn.Text      = "KILLED"
    $killBtn.BackColor = [System.Drawing.Color]::FromArgb(50, 10, 10)
    [System.Windows.Forms.Application]::DoEvents()
})

$form.Add_Shown({
    $form.Activate()
    [System.Windows.Forms.Application]::DoEvents()
    New-Item -ItemType File -Path $LogFile -Force | Out-Null

    # -- Download script ------------------------------------------------------
    $summaryLbl.Text      = "Downloading Install-Drivers-auto-7z.ps1 from repo..."
    $summaryLbl.ForeColor = $C.Accent
    UILog "Repo: $RepoUrl"

    $cx = (Start-Process curl.exe -ArgumentList "--silent --location --output `"$ScriptCache`" `"$RepoUrl`"" -Wait -PassThru).ExitCode
    if ($cx -ne 0 -or -not (Test-Path $ScriptCache)) {
        $summaryLbl.Text = "ERROR: Could not download script (curl exit $cx)"
        $summaryLbl.ForeColor = $C.Red
        UILog "FAIL: Download failed"
        return
    }
    $kb  = [math]::Round((Get-Item $ScriptCache).Length / 1KB, 1)
    $vl  = Get-Content $ScriptCache | Select-String '^\$ScriptVersion\s*=' | Select-Object -First 1
    $ver = if ($vl) { $vl -replace '.*"(.*)".*','$1' } else { "?" }
    UILog "Downloaded $kb KB  |  Script v$ver"
    $summaryLbl.Text      = "Script v$ver ready. Launching parallel jobs..."
    $summaryLbl.ForeColor = $C.TextBright

    # -- Launch jobs ----------------------------------------------------------
    $results = [System.Collections.Generic.List[object]]::new()
    $jobs    = [System.Collections.Generic.List[object]]::new()

    foreach ($oemName in $OEM) {
        if (-not $OEMConfig.ContainsKey($oemName)) { UILog "Unknown OEM '$oemName' - skipping"; continue }

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
        UILog "[$oemName] Launching job..."

        $jAl = $al; $jN = $oemName; $jDr = $driverRoot; $jEd = $extractDir
        $jLf = $LogFile; $jRi = $RunInstall

        $job = Start-Job -ScriptBlock {
            param($al,$n,$dr,$ed,$lf,$ri)
            function JL { param([string]$m)
                Add-Content $lf "[$([datetime]::Now.ToString('HH:mm:ss'))][$n] $m" -Encoding UTF8
                $m
            }
            $st  = Get-Date
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName=$("powershell.exe"); $psi.Arguments=$al
            $psi.UseShellExecute=$false; $psi.CreateNoWindow=$true
            $psi.RedirectStandardOutput=$true; $psi.RedirectStandardError=$true
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo=$psi; $proc.Start()|Out-Null
            $out=[System.Collections.Generic.List[string]]::new()
            while(-not $proc.HasExited){$l=$proc.StandardOutput.ReadLine();if($l-ne$null){$out.Add($l);JL "OUT: $l"|Out-Null}}
            ($proc.StandardOutput.ReadToEnd()-split"`n"|Where-Object{$_.Trim()})|ForEach-Object{$out.Add($_);JL "OUT: $_"|Out-Null}
            $err=$proc.StandardError.ReadToEnd().Trim();if($err){JL "ERR: $err"|Out-Null}
            $exit=$proc.ExitCode; $dur=[math]::Round(((Get-Date)-$st).TotalSeconds,1)
            $ok=$true; $notes=[System.Collections.Generic.List[string]]::new(); $fc=0; $ic=0
            if($exit-ne 0){$notes.Add("Exit:$exit");$ok=$false}
            if(Test-Path $ed){
                $fc=@(Get-ChildItem $ed -Recurse -File -EA SilentlyContinue).Count
                $ic=@(Get-ChildItem $ed -Recurse -Filter "*.inf" -EA SilentlyContinue).Count
                if($fc-eq 0){$notes.Add("0 files");$ok=$false}
                if($ic-eq 0){$notes.Add("0 INFs");$ok=$false}
            } else {$notes.Add("No extract dir");$ok=$false}
            if($n-in@("Dell","HP")-and(Test-Path "C:\Program Files\7-Zip\7z.exe")){$notes.Add("7z not removed");$ok=$false}
            $sl=$out|Where-Object{$_-match"SUCCESS:|complete"}|Select-Object -First 1
            $fl=$out|Where-Object{$_-match"FAILED:|did not"}|Select-Object -First 1
            if(-not $sl-and $fl){$notes.Add("Script fail");$ok=$false}
            if(-not $sl-and-not $fl){$notes.Add("Unclear outcome")}
            if(Test-Path $dr){Remove-Item $dr -Recurse -Force -EA SilentlyContinue}
            JL "END $n $(if($ok){'PASS'}else{'FAIL'})"|Out-Null
            [ordered]@{OEM=$n;Status=if($ok){"PASS"}else{"FAIL"};Duration=$dur;Files=$fc;INFs=$ic;Notes=($notes-join"; ")}
        } -ArgumentList $jAl,$jN,$jDr,$jEd,$jLf,$jRi

        $jobs.Add([ordered]@{ OEM=$oemName; Job=$job; Start=(Get-Date) })
        $script:ActiveJobs = $jobs
        UILog "[$oemName] Job $($job.Id) started"
    }

    # -- Wait loop ------------------------------------------------------------
    $done = @{}
    while ($done.Count -lt $jobs.Count) {
        foreach ($e in $jobs) {
            $n = $e.OEM; $j = $e.Job
            if ($done.ContainsKey($n)) { continue }
            $sec = [int]((Get-Date)-$e.Start).TotalSeconds
            $oemPanels[$n].Stats["Sec"].Text = "$sec"
            if ($j.State -in @("Completed","Failed","Stopped")) {
                $done[$n] = $true
                UILog "[$n] Job done in ${sec}s"
            }
        }
        $still = ($jobs | Where-Object { -not $done.ContainsKey($_.OEM) } | ForEach-Object { $_.OEM }) -join ", "
        if ($still) {
            $summaryLbl.Text      = "Running: $still"
            $summaryLbl.ForeColor = $C.Yellow
        }
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:KillRequested) { break }
        if ($done.Count -lt $jobs.Count) { Start-Sleep -Milliseconds 800 }
    }

    # -- Collect results ------------------------------------------------------
    foreach ($e in $jobs) {
        $r = Receive-Job $e.Job -EA SilentlyContinue
        if (-not $r) { $r = [ordered]@{OEM=$e.OEM;Status="FAIL";Duration=0;Files=0;INFs=0;Notes="No result returned"} }
        $results.Add($r)
        Set-OEMStatus -OEM $r.OEM -Status $r.Status
        Set-OEMStats  -OEM $r.OEM -Files $r.Files -INFs $r.INFs -Sec $r.Duration -Notes $r.Notes
        UILog "[$($r.OEM)] $($r.Status) - Files:$($r.Files) INFs:$($r.INFs) Time:$($r.Duration)s$(if($r.Notes){" | $($r.Notes)"})"
        Remove-Job $e.Job -Force -EA SilentlyContinue
    }
    Remove-Item $ScriptCache -Force -EA SilentlyContinue

    # -- Final summary --------------------------------------------------------
    $pass  = ($results | Where-Object { $_.Status -eq "PASS" }).Count
    $fail  = ($results | Where-Object { $_.Status -eq "FAIL" }).Count
    $total = $results.Count
    $dur   = [int]((Get-Date)-$globalStart).TotalSeconds
    $ticker.Stop()

    Add-Content $LogFile "=== SUMMARY: $pass passed, $fail failed ===" -Encoding UTF8
    foreach ($r in $results) { Add-Content $LogFile "$($r.OEM): $($r.Status)  $($r.Notes)" -Encoding UTF8 }

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