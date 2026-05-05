# =============================================================
# DriverInstaller.ps1
# Version: 1.0.0
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

$ScriptVersion = "1.0.0"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
# FORM SETUP
# =========================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Driver Installer Tool  v$ScriptVersion"
$form.Size            = New-Object System.Drawing.Size(580, 460)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 245)

# Title label
$title           = New-Object System.Windows.Forms.Label
$title.AutoSize  = $true
$title.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$title.Location  = New-Object System.Drawing.Point(20, 15)
$title.Text      = "Driver Installer"
$form.Controls.Add($title)

# Version label (top-right)
$versionLabel           = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize  = $true
$versionLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$versionLabel.ForeColor = [System.Drawing.Color]::Gray
$versionLabel.Text      = "v$ScriptVersion"
$versionLabel.Location  = New-Object System.Drawing.Point(510, 20)
$form.Controls.Add($versionLabel)

# Status box (dark terminal style)
$statusBox             = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline   = $true
$statusBox.ScrollBars  = "Vertical"
$statusBox.Size        = New-Object System.Drawing.Size(536, 240)
$statusBox.Location    = New-Object System.Drawing.Point(20, 55)
$statusBox.ReadOnly    = $true
$statusBox.BackColor   = [System.Drawing.Color]::FromArgb(30, 30, 30)
$statusBox.ForeColor   = [System.Drawing.Color]::FromArgb(180, 255, 180)
$statusBox.Font        = New-Object System.Drawing.Font("Consolas", 8.5)
$statusBox.BorderStyle = "FixedSingle"
$form.Controls.Add($statusBox)

# Progress bar
$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Size     = New-Object System.Drawing.Size(536, 22)
$progress.Location = New-Object System.Drawing.Point(20, 308)
$progress.Style    = "Continuous"
$form.Controls.Add($progress)

# Stage label (under progress bar)
$stageLabel           = New-Object System.Windows.Forms.Label
$stageLabel.AutoSize  = $false
$stageLabel.Size      = New-Object System.Drawing.Size(536, 18)
$stageLabel.Location  = New-Object System.Drawing.Point(20, 333)
$stageLabel.ForeColor = [System.Drawing.Color]::DimGray
$stageLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$stageLabel.Text      = ""
$form.Controls.Add($stageLabel)

# Log path label
$logLabel           = New-Object System.Windows.Forms.Label
$logLabel.AutoSize  = $false
$logLabel.Size      = New-Object System.Drawing.Size(536, 16)
$logLabel.Location  = New-Object System.Drawing.Point(20, 355)
$logLabel.ForeColor = [System.Drawing.Color]::Gray
$logLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
$logLabel.Text      = "Log: $LogFile"
$form.Controls.Add($logLabel)

# Button
$button            = New-Object System.Windows.Forms.Button
$button.Text       = "Install Drivers"
$button.Size       = New-Object System.Drawing.Size(160, 36)
$button.Location   = New-Object System.Drawing.Point(200, 380)
$button.Font       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
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

function SetStage($msg) {
    $stageLabel.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

function SetProgress($val) {
    $progress.Value = [math]::Min([math]::Max([int]$val, 0), 100)
    [System.Windows.Forms.Application]::DoEvents()
}

function Assert-Curl {
    if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
        Log "ERROR: curl.exe not found. Windows 10 1803+ required."
        return $false
    }
    return $true
}

# Silent download via curl with live size logging
function Invoke-CurlDownload {
    param(
        [string]$Url,
        [string]$OutFile,
        [int]$TimeoutSec = 600
    )
    Log "Downloading: $([System.IO.Path]::GetFileName($OutFile))"
    Log "  URL: $Url"
    SetStage "Downloading $([System.IO.Path]::GetFileName($OutFile))..."

    $curlArgs = "--location --fail --max-time $TimeoutSec --connect-timeout 30 " +
                "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64)`" " +
                "--output `"$OutFile`" `"$Url`""

    $psi             = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName    = "curl.exe"
    $psi.Arguments   = $curlArgs
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc            = New-Object System.Diagnostics.Process
    $proc.StartInfo  = $psi
    $proc.Start() | Out-Null

    $lastSize = 0; $stall = 0
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 800
        $sz = if (Test-Path $OutFile) { (Get-Item $OutFile -EA SilentlyContinue).Length } else { 0 }
        if ($sz -gt $lastSize) { $stall = 0; $lastSize = $sz } else { $stall++ }
        $mb = [math]::Round($sz / 1MB, 1)
        Log "  Received: $mb MB"
        [System.Windows.Forms.Application]::DoEvents()
        if ($stall -gt 90) { Log "  WARNING: download stalled 72s — aborting."; $proc.Kill(); return $false }
    }

    if ($proc.ExitCode -ne 0) { Log "  curl failed (exit $($proc.ExitCode))"; return $false }
    if (-not (Test-Path $OutFile) -or (Get-Item $OutFile).Length -eq 0) {
        Log "  File missing or empty after download."
        return $false
    }
    $finalMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 1)
    Log "  Download complete: $finalMB MB"
    return $true
}

# Recursively install all INFs found under a path using pnputil
function Install-DriversFromPath {
    param([string]$BasePath)

    $infs = Get-ChildItem $BasePath -Recurse -Filter *.inf -ErrorAction SilentlyContinue
    if (-not $infs -or $infs.Count -eq 0) {
        Log "No INF files found under: $BasePath"
        return $false
    }

    $total = $infs.Count; $i = 0
    Log "Found $total INF file(s) — installing via pnputil..."
    SetStage "Installing drivers ($total INFs)..."

    foreach ($inf in $infs) {
        $i++
        $pct = 60 + [int](($i / $total) * 38)   # maps to 60–98%
        SetProgress $pct
        Log "[$i/$total] $($inf.Name)"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        foreach ($l in $out) { Log "  $l" }
        [System.Windows.Forms.Application]::DoEvents()
    }

    SetProgress 100
    Log "All INFs processed."
    return $true
}

# =========================
# DELL
# =========================
function Start-DellDriverInstall {
    param([string]$DriverRoot)

    Log "=== DELL: Starting automated driver install ==="

    # Get Service Tag
    $serviceTag = $null
    try {
        $serviceTag = (Get-CimInstance Win32_BIOS).SerialNumber.Trim()
        Log "Service Tag: $serviceTag"
    } catch {
        Log "Could not read Service Tag: $($_.Exception.Message)"
        return $false
    }

    if (-not $serviceTag -or $serviceTag.Length -lt 4) {
        Log "Invalid Service Tag detected."
        return $false
    }

    # Get System SKU (used to match catalog)
    $sysId = $null
    try {
        $sysId = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber.Trim()
        Log "System SKU: $sysId"
    } catch {
        Log "Could not read SystemSKUNumber: $($_.Exception.Message)"
    }

    # Download Dell catalog cab
    SetStage "Downloading Dell driver catalog..."
    Log "Downloading Dell CatalogPC.cab..."
    $catalogCab = Join-Path $env:TEMP "DellCatalogPC.cab"
    $catalogXml = Join-Path $env:TEMP "CatalogPC.xml"

    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/CatalogPC.cab" -OutFile $catalogCab -TimeoutSec 180)) {
        Log "Failed to download Dell catalog."
        return $false
    }
    SetProgress 15

    # Extract catalog
    SetStage "Extracting Dell catalog..."
    Log "Extracting Dell catalog..."
    expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1 | Out-Null
    if (-not (Test-Path $catalogXml)) {
        Log "Dell catalog extraction failed — expand.exe may not have output CatalogPC.xml."
        return $false
    }
    SetProgress 20

    # Parse catalog XML (strip BOM)
    Log "Parsing Dell catalog..."
    SetStage "Searching Dell catalog for driver pack..."
    try {
        $rawXml = [System.IO.File]::ReadAllText($catalogXml)
        $rawXml = $rawXml.TrimStart([char]0xFEFF)
        [xml]$cat = $rawXml
    } catch {
        Log "Failed to parse Dell catalog XML: $($_.Exception.Message)"
        return $false
    }

    # Find a SoftwareComponent with "Driver Pack" in its display name that matches this SKU
    $packNode = $null
    foreach ($comp in $cat.SelectNodes("//SoftwareComponent")) {
        # Display name check
        $dispName = ""
        try { $dispName = $comp.Name.Display.InnerText } catch {}
        if ($dispName -notmatch "(?i)driver\s*pack") { continue }

        # SKU match
        if ($sysId) {
            $matched = $false
            foreach ($s in $comp.SelectNodes(".//SystemID")) {
                if ($s.InnerText.Trim() -ieq $sysId) { $matched = $true; break }
            }
            if (-not $matched) { continue }
        }

        $packNode = $comp
        break
    }

    if (-not $packNode) {
        Log "No driver pack found in Dell catalog for SKU '$sysId'."
        Log "Opening Dell support page for manual download..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    $packPath = $packNode.GetAttribute("path")
    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Driver pack found: $([System.IO.Path]::GetFileName($packPath))"
    SetProgress 25

    # Download pack
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "Dell driver pack download failed."
        return $false
    }
    SetProgress 55

    # Extract silently
    $extractPath = Join-Path $DriverRoot "Dell_Extracted"
    if (-not (Test-Path $extractPath)) { New-Item -Path $extractPath -ItemType Directory -Force | Out-Null }
    Log "Extracting Dell pack..."
    SetStage "Extracting Dell driver pack..."
    $proc = Start-Process -FilePath $packFile `
        -ArgumentList "/s /e=`"$extractPath`"" `
        -Wait -PassThru -ErrorAction SilentlyContinue
    if ($proc -and $proc.ExitCode -ne 0) {
        Log "Extraction exit code: $($proc.ExitCode) — checking for INFs anyway..."
    }
    Start-Sleep -Seconds 5
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# HP
# =========================
function Start-HpDriverInstall {
    param([string]$DriverRoot)

    Log "=== HP: Starting automated driver install ==="

    # HP Platform ID from BaseBoard.Product
    $platformId = $null
    try {
        $platformId = (Get-CimInstance Win32_BaseBoard).Product.Trim()
        Log "HP Platform ID: $platformId"
    } catch {
        Log "Could not read HP platform ID: $($_.Exception.Message)"
        return $false
    }

    if (-not $platformId) {
        Log "Empty HP platform ID — cannot continue."
        return $false
    }

    # Download HP platform-specific catalog cab
    SetStage "Downloading HP driver catalog..."
    Log "Fetching HP platform catalog for: $platformId"
    $hpCatalogUrl = "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$platformId/$platformId.cab"
    $hpCatalogCab = Join-Path $env:TEMP "HP_$platformId.cab"
    $hpCatalogXml = Join-Path $env:TEMP "HP_${platformId}.xml"

    if (-not (Invoke-CurlDownload -Url $hpCatalogUrl -OutFile $hpCatalogCab -TimeoutSec 120)) {
        Log "HP platform catalog not available for '$platformId'."
        Log "Opening HP driver pack matrix for manual selection..."
        Start-Process "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html"
        return $false
    }
    SetProgress 15

    # Extract catalog
    SetStage "Extracting HP catalog..."
    expand.exe "`"$hpCatalogCab`"" "`"$hpCatalogXml`"" 2>&1 | Out-Null
    if (-not (Test-Path $hpCatalogXml)) {
        Log "HP catalog extraction failed."
        return $false
    }
    SetProgress 20

    # Parse catalog
    Log "Parsing HP catalog..."
    try {
        $rawXml  = [System.IO.File]::ReadAllText($hpCatalogXml).TrimStart([char]0xFEFF)
        [xml]$hpCat = $rawXml
    } catch {
        Log "Failed to parse HP catalog: $($_.Exception.Message)"
        return $false
    }

    # Find driver pack node
    $packNode = $null
    foreach ($sp in $hpCat.SelectNodes("//SoftPaq")) {
        $catNode = $sp.SelectSingleNode("Category")
        if ($catNode -and $catNode.InnerText -match "(?i)driver\s*pack") {
            $packNode = $sp; break
        }
    }
    # Broader fallback
    if (-not $packNode) {
        $packNode = $hpCat.SelectSingleNode(
            "//SoftPaq[contains(translate(Category,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'driver')]"
        )
    }

    if (-not $packNode) {
        Log "No HP driver pack found in catalog for platform '$platformId'."
        return $false
    }

    $packUrl  = $packNode.SelectSingleNode("Url").InnerText.Trim()
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    Log "HP driver pack: $([System.IO.Path]::GetFileName($packFile))"
    SetProgress 25

    # Download
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "HP driver pack download failed."
        return $false
    }
    SetProgress 55

    # Extract
    $extractPath = Join-Path $DriverRoot "HP_Extracted"
    if (-not (Test-Path $extractPath)) { New-Item -Path $extractPath -ItemType Directory -Force | Out-Null }
    Log "Extracting HP SoftPaq silently..."
    SetStage "Extracting HP driver pack..."
    $proc = Start-Process -FilePath $packFile `
        -ArgumentList "/s /e /f `"$extractPath`"" `
        -Wait -PassThru -ErrorAction SilentlyContinue
    if ($proc -and $proc.ExitCode -ne 0) {
        Log "HP extraction exit code: $($proc.ExitCode) — checking for INFs anyway..."
    }
    Start-Sleep -Seconds 5
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# LENOVO
# =========================
function Start-LenovoDriverInstall {
    param([string]$DriverRoot)

    Log "=== LENOVO: Starting automated driver install ==="

    # Machine type prefix (first 4 chars of Win32_ComputerSystemProduct.Name)
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
        Log "Cannot determine Lenovo machine type — aborting."
        return $false
    }

    # Detect Windows version
    $winVer     = (Get-CimInstance Win32_OperatingSystem).Version
    $osAttr     = if ($winVer -match "^10\.0\.2") { "win11" } else { "win10" }
    $osFallback = if ($osAttr -eq "win11") { "win10" } else { "win11" }
    Log "Detected OS tag: $osAttr"

    # Download catalogv2.xml
    SetStage "Downloading Lenovo driver catalog..."
    Log "Fetching Lenovo catalogv2.xml..."
    $catalogFile = Join-Path $env:TEMP "lenovo_catalogv2.xml"

    if (-not (Invoke-CurlDownload -Url "https://download.lenovo.com/cdrt/td/catalogv2.xml" -OutFile $catalogFile -TimeoutSec 120)) {
        Log "Failed to download Lenovo catalog."
        return $false
    }
    SetProgress 20

    # Parse catalog — read bytes to reliably strip UTF-8 BOM
    Log "Parsing Lenovo catalog..."
    SetStage "Parsing Lenovo catalog..."
    try {
        $bytes   = [System.IO.File]::ReadAllBytes($catalogFile)
        $rawText = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        [xml]$cat = $rawText
        Log "Catalog parsed OK."
    } catch {
        Log "Failed to parse Lenovo catalog: $($_.Exception.Message)"
        return $false
    }
    SetProgress 25

    # Find SCCM driver pack URL for this machine type
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
        Log "Opening Lenovo RecipeCard for manual selection..."
        Start-Process "https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html"
        return $false
    }
    SetProgress 28

    # Download driver pack
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$packUrl).LocalPath))
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile -TimeoutSec 600)) {
        Log "Lenovo driver pack download failed."
        return $false
    }
    SetProgress 55

    # Extract
    $extractPath = Join-Path $DriverRoot "Lenovo_Extracted"
    if (-not (Test-Path $extractPath)) { New-Item -Path $extractPath -ItemType Directory -Force | Out-Null }
    Log "Extracting Lenovo pack silently..."
    SetStage "Extracting Lenovo driver pack..."

    $ext = [System.IO.Path]::GetExtension($packFile).ToLower()
    switch ($ext) {
        ".zip" {
            Expand-Archive -Path $packFile -DestinationPath $extractPath -Force -ErrorAction SilentlyContinue
            Log "ZIP extraction complete."
        }
        ".cab" {
            expand.exe "`"$packFile`"" -F:* "`"$extractPath`"" 2>&1 | ForEach-Object { Log "  $_" }
            Log "CAB extraction complete."
        }
        default {
            # Treat as Inno Setup EXE (most Lenovo SCCM packs)
            $proc = Start-Process -FilePath $packFile `
                -ArgumentList "/VERYSILENT /DIR=`"$extractPath`" /EXTRACT=YES" `
                -Wait -PassThru -ErrorAction SilentlyContinue
            if ($proc -and $proc.ExitCode -ne 0) {
                Log "EXE extraction exit code: $($proc.ExitCode) — checking for INFs anyway..."
            }
            Start-Sleep -Seconds 5
        }
    }
    SetProgress 60

    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# MAIN
# =========================
function Start-Install {

    $button.Enabled = $false
    SetProgress 0

    # Elevation check
    $id  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pri = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $pri.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Log "Not running as admin — re-launching elevated..."
        Start-Process powershell `
            "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/DriverInstaller.ps1 | iex`"" `
            -Verb RunAs
        exit
    }

    Log "Driver Installer v$ScriptVersion"
    Log "Log: $LogFile"
    Log "--------------------------------------------"

    # Detect system
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

    # Result
    Log "--------------------------------------------"
    if ($success) {
        SetProgress 100
        SetStage "Done!"
        Log "Driver installation complete!"
        Log "Log saved to: $LogFile"

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Drivers installed successfully for:`n$model`n`nReboot now to complete installation?",
            "Installation Complete", "YesNo", "Information"
        )
        if ($result -eq "Yes") { Restart-Computer -Force }
        else { $button.Enabled = $true }

    } else {
        SetStage "Failed — check log for details."
        Log "Driver installation did not complete. Check log: $LogFile"
        [System.Windows.Forms.MessageBox]::Show(
            "Driver installation failed or no pack was found.`nCheck the log for details:`n`n$LogFile",
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