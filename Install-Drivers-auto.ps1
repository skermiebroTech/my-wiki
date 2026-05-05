# =============================================================================
# Driver Installer Tool
# Version: 1.0.7
# Date:    2026-05-05
#
# Changelog:
#   1.0.7 - Use Lenovo.Client.Scripting PSGallery module + catalogv2.xml for
#            direct driver pack URL resolution; page-scrape as final fallback
#   1.0.6 - Switch to Invoke-WebRequest for support page fetch (curl blocked);
#            extract URLs from customData JSON blob in page HTML
#   1.0.5 - Add progress label under progress bar; fix not-responding by using
#            temp file for page fetch instead of blocking ReadToEnd()
#   1.0.4 - Switch to Lenovo eSupport Content API (v2.5) for URL resolution;
#            fix $var: drive-reference parse errors throughout
#   1.0.3 - Add OS fallback (Win11->Win10) using lowercase docId; multi-region
#            API endpoint retry; real download progress bar with MB display
#   1.0.2 - Add log file to Downloads folder; fix $true array leak in URL
#            pipeline; add Test-ModelKeyword helper
#   1.0.1 - Integrate Lenovo recipecard.json auto-download; SMBIOS detection;
#            string ID comparison fixes; Win11/Win10 OS fallback logic
#   1.0.0 - Initial release: GUI driver installer for Dell, HP, Lenovo
# =============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Disable sleep and display timeout (AC power only)
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# =========================
# LOG FILE SETUP
# =========================
$LogFile = Join-Path ([Environment]::GetFolderPath("UserProfile")) ("Downloads\DriverInstaller_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
New-Item -ItemType File -Path $LogFile -Force | Out-Null

# =========================
# FORM SETUP
# =========================
$form = New-Object System.Windows.Forms.Form
$ScriptVersion = "1.0.7"
$form.Text = "Driver Installer Tool v${ScriptVersion}"
$form.Size = New-Object System.Drawing.Size(560, 445)
$form.StartPosition = "CenterScreen"

# Title
$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120, 15)
$form.Controls.Add($title)

# Status box
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Size = New-Object System.Drawing.Size(500, 220)
$statusBox.Location = New-Object System.Drawing.Point(20, 60)
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

# Progress bar
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(500, 20)
$progress.Location = New-Object System.Drawing.Point(20, 295)
$form.Controls.Add($progress)

# Progress label (shows current stage under the bar)
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.AutoSize = $false
$progressLabel.Size = New-Object System.Drawing.Size(500, 18)
$progressLabel.Location = New-Object System.Drawing.Point(20, 318)
$progressLabel.ForeColor = [System.Drawing.Color]::DimGray
$progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$progressLabel.Text = ""
$form.Controls.Add($progressLabel)

# Version label (bottom-right corner)
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize = $true
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$versionLabel.ForeColor = [System.Drawing.Color]::DarkGray
$versionLabel.Text = "v${ScriptVersion}"
$versionLabel.Location = New-Object System.Drawing.Point(490, 420)
$form.Controls.Add($versionLabel)

# Button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Install Drivers"
$button.Size = New-Object System.Drawing.Size(160, 35)
$button.Location = New-Object System.Drawing.Point(195, 355)
$form.Controls.Add($button)

# =========================
# LOG FUNCTION
# =========================
function Log($msg) {
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $line = "[$timestamp] $msg"
    $statusBox.AppendText("$line`r`n")
    $statusBox.ScrollToCaret()
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    [System.Windows.Forms.Application]::DoEvents()
}

# Update the progress label shown under the progress bar
function SetStage($msg) {
    $progressLabel.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

# =========================
# SYSTEM INFO
# =========================
function Get-Model {
    (Get-CimInstance Win32_ComputerSystem).Model.Trim()
}

function Copy-ModelToClipboard {
    $model = Get-Model
    try { [System.Windows.Forms.Clipboard]::SetText($model) } catch {}
    return $model
}

# =========================
# OEM LINKS
# =========================
function Get-DriverLink($manufacturer) {
    switch -Wildcard ($manufacturer) {
        "*Dell*"   { return "https://www.dell.com/support/kbdoc/en-au/000124139/dell-command-deploy-driver-packs-for-enterprise-client-os-deployment" }
        "*HP*"     { return "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html" }
        "*Lenovo*" { return "https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html" }
        default     { return "" }
    }
}

# =========================
# DOWNLOADS EXE CHECK
# =========================
function Check-DownloadsForExe {
    $downloads = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    if (Test-Path $downloads) {
        $exe = Get-ChildItem $downloads -Filter *.exe -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending |
               Select-Object -First 1
        if ($exe) {
            Log "Found EXE in Downloads: $($exe.Name)"
            Log "Launching installer..."
            Start-Process $exe.FullName
            return $true
        }
    }
    return $false
}

# =========================
# LENOVO: GET MACHINE TYPE PREFIX VIA WMI
# =========================
function Get-LenovoMachineTypePrefix {
    Log "Querying WMI for Lenovo machine type..."
    try {
        $csProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop
        $skuName = $csProduct.Name
        if ($skuName -and $skuName.Trim().Length -ge 4) {
            $prefix = $skuName.Trim().Substring(0, 4).ToUpper()
            Log "Detected machine type: $skuName  ->  using prefix: $prefix"
            return $prefix
        } else {
            Log "WMI returned an unexpected or short SKU name: '$skuName'"
            return $null
        }
    } catch {
        Log "WMI query failed: $($_.Exception.Message)"
        return $null
    }
}

# =========================
# LENOVO: GET DRIVER PACK URL VIA OFFICIAL MODULE OR DIRECT BITS DOWNLOAD
# =========================

function Install-LenovoClientModule {
    # Install Lenovo's official Client Scripting Module from PowerShell Gallery
    # This module uses catalogv2.xml and handles all URL resolution internally
    try {
        if (-not (Get-Module -ListAvailable -Name "Lenovo.Client.Scripting" -ErrorAction SilentlyContinue)) {
            Log "Installing Lenovo.Client.Scripting module from PowerShell Gallery..."
            SetStage "Installing Lenovo module..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-Module -Name "Lenovo.Client.Scripting" -Force -Scope CurrentUser -ErrorAction Stop
            Log "Module installed OK."
        }
        Import-Module "Lenovo.Client.Scripting" -Force -ErrorAction Stop
        Log "Lenovo.Client.Scripting module loaded."
        return $true
    } catch {
        Log "Could not install/load Lenovo.Client.Scripting: $($_.Exception.Message)"
        return $false
    }
}

function Get-LenovoSccmUrlFromCatalog {
    param(
        [string[]]$SmbiosCodes,
        [string]$TargetOSName = "Windows 11",
        [string]$MachineType  = ""
    )

    $winVer = if ($TargetOSName -match "11") { "11" } else { "10" }

    # Try Lenovo's official PowerShell module first (uses catalogv2.xml internally)
    if (Install-LenovoClientModule) {
        try {
            Log "Using Get-LnvDriverPack for machine type ${MachineType}, Windows ${winVer}, Latest..."
            SetStage "Finding driver pack via Lenovo module..."

            # Find-LnvDriverPack returns URL without downloading; try Latest build version first
            $packInfo = Find-LnvDriverPack -MachineType $MachineType -WindowsVersion $winVer -OSBuildVersion "Latest" -ErrorAction Stop
            if ($packInfo -and $packInfo.URL) {
                Log "Found SCCM pack URL via module: $($packInfo.URL)"
                return $packInfo.URL
            }
        } catch {
            Log "Find-LnvDriverPack failed: $($_.Exception.Message)"
        }

        # Try without OSBuildVersion (older module versions)
        try {
            $packInfo = Find-LnvDriverPack -MachineType $MachineType -WindowsVersion $winVer -ErrorAction Stop
            if ($packInfo -and $packInfo.URL) {
                Log "Found SCCM pack URL via module (no build ver): $($packInfo.URL)"
                return $packInfo.URL
            }
        } catch {
            Log "Find-LnvDriverPack (no build ver) failed: $($_.Exception.Message)"
        }

        # Win10 fallback
        if ($winVer -eq "11") {
            Log "No Win11 pack found — trying Win10 fallback via module..."
            try {
                $packInfo = Find-LnvDriverPack -MachineType $MachineType -WindowsVersion "10" -OSBuildVersion "Latest" -ErrorAction Stop
                if ($packInfo -and $packInfo.URL) {
                    Log "Found Win10 fallback SCCM URL: $($packInfo.URL)"
                    return $packInfo.URL
                }
            } catch {
                Log "Win10 fallback also failed: $($_.Exception.Message)"
            }
        }
    }

    # Module failed - try catalogv2.xml directly (the XML the module uses internally)
    Log "Trying catalogv2.xml directly..."
    SetStage "Reading Lenovo driver catalog..."
    try {
        $Global:ProgressPreference = 'SilentlyContinue'
        $resp = Invoke-WebRequest -Uri "https://download.lenovo.com/cdrt/td/catalogv2.xml" `
            -UseBasicParsing -TimeoutSec 60 `
            -Headers @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" } `
            -ErrorAction Stop
        $Global:ProgressPreference = 'Continue'

        [xml]$cat = $resp.Content
        [System.Windows.Forms.Application]::DoEvents()

        # catalogv2.xml structure: <Catalog><Model><Types><Type>{MT}</Type></Types><SCCM os="win11">{URL}</SCCM></Model>
        $osAttr = if ($winVer -eq "11") { "win11" } else { "win10" }
        $osFallback = if ($osAttr -eq "win11") { "win10" } else { "win11" }

        foreach ($model in $cat.Catalog.Model) {
            $types = @($model.Types.Type)
            $matched = $types | Where-Object { $_ -like "${MachineType}*" }
            if (-not $matched) { continue }

            Log "Matched model in catalogv2.xml for ${MachineType}"
            # Try preferred OS first then fallback
            foreach ($os in @($osAttr, $osFallback)) {
                $sccmNode = $model.SCCM | Where-Object { $_.os -eq $os } | Select-Object -First 1
                if ($sccmNode -and $sccmNode."#text" -match "^https?://") {
                    $url = $sccmNode."#text"
                    Log "catalogv2.xml SCCM URL [${os}]: ${url}"
                    return $url
                }
            }
        }
        Log "Machine type ${MachineType} not found in catalogv2.xml"
    } catch {
        $Global:ProgressPreference = 'Continue'
        Log "catalogv2.xml fetch failed: $($_.Exception.Message)"
    }

    return $null
}

function Get-LenovoBiosUrlFromCatalog {
    param(
        [string[]]$SmbiosCodes,
        [string]$TargetOSName = "Windows 11",
        [string]$MachineType  = ""
    )

    $winVer = if ($TargetOSName -match "11") { "11" } else { "10" }

    if (Get-Module -Name "Lenovo.Client.Scripting" -ErrorAction SilentlyContinue) {
        try {
            $biosInfo = Get-LnvBiosUpdateUrl -MachineType $MachineType -ErrorAction Stop
            if ($biosInfo -and $biosInfo.URL) {
                Log "BIOS URL via module: $($biosInfo.URL)"
                return $biosInfo.URL
            }
        } catch {
            Log "Get-LnvBiosUpdateUrl failed: $($_.Exception.Message)"
        }
    }

    # catalogv2.xml BIOS fallback
    try {
        if (-not $script:catalogXml) {
            $Global:ProgressPreference = 'SilentlyContinue'
            $resp = Invoke-WebRequest -Uri "https://download.lenovo.com/cdrt/td/catalogv2.xml" `
                -UseBasicParsing -TimeoutSec 60 `
                -Headers @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" } `
                -ErrorAction Stop
            $Global:ProgressPreference = 'Continue'
            [xml]$script:catalogXml = $resp.Content
        }
        foreach ($model in $script:catalogXml.Catalog.Model) {
            $types = @($model.Types.Type)
            if (-not ($types | Where-Object { $_ -like "${MachineType}*" })) { continue }
            $biosUrl = $model.BIOS
            if ($biosUrl -and $biosUrl -match "^https?://") {
                Log "BIOS URL from catalogv2.xml: ${biosUrl}"
                return $biosUrl
            }
        }
    } catch {
        $Global:ProgressPreference = 'Continue'
        Log "catalogv2.xml BIOS lookup failed: $($_.Exception.Message)"
    }

    return $null
}

# Page-scrape fallback (support.lenovo.com may be blocked on some networks)
function Get-LenovoDirectDownloadLink {
    param(
        [string]$SupportPageUrl,
        [string]$ModelName,
        [string]$TargetOSName,
        [string]$LinkType = "SCCM"
    )

    Log "Fetching ${LinkType} support page: ${SupportPageUrl}"
    SetStage "Fetching ${LinkType} support page..."

    $pageContent = $null
    try {
        $Global:ProgressPreference = 'SilentlyContinue'
        $resp = Invoke-WebRequest -Uri $SupportPageUrl -UseBasicParsing -TimeoutSec 45 `
            -Headers @{
                "Accept"          = "text/html,application/xhtml+xml,*/*;q=0.8"
                "Accept-Language" = "en-US,en;q=0.9"
                "User-Agent"      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            } -ErrorAction Stop
        $Global:ProgressPreference = 'Continue'
        $pageContent = $resp.Content
        Log "Page fetched OK ($($pageContent.Length) bytes)"
    } catch {
        $Global:ProgressPreference = 'Continue'
        Log "Failed to fetch ${LinkType} page: $($_.Exception.Message)"
        return $null
    }

    [System.Windows.Forms.Application]::DoEvents()
    if (-not $pageContent -or $pageContent.Length -lt 100) { return $null }

    $urlOsPairs = @()
    $pairPattern = '"URL":"(https?://download\.lenovo\.com/[^"]+\.(?:exe|zip|cab))"[^}]{0,500}?"OperatingSystemKeys":\["([^"]+)"\]'
    foreach ($m in [regex]::Matches($pageContent, $pairPattern)) {
        $urlOsPairs += [PSCustomObject]@{ URL = $m.Groups[1].Value; OS = $m.Groups[2].Value }
    }
    if ($urlOsPairs.Count -eq 0) {
        foreach ($m in [regex]::Matches($pageContent, '"URL":"(https?://download\.lenovo\.com/[^"]+\.(?:exe|zip|cab))"')) {
            $url = $m.Groups[1].Value
            if ($url -notmatch "readme") { $urlOsPairs += [PSCustomObject]@{ URL = $url; OS = "unknown" } }
        }
    }

    if ($urlOsPairs.Count -eq 0) { Log "No URLs found on ${LinkType} page."; return $null }
    Log "Found $($urlOsPairs.Count) candidate(s) from page."

    if ($LinkType -eq "SCCM") {
        $win11 = $urlOsPairs | Where-Object { $_.OS -match "Windows 11" -and $_.URL -notmatch "readme" }
        $preferred = $win11 | Where-Object { $_.URL -match "w11_24|win11_24" } | Select-Object -Last 1
        if (-not $preferred) { $preferred = $win11 | Where-Object { $_.URL -match "w11_22|win11_22" } | Select-Object -Last 1 }
        if (-not $preferred) { $preferred = $win11 | Select-Object -Last 1 }
        if (-not $preferred) {
            $win10 = $urlOsPairs | Where-Object { $_.OS -match "Windows 10" -and $_.URL -notmatch "readme" }
            $preferred = $win10 | Where-Object { $_.URL -match "22h2|21h2|21h1" } | Select-Object -Last 1
            if (-not $preferred) { $preferred = $win10 | Select-Object -Last 1 }
        }
        if (-not $preferred) { $preferred = $urlOsPairs | Where-Object { $_.URL -notmatch "readme" } | Select-Object -Last 1 }
        if ($preferred) { Log "Selected SCCM URL [$($preferred.OS)]: $($preferred.URL)"; return $preferred.URL }
    } else {
        $preferred = $urlOsPairs | Where-Object { $_.URL -match "\.exe$" -and ($_.URL -match "bios|uefi") -and $_.URL -notmatch "readme" } | Select-Object -First 1
        if (-not $preferred) { $preferred = $urlOsPairs | Where-Object { $_.URL -match "\.exe$" -and $_.URL -notmatch "readme" } | Select-Object -First 1 }
        if ($preferred) { Log "Selected BIOS URL: $($preferred.URL)"; return $preferred.URL }
    }
    return $null
}
# =========================
# LENOVO: FULL AUTO-DOWNLOAD FLOW
# =========================
function Start-LenovoDriverDownload {
    param(
        [string]$MachineTypePrefix,
        [string]$TargetOS = "Windows 11",
        [string]$ProductFamily = "ThinkPad",
        [string]$DownloadPath = "C:\DriverPacks",
        [string]$ExtractionPath = "C:\ExtractedDrivers"
    )

    Log "--- Lenovo Auto-Download: $ProductFamily ($MachineTypePrefix / $TargetOS) ---"

    # ---- Fetch recipecard.json ----
    $RecipeJsonUrl = "https://download.lenovo.com/cdrt/ddrc/recipecard.json"
    Log "Fetching Lenovo recipecard.json..."
        SetStage "Fetching Lenovo driver catalogue..."
    try {
        $Global:ProgressPreference = 'SilentlyContinue'
        $jsonResponse = Invoke-WebRequest -Uri $RecipeJsonUrl -UseBasicParsing -TimeoutSec 60
        $Global:ProgressPreference = 'Continue'
        $jsonData = $jsonResponse.Content | ConvertFrom-Json
        Log "recipecard.json loaded OK."
    } catch {
        Log "Failed to fetch recipecard.json: $($_.Exception.Message)"
        return $false
    }

    # ---- Find model ID ----
    $modelFamilyData = $jsonData.$ProductFamily
    if (-not $modelFamilyData) {
        Log "Product family '$ProductFamily' not found in recipecard.json."
        return $false
    }

    $modelId = $null; $modelName = $null; $modelEntry = $null
    foreach ($entry in $modelFamilyData) {
        if ($entry.types -is [array]) {
            foreach ($t in $entry.types) {
                if ($t -like "$($MachineTypePrefix)*") {
                    $modelId    = $entry.id
                    $modelName  = $entry.name
                    $modelEntry = $entry   # preserve full entry for SMBIOS lookup
                    break
                }
            }
        }
        if ($modelId) { break }
    }

    if (-not $modelId) {
        Log "Machine type '$MachineTypePrefix' not found in '$ProductFamily'. Trying other families..."
        foreach ($family in @("ThinkCentre","ThinkStation","IdeaPad","IdeaCentre")) {
            $familyData = $jsonData.$family
            if (-not $familyData) { continue }
            foreach ($entry in $familyData) {
                if ($entry.types -is [array]) {
                    foreach ($t in $entry.types) {
                        if ($t -like "$($MachineTypePrefix)*") {
                            $modelId      = $entry.id
                            $modelName    = $entry.name
                            $ProductFamily = $family
                            Log "Found in family '$family': $modelName"
                            break
                        }
                    }
                }
                if ($modelId) { break }
            }
            if ($modelId) { break }
        }
    }

    if (-not $modelId) {
        Log "Machine type '$MachineTypePrefix' not found in any Lenovo product family."
        return $false
    }
    Log "Model matched: $modelName (ID: $modelId)"

    # All IDs in recipecard.json are strings — cast everything for safe comparison
    $modelIdStr = [string]$modelId

    # ---- Find OS ID ----
    $foundOS = $jsonData.OperatingSystems | Where-Object { $_.name -eq $TargetOS }
    if ($foundOS -is [array]) { $foundOS = $foundOS[0] }
    $osIdStr = if ($foundOS) { [string]$foundOS.id } else { $null }

    # ---- Find RecipeCard (with fallback to most recent available OS) ----
    $recipeCard = $null

    if ($osIdStr) {
        $recipeCard = $jsonData.RecipeCards | Where-Object { [string]$_.modelId -eq $modelIdStr -and [string]$_.osId -eq $osIdStr }
        if ($recipeCard -is [array]) { $recipeCard = $recipeCard[0] }
    }

    if (-not $recipeCard) {
        if ($osIdStr) {
            Log "No RecipeCard for '$modelName' + '$TargetOS'. Falling back to most recent available OS..."
        } else {
            Log "OS '$TargetOS' not found in recipecard.json. Finding most recent available OS for this model..."
        }

        # Get all RecipeCards for this model, sort by osId descending — higher ID = newer OS
        $allModelCards = $jsonData.RecipeCards | Where-Object { [string]$_.modelId -eq $modelIdStr }
        if ($allModelCards) {
            $bestCard   = $allModelCards | Sort-Object { [int][string]$_.osId } -Descending | Select-Object -First 1
            $fallbackOS = $jsonData.OperatingSystems | Where-Object { [string]$_.id -eq [string]$bestCard.osId }
            if ($fallbackOS -is [array]) { $fallbackOS = $fallbackOS[0] }
            $recipeCard = $bestCard
            $osIdStr    = [string]$bestCard.osId
            Log "Falling back to: $($fallbackOS.name) (OS ID: $osIdStr)"
        }
    } else {
        Log "OS matched: $TargetOS (ID: $osIdStr)"
    }

    if (-not $recipeCard) {
        Log "No RecipeCard found for model '$modelName' under any OS."
        return $false
    }

    $sccmPackIdRef = [string]$recipeCard.sccmPack
    $recipeNote    = $recipeCard.note
    Log "SCCM Pack Ref: $sccmPackIdRef"

    # ---- Extract BIOS support page from note ----
    $biosPageUrl = $null
    if ($recipeNote -match '\[\[.*?\|(https?://[^\]]+)\]\]') {
        $biosPageUrl = $matches[1]
        Log "BIOS support page found in note: $biosPageUrl"
    }

    # ---- Find SCCM pack support page ----
    $foundSccmPack = $jsonData.SCCMPacks | Where-Object { [string]$_.id -eq $sccmPackIdRef }
    if (-not $foundSccmPack) {
        Log "SCCM pack details not found for ref '$sccmPackIdRef'."
        return $false
    }
    if ($foundSccmPack -is [array]) { $foundSccmPack = $foundSccmPack[0] }
    $sccmSupportPageUrl = $foundSccmPack.url
    Log "SCCM support page: $sccmSupportPageUrl"

    # ---- Get direct download URLs ----
    # Primary: Thin Installer XML catalog (download.lenovo.com CDN — not blocked like support pages)
    $smbiosCodes = @()
    if ($modelEntry.smbios) { $smbiosCodes = @($modelEntry.smbios) }

    $driverPackUrl = $null
    $biosFileUrl   = $null

    if ($smbiosCodes.Count -gt 0 -or $MachineTypePrefix) {
        Log "Catalog lookup for machine type: ${MachineTypePrefix}"
        $driverPackUrl = Get-LenovoSccmUrlFromCatalog -SmbiosCodes $smbiosCodes -TargetOSName $TargetOS -MachineType $MachineTypePrefix
        $biosFileUrl   = Get-LenovoBiosUrlFromCatalog  -SmbiosCodes $smbiosCodes -TargetOSName $TargetOS -MachineType $MachineTypePrefix
    }

    # Fallback: scrape the support page HTML (may be blocked on some networks)
    if (-not $driverPackUrl) {
        Log "Catalog lookup failed or no SMBIOS — trying support page scrape..."
        $driverPackUrl = Get-LenovoDirectDownloadLink -SupportPageUrl $sccmSupportPageUrl -ModelName $modelName -TargetOSName $TargetOS -LinkType "SCCM"
    }
    if (-not $biosFileUrl -and $biosPageUrl) {
        $rawBios = Get-LenovoDirectDownloadLink -SupportPageUrl $biosPageUrl -ModelName $modelName -TargetOSName $TargetOS -LinkType "BIOS"
        if ($rawBios -is [array]) { $biosFileUrl = $rawBios | Where-Object { $_ -is [string] -and $_ -match "^https?://" } | Select-Object -First 1 }
        else { $biosFileUrl = $rawBios }
    }

    # ---- Validate SCCM URL ----
    # Unwrap array if function returned ($true, "url") due to pipeline quirk
    if ($driverPackUrl -is [array]) {
        $driverPackUrl = $driverPackUrl | Where-Object { $_ -is [string] -and $_ -match "^https?://" } | Select-Object -First 1
    }
    Log "DEBUG: SCCM URL resolved: $driverPackUrl"
    if (-not $driverPackUrl -or $driverPackUrl -notmatch "^https?://") {
        Log "Could not automatically find SCCM driver pack URL."
        Log "Please visit manually: $sccmSupportPageUrl"
        Start-Process $sccmSupportPageUrl
        return $false
    }

    # ---- Create directories ----
    foreach ($dir in @($DownloadPath, $ExtractionPath)) {
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    }

    # ---- Download SCCM driver pack ----
    $packFileName = try { [System.IO.Path]::GetFileName(([System.Uri]$driverPackUrl).LocalPath) } catch { "sccm_driverpack.exe" }
    $packFilePath = Join-Path $DownloadPath $packFileName
    Log "Downloading SCCM pack: $packFileName"
    $progress.Value = 10

    try {
        Log "Download starting — this may take several minutes for a large pack..."
        SetStage "Downloading driver pack..."
        $progress.Value = 10

        # Use a temp stderr file so we can read curl's progress without blocking
        $dlErrFile = [System.IO.Path]::GetTempFileName()

        $curlDlArgs = "--location --fail --max-time 600 --connect-timeout 30 " +
                      "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36`" " +
                      "--output `"$packFilePath`" " +
                      "`"$driverPackUrl`""

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "curl.exe"
        $psi.Arguments = $curlDlArgs
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $dlProc = New-Object System.Diagnostics.Process
        $dlProc.StartInfo = $psi
        $dlProc.Start() | Out-Null

        $lastSize = 0; $stall = 0; $totalMB = 0
        while (-not $dlProc.HasExited) {
            Start-Sleep -Milliseconds 800

            $fileSize = if (Test-Path $packFilePath) { (Get-Item $packFilePath -ErrorAction SilentlyContinue).Length } else { 0 }
            if (-not $fileSize) { $fileSize = 0 }

            if ($fileSize -gt $lastSize) {
                $stall = 0
                # Try to infer total size from a HEAD request result already in memory,
                # or derive progress from known pack sizes (~1.9 GB typical)
                if ($totalMB -eq 0 -and $fileSize -gt 10MB) { $totalMB = 1950 }   # typical Lenovo SCCM pack
                $lastSize = $fileSize
            } else { $stall++ }

            $mbDone = [math]::Round($fileSize / 1MB, 1)

            if ($totalMB -gt 0) {
                # Real percentage: map download progress into 10-39% of overall bar
                $pct = [math]::Min(($fileSize / ($totalMB * 1MB)), 1.0)
                $progress.Value = 10 + [int]($pct * 29)
                $pctDisp = [math]::Round($pct * 100, 0)
                Log "Downloading... $mbDone MB / ~$totalMB MB ($pctDisp%)"
            } else {
                # No size known yet — pulse
                $cur = $progress.Value
                $progress.Value = if ($cur -ge 18) { 10 } else { $cur + 1 }
                Log "Downloading... $mbDone MB received"
            }

            [System.Windows.Forms.Application]::DoEvents()
            if ($stall -gt 75) { Log "WARNING: Download stalled 60s — aborting."; $dlProc.Kill(); break }
        }

        if ($dlProc.ExitCode -ne 0) {
            Log "curl.exe download failed (exit code $($dlProc.ExitCode))"
            return $false
        }
        if (-not (Test-Path $packFilePath) -or (Get-Item $packFilePath).Length -eq 0) {
            Log "Download finished but file is missing or empty."
            return $false
        }
        $finalMB = [math]::Round((Get-Item $packFilePath).Length / 1MB, 1)
        Log "Download complete: $packFilePath ($finalMB MB)"
        SetStage "Download complete — extracting..."
    } catch {
        Log "SCCM download exception: $($_.Exception.Message)"
        return $false
    }

    $progress.Value = 40

    # ---- Extract SCCM driver pack ----
    Log "Extracting drivers to: $ExtractionPath"
    SetStage "Extracting driver pack..."
    try {
        if ($packFileName -match "\.zip$") {
            Expand-Archive -Path $packFilePath -DestinationPath $ExtractionPath -Force -ErrorAction Stop
            Log "ZIP extraction complete."
        } elseif ($packFileName -match "\.exe$") {
            $extractArgs = "/VERYSILENT /DIR=`"$($ExtractionPath.TrimEnd('\'))`" /EXTRACT=YES"
            $proc = Start-Process -FilePath $packFilePath -ArgumentList $extractArgs -Wait -PassThru -ErrorAction Stop
            if ($proc.ExitCode -ne 0) { Log "EXE extractor exited with code $($proc.ExitCode) (may still be OK)." }
            else { Log "EXE extraction complete." }
            Start-Sleep -Seconds 5
        } elseif ($packFileName -match "\.cab$") {
            $expandArgs = "`"$packFilePath`" -F:* `"$ExtractionPath`""
            Start-Process -FilePath "expand.exe" -ArgumentList $expandArgs -Wait -NoNewWindow -ErrorAction Stop
            Log "CAB extraction complete."
        } else {
            Log "Unsupported pack file type: $packFileName"
        }
    } catch {
        Log "Extraction error: $($_.Exception.Message)"
        return $false
    }

    $progress.Value = 60

    # ---- Download BIOS (optional) ----
    if ($biosFileUrl -and $biosFileUrl -match "^https?://") {
        $biosFolder   = Join-Path $DownloadPath "BIOS"
        if (-not (Test-Path $biosFolder)) { New-Item -Path $biosFolder -ItemType Directory -Force | Out-Null }
        $biosFileName = try { [System.IO.Path]::GetFileName(([System.Uri]$biosFileUrl).LocalPath) } catch { "bios_update.exe" }
        $biosFilePath = Join-Path $biosFolder $biosFileName
        Log "Downloading BIOS update: $biosFileName"
        try {
            Log "Downloading BIOS update..."
            $biosCurlArgs = "--location --show-error --fail --max-time 300 --connect-timeout 30 " +
                            "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36`" " +
                            "--output `"$biosFilePath`" `"$biosFileUrl`""
            $biosProc = Start-Process -FilePath "curl.exe" -ArgumentList $biosCurlArgs -Wait -PassThru -NoNewWindow -RedirectStandardError "$env:TEMP\bios_curl_err.txt"
            if ($biosProc.ExitCode -ne 0) {
                $biosErr = if (Test-Path "$env:TEMP\bios_curl_err.txt") { Get-Content "$env:TEMP\bios_curl_err.txt" -Raw } else { "no stderr" }
                Log "BIOS download failed (exit $($biosProc.ExitCode)): $biosErr"
            } else {
                Log "BIOS downloaded to: $biosFilePath"
                Log "NOTE: BIOS flashing must be done manually — do NOT run automatically."
            }
        } catch {
            Log "BIOS download exception: $($_.Exception.Message)"
        }
    } elseif ($biosPageUrl) {
        Log "BIOS page found but no direct link resolved. Visit manually: $biosPageUrl"
    }

    # ---- Return extraction path so the INF installer loop can pick up from here ----
    return $ExtractionPath
}

# =========================
# LENOVO: DETECT PRODUCT FAMILY
# =========================
function Get-LenovoProductFamily($model) {
    if ($model -match "ThinkPad")    { return "ThinkPad" }
    if ($model -match "ThinkCentre") { return "ThinkCentre" }
    if ($model -match "ThinkStation") { return "ThinkStation" }
    if ($model -match "IdeaPad")     { return "IdeaPad" }
    if ($model -match "IdeaCentre")  { return "IdeaCentre" }
    return "ThinkPad"   # safest default for business fleet
}

# =========================
# MAIN INSTALL FUNCTION
# =========================
function Start-Install {

    $button.Enabled = $false
    $progress.Value = 0

    # Admin check
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }

    Log "Driver Installer Tool v${ScriptVersion}"
    Log "Starting driver installation..."
    Log "Log file: $LogFile"

    $manufacturer = (Get-CimInstance Win32_ComputerSystem).Manufacturer
    $model        = Copy-ModelToClipboard

    Log "Manufacturer: $manufacturer"
    Log "Model: $model (copied to clipboard)"
    $title.Text = "Driver Installer - $model"

    # =========================
    # DRIVER PATHS
    # =========================
    $paths     = @()
    $driverURL = Get-DriverLink $manufacturer

    if ($manufacturer -match "Dell") {
        $paths += "C:\Users\Administrator\"
    } elseif ($manufacturer -match "HP") {
        $paths += "C:\SWSetup\"
    } elseif ($manufacturer -match "Lenovo") {
        $paths += "C:\DRIVERS\", "C:\ExtractedDrivers\"
    } else {
        $paths += "C:\Drivers\", "C:\SWSetup\", "C:\DRIVERS\", "C:\Users\Administrator\"
    }

    # =========================
    # FIND DRIVER FOLDERS (INF scan)
    # =========================
    $validPaths  = @()
    $driverFound = $false

    foreach ($base in $paths) {
        if (Test-Path $base) {
            $dirs = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue
            foreach ($dir in $dirs) {
                if (Get-ChildItem $dir.FullName -Recurse -Filter *.inf -ErrorAction SilentlyContinue) {
                    $validPaths += $dir.FullName
                    $driverFound = $true
                }
            }
        }
    }

    # =========================
    # LENOVO: AUTO-DOWNLOAD IF NO LOCAL DRIVERS
    # =========================
    if (-not $driverFound -and $manufacturer -match "Lenovo") {

        Log "No local Lenovo drivers found. Attempting auto-download from Lenovo recipecard..."

        $machineTypePrefix = Get-LenovoMachineTypePrefix
        if (-not $machineTypePrefix) {
            Log "Could not detect machine type. Falling back to OEM page."
        } else {
            $productFamily  = Get-LenovoProductFamily $model
            $extractionPath = Start-LenovoDriverDownload `
                -MachineTypePrefix $machineTypePrefix `
                -TargetOS "Windows 11" `
                -ProductFamily $productFamily `
                -DownloadPath "C:\DriverPacks" `
                -ExtractionPath "C:\ExtractedDrivers"

            if ($extractionPath -and (Test-Path $extractionPath)) {
                # Re-scan extraction path for INFs
                $dirs = Get-ChildItem $extractionPath -Directory -ErrorAction SilentlyContinue
                foreach ($dir in $dirs) {
                    if (Get-ChildItem $dir.FullName -Recurse -Filter *.inf -ErrorAction SilentlyContinue) {
                        $validPaths += $dir.FullName
                        $driverFound = $true
                    }
                }
                # Also check extraction root itself
                if (Get-ChildItem $extractionPath -Recurse -Filter *.inf -ErrorAction SilentlyContinue) {
                    if ($validPaths -notcontains $extractionPath) {
                        $validPaths += $extractionPath
                        $driverFound = $true
                    }
                }
            }
        }
    }

    # =========================
    # PRIORITY 1: DRIVERS FOUND — INSTALL INFs
    # =========================
    if ($driverFound) {

        $allDrivers = @()
        foreach ($path in $validPaths) {
            $allDrivers += Get-ChildItem $path -Recurse -Filter *.inf -ErrorAction SilentlyContinue
        }

        $total   = $allDrivers.Count
        $current = 0

        if ($total -eq 0) {
            Log "No driver files found after scan."
            $button.Enabled = $true
            return
        }

        Log "Found $total driver file(s). Starting installation..."
        SetStage "Installing drivers..."

        foreach ($driver in $allDrivers) {
            $current++
            $percent        = [math]::Round(($current / $total) * 100)
            $progress.Value = [math]::Min($percent, 99)   # keep 100 for completion

            Log "[$current/$total] Installing: $($driver.Name)"
            $output = pnputil /add-driver "`"$($driver.FullName)`"" /install 2>&1
            foreach ($line in $output) { Log $line }
            [System.Windows.Forms.Application]::DoEvents()
        }

        $progress.Value = 100
        Log "Installation complete!"
        SetStage "Done!"

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Drivers installed for $model.`nReboot now?",
            "Complete",
            "YesNo"
        )
        if ($result -eq "Yes") { Restart-Computer -Force }
        else { $button.Enabled = $true }
        return
    }

    # =========================
    # PRIORITY 2: DOWNLOADS EXE
    # =========================
    Log "No drivers found. Checking Downloads folder..."
    $exeFound = Check-DownloadsForExe
    if ($exeFound) {
        Log "Installer launched from Downloads."
        $button.Enabled = $true
        return
    }

    # =========================
    # PRIORITY 3: OEM PAGE
    # =========================
    Log "No drivers or installers found. Opening OEM support page..."
    if ($driverURL) { Start-Process $driverURL }

    [System.Windows.Forms.MessageBox]::Show(
        "No drivers or installers found.`nModel: $model",
        "Nothing Found"
    )
    $button.Enabled = $true
}

# =========================
# BUTTON
# =========================
$button.Add_Click({ Start-Install })

# =========================
# AUTO START
# =========================
$form.Add_Shown({
    $form.Activate()
    Start-Sleep -Milliseconds 300
    Log "Running startup checks..."
    Start-Install
})

# =========================
# RUN
# =========================
[void]$form.ShowDialog()