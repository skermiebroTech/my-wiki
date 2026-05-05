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
$form.Text = "Driver Installer Tool"
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
# LENOVO: FIND DIRECT DOWNLOAD LINK FROM SUPPORT PAGE
# =========================
function Get-LenovoDirectDownloadLink {
    param(
        [string]$SupportPageUrl,
        [string]$ModelName,
        [string]$TargetOSName,
        [string]$LinkType = "SCCM"
    )

    # Extract the docId from the support page URL (e.g. ds549249 from .../downloads/ds549249)
    $docId = $null
    if ($SupportPageUrl -match '/downloads/(ds\d+)') {
        $docId = $Matches[1].ToUpper()
    }
    if (-not $docId) {
        # Try last path segment as fallback
        $seg = ($SupportPageUrl.TrimEnd('/') -split '/')[-1]
        if ($seg -match '^ds\d+') { $docId = $seg.ToUpper() }
    }
    if (-not $docId) {
        Log "Could not extract docId from URL: $SupportPageUrl"
        return $null
    }

    Log "Querying Lenovo driver API for ${docIdLower}..."
    SetStage "Querying Lenovo driver API (${docIdLower})..."

    # Try multiple API endpoint variants — Lenovo's API is region-sensitive
    # Lenovo API requires lowercase docId - uppercase returns body:null
    $docIdLower = $docId.ToLower()
    $apiUrls = @(
        "https://pcsupport.lenovo.com/us/en/api/v4/downloads/drivers?docId=${docIdLower}",
        "https://pcsupport.lenovo.com/au/en/api/v4/downloads/drivers?docId=${docIdLower}",
        "https://pcsupport.lenovo.com/gb/en/api/v4/downloads/drivers?docId=${docIdLower}",
        "https://pcsupport.lenovo.com/jp/en/api/v4/downloads/drivers?docId=${docIdLower}"
    )

    $jsonText = $null
    $headers = @{
        "Accept"          = "application/json, text/plain, */*"
        "Accept-Language" = "en-US,en;q=0.9"
        "Referer"         = "https://pcsupport.lenovo.com/"
        "User-Agent"      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    foreach ($apiUrl in $apiUrls) {
        Log "Trying: $apiUrl"
        try {
            $Global:ProgressPreference = 'SilentlyContinue'
            $resp = Invoke-WebRequest -Uri $apiUrl -Headers $headers -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            $Global:ProgressPreference = 'Continue'
            [System.Windows.Forms.Application]::DoEvents()

            if ($resp.StatusCode -eq 200 -and $resp.Content.Length -gt 20) {
                $candidate = $resp.Content
                # Check the response actually has file data (not just a null body error)
                if ($candidate -notmatch '"body"\s*:\s*null') {
                    $jsonText = $candidate
                    Log "Got valid response from: ${apiUrl}"
                    break
                } else {
                    Log "API returned null body from: ${apiUrl}"
                }
            }
        } catch {
            Log "API attempt failed (${apiUrl}): $($_.Exception.Message)"
            $Global:ProgressPreference = 'Continue'
        }
        $cur = $progress.Value
        $progress.Value = if ($cur -ge 9) { 1 } else { $cur + 1 }
        [System.Windows.Forms.Application]::DoEvents()
    }

    if (-not $jsonText -or $jsonText.Length -lt 10) {
        Log "All API endpoints failed for ${docIdLower} — trying direct URL regex from support page..."
        # Emergency fallback: use known Lenovo download URL pattern from recipecard version
        return $null
    }

    try {
        $apiData = $jsonText | ConvertFrom-Json
    } catch {
        Log "Failed to parse API JSON for ${docIdLower}: $($_.Exception.Message)"
        $urlPat = '"URL"\s*:\s*"(https?://download\.lenovo\.com/[^"]+\.(?:exe|zip|cab))"'
        $m = [regex]::Match($jsonText, $urlPat)
        if ($m.Success) { $url = $m.Groups[1].Value; Log "Regex fallback URL: ${url}"; return $url }
        return $null
    }

    # Navigate the response — Lenovo API returns body.DriverDetails.Files[]
    $files = $null
    try { $files = $apiData.body.DriverDetails.Files } catch {}
    if (-not $files) { try { $files = $apiData.Files } catch {} }
    if (-not $files -or $files.Count -eq 0) {
        Log "No files in API response for ${docIdLower}. Raw: $($jsonText.Substring(0, [Math]::Min(300,$jsonText.Length)))"
        return $null
    }

    Log "API returned $($files.Count) file(s) for $docId."

    if ($LinkType -eq "SCCM") {
        # Filter to EXE files only, excluding READMEs
        $exeFiles = $files | Where-Object {
            $_.TypeString -match "EXE" -and
            $_.URL -notmatch "readme|\.txt$|\.html$" -and
            $_.URL -match "^https?://"
        }

        Log "EXE candidates:"
        $exeFiles | ForEach-Object {
            $osLabel = if ($_.OperatingSystemKeys) { $_.OperatingSystemKeys -join ", " } else { "unknown" }
            Log "  [$osLabel] $($_.URL)"
        }

        # Priority: Win11 newest version first
        $win11 = $exeFiles | Where-Object { $_.OperatingSystemKeys -match "Windows 11" }
        $preferred = $win11 | Where-Object { $_.URL -match "w11_24|win11_24" } | Select-Object -Last 1
        if (-not $preferred) { $preferred = $win11 | Where-Object { $_.URL -match "w11_22|win11_22" } | Select-Object -Last 1 }
        if (-not $preferred) { $preferred = $win11 | Where-Object { $_.URL -match "w11_21|win11_21" } | Select-Object -Last 1 }
        if (-not $preferred) { $preferred = $win11 | Select-Object -Last 1 }

        # Fallback: newest Win10 version
        if (-not $preferred) {
            Log "No Windows 11 SCCM pack found — falling back to Windows 10..."
            $win10 = $exeFiles | Where-Object { $_.OperatingSystemKeys -match "Windows 10" }
            $preferred = $win10 | Where-Object { $_.URL -match "22h2|21h2|21h1|20h2" } | Select-Object -Last 1
            if (-not $preferred) { $preferred = $win10 | Select-Object -Last 1 }
        }

        # Last resort: any EXE
        if (-not $preferred) { $preferred = $exeFiles | Select-Object -Last 1 }

        if ($preferred) {
            $osLabel = if ($preferred.OperatingSystemKeys) { $preferred.OperatingSystemKeys -join ", " } else { "" }
            Log "Selected SCCM URL [$osLabel]: $($preferred.URL)"
            return $preferred.URL
        }

        Log "Could not select a preferred SCCM URL."
        return $null
    }

    # BIOS: pick the first EXE that looks like a BIOS update
    $biosFile = $files | Where-Object {
        $_.TypeString -match "EXE" -and
        ($_.Name -match "bios|uefi|firmware" -or $_.URL -match "bios|uefi") -and
        $_.URL -notmatch "readme"
    } | Select-Object -First 1

    if (-not $biosFile) {
        $biosFile = $files | Where-Object { $_.TypeString -match "EXE" -and $_.URL -notmatch "readme" } | Select-Object -First 1
    }

    if ($biosFile) {
        Log "Selected BIOS URL: $($biosFile.URL)"
        return $biosFile.URL
    }

    Log "Could not find a BIOS download URL."
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

    $modelId = $null; $modelName = $null
    foreach ($entry in $modelFamilyData) {
        if ($entry.types -is [array]) {
            foreach ($t in $entry.types) {
                if ($t -like "$($MachineTypePrefix)*") {
                    $modelId   = $entry.id
                    $modelName = $entry.name
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
    $driverPackUrl = Get-LenovoDirectDownloadLink -SupportPageUrl $sccmSupportPageUrl -ModelName $modelName -TargetOSName $TargetOS -LinkType "SCCM"
    $biosFileUrl   = if ($biosPageUrl) {
        $rawBios = Get-LenovoDirectDownloadLink -SupportPageUrl $biosPageUrl -ModelName $modelName -TargetOSName $TargetOS -LinkType "BIOS"
        if ($rawBios -is [array]) { $rawBios | Where-Object { $_ -is [string] -and $_ -match "^https?://" } | Select-Object -First 1 }
        else { $rawBios }
    } else { $null }

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