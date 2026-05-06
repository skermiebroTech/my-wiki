Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Disable sleep and display timeout (AC power only)
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# =========================
# FORM SETUP
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool"
$form.Size = New-Object System.Drawing.Size(520,340)
$form.StartPosition = "CenterScreen"

# Title
$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,15)
$form.Controls.Add($title)

# Status box
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.Size = New-Object System.Drawing.Size(460,140)
$statusBox.Location = New-Object System.Drawing.Point(20,60)
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

# Progress bar
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(460,20)
$progress.Location = New-Object System.Drawing.Point(20,210)
$form.Controls.Add($progress)

# Button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Install Drivers"
$button.Size = New-Object System.Drawing.Size(160,35)
$button.Location = New-Object System.Drawing.Point(180,245)
$form.Controls.Add($button)

# =========================
# LOG FUNCTION
# =========================
function Log($msg) {
    $statusBox.AppendText("$msg`r`n")
    $statusBox.ScrollToCaret()
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
# MAIN INSTALL FUNCTION
# =========================
function Start-Install {

    $button.Enabled = $false
    $progress.Value = 0

    # Admin check
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }

    Log "Starting driver installation..."

    $manufacturer = (Get-CimInstance Win32_ComputerSystem).Manufacturer
    $model = Copy-ModelToClipboard

    Log "Manufacturer: $manufacturer"
    Log "Model: $model (copied to clipboard)"

    $title.Text = "Driver Installer - $model"

    # =========================
    # DRIVER PATHS
    # =========================
    $paths = @()
    $driverURL = Get-DriverLink $manufacturer

    if ($manufacturer -match "Dell") {
        $paths += "C:\Users\Administrator\"
    }
    elseif ($manufacturer -match "HP") {
        $paths += "C:\SWSetup\"
    }
    elseif ($manufacturer -match "Lenovo") {
        $paths += "C:\DRIVERS\"
    }
    else {
        $paths += "C:\Drivers\","C:\SWSetup\","C:\DRIVERS\","C:\Users\Administrator\"
    }

    # =========================
    # FIND DRIVER FOLDERS
    # =========================
    $validPaths = @()
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
    # PRIORITY 1: DRIVERS FOUND
    # =========================
    if ($driverFound) {

        # Collect all INF files
        $allDrivers = @()

        foreach ($path in $validPaths) {
            $drivers = Get-ChildItem $path -Recurse -Filter *.inf -ErrorAction SilentlyContinue
            $allDrivers += $drivers
        }

        $total = $allDrivers.Count
        $current = 0

        if ($total -eq 0) {
            Log "No driver files found."
            return
        }

        Log "Found $total driver files. Starting install..."

        foreach ($driver in $allDrivers) {

            $current++
            $percent = [math]::Round(($current / $total) * 100)
            $progress.Value = $percent

            Log "[$current/$total] Installing: $($driver.Name)"

            $output = pnputil /add-driver "`"$($driver.FullName)`"" /install 2>&1

            foreach ($line in $output) {
                Log $line
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $progress.Value = 100
        Log "Installation complete!"

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Drivers installed for $model.`nReboot now?",
            "Complete",
            "YesNo"
        )

        if ($result -eq "Yes") {
            Restart-Computer -Force
        } else {
            $button.Enabled = $true
        }

        return
    }

    # =========================
    # PRIORITY 2: DOWNLOADS EXE
    # =========================
    Log "No drivers found. Checking Downloads..."

    $exeFound = Check-DownloadsForExe

    if ($exeFound) {
        Log "Installer launched from Downloads."
        $button.Enabled = $true
        return
    }

    # =========================
    # PRIORITY 3: OEM PAGE
    # =========================
    Log "No drivers or installers found. Opening OEM page..."

    if ($driverURL) {
        Start-Process $driverURL
    }

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