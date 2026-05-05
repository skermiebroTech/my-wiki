Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# ADMIN CHECK
# =========================
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    $scriptUrl = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers.ps1"

    Start-Process powershell "-ExecutionPolicy Bypass -Command `"irm $scriptUrl | iex`"" -Verb RunAs
    exit
}

# =========================
# PREVENT SLEEP
# =========================
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# =========================
# LENOVO FUNCTIONS
# =========================
function Get-LenovoMachineType {
    try {
        $sku = (Get-CimInstance Win32_ComputerSystemProduct).Name
        if ($sku) { return $sku.Substring(0,4).ToUpper() }
    } catch {}
    return $null
}

function Get-LenovoDriverPackUrl($machineType, $os) {
    try {
        $json = Invoke-WebRequest "https://download.lenovo.com/cdrt/ddrc/recipecard.json" -UseBasicParsing | ConvertFrom-Json

        $family = $json.ThinkPad
        if (-not $family) { return $null }

        $model = $family | Where-Object { $_.types -match "^$machineType" } | Select-Object -First 1
        if (-not $model) { return $null }

        $osId = ($json.OperatingSystems | Where-Object name -eq $os).id

        $recipe = $json.RecipeCards | Where-Object {
            $_.modelId -eq $model.id -and $_.osId -eq $osId
        }

        if (-not $recipe) { return $null }

        return ($json.SCCMPacks | Where-Object id -eq $recipe.sccmPack).url
    } catch {
        return $null
    }
}

# =========================
# FORM
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool"
$form.Size = New-Object System.Drawing.Size(520,340)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,15)
$title.Text = "Driver Installer Ready"
$form.Controls.Add($title)

# =========================
# STATUS BOX (UPDATED TEXT)
# =========================
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.Size = New-Object System.Drawing.Size(460,140)
$statusBox.Location = New-Object System.Drawing.Point(20,60)
$statusBox.ReadOnly = $true
$statusBox.Text = "Waiting for installation to begin...`r`nThis tool will automatically detect drivers, download packages, and install them."
$form.Controls.Add($statusBox)

# =========================
# PROGRESS BAR
# =========================
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(460,20)
$progress.Location = New-Object System.Drawing.Point(20,210)
$form.Controls.Add($progress)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Install Drivers"
$button.Size = New-Object System.Drawing.Size(160,35)
$button.Location = New-Object System.Drawing.Point(180,245)
$form.Controls.Add($button)

# =========================
# LOG
# =========================
function Log($msg) {
    $statusBox.AppendText("$msg`r`n")
    $statusBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# =========================
# MAIN
# =========================
function Start-Install {

    $button.Enabled = $false
    $progress.Value = 0

    Log "Starting driver installation..."
    Log "Scanning system..."

    $manufacturer = (Get-CimInstance Win32-ComputerSystem).Manufacturer
    $model = (Get-CimInstance Win32-ComputerSystem).Model

    Log "Device: $manufacturer $model"

    # =========================
    # FIND LOCAL DRIVERS
    # =========================
    $paths = @("C:\Drivers","C:\SWSetup","C:\DRIVERS","C:\Users\Administrator\")
    $allDrivers = @()

    foreach ($base in $paths) {
        if (Test-Path $base) {
            $allDrivers += Get-ChildItem $base -Recurse -Filter *.inf -ErrorAction SilentlyContinue
        }
    }

    # =========================
    # INSTALL LOCAL DRIVERS
    # =========================
    if ($allDrivers.Count -gt 0) {

        Log "Found $($allDrivers.Count) local drivers."
        $i = 0

        foreach ($driver in $allDrivers) {
            $i++

            $progress.Value = [math]::Round(($i / $allDrivers.Count) * 100)
            Log "Installing: $($driver.Name)"

            pnputil /add-driver "$($driver.FullName)" /install | Out-Null

            [System.Windows.Forms.Application]::DoEvents()
        }

        Log "Local driver installation complete."
        $progress.Value = 100
        return
    }

    # =========================
    # LENOVO AUTO DRIVER DOWNLOAD
    # =========================
    if ($manufacturer -match "Lenovo") {

        Log "Lenovo detected. Checking online driver packs..."
        $progress.Value = 10
        [System.Windows.Forms.Application]::DoEvents()

        $machineType = Get-LenovoMachineType
        Log "Machine Type: $machineType"

        if ($machineType) {

            $url = Get-LenovoDriverPackUrl $machineType "Windows 11"

            if ($url) {

                Log "Downloading Lenovo driver pack..."
                Log $url

                $file = "$env:TEMP\lenovo_driverpack.exe"

                Invoke-WebRequest $url -OutFile $file

                Log "Download complete. Running installer..."
                $progress.Value = 60
                [System.Windows.Forms.Application]::DoEvents()

                Start-Process $file -Wait

                $progress.Value = 100
                Log "Lenovo driver installation complete."

                return
            }

            Log "No Lenovo driver pack found."
        }
    }

    # =========================
    # FALLBACK
    # =========================
    Log "Checking Downloads..."
    $progress.Value = 80
    [System.Windows.Forms.Application]::DoEvents()

    $downloads = "$env:USERPROFILE\Downloads"

    $exe = Get-ChildItem $downloads -Filter *.exe -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

    if ($exe) {
        Log "Running installer: $($exe.Name)"
        Start-Process $exe.FullName
        return
    }

    Log "Opening OEM page..."
    Start-Process (Get-DriverLink $manufacturer)

    $progress.Value = 100
}

$button.Add_Click({ Start-Install })

$form.Add_Shown({
    $form.Activate()
    Log "Ready. Click Install to begin."
})

[void]$form.ShowDialog()