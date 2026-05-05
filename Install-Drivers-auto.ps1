Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# ADMIN CHECK (EARLY)
# =========================
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    $scriptUrl = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers.ps1"

    Start-Process powershell "-ExecutionPolicy Bypass -Command `"irm $scriptUrl | iex`"" -Verb RunAs
    exit
}

# =========================
# PREVENT SLEEP (AC ONLY)
# =========================
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# =========================
# LENOVO FUNCTIONS (NEW)
# =========================
function Get-LenovoMachineType {
    try {
        $sku = (Get-CimInstance Win32_ComputerSystemProduct).Name
        if ($sku) {
            return $sku.Substring(0,4).ToUpper()
        }
    } catch {}
    return $null
}

function Get-LenovoDriverPackUrl($machineType, $os) {

    try {
        $json = Invoke-WebRequest "https://download.lenovo.com/cdrt/ddrc/recipecard.json" -UseBasicParsing | ConvertFrom-Json

        $family = $json.ThinkPad
        if (-not $family) { return $null }

        $model = $family | Where-Object {
            $_.types -match "^$machineType"
        } | Select-Object -First 1

        if (-not $model) { return $null }

        $osId = ($json.OperatingSystems | Where-Object name -eq $os).id

        $recipe = $json.RecipeCards | Where-Object {
            $_.modelId -eq $model.id -and $_.osId -eq $osId
        }

        if (-not $recipe) { return $null }

        $pack = $json.SCCMPacks | Where-Object id -eq $recipe.sccmPack

        return $pack.url
    }
    catch {
        return $null
    }
}

# =========================
# FORM SETUP
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool"
$form.Size = New-Object System.Drawing.Size(520,340)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,15)
$title.Text = "Detecting system..."
$form.Controls.Add($title)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.Size = New-Object System.Drawing.Size(460,140)
$statusBox.Location = New-Object System.Drawing.Point(20,60)
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(460,20)
$progress.Location = New-Object System.Drawing.Point(20,210)
$form.Controls.Add($progress)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Install Drivers"
$button.Size = New-Object System.Drawing.Size(160,35)
$button.Location = New-Object System.Drawing.Point(180,245)
$form.Controls.Add($button)

function Log($msg) {
    $statusBox.AppendText("$msg`r`n")
    $statusBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-Model {
    (Get-CimInstance Win32_ComputerSystem).Model.Trim()
}

function Get-DriverLink($manufacturer) {
    switch -Wildcard ($manufacturer) {
        "*Dell*"   { return "https://www.dell.com/support" }
        "*HP*"     { return "https://support.hp.com" }
        "*Lenovo*" { return "https://pcsupport.lenovo.com" }
        default     { return "" }
    }
}

function Check-DownloadsForExe {
    $downloads = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

    if (Test-Path $downloads) {

        $exe = Get-ChildItem $downloads -Filter *.exe -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

        if ($exe) {
            Log "Found EXE in Downloads: $($exe.Name)"
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

    Log "Starting driver installation..."

    $manufacturer = (Get-CimInstance Win32-ComputerSystem).Manufacturer
    $model = Get-Model
    $title.Text = "Driver Installer - $model"

    Log "Manufacturer: $manufacturer"
    Log "Model: $model"

    # =========================
    # COLLECT LOCAL INF DRIVERS
    # =========================
    $paths = @("C:\Drivers","C:\SWSetup","C:\DRIVERS","C:\Users\Administrator\")
    $allDrivers = @()

    foreach ($base in $paths) {
        if (Test-Path $base) {
            $allDrivers += Get-ChildItem $base -Recurse -Filter *.inf -ErrorAction SilentlyContinue
        }
    }

    # =========================
    # PRIORITY 1: LOCAL DRIVERS
    # =========================
    if ($allDrivers.Count -gt 0) {

        Log "Found $($allDrivers.Count) local driver files."

        $i = 0
        foreach ($driver in $allDrivers) {
            $i++
            $progress.Value = ($i / $allDrivers.Count) * 100

            Log "Installing: $($driver.Name)"
            pnputil /add-driver "$($driver.FullName)" /install | Out-Null
        }

        Log "Local driver installation complete."
        return
    }

    # =========================
    # PRIORITY 2: LENOVO AUTO DRIVER PACK (NEW)
    # =========================
    if ($manufacturer -match "Lenovo") {

        Log "Lenovo device detected. Checking Lenovo driver repository..."

        $machineType = Get-LenovoMachineType

        if ($machineType) {

            Log "Machine Type: $machineType"

            $lenovoUrl = Get-LenovoDriverPackUrl $machineType "Windows 11"

            if ($lenovoUrl) {

                Log "Lenovo driver pack found."
                Log $lenovoUrl

                $file = "$env:TEMP\lenovo_driverpack.exe"

                Invoke-WebRequest $lenovoUrl -OutFile $file

                Log "Downloaded Lenovo driver pack."

                Start-Process $file -ArgumentList "/VERYSILENT" -Wait

                Log "Lenovo driver pack executed."
                return
            }

            Log "No Lenovo pack found."
        }
    }

    # =========================
    # PRIORITY 3: DOWNLOADS EXE
    # =========================
    Log "Checking Downloads for installer..."
    if (Check-DownloadsForExe) { return }

    # =========================
    # PRIORITY 4: OEM FALLBACK
    # =========================
    Log "Opening OEM support page..."
    Start-Process (Get-DriverLink $manufacturer)

    $button.Enabled = $true
}

$button.Add_Click({ Start-Install })

$form.Add_Shown({
    $form.Activate()
    Start-Install
})

[void]$form.ShowDialog()