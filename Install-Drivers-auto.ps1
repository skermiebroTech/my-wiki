Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# ADMIN CHECK
# =========================
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptUrl = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers.ps1"
    Start-Process powershell "-ExecutionPolicy Bypass -Command `"irm $scriptUrl | iex`"" -Verb RunAs
    exit
}

# =========================
# STATE
# =========================
$global:sync = [hashtable]::Synchronized(@{})
$global:sync.Cancel = $false
$global:sync.LogFile = "$env:TEMP\driver_tool_v5_log.txt"

# =========================
# UI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool V5"
$form.Size = New-Object System.Drawing.Size(520,340)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,15)
$title.Text = "Driver Installer Ready"
$form.Controls.Add($title)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.Size = New-Object System.Drawing.Size(460,140)
$statusBox.Location = New-Object System.Drawing.Point(20,60)
$statusBox.ReadOnly = $true
$statusBox.Text = "Ready... Click Install to begin.`r`n"
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

$cancel = New-Object System.Windows.Forms.Button
$cancel.Text = "Cancel"
$cancel.Size = New-Object System.Drawing.Size(160,35)
$cancel.Location = New-Object System.Drawing.Point(20,245)
$form.Controls.Add($cancel)

$global:sync.StatusBox = $statusBox
$global:sync.Progress = $progress
$global:sync.Button = $button

# =========================
# LOG
# =========================
function Log($msg) {
    $time = Get-Date -Format "HH:mm:ss"
    $line = "[$time] $msg"

    $statusBox.AppendText("$line`r`n")
    $statusBox.ScrollToCaret()

    Add-Content -Path $global:sync.LogFile -Value $line
}

# =========================
# CANCEL
# =========================
$cancel.Add_Click({
    $global:sync.Cancel = $true
    Log "Cancel requested..."
})

# =========================
# LENOVO ENGINE (FIXED V5)
# =========================
function Invoke-LenovoDriverPack($sync) {

    try {

        Log "Checking Lenovo system..."

        $sku = (Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber
        if (-not $sku) {
            Log "No Lenovo SKU detected."
            return $false
        }

        $machineType = ($sku -replace "[^0-9A-Z]", "").Substring(0,4)
        Log "Machine Type: $machineType"

        $json = Invoke-WebRequest "https://download.lenovo.com/cdrt/ddrc/recipecard.json" -UseBasicParsing | ConvertFrom-Json

        $familyKey = $json.PSObject.Properties.Name | Where-Object { $_ -match "Think" } | Select-Object -First 1
        $family = $json.$familyKey

        $model = $family | Where-Object { $_.types -contains $machineType } | Select-Object -First 1

        if (-not $model) {
            Log "No Lenovo model match."
            return $false
        }

        # =========================
        # AUTO OS DETECTION (FIX)
        # =========================
        $osName = if ([Environment]::OSVersion.Version.Build -ge 22000) {
            "Windows 11"
        } else {
            "Windows 10"
        }

        Log "Detected OS: $osName"

        $osId = ($json.OperatingSystems | Where-Object name -eq $osName).id

        # find recipe
        $recipe = $json.RecipeCards | Where-Object {
            $_.modelId -eq $model.id -and $_.osId -eq $osId
        }

        if (-not $recipe) {
            Log "No Lenovo recipe found for $osName."
            return $false
        }

        $pack = $json.SCCMPacks | Where-Object id -eq $recipe.sccmPack

        if (-not $pack.url) {
            Log "No Lenovo driver pack URL found."
            return $false
        }

        Log "Downloading Lenovo driver pack..."
        Log $pack.url

        $file = "$env:TEMP\lenovo_driverpack.exe"

        Invoke-WebRequest $pack.url -OutFile $file

        Log "Running Lenovo driver pack..."
        Start-Process $file -Wait

        Log "Lenovo driver installation complete."
        return $true
    }
    catch {
        Log "Lenovo error: $($_.Exception.Message)"
        return $false
    }
}

# =========================
# RUNSPACE WORKER
# =========================
$button.Add_Click({

    $button.Enabled = $false
    $global:sync.Cancel = $false
    $progress.Value = 0

    $ps = [PowerShell]::Create()

    $ps.AddScript({

        param($sync)

        function Log($m) {
            $sync.StatusBox.Invoke([action]{
                $sync.StatusBox.AppendText("$m`r`n")
                $sync.StatusBox.ScrollToCaret()
            })
        }

        function SetProgress($v) {
            $sync.Progress.Invoke([action]{ $sync.Progress.Value = $v })
        }

        function CheckCancel {
            if ($sync.Cancel) {
                Log "Cancelled."
                return $true
            }
            return $false
        }

        Log "Starting driver installation..."
        SetProgress 10

        # =========================
        # LOCAL DRIVERS
        # =========================
        $paths = @("C:\Drivers","C:\SWSetup","C:\DRIVERS")
        $drivers = @()

        foreach ($p in $paths) {
            if (Test-Path $p) {
                $drivers += Get-ChildItem $p -Recurse -Filter *.inf -ErrorAction SilentlyContinue
            }
        }

        if ($drivers.Count -gt 0) {

            Log "Found $($drivers.Count) local drivers."

            $i = 0
            foreach ($d in $drivers) {

                if (CheckCancel) { return }

                $i++
                SetProgress (10 + (($i / $drivers.Count) * 40))

                Log "Installing $($d.Name)"

                try {
                    pnputil /add-driver "$($d.FullName)" /install | Out-Null
                } catch {}
            }
        }

        # =========================
        # LENOVO (FIXED)
        # =========================
        if (CheckCancel) { return }

        $manufacturer = (Get-CimInstance Win32-ComputerSystem).Manufacturer

        if ($manufacturer -match "Lenovo") {

            SetProgress 60

            $ok = Invoke-LenovoDriverPack $sync

            if ($ok) {
                SetProgress 100
                Log "Lenovo completed."
                $sync.Button.Invoke([action]{ $sync.Button.Enabled = $true })
                return
            }

            Log "Lenovo not applicable or failed."
        }

        # =========================
        # OEM FALLBACK
        # =========================
        if (CheckCancel) { return }

        SetProgress 85
        Log "Opening OEM support page..."

        Start-Process "https://pcsupport.lenovo.com"

        SetProgress 100
        Log "Complete."

        $sync.Button.Invoke([action]{ $sync.Button.Enabled = $true })

    }).AddArgument($global:sync)

    $ps.BeginInvoke()
})

# =========================
# SHOW
# =========================
$form.Add_Shown({ $form.Activate() })

[void]$form.ShowDialog()