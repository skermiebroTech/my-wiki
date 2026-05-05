Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# ADMIN CHECK
# =========================
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptUrl = "https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1"
    Start-Process powershell "-ExecutionPolicy Bypass -Command `"irm $scriptUrl | iex`"" -Verb RunAs
    exit
}

# =========================
# STATE
# =========================
$global:sync = [hashtable]::Synchronized(@{})
$global:sync.Cancel = $false
$global:sync.LogFile = "$env:TEMP\driver_installer_log.txt"

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

# attach sync objects
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
# LENOVO SUPPORT
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
# CANCEL
# =========================
$cancel.Add_Click({
    $global:sync.Cancel = $true
    Log "Cancel requested..."
})

# =========================
# RUNSPACE INSTALL
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
        SetProgress 5

        # =========================
        # SCAN DRIVERS
        # =========================
        $paths = @("C:\Drivers","C:\SWSetup","C:\DRIVERS")
        $drivers = @()

        foreach ($p in $paths) {
            if (Test-Path $p) {
                $drivers += Get-ChildItem $p -Recurse -Filter *.inf -ErrorAction SilentlyContinue
            }
        }

        # =========================
        # LOCAL INSTALL
        # =========================
        if ($drivers.Count -gt 0) {

            Log "Found $($drivers.Count) local drivers."
            $i = 0

            foreach ($d in $drivers) {

                if (CheckCancel) { return }

                $i++
                $percent = ($i / $drivers.Count) * 60
                SetProgress $percent

                Log "Installing $($d.Name)"

                $attempts = 0
                $done = $false

                while (-not $done -and $attempts -lt 2) {
                    try {
                        pnputil /add-driver "$($d.FullName)" /install | Out-Null
                        $done = $true
                    } catch {
                        $attempts++
                    }
                }
            }

            SetProgress 70
        }

        # =========================
        # LENOVO AUTO
        # =========================
        $manufacturer = (Get-CimInstance Win32-ComputerSystem).Manufacturer

        if ($manufacturer -match "Lenovo") {

            Log "Lenovo detected..."
            SetProgress 75

            $mt = (Get-CimInstance Win32-ComputerSystemProduct).Name.Substring(0,4).ToUpper()
            Log "Machine Type: $mt"

            $url = Get-LenovoDriverPackUrl $mt "Windows 11"

            if ($url) {
                Log "Downloading Lenovo driver pack..."
                $file = "$env:TEMP\lenovo.exe"

                Invoke-WebRequest $url -OutFile $file

                Log "Running Lenovo installer..."
                Start-Process $file -Wait
            }
        }

        # =========================
        # OEM FALLBACK
        # =========================
        if (CheckCancel) { return }

        SetProgress 90
        Log "Opening OEM support page..."

        Start-Process "https://pcsupport.lenovo.com"

        SetProgress 100
        Log "Complete."

        $sync.Button.Invoke([action]{ $sync.Button.Enabled = $true })

    }).AddArgument($global:sync)

    $ps.BeginInvoke()
})

# =========================
# SHOW FORM
# =========================
$form.Add_Shown({
    $form.Activate()
})

[void]$form.ShowDialog()