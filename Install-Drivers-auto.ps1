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
$global:sync.LogFile = "$env:TEMP\driver_tool_v6_log.txt"

# =========================
# UI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool V6"
$form.Size = New-Object System.Drawing.Size(520,340)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,15)
$title.Text = "Driver Installer"
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

    Add-Content $global:sync.LogFile $line
}

# =========================
# CANCEL
# =========================
$cancel.Add_Click({
    $global:sync.Cancel = $true
    Log "Cancel requested..."
})

# =========================
# LENOVO SAFE ENGINE (V6 FIXED)
# =========================
function Invoke-Lenovo($sync) {

    try {

        Log "Checking Lenovo device..."

        $cs = Get-CimInstance Win32_ComputerSystem
        if ($cs.Manufacturer -notmatch "Lenovo") {
            Log "Not Lenovo device."
            return $false
        }

        $sku = (Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber
        if (-not $sku) {
            Log "No SKU found."
            return $false
        }

        $machineType = ($sku -replace "[^0-9A-Z]", "").Substring(0,4)
        Log "Machine Type: $machineType"

        Log "Downloading Lenovo recipe JSON..."
        $json = Invoke-WebRequest "https://download.lenovo.com/cdrt/ddrc/recipecard.json" -UseBasicParsing | ConvertFrom-Json

        # SAFE lookup (no guessing families)
        $modelMatch = $null

        foreach ($family in $json.PSObject.Properties.Name) {
            $items = $json.$family
            foreach ($item in $items) {
                if ($item.types -contains $machineType) {
                    $modelMatch = $item
                    break
                }
            }
            if ($modelMatch) { break }
        }

        if (-not $modelMatch) {
            Log "No matching Lenovo model found."
            return $false
        }

        Log "Model found: $($modelMatch.name)"

        # OS auto detect
        $osName = if ([Environment]::OSVersion.Version.Build -ge 22000) { "Windows 11" } else { "Windows 10" }
        Log "Detected OS: $osName"

        $osId = ($json.OperatingSystems | Where-Object name -eq $osName).id

        $recipe = $json.RecipeCards | Where-Object {
            $_.modelId -eq $modelMatch.id -and $_.osId -eq $osId
        }

        if (-not $recipe) {
            Log "No recipe for OS."
            return $false
        }

        $pack = $json.SCCMPacks | Where-Object id -eq $recipe.sccmPack

        if (-not $pack -or -not $pack.url) {
            Log "No valid SCCM pack URL."
            return $false
        }

        Log "Driver pack found:"
        Log $pack.url

        # validate URL
        if ($pack.url -notmatch "^https?://") {
            Log "Invalid URL format."
            return $false
        }

        $file = "$env:TEMP\lenovo_driverpack.exe"

        Log "Downloading..."
        Invoke-WebRequest $pack.url -OutFile $file

        Log "Running installer..."
        Start-Process $file -Wait

        Log "Lenovo complete."
        return $true
    }
    catch {
        Log "Lenovo error: $($_.Exception.Message)"
        return $false
    }
}

# =========================
# MAIN RUN
# =========================
$button.Add_Click({

    $button.Enabled = $false
    $progress.Value = 0
    $global:sync.Cancel = $false

    $ps = [PowerShell]::Create()

    $ps.AddScript({

        param($sync)

        function Log($m) {
            $sync.StatusBox.Invoke([action]{
                $sync.StatusBox.AppendText("$m`r`n")
            })
        }

        function Set($v) {
            $sync.Progress.Invoke([action]{ $sync.Progress.Value = $v })
        }

        Log "Starting..."
        Set 10

        # LOCAL DRIVERS
        $paths = @("C:\Drivers","C:\SWSetup","C:\DRIVERS")
        foreach ($p in $paths) {
            if (Test-Path $p) {
                Get-ChildItem $p -Recurse -Filter *.inf -ErrorAction SilentlyContinue | ForEach-Object {
                    Log "Installing $($_.Name)"
                    pnputil /add-driver $_.FullName /install | Out-Null
                }
            }
        }

        Set 60

        # LENOVO
        $ok = Invoke-Lenovo $sync

        if ($ok) {
            Set 100
            Log "Done."
            $sync.Button.Invoke([action]{ $sync.Button.Enabled = $true })
            return
        }

        # FALLBACK
        Log "Opening OEM page..."
        Start-Process "https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html"

        Set 100
        $sync.Button.Invoke([action]{ $sync.Button.Enabled = $true })

    }).AddArgument($global:sync)

    $ps.BeginInvoke()
})

# =========================
# SHOW
# =========================
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()