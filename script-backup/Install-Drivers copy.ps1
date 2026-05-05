Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Installer Tool"
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = "CenterScreen"

# Title Label
$title = New-Object System.Windows.Forms.Label
$title.Text = "One-Click Driver Installer"
$title.AutoSize = $true
$title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(120,20)
$form.Controls.Add($title)

# Status Box
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.Size = New-Object System.Drawing.Size(440,120)
$statusBox.Location = New-Object System.Drawing.Point(20,70)
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

# Progress Bar
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(440,20)
$progress.Location = New-Object System.Drawing.Point(20,200)
$form.Controls.Add($progress)

# Button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Install Drivers"
$button.Size = New-Object System.Drawing.Size(150,30)
$button.Location = New-Object System.Drawing.Point(170,230)
$form.Controls.Add($button)

# Function to log text
function Log($msg) {
    $statusBox.AppendText("$msg`r`n")
    $statusBox.ScrollToCaret()
}

# Button Click Event
$button.Add_Click({

    $button.Enabled = $false
    $progress.Value = 0

    Log "Starting driver installation..."

    $manufacturer = (Get-CimInstance Win32_ComputerSystem).Manufacturer
    Log "Detected: $manufacturer"

    $paths = @()

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
        Log "Unknown manufacturer, scanning common paths..."
        $paths += "C:\Drivers\","C:\SWSetup\","C:\DRIVERS\","C:\Users\Administrator\"
    }

    $validPaths = @()

    foreach ($base in $paths) {
        if (Test-Path $base) {
            $dirs = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue
            foreach ($dir in $dirs) {
                if (Get-ChildItem $dir.FullName -Recurse -Filter *.inf -ErrorAction SilentlyContinue) {
                    $validPaths += $dir.FullName
                }
            }
        }
    }

    if ($validPaths.Count -eq 0) {
        Log "No driver folders found!"
        $button.Enabled = $true
        return
    }

    $step = 100 / $validPaths.Count
    $current = 0

    foreach ($path in $validPaths) {
        Log "Installing from: $path"
        pnputil /add-driver "$path\*.inf" /subdirs /install | Out-Null

        $current += $step
        $progress.Value = [math]::Min($current,100)
    }

    $progress.Value = 100
    Log "Installation complete!"

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Drivers installed. Reboot now?",
        "Done",
        "YesNo"
    )

    if ($result -eq "Yes") {
        Restart-Computer -Force
    } else {
        $button.Enabled = $true
    }
})

# Run form
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()