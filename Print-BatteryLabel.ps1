# Print-BatteryLabel.ps1
# Prints battery health percent to a Zebra printer (ZPL over TCP 9100).
# Run from Win+R:
#   powershell -NoProfile -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Print-BatteryLabel.ps1|iex"

$PrinterIp = '172.17.21.186'
$PrinterPort = 9100

$fcc = (Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity |
    Measure-Object -Property FullChargedCapacity -Sum).Sum

# Design capacity: try BatteryStaticData (needs admin on some boxes), fall back to Win32_PortableBattery
$design = $null
try {
    $design = (Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction Stop |
        Measure-Object -Property DesignedCapacity -Sum).Sum
} catch {}
if (-not $design) {
    $batt = Get-CimInstance Win32_PortableBattery
    $design = ($batt | Measure-Object -Property DesignCapacity -Sum).Sum
    # Some firmware reports mAh here while FCC is mWh; convert if it looks that way
    if ($design -and $batt[0].DesignVoltage -and ($fcc / $design) -gt 3) {
        $design = $design * $batt[0].DesignVoltage / 1000
    }
}

if (-not $fcc -or -not $design) {
    Write-Host "Could not read battery capacities (FCC=$fcc Design=$design). No battery, or WMI blocked."
    Read-Host 'Press Enter to exit'
    return
}

$health = [math]::Round($fcc / $design * 100)
Write-Host "FullCharged: $fcc mWh  Design: $design mWh  Health: $health%"

$model = (Get-CimInstance Win32_ComputerSystem).Model.Trim()
$tag = (Get-CimInstance Win32_BIOS).SerialNumber.Trim()
$cycles = $null
try {
    $cycles = (Get-CimInstance -Namespace root\wmi -ClassName BatteryCycleCount -ErrorAction Stop |
        Measure-Object -Property CycleCount -Sum).Sum
} catch {}
$line4 = Get-Date -Format 'dd/MM/yy'
if ($cycles) { $line4 += '  Cycles: ' + $cycles }

# Build one ZPL text line, shrinking font width (and height to match) so text fits the label
function Get-ZplLine([int]$y, [int]$height, [int]$width, [string]$text) {
    $maxDots = 340  # printable width at x=30 on a 400-dot label
    # ZPL font 0 average char advance is ~0.6 x the width parameter
    $fitWidth = [math]::Floor($maxDots / ($text.Length * 0.6))
    if ($fitWidth -lt $width) {
        $height = [math]::Max(14, [math]::Round($height * $fitWidth / $width))
        $width = [math]::Max(14, $fitWidth)
    }
    return '^FO30,' + $y + '^A0N,' + $height + ',' + $width + '^FD' + $text + '^FS'
}

$zpl = '^XA^PW400^LL200' +
    (Get-ZplLine 17 60 50 ('Battery: ' + $health + '%')) +
    (Get-ZplLine 84 28 28 ($model + '  ' + $tag)) +
    (Get-ZplLine 119 28 28 ([math]::Round($design).ToString() + '/' + [math]::Round($fcc) + ' mWh')) +
    (Get-ZplLine 154 28 28 $line4) +
    '^FO305,95^A0N,70,60^FD' + $(if ($health -ge 70) { ':)' } else { ':(' }) + '^FS^XZ'
try {
    $client = New-Object Net.Sockets.TcpClient($PrinterIp, $PrinterPort)
    $stream = $client.GetStream()
    $bytes = [Text.Encoding]::ASCII.GetBytes($zpl)
    $stream.Write($bytes, 0, $bytes.Length)
    $client.Close()
    Write-Host 'Label sent.'
} catch {
    Write-Host "Failed to reach printer ${PrinterIp}:${PrinterPort} - $_"
    Read-Host 'Press Enter to exit'
}
