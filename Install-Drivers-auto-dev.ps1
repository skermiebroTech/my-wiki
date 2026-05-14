$sysid = (Get-CimInstance Win32_BaseBoard).Product.Trim().ToLower()
$build = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
$dv    = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion
Write-Host "SysID:$sysid  Build:$build  Display:$dv"

$base = "https://hpia.hpcloud.hp.com/ref/$sysid"
# Lowercase, with the 10.0./11.0. prefix HPIA uses
$candidates = @(
    "$base/${sysid}_64_11.0.25h2.cab",
    "$base/${sysid}_64_11.0.24h2.cab",
    "$base/${sysid}_64_11.0.23h2.cab",
    "$base/${sysid}_64_11.0.22h2.cab",
    "$base/${sysid}_64_11.0.21h2.cab"
)
foreach ($u in $candidates) {
    try {
        $r = Invoke-WebRequest -Uri $u -Method Head -UseBasicParsing -ErrorAction Stop
        Write-Host "OK    $($r.StatusCode)  $u"
    } catch {
        $code = $_.Exception.Response.StatusCode.value__
        Write-Host "FAIL  $code  $u"
    }
}