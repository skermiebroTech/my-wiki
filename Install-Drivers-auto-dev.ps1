$sysid = (Get-CimInstance Win32_BaseBoard).Product.Trim().ToLower()
$os    = (Get-CimInstance Win32_OperatingSystem)
$ver   = $os.Version              # e.g. 10.0.26200
$build = [int]$os.BuildNumber     # e.g. 26200

# Build version code: 22000-22631 = 22H2, 22631-26100 = 23H2/24H2, etc.
# Easiest: just tell me what these print.
Write-Host "SysID:   $sysid"
Write-Host "Version: $ver"
Write-Host "Build:   $build"
Write-Host "Display: $((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion)"

# Now try the likely URL patterns — paste whichever one returns 200:
$candidates = @(
    "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$sysid/${sysid}_64_11.0.24H2.cab",
    "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$sysid/${sysid}_64_11.0.23H2.cab",
    "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$sysid/${sysid}_64_24H2.cab",
    "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$sysid/${sysid}_64_23H2.cab",
    "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/$sysid/${sysid}_64_22H2.cab"
)
foreach ($u in $candidates) {
    try {
        $r = Invoke-WebRequest -Uri $u -Method Head -UseBasicParsing -ErrorAction Stop
        Write-Host "OK    $($r.StatusCode)  $u"
    } catch {
        Write-Host "FAIL  $($_.Exception.Response.StatusCode.value__)  $u"
    }
}