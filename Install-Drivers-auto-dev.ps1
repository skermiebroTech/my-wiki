$sysid = (Get-CimInstance Win32_BaseBoard).Product.Trim().ToLower()
$url   = "https://hpia.hpcloud.hp.com/ref/$sysid/${sysid}_64_11.0.25h2.cab"
$out   = "$env:USERPROFILE\Downloads\${sysid}_25h2.cab"
Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
Write-Host "Saved to: $out  ($([math]::Round((Get-Item $out).Length / 1MB, 2)) MB)"