# =============================================================
# Install-Drivers-auto.Tests.ps1
#
# Fixture tests for the pure parser functions in the driver installer.
# No Pester needed: plain PowerShell (5.1 or 7). Exit code 1 on any failure.
#
# Run from the repo root:
#   powershell -ExecutionPolicy Bypass -File tests\Install-Drivers-auto.Tests.ps1
#   pwsh -File tests/Install-Drivers-auto.Tests.ps1 -ScriptPath Install-Drivers-auto.ps1
#
# Fixtures (tests\fixtures) are trimmed captures of the live vendor catalogs
# taken on 2026-09-14. Refresh them when a vendor changes a format.
# =============================================================
param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot "..\Install-Drivers-auto-dev.ps1"),
    [string]$Fixtures   = (Join-Path $PSScriptRoot "fixtures")
)
$ErrorActionPreference = 'Stop'
# Pull only the pure functions out of the one-file script via the AST (the
# script's main body creates a WinForms form, so it cannot be dot-sourced here).
$tok=$null;$err=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($ScriptPath,[ref]$tok,[ref]$err)
$want='Get-MsCatalogResultRows','Select-MsCatalogDriverGuid','Find-DellIndexManifestPath','Test-DellOsCode','Test-DellOsArch','Test-ActionableProblemCode','Get-ProblemCodeHint','Select-HpDriverPack','Find-DellCatalogDeviceMatches'
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)) {
    if ($want -contains $f.Name) { Invoke-Expression $f.Extent.Text }
}
function Log { param([string]$msg,$Level,$Event,$Context) }   # stub
$script:IgnoredProblemCodes = @(22, 24, 29, 45, 46, 53, 55)
$pass=0;$fail=0
function Check($name,$cond){ if($cond){$script:pass++; "  ok   $name"} else {$script:fail++; "  FAIL $name"} }

"== Update Catalog =="
$html = Get-Content (Join-Path $Fixtures 'mscatalog-search-8086-51F0.html') -Raw
$rows = @(Get-MsCatalogResultRows -Html $html)
Check "parses 25 result rows" ($rows.Count -eq 25)
Check "row has guid/title/products/date" ($rows[0].Guid -match '^[0-9a-f-]{36}$' -and $rows[0].Title -and $rows[0].Products -and $rows[0].Updated -gt [datetime]'2020-01-01')
$g11 = Select-MsCatalogDriverGuid -Rows $rows -OsBuild 26100
$g10 = Select-MsCatalogDriverGuid -Rows $rows -OsBuild 19045
$r11 = $rows | Where-Object Guid -eq $g11; $r10 = $rows | Where-Object Guid -eq $g10
Check "Win11 pick names Windows 11" ($r11.Products -match 'Windows 11')
Check "Win10 pick names Windows 10" ($r10.Products -match 'Windows 10')
Check "Win11 pick is newest of its generation" ($r11.Updated -eq (($rows | Where-Object { $_.Products -match 'Windows 11' -and $_.Products -notmatch 'Server' } | Measure-Object -Property Updated -Maximum).Maximum))
"  picked (Win11): $($r11.Title) | $($r11.Products) | $($r11.Updated.ToString('yyyy-MM-dd'))"

"== Dell index =="
[xml]$idx = Get-Content (Join-Path $Fixtures 'dell-CatalogIndexPC-trimmed.xml') -Raw
$p = Find-DellIndexManifestPath -Index $idx -SystemSKU '0b0b'
Check "0B0B resolves to Latitude_0B0B.cab" ($p -like '*/Latitude_0B0B.cab')
Check "0B3D (Vostro 5320) present" ((Find-DellIndexManifestPath -Index $idx -SystemSKU '0B3D') -like '*.cab')
Check "unknown SKU returns null" ($null -eq (Find-DellIndexManifestPath -Index $idx -SystemSKU 'ZZZZ'))
Check "baseLocation attr" ($idx.DocumentElement.GetAttribute('baseLocation') -eq 'downloads.dell.com')

"== Dell per-SKU matcher =="
$raw = Get-Content (Join-Path $Fixtures 'dell-Latitude_0B0B-trimmed.xml') -Raw
$raw = $raw -replace '^\s*<\?xml[^>]*\?>', ''
[xml]$cat = $raw
$w11 = @('W21H4','W21P4','W21S4','W21S5','W11AH','W11AP','W11S5','W11TM','IOTL5'); $w10 = @('W10H4','W10P4','W10H2','W10P2','IOT01','IOTL3','IOTL4','WTCLD')
$env:PROCESSOR_ARCHITECTURE = 'AMD64'
$m = @(Find-DellCatalogDeviceMatches -Catalog $cat -DevVenDev @([pscustomobject]@{VEN='8086';DEV='51F0'}) -DevSubsysPairs @() -SystemSKU '0B0B' -IsWin11 $true -ApplySkuFilter $true -Win11Codes $w11 -Win10Codes $w10)
Check "AX211 matches >= 10 packages with SKU filter on" ($m.Count -ge 10)
Check "every match carries a SHA256" (@($m | Where-Object { $_.Sha256 -match '^[0-9a-f]{64}$' }).Count -eq $m.Count)
$newest = $m | Sort-Object Date -Descending | Select-Object -First 1
Check "newest AX211 is 24.40.0.4" ($newest.Version -eq '24.40.0.4')
Check "osCode prefix accepted (W11XX)" (Test-DellOsCode -Code 'W11ZZ' -IsWin11 $true -Win11Codes $w11 -Win10Codes $w10)
Check "osCode W10 rejected on Win11" (-not (Test-DellOsCode -Code 'W10P4' -IsWin11 $true -Win11Codes $w11 -Win10Codes $w10))
Check "arch x64 accepted on AMD64" (Test-DellOsArch 'x64')
Check "arch ARM64 rejected on AMD64" (-not (Test-DellOsArch 'ARM64'))
$env:PROCESSOR_ARCHITECTURE = 'ARM64'
Check "arch ARM64 accepted on ARM64" (Test-DellOsArch 'arm64')
$env:PROCESSOR_ARCHITECTURE = 'AMD64'

"== HP driver pack catalog =="
[xml]$dpc = (Get-Content (Join-Path $Fixtures 'hp-HPClientDriverPackCatalog-trimmed.xml') -Raw).TrimStart([char]0xFEFF)
$pk = Select-HpDriverPack -Catalog $dpc -SysId '8b41' -IsWin11 $true -DisplayVersion '24H2'
Check "8b41 24H2 -> sp160195" ($pk.SoftPaqId -eq 'sp160195' -and $pk.OSName -like '*24H2*')
Check "url + sha present" ($pk.Url -like 'https://*sp160195.exe' -and $pk.SHA256 -match '^[0-9A-Fa-f]{64}$')
$pk2 = Select-HpDriverPack -Catalog $dpc -SysId '8B41' -IsWin11 $true -DisplayVersion '26H1'
Check "unknown DisplayVersion falls back to newest Win11 (25H2)" ($pk2.OSName -like '*25H2*')
$pk3 = Select-HpDriverPack -Catalog $dpc -SysId '8b41' -IsWin11 $false -DisplayVersion '22H2'
Check "Win10 22H2 -> sp147159" ($pk3.SoftPaqId -eq 'sp147159')
Check "unknown SystemId returns null" ($null -eq (Select-HpDriverPack -Catalog $dpc -SysId 'zzzz' -IsWin11 $true -DisplayVersion '24H2'))

"== problem codes =="
Check "22 (disabled) not actionable" (-not (Test-ActionableProblemCode 22))
Check "45 (phantom) not actionable" (-not (Test-ActionableProblemCode '45'))
Check "28 actionable" (Test-ActionableProblemCode 28)
Check "0 not actionable" (-not (Test-ActionableProblemCode 0))
Check "hint for 48 mentions policy" ((Get-ProblemCodeHint 48) -match 'POLICY')
""
"passed: $pass  failed: $fail"
if ($fail -gt 0) { exit 1 }
