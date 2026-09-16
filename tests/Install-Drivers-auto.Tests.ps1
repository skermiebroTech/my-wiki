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
$want='Test-HpVersionNewer','ConvertTo-HpMatchId','Select-HpConsumerSoftpaqs','Invoke-HpSupportJson','ConvertFrom-HpJson','ConvertTo-HpBytes','Find-HpSupportProductOids','Select-HpSupportOs','Get-HpCvaInfo','Test-HpCvaAppliesToDevices','Get-MsCatalogResultRows','Select-MsCatalogDriverGuid','Find-DellIndexManifestPath','Test-DellOsCode','Test-DellOsArch','Test-ActionableProblemCode','Get-ProblemCodeHint','Select-HpDriverPack','Find-DellCatalogDeviceMatches'
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)) {
    if ($want -contains $f.Name) { Invoke-Expression $f.Extent.Text }
}
function Log { param([string]$msg,$Level,$Event,$Context) $script:LogLines += $msg }   # stub
function Log-Diag { param([string]$msg,$Context) }
$script:LogLines = @()
if (-not $env:TEMP) { $env:TEMP = [System.IO.Path]::GetTempPath() }
if (-not (Get-Command curl.exe -EA SilentlyContinue)) { function curl.exe { & (Get-Command curl -CommandType Application | Select-Object -First 1).Source @args } }   # macOS/Linux shim for the real-call tests
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

"== HP consumer tier (v1.30.0) =="
$sj = Get-Content (Join-Path $Fixtures 'hp-support-search-model.json') -Raw | ConvertFrom-Json
$oids = @(Find-HpSupportProductOids -SearchJson $sj)
Check "model-name search yields product OIDs" ($oids.Count -ge 1 -and $oids[0].Oid -match '^\d+$')
$sj2 = Get-Content (Join-Path $Fixtures 'hp-support-search-product.json') -Raw | ConvertFrom-Json
$oids2 = @(Find-HpSupportProductOids -SearchJson $sj2)
Check "product-name search yields OID 31291399" (@($oids2 | Where-Object { $_.Oid -eq '31291399' }).Count -eq 1)
Check "empty search yields nothing" (@(Find-HpSupportProductOids -SearchJson $null).Count -eq 0)
$rawTxt = Get-Content (Join-Path $Fixtures 'hp-support-search-model.json') -Raw
$oidsRaw = @(Find-HpSupportProductOids -SearchJson $rawTxt)
Check "raw-text search yields the same OIDs as the object walk" ($oidsRaw.Count -eq $oids.Count -and $oidsRaw[0].Oid -eq $oids[0].Oid)
$rawTxt2 = Get-Content (Join-Path $Fixtures 'hp-support-search-product.json') -Raw
# Real call through curl against a local file: proves the GET path, the -AsText switch and the body check.
$fileUrl = "file://" + ((Resolve-Path (Join-Path $Fixtures 'hp-support-search-product.json')).Path -replace '\\','/' -replace '^([A-Za-z]):','/$1:')
$script:LogLines = @()
$viaHelper = Invoke-HpSupportJson -AsText -Url $fileUrl
Check "Invoke-HpSupportJson -AsText returns the file text with no error" ($viaHelper -is [string] -and $viaHelper.Length -gt 500 -and @($script:LogLines | Where-Object { $_ -match 'failed|not JSON|empty' }).Count -eq 0)
Check "Invoke-HpSupportJson GET parses to an object" ((Invoke-HpSupportJson -Url $fileUrl).data.kaaSResponse.code -eq 200)
Check "raw-text product search yields OID 31291399" (@(Find-HpSupportProductOids -SearchJson $rawTxt2 | Where-Object { $_.Oid -eq '31291399' }).Count -eq 1)
$oj = Get-Content (Join-Path $Fixtures 'hp-support-osVersionData.json') -Raw | ConvertFrom-Json
$os = Select-HpSupportOs -OsJson $oj -IsWin11 $true -DisplayVersion '24H2'
Check "Win11 machine falls back to Windows 10 (64-bit) generic" ($os -and $os.OsName -eq 'Windows 10 (64-bit)' -and $os.OsTmsId -eq '792898937266030878164166465223921' -and $os.PlatformId -eq '487192269364721453674728010296573')
$os10 = Select-HpSupportOs -OsJson $oj -IsWin11 $false -DisplayVersion '20H2'
Check "Win10 20H2 machine picks the 20H2 row" ($os10.OsName -like '*20H2*')
$cva = Get-HpCvaInfo -Text (Get-Content (Join-Path $Fixtures 'hp-sp144777.cva') -Raw)
Check "CVA softpaq number + SHA256" ($cva.SoftpaqNumber -eq 'sp144777' -and $cva.SHA256 -eq 'F009CC64AE92EBD6F32CE6348E5E5171E2E0EB1E4A628BC96913A5C792C0071A')
Check "CVA type Driver, 8 device ids, silent install" ($cva.Type -eq 'Driver' -and $cva.Devices.Count -eq 8 -and $cva.Devices -contains 'USB\VID_06CB&PID_00DF' -and $cva.SilentInstall)
Check "size text parses (15.5 MB, 627.42 KB, 1.3 GB, 12345)" ((ConvertTo-HpBytes "15.5 MB") -eq 16252928 -and (ConvertTo-HpBytes "627.42 KB") -eq 642478 -and (ConvertTo-HpBytes "1.3 GB") -eq 1395864371 -and (ConvertTo-HpBytes "12345") -eq 12345 -and (ConvertTo-HpBytes "") -eq 0)
Check "CVA sysids include 8830 (ProBook 635 Aero G7)" ($cva.SysIds -contains '8830')
$devs = @([pscustomobject]@{ Name='Synaptics WBDI'; HardwareIDs=@('USB\VID_06CB&PID_00DF&REV_0100','USB\VID_06CB&PID_00DF'); CompatibleIDs=@('USB\CLASS_FF') },
          [pscustomobject]@{ Name='Other'; HardwareIDs=@('PCI\VEN_8086&DEV_51F0'); CompatibleIDs=@() })
Check "CVA matches the fingerprint device by prefix" (@(Test-HpCvaAppliesToDevices -Cva $cva -MissingDevices $devs -SysId '8830') -eq @('Synaptics WBDI'))
Check "CVA rejects a machine whose SysId is not listed" (@(Test-HpCvaAppliesToDevices -Cva $cva -MissingDevices $devs -SysId '9999').Count -eq 0)
$cvaAudio = @{ Devices=@('HDAUDIO\FUNC_01&VEN_10EC&DEV_0285&SUBSYS_103C86B2'); SysIds=@('86B2') }
$devAudio = @([pscustomobject]@{ Name='Intel High Definition Audio'; HardwareIDs=@('INTELAUDIO\FUNC_01&VEN_10EC&DEV_0285&SUBSYS_103C86B2&REV_1000','INTELAUDIO\FUNC_01&VEN_10EC&DEV_0285&SUBSYS_103C86B2'); CompatibleIDs=@() })
Check "INTELAUDIO codec matches an HDAUDIO CVA id" (@(Test-HpCvaAppliesToDevices -Cva $cvaAudio -MissingDevices $devAudio -SysId '86B2').Count -eq 1)
$cvaSst = @{ Devices=@('PCI\VEN_8086&DEV_02C8&SUBSYS_86B2103C'); SysIds=@() }
$devDsp = @([pscustomobject]@{ Name='Intel High Definition DSP'; HardwareIDs=@('INTELAUDIO\DSP_CTLR_DEV_02C8&VEN_8086&DEV_0222'); CompatibleIDs=@(); ParentHardwareIDs=@('PCI\VEN_8086&DEV_02C8&SUBSYS_86B2103C&REV_00') })
Check "child DSP device matches via its PCI parent id" (@(Test-HpCvaAppliesToDevices -Cva $cvaSst -MissingDevices $devDsp -SysId '86B2').Count -eq 1)
$cands = @(@{ Id='sp138969'; Name='Synaptics Fingerprint Driver'; Version='6.0.60.1111'; Devices=@('WBDI') },
           @{ Id='sp98600';  Name='Synaptics Fingerprint Driver - Comet Lake'; Version='6.0.20.1111'; Devices=@('WBDI') },
           @{ Id='sp105483'; Name='Synaptics Fingerprint Sensor Driver'; Version='6.0.39.1111'; Devices=@('WBDI') },
           @{ Id='sp111684'; Name='NVIDIA Graphics Driver'; Version='27.21.14.5206'; Devices=@('3D Video Controller') },
           @{ Id='sp100206'; Name='NVIDIA Graphics Driver - Comet Lake'; Version='26.21.14.3198'; Devices=@('3D Video Controller') })
$win = @(Select-HpConsumerSoftpaqs -Candidates $cands)
Check "newest package per device: 2 winners (sp138969, sp111684)" ($win.Count -eq 2 -and @($win | Where-Object { $_.Id -eq 'sp138969' }).Count -eq 1 -and @($win | Where-Object { $_.Id -eq 'sp111684' }).Count -eq 1)
# v1.30.5 regression: a New-Object List[object] is PSObject-wrapped; @() on it
# threw "Argument types do not match" and aborted the run.
$candList = New-Object System.Collections.Generic.List[object]
foreach ($c in $cands) { $candList.Add($c) | Out-Null }
$winList = @()
try { $winList = @(Select-HpConsumerSoftpaqs -Candidates $candList) } catch { $winList = @() }
Check "selector accepts a New-Object List[object] (2 winners)" ($winList.Count -eq 2)
Check "version compare: 6.0.60.1111 newer than 6.0.39.1111" ((Test-HpVersionNewer -Candidate '6.0.60.1111' -Current '6.0.39.1111') -eq $true)
Check "version compare: text version loses to a numeric one" ((Test-HpVersionNewer -Candidate 'Rev.H' -Current '1.0') -eq $false)
Check "CVA with no SysId check still matches" (@(Test-HpCvaAppliesToDevices -Cva $cva -MissingDevices $devs -SysId '').Count -eq 1)
""
"passed: $pass  failed: $fail"
if ($fail -gt 0) { exit 1 }
