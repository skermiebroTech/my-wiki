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
$want='Get-MsCatalogCabUrls','Test-HpVersionNewer','ConvertTo-HpMatchId','Select-HpConsumerSoftpaqs','Invoke-HpSupportJson','ConvertFrom-HpJson','ConvertTo-HpBytes','Find-HpSupportProductOids','Select-HpSupportOs','Get-HpCvaInfo','Test-HpCvaAppliesToDevices','Get-MsCatalogResultRows','Select-MsCatalogDriverGuid','Find-DellIndexManifestPath','Test-DellOsCode','Test-DellOsArch','Test-ActionableProblemCode','Get-ProblemCodeHint','Select-HpDriverPack','Find-DellCatalogDeviceMatches','Get-NormalizedManufacturer','Test-PlaceholderManufacturer','Get-AsusModelCandidates','Select-AsusOsId','ConvertTo-AsusVersion','Get-AsusDriverEntries','Find-AsusDriverMatches'
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
# v1.30.6 regression: one .cab URL must survive as an array element, not a string.
$ddText = "downloadInformation[0].files[0].url = 'https://catalog.s.download.windowsupdate.com/d/msdownload/update/driver/drvs/2023/06/synaptics_1234abcd.cab';`nfiles[0].digest = 'x';"
$cabUrls = @(Get-MsCatalogCabUrls -Text $ddText)
Check "catalog download: one .cab URL stays a full URL at index 0" ($cabUrls.Count -eq 1 -and $cabUrls[0] -like 'https://*.cab')
$callerLine = (Get-Content $ScriptPath | Where-Object { $_ -match 'Resolve-MsCatalogDownloadUrls -Guid' -and $_ -notmatch '^\s*#' -and $_ -match '=' })
Check "catalog download caller wraps the resolver in @()" (@($callerLine).Count -eq 1 -and $callerLine -match '@\(Resolve-MsCatalogDownloadUrls')
Check "CVA with no SysId check still matches" (@(Test-HpCvaAppliesToDevices -Cva $cva -MissingDevices $devs -SysId '').Count -eq 1)

"== Manufacturer normalisation (v1.30.7) =="
# Run 20260921_110550: WMI said "Alienware", the dispatch tests "Dell", and the
# machine fell through to the Surface model picker with 26 drivers missing.
Check "Alienware maps to Dell"                 ((Get-NormalizedManufacturer -Manufacturer 'Alienware') -eq 'Dell')
Check "case and padding do not matter"         ((Get-NormalizedManufacturer -Manufacturer '  ALIENWARE ') -eq 'Dell')
Check "normalised value satisfies the dispatch" ((Get-NormalizedManufacturer -Manufacturer 'Alienware') -match 'Dell')
Check "normalised value satisfies 7-Zip prep"  ((Get-NormalizedManufacturer -Manufacturer 'Alienware') -match 'Dell|HP|Hewlett')
Check "Dell Inc. passes through unchanged"     ((Get-NormalizedManufacturer -Manufacturer 'Dell Inc.') -eq 'Dell Inc.')
Check "HP passes through unchanged"            ((Get-NormalizedManufacturer -Manufacturer 'HP') -eq 'HP')
Check "LENOVO passes through unchanged"        ((Get-NormalizedManufacturer -Manufacturer 'LENOVO') -eq 'LENOVO')
Check "Microsoft passes through unchanged"     ((Get-NormalizedManufacturer -Manufacturer 'Microsoft Corporation') -eq 'Microsoft Corporation')
Check "an unknown OEM still reaches the unknown branch" ((Get-NormalizedManufacturer -Manufacturer 'OEMBY') -eq 'OEMBY')
Check "an empty string stays empty"            ((Get-NormalizedManufacturer -Manufacturer '') -eq '')
# Ordering guard: analytics must capture the RAW string BEFORE normalisation,
# so the Sheet keeps Alienware and Dell Inc. apart.
$body      = Get-Content $ScriptPath
$iAnalytic = ($body | Select-String -SimpleMatch '$script:AnalyticsManufacturer = $manufacturer' | Select-Object -First 1).LineNumber
$iNorm     = ($body | Select-String -SimpleMatch '$manufacturer = Get-NormalizedManufacturer' | Select-Object -First 1).LineNumber
$iDispatch = ($body | Select-String -SimpleMatch 'elseif ($manufacturer -match "Dell")' | Select-Object -First 1).LineNumber
Check "analytics captures the raw string first" ($iAnalytic -and $iNorm -and $iAnalytic -lt $iNorm)
Check "normalisation runs before the dispatch"  ($iNorm -and $iDispatch -and $iNorm -lt $iDispatch)


"== Unknown brands skip the Surface picker (v1.30.8) =="
# Run 20260921_201150: a Razer Blade 16 got the Surface model list, the operator
# cancelled, and the run returned before the Windows Update / catalog fallbacks.
Check "Razer is a real brand"                  (-not (Test-PlaceholderManufacturer -Manufacturer 'Razer'))
Check "Micro-Star is a real brand"             (-not (Test-PlaceholderManufacturer -Manufacturer 'Micro-Star International Co., Ltd.'))
Check "ASUSTeK is a real brand"                (-not (Test-PlaceholderManufacturer -Manufacturer 'ASUSTeK COMPUTER INC.'))
Check "OEMBY is a placeholder"                 (Test-PlaceholderManufacturer -Manufacturer 'OEMBY')
Check "blank is a placeholder"                 (Test-PlaceholderManufacturer -Manufacturer '')
Check "whitespace is a placeholder"            (Test-PlaceholderManufacturer -Manufacturer '   ')
Check "To Be Filled By O.E.M. is a placeholder" (Test-PlaceholderManufacturer -Manufacturer 'To Be Filled By O.E.M.')
Check "System manufacturer is a placeholder"   (Test-PlaceholderManufacturer -Manufacturer 'System manufacturer')
Check "Default string is a placeholder"        (Test-PlaceholderManufacturer -Manufacturer 'Default string')
# Ordering guard: the real-brand skip must come before the picker and the headless error.
$iSkip   = ($body | Select-String -SimpleMatch 'No vendor driver source for' | Select-Object -First 1).LineNumber
$iPicker = ($body | Select-String -SimpleMatch '$pickedModel = Show-SurfaceModelPicker' | Select-Object -First 1).LineNumber
$iHlErr  = ($body | Select-String -SimpleMatch "doesn't match a known Surface" | Select-Object -First 1).LineNumber
Check "real-brand skip runs before the picker"        ($iSkip -and $iPicker -and $iSkip -lt $iPicker)
Check "real-brand skip runs before the headless error" ($iSkip -and $iHlErr -and $iSkip -lt $iHlErr)


"== ASUS support API (v1.31.0) =="
# Fixtures: live GetPDOS / GetPDDrivers answers for UX3405MA (Zenbook 14),
# captured 2026-09-21, trimmed to Networking/Chipset/Audio/Software entries.
$c1 = @(Get-AsusModelCandidates -Model 'ASUS Zenbook 14 UX3405MA_UX3405MA' -BaseboardProduct 'UX3405MA')
Check "model code: text after the last _ comes first"   ($c1.Count -ge 1 -and $c1[0] -eq 'UX3405MA')
Check "model code: no duplicates"                        (@($c1 | Where-Object { $_ -eq 'UX3405MA' }).Count -eq 1)
$c2 = @(Get-AsusModelCandidates -Model 'ROG Strix G614JV_G614JV' -BaseboardProduct 'G614JV')
Check "model code: ROG model"                            ($c2[0] -eq 'G614JV')
$c3 = @(Get-AsusModelCandidates -Model 'Vivobook 15 X1504VA' -BaseboardProduct '')
Check "model code: code-shaped token without _"          ($c3 -contains 'X1504VA' -and $c3 -notcontains 'VIVOBOOK')
Check "model code: empty input gives nothing"            (@(Get-AsusModelCandidates -Model '' -BaseboardProduct '').Count -eq 0)

$osJ = Get-Content (Join-Path $Fixtures 'asus-GetPDOS-UX3405MA.json') -Raw | ConvertFrom-Json
Check "OS id: Windows 11 64-bit is 52"                   ((Select-AsusOsId -OsJson $osJ -IsWin11 $true).Id -eq '52')
Check "OS id: Win10 machine falls back to the Win11 list" ((Select-AsusOsId -OsJson $osJ -IsWin11 $false).Id -eq '52')
Check "OS id: no list gives null"                        ($null -eq (Select-AsusOsId -OsJson $null -IsWin11 $true))

Check "version: V prefix removed"                        ((ConvertTo-AsusVersion 'V30.100.2318.58') -eq '30.100.2318.58')
Check "version: Sub suffix removed"                      ((ConvertTo-AsusVersion 'V6001.15.155.1Sub2') -eq '6001.15.155.1')

$drvJ = Get-Content (Join-Path $Fixtures 'asus-GetPDDrivers-UX3405MA-trimmed.json') -Raw | ConvertFrom-Json
$ent  = Get-AsusDriverEntries -DriverJson $drvJ
Check "entries: store links and no-hardware-id entries dropped" ($ent.Count -eq 9)
Check "entries: no Microsoft Store URL survives"         (@($ent | Where-Object { $_.Url -match 'microsoft\.com' }).Count -eq 0)
Check "entries: every entry has a 64-hex sha256"         (@($ent | Where-Object { $_.SHA256 -notmatch '^[0-9A-F]{64}$' }).Count -eq 0)
Check "entries: .exe extension and size parsed"          (@($ent | Where-Object { $_.Ext -ne '.exe' -or $_.Size -le 0 }).Count -eq 0)
Check "entries: hardware ids upper-case"                 ((@($ent | Where-Object { $_.Name -eq 'Intel Serial IO Driver' })[0].HwIds -contains 'PCI\VEN_8086&DEV_7E50'))

$asusMissing = @(
    [pscustomobject]@{ Name='Serial IO I2C Host Controller'; DeviceID='PCI\VEN_8086&DEV_7E50&SUBSYS_1C131043&REV_20\3&11583659&0&A8'; HardwareIDs=@('PCI\VEN_8086&DEV_7E50&SUBSYS_1C131043&REV_20','PCI\VEN_8086&DEV_7E50&SUBSYS_1C131043'); CompatibleIDs=@('PCI\VEN_8086&DEV_7E50&REV_20','PCI\VEN_8086&DEV_7E50'); ParentHardwareIDs=@() },
    [pscustomobject]@{ Name='Sensor Hub'; DeviceID='PCI\VEN_8086&DEV_7E45\x'; HardwareIDs=@('PCI\VEN_8086&DEV_7E45&SUBSYS_1C131043&REV_20'); CompatibleIDs=@(); ParentHardwareIDs=@() },
    [pscustomobject]@{ Name='High Definition Audio Device'; DeviceID='HDAUDIO\FUNC_01&VEN_10EC&DEV_0233&SUBSYS_104312A0\x'; HardwareIDs=@('HDAUDIO\FUNC_01&VEN_10EC&DEV_0233&SUBSYS_104312A0&REV_1000'); CompatibleIDs=@(); ParentHardwareIDs=@() },
    [pscustomobject]@{ Name='Unknown Razer thing'; DeviceID='ACPI\RZR0001\0'; HardwareIDs=@('ACPI\RZR0001'); CompatibleIDs=@(); ParentHardwareIDs=@() }
)
$win = @(Find-AsusDriverMatches -Entries $ent -MissingDevices $asusMissing)
Check "match: three packages for three matchable devices" ($win.Count -eq 3)
Check "match: Serial IO matched by prefix"               (@($win | Where-Object { $_.Name -eq 'Intel Serial IO Driver' }).Count -eq 1)
Check "match: newest sensor package kept (5.8.0.5)"      ((@($win | Where-Object { $_.Name -like '*Sensor*' })[0].Version) -eq '5.8.0.5')
Check "match: newest audio package kept (6.0.9780.1)"    ((@($win | Where-Object { $_.Name -eq 'Realtek Audio Driver' })[0].Version) -eq '6.0.9780.1')
Check "match: an unmatched device pulls no package"      (@($win | Where-Object { $_.Devices -contains 'Unknown Razer thing' }).Count -eq 0)
Check "match: no missing devices gives nothing"          (@(Find-AsusDriverMatches -Entries $ent -MissingDevices @()).Count -eq 0)

Check "dispatch: ASUSTeK string reaches the ASUS branch" ('ASUSTeK COMPUTER INC.' -match 'ASUS')
Check "dispatch: ASUS is not normalised away"            ((Get-NormalizedManufacturer -Manufacturer 'ASUSTeK COMPUTER INC.') -match 'ASUS')
$iAsus    = ($body | Select-String -SimpleMatch 'elseif ($manufacturer -match "ASUS")' | Select-Object -First 1).LineNumber
$iUnknown = ($body | Select-String -SimpleMatch "Unrecognised manufacturer: '" | Select-Object -First 1).LineNumber
Check "dispatch: ASUS branch sits before the unknown branch" ($iAsus -and $iUnknown -and $iAsus -lt $iUnknown)
Check "7-Zip prep includes ASUS"                          (@($body | Where-Object { $_ -match 'manufacturer -match "Dell\|HP\|Hewlett\|ASUS"' }).Count -eq 1)

""
"passed: $pass  failed: $fail"
if ($fail -gt 0) { exit 1 }
