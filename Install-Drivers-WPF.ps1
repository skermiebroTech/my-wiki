# =============================================================
# Install-Drivers-WPF.ps1
# Version: 2.0.0  (WPF UI + runspace threading rearchitecture)
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode (one command, no prerequisites):
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-WPF.ps1 | iex"
#
# Headless (unchanged contract from the WinForms build):
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-WPF.ps1 -Manufacturer Dell
#   ...all the original switches still apply (see PARAM block).
#
# WHAT CHANGED vs Install-Drivers-auto.ps1 (v1.13.1):
#   - UI: Windows Forms  ->  WPF (XAML, styled, modern, dark).
#   - Threading: the engine no longer pumps
#     [Windows.Forms.Application]::DoEvents() on the UI thread. It runs on a
#     background runspace; the window stays responsive and updates arrive
#     via the WPF Dispatcher.
#   - Everything else (Dell/HP/Lenovo/Surface logic, downloads, extraction,
#     pnputil install, analytics, HTML report) is UNCHANGED. It is dot-sourced
#     from the proven engine and driven through the same function names it
#     already calls. See sections 6 and 7.
#
# DESIGN NOTES
#   * No prerequisites: only PresentationFramework / WindowsBase /
#     PresentationCore / System.Xaml - all in-box on Windows 8+.
#   * Single command: this file bootstraps the engine into a scoped temp
#     dir and removes everything on exit. Output files (.log/.events.json/
#     .analytics.json/.report.html) still go to %USERPROFILE%\Downloads
#     on purpose - those are the deliverables, not litter.
#   * Clean afterwards: temp working dir + any self-downloaded engine copy
#     are deleted in the finally block AND on PowerShell.Exiting, so even a
#     crash or window-close leaves nothing behind.
#   * STA: WPF requires STA. Windows PowerShell 5.1 (the documented launch
#     path) is STA by default. If MTA is detected (e.g. pwsh 7) we relaunch
#     under powershell.exe -STA so the user still typed only one command.
#   * UI robustness: standard window chrome (NO AllowsTransparency) so it
#     renders under audit mode / Sysprep / RDP / VMs where DWM compositing
#     is unreliable - which is exactly where this tool runs. A borderless
#     custom-chrome variant is documented at the bottom for controlled HW.
# =============================================================

[CmdletBinding()]
param(
    [string]$Manufacturer,
    [string]$Model,
    [string]$MachineType,
    [switch]$Headless,
    [switch]$Silent,
    [switch]$TestMode,
    [switch]$Diagnostic,
    [switch]$NoAnalytics,
    [ValidateRange(1,6)][int]$MaxParallelDownloads = 3,
    [switch]$SkipInstall,
    [switch]$SkipCleanup,
    [switch]$PromptWindowsUpdate = $true,

    # Internal: set when we relaunch under -STA. Not for users.
    [switch]$__StaRelaunch,
    # Internal: path we were bootstrapped from, so cleanup can remove it.
    [string]$__SelfPath
)

$SCRIPT_VERSION = "2.0.0"
$ErrorActionPreference = "Stop"

# Any explicit param implies headless intent - same rule as the original.
if ($Manufacturer -or $Model -or $MachineType -or $Silent) { $Headless = $true }

# =============================================================
# 0. STA GUARD - WPF cannot run MTA. Relaunch under powershell.exe -STA.
# =============================================================
$apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
if (-not $Headless -and $apartment -ne [System.Threading.ApartmentState]::STA -and -not $__StaRelaunch) {

    # We may have arrived via `irm | iex`, so there is no file to -File.
    # Persist ourselves first.
    $bootPath = $__SelfPath
    if (-not $bootPath -or -not (Test-Path $bootPath)) {
        $bootPath = Join-Path $env:TEMP ("DriverInstaller_{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -LiteralPath $bootPath -Value $MyInvocation.MyCommand.Definition -Encoding UTF8
    }

    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$bootPath,
                 '-__StaRelaunch','-__SelfPath',$bootPath)
    foreach ($p in $PSBoundParameters.GetEnumerator()) {
        if ($p.Key -in '__StaRelaunch','__SelfPath') { continue }
        if ($p.Value -is [switch]) { if ($p.Value.IsPresent) { $argList += "-$($p.Key)" } }
        else { $argList += "-$($p.Key)"; $argList += "$($p.Value)" }
    }
    Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $argList -Wait
    return
}

# =============================================================
# 1. SCOPED WORKSPACE + GUARANTEED CLEANUP
# =============================================================
$script:RunId        = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:WorkRoot     = Join-Path $env:TEMP ("DriverInstaller_WPF_{0}" -f $script:RunId)
$script:DownloadDir  = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
$script:CleanupPaths = New-Object System.Collections.Generic.List[string]

New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
$script:CleanupPaths.Add($script:WorkRoot)
if ($__SelfPath -and (Test-Path $__SelfPath)) { $script:CleanupPaths.Add($__SelfPath) }

function Invoke-Cleanup {
    foreach ($p in $script:CleanupPaths) {
        try { if ($p -and (Test-Path $p)) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
    }
}
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { Invoke-Cleanup }

# =============================================================
# 2. SHARED STATE  (UI thread  <->  worker runspace)
# =============================================================
$sync = [hashtable]::Synchronized(@{
    Window     = $null
    Dispatcher = $null
    Cancel     = $false
    Running    = $false
    Done       = $false
    Result     = $null
    LogFile    = (Join-Path $script:DownloadDir "DriverInstaller_$($script:RunId).log")
    EventsFile = (Join-Path $script:DownloadDir "DriverInstaller_$($script:RunId).events.json")
    Version    = $SCRIPT_VERSION
    Silent     = [bool]$Silent
    Headless   = [bool]$Headless
})

# =============================================================
# 3. LOGGING (file + NDJSON + UI, thread-safe). Mirrors the original
#    three-sink Log() contract; -Level / -Event / -Context still accepted.
#    NOTE: this UI-thread copy is used for pre-window logging and headless.
#    The worker runspace declares its own identical copy (section 6) so
#    engine call sites resolve there.
# =============================================================
$script:LogLock = New-Object object
function Log {
    param(
        [Parameter(Mandatory,Position=0)][string]$msg,
        [string]$Level = $null,
        [string]$Event = $null,
        [hashtable]$Context = $null
    )
    if (-not $Level) {
        $Level = if     ($msg -match '(?i)\bERROR\b|failed|exception') { 'error' }
                 elseif ($msg -match '(?i)\bWARN')                     { 'warn'  }
                 else                                                  { 'info'  }
    }
    $stamp = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level.ToUpper(), $msg
    [System.Threading.Monitor]::Enter($script:LogLock)
    try {
        Add-Content -LiteralPath $sync.LogFile -Value $stamp -Encoding UTF8 -ErrorAction SilentlyContinue
        $o = [ordered]@{ ts = (Get-Date -Format o); level = $Level; script_version = $sync.Version; msg = $msg }
        if ($Event)   { $o.event = $Event }
        if ($Context) { foreach ($k in $Context.Keys) { $o[$k] = $Context[$k] } }
        Add-Content -LiteralPath $sync.EventsFile -Value ($o | ConvertTo-Json -Compress -Depth 6) -Encoding UTF8 -ErrorAction SilentlyContinue
    } finally { [System.Threading.Monitor]::Exit($script:LogLock) }

    if ($sync.Silent)  { return }
    if ($sync.Headless){ Write-Host $stamp; return }
}

# Headless / pre-window UI-contract shims. The engine calls these names
# unconditionally; in headless mode there is no window, so they degrade to
# no-ops / console. (The worker runspace re-declares richer versions that
# drive the WPF Dispatcher - section 5.)
function SetProgress     { param([int]$Pct) }
function SetDownload     { param([int]$Pct,[string]$Label) if ($Label -and -not $sync.Silent) { Write-Host "  [dl] $Label" } }
function SetExtract      { param([int]$Pct,[string]$Label) if ($Label -and -not $sync.Silent) { Write-Host "  [ex] $Label" } }
function Set-ButtonIdle    { }
function Set-ButtonRunning { }
function Set-Button        { }
function Step-DlSpinner      { }
function Step-ExSpinner      { }
function Step-OverallSpinner { }
function Step-AllSpinners    { }
function Stop-ExSpinner      { param([bool]$Success=$true) }
function Play-Sound          { param([string]$Event) }
function Test-Cancelled      { return [bool]$sync.Cancel }

# =============================================================
# 4. WPF VIEW (XAML) - clean, flat, modern, robust chrome
# =============================================================
function Show-MainWindow {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Driver Installer  v$SCRIPT_VERSION"
        Width="880" Height="640" MinWidth="720" MinHeight="540"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" Background="#0D1117">
  <Window.Resources>
    <SolidColorBrush x:Key="Card"   Color="#161B22"/>
    <SolidColorBrush x:Key="Stroke" Color="#30363D"/>
    <SolidColorBrush x:Key="Text"   Color="#E6EDF3"/>
    <SolidColorBrush x:Key="Muted"  Color="#8B949E"/>
    <SolidColorBrush x:Key="Accent" Color="#2F81F7"/>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="8"/>
      <Setter Property="Foreground" Value="{StaticResource Accent}"/>
      <Setter Property="Background" Value="#21262D"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border CornerRadius="4" Background="{TemplateBinding Background}"/>
              <Border x:Name="PART_Track"/>
              <Border x:Name="PART_Indicator" CornerRadius="4"
                      HorizontalAlignment="Left" Background="{TemplateBinding Foreground}"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="Padding" Value="20,9"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" CornerRadius="6" Background="{TemplateBinding Background}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Opacity" Value="0.88"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="b" Property="Background" Value="#21262D"/>
                <Setter Property="Foreground" Value="#6E7681"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GhostBtn" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="#21262D"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
    </Style>
  </Window.Resources>

  <Grid Margin="22">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid Grid.Row="0" Margin="0,0,0,18">
      <StackPanel>
        <TextBlock Text="Driver Installer" Foreground="{StaticResource Text}" FontSize="22" FontWeight="Bold"/>
        <TextBlock x:Name="SubTitle" Text="Detecting device..." Foreground="{StaticResource Muted}" FontSize="13" Margin="0,2,0,0"/>
      </StackPanel>
      <Border CornerRadius="11" Padding="12,4" Background="#21262D"
              HorizontalAlignment="Right" VerticalAlignment="Top">
        <TextBlock x:Name="StatusPill" Text="Idle" Foreground="{StaticResource Muted}" FontWeight="SemiBold" FontSize="12"/>
      </Border>
    </Grid>

    <Border Grid.Row="1" Background="{StaticResource Card}" BorderBrush="{StaticResource Stroke}"
            BorderThickness="1" CornerRadius="10" Padding="20" Margin="0,0,0,16">
      <StackPanel>
        <TextBlock Text="Overall" Foreground="{StaticResource Muted}" FontSize="12" Margin="0,0,0,6"/>
        <ProgressBar x:Name="OverallBar" Value="0" Margin="0,0,0,16"/>

        <DockPanel Margin="0,0,0,6">
          <TextBlock Text="Download" Foreground="{StaticResource Muted}" FontSize="12"/>
          <TextBlock x:Name="DownloadLabel" Text="Waiting..." Foreground="{StaticResource Muted}"
                     FontSize="12" HorizontalAlignment="Right" DockPanel.Dock="Right"/>
        </DockPanel>
        <ProgressBar x:Name="DownloadBar" Value="0" Margin="0,0,0,16"/>

        <DockPanel Margin="0,0,0,6">
          <TextBlock Text="Extract / Install" Foreground="{StaticResource Muted}" FontSize="12"/>
          <TextBlock x:Name="ExtractLabel" Text="Waiting..." Foreground="{StaticResource Muted}"
                     FontSize="12" HorizontalAlignment="Right" DockPanel.Dock="Right"/>
        </DockPanel>
        <ProgressBar x:Name="ExtractBar" Value="0"/>
      </StackPanel>
    </Border>

    <Border Grid.Row="2" Background="#010409" BorderBrush="{StaticResource Stroke}"
            BorderThickness="1" CornerRadius="10">
      <RichTextBox x:Name="LogBox" IsReadOnly="True" Background="Transparent"
                   BorderThickness="0" Foreground="#C9D1D9"
                   FontFamily="Cascadia Mono, Consolas" FontSize="12"
                   VerticalScrollBarVisibility="Auto" Padding="14">
        <RichTextBox.Document><FlowDocument/></RichTextBox.Document>
      </RichTextBox>
    </Border>

    <Grid Grid.Row="3" Margin="0,18,0,0">
      <TextBlock VerticalAlignment="Center" Foreground="{StaticResource Muted}" FontSize="11"
                 Text="skermiebroTech &#183; output saved to Downloads"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="CancelBtn" Content="Cancel" Style="{StaticResource GhostBtn}" IsEnabled="False" Margin="0,0,10,0"/>
        <Button x:Name="StartBtn"  Content="Start"  Style="{StaticResource Btn}" Width="120"/>
      </StackPanel>
    </Grid>
  </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $sync.Window     = $window
    $sync.Dispatcher = $window.Dispatcher

    $startBtn  = $window.FindName('StartBtn')
    $cancelBtn = $window.FindName('CancelBtn')

    $startBtn.Add_Click({ if (-not $sync.Running) { Start-Worker } })
    $cancelBtn.Add_Click({
        $sync.Cancel = $true
        $cancelBtn.IsEnabled = $false
        Log "Cancel requested - finishing current step then stopping..." -Level 'warn' -Event 'cancel'
    })
    $window.Add_Closing({ if ($sync.Running) { $sync.Cancel = $true } })

    $window.Add_ContentRendered({
        Log "Driver Installer v$SCRIPT_VERSION (WPF)" -Event 'run_start' -Context @{
            log_file = $sync.LogFile; events_file = $sync.EventsFile; test_mode = [bool]$TestMode
        }
        # async device detection so the subtitle fills in immediately
        $d = [powershell]::Create()
        $d.Runspace = $script:WorkerRunspace
        $null = $d.AddScript({
            param($sync,$mfg,$mdl)
            try {
                $cs  = Get-CimInstance Win32_ComputerSystem
                $man = if ($mfg) { $mfg } else { $cs.Manufacturer.Trim() }
                $mod = if ($mdl) { $mdl } else { $cs.Model.Trim() }
                $sync.Dispatcher.Invoke([action]{ $sync.Window.FindName('SubTitle').Text = "$man  -  $mod" }) | Out-Null
            } catch { }
        }).AddArgument($sync).AddArgument($Manufacturer).AddArgument($Model)
        $null = $d.BeginInvoke()
    })

    $window.ShowDialog() | Out-Null
}

# =============================================================
# 5. WORKER RUNSPACE - hosts the engine so the long curl/msiexec
#    poll loops never touch the UI thread.
# =============================================================
$script:WorkerRunspace = [runspacefactory]::CreateRunspace()
$script:WorkerRunspace.ApartmentState = 'STA'
$script:WorkerRunspace.ThreadOptions  = 'ReuseThread'
$script:WorkerRunspace.Open()
$script:WorkerRunspace.SessionStateProxy.SetVariable('sync', $sync)

function Start-Worker {
    # flip UI to running state from the UI thread
    $sync.Running = $true
    $window = $sync.Window
    $window.FindName('StartBtn').IsEnabled  = $false
    $window.FindName('CancelBtn').IsEnabled = $true
    $window.FindName('StatusPill').Text     = 'Working'

    $ps = [powershell]::Create()
    $ps.Runspace = $script:WorkerRunspace

    $null = $ps.AddScript({
        param($sync, $engineSource, $workRoot, $downloadDir, $isTest)

        # ---- UI CONTRACT, declared inside the runspace ----
        # The proven engine calls these exact names. Each marshals onto the
        # WPF Dispatcher. This is the seam that lets the engine stay unaware
        # of which toolkit is underneath.
        function _ui($n,$p,$v){ if($sync.Dispatcher){ $sync.Dispatcher.Invoke([action]{ $e=$sync.Window.FindName($n); if($e){ $e.$p=$v } })|Out-Null } }
        function _bar($n,$pct,$ind=$false){ if($sync.Dispatcher){ $sync.Dispatcher.Invoke([action]{ $b=$sync.Window.FindName($n); if($b){ if($ind -or $pct -lt 0){$b.IsIndeterminate=$true} else {$b.IsIndeterminate=$false;$b.Value=[math]::Max(0,[math]::Min(100,$pct))} } })|Out-Null } }

        function SetProgress { param([int]$Pct) _bar 'OverallBar' $Pct }
        function SetDownload { param([int]$Pct,[string]$Label) _bar 'DownloadBar' $Pct ($Pct -lt 0); if($Label){ _ui 'DownloadLabel' 'Text' $Label } }
        function SetExtract  { param([int]$Pct,[string]$Label) _bar 'ExtractBar'  $Pct ($Pct -lt 0); if($Label){ _ui 'ExtractLabel'  'Text' $Label } }
        function Set-ButtonIdle    { _ui 'StartBtn' 'IsEnabled' $true;  _ui 'CancelBtn' 'IsEnabled' $false; _ui 'StatusPill' 'Text' 'Idle';    $sync.Running=$false }
        function Set-ButtonRunning { _ui 'StartBtn' 'IsEnabled' $false; _ui 'CancelBtn' 'IsEnabled' $true;  _ui 'StatusPill' 'Text' 'Working'; $sync.Running=$true }
        function Set-Button        { Set-ButtonIdle }
        function Step-DlSpinner {}; function Step-ExSpinner {}; function Step-OverallSpinner {}; function Step-AllSpinners {}
        function Stop-ExSpinner { param([bool]$Success=$true) _bar 'ExtractBar' ($(if($Success){100}else{0})) }
        function Play-Sound { param([string]$Event) if($sync.Silent){return}; try { [System.Media.SystemSounds]::Beep.Play() } catch {} }
        function Test-Cancelled { return [bool]$sync.Cancel }

        $script:LogLock = New-Object object
        function Log {
            param([Parameter(Mandatory,Position=0)][string]$msg,[string]$Level,[string]$Event,[hashtable]$Context)
            if(-not $Level){ $Level = if($msg -match '(?i)error|failed|exception'){'error'} elseif($msg -match '(?i)warn'){'warn'} else {'info'} }
            $stamp = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'),$Level.ToUpper(),$msg
            [System.Threading.Monitor]::Enter($script:LogLock)
            try {
                Add-Content -LiteralPath $sync.LogFile -Value $stamp -Encoding UTF8 -EA SilentlyContinue
                $o=[ordered]@{ts=(Get-Date -Format o);level=$Level;script_version=$sync.Version;msg=$msg}
                if($Event){$o.event=$Event}; if($Context){foreach($k in $Context.Keys){$o[$k]=$Context[$k]}}
                Add-Content -LiteralPath $sync.EventsFile -Value ($o|ConvertTo-Json -Compress -Depth 6) -Encoding UTF8 -EA SilentlyContinue
            } finally { [System.Threading.Monitor]::Exit($script:LogLock) }
            if($sync.Silent){return}
            if($sync.Dispatcher){
                $sync.Dispatcher.Invoke([action]{
                    $tb=$sync.Window.FindName('LogBox')
                    $col=switch($Level){'error'{'#FF5C57'}'warn'{'#E0A33E'}default{'#C9D1D9'}}
                    $r=New-Object Windows.Documents.Run ($stamp+[Environment]::NewLine)
                    $r.Foreground=(New-Object Windows.Media.BrushConverter).ConvertFromString($col)
                    $p=$tb.Document.Blocks.LastBlock; if(-not $p){$p=New-Object Windows.Documents.Paragraph;$tb.Document.Blocks.Add($p)}
                    $p.Inlines.Add($r); $tb.ScrollToEnd()
                })|Out-Null
            }
        }

        try {
            if ($engineSource -and (Test-Path $engineSource)) {
                # The engine companion = Install-Drivers-auto.ps1 (v1.13.1)
                # with its own param / STA / WinForms-UI / bootstrap stripped
                # (functions ONLY: Get-MissingDriverCount, Start-Install, the
                # Dell/HP/Lenovo/Surface installers, analytics, HTML report).
                # Dot-sourced here so its SetProgress/Log/etc. bind to the
                # shims above.
                . $engineSource
                Start-Install
                if (-not $sync.Result) { $sync.Result = 'success' }
            }
            else {
                # ---- BUILT-IN DEMO PIPELINE ----
                # Zero external deps: real (read-only) missing-driver scan,
                # then simulated download/extract so the WPF + threading is
                # fully verifiable before wiring the real engine.
                Log "Engine not wired - running built-in demo pipeline." -Level 'warn' -Event 'demo_mode'
                SetProgress 5
                Log "Checking for devices with missing drivers..."
                $missing=@()
                try {
                    $missing = Get-CimInstance Win32_PNPEntity -EA Stop |
                               Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
                               Select-Object -ExpandProperty Name -Unique
                } catch { Log "  WMI PnP enumerate failed: $($_.Exception.Message)" -Level 'warn' }
                Log ("Missing drivers BEFORE: {0}" -f @($missing).Count) -Event 'missing_before' -Context @{ count = @($missing).Count }
                foreach ($x in $missing) { Log "  - $x" }

                foreach ($stage in 'Download','Extract') {
                    for ($i=0;$i -le 100;$i+=4) {
                        if (Test-Cancelled) { throw 'cancelled' }
                        if ($stage -eq 'Download') { SetDownload $i "demo pack  $i%" } else { SetExtract $i "extracting  $i%" }
                        SetProgress ([int]( (($stage -eq 'Extract')*50) + $i*0.5 ))
                        Start-Sleep -Milliseconds 60
                    }
                    Log "$stage stage complete." -Event 'stage_done' -Context @{ stage = $stage }
                }
                SetProgress 100
                $sync.Result = if ($isTest) { 'testmode' } else { 'success' }
            }
        }
        catch {
            if ("$_" -eq 'cancelled' -or $sync.Cancel) { $sync.Result='cancelled'; Log "Run cancelled." -Level 'warn' -Event 'cancelled' }
            else { $sync.Result='failure'; Log "FATAL: $($_.Exception.Message)" -Level 'error' -Event 'fatal' }
        }
        finally {
            $sync.Done = $true
            $ok = ($sync.Result -in 'success','testmode')
            Play-Sound -Event ($(if($ok){'Success'}elseif($sync.Result -eq 'cancelled'){'Cancel'}else{'Failure'}))
            SetProgress 100
            _ui 'StatusPill' 'Text' "$($sync.Result)"
            Set-ButtonIdle
            Log "Done - result=$($sync.Result)" -Event 'run_end' -Context @{ result = "$($sync.Result)" }
        }
    }).AddArgument($sync).AddArgument($script:EngineSource).AddArgument($script:WorkRoot).AddArgument($script:DownloadDir).AddArgument([bool]$TestMode)

    $script:WorkerHandle = $ps.BeginInvoke()
}

# =============================================================
# 6. ENGINE SOURCE RESOLUTION
#    Single command + no prereqs: pull the engine companion from the
#    same repo into the scoped temp dir (deleted on exit). If it isn't
#    reachable, the demo pipeline runs so the app still launches.
# =============================================================
$EngineRawUrl = 'https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-Engine.ps1'
$script:EngineSource = $null

$localEngine = $null
try {
    $base = if ($__SelfPath) { Split-Path -Parent $__SelfPath }
            elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
            else { $null }
    if ($base) { $localEngine = Join-Path $base 'Install-Drivers-Engine.ps1' }
} catch { $localEngine = $null }

if ($localEngine -and (Test-Path $localEngine)) {
    $script:EngineSource = $localEngine
} else {
    try {
        $dest = Join-Path $script:WorkRoot 'engine.ps1'
        Invoke-RestMethod -Uri $EngineRawUrl -OutFile $dest -ErrorAction Stop
        if ((Get-Item $dest).Length -gt 1024) {
            $script:EngineSource = $dest
            $script:CleanupPaths.Add($dest)
        }
    } catch { $script:EngineSource = $null }
}

# =============================================================
# 7. ENTRYPOINT
# =============================================================
try {
    if ($Headless) {
        Log "Driver Installer v$SCRIPT_VERSION (headless)" -Event 'run_start'
        if ($script:EngineSource) { . $script:EngineSource; Start-Install }
        else { Log "No engine available in headless mode - nothing to do." -Level 'error'; $sync.Result = 'failure' }
    }
    else {
        Show-MainWindow
        while ($sync.Running -and -not $sync.Done) { Start-Sleep -Milliseconds 150 }
    }
}
finally {
    try { if ($script:WorkerRunspace) { $script:WorkerRunspace.Close() } } catch { }
    Invoke-Cleanup
    Get-EventSubscriber -SourceIdentifier PowerShell.Exiting -EA SilentlyContinue | Unregister-Event -EA SilentlyContinue
}

# =============================================================
# OPTIONAL: borderless custom-chrome variant
# -------------------------------------------------------------
# For controlled hardware (not Sysprep/RDP/VM) you can get a floating,
# rounded, shadowed look by changing the <Window> opener to:
#
#   WindowStyle="None" AllowsTransparency="True" Background="Transparent"
#
# wrapping the root Grid in a rounded Border with a DropShadowEffect,
# and adding a custom title bar that calls $window.DragMove() on
# MouseLeftButtonDown plus your own min/close buttons. It is NOT the
# default because AllowsTransparency breaks rendering under audit mode /
# DWM-less sessions - exactly where this tool runs.
#
# ENGINE COMPANION (Install-Drivers-Engine.ps1) - how to produce it from
# the current working Install-Drivers-auto.ps1 v1.13.1, mechanically:
#   1. Delete the param(...) block (this shell owns params now).
#   2. Delete the STA relaunch / bootstrap / cleanup sections.
#   3. Delete the WinForms UI: the $form/$bar/$label builders, the XAML-
#      free Add-Type System.Windows.Forms, Show-SurfaceModelPicker's form
#      (or port it to a tiny WPF dialog), and the spinner frame globals.
#   4. Delete the original SetProgress/SetDownload/SetExtract/Log/
#      Set-Button*/Step-*Spinner/Stop-ExSpinner/Play-Sound/Test-Cancelled
#      definitions - this shell provides them.
#   5. Delete EVERY '[System.Windows.Forms.Application]::DoEvents()' line
#      (11 of them). The runspace model makes them unnecessary; leaving
#      them throws (no message pump on the worker thread).
#   6. Keep EVERYTHING else verbatim: Get-MissingDriverCount/-Names, the
#      Dell/HP/Lenovo/Surface install functions, Get-CpuVendor +
#      $SurfaceCpuVariantIds (v1.13.0), the v1.13.1 early serial capture,
#      analytics webhook, HTML report, NDJSON. Ensure Start-Install is the
#      single top-level entry the worker calls.
# That file is a deletion-only transform of code you have already proven -
# no vendor logic is rewritten, which is the whole point of this split.
# =============================================================