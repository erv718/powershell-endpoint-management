# Barcode scanner USB reset - added August 2025
# Fixes intermittent scanner dropouts at club front desks
# Barcode scanner USB reset - added April 2022
# Fixes intermittent scanner dropouts at club front desks
<#
  Check-ScannerCOM-FAST.ps1  (Run as Administrator)
  - Closes only eXerp across ALL users (~5s total)
  - Auto-detects scanner from current devices (no unplug/re-plug)
  - FINAL SUMMARY box: Hostname, RAM, OS, Barcode Scanner, Disk MediaType
  - Logs to C:\Temp\eXerp\BarcodeScanner\
  - PowerShell 5.1 compatible, StrictMode-safe, paste-friendly (no elseif in main flow)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ===== Logging =====
$LogDir = 'C:\Temp\eXerp\BarcodeScanner'
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$Stamp  = (Get-Date).ToString('yyyyMMdd_HHmmss')
$Log    = Join-Path $LogDir "Scanner_Check_$Stamp.log"
try { Start-Transcript -Path $Log -Force -ErrorAction Stop } catch {}

# ===== Admin check =====
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { Write-Host "Please run this script as Administrator." -ForegroundColor Red; try { Stop-Transcript | Out-Null } catch {}; exit 1 }

# ===== UI helpers =====
function Section([string]$t){ Write-Host "`n==================== $t ====================" -ForegroundColor Cyan }
function MsgInfo ([string]$m){ Write-Host ("[i] "+$m) -ForegroundColor Cyan }
function MsgOk   ([string]$m){ Write-Host ("[] "+$m) -ForegroundColor Green }
function MsgWarn ([string]$m){ Write-Host ("[!] "+$m) -ForegroundColor Yellow }
function MsgErr  ([string]$m){ Write-Host ("[x] "+$m) -ForegroundColor Red }

# ===== Pretty summary box =====
function Show-Box {
    param([string[]]$Lines)
    $width = ($Lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    if (-not $width -or $width -lt 20) { $width = 20 }
    $top = '+' + ('-' * ($width + 2)) + '+'
    Write-Host $top -ForegroundColor Cyan
    foreach($line in $Lines){
        $pad = ' ' * ($width - $line.Length)
        Write-Host ('| ' + $line + $pad + ' |') -ForegroundColor Cyan
    }
    Write-Host $top -ForegroundColor Cyan
}

# ===== Quick system info =====
function Get-RamGB { try { [int][math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB) } catch { $null } }
function Get-OSName {
    try { $b = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber; if ($b -ge 22000) { 'Windows 11' } else { 'Windows 10' } }
    catch { '(unknown OS)' }
}

# ===== Device helpers =====
function Get-ComSnapshot {
    $ports = @()
    try {
        $ports = @(Get-CimInstance Win32_SerialPort | Sort-Object DeviceID | ForEach-Object {
            [pscustomobject]@{
                DeviceID    = $_.DeviceID
                Caption     = $_.Caption
                PNPDeviceID = $_.PNPDeviceID
                IsUSB       = [bool]($_.PNPDeviceID -like '*USB*')
            }
        })
    } catch {}
    $reg = @()
    try {
        $reg = @(Get-ItemProperty -Path 'HKLM:\HARDWARE\DEVICEMAP\SERIALCOMM' -ErrorAction SilentlyContinue |
              Select-Object -ExcludeProperty PS* |
              ForEach-Object {
                  $_.PSObject.Properties |
                  Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' } |
                  ForEach-Object { [pscustomobject]@{ DeviceSymbol=$_.Name; PortName=$_.Value } }
              })
    } catch {}
    [pscustomobject]@{ Ports=$ports; Map=$reg }
}
function Get-HIDKeyboardSnapshot {
    $kbGuid = '{4d36e96b-e325-11ce-bfc1-08002be10318}'
    try {
        @(Get-CimInstance Win32_PnPEntity) | Where-Object {
            $_.PNPClass -eq 'Keyboard' -or
            ($_.ClassGuid -and $_.ClassGuid -ieq $kbGuid) -or
            ($_.Name -match '(?i)HID.*Keyboard')
        } | Select-Object Name, PNPDeviceID, Manufacturer, Status
    } catch { @() }
}

# ===== Close only eXerp (FAST, 5s cap) =====
function ForceKill-Exerp {
    param([int]$MaxSeconds = 5,[int]$PollMs = 150)
    $deadline = (Get-Date).AddSeconds($MaxSeconds)
    do {
        cmd.exe /c "taskkill /IM exerp*.exe /T /F >NUL 2>&1"
        Start-Sleep -Milliseconds $PollMs
        $alive = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'exerp*.exe' })
        foreach ($p in $alive) { if (Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue) { cmd.exe /c "taskkill /PID $($p.ProcessId) /T /F >NUL 2>&1" } }
        Start-Sleep -Milliseconds $PollMs
    } while ((Get-Date) -lt $deadline)
    $left = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'exerp*.exe' })
    if ($left.Count -gt 0) { MsgWarn ("eXerp still running after forced kill: " + (($left | Group-Object Name | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ')); $false } else { MsgOk "eXerp terminated."; $true }
}

# ===== Begin run =====
Section "Scanner Quick Check"
MsgInfo ("Computer : {0}" -f $env:COMPUTERNAME)
$ramGB = Get-RamGB
if ($ramGB) { MsgInfo ("Installed RAM : {0} GB" -f $ramGB) } else { MsgWarn "Installed RAM : (could not determine)" }

# Kill only eXerp (~5s)
Section "Closing eXerp"
MsgWarn "Force killing eXerp only (leaving Edge/WebView2 open)..."
$null = ForceKill-Exerp -MaxSeconds 5 -PollMs 150

# Current device snapshots (no unplug/replug required)
$comSnap = Get-ComSnapshot
$hidSnap = @(Get-HIDKeyboardSnapshot)

Section "Current (COM Ports)"
if(@($comSnap.Ports).Count -gt 0){ $comSnap.Ports | Select DeviceID,Caption,IsUSB,PNPDeviceID | Format-Table -AutoSize } else { MsgWarn "No COM ports via WMI." }

Section "Current (Keyboard/HID)"
if(@($hidSnap).Count -gt 0){ $hidSnap | Format-Table Name,Manufacturer,Status -AutoSize } else { MsgWarn "No Keyboard/HID devices found." }

# ===== Determine likely scanner from current devices =====
Section "Determine Scanner (no unplug required)"
$usbComs      = @($comSnap.Ports | Where-Object { $_.IsUSB })
$zebraComs    = @($usbComs     | Where-Object { $_.PNPDeviceID -match '(?i)VID_05E0' -or $_.Caption -match '(?i)Zebra|Symbol|Scanner' })
$zebraHidKeys = @($hidSnap     | Where-Object { $_.PNPDeviceID -match '(?i)VID_05E0' -or $_.Manufacturer -match '(?i)Zebra|Symbol' -or $_.Name -match '(?i)Zebra|Symbol' })

$ScannerType  = 'Unknown'
$ScannerPorts = @()

if (@($zebraComs).Count -gt 0) {
    $ScannerType  = 'COM'
    $ScannerPorts = @($zebraComs.DeviceID)
    MsgOk ("Likely scanner on COM (Zebra/Symbol signature): {0}" -f ($ScannerPorts -join ', '))
}
if ($ScannerType -eq 'Unknown' -and @($usbComs).Count -eq 1) {
    $ScannerType  = 'COM (inferred)'
    $ScannerPorts = @($usbComs[0].DeviceID)
    MsgOk ("Single USB COM present, inferred scanner: {0}" -f ($ScannerPorts -join ', '))
}
if ($ScannerType -eq 'Unknown' -and @($zebraHidKeys).Count -gt 0) {
    $ScannerType = 'HID Keyboard'
    MsgWarn "HID Keyboard with Zebra/Symbol signature found (scanner likely in HID mode)."
}
if ($ScannerType -eq 'Unknown' -and @($usbComs).Count -gt 1) {
    $ScannerType  = 'COM (multiple)'
    $ScannerPorts = @($usbComs.DeviceID)
    MsgWarn ("Multiple USB COM ports present: {0}" -f ($ScannerPorts -join ', '))
}

# ===== FINAL SUMMARY (single tidy block) =====
Section "FINAL SUMMARY"
$hostname = $env:COMPUTERNAME
$osName   = Get-OSName
$ramText  = $(if ($ramGB) { "$ramGB GB" } else { "(unknown)" })

# Disk MediaType via Get-PhysicalDisk (as requested)
$diskMediaType = '(unknown)'
try {
    $sysLetter = ($env:SystemDrive).TrimEnd(':','\')
    $sysDisk   = Get-Partition -DriveLetter $sysLetter -ErrorAction Stop | Get-Disk | Select-Object -First 1
    $pd = @(
        Get-PhysicalDisk | Where-Object {
            ($_.UniqueId -and $_.UniqueId -eq $sysDisk.UniqueId) -or
            ($_.FriendlyName -and $_.FriendlyName -eq $sysDisk.FriendlyName) -or
            ($_.DeviceId -eq $sysDisk.Number)
        }
    ) | Select-Object -First 1
    if (-not $pd) { $pd = Get-PhysicalDisk | Sort-Object Size -Descending | Select-Object -First 1 }
    $diskMediaType = [string]$pd.MediaType  # SSD / HDD / SCM / Unspecified
} catch { $diskMediaType = '(unknown)' }

# Build scanner text safely
$scannerText = 'Unknown'
if ($ScannerType -like 'COM*') {
    $portsText = if (@($ScannerPorts).Count -gt 0) { "Port(s): $(@($ScannerPorts) -join ', ')" } else { '' }
    $scannerText = $ScannerType + $(if ($portsText) { "  $portsText" } else { '' })
} elseif ($ScannerType -eq 'HID Keyboard') {
    $scannerText = 'HID Keyboard'
}

$lines = @(
    "Hostname            : $hostname",
    "Total RAM           : $ramText",
    "OS                  : $osName",
    "Barcode Scanner     : $scannerText",
    "Storage Disk Type   : $diskMediaType",
    "Log File            : $Log"
)
Show-Box -Lines $lines

Write-Host ""
MsgOk ("Log saved: {0}" -f $Log)
try{ Stop-Transcript | Out-Null }catch{}
