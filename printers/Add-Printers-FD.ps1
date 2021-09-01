get-printer | remove-printer
Get-PrinterPort | Remove-PrinterPort -ErrorAction SilentlyContinue

# Variables
$driverDownloadPath = '\\SYSADMIN01\Tools\Printers\Clubs\Back Office\Canon imageRUNNER1730iF\Driver'
$portName = "IP_10.1.6.5" # Printer port name
$portAddress = "10.1.6.5" # Printer IP
# Retrieve the system's hostname
$systemName = hostname
# Extract the first three numbers from the hostname
$clubID = if ($systemName -match '(\d{3})') { $matches[1] } else { "000" }

# Use the club ID as part of the printer name
$printerName = "${clubID} - Back Office Printer" # Printer display name with club ID from hostname

$driverName = "Canon Generic Plus PS3" # Must be exact! Install driver on your computer and copy its name and paste here!
# Do not modify these variables
$portExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue
$printerExists = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
#####################
# Waving magic wand #
#####################
Get-ChildItem $driverDownloadPath -Recurse -Filter "*.inf" -Force | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install } # Add to Windows Driver Store
Add-PrinterDriver -Name $driverName # Add Driver
# Add Printer Port
if (-not $portExists) {
  Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress -PortNumber '9100'
}
# Install Printer
if (-not $printerExists) {
Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
}
Set-PrintConfiguration -PrinterName $printerName -Color $false # Set Default Color to Grey Scale
(Get-WmiObject -ClassName Win32_Printer | Where-Object -Property Name -EQ $printerName).SetDefaultPrinter() # Set as default printer
Write-Output 'Done'
# Variables
$driverDownloadPath = '\\SYSADMIN01\Tools\Printers\Clubs\Front Desk\HP LaserJet 400 M401n\Driver'
$portName = "IP_10.1.6.6" # Printer port name
$portAddress = "10.1.6.6" # Printer IP
# Retrieve the system's hostname
$systemName = hostname
# Extract the first three numbers from the hostname
$clubID = if ($systemName -match '(\d{3})') { $matches[1] } else { "000" }

# Use the club ID as part of the printer name
$printerName = "${clubID} - Front Desk Printer" # Printer display name with club ID from hostname

$driverName = "HP LaserJet 400 M401 PCL 6" # Must be exact! Install driver on your computer and copy its name and paste here!
# Do not modify these variables
$portExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue
$printerExists = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
#####################
# Waving magic wand #
#####################
Get-ChildItem $driverDownloadPath -Recurse -Filter "*.inf" -Force | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install } # Add to Windows Driver Store
Add-PrinterDriver -Name $driverName # Add Driver
# Add Printer Port
if (-not $portExists) {
  Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
}
# Install Printer
if (-not $printerExists) {
Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
}
Set-PrintConfiguration -PrinterName $printerName -Color $false # Set Default Color to Grey Scale
(Get-WmiObject -ClassName Win32_Printer | Where-Object -Property Name -EQ $printerName).SetDefaultPrinter() # Set as default printer
Write-Output 'Done'