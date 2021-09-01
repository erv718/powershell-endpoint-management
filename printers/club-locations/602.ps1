#BO Printer
# Variables
$driverDownloadPath = '\\SYSADMIN01\Tools\Printers\Clubs\Back Office\HP PageWide MFP P57750 XC\Driver'
# Get the first non-loopback IPv4 address
$localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.InterfaceAlias -notlike "*Loopback*" -and $_.PrefixOrigin -ne "WellKnown"
}).IPAddress

# If an IP address was found
if ($localIP) {
    # Extract the subnet from the IP address
    $subnetBase = $localIP -replace '(\d+\.\d+\.\d+\.)\d+', '$1'

    # Set the port address to be x.x.x.5
    $portAddress = $subnetBase + "5"

    # Create the port name by replacing '.' with '_'
    $portName = "IP_" + $portAddress.Replace('.', '_')

    # Output the results
    Write-Output "Port Name: $portName"
    Write-Output "Port Address: $portAddress"
} else {
    Write-Error "No non-loopback IPv4 address found on this computer."
}

# Use $portName and $portAddress as needed in the rest of your script
# Retrieve the system's hostname
$systemName = hostname
# Extract the first three numbers from the hostname
$clubID = if ($systemName -match '(\d{3})') { $matches[1] } else { "000" }

# Use the club ID as part of the printer name
$printerName = "${clubID} - Back Office Printer" # Printer display name with club ID from hostname

#$driverName = "HP PageWide MFP P57750 (NET)" # Must be exact! Install driver on your computer and copy its name and paste here!
$driverName = "HP Universal Printing PCL 6" # Must be exact! Install driver on your computer and copy its name and paste here!
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


#FD Printer
# Variables
$driverDownloadPath = '\\SYSADMIN01\Tools\Printers\Clubs\Front Desk\HP LaserJet 400 M401n\Driver'
# Get the first non-loopback IPv4 address
$localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.InterfaceAlias -notlike "*Loopback*" -and $_.PrefixOrigin -ne "WellKnown"
}).IPAddress

# If an IP address was found
if ($localIP) {
    # Extract the subnet from the IP address
    $subnetBase = $localIP -replace '(\d+\.\d+\.\d+\.)\d+', '$1'

    # Set the port address to be x.x.x.6
    $portAddress = $subnetBase + "76"

    # Create the port name by replacing '.' with '_'
    $portName = "IP_" + $portAddress.Replace('.', '_')

    # Output the results
    Write-Output "Port Name: $portName"
    Write-Output "Port Address: $portAddress"
} else {
    Write-Error "No non-loopback IPv4 address found on this computer."
}

# Use $portName and $portAddress as needed in the rest of your script
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